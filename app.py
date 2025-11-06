# -*- coding: utf-8 -*-
"""Minimal web UI to run debt collection automation scripts."""

from __future__ import annotations

import uuid
import json
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional

from flask import (
    Flask,
    redirect,
    render_template,
    request,
    send_from_directory,
    url_for,
)

from workflows import (
    bc_claims,
    default_judgment,
    demand_letter,
    dismissal,
    ontario_claims,
)

BASE_DIR = Path(__file__).resolve().parent
SESSIONS_DIR = BASE_DIR / "sessions"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class StepConfig:
    label: str
    description: str
    file_hint: str
    expected_files: int
    allowed_extensions: List[str]
    processor: Callable[[Path, List[Path]], List[Path]]
    dropzones: List[Dict[str, object]] = field(default_factory=list)
    optional_extensions: List[str] = field(default_factory=list)
    optional_max: int = 0


def process_demand_letter(session_dir: Path, uploads: List[Path]) -> List[Path]:
    output_dir = session_dir / "output"
    output_dir.mkdir(exist_ok=True)
    source = uploads[0]
    output_path, _summary = demand_letter.fill_demand_letter(
        pdf_path=source,
        template_path=None,
        output_dir=output_dir,
    )
    return [output_path]


def process_bc_claims(session_dir: Path, uploads: List[Path]) -> List[Path]:
    files_by_type: Dict[str, Path] = {}
    for candidate in uploads:
        name = candidate.name.upper()
        if "MRP" in name and "MRP" not in files_by_type:
            files_by_type["MRP"] = candidate
        elif "MRC" in name and "MRC" not in files_by_type:
            files_by_type["MRC"] = candidate
        elif "MRS" in name and "MRS" not in files_by_type:
            files_by_type["MRS"] = candidate
        elif ("CBR" in name or "CREDIT REPORT" in name) and "CBR" not in files_by_type:
            files_by_type["CBR"] = candidate

    missing = [key for key in ("MRP", "MRC", "MRS", "CBR") if key not in files_by_type]
    if missing:
        raise ValueError(
            f"Missing required files for BC Claims: {', '.join(missing)} "
            "(ensure filenames include MRP, MRC, MRS, CBR)."
        )

    inputs = bc_claims.ClaimsInputs(
        mrp=files_by_type["MRP"],
        mrc=files_by_type["MRC"],
        mrs=files_by_type["MRS"],
        cbr=files_by_type["CBR"],
    )
    combined = bc_claims.aggregate_claims_data(inputs)
    formatted = bc_claims.format_data_for_template(combined)

    claims_output_dir = session_dir / "output"
    claims_output_dir.mkdir(exist_ok=True)

    schedule_bytes = bc_claims.create_schedule_a_document(bc_claims.SCHEDULE_TEMPLATE, formatted)
    notice_bytes = bc_claims.create_notice_of_claim_pdf(bc_claims.NOTICE_TEMPLATE, formatted)

    defendant = bc_claims.safe_filename(formatted.get("SIMPLE NAME INSERT", "Defendant"))

    schedule_path = claims_output_dir / f"{defendant} - Schedule A.docx"
    notice_path = claims_output_dir / f"{defendant} - Notice of Claim.pdf"

    schedule_path.write_bytes(schedule_bytes)
    notice_path.write_bytes(notice_bytes)

    return [schedule_path, notice_path]


def process_on_claims(session_dir: Path, uploads: List[Path]) -> List[Path]:
    metadata: Dict[str, str] = {}
    schedule_template: Optional[Path] = None
    claim_template: Optional[Path] = None
    docx_candidates: List[Path] = []
    pdf_candidates: List[Path] = []

    for candidate in uploads:
        suffix = candidate.suffix.lower()
        if suffix == ".json":
            raw_data = json.loads(candidate.read_text(encoding="utf-8"))
            for key, value in raw_data.items():
                metadata[str(key)] = str(value)
        elif suffix == ".docx":
            docx_candidates.append(candidate)
        elif suffix == ".pdf":
            pdf_candidates.append(candidate)

    for docx_path in docx_candidates:
        identified = ontario_claims.identify_template(docx_path)
        upper_name = docx_path.name.upper()
        if identified == "SCHEDULE" or ("SCHEDULE" in upper_name and schedule_template is None):
            if schedule_template is None:
                schedule_template = docx_path
            continue
        if identified == "CLAIM" or (
            ("CLAIM" in upper_name or "7A" in upper_name) and claim_template is None
        ):
            if claim_template is None:
                claim_template = docx_path
            continue
    if schedule_template is None and docx_candidates:
        schedule_template = docx_candidates[0]
    if claim_template is None and len(docx_candidates) > 1:
        for candidate in docx_candidates:
            if candidate is not schedule_template:
                claim_template = candidate
                break

    if schedule_template is None or claim_template is None:
        raise ValueError("Upload both Schedule A and ON Form 7A templates (DOCX).")

    files_by_type: Dict[str, Path] = {}
    leftovers: List[Path] = []
    for pdf_path in pdf_candidates:
        name = pdf_path.name.upper()
        if "MRP" in name and "MRP" not in files_by_type:
            files_by_type["MRP"] = pdf_path
        elif "MRC" in name and "MRC" not in files_by_type:
            files_by_type["MRC"] = pdf_path
        elif "MRS" in name and "MRS" not in files_by_type:
            files_by_type["MRS"] = pdf_path
        elif ("CBR" in name or "CREDIT REPORT" in name) and "CBR" not in files_by_type:
            files_by_type["CBR"] = pdf_path
        else:
            leftovers.append(pdf_path)

    for pdf_path in leftovers:
        text, _ = ontario_claims.load_pdf_text_and_lines(pdf_path)
        lowered = text.lower()
        if "new payments" in lowered and "total of payment activity" in lowered and "MRP" not in files_by_type:
            files_by_type["MRP"] = pdf_path
        elif "new transactions for" in lowered and "total of new transactions" in lowered and "MRC" not in files_by_type:
            files_by_type["MRC"] = pdf_path
        elif "statement of account" in lowered and "new balance" in lowered and "MRS" not in files_by_type:
            files_by_type["MRS"] = pdf_path
        elif "this completes the file for" in lowered and "CBR" not in files_by_type:
            files_by_type["CBR"] = pdf_path

    missing = [key for key in ("MRP", "MRC", "MRS", "CBR") if key not in files_by_type]
    if missing:
        raise ValueError(
            f"Missing required files for ON Claims: {', '.join(missing)}. "
            "Ensure filenames include MRP, MRC, MRS, and Credit Report."
        )

    outputs = ontario_claims.generate_claim_documents(
        mrs_path=files_by_type["MRS"],
        mrc_path=files_by_type["MRC"],
        mrp_path=files_by_type["MRP"],
        cbr_path=files_by_type["CBR"],
        schedule_template=schedule_template,
        claim_template=claim_template,
        demand_letter_date=metadata.get("demand_letter_date"),
        claim_prepared_date=metadata.get("claim_prepared_date"),
    )

    output_dir = session_dir / "output"
    output_dir.mkdir(exist_ok=True)
    result_paths: List[Path] = []
    for key in ("schedule", "claim"):
        filename, binary = outputs[key]
        dest = output_dir / filename
        dest.write_bytes(binary)
        result_paths.append(dest)
    return result_paths


STEPS: Dict[str, StepConfig] = {
    "demand_letter": StepConfig(
        label="BC Demand Letter",
        description="Generate a BC demand letter from the intake document.",
        file_hint="Upload the filled demand letter intake DOCX (or the original statement PDF).",
        expected_files=1,
        allowed_extensions=[".docx", ".pdf"],
        processor=process_demand_letter,
        dropzones=[
            {
                "id": "statement",
                "label": "Demand Letter Intake",
                "hint": "Upload the intake Word document or statement PDF for this debtor.",
                "accept": [".docx", ".pdf"],
                "max": 1,
            }
        ],
    ),
    "on_claims": StepConfig(
        label="ON Claims",
        description="Generate Ontario Schedule A and Plaintiff's Claim from four statements.",
        file_hint="Upload the four PDFs (MRP, MRC, MRS, Credit Report) and the two Ontario templates.",
        expected_files=6,
        allowed_extensions=[".pdf", ".docx"],
        processor=process_on_claims,
        dropzones=[
            {
                "id": "on-mrp",
                "label": "MRP PDF (Monthly Payment Report)",
                "hint": "Upload the Ontario MRP statement.",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "on-mrc",
                "label": "MRC PDF (Monthly Recap Report)",
                "hint": "Upload the Ontario MRC statement.",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "on-mrs",
                "label": "MRS PDF (Most Recent Statement)",
                "hint": "Upload the Ontario MRS statement.",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "on-cbr",
                "label": "Credit Report PDF",
                "hint": "Upload the TransUnion credit report.",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "on-schedule",
                "label": "Ontario Schedule A Template",
                "hint": "Upload the ON Schedule A DOCX template.",
                "accept": [".docx"],
                "max": 1,
            },
            {
                "id": "on-claim",
                "label": "Ontario Form 7A Template",
                "hint": "Upload the ON Plaintiff's Claim (Form 7A) DOCX template.",
                "accept": [".docx"],
                "max": 1,
            },
        ],
    ),
    "bc_claims": StepConfig(
        label="BC Claims",
        description="Prepare Schedule A and Notice of Claim from four monthly reports.",
        file_hint="Upload the four PDFs (MRP, MRC, MRS, and Credit Report).",
        expected_files=4,
        allowed_extensions=[".pdf"],
        processor=process_bc_claims,
        dropzones=[
            {
                "id": "mrp",
                "label": "MRP PDF (Monthly Payment Report)",
                "hint": "Upload the MRP file (payments).",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "mrc",
                "label": "MRC PDF (Monthly Charge Report)",
                "hint": "Upload the MRC file (charges).",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "mrs",
                "label": "MRS PDF (Monthly Statement)",
                "hint": "Upload the MRS statement PDF.",
                "accept": [".pdf"],
                "max": 1,
            },
            {
                "id": "cbr",
                "label": "Credit Bureau Report (CBR)",
                "hint": "Upload the credit bureau report PDF.",
                "accept": [".pdf"],
                "max": 1,
            },
        ],
    ),
    "default_judgment": StepConfig(
        label="BC Default Judgment",
        description="Populate the Application for Default Order using Schedule A and Notice of Claim.",
        file_hint="Upload the filled Schedule A DOCX and the filed Notice of Claim PDF.",
        expected_files=2,
        allowed_extensions=[".docx", ".pdf"],
        processor=None,  # placeholder
        dropzones=[
            {
                "id": "schedule",
                "label": "Schedule A (DOCX)",
                "hint": "Upload the filled Schedule A Word document.",
                "accept": [".docx"],
                "max": 1,
            },
            {
                "id": "notice",
                "label": "Notice of Claim (PDF)",
                "hint": "Upload the filed Notice of Claim PDF.",
                "accept": [".pdf"],
                "max": 1,
            },
        ],
    ),
    "dismissal": StepConfig(
        label="BC Dismissal",
        description="Generate a Notice of Withdrawal using Schedule A and Notice of Claim.",
        file_hint="Upload the filled Schedule A DOCX and the filed Notice of Claim PDF.",
        expected_files=2,
        allowed_extensions=[".docx", ".pdf"],
        processor=None,  # placeholder
        dropzones=[
            {
                "id": "schedule",
                "label": "Schedule A (DOCX)",
                "hint": "Upload the filled Schedule A Word document.",
                "accept": [".docx"],
                "max": 1,
            },
            {
                "id": "notice",
                "label": "Notice of Claim (PDF)",
                "hint": "Upload the filed Notice of Claim PDF.",
                "accept": [".pdf"],
                "max": 1,
            },
        ],
    ),
}

app = Flask(__name__)


def process_default_judgment(session_dir: Path, uploads: List[Path]) -> List[Path]:
    docx_files = [p for p in uploads if p.suffix.lower() == ".docx"]
    pdf_files = [p for p in uploads if p.suffix.lower() == ".pdf"]
    if len(docx_files) != 1 or len(pdf_files) != 1:
        raise ValueError("Upload exactly one DOCX (Schedule A) and one PDF (Notice of Claim).")
    output_dir = session_dir / "output"
    output_dir.mkdir(exist_ok=True)
    out_path, _summary = default_judgment.fill_default_order(
        claim_docx=docx_files[0],
        notice_pdf=pdf_files[0],
        template_pdf=default_judgment.ASSET_TEMPLATE,
        output_dir=output_dir,
        registry_location_override="",
    )
    return [out_path]


def process_dismissal(session_dir: Path, uploads: List[Path]) -> List[Path]:
    docx_files = [p for p in uploads if p.suffix.lower() == ".docx"]
    pdf_files = [p for p in uploads if p.suffix.lower() == ".pdf"]
    if len(docx_files) != 1 or len(pdf_files) != 1:
        raise ValueError("Upload exactly one DOCX (Schedule A) and one PDF (Notice of Claim).")
    output_dir = session_dir / "output"
    output_dir.mkdir(exist_ok=True)
    out_path, _summary = dismissal.fill_dismissal_form(
        claim_docx=docx_files[0],
        notice_pdf=pdf_files[0],
        template_pdf=dismissal.ASSET_TEMPLATE,
        output_dir=output_dir,
    )
    return [out_path]


STEPS["default_judgment"].processor = process_default_judgment
STEPS["dismissal"].processor = process_dismissal


@app.route("/")
def index():
    step_key = request.args.get("step")
    if step_key and step_key in STEPS:
        return render_template("upload.html", step_key=step_key, step=STEPS[step_key])
    return render_template("index.html", steps=STEPS, year=datetime.now().year)


def _validate_files(step: StepConfig, files) -> List[Path]:
    saved_paths: List[Path] = []
    min_required = step.expected_files
    max_allowed = step.expected_files + step.optional_max
    if len(files) < min_required or len(files) > max_allowed:
        if step.optional_max:
            msg = (
                f"Expected between {min_required} and {max_allowed} files for {step.label}, "
                f"received {len(files)}."
            )
        else:
            msg = (
                f"Expected {min_required} file(s) for {step.label}, "
                f"received {len(files)}."
            )
        raise ValueError(msg)
    session_id = uuid.uuid4().hex
    session_dir = SESSIONS_DIR / session_id
    uploads_dir = session_dir / "uploads"
    uploads_dir.mkdir(parents=True, exist_ok=True)
    required_count = 0
    optional_count = 0
    for storage in files:
        if not storage or not storage.filename:
            raise ValueError("Missing file upload.")
        filename = storage.filename
        suffix = Path(filename).suffix.lower()
        if suffix in step.allowed_extensions:
            required_count += 1
        elif suffix in step.optional_extensions and optional_count < step.optional_max:
            optional_count += 1
        else:
            allowed = ", ".join(step.allowed_extensions + step.optional_extensions)
            raise ValueError(
                f"Invalid file type for {filename}. Allowed: {allowed or 'none'}"
            )
        dest = uploads_dir / filename
        storage.save(dest)
        saved_paths.append(dest)
    if required_count != step.expected_files:
        raise ValueError(
            f"Expected {step.expected_files} required file(s) for {step.label}, "
            f"received {required_count}."
        )
    return saved_paths


@app.route("/process/<step_key>", methods=["POST"])
def process_step(step_key: str):
    if step_key not in STEPS:
        return redirect(url_for("index"))

    step = STEPS[step_key]
    try:
        files = request.files.getlist("files")
        saved_paths = _validate_files(step, files)
        session_dir = saved_paths[0].parent.parent
        outputs = step.processor(session_dir, saved_paths)
        download_links = [
            url_for("download_file", session_id=session_dir.name, filename=path.name)
            for path in outputs
        ]
        return render_template(
            "result.html",
            step=step,
            output_files=zip(outputs, download_links),
        )
    except Exception as exc:  # pragma: no cover - defensive
        return render_template(
            "upload.html",
            step_key=step_key,
            step=step,
            error=str(exc),
        )


@app.route("/download/<session_id>/<path:filename>")
def download_file(session_id: str, filename: str):
    session_dir = SESSIONS_DIR / session_id / "output"
    if not session_dir.exists():
        return redirect(url_for("index"))
    return send_from_directory(session_dir, filename, as_attachment=True)


if __name__ == "__main__":  # pragma: no cover
    app.run(debug=True, host="0.0.0.0", port=5001)
