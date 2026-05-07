from __future__ import annotations

import json
import sys
import uuid
import zipfile
from pathlib import Path

from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from scripts.build_rms_import_master import build_master, write_outputs  # noqa: E402


APP_DATA = ROOT / "web_runs"
UPLOAD_DIR = APP_DATA / "uploads"
OUTPUT_DIR = APP_DATA / "outputs"
MAX_UPLOAD_MB = 250

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


def save_required_file(run_upload_dir: Path, field_name: str) -> Path:
    file = request.files.get(field_name)
    if file is None or not file.filename:
        raise ValueError(f"Missing required upload: {field_name}")
    filename = secure_filename(file.filename)
    if not filename:
        raise ValueError(f"Invalid filename for: {field_name}")
    path = run_upload_dir / filename
    file.save(path)
    return path


def save_optional_file(run_upload_dir: Path, field_name: str) -> Path | None:
    file = request.files.get(field_name)
    if file is None or not file.filename:
        return None
    filename = secure_filename(file.filename)
    if not filename:
        raise ValueError(f"Invalid filename for: {field_name}")
    path = run_upload_dir / filename
    file.save(path)
    return path


def save_siteminder_files(run_upload_dir: Path) -> list[Path]:
    files = request.files.getlist("siteminder_files")
    paths: list[Path] = []
    for file in files:
        if not file or not file.filename:
            continue
        filename = secure_filename(file.filename)
        if not filename:
            continue
        path = run_upload_dir / filename
        file.save(path)
        paths.append(path)
    if not paths:
        raise ValueError("Missing required upload: siteminder_files")
    return paths


def make_zip(run_id: str, out_dir: Path) -> Path:
    zip_path = out_dir / f"rms_import_results_{run_id}.zip"
    names = [
        "rms_absolute_master_merged.xlsx",
        "cancelled_bookings.xlsx",
        "status_review.xlsx",
    ]
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            path = out_dir / name
            if path.exists():
                zf.write(path, arcname=name)
    return zip_path


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/process")
def process_uploads():
    run_id = uuid.uuid4().hex[:12]
    run_upload_dir = UPLOAD_DIR / run_id
    run_output_dir = OUTPUT_DIR / run_id
    run_upload_dir.mkdir(parents=True, exist_ok=True)
    run_output_dir.mkdir(parents=True, exist_ok=True)

    try:
        master_file = save_required_file(run_upload_dir, "master_file")
        arrival_file = save_required_file(run_upload_dir, "arrival_file")
        expedia_file = save_required_file(run_upload_dir, "expedia_file")
        asi_booking_report_file = save_optional_file(run_upload_dir, "asi_booking_report_file")
        siteminder_files = save_siteminder_files(run_upload_dir)

        output, summary, audit = build_master(
            master_file=master_file,
            arrival_file=arrival_file,
            expedia_file=expedia_file,
            siteminder_files=siteminder_files,
            asi_booking_report_file=asi_booking_report_file,
        )
        write_outputs(output, summary, audit, run_output_dir)
        make_zip(run_id, run_output_dir)
    except Exception as exc:
        return render_template("error.html", error=str(exc)), 400

    return redirect(url_for("result", run_id=run_id))


@app.get("/result/<run_id>")
def result(run_id: str):
    out_dir = OUTPUT_DIR / run_id
    summary_path = out_dir / "rms_absolute_master_summary.json"
    if not summary_path.exists():
        abort(404)
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    return render_template(
        "result.html",
        run_id=run_id,
        summary=summary,
        zip_filename=f"rms_import_results_{run_id}.zip",
    )


@app.get("/download/<run_id>/<path:filename>")
def download(run_id: str, filename: str):
    out_dir = OUTPUT_DIR / run_id
    path = (out_dir / filename).resolve()
    if out_dir.resolve() not in path.parents or not path.exists():
        abort(404)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
