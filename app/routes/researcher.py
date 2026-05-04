"""
Routes for the Researcher module.

Endpoints
─────────
GET  /researcher/           – upload form
POST /researcher/process    – upload & parse syllabus
GET  /researcher/export     – download results as CSV or XLSX
POST /researcher/export     – same, triggered from results page
"""

from __future__ import annotations

import os
import uuid

from flask import (
    Blueprint,
    current_app,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
import io

from app.services.parser import parse_syllabus_bytes
from app.services.exporter import to_csv, to_xlsx, entries_to_dataframe, COLUMNS

researcher_bp = Blueprint("researcher", __name__)

ALLOWED = {"pdf", "docx", "doc", "txt"}


def _allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED


def _save_entries_to_session(entries) -> None:
    """Persist parsed entries in the server-side session (as list of dicts)."""
    session["researcher_entries"] = [e.as_dict() for e in entries]
    session["researcher_count"] = len(entries)


def _load_entries_from_session():
    """Reload CaseEntry-like dicts from session."""
    return session.get("researcher_entries", [])


# ── views ────────────────────────────────────────────────────────────────────

@researcher_bp.route("/", methods=["GET"])
def index():
    return render_template("researcher/index.html")


@researcher_bp.route("/process", methods=["POST"])
def process():
    if "syllabus" not in request.files:
        flash("No file part in the request.", "danger")
        return redirect(url_for("researcher.index"))

    file = request.files["syllabus"]
    if file.filename == "":
        flash("No file selected.", "danger")
        return redirect(url_for("researcher.index"))

    if not _allowed(file.filename):
        flash(
            "Unsupported file type. Please upload a PDF, DOCX, DOC, or TXT file.",
            "danger",
        )
        return redirect(url_for("researcher.index"))

    data = file.read()
    try:
        entries = parse_syllabus_bytes(data, file.filename)
    except Exception as exc:
        flash(f"Error parsing file: {exc}", "danger")
        return redirect(url_for("researcher.index"))

    _save_entries_to_session(entries)

    if not entries:
        flash(
            "No case citations were found in the uploaded syllabus. "
            "Please check the file format.",
            "warning",
        )

    return redirect(url_for("researcher.results"))


@researcher_bp.route("/results", methods=["GET"])
def results():
    rows = _load_entries_from_session()
    count = session.get("researcher_count", 0)
    return render_template(
        "researcher/results.html",
        rows=rows,
        count=count,
        columns=COLUMNS,
    )


@researcher_bp.route("/export", methods=["GET", "POST"])
def export():
    fmt = request.values.get("format", "csv").lower()
    rows = _load_entries_from_session()

    # Reconstruct minimal entry-like objects (dicts are fine for exporters)
    from app.services.parser import CaseEntry

    entries = [
        CaseEntry(
            case=r.get("Case", ""),
            topic=r.get("Topic", ""),
            subtopic=r.get("Subtopic", ""),
            sub_subtopic=r.get("Sub-subtopic", ""),
            further_division=r.get("Further Division", ""),
        )
        for r in rows
    ]

    if fmt == "xlsx":
        data = to_xlsx(entries)
        mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = "syllabus_cases.xlsx"
    else:
        data = to_csv(entries)
        mimetype = "text/csv"
        filename = "syllabus_cases.csv"

    return send_file(
        io.BytesIO(data),
        mimetype=mimetype,
        as_attachment=True,
        download_name=filename,
    )
