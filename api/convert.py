"""Vercel serverless endpoint: POST /api/convert (multipart, field "file")."""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from flask import Flask, request, send_file, jsonify  # noqa: E402

import convert as core  # noqa: E402

app = Flask(__name__)
MAX_BYTES = 4 * 1024 * 1024  # Vercel serverless request body limit is ~4.5 MB


@app.errorhandler(Exception)
def on_error(exc):
    return jsonify(error=f"{type(exc).__name__}: {exc}"), 500


@app.post("/api/convert")
def convert_endpoint():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="No file uploaded"), 400

    name = Path(f.filename).name
    suffix = Path(name).suffix.lower()
    if suffix not in (".pdf", ".xls"):
        return jsonify(error="Only .pdf and .xls files are supported"), 400

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"upload{suffix}"
        f.save(src)
        if src.stat().st_size > MAX_BYTES:
            return jsonify(error="File too large (max 4 MB)"), 413

        out_dir = Path(tmp) / "out"
        out_dir.mkdir()
        if suffix == ".pdf":
            status = core.convert_pdf_to_excel(src, out_dir)
        else:
            status = core.convert_xls_to_xlsx(src, out_dir)

        result = out_dir / "upload.xlsx"
        if not result.exists():
            return jsonify(error=status), 422

        data = result.read_bytes()

    import io
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=Path(name).stem + ".xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/merge")
def merge_endpoint():
    """Stack all sheets of all uploaded .xlsx files into one table."""
    import io
    import pandas as pd

    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No files uploaded"), 400

    frames = []
    multi_sheet = False
    for f in files:
        xls = pd.ExcelFile(io.BytesIO(f.read()), engine="openpyxl")
        multi_sheet = multi_sheet or len(xls.sheet_names) > 1
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df.insert(0, "source_sheet", sheet)
            df.insert(0, "source_file", Path(f.filename).name)
            frames.append(df)

    merged = pd.concat(frames, ignore_index=True, sort=False)
    if not multi_sheet:
        merged = merged.drop(columns="source_sheet")

    out = io.BytesIO()
    merged.to_excel(out, index=False, sheet_name="Merged")
    out.seek(0)
    return send_file(
        out,
        as_attachment=True,
        download_name="merged.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
