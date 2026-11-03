"""Vercel serverless endpoint: POST /api/convert (multipart, field "file")."""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from flask import Flask, request, send_file, jsonify  # noqa: E402

import convert as core  # noqa: E402

app = Flask(__name__)
MAX_BYTES = 4 * 1024 * 1024  # Vercel serverless request body limit is ~4.5 MB


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
