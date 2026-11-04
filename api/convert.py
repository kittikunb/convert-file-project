"""Vercel serverless endpoint: POST /api/convert (multipart, field "file")."""
import io
import json
import re
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


def convert_csv(src: Path, out_dir: Path) -> str:
    """CSV -> xlsx. Tries UTF-8 (with BOM), then Thai Windows-874 encoding."""
    import pandas as pd

    last = None
    for enc in ("utf-8-sig", "cp874"):
        try:
            df = pd.read_csv(src, sep=None, engine="python", encoding=enc)
            break
        except UnicodeDecodeError as exc:
            last = exc
        except Exception as exc:
            return f"ERROR: {exc}"
    else:
        return f"ERROR: unsupported encoding ({last})"
    out = out_dir / (src.stem + ".xlsx")
    df.to_excel(out, index=False)
    return f"ok ({len(df)} row(s))"


@app.post("/api/convert")
def convert_endpoint():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="No file uploaded"), 400

    name = Path(f.filename).name
    suffix = Path(name).suffix.lower()
    if suffix not in (".pdf", ".xls", ".csv"):
        return jsonify(error="Only .pdf, .xls and .csv files are supported"), 400

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"upload{suffix}"
        f.save(src)
        if src.stat().st_size > MAX_BYTES:
            return jsonify(error="File too large (max 4 MB)"), 413

        out_dir = Path(tmp) / "out"
        out_dir.mkdir()
        if suffix == ".pdf":
            status = core.convert_pdf_to_excel(src, out_dir)
        elif suffix == ".csv":
            status = convert_csv(src, out_dir)
        else:
            status = core.convert_xls_to_xlsx(src, out_dir)

        result = out_dir / "upload.xlsx"
        if not result.exists():
            return jsonify(error=status), 422

        data = result.read_bytes()

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=Path(name).stem + ".xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


DATE_RE = re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)")


def build_merged(files):
    import pandas as pd

    frames = []
    multi_sheet = False
    for f in files:
        name = Path(f.filename).name
        m = DATE_RE.search(name)
        file_date = pd.Timestamp(f"{m.group(1)}-{m.group(2)}-{m.group(3)}") if m else None
        xls = pd.ExcelFile(io.BytesIO(f.read()), engine="openpyxl")
        multi_sheet = multi_sheet or len(xls.sheet_names) > 1
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df.insert(0, "source_sheet", sheet)
            if file_date is not None:
                df.insert(0, "file_date", file_date)
            df.insert(0, "source_file", name)
            frames.append(df)

    merged = pd.concat(frames, ignore_index=True, sort=False)
    if not multi_sheet:
        merged = merged.drop(columns="source_sheet")
    return merged


@app.post("/api/merge")
def merge_endpoint():
    """Stack all sheets of all uploaded .xlsx files into one table.

    With ?preview=1 returns the first rows as JSON instead of the file.
    """
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No files uploaded"), 400

    merged = build_merged(files)

    if request.args.get("preview"):
        head = merged.head(20)
        return app.response_class(
            json.dumps({
                "total": len(merged),
                "columns": [str(c) for c in head.columns],
                "rows": json.loads(head.to_json(orient="values", date_format="iso", date_unit="s")),
            }),
            mimetype="application/json",
        )

    if request.args.get("format") == "csv":
        if "file_date" in merged.columns:
            merged["file_date"] = merged["file_date"].dt.strftime("%Y-%m-%d")
        # utf-8-sig so Excel opens Thai text correctly
        data = merged.to_csv(index=False).encode("utf-8-sig")
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name="merged.csv",
            mimetype="text/csv",
        )

    out = io.BytesIO()
    merged.to_excel(out, index=False, sheet_name="Merged")
    out.seek(0)
    return send_file(
        out,
        as_attachment=True,
        download_name="merged.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
