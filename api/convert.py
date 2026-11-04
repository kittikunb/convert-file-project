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


@app.after_request
def no_store(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp


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
    if suffix not in (".pdf", ".xls", ".csv", ".xlsx"):
        return jsonify(error="Only .pdf, .xls, .xlsx and .csv files are supported"), 400

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
        elif suffix == ".xlsx":
            import shutil
            shutil.copy(src, out_dir / "upload.xlsx")
            status = "ok"
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


def read_csv_df(data: bytes):
    import pandas as pd

    last = None
    for enc in ("utf-8-sig", "cp874"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc)
        except UnicodeDecodeError as exc:
            last = exc
    raise ValueError(f"unsupported encoding ({last})")


def stack(entries):
    """entries: [(file_name, {sheet_name: DataFrame})] -> one stacked DataFrame."""
    import pandas as pd

    frames = []
    multi_sheet = False
    for name, sheets in entries:
        m = DATE_RE.search(name)
        file_date = pd.Timestamp(f"{m.group(1)}-{m.group(2)}-{m.group(3)}") if m else None
        multi_sheet = multi_sheet or len(sheets) > 1
        for sheet, df in sheets.items():
            df = df.copy()
            df.insert(0, "source_sheet", sheet)
            if file_date is not None:
                df.insert(0, "file_date", file_date)
            df.insert(0, "source_file", name)
            frames.append(df)

    merged = pd.concat(frames, ignore_index=True, sort=False)
    if not multi_sheet:
        merged = merged.drop(columns="source_sheet")
    return merged


def build_merged(files):
    """Merge uploaded .xlsx files."""
    import pandas as pd

    entries = []
    for f in files:
        xls = pd.ExcelFile(io.BytesIO(f.read()), engine="openpyxl")
        entries.append((Path(f.filename).name, {sh: xls.parse(sh) for sh in xls.sheet_names}))
    return stack(entries)


def csv_response(merged):
    if "file_date" in merged.columns:
        merged["file_date"] = merged["file_date"].dt.strftime("%Y-%m-%d")
    # utf-8-sig so Excel opens Thai text correctly
    data = merged.to_csv(index=False).encode("utf-8-sig")
    return send_file(io.BytesIO(data), as_attachment=True, download_name="merged.csv", mimetype="text/csv")


def xlsx_response(merged):
    out = io.BytesIO()
    merged.to_excel(out, index=False, sheet_name="Merged")
    out.seek(0)
    return send_file(
        out,
        as_attachment=True,
        download_name="merged.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/merge-original")
def merge_original_endpoint():
    """Merge original files of ONE type, keeping the extension.

    csv -> merged.csv, xlsx -> merged.xlsx, pdf -> merged.pdf (pages appended),
    xls -> converted to xlsx first, result is merged.xlsx.
    """
    import pandas as pd

    files = request.files.getlist("files")
    if len(files) < 2:
        return jsonify(error="Need at least 2 files"), 400

    kinds = {Path(f.filename).suffix.lower() for f in files}
    if len(kinds) != 1:
        return jsonify(error="ต้องเป็นไฟล์ชนิดเดียวกันเท่านั้น (พบ: " + ", ".join(sorted(kinds)) + ")"), 400
    kind = kinds.pop()

    if kind == ".csv":
        entries = [(Path(f.filename).name, {"Sheet1": read_csv_df(f.read())}) for f in files]
        return csv_response(stack(entries))

    if kind == ".xlsx":
        return xlsx_response(build_merged(files))

    if kind == ".xls":
        entries = []
        with tempfile.TemporaryDirectory() as tmp:
            for i, f in enumerate(files):
                src = Path(tmp) / f"in{i}.xls"
                f.save(src)
                xls = pd.ExcelFile(src, engine="xlrd")
                entries.append((Path(f.filename).name, {sh: xls.parse(sh) for sh in xls.sheet_names}))
        return xlsx_response(stack(entries))

    if kind == ".pdf":
        from pypdf import PdfWriter

        writer = PdfWriter()
        for f in files:
            writer.append(io.BytesIO(f.read()))
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return send_file(out, as_attachment=True, download_name="merged.pdf", mimetype="application/pdf")

    return jsonify(error="Unsupported file type"), 400


@app.post("/api/merge")
def merge_endpoint():
    """Stack all sheets of all uploaded .xlsx files into one table."""
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No files uploaded"), 400

    merged = build_merged(files)

    if request.args.get("format") == "csv":
        return csv_response(merged)
    return xlsx_response(merged)
