from pathlib import Path
from uuid import uuid4

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from time_columns_to_rows import transpose_time_columns

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}

app = Flask(__name__)
app.secret_key = "time-columns-to-rows-local-app"
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def safe_display_columns(columns, limit=12):
    values = [str(column) for column in columns]
    if len(values) <= limit:
        return "、".join(values)
    return "、".join(values[:limit]) + f" 等 {len(values)} 个"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    upload = request.files.get("file")
    if not upload or upload.filename == "":
        flash("请先选择一个表格文件。")
        return redirect(url_for("index"))

    if not allowed_file(upload.filename):
        flash("文件格式不支持。请上传 xlsx、xls、xlsm、csv 或 tsv。")
        return redirect(url_for("index"))

    original_name = Path(upload.filename).name
    suffix = Path(original_name).suffix.lower()
    safe_name = secure_filename(original_name) or f"table{suffix}"
    token = uuid4().hex
    input_path = UPLOAD_DIR / f"{token}_{safe_name}"
    output_path = OUTPUT_DIR / f"{Path(safe_name).stem}_时间列转行.xlsx"
    upload.save(input_path)

    sheet_name = request.form.get("sheet_name", "").strip() or None
    time_name = request.form.get("time_name", "").strip() or "时间"
    value_name = request.form.get("value_name", "").strip() or "值"

    try:
        result_path, time_columns, row_count = transpose_time_columns(
            input_path,
            output_path,
            sheet_name=sheet_name,
            time_column_name=time_name,
            value_column_name=value_name,
        )
    except Exception as error:
        flash(str(error))
        return redirect(url_for("index"))

    return render_template(
        "result.html",
        download_id=result_path.name,
        original_name=original_name,
        time_columns=safe_display_columns(time_columns),
        time_column_count=len(time_columns),
        row_count=row_count,
    )


@app.route("/download/<path:filename>")
def download(filename):
    file_path = OUTPUT_DIR / Path(filename).name
    if not file_path.exists():
        flash("下载文件不存在，请重新上传转换。")
        return redirect(url_for("index"))
    return send_file(file_path, as_attachment=True, download_name=file_path.name)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
