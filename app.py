import io
import json
import os
import re
from collections import Counter
from flask import Flask, request, render_template, send_file, jsonify
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

PATTERNS = [
    ("Khoảng trắng lặp", re.compile(r"(?<=\\S)[ \\t]{2,}"), " "),
    ("Dấu cách trước dấu câu", re.compile(r"[ \\t]+(?=[,:;!?])"), ""),
    ("Dấu cách trước dấu chấm", re.compile(r"[ \\t]+(?=\\.(?!\\.))"), ""),
    ("Thiếu khoảng trắng sau dấu phẩy", re.compile(r",(?=[^\\s\\d\\\"”’.,])"), ", "),
    ("Ký hiệu ^+^+ dính chữ", re.compile(r"(?<=[^\\W\\d_])(?=\\^\\+\\^\\+)"), " "),
]

def fix_text(text, counts):
    for name, pattern, replacement in PATTERNS:
        text, changed = pattern.subn(replacement, text)
        counts[name] += changed
    return text

def is_heading(text):
    title = text.strip()
    return bool(re.match(r"^(?:Chương|Chapter)\\s+\\d+(?:\\s*[:：.\\-–].*)?$", title, re.I)) or title.casefold() in ("văn án", "giới thiệu", "mục lục")

def iter_paragraphs(doc):
    for paragraph in doc.paragraphs:
        yield paragraph
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph

def clean(data, options):
    doc = Document(io.BytesIO(data))
    counts = Counter()
    if options.get("narrow"):
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(1.27)
    for paragraph in iter_paragraphs(doc):
        original = paragraph.text
        local = Counter()
        cleaned = fix_text(original, local)
        if cleaned != original:
            if len(paragraph.runs) == 1:
                paragraph.runs[0].text = cleaned
                counts.update(local)
            elif not paragraph.runs:
                paragraph.add_run(cleaned)
                counts.update(local)
            else:
                # Preserve text formatting when a paragraph contains multiple styled runs.
                counts["Đoạn nhiều định dạng cần kiểm tra"] += 1
        if options.get("headings") and is_heading(paragraph.text):
            paragraph.style = "Heading 2"
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in paragraph.runs:
                run.font.name = "Tahoma"
                run.font.size = Pt(11)
            counts["Chuẩn hóa Heading 2"] += 1
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, dict(counts)

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/api/clean")
def api_clean():
    upload = request.files.get("file")
    if not upload or not upload.filename.lower().endswith(".docx"):
        return jsonify(error="Vui lòng tải file .docx"), 400
    options = {
        "narrow": request.form.get("narrow") == "true",
        "headings": request.form.get("headings") == "true"
    }
    try:
        output, counts = clean(upload.read(), options)
    except Exception:
        return jsonify(error="Không đọc được file Word hợp lệ."), 400
    original_filename = os.path.splitext(secure_filename(upload.filename))[0] or "document"
    response = send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=original_filename + "_da_sua.docx",
    )
    response.headers["X-Corrections"] = json.dumps(counts, ensure_ascii=True)
    response.headers["Access-Control-Expose-Headers"] = "X-Corrections"
    return response

@app.get("/health")
def health():
    return {"ok": True}

if __name__ == "__main__":
    app.run(debug=True)
