import os
import uuid
import threading
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
app.config["DOWNLOAD_FOLDER"] = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")

ALLOWED_EXTENSIONS = {"pdf", "docx", "doc"}

# In-memory job tracking
jobs = {}
jobs_lock = threading.Lock()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def pdf_to_word(input_path, output_path):
    """Convert PDF to DOCX. Uses Microsoft Word's engine for best fidelity,
    falls back to pdf2docx if Word is unavailable."""
    # Try Microsoft Word COM first (best quality for Chinese documents)
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            from win32com import client
            word = client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(input_path)
                doc.SaveAs2(output_path, FileFormat=16)  # 16 = wdFormatDocx
                doc.Close()
                return
            finally:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        pass

    # Fallback: pdf2docx
    from pdf2docx import Converter
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()


def word_to_pdf_libreoffice(input_path, output_dir):
    """Convert Word to PDF using LibreOffice (cross-platform)."""
    import subprocess
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_path],
        check=True, capture_output=True, timeout=120
    )
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(output_dir, f"{base}.pdf")


def word_to_pdf_win32(input_path, output_path):
    """Convert Word to PDF using Microsoft Word COM (Windows only)."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        from win32com import client
        word = client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(input_path)
            doc.ExportAsFixedFormat(output_path, 17)  # 17 = wdFormatPDF
            doc.Close()
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()


def word_to_pdf(input_path, output_path):
    """Try win32com first (Windows), fall back to LibreOffice."""
    try:
        word_to_pdf_win32(input_path, output_path)
        return
    except Exception:
        pass
    output_dir = os.path.dirname(output_path)
    result = word_to_pdf_libreoffice(input_path, output_dir)
    base = os.path.splitext(os.path.basename(input_path))[0]
    generated = os.path.join(output_dir, f"{base}.pdf")
    if generated != output_path:
        os.rename(generated, output_path)


def process_job(job_id, filename, direction):
    try:
        with jobs_lock:
            jobs[job_id]["status"] = "processing"

        ext = filename.rsplit(".", 1)[1].lower()
        base = filename.rsplit(".", 1)[0]
        input_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{job_id}_{filename}")
        download_filename = f"{base}.docx" if direction == "pdf2word" else f"{base}.pdf"
        output_path = os.path.join(app.config["DOWNLOAD_FOLDER"], f"{job_id}_{download_filename}")

        if direction == "pdf2word":
            pdf_to_word(input_path, output_path)
        else:
            word_to_pdf(input_path, output_path)

        with jobs_lock:
            jobs[job_id]["status"] = "done"
            jobs[job_id]["download"] = output_path
            jobs[job_id]["download_filename"] = download_filename
    except Exception as e:
        with jobs_lock:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["error"] = str(e)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/network-info")
def network_info():
    import socket
    local_ip = "127.0.0.1"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        pass
    return jsonify({"local_ip": local_ip})


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "未找到文件"}), 400

    file = request.files["file"]
    direction = request.form.get("direction", "pdf2word")

    if file.filename == "":
        return jsonify({"error": "未选择文件"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "仅支持 PDF、DOCX、DOC 格式"}), 400

    ext = file.filename.rsplit(".", 1)[1].lower()
    if direction == "pdf2word" and ext != "pdf":
        return jsonify({"error": "PDF 转 Word 需要上传 .pdf 文件"}), 400
    if direction == "word2pdf" and ext not in ("docx", "doc"):
        return jsonify({"error": "Word 转 PDF 需要上传 .docx 或 .doc 文件"}), 400

    job_id = str(uuid.uuid4())
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{job_id}_{filename}")
    file.save(input_path)

    with jobs_lock:
        jobs[job_id] = {"status": "queued", "filename": filename, "direction": direction}

    thread = threading.Thread(target=process_job, args=(job_id, filename, direction), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id, "filename": filename})


@app.route("/status/<job_id>")
def status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "任务不存在"}), 404
    return jsonify(job)


@app.route("/download/<job_id>")
def download(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job.get("status") != "done":
        return jsonify({"error": "文件尚未准备好"}), 404
    return send_file(
        job["download"],
        as_attachment=True,
        download_name=job["download_filename"],
    )


if __name__ == "__main__":
    import socket
    for d in [app.config["UPLOAD_FOLDER"], app.config["DOWNLOAD_FOLDER"]]:
        os.makedirs(d, exist_ok=True)

    port = int(os.environ.get("PORT", 5000))
    local_ip = "127.0.0.1"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        pass

    print("=" * 50)
    print("  PDF <-> Word Converter Started")
    print(f"  Local:   http://127.0.0.1:{port}")
    print(f"  Mobile:  http://{local_ip}:{port}")
    print("=" * 50)
    app.run(debug=False, host="0.0.0.0", port=port)
