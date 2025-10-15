from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
import subprocess
import uuid
import shutil
from pathlib import Path
import sys

app = Flask(__name__)
app.secret_key = "replace-with-a-secret-key"
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_PDF = "stickers_msg_final.pdf"  # as your script writes by default

UPLOAD_DIR.mkdir(exist_ok=True)

# Path to the original script (unchanged)
ORIGINAL_SCRIPT = BASE_DIR / "sticker_generator_clean.py"

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    # validate file
    f = request.files.get("excel_file")
    if not f or f.filename == "":
        flash("Please upload an Excel (.xlsx) file.")
        return redirect(url_for("index"))

    filename = f.filename
    # create unique working dir for safety
    runid = uuid.uuid4().hex[:8]
    workdir = UPLOAD_DIR / runid
    workdir.mkdir(parents=True, exist_ok=True)

    saved_excel = workdir / filename
    f.save(saved_excel)

    # options from form
    font_choice = request.form.get("font_choice", "1")  # "1","2","3"
    include_logo = request.form.get("include_logo", "n")  # 'y' or 'n'
    include_qr = request.form.get("include_qr", "n")      # 'y' or 'n'

    # Build stdin input for the original script.
    answers = "\n".join([
        font_choice,
        str(saved_excel),   # Excel file path (we saved it inside workdir)
        "y" if include_logo == "y" else "n",
        "y" if include_qr == "y" else "n"
    ]) + "\n"

    # Run the original script as subprocess in the workdir.
    # Use the same Python interpreter and force UTF-8 for child IO so emoji / unicode doesn't crash.
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    cmd = [sys.executable, str(ORIGINAL_SCRIPT)]
    try:
        proc = subprocess.run(
            cmd,
            input=answers.encode("utf-8"),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=str(workdir),
            timeout=300,    # 5 minutes; increase if your script needs longer
            env=env,
            check=False
        )
    except subprocess.TimeoutExpired:
        flash("Processing timed out. The script took too long.")
        shutil.rmtree(workdir, ignore_errors=True)
        return redirect(url_for("index"))

    # decode outputs safely
    out = proc.stdout.decode("utf-8", errors="replace") if proc.stdout else ""
    err = proc.stderr.decode("utf-8", errors="replace") if proc.stderr else ""

    # Check output PDF path: script may save in workdir or project root
    generated_pdf = workdir / OUTPUT_PDF
    project_pdf = BASE_DIR / OUTPUT_PDF

    # if generated in project root, move it into the run-specific workdir
    if not generated_pdf.exists() and project_pdf.exists():
        try:
            shutil.move(str(project_pdf), str(generated_pdf))
        except Exception as e:
            (workdir / "move_error.txt").write_text(str(e), encoding="utf-8")

    # If PDF still missing -> save logs and show error page
    if not generated_pdf.exists():
        (workdir / "script_stdout.txt").write_text(out, encoding="utf-8")
        (workdir / "script_stderr.txt").write_text(err, encoding="utf-8")
        flash("Failed to create PDF. See server logs / messages.")
        return render_template("result.html", success=False, stdout=out, stderr=err, runid=runid)

    # Success — render result with download link
    return render_template("result.html", success=True, runid=runid, pdf_name=generated_pdf.name)

@app.route("/download/<runid>/<filename>", methods=["GET"])
def download(runid, filename):
    folder = UPLOAD_DIR / runid
    if not folder.exists():
        flash("File not found.")
        return redirect(url_for("index"))
    # send_from_directory expects (directory, filename)
    return send_from_directory(str(folder), filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
