from flask import Flask, render_template, request, send_file, flash
import os
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from pdf2docx import Converter
from docx2pdf import convert as docx2pdf
import pandas as pd
import img2pdf
from fpdf import FPDF

app = Flask(__name__)
app.secret_key = "supersecret"

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"

for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ---------- Helper ----------
def save_file(file):
    filename = secure_filename(file.filename)
    path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(path)
    return path

# ---------- Routes ----------
@app.route("/")
def index():
    return render_template("index.html")

# ---------- PDF Operations ----------
@app.route("/encrypt", methods=["GET", "POST"])
def encrypt_pdf():
    if request.method == "POST":
        file = request.files["pdf"]
        password = request.form["password"]
        path = save_file(file)

        reader = PdfReader(path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)

        output_path = os.path.join(PROCESSED_FOLDER, "encrypted.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True)
    return render_template("encrypt.html")

@app.route("/decrypt", methods=["GET", "POST"])
def decrypt_pdf():
    if request.method == "POST":
        file = request.files["pdf"]
        password = request.form["password"]
        path = save_file(file)

        reader = PdfReader(path)
        if reader.is_encrypted:
            reader.decrypt(password)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        output_path = os.path.join(PROCESSED_FOLDER, "decrypted.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True)
    return render_template("decrypt.html")

@app.route("/split", methods=["GET", "POST"])
def split_pdf():
    if request.method == "POST":
        file = request.files["pdf"]
        start = int(request.form["start"])
        end = int(request.form["end"])
        path = save_file(file)

        reader = PdfReader(path)
        writer = PdfWriter()
        for i in range(start - 1, end):
            writer.add_page(reader.pages[i])

        output_path = os.path.join(PROCESSED_FOLDER, "split.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True)
    return render_template("split.html")

@app.route("/compress", methods=["GET", "POST"])
def compress_pdf():
    if request.method == "POST":
        file = request.files["pdf"]
        path = save_file(file)

        doc = fitz.open(path)
        output_path = os.path.join(PROCESSED_FOLDER, "compressed.pdf")
        doc.save(output_path, deflate=True, garbage=4)
        doc.close()
        return send_file(output_path, as_attachment=True)
    return render_template("compress.html")

@app.route("/nup", methods=["GET", "POST"])
def nup_pdf():
    if request.method == "POST":
        file = request.files["pdf"]
        pages_per_sheet = int(request.form["nup"])
        path = save_file(file)

        doc = fitz.open(path)
        new_doc = fitz.open()

        rect = doc[0].rect
        w, h = rect.width, rect.height
        if pages_per_sheet == 2:
            new_rect = fitz.Rect(0, 0, w * 2, h)
        else:
            new_rect = fitz.Rect(0, 0, w * 2, h * 2)

        for i in range(0, len(doc), pages_per_sheet):
            page = new_doc.new_page(width=new_rect.width, height=new_rect.height)
            for j in range(pages_per_sheet):
                if i + j >= len(doc):
                    break
                if pages_per_sheet == 2:
                    rect_dest = fitz.Rect(j * w, 0, (j + 1) * w, h)
                else:
                    rect_dest = fitz.Rect((j % 2) * w, (j // 2) * h,
                                          (j % 2 + 1) * w, (j // 2 + 1) * h)
                page.show_pdf_page(rect_dest, doc, i + j)

        output_path = os.path.join(PROCESSED_FOLDER, "nup.pdf")
        new_doc.save(output_path)
        new_doc.close()
        return send_file(output_path, as_attachment=True)
    return render_template("nup.html")

@app.route("/merge", methods=["GET", "POST"])
def merge_pdf():
    if request.method == "POST":
        files = request.files.getlist("pdfs")
        writer = PdfWriter()
        for file in files:
            path = save_file(file)
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
        output_path = os.path.join(PROCESSED_FOLDER, "merged.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True)
    return render_template("merge.html")

# ---------- Conversions ----------
@app.route("/word_to_pdf", methods=["GET", "POST"])
def word_to_pdf():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        output_path = os.path.join(PROCESSED_FOLDER, "converted.pdf")
        docx2pdf(path, output_path)
        return send_file(output_path, as_attachment=True)
    return render_template("word_to_pdf.html")

@app.route("/pdf_to_word", methods=["GET", "POST"])
def pdf_to_word():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        output_path = os.path.join(PROCESSED_FOLDER, "converted.docx")
        cv = Converter(path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        return send_file(output_path, as_attachment=True)
    return render_template("pdf_to_word.html")

@app.route("/excel_to_pdf", methods=["GET", "POST"])
def excel_to_pdf():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        df = pd.read_excel(path)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        for row in df.values.tolist():
            pdf.cell(200, 10, txt=" | ".join(map(str, row)), ln=1)
        output_path = os.path.join(PROCESSED_FOLDER, "converted.pdf")
        pdf.output(output_path)
        return send_file(output_path, as_attachment=True)
    return render_template("excel_to_pdf.html")

@app.route("/pdf_to_excel", methods=["GET", "POST"])
def pdf_to_excel():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        doc = fitz.open(path)
        text = ""
        for page in doc:
            text += page.get_text()
        df = pd.DataFrame({"content": [text]})
        output_path = os.path.join(PROCESSED_FOLDER, "converted.xlsx")
        df.to_excel(output_path, index=False)
        return send_file(output_path, as_attachment=True)
    return render_template("pdf_to_excel.html")

@app.route("/image_to_pdf", methods=["GET", "POST"])
def image_to_pdf():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        output_path = os.path.join(PROCESSED_FOLDER, "converted.pdf")
        with open(output_path, "wb") as f:
            f.write(img2pdf.convert(path))
        return send_file(output_path, as_attachment=True)
    return render_template("image_to_pdf.html")

@app.route("/pdf_to_image", methods=["GET", "POST"])
def pdf_to_image():
    if request.method == "POST":
        file = request.files["file"]
        path = save_file(file)
        doc = fitz.open(path)
        images = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap()
            output_path = os.path.join(PROCESSED_FOLDER, f"page_{i+1}.png")
            pix.save(output_path)
            images.append(output_path)
        return send_file(images[0], as_attachment=True)
    return render_template("pdf_to_image.html")

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)