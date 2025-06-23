import PyPDF2
from docx import Document
import pytesseract
from PIL import Image
import os

def process_uploaded_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = "".join(page.extract_text() for page in reader.pages)
        return text
    elif ext == ".docx":
        doc = Document(file_path)
        text = "\n".join(p.text for p in doc.paragraphs)
        return text
    elif ext in [".png", ".jpg", ".jpeg"]:
        text = pytesseract.image_to_string(Image.open(file_path), lang="chi_sim+eng")
        return text
    else:
        raise ValueError("不支持的文件格式")