
import io
import os

import fitz
import pytesseract
from PIL import Image

# Tesseract OCR installation path
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)


def extract_text_from_pdf(uploaded_file):
    if uploaded_file is None:
        raise ValueError("Please upload a PDF file.")

    pdf_bytes = uploaded_file.getvalue()

    if not pdf_bytes:
        raise ValueError("The uploaded PDF is empty.")

    try:
        pdf_document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )
    except Exception as error:
        raise ValueError("Unable to open this PDF file.") from error

    extracted_pages = []

    try:
        for page_number, page in enumerate(pdf_document, start=1):
            page_text = page.get_text("text").strip()

            # Use OCR when the page has little or no selectable text
            if len(page_text) < 20:
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(2, 2),
                    alpha=False
                )

                image = Image.open(
                    io.BytesIO(pixmap.tobytes("png"))
                )

                ocr_text = pytesseract.image_to_string(
                    image,
                    lang="eng"
                ).strip()

                if len(ocr_text) > len(page_text):
                    page_text = ocr_text

            if page_text:
                extracted_pages.append(
                    f"--- Page {page_number} ---\n{page_text}"
                )

    finally:
        pdf_document.close()

    full_text = "\n\n".join(extracted_pages).strip()

    if not full_text:
        raise ValueError(
            "No readable text found. Check whether the PDF is blank "
            "or the scan quality is too low."
        )

    return full_text
