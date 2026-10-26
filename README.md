# DocuMind AI — Intelligent Document Comparison

## Project Overview

DocuMind AI is an AI-powered document comparison application built using Python and Streamlit. It compares two PDF documents to identify similarities, added content, removed content, and modified information. The application also generates a professional Word report containing the comparison results.

## Features

- Upload and compare two PDF documents.
- Extract text from PDF files.
- Support scanned PDFs using Optical Character Recognition (OCR).
- Compare documents using Google Gemini AI.
- Display a comparison summary and finding counts.
- Highlight added, removed, and modified information using different colors.
- Preview extracted text from both documents.
- Download a professional Word comparison report.
- Professional midnight-blue user interface.
- Validate PDF file size and prevent uploading the same document twice.

## Technologies Used

- **Programming Language:** Python
- **Frontend:** Streamlit
- **AI Model:** Google Gemini
- **PDF Processing:** PyMuPDF
- **OCR:** Tesseract OCR and Pytesseract
- **Word Report Generation:** python-docx
- **Environment Configuration:** python-dotenv

## Project Structure

```text
AI doc/
├── app.py
├── document_reader.py
├── gemini_compare.py
├── requirements.txt
├── README.md
├── .gitignore
└── .env
```

**Note:** The `.env` file contains the Gemini API key and must not be uploaded to GitHub.

## Installation and Setup

### 1. Install Python

Install Python and ensure it is available in your system PATH.

### 2. Install Dependencies

Open a terminal in the project directory and run:

```bash
pip install -r requirements.txt
```

### 3. Install Tesseract OCR

Install Tesseract OCR separately. Ensure its installation path matches the path configured in `document_reader.py`.

### 4. Configure the Gemini API Key

Create a `.env` file in the project directory and add your Gemini API key:

```text
GEMINI_API_KEY=your_api_key_here
```

Replace `your_api_key_here` with your own API key. Do not share or commit this file.

### 5. Run the Application

Run the following command in the terminal:

```bash
python -m streamlit run app.py
```

The application will open in your web browser.

## How to Use

1. Open the DocuMind AI application.
2. Upload the original PDF document.
3. Upload the updated PDF document.
4. Click **Extract Text and Compare Documents**.
5. Review the extracted text and AI comparison results.
6. Examine the identified additions, removals, and modifications.
7. Download the professional Word comparison report.

## Security

- Keep API keys and credentials private.
- Exclude `.env` and `.venv/` from version control.
- Do not upload confidential documents or sensitive information to external AI services without authorization.

