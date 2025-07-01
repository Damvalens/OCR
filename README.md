# OCR Streamlit App

This project provides a small [Streamlit](https://streamlit.io/) application to perform Optical Character Recognition (OCR) on scanned invoices or other PDF/image files. The app can extract text and tables from uploaded documents and optionally send the text to OpenAI for additional analysis.

## Project Purpose

The application aims to simplify the extraction of data from invoices and other scanned documents. Users can upload PDFs or images, preview extracted tables in the browser, and download the results to an Excel file. When configured, the text can also be sent to OpenAI for further processing.

## Prerequisites

The code depends on a few system packages in addition to the Python libraries listed in `requirements.txt`:

- **Tesseract OCR** – required by `pytesseract` for performing OCR on images.
- **Poppler** – used by `pdf2image` to convert PDF pages into images.

Install these utilities using your system package manager (for example, on Ubuntu: `sudo apt-get install tesseract-ocr poppler-utils`).

## Environment Setup

1. Install the Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Ensure that `tesseract` and `poppler` are available on your system path.

3. (Optional) Install the `openai` package if you plan to use the OpenAI integration:

   ```bash
   pip install openai
   ```

## Running the Application

Launch the Streamlit interface with:

```bash
streamlit run ocr.py
```

Upload a PDF or image when prompted and the extracted data will appear in the browser. You can download the results as an Excel file directly from the app.

## Optional OpenAI Integration

The script contains optional functionality to send extracted text to OpenAI. To enable this feature, install the `openai` package as shown above and set the `OPENAI_API_KEY` environment variable with your API key:

```bash
export OPENAI_API_KEY="sk-..."
```

When this variable is present, the app can call OpenAI's API to analyze the extracted text (assuming the corresponding function is implemented).
