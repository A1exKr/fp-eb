from io import BytesIO
from pathlib import Path

import pdfplumber
from docx import Document
from fastapi import UploadFile
from pypdf import PdfReader


SUPPORTED_RFP_EXTENSIONS = {".txt", ".md", ".json", ".pdf", ".docx"}


def _decode_text_bytes(content: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode("utf-8", errors="ignore")


def _extract_pdf_text(content: bytes) -> str:
    pages: list[str] = []
    reader = PdfReader(BytesIO(content))
    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():
            pages.append(text)

    if pages:
        return "\n\n".join(pages)

    with pdfplumber.open(BytesIO(content)) as pdf:
        plumber_pages = [(page.extract_text() or "") for page in pdf.pages]
    return "\n\n".join([page for page in plumber_pages if page.strip()])


def _extract_docx_text(content: bytes) -> str:
    document = Document(BytesIO(content))
    return "\n".join([paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()])


def extract_rfp_text(upload: UploadFile) -> tuple[str, str]:
    filename = upload.filename or "uploaded_rfp.txt"
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED_RFP_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type '{suffix or 'unknown'}'. Supported types: {', '.join(sorted(SUPPORTED_RFP_EXTENSIONS))}"
        )

    content = upload.file.read()
    if not content:
        raise ValueError("Uploaded file is empty.")

    if suffix in {".txt", ".md", ".json"}:
        text = _decode_text_bytes(content)
    elif suffix == ".pdf":
        text = _extract_pdf_text(content)
    elif suffix == ".docx":
        text = _extract_docx_text(content)
    else:
        raise ValueError(f"Unsupported file type '{suffix}'.")

    normalized = text.strip()
    if len(normalized) < 20:
        raise ValueError("The uploaded file did not yield enough text to parse.")

    return filename, normalized