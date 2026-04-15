"""Extract text from PDF files using PyMuPDF."""

import fitz  # PyMuPDF


def extract_text(pdf_path: str) -> str:
    """Extract all text from a PDF file."""
    doc = fitz.open(pdf_path)
    text_parts = []
    for page in doc:
        text_parts.append(page.get_text())
    doc.close()
    return "\n".join(text_parts)


def extract_text_by_page(pdf_path: str) -> list:
    """Extract text from each page of a PDF, returned as a list."""
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        pages.append(page.get_text())
    doc.close()
    return pages


def extract_report_metadata(pdf_path: str) -> dict:
    """Try to extract report metadata (title, author, revision) from first pages."""
    doc = fitz.open(pdf_path)
    metadata = {
        "title": doc.metadata.get("title", ""),
        "author": doc.metadata.get("author", ""),
        "filename": pdf_path.split("/")[-1].split("\\")[-1],
    }
    doc.close()
    return metadata
