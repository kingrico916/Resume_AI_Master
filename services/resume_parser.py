"""
services.resume_parser
======================
Extract plain text from .pdf, .docx, or .txt resume uploads.
Returns empty string + error message on failure — never raises.

Usage:
    from services.resume_parser import parse_resume
    text, error = parse_resume(file_storage_object)
    # file_storage_object is a Flask FileStorage (request.files['resume'])
"""

import os


def parse_resume(file_storage) -> tuple:
    """
    Extract plain text from a Flask FileStorage object.

    Parameters
    ----------
    file_storage : werkzeug.datastructures.FileStorage
        The uploaded file from request.files['resume']

    Returns
    -------
    tuple[str, str]
        (text, error_message)
        text is empty string on failure.
        error_message is empty string on success.
    """
    if file_storage is None:
        return ("", "No file provided")

    filename = file_storage.filename or ""
    ext = os.path.splitext(filename)[1].lower()

    if ext == ".pdf":
        return _parse_pdf(file_storage)
    elif ext == ".docx":
        return _parse_docx(file_storage)
    elif ext == ".txt":
        return _parse_txt(file_storage)
    else:
        return ("", f"Unsupported file type: '{ext}'. Supported types: .pdf, .docx, .txt")


def _parse_pdf(file_storage) -> tuple:
    """Extract text from PDF using pypdf."""
    try:
        from pypdf import PdfReader
    except ImportError as e:
        return ("", f"pypdf library not available: {e}")

    try:
        reader = PdfReader(file_storage.stream)
        pages_text = []
        for page in reader.pages:
            try:
                page_text = page.extract_text()
                if page_text:
                    pages_text.append(page_text)
            except Exception:
                pass  # Skip unreadable pages
        text = "\n".join(pages_text).strip()
        if not text:
            return ("", "PDF appears to be empty or image-only (no extractable text found)")
        return (text, "")
    except Exception as e:
        return ("", f"PDF parsing error: {e}")


def _parse_docx(file_storage) -> tuple:
    """Extract text from DOCX using python-docx."""
    try:
        import docx
    except ImportError as e:
        return ("", f"python-docx library not available: {e}")

    try:
        document = docx.Document(file_storage.stream)
        paragraphs = []
        for para in document.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text)
        text = "\n".join(paragraphs).strip()
        if not text:
            return ("", "DOCX file appears to be empty (no paragraph text found)")
        return (text, "")
    except Exception as e:
        return ("", f"DOCX parsing error: {e}")


def _parse_txt(file_storage) -> tuple:
    """Read plain text file as UTF-8."""
    try:
        raw = file_storage.stream.read()
        text = raw.decode("utf-8", errors="replace").strip()
        if not text:
            return ("", "Text file appears to be empty")
        return (text, "")
    except Exception as e:
        return ("", f"Text file read error: {e}")
