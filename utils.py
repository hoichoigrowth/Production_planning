"""
Shared utility functions for the hoichoi Film Script Production Breakdown tool.
All functions here are pure data/IO — they do not call any Streamlit UI functions
(no st.error / st.warning / st.success). Callers are responsible for UI feedback.
"""

import time
import unicodedata
from typing import Optional, Tuple

import streamlit as st

try:
    import requests
    _REQUESTS_AVAILABLE = True
except ImportError:
    _REQUESTS_AVAILABLE = False


# ── Text helpers ──────────────────────────────────────────────────────────────

def safe_unicode_text(text) -> str:
    """Safely normalise text to NFC Unicode, preserving Bengali and other scripts."""
    if not text:
        return ""
    try:
        if isinstance(text, bytes):
            text = text.decode("utf-8", errors="replace")
        elif not isinstance(text, str):
            text = str(text)
        # Remove invisible/control characters that cause rendering issues
        text = text.replace("\u200b", "")  # zero-width space
        text = text.replace("\ufeff", "")  # BOM
        text = text.replace("\u200c", "")  # zero-width non-joiner
        text = text.replace("\u200d", "")  # zero-width joiner
        return unicodedata.normalize("NFC", text)
    except Exception:
        return str(text)


# ── API key helpers ───────────────────────────────────────────────────────────

def get_api_key() -> Optional[str]:
    """Return the OpenAI API key from Streamlit secrets, or None if not set."""
    try:
        return st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        return None


def get_mistral_api_key() -> Optional[str]:
    """Return the Mistral API key from Streamlit secrets, or None if not set."""
    try:
        return st.secrets.get("MISTRAL_API_KEY", None)
    except Exception:
        return None


def get_mistral_api_key_with_session() -> Optional[str]:
    """Return the Mistral API key, preferring a session-scoped override if set."""
    if st.session_state.get("temp_mistral_key"):
        return st.session_state.temp_mistral_key
    try:
        return st.secrets.get("MISTRAL_API_KEY", None)
    except Exception:
        return None


# ── OCR availability ──────────────────────────────────────────────────────────

def check_mistral_ocr_availability() -> Tuple[bool, str]:
    """
    Test whether the Mistral OCR API is reachable with the configured key.

    Writes results to st.session_state.ocr_available and
    st.session_state.ocr_error_message so that multiple callers in the same
    session share the result without re-checking.
    Returns (is_available: bool, message: str).
    """
    # Initialise session state keys if not present
    if "ocr_available" not in st.session_state:
        st.session_state.ocr_available = False
    if "ocr_error_message" not in st.session_state:
        st.session_state.ocr_error_message = ""

    mistral_key = get_mistral_api_key_with_session()
    if not mistral_key:
        st.session_state.ocr_available = False
        st.session_state.ocr_error_message = "Mistral API key not configured"
        return False, "Mistral API key not configured"

    if not _REQUESTS_AVAILABLE:
        st.session_state.ocr_available = False
        st.session_state.ocr_error_message = "requests library not available"
        return False, "requests library not available"

    try:
        headers = {"Authorization": f"Bearer {mistral_key}"}
        response = requests.get(
            "https://api.mistral.ai/v1/models", headers=headers, timeout=10
        )
        if response.status_code == 200:
            st.session_state.ocr_available = True
            st.session_state.ocr_error_message = ""
            return True, "Mistral OCR available"
        elif response.status_code == 401:
            msg = "Invalid Mistral API key"
        elif response.status_code == 403:
            msg = "OCR access not enabled for this account"
        else:
            msg = f"Mistral API error: {response.status_code}"
        st.session_state.ocr_available = False
        st.session_state.ocr_error_message = msg
        return False, msg
    except Exception as exc:
        msg = f"Mistral API connection failed: {exc}"
        st.session_state.ocr_available = False
        st.session_state.ocr_error_message = msg
        return False, msg


# ── Mistral file upload + OCR ─────────────────────────────────────────────────

def upload_file_to_mistral(file_data: bytes, filename: str, mistral_key: str) -> str:
    """
    Upload a document to the Mistral Files API for OCR processing.

    Returns the file_id string on success. Raises RuntimeError on failure.
    """
    url = "https://api.mistral.ai/v1/files"
    headers = {"Authorization": f"Bearer {mistral_key}"}
    mime = "image/jpeg" if filename.lower().endswith((".jpg", ".jpeg")) else "image/png"
    files = {"file": (filename, file_data, mime)}
    data = {"purpose": "ocr"}

    response = requests.post(url, headers=headers, files=files, data=data, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(
            f"Mistral file upload failed ({response.status_code}): {response.text}"
        )
    file_id = response.json().get("id")
    if not file_id:
        raise RuntimeError("Mistral file upload returned no file ID.")
    return file_id


def get_mistral_ocr_result(file_id: str, mistral_key: str, language: str = "ben+eng") -> str:
    """
    Run OCR on an already-uploaded Mistral file and return extracted markdown text.

    Step 1: GET /v1/files/{file_id}/url to obtain a signed URL.
    Step 2: POST /v1/ocr with that signed URL.
    Raises RuntimeError on any API failure.
    Returns concatenated markdown from all pages.
    """
    auth_header = {"Authorization": f"Bearer {mistral_key}"}

    # Step 1 — signed URL
    url_resp = requests.get(
        f"https://api.mistral.ai/v1/files/{file_id}/url",
        headers=auth_header,
        timeout=30,
    )
    if url_resp.status_code != 200:
        raise RuntimeError(
            f"Failed to retrieve signed URL for file {file_id} "
            f"({url_resp.status_code}): {url_resp.text}"
        )
    signed_url = url_resp.json().get("url")
    if not signed_url:
        raise RuntimeError(f"No signed URL returned for file {file_id}.")

    # Step 2 — OCR
    payload = {
        "model": "mistral-ocr-latest",
        "document": {"type": "document_url", "document_url": signed_url},
        "include_image_base64": False,
    }
    ocr_resp = requests.post(
        "https://api.mistral.ai/v1/ocr",
        headers={**auth_header, "Content-Type": "application/json"},
        json=payload,
        timeout=120,
    )
    if ocr_resp.status_code != 200:
        raise RuntimeError(
            f"Mistral OCR failed ({ocr_resp.status_code}): {ocr_resp.text}"
        )

    pages = ocr_resp.json().get("pages", [])
    return "\n\n".join(p.get("markdown", "") for p in pages)


def extract_text_with_mistral_ocr(
    image_file,
    language: str = "ben+eng",
    progress_callback=None,
) -> str:
    """
    Upload an image and extract text via Mistral OCR.

    progress_callback(float, str) is optional and called with (fraction, status_msg).
    Raises RuntimeError if the API key is missing or any API call fails.
    Returns the extracted text (normalised Unicode).
    """
    mistral_key = get_mistral_api_key_with_session()
    if not mistral_key:
        raise RuntimeError("Mistral API key not configured for OCR.")

    if hasattr(image_file, "getvalue"):
        file_data = image_file.getvalue()
        filename = image_file.name
    else:
        file_data = image_file
        filename = "uploaded_image.jpg"

    if progress_callback:
        progress_callback(0.1, f"Uploading {filename} ({len(file_data) / 1024:.1f} KB) to Mistral...")

    file_id = upload_file_to_mistral(file_data, filename, mistral_key)

    if progress_callback:
        progress_callback(0.4, f"File uploaded: {file_id}")

    time.sleep(3)  # Allow Mistral to index the file before OCR

    if progress_callback:
        progress_callback(0.7, "Running OCR...")

    text = get_mistral_ocr_result(file_id, mistral_key, language)

    if progress_callback:
        progress_callback(1.0, f"OCR complete — {len(text)} characters extracted.")

    return safe_unicode_text(text)
