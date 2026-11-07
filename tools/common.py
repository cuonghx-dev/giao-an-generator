"""Helpers shared by tool pages."""

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def tool_state(slug: str):
    """Return a key builder that namespaces session state per tool.

    Streamlit session state is shared by every page, so two tools using the
    same key (``"workdir"``, ``"log"``...) would overwrite each other.
    """
    return lambda name: f"{slug}:{name}"


def tool_workdir(slug: str) -> Path:
    """A temp directory that lives for the browser session, one per tool."""
    key = f"{slug}:workdir"
    if key not in st.session_state:
        st.session_state[key] = tempfile.mkdtemp(prefix=f"{slug}_")
    return Path(st.session_state[key])


def save_upload(uploaded_file, dest_dir: Path) -> Path:
    path = dest_dir / uploaded_file.name
    path.write_bytes(uploaded_file.getbuffer())
    return path


def zip_files(files: list[Path]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=f.name)
    return buf.getvalue()
