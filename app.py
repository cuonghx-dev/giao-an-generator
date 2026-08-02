"""Streamlit UI for Giáo Án Generator."""

import io
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from generate_giao_an_v2 import get_all_classes

REPO_ROOT = Path(__file__).resolve().parent
CLI_SCRIPT = REPO_ROOT / "generate_giao_an_v2.py"

# A class that runs longer than this is skipped.
TIMEOUT_SECONDS = 180


st.set_page_config(page_title="Giáo Án Generator", page_icon="📚", layout="centered")
st.title("📚 Giáo Án Generator")
st.caption("Tạo giáo án từ kế hoạch giảng dạy theo tuần và file nội dung chi tiết.")


def save_upload(uploaded_file, dest_dir: Path) -> Path:
    path = dest_dir / uploaded_file.name
    path.write_bytes(uploaded_file.getbuffer())
    return path


def zip_outputs(output_dir: Path) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(output_dir.glob("*.xlsx")):
            zf.write(f, arcname=f.name)
    return buf.getvalue()


# --- Step 1: Upload files ---
st.subheader("1. Tải file lên")
col1, col2 = st.columns(2)
with col1:
    schedule_upload = st.file_uploader(
        "Kế hoạch giảng dạy (.xlsx)",
        type=["xlsx"],
        key="schedule",
    )
with col2:
    content_upload = st.file_uploader(
        'Nội dung chi tiết "KH chi tiết" (.xlsx)',
        type=["xlsx"],
        key="content",
    )

if not schedule_upload or not content_upload:
    st.info("Vui lòng tải lên cả 2 file để tiếp tục.")
    st.stop()

# Persist uploads to a session-scoped temp dir so re-runs don't re-write.
if "workdir" not in st.session_state:
    st.session_state["workdir"] = tempfile.mkdtemp(prefix="giao_an_")
workdir = Path(st.session_state["workdir"])

schedule_path = save_upload(schedule_upload, workdir)
content_path = save_upload(content_upload, workdir)

# --- Step 2: Pick classes ---
st.subheader("2. Chọn lớp")
try:
    all_classes = get_all_classes(schedule_path)
except Exception as e:
    st.error(f"Không đọc được file kế hoạch: {e}")
    st.stop()

if not all_classes:
    st.error("File kế hoạch không có sheet lớp nào.")
    st.stop()

# All classes start ticked — generating for every class is the common case.
checkbox_cols = st.columns(3)
selected_classes = []
for i, name in enumerate(all_classes):
    key = f"class_{name}"
    st.session_state.setdefault(key, True)
    if checkbox_cols[i % 3].checkbox(name, key=key):
        selected_classes.append(name)

if not selected_classes:
    st.info("Chọn ít nhất 1 lớp để tiếp tục.")
    st.stop()

# --- Step 3: Generate ---
st.subheader("3. Tạo giáo án")
if st.button("🚀 Bắt đầu tạo", type="primary", use_container_width=True):
    output_dir = workdir / "output"
    if output_dir.exists():
        for f in output_dir.glob("*.xlsx"):
            f.unlink()
    output_dir.mkdir(parents=True, exist_ok=True)

    progress = st.progress(0.0, text="Đang chuẩn bị...")
    status_box = st.container()
    succeeded: list[str] = []
    errors: list[tuple[str, str, str]] = []  # (class_name, error_message, log)
    full_log = io.StringIO()
    total = len(selected_classes)

    for i, class_name in enumerate(selected_classes):
        progress.progress(
            i / total,
            text=f"({i + 1}/{total}) Đang xử lý: {class_name}",
        )

        cmd = [
            sys.executable,
            str(CLI_SCRIPT),
            "-s", str(schedule_path),
            "-c", str(content_path),
            "-C", class_name,
            "-o", str(output_dir),
        ]
        proc = subprocess.Popen(
            cmd,
            cwd=str(REPO_ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )
        try:
            log, _ = proc.communicate(timeout=TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            proc.kill()
            log, _ = proc.communicate()
            msg = f"Quá thời gian chờ ({TIMEOUT_SECONDS}s)"
            errors.append((class_name, msg, log or ""))
            with status_box:
                st.write(f"⏱️ {class_name} — {msg}, đã bỏ qua")
            full_log.write(f"\n{'=' * 60}\n[{class_name}] TIMEOUT\n{'=' * 60}\n{log or ''}")
            continue

        full_log.write(f"\n{'=' * 60}\n[{class_name}]\n{'=' * 60}\n{log}")

        if proc.returncode == 0:
            succeeded.append(class_name)
            with status_box:
                st.write(f"✅ {class_name}")
        else:
            err_msg = (log.strip().splitlines() or [f"exit {proc.returncode}"])[-1][:200]
            errors.append((class_name, err_msg, log))
            with status_box:
                st.write(f"❌ {class_name} — {err_msg}")

    progress.progress(
        1.0,
        text=f"Hoàn thành: {len(succeeded)}/{total} thành công",
    )

    st.session_state["output_dir"] = str(output_dir)
    st.session_state["log"] = full_log.getvalue()
    st.session_state["succeeded"] = succeeded
    st.session_state["errors"] = errors

# --- Step 4: Result + download ---
output_dir_str = st.session_state.get("output_dir")
if output_dir_str:
    output_dir = Path(output_dir_str)
    files = sorted(output_dir.glob("*.xlsx"))
    succeeded = st.session_state.get("succeeded", [])
    errors = st.session_state.get("errors", [])

    if succeeded:
        st.success(f"Đã tạo {len(succeeded)} file giáo án.")
    if errors:
        st.warning(f"{len(errors)} lớp lỗi, đã bỏ qua:")
        for class_name, err, log in errors:
            with st.expander(f"❌ {class_name}: {err}"):
                st.code(log or err)

    if errors:
        st.info("Có lớp bị lỗi — vui lòng xử lý lỗi rồi chạy lại để tải file.")
    elif files:
        st.subheader("4. Tải về")
        st.download_button(
            label=f"⬇️ Tải zip tất cả ({len(files)} file)",
            data=zip_outputs(output_dir),
            file_name="giao_an.zip",
            mime="application/zip",
            use_container_width=True,
        )
        with st.expander("Tải từng file"):
            for f in files:
                st.download_button(
                    label=f.name,
                    data=f.read_bytes(),
                    file_name=f.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{f.name}",
                )

    if log := st.session_state.get("log"):
        with st.expander("Log chi tiết"):
            st.code(log)
