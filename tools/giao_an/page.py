"""Streamlit page for the giáo án generator."""

import io
import subprocess
import sys
from pathlib import Path

import streamlit as st

from tools.common import XLSX_MIME, save_upload, tool_state, tool_workdir, zip_files
from tools.giao_an.generator import get_all_classes

SLUG = "giao-an"
REPO_ROOT = Path(__file__).resolve().parents[2]

# A class that runs longer than this is skipped.
TIMEOUT_SECONDS = 180

k = tool_state(SLUG)


def _run_class(schedule_path: Path, content_path: Path, class_name: str, output_dir: Path):
    """Generate one class in a subprocess so a hang can be killed by timeout.

    Returns (returncode or None on timeout, log).
    """
    cmd = [
        sys.executable, "-m", "tools.giao_an",
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
        return proc.returncode, log
    except subprocess.TimeoutExpired:
        proc.kill()
        log, _ = proc.communicate()
        return None, log or ""


def render() -> None:
    st.title("📚 Soạn giáo án")
    st.caption("Tạo giáo án từ kế hoạch giảng dạy theo tuần và file nội dung chi tiết.")

    # --- Step 1: Upload files ---
    st.subheader("1. Tải file lên")
    col1, col2 = st.columns(2)
    with col1:
        schedule_upload = st.file_uploader(
            "Kế hoạch giảng dạy (.xlsx)", type=["xlsx"], key=k("schedule")
        )
    with col2:
        content_upload = st.file_uploader(
            'Nội dung chi tiết "KH chi tiết" (.xlsx)', type=["xlsx"], key=k("content")
        )

    if not schedule_upload or not content_upload:
        st.info("Vui lòng tải lên cả 2 file để tiếp tục.")
        return

    workdir = tool_workdir(SLUG)
    schedule_path = save_upload(schedule_upload, workdir)
    content_path = save_upload(content_upload, workdir)

    # --- Step 2: Pick classes ---
    st.subheader("2. Chọn lớp")
    try:
        all_classes = get_all_classes(schedule_path)
    except Exception as e:
        st.error(f"Không đọc được file kế hoạch: {e}")
        return

    if not all_classes:
        st.error("File kế hoạch không có sheet lớp nào.")
        return

    # All classes start ticked — generating for every class is the common case.
    checkbox_cols = st.columns(3)
    selected_classes = []
    for i, name in enumerate(all_classes):
        key = k(f"class_{name}")
        st.session_state.setdefault(key, True)
        if checkbox_cols[i % 3].checkbox(name, key=key):
            selected_classes.append(name)

    if not selected_classes:
        st.info("Chọn ít nhất 1 lớp để tiếp tục.")
        return

    # --- Step 3: Generate ---
    st.subheader("3. Tạo giáo án")
    if st.button("🚀 Bắt đầu tạo", type="primary", use_container_width=True):
        _generate(workdir, schedule_path, content_path, selected_classes)

    _show_results()


def _generate(workdir: Path, schedule_path: Path, content_path: Path, classes: list[str]) -> None:
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
    total = len(classes)

    for i, class_name in enumerate(classes):
        progress.progress(i / total, text=f"({i + 1}/{total}) Đang xử lý: {class_name}")
        returncode, log = _run_class(schedule_path, content_path, class_name, output_dir)

        if returncode is None:
            msg = f"Quá thời gian chờ ({TIMEOUT_SECONDS}s)"
            errors.append((class_name, msg, log))
            with status_box:
                st.write(f"⏱️ {class_name} — {msg}, đã bỏ qua")
            full_log.write(f"\n{'=' * 60}\n[{class_name}] TIMEOUT\n{'=' * 60}\n{log}")
            continue

        full_log.write(f"\n{'=' * 60}\n[{class_name}]\n{'=' * 60}\n{log}")
        if returncode == 0:
            succeeded.append(class_name)
            with status_box:
                st.write(f"✅ {class_name}")
        else:
            err_msg = (log.strip().splitlines() or [f"exit {returncode}"])[-1][:200]
            errors.append((class_name, err_msg, log))
            with status_box:
                st.write(f"❌ {class_name} — {err_msg}")

    progress.progress(1.0, text=f"Hoàn thành: {len(succeeded)}/{total} thành công")

    st.session_state[k("output_dir")] = str(output_dir)
    st.session_state[k("log")] = full_log.getvalue()
    st.session_state[k("succeeded")] = succeeded
    st.session_state[k("errors")] = errors


def _show_results() -> None:
    output_dir_str = st.session_state.get(k("output_dir"))
    if not output_dir_str:
        return

    files = sorted(Path(output_dir_str).glob("*.xlsx"))
    succeeded = st.session_state.get(k("succeeded"), [])
    errors = st.session_state.get(k("errors"), [])

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
            data=zip_files(files),
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
                    mime=XLSX_MIME,
                    key=k(f"dl_{f.name}"),
                )

    if log := st.session_state.get(k("log")):
        with st.expander("Log chi tiết"):
            st.code(log)
