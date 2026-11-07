"""Streamlit page for batch-placing photos into a template's photo frame."""

import hashlib
import io
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont

from tools.common import tool_state, tool_workdir, zip_files
from tools.ghep_anh.generator import (
    FIT_CONTAIN,
    FIT_COVER,
    HEIC_SUPPORTED,
    PHOTO_EXTENSIONS,
    Box,
    compose,
    detect_frames,
    load_image,
    save_image,
    save_pdf,
)

SLUG = "ghep-anh"
MANUAL = "manual"
PREVIEW_WIDTH = 900  # px; previews are downscaled to keep the page snappy

ALIGN_Y = {"Giữa": 0.5, "Trên": 0.0, "Dưới": 1.0}
ALIGN_X = {"Giữa": 0.5, "Trái": 0.0, "Phải": 1.0}
FIT_LABELS = {FIT_COVER: "Lấp đầy khung (cắt bớt ảnh)", FIT_CONTAIN: "Hiện trọn ảnh (có thể hở viền)"}

k = tool_state(SLUG)


@st.cache_data(show_spinner=False)
def _detect(data: bytes) -> list[tuple[int, int, int, int]]:
    boxes = detect_frames(load_image(io.BytesIO(data)))
    return [(b.left, b.top, b.right, b.bottom) for b in boxes]


def _preview(img: Image.Image) -> Image.Image:
    if img.width <= PREVIEW_WIDTH:
        return img
    ratio = PREVIEW_WIDTH / img.width
    return img.resize((PREVIEW_WIDTH, round(img.height * ratio)), Image.Resampling.LANCZOS)


def _draw_boxes(template: Image.Image, boxes: list[Box], selected: Box | None) -> Image.Image:
    """Outline candidate frames (numbered) and the selected one in red."""
    img = template.convert("RGB")
    draw = ImageDraw.Draw(img)
    line = max(2, template.width // 250)
    font = ImageFont.load_default(size=max(16, template.width // 25))
    for i, b in enumerate(boxes, 1):
        draw.rectangle((b.left, b.top, b.right - 1, b.bottom - 1), outline="#1e88e5", width=line)
        draw.text((b.left + 2 * line, b.top + line), str(i), fill="#1e88e5", font=font,
                  stroke_width=line, stroke_fill="white")
    if selected:
        draw.rectangle((selected.left, selected.top, selected.right - 1, selected.bottom - 1),
                       outline="red", width=line * 2)
    return img


def _unique_names(names: list[str], ext: str) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for name in names:
        stem = Path(name).stem
        n = seen.get(stem, 0)
        seen[stem] = n + 1
        out.append(f"{stem}.{ext}" if n == 0 else f"{stem}_{n + 1}.{ext}")
    return out


def render() -> None:
    st.title("🖼️ Ghép ảnh hàng loạt")
    st.caption("Ghép từng ảnh vào khung ảnh của mẫu thiết kế — mỗi ảnh ra một trang.")

    # --- Step 1: Template ---
    st.subheader("1. Mẫu thiết kế")
    template_upload = st.file_uploader(
        "Ảnh mẫu (xuất từ Canva dạng PNG/JPG, còn trống khung ảnh)",
        type=["png", "jpg", "jpeg", "webp"],
        key=k("template"),
    )
    if not template_upload:
        st.info("Tải ảnh mẫu lên để tiếp tục.")
        return

    template_bytes = template_upload.getvalue()
    template_id = hashlib.md5(template_bytes).hexdigest()[:8]
    try:
        template = load_image(io.BytesIO(template_bytes))
    except Exception as e:
        st.error(f"Không đọc được ảnh mẫu: {e}")
        return
    tw, th = template.size

    # --- Step 2: Frame ---
    st.subheader("2. Chọn khung ảnh")
    with st.spinner("Đang tìm khung ảnh..."):
        candidates = [Box(*b) for b in _detect(template_bytes)]

    options = list(range(len(candidates))) + [MANUAL]
    choice = st.radio(
        "Khung phát hiện được (viền xanh, đánh số trên hình)" if candidates
        else "Không tự tìm được khung — hãy chỉnh tay bên dưới",
        options,
        format_func=lambda o: "Tự chỉnh" if o == MANUAL else f"Khung {o + 1}",
        horizontal=True,
        key=k(f"choice_{template_id}"),
    )
    if choice == MANUAL:
        base = Box(tw // 4, th // 4, tw * 3 // 4, th * 3 // 4)
    else:
        base = candidates[choice]

    # Widget keys include the template and the choice, so picking another
    # candidate resets the sliders to that candidate's box.
    with st.expander("Chỉnh vị trí khung (pixel)", expanded=choice == MANUAL):
        x0, x1 = st.slider("Trái – Phải", 0, tw, (base.left, base.right),
                           key=k(f"x_{template_id}_{choice}"))
        y0, y1 = st.slider("Trên – Dưới", 0, th, (base.top, base.bottom),
                           key=k(f"y_{template_id}_{choice}"))
    if x1 - x0 < 2 or y1 - y0 < 2:
        st.error("Khung quá nhỏ.")
        return
    box = Box(x0, y0, x1, y1)

    # --- Step 3: Photos ---
    st.subheader("3. Ảnh cần ghép")
    photo_uploads = st.file_uploader(
        "Chọn nhiều ảnh cùng lúc" + ("" if HEIC_SUPPORTED else " (HEIC chưa hỗ trợ — đổi sang JPG)"),
        type=PHOTO_EXTENSIONS,
        accept_multiple_files=True,
        key=k("photos"),
    )

    with st.expander("Tuỳ chọn", expanded=False):
        c1, c2 = st.columns(2)
        fit = c1.radio("Cách đặt ảnh", list(FIT_LABELS), format_func=FIT_LABELS.get, key=k("fit"))
        radius = c2.number_input("Bo góc (pixel)", 0, min(box.width, box.height) // 2, 0,
                                 key=k("radius"))
        c3, c4 = st.columns(2)
        align_y = c3.selectbox("Giữ phần nào theo chiều dọc", list(ALIGN_Y), key=k("align_y"))
        align_x = c4.selectbox("Giữ phần nào theo chiều ngang", list(ALIGN_X), key=k("align_x"))
        c5, c6 = st.columns(2)
        fmt = c5.radio("Định dạng", ["jpg", "png"], horizontal=True, key=k("fmt"))
        want_pdf = c6.checkbox("Kèm 1 file PDF gộp tất cả (để in)", value=True, key=k("pdf"))
    align = (ALIGN_X[align_x], ALIGN_Y[align_y])

    # Preview: frames on the template, and the first photo placed in the frame.
    col_a, col_b = st.columns(2)
    col_a.image(_preview(_draw_boxes(template, candidates, box)),
                caption="Khung đang chọn (viền đỏ)", width="stretch")
    if photo_uploads:
        try:
            sample = compose(template, load_image(photo_uploads[0]), box, fit, align, radius)
            col_b.image(_preview(sample), caption=f"Xem trước: {photo_uploads[0].name}",
                        width="stretch")
        except Exception as e:
            col_b.error(f"Không ghép được {photo_uploads[0].name}: {e}")

    if not photo_uploads:
        st.info("Tải ảnh lên để tiếp tục.")
        return

    # --- Step 4: Generate ---
    st.subheader("4. Tạo ảnh")
    if st.button(f"🚀 Ghép {len(photo_uploads)} ảnh", type="primary", width="stretch"):
        _generate(template, photo_uploads, box, fit, align, radius, fmt, want_pdf)

    _show_results()


def _generate(template, uploads, box, fit, align, radius, fmt, want_pdf) -> None:
    output_dir = tool_workdir(SLUG) / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    for f in output_dir.iterdir():
        f.unlink()

    names = _unique_names([u.name for u in uploads], fmt)
    files: list[Path] = []
    images: list[Image.Image] = []
    errors: list[str] = []
    progress = st.progress(0.0, text="Đang ghép...")
    for i, (upload, name) in enumerate(zip(uploads, names)):
        progress.progress(i / len(uploads), text=f"({i + 1}/{len(uploads)}) {upload.name}")
        try:
            img = compose(template, load_image(upload), box, fit, align, radius)
        except Exception as e:
            errors.append(f"{upload.name}: {e}")
            continue
        files.append(save_image(img, output_dir / name))
        if want_pdf:
            images.append(img)
    progress.progress(1.0, text=f"Xong {len(files)}/{len(uploads)} ảnh")

    pdf = save_pdf(images, output_dir / "tat_ca.pdf") if images else None
    st.session_state[k("files")] = [str(f) for f in files]
    st.session_state[k("pdf_file")] = str(pdf) if pdf else None
    st.session_state[k("errors")] = errors


def _show_results() -> None:
    files = [Path(f) for f in st.session_state.get(k("files"), []) if Path(f).exists()]
    errors = st.session_state.get(k("errors"), [])
    pdf = st.session_state.get(k("pdf_file"))

    for err in errors:
        st.warning(f"Bỏ qua {err}")
    if not files:
        return

    st.success(f"Đã ghép {len(files)} ảnh.")
    c1, c2 = st.columns(2)
    c1.download_button(
        f"⬇️ Tải zip ({len(files)} ảnh)",
        data=zip_files(files),
        file_name="ghep_anh.zip",
        mime="application/zip",
        width="stretch",
        key=k("dl_zip"),
    )
    if pdf and Path(pdf).exists():
        c2.download_button(
            "⬇️ Tải PDF gộp",
            data=Path(pdf).read_bytes(),
            file_name="ghep_anh.pdf",
            mime="application/pdf",
            width="stretch",
            key=k("dl_pdf"),
        )
    with st.expander("Xem tất cả"):
        cols = st.columns(3)
        for i, f in enumerate(files):
            cols[i % 3].image(str(f), caption=f.name, width="stretch")
