"""Ghép ảnh: batch-place photos into the photo frame of a template image."""

from tools.base import Tool


def _render() -> None:
    # Imported lazily so loading the registry doesn't pull in every tool's UI.
    from tools.ghep_anh.page import render

    render()


TOOL = Tool(
    slug="ghep-anh",
    title="Ghép ảnh hàng loạt",
    icon="🖼️",
    description="Ghép nhiều ảnh vào khung ảnh của một mẫu thiết kế, mỗi ảnh ra một trang.",
    render=_render,
)
