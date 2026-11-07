"""Giáo án generator: builds lesson plans from a weekly schedule + "KH chi tiết"."""

from tools.base import Tool


def _render() -> None:
    # Imported lazily so loading the registry doesn't pull in every tool's UI.
    from tools.giao_an.page import render

    render()


TOOL = Tool(
    slug="giao-an",
    title="Soạn giáo án",
    icon="📚",
    description="Tạo giáo án từ kế hoạch giảng dạy theo tuần và file nội dung chi tiết.",
    render=_render,
)
