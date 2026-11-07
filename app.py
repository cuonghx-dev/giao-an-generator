"""Streamlit entry point: a home page plus one page per registered tool.

Run with: streamlit run app.py
"""

import streamlit as st

from tools import TOOLS

APP_TITLE = "Teacher Toolbox"

st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="centered")


def home() -> None:
    st.title(f"🧰 {APP_TITLE}")
    st.caption("Chọn công cụ ở thanh bên hoặc bên dưới.")
    for tool in TOOLS:
        with st.container(border=True):
            st.page_link(pages[tool.slug], label=f"**{tool.title}**", icon=tool.icon)
            st.write(tool.description)


pages = {
    tool.slug: st.Page(tool.render, title=tool.title, icon=tool.icon, url_path=tool.slug)
    for tool in TOOLS
}
home_page = st.Page(home, title="Trang chủ", icon="🏠", default=True)

st.navigation({"": [home_page], "Công cụ": list(pages.values())}).run()
