"""Tool registry.

To add a tool: create ``tools/<name>/`` exposing a ``TOOL`` (see
``tools/base.py``), then append it to ``TOOLS`` below. The web UI builds its
sidebar and home page from this list.
"""

from tools.base import Tool
from tools.giao_an import TOOL as GIAO_AN

TOOLS: list[Tool] = [
    GIAO_AN,
]
