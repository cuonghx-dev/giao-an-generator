"""Contract every tool in the toolbox implements."""

from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class Tool:
    """A self-contained tool shown as one page in the web UI.

    Attributes:
        slug: URL path of the page (``/<slug>``); must be unique.
        title: Name shown in the sidebar and on the home page.
        icon: Emoji shown next to the title.
        description: One-line summary shown on the home page card.
        render: Draws the tool's Streamlit page. Called on every rerun.
    """

    slug: str
    title: str
    icon: str
    description: str
    render: Callable[[], None]
