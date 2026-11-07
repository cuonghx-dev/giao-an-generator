# Teacher Toolbox

A collection of small tools for teachers, served as one Streamlit app with one
page per tool. Each tool also has its own CLI.

| Tool | Page | CLI | Docs |
|------|------|-----|------|
| Soạn giáo án | `/giao-an` | `python3 -m tools.giao_an` | [docs/giao_an.md](docs/giao_an.md) |

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Layout

```
app.py                 # entry: home page + sidebar built from the registry
tools/
  __init__.py          # TOOLS registry
  base.py              # Tool dataclass (slug, title, icon, description, render)
  common.py            # shared helpers: uploads, zip, per-tool session state
  giao_an/
    __init__.py        # exports TOOL
    __main__.py        # CLI: python3 -m tools.giao_an
    generator.py       # core logic, no Streamlit dependency
    page.py            # Streamlit page
docs/                  # one doc per tool
examples/              # sample input files
```

## Adding a tool

1. Create `tools/<name>/` with:
   - `generator.py` (or similar) — the core logic, free of Streamlit, so it
     can be tested and run from the CLI.
   - `page.py` with a `render()` function that draws the page.
   - `__init__.py` exporting `TOOL = Tool(slug=..., title=..., icon=...,
     description=..., render=...)`. Import `page` lazily inside `render` (see
     `tools/giao_an/__init__.py`).
   - Optionally `__main__.py` for a CLI.
2. Append the tool to `TOOLS` in `tools/__init__.py`.
3. In `page.py`, namespace session-state and widget keys with
   `tools.common.tool_state(slug)` and use `tool_workdir(slug)` for temp
   files. Session state is shared across pages, so unprefixed keys collide.
4. Add `docs/<name>.md` and a row to the table above.
