# AGENTS.md

Slide-deck repo: Fibonacci implementations deck (per `README.md`). One file per slide in `slides/` (each exposes `build(prs)`); `slides/__init__.py` holds the `SLIDES` registry in deck order; `main.py` runs every builder via the `pptx` skill (`Presentation()` → slide layouts → save). Shared helpers live in `slides/utils.py` (`SlideText` for text formatting, `gif_font` for build-time GIFs — not a slide). Only `slides/title.py`, `slides/person.py`, `slides/problem.py`, `slides/recurrence.py`, `slides/recursion.py`, and `slides/memoization.py` exist so far. The recursion slide regenerates its call-tree GIF into `target/` at build time (PIL) — no committed binaries.

## Environment
- Python 3.11 (`pyvenv.cfg`: CPython 3.11.4); repo-local venv at `.venv/` (gitignored, Windows).
- No manifest/lockfile (`requirements.txt`, `pyproject.toml`) and no test/lint/typecheck/CI config. Do not assume pytest, ruff, mypy, or tox exist.
- `.venv` already contains `python-pptx 1.0.2` and `XlsxWriter 3.2.9` (site-packages only, not declared anywhere). If you add a dependency, add a manifest (`requirements.txt` or `pyproject.toml`) so it is reproducible.

## Commands (PowerShell, from repo root)
- Run: `.venv\Scripts\python.exe main.py`
- Inspect installed deps: `.venv\Scripts\pip list`
- Install: `.venv\Scripts\pip install <pkg>` (then record it in a new manifest — see above)

## Notes
- PyCharm project (`.idea/` untracked); entrypoint is `main.py`, which only orchestrates slide builders — put slide content in a new `slides/<name>.py` module and register its `build` in `SLIDES` order instead of editing `main.py`.
- Generated deck goes to `target/` (`target/fibonacci-slides.pptx`), already gitignored by the `target/` line in `.gitignore` — never commit generated `.pptx`/`.xlsx` output.
- Verify with: re-run `main.py` and confirm the `target/` deck opens. There is no test suite to run.
