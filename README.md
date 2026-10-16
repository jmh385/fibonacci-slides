# fibonacci-slides

Python project that generates a PowerPoint slide deck comparing different ways of implementing the Fibonacci algorithm and visualizing them (e.g. naive recursion, memoization, iterative / DP, closed-form).

Built with `python-pptx` following the OpenCode `pptx` skill workflow (layouts for consistency, one idea per slide, 24pt titles / 18pt body minimum).

## Requirements
- Python 3.11, repo-local venv at `.venv/`
- `python-pptx 1.0.2`, `XlsxWriter 3.2.9` (currently installed in `.venv`, not yet declared in a manifest)

## Usage (PowerShell, from repo root)
- Run: `.venv\Scripts\python.exe main.py`
- Output: `target/fibonacci-slides.pptx` (gitignored build output — open in PowerPoint or Google Slides to verify)

## Structure
- `main.py` — builds the full deck: creates a `Presentation()`, runs every builder in `slides.SLIDES` in order, saves to `target/`.
- `slides/` — one file per slide, each exposing `build(prs)`. Deck order is the `SLIDES` list in `slides/__init__.py`. Shared helpers live in `slides/utils.py` (`SlideText`, `gif_font`; not a slide).
- To add a slide: create `slides/<name>.py` with a `build(prs)` function, then append it to `SLIDES`.

## Status
Nine slides: title (`slides/title.py`), Fibonacci intro with portrait (`slides/person.py`), rabbit problem with bunny photo (`slides/problem.py`), recurrence with value table (`slides/recurrence.py`), naive recursion with pseudocode and animated call-tree GIF (`slides/recursion.py`), memoization with for-loop pseudocode and animated memo-table GIF (`slides/memoization.py`), two-variable iteration with rolling-window diagram (`slides/iterative.py`), Binet's formula with portrait and rendered fraction (`slides/binet.py`), and an all-algorithms roadmap finale (`slides/overview.py`). Images, GIFs, and diagrams are regenerated into `target/` on every build.
