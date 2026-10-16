"""Memoization slide.

Same recurrence as naive recursion, but each value is computed once with a
for loop and reused from a table: O(2^n) time becomes O(n). The animated GIF
shows the memo table for fib(8) filling left to right (generated with PIL at
build time into target/, plays in slideshow mode).
"""

from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from slides.utils import SlideText, gif_font

TARGET_DIR = Path(__file__).resolve().parent.parent / "target"
GIF_PATH = TARGET_DIR / "memo-table.gif"

PSEUDOCODE = [
    "fib(n):",
    "  memo = [0, 1]",
    "  for i = 2 to n:",
    "    memo[i] = memo[i-1] + memo[i-2]",
    "  return memo[n]",
]

BULLETS = [
    "Same logic — memo[i] is still the sum of the previous two",
    "Each value computed once, then reused: O(n) time",
    "Price: O(n) memory for the table",
]

# GIF layout
CALL_N = 8
IMG_W, IMG_H = 800, 600
BOX_W, BOX_H, BOX_Y, X0 = 76, 110, 220, 40
STEP = (IMG_W - 2 * X0) / (CALL_N + 1)

EMPTY_FILL, EMPTY_LINE = "white", (128, 128, 128)
ORANGE_FILL = (255, 217, 102)
BLUE_FILL, BLUE_LINE = (219, 234, 249), (31, 78, 121)
GREEN_FILL, GREEN_LINE = (198, 224, 180), (56, 118, 29)
INK, GREY = (31, 31, 31), (110, 110, 110)


def _box_x(j: int) -> float:
    return X0 + STEP * (j + 0.5)


def _draw_box(draw: ImageDraw.ImageDraw, j: int, text: str, fill: tuple, line: tuple) -> None:
    x = _box_x(j)
    draw.rectangle([x - BOX_W / 2, BOX_Y, x + BOX_W / 2, BOX_Y + BOX_H], fill=fill, outline=line, width=2)
    draw.text((x, BOX_Y - 28), f"i={j}", font=gif_font(17), fill=GREY, anchor="ma")
    draw.text((x, BOX_Y + BOX_H / 2), text, font=gif_font(30), fill=INK, anchor="mm")


def _draw_frame(values: dict[int, int], highlight: tuple[int, int] | None, status: str) -> Image.Image:
    img = Image.new("RGB", (IMG_W, IMG_H), "white")
    draw = ImageDraw.Draw(img)
    draw.text((IMG_W / 2, 24), f"memo table for fib({CALL_N})", font=gif_font(22), fill=INK, anchor="ma")
    legend = [(ORANGE_FILL, "read"), (GREEN_FILL, "written"), (EMPTY_FILL, "empty")]
    lx = IMG_W / 2 - 160
    for fill, name in legend:
        draw.rectangle([lx, 66, lx + 30, 92], fill=fill, outline=(128, 128, 128), width=2)
        draw.text((lx + 38, 79), name, font=gif_font(17), fill=GREY, anchor="lm")
        lx += 130
    current = highlight[1] if highlight else None
    for j in range(CALL_N + 1):
        if j in values:
            fill, line = (BLUE_FILL, BLUE_LINE) if j < 2 else (GREEN_FILL, GREEN_LINE)
            _draw_box(draw, j, str(values[j]), fill, line)
        else:
            _draw_box(draw, j, "?", EMPTY_FILL, EMPTY_LINE)
    if highlight is not None:
        a, b = highlight
        for j in (a, b):
            _draw_box(draw, j, str(values[j]), ORANGE_FILL, BLUE_LINE)
        # Orange outline on the target box shows where the sum lands.
        cx = _box_x(current)
        draw.rectangle(
            [cx - BOX_W / 2, BOX_Y, cx + BOX_W / 2, BOX_Y + BOX_H],
            outline=(237, 125, 49),
            width=3,
        )
    draw.text((IMG_W / 2, 500), status, font=gif_font(21), fill=INK, anchor="ma")
    return img


def _make_gif() -> None:
    memo = {0: 0, 1: 1}
    frames = [_draw_frame(dict(memo), None, "base cases: memo[0]=0, memo[1]=1")]
    durations = [900]
    for i in range(2, CALL_N + 1):
        a, b = memo[i - 1], memo[i - 2]
        frames.append(_draw_frame(dict(memo), (i - 2, i - 1), f"i={i}: memo[{i}] = memo[{i - 1}] + memo[{i - 2}]"))
        durations.append(800)
        memo[i] = a + b
        frames.append(_draw_frame(dict(memo), None, f"i={i}: memo[{i}] = {a} + {b} = {memo[i]}"))
        durations.append(800)
    # Single hold frame: Pillow merges pixel-identical consecutive frames,
    # so two identical holds would collapse into one.
    frames.append(
        _draw_frame(dict(memo), None, f"fib({CALL_N}) = {memo[CALL_N]} — 9 table steps vs 67 naive calls")
    )
    durations.append(5000)
    frames[0].save(GIF_PATH, save_all=True, append_images=frames[1:], duration=durations, loop=0)


def build(prs: Presentation) -> None:
    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    _make_gif()

    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only layout
    slide.shapes.title.text = "Memoization: Linear Time with a for Loop"

    code = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.2), Inches(1.9))
    code.fill.solid()
    code.fill.fore_color.rgb = RGBColor(243, 244, 246)
    frame = code.text_frame
    frame.word_wrap = True
    for i, line in enumerate(PSEUDOCODE):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.text = line
        run = paragraph.runs[0]
        run.font.size = Pt(18)
        run.font.name = "Consolas"

    box = slide.shapes.add_textbox(Inches(0.5), Inches(3.6), Inches(4.2), Inches(2.6))
    frame = box.text_frame
    frame.word_wrap = True
    for i, bullet in enumerate(BULLETS):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.text = bullet
        paragraph.level = 0
        paragraph.font.size = Pt(20)
        SlideText.add_bullet(paragraph)
        if i > 0:
            paragraph.space_before = Pt(12)

    # GIF is 800×600; width only, height scales proportionally.
    slide.shapes.add_picture(str(GIF_PATH), Inches(5.0), Inches(1.5), width=Inches(4.5))

    caption = slide.shapes.add_textbox(Inches(5.0), Inches(5.1), Inches(4.5), Inches(0.8))
    caption.text_frame.word_wrap = True
    run = caption.text_frame.paragraphs[0]
    run.text = "Memo table for fib(8) filling left to right (animated GIF — plays in slideshow mode)"
    run.font.size = Pt(10)
