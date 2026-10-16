from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

from slides.utils import SlideText

IMAGE_PATH = Path(__file__).resolve().parent.parent / "assets" / "fibonacci.jpg"

BULLETS = [
    "Italian mathematician, c. 1170 – c. 1250",
    "Liber Abaci (1202) brought Hindu–Arabic numerals to Europe",
    "The famous number sequence is named after him",
    "His rabbit puzzle is where our story starts →",
]


def build(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only layout
    slide.shapes.title.text = "Leonardo of Pisa (Fibonacci)"

    left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.3), Inches(4.5))
    frame = left.text_frame
    frame.word_wrap = True
    for i, bullet in enumerate(BULLETS):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.text = bullet
        paragraph.level = 0
        paragraph.font.size = Pt(20)
        SlideText.add_bullet(paragraph)
        if i > 0:
            paragraph.space_before = Pt(12)

    # Portrait is near-square (960 × 957); width only, height scales proportionally.
    slide.shapes.add_picture(str(IMAGE_PATH), Inches(6.3), Inches(1.5), width=Inches(3.2))

    caption = slide.shapes.add_textbox(Inches(6.3), Inches(6.3), Inches(3.2), Inches(0.5))
    caption.text_frame.word_wrap = True
    run = caption.text_frame.paragraphs[0]
    run.text = "Portrait: unknown artist, 1905 (public domain, via Wikimedia Commons)"
    run.font.size = Pt(10)
