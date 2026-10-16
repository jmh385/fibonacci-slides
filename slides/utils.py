from lxml.etree import QName
from PIL import ImageDraw, ImageFont
from pptx.oxml.ns import qn
from pptx.oxml.text import CT_TextParagraphProperties
from pptx.oxml.xmlchemy import OxmlElement
from pptx.text.text import _Paragraph


def gif_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    """Font for build-time PIL GIFs (Arial on Windows, PIL default elsewhere)."""
    try:
        return ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", size)
    except OSError:
        return ImageFont.load_default()


def draw_fraction(
    draw: ImageDraw.ImageDraw,
    x: float,
    y_top: float,
    num: str,
    den: str,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    color: tuple[int, int, int] = (31, 31, 31),
    pad: int = 14,
    rule: int = 4,
) -> tuple[float, float]:
    """Draw num over a full-width division line over den, starting at left x.

    Returns (width, total height). The rule spans the wider of the two parts.
    """
    nb = draw.textbbox((0, 0), num, font=font)
    db = draw.textbbox((0, 0), den, font=font)
    width = max(nb[2] - nb[0], db[2] - db[0])
    cx = x + width / 2
    num_h = nb[3] - nb[1]
    draw.text((cx, y_top), num, font=font, fill=color, anchor="ma")
    line_y = y_top + num_h + pad
    draw.line([(x, line_y), (x + width, line_y)], fill=color, width=rule)
    den_h = db[3] - db[1]
    draw.text((cx, line_y + 10), den, font=font, fill=color, anchor="ma")
    return width, (line_y + 10 + den_h) - y_top


class SlideText:
    """Shared text-formatting helpers for slide builders."""

    # Canonical child order of a:pPr per the DrawingML schema (EG_TextParagraphProperties).
    # PowerPoint ignores bullet formatting when these elements are out of order,
    # which happens easily because python-pptx setters append new children at the end.
    _PPR_CHILD_ORDER = (
        "lnSpc",
        "spcBef",
        "spcAft",
        "buClrTx",
        "buSzTx",
        "buSzPct",
        "buSzPts",
        "buFontTx",
        "buFont",
        "buNone",
        "buAutoNum",
        "buChar",
        "tabLst",
        "defRPr",
        "extLst",
    )

    @staticmethod
    def add_bullet(paragraph: _Paragraph) -> None:
        """Give a plain-textbox paragraph a real bullet (textboxes have none by default)."""
        pPr = paragraph._p.get_or_add_pPr()
        for tag in ("buNone", "buAutoNum"):
            existing = pPr.find(qn(f"a:{tag}"))
            if existing is not None:
                pPr.remove(existing)
        if pPr.find(qn("a:buChar")) is None:
            bu_char = OxmlElement("a:buChar")
            bu_char.set("char", "•")
            pPr.append(bu_char)
        if pPr.find(qn("a:buSzPct")) is None:
            bu_size = OxmlElement("a:buSzPct")
            bu_size.set("val", "100000")
            pPr.append(bu_size)
        pPr.set("marL", "228600")  # 0.25" left margin + hanging indent
        pPr.set("indent", "-228600")
        SlideText._sort_pPr_children(pPr)

    @staticmethod
    def _sort_pPr_children(pPr: CT_TextParagraphProperties) -> None:
        """Reorder a:pPr children into schema order; unknown elements keep their place at the end."""
        order = SlideText._PPR_CHILD_ORDER
        children = list(pPr)
        for child in children:
            pPr.remove(child)
        children.sort(
            key=lambda el: order.index(QName(el).localname)
            if QName(el).localname in order
            else len(order)
        )
        for child in children:
            pPr.append(child)
