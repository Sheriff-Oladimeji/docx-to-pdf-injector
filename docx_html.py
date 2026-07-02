"""
Convert templates/engagement_letter_template.docx into HTML for headless,
pixel-controlled PDF rendering via WeasyPrint.

Content (paragraph text, bold/italic runs, hyperlinks, list numbering, the
pricing table) is walked directly out of the DOCX XML -- never retyped --
so legal wording can't drift from the source document during the rewrite.
Only layout (CSS in render.py) is hand-authored, tuned against
Pages-rendered reference screenshots of the client's original document.

{{ jinja }} merge tags already live as literal text inside the template's
paragraphs (same tags used by the docxtpl path), so they pass through
untouched and get rendered later by jinja2.
"""

import base64
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# Aptos (Microsoft font, used since the client's doc is authored in
# Word/M365) is bundled here so rendering works on any host, not just one
# with it installed system-wide. Note: Microsoft doesn't grant free
# redistribution of Aptos outside a Windows/Office/M365 install, so treat
# this as a POC convenience, not a cleared-for-production asset -- get
# proper licensing sorted before shipping this to a public/commercial repo.
FONT_DIR = Path(__file__).parent / "templates" / "fonts"
FONT_REGULAR = FONT_DIR / "aptos.ttf"
FONT_BOLD = FONT_DIR / "aptos-bold.ttf"
FONT_ITALIC = FONT_DIR / "aptos-italic.ttf"
FONT_BOLD_ITALIC = FONT_DIR / "aptos-bold-italic.ttf"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# (numId) -> {ilvl: (numFmt, left_pt, hanging_pt)}
# Hand-extracted once from word/numbering.xml (see abstractNum lvl defs).
NUMBERING = {
    "23": {0: ("decimal", 36, 18), 1: ("lowerLetter", 72, 18)},
    "13": {0: ("bullet", 36, 18)},
    "18": {0: ("bullet", 72, 18)},
    "19": {0: ("bullet", 72, 18)},
    "20": {0: ("bullet", 36, 18)},
    "21": {0: ("bullet", 39.25, 17.9)},
    "22": {0: ("bullet", 39.25, 17.9)},
    "24": {0: ("bullet", 36, 18)},
    "25": {0: ("bullet", 36, 18)},
}


# style_id -> (before_pt, after_pt, line_multiplier, font_pt), from styles.xml.
# 232/299 paragraphs in the source doc carry a *direct* pPr/spacing override
# on top of these -- resolve_fmt() reads that override when present.
STYLE_DEFAULTS = {
    None: (0, 8, 278 / 240, 12),
    "Normal": (0, 8, 278 / 240, 12),
    "Heading1": (18, 4, 278 / 240, 20),
    "Heading2": (8, 4, 278 / 240, 16),
    "Heading3": (8, 4, 278 / 240, 14),
}


def resolve_fmt(p: ET.Element, style_id: str | None):
    before_pt, after_pt, line_mult, font_pt = STYLE_DEFAULTS.get(style_id, STYLE_DEFAULTS[None])
    pPr = p.find(W + "pPr")
    # Font size: for a paragraph with visible text, read the first run's own
    # rPr, not w:pPr/w:rPr -- that describes the invisible paragraph-mark
    # character and can drift out of sync with the visible run text (a
    # common Word autofit artifact), which was causing headings to render at
    # the wrong (mark's) size. A genuinely empty paragraph has no runs at
    # all, so its height *is* governed by the mark's own rPr -- use that.
    runs_with_text = [r for r in p.findall(W + "r") if r.find(W + "t") is not None]
    if runs_with_text:
        rPr = runs_with_text[0].find(W + "rPr")
    else:
        rPr = pPr.find(W + "rPr") if pPr is not None else None
    sz_el = rPr.find(W + "sz") if rPr is not None else None
    if sz_el is not None:
        font_pt = int(sz_el.get(W + "val")) / 2
    if pPr is None:
        return before_pt, after_pt, line_mult, font_pt
    spacing_el = pPr.find(W + "spacing")
    if spacing_el is not None:
        if spacing_el.get(W + "before") is not None:
            before_pt = int(spacing_el.get(W + "before")) / 20
        if spacing_el.get(W + "after") is not None:
            after_pt = int(spacing_el.get(W + "after")) / 20
        if spacing_el.get(W + "line") is not None:
            line_mult = int(spacing_el.get(W + "line")) / 240
    return before_pt, after_pt, line_mult, font_pt


def fmt_style_attr(before_pt, after_pt, line_mult, font_pt) -> str:
    # Word adds a paragraph's "space before" to the previous paragraph's
    # "space after" -- it never collapses them. CSS sibling margins *do*
    # collapse (max, not sum) by default, which was silently discarding most
    # "before" spacing. padding-top doesn't collapse, so use that for
    # "before" and keep margin-bottom (simple, no adjacent collapse issue
    # since nothing above it) for "after".
    return (
        f"font-size:{font_pt}pt;line-height:{line_mult};"
        f"padding-top:{before_pt}pt;margin-bottom:{after_pt}pt;"
    )


def esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def marker_for(numFmt: str, n: int) -> str:
    if numFmt == "decimal":
        return f"{n}."
    if numFmt == "lowerLetter":
        return f"{chr(ord('a') + n - 1)}."
    return "•"


class NumberingState:
    def __init__(self):
        self.counters = {}  # numId -> {ilvl: n}

    def next_marker(self, numId: str, ilvl: int) -> str:
        levels = NUMBERING.get(numId, {0: ("bullet", 36, 18)})
        numFmt = levels.get(ilvl, levels[0])[0]
        if numFmt == "bullet":
            return marker_for(numFmt, 0)
        counts = self.counters.setdefault(numId, {})
        counts[ilvl] = counts.get(ilvl, 0) + 1
        # a bump at this level resets any deeper levels (standard Word behaviour)
        for deeper in list(counts):
            if deeper > ilvl:
                counts[deeper] = 0
        return marker_for(numFmt, counts[ilvl])


def run_html(run: ET.Element, hyperlink_href: str | None) -> str:
    bold = run.find(W + "rPr/" + W + "b") is not None
    italic = run.find(W + "rPr/" + W + "i") is not None
    text_parts = []
    for child in run:
        if child.tag == W + "t":
            text_parts.append(esc(child.text or ""))
        elif child.tag == W + "tab":
            text_parts.append("&nbsp;&nbsp;&nbsp;&nbsp;")
        elif child.tag == W + "br":
            text_parts.append("<br/>")
    text = "".join(text_parts)
    if not text:
        return ""
    if bold:
        text = f"<strong>{text}</strong>"
    if italic:
        text = f"<em>{text}</em>"
    if hyperlink_href:
        text = f'<a href="{esc(hyperlink_href)}">{text}</a>'
    return text


def paragraph_inline_html(p: ET.Element, rels: dict) -> str:
    out = []
    for child in list(p):
        if child.tag == W + "r":
            out.append(run_html(child, None))
        elif child.tag == W + "hyperlink":
            rid = child.get(R + "id")
            href = rels.get(rid, "")
            for run in child.findall(W + "r"):
                out.append(run_html(run, href))
    return "".join(out)


def get_numpr(p: ET.Element):
    numPr = p.find(W + "pPr/" + W + "numPr")
    if numPr is None:
        return None
    numId_el = numPr.find(W + "numId")
    ilvl_el = numPr.find(W + "ilvl")
    if numId_el is None:
        return None
    numId = numId_el.get(W + "val")
    ilvl = int(ilvl_el.get(W + "val")) if ilvl_el is not None else 0
    return numId, ilvl


HEADING_TAGS = {"Heading1": "h1", "Heading2": "h2", "Heading3": "h3"}


def convert(docx_path) -> tuple[str, bytes, bytes]:
    """Returns (body_html, logo_header_png_bytes, logo_closing_png_bytes)."""
    with zipfile.ZipFile(docx_path) as z:
        doc = ET.fromstring(z.read("word/document.xml"))
        rels_xml = ET.fromstring(z.read("word/_rels/document.xml.rels"))
        header_logo = z.read("word/media/image1.png")
        closing_logo = z.read("word/media/image2.png")

    rels = {}
    for rel in rels_xml:
        if rel.get("TargetMode") == "External":
            rels[rel.get("Id")] = rel.get("Target")

    body = doc.find(W + "body")
    numbering = NumberingState()
    html = []
    header_logo_b64 = base64.b64encode(header_logo).decode("ascii")

    for el in list(body):
        if el.tag == W + "tbl":
            html.append(render_table(el, rels))
            continue
        if el.tag != W + "p":
            continue

        # the header logo lives in its own (text-less) paragraph; place it
        # inline, at its natural position in the document flow.
        blip = next((n for n in el.iter(W + "blip") if n.get(R + "embed") == "rId8"), None)
        if blip is not None:
            html.append(f'<img class="header-logo" src="data:image/png;base64,{header_logo_b64}">')
            continue

        text = paragraph_inline_html(el, rels)
        pStyle_el = el.find(W + "pPr/" + W + "pStyle")
        style = pStyle_el.get(W + "val") if pStyle_el is not None else None
        numpr = get_numpr(el)

        fmt_attr = fmt_style_attr(*resolve_fmt(el, style))

        if not text.strip() and numpr is None:
            # A genuinely blank paragraph is intentional spacing in the
            # source doc (the author used blank lines as spacers) -- keep it
            # as a blank line rather than dropping the vertical space.
            html.append(f'<p class="blank" style="{fmt_attr}">&nbsp;</p>')
            continue

        if numpr is not None:
            numId, ilvl = numpr
            marker = numbering.next_marker(numId, ilvl)
            fmt, left, hanging = NUMBERING.get(numId, {}).get(ilvl, ("bullet", 36, 18))
            marker_w = hanging
            pad_left = left - hanging
            html.append(
                f'<div class="li" style="{fmt_attr}padding-left:{pad_left}pt;">'
                f'<span class="marker" style="width:{marker_w}pt;">{marker}</span>'
                f'<span class="li-content">{text}</span></div>'
            )
            continue

        tag = HEADING_TAGS.get(style)
        if tag:
            text = text.removeprefix("<br/>").removesuffix("<br/>")
            html.append(f'<{tag} style="{fmt_attr}">{text}</{tag}>')
        else:
            html.append(f'<p style="{fmt_attr}">{text}</p>')

    return "\n".join(html), header_logo, closing_logo


def render_table(tbl: ET.Element, rels: dict) -> str:
    rows_html = []
    trs = tbl.findall(W + "tr")
    for ri, tr in enumerate(trs):
        cells = []
        for tc in tr.findall(W + "tc"):
            cell_paras = tc.findall(W + "p")
            cell_text = "<br/>".join(
                paragraph_inline_html(p, rels) for p in cell_paras if paragraph_inline_html(p, rels)
            )
            cells.append(f"<td>{cell_text}</td>")
        row_class = ' class="hdr"' if ri == 0 else (' class="total"' if ri == len(trs) - 1 else "")
        rows_html.append(f"<tr{row_class}>{''.join(cells)}</tr>")
    return f'<table class="pricing"><colgroup><col style="width:78.6%"/><col style="width:21.4%"/></colgroup>{"".join(rows_html)}</table>'


def _font_face(family: str, path: Path, weight: str, style: str) -> str:
    return (
        f"@font-face {{ font-family: '{family}'; src: url('{path.as_uri()}'); "
        f"font-weight: {weight}; font-style: {style}; }}"
    )


CSS = """
@page { size: A4; margin: 0.5in; }
* { box-sizing: border-box; }
body { font-family: 'Aptos', sans-serif; font-size: 12pt; line-height: 1.158; color: #000; margin: 0; }
/* font-size/line-height/margin-top/margin-bottom are set inline per element
   from the paragraph's actual resolved Word formatting (see resolve_fmt) --
   these rules only cover what doesn't vary per paragraph. */
p { text-align: justify; }
p.blank { text-align: left; }
h1, h2, h3 { font-weight: bold; color: #000; }
.header-logo { display: block; margin: 4pt auto -63pt auto; height: 1.1in; }
.title-rule { border: none; border-top: 1pt solid #999; margin: 4pt 0 12pt 0; }
.li { display: flex; text-align: justify; }
.li .marker { flex-shrink: 0; }
.li .li-content { flex: 1; }
table.pricing { border-collapse: collapse; width: 100%; margin: 4pt 0 8pt 0; }
table.pricing td { border: 0.5pt solid #000; padding: 2pt 6pt; vertical-align: top; }
table.pricing tr.hdr td { text-align: center; }
table.pricing tr.total td { font-weight: bold; }
table.pricing td:last-child { text-align: right; }
table.pricing tr.hdr td:last-child { text-align: center; }
a { color: inherit; text-decoration: underline; }
"""


PAGE_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
{fonts}
{css}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def build_page_html(docx_path) -> str:
    body_html, _header_logo, _closing_logo = convert(docx_path)
    # the title-rule under the H1 is a direct-formatted line in the source
    # doc, not captured by the generic paragraph walk -- splice it in.
    body_html = body_html.replace("</h1>", '</h1><hr class="title-rule">', 1)
    fonts = "\n".join(
        [
            _font_face("Aptos", FONT_REGULAR, "normal", "normal"),
            _font_face("Aptos", FONT_BOLD, "bold", "normal"),
            _font_face("Aptos", FONT_ITALIC, "normal", "italic"),
            _font_face("Aptos", FONT_BOLD_ITALIC, "bold", "italic"),
        ]
    )
    return PAGE_TEMPLATE.format(fonts=fonts, css=CSS, body=body_html)
