"""
Engagement Letter POC — DOCX template -> data injection -> high-fidelity PDF.

Usage:
    uv run generate.py                      # uses data.json
    uv run generate.py --data other.json    # use a different data file

Output (in ./output/):
    - engagement_letter_filled.docx
    - engagement_letter_filled.pdf
"""

import argparse
import io
import json
import logging
import re
import shutil
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from docxtpl import DocxTemplate

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "templates" / "engagement_letter_template.docx"
OUTPUT_DIR = ROOT / "output"

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-5s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("poc")


def find_soffice() -> str:
    for c in ["soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]:
        if shutil.which(c) or Path(c).exists():
            log.info("LibreOffice: %s", c)
            return c
    sys.exit("LibreOffice not found.\n  macOS: brew install --cask libreoffice\n  Linux: sudo apt install libreoffice")


def patch_docx(docx_bytes: bytes) -> bytes:
    """
    Three patches applied to raw DOCX bytes before rendering:

    1. Theme font resolution — replaces majorHAnsi/minorHAnsi with explicit font
       names read from the document theme (e.g. Aptos Display / Aptos).
       LibreOffice doesn't resolve theme font refs, so without this headings
       render in the wrong weight.

    2. docDefault line spacing — marks the default paragraph spacing as explicit
       so LibreOffice honours it instead of falling back to single spacing.

    3. Empty paragraph height — empty paragraphs act as blank lines in Word.
       LibreOffice sometimes collapses them. This ensures they carry the same
       line height as the rest of the document.
    """
    # --- 1. Read theme font names ---
    major_font, minor_font = "Aptos Display", "Aptos"
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if "word/theme/theme1.xml" in z.namelist():
            tree = ET.fromstring(z.read("word/theme/theme1.xml"))
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            mj = tree.find(".//a:fontScheme/a:majorFont/a:latin", ns)
            mn = tree.find(".//a:fontScheme/a:minorFont/a:latin", ns)
            if mj is not None and mj.get("typeface"): major_font = mj.get("typeface")
            if mn is not None and mn.get("typeface"): minor_font = mn.get("typeface")
    log.info("Theme fonts: major=%s  minor=%s", major_font, minor_font)

    def fix_empty_paras(s: str) -> str:
        """Add explicit line height to empty paragraphs so LibreOffice doesn't collapse them."""
        count = [0]
        def _fix(m):
            block = m.group(0)
            if "<w:t" in block or "w:line=" in block:
                return block
            if "<w:pPr>" in block:
                block = block.replace("<w:pPr>", '<w:pPr><w:spacing w:line="278" w:lineRule="auto"/>', 1)
            elif re.search(r"<w:pPr\s", block):
                block = re.sub(r"(<w:pPr[^>]*>)", r'\1<w:spacing w:line="278" w:lineRule="auto"/>', block, count=1)
            else:
                block = re.sub(r"(<w:p(?:\s[^>]*)?>)", r'\1<w:pPr><w:spacing w:line="278" w:lineRule="auto"/></w:pPr>', block, count=1)
            count[0] += 1
            return block
        result = re.sub(r"<w:p[ >].*?</w:p>", _fix, s, flags=re.DOTALL)
        log.info("Empty paragraph fix: %d paragraphs updated", count[0])
        return result

    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zin, \
         zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename.endswith(".xml"):
                s = data.decode("utf-8")

                # Patch 1: theme font resolution
                if item.filename in ("word/styles.xml", "word/document.xml"):
                    s = re.sub(r'w:asciiTheme="majorHAnsi"', f'w:ascii="{major_font}"', s)
                    s = re.sub(r'w:hAnsiTheme="majorHAnsi"',  f'w:hAnsi="{major_font}"', s)
                    s = re.sub(r'w:asciiTheme="minorHAnsi"', f'w:ascii="{minor_font}"', s)
                    s = re.sub(r'w:hAnsiTheme="minorHAnsi"',  f'w:hAnsi="{minor_font}"', s)

                # Patch 2: docDefault line spacing
                if item.filename == "word/styles.xml":
                    s = re.sub(
                        r'(w:after="\d+" w:line="\d+" w:lineRule="auto")(?! w:afterLines)',
                        r'\1 w:afterLines="0"', s, count=1
                    )

                # Patch 3: empty paragraph height
                if item.filename == "word/document.xml":
                    s = fix_empty_paras(s)

                # Patch 4: table cell margins — Word's TableGrid style has vertical
                # cell padding that LibreOffice doesn't apply, making rows too short.
                # Inject explicit tblCellMar so both render rows at the same height.
                if item.filename == "word/document.xml" and "<w:tblCellMar>" not in s:
                    cellmar = (
                        '<w:tblCellMar>'
                        '<w:top w:w="55" w:type="dxa"/>'
                        '<w:left w:w="108" w:type="dxa"/>'
                        '<w:bottom w:w="55" w:type="dxa"/>'
                        '<w:right w:w="108" w:type="dxa"/>'
                        '</w:tblCellMar>'
                    )
                    s = s.replace("</w:tblPr>", cellmar + "</w:tblPr>")
                    log.info("Table cell margins applied")

                data = s.encode("utf-8")

            zout.writestr(item, data)

    return buf.getvalue()


def build_context(data: dict) -> dict:
    ctx = dict(data)
    ctx["client_first_name"] = data["client_name"].split()[0]
    return ctx


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data.json")
    args = parser.parse_args()
    OUTPUT_DIR.mkdir(exist_ok=True)

    log.info("Template : %s", TEMPLATE)
    log.info("Data file: %s", ROOT / args.data)

    if not TEMPLATE.exists():
        sys.exit(f"ERROR: template not found at {TEMPLATE}")
    data_path = ROOT / args.data
    if not data_path.exists():
        sys.exit(f"ERROR: data file not found at {data_path}")

    data = json.loads(data_path.read_text())
    patched = patch_docx(TEMPLATE.read_bytes())

    doc = DocxTemplate(io.BytesIO(patched))
    tags = sorted(doc.get_undeclared_template_variables())
    log.info("Merge tags: %s", tags or "NONE")

    if not tags:
        log.error("Template has NO {{ }} merge fields — wrong file in templates/")
        sys.exit(1)

    ctx = build_context(data)
    log.info("Injecting:")
    for k in sorted(ctx):
        log.info("    %-18s = %s", k, ctx[k])

    missing = [t for t in tags if t not in ctx]
    if missing:
        log.warning("Tags with no data (blank in output): %s", missing)

    doc.render(ctx)
    filled_docx = OUTPUT_DIR / "engagement_letter_filled.docx"
    doc.save(filled_docx)
    log.info("[1/2] Injected -> %s", filled_docx.name)

    text = "\n".join(p.text for p in Document(filled_docx).paragraphs)
    if data["client_name"] in text:
        log.info("VERIFY ok: '%s' in output", data["client_name"])
    else:
        log.error("VERIFY FAILED: '%s' not found", data["client_name"])

    soffice = find_soffice()
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(OUTPUT_DIR), str(filled_docx)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        log.error("LibreOffice failed:\n%s", result.stderr.strip())
        sys.exit(1)

    log.info("[2/2] PDF -> engagement_letter_filled.pdf")
    log.info("Done: %s", OUTPUT_DIR / "engagement_letter_filled.pdf")


if __name__ == "__main__":
    main()
