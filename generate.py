"""
Engagement Letter — DOCX template -> data injection -> high-fidelity PDF.

Usage:
    uv run generate.py                         # uses data.json
    uv run generate.py --data other.json       # different data file
    uv run generate.py --renderer pages        # Apple Pages via AppleScript (macOS, default)
    uv run generate.py --renderer word         # Word via docx2pdf (macOS/Windows)
    uv run generate.py --renderer libreoffice  # LibreOffice headless

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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-5s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("poc")


def render_with_word(filled_docx: Path, output_dir: Path) -> Path:
    try:
        from docx2pdf import convert
    except ImportError:
        log.error("docx2pdf not installed. Run: uv add docx2pdf")
        sys.exit(1)
    out_pdf = output_dir / "engagement_letter_filled.pdf"
    log.info("Rendering with Word (docx2pdf)...")
    convert(str(filled_docx), str(out_pdf))
    log.info("[2/2] PDF (Word) -> %s", out_pdf.name)
    return out_pdf


def render_with_pages(filled_docx: Path, output_dir: Path) -> Path:
    """
    Render via Apple Pages (AppleScript), macOS only.

    LibreOffice's layout engine computes paragraph spacing, table cell
    padding and line-height differently from Word/Pages, so even byte-
    identical DOCX XML produces different pagination between engines.
    Pages is the app used to view/compare the original document, so
    rendering through it (rather than LibreOffice) gives a 1:1 match by
    construction instead of chasing per-element spacing patches.
    """
    out_pdf = output_dir / "engagement_letter_filled.pdf"
    docx_path = str(filled_docx.resolve())
    pdf_path = str(out_pdf.resolve())
    log.info("Rendering with Pages (AppleScript)...")
    script = f'''
    tell application "Pages"
        repeat with d in documents
            try
                if POSIX path of (file of d) is "{docx_path}" then
                    close d saving no
                end if
            end try
        end repeat
        set theDoc to open POSIX file "{docx_path}"
        export theDoc to POSIX file "{pdf_path}" as PDF
        close theDoc saving no
    end tell
    '''
    result = subprocess.run(
        ["osascript", "-e", script], capture_output=True, text=True
    )
    if result.returncode != 0:
        log.error("Pages export failed:\n%s", result.stderr.strip())
        sys.exit(1)
    log.info("[2/2] PDF (Pages) -> %s", out_pdf.name)
    return out_pdf


def render_with_libreoffice(filled_docx: Path, output_dir: Path) -> Path:
    candidates = [
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    soffice = next((c for c in candidates if shutil.which(c) or Path(c).exists()), None)
    if not soffice:
        sys.exit("LibreOffice not found.\n  macOS: brew install --cask libreoffice")
    log.info("Rendering with LibreOffice: %s", soffice)
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(filled_docx),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log.error("LibreOffice failed:\n%s", result.stderr.strip())
        sys.exit(1)
    log.info("[2/2] PDF (LibreOffice) -> engagement_letter_filled.pdf")
    return output_dir / "engagement_letter_filled.pdf"


def patch_docx(docx_bytes: bytes) -> bytes:
    """
    Two targeted patches applied to raw DOCX bytes before rendering.
    Verified directly against the original untagged source file (not a
    re-rendered copy) — the title underline gap at the DEFAULT value (80)
    already matches the original exactly (33.7pt). No Heading1 patch needed.

    1. Theme font resolution — replaces majorHAnsi/minorHAnsi with explicit
       font names from the document theme (Aptos Display / Aptos).
       Without this LibreOffice renders headings in the wrong weight.

    2. docDefault afterLines=0 — makes the docDefault line spacing explicit
       so LibreOffice honours it instead of adding extra contextual space.
    """
    # Read theme fonts
    major_font, minor_font = "Aptos Display", "Aptos"
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if "word/theme/theme1.xml" in z.namelist():
            tree = ET.fromstring(z.read("word/theme/theme1.xml"))
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            mj = tree.find(".//a:fontScheme/a:majorFont/a:latin", ns)
            mn = tree.find(".//a:fontScheme/a:minorFont/a:latin", ns)
            if mj is not None and mj.get("typeface"):
                major_font = mj.get("typeface")
            if mn is not None and mn.get("typeface"):
                minor_font = mn.get("typeface")
    log.info("Theme fonts: major=%s  minor=%s", major_font, minor_font)

    buf = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(docx_bytes)) as zin,
        zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                s = data.decode("utf-8")
                # Patch 1: theme font resolution
                if item.filename in ("word/styles.xml", "word/document.xml"):
                    s = re.sub(
                        r'w:asciiTheme="majorHAnsi"', f'w:ascii="{major_font}"', s
                    )
                    s = re.sub(
                        r'w:hAnsiTheme="majorHAnsi"', f'w:hAnsi="{major_font}"', s
                    )
                    s = re.sub(
                        r'w:asciiTheme="minorHAnsi"', f'w:ascii="{minor_font}"', s
                    )
                    s = re.sub(
                        r'w:hAnsiTheme="minorHAnsi"', f'w:hAnsi="{minor_font}"', s
                    )
                # Patch 2: docDefault line spacing explicit
                if item.filename == "word/styles.xml":
                    s = re.sub(
                        r'(w:after="\d+" w:line="\d+" w:lineRule="auto")(?! w:afterLines)',
                        r'\1 w:afterLines="0"',
                        s,
                        count=1,
                    )
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
    parser.add_argument(
        "--renderer",
        choices=["pages", "word", "libreoffice"],
        default="pages",
    )
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(exist_ok=True)
    log.info("Template  : %s", TEMPLATE)
    log.info("Data file : %s", ROOT / args.data)
    log.info("Renderer  : %s", args.renderer)

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
        log.error("Template has NO {{ }} merge fields — check templates/")
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

    if args.renderer == "word":
        render_with_word(filled_docx, OUTPUT_DIR)
    elif args.renderer == "libreoffice":
        render_with_libreoffice(filled_docx, OUTPUT_DIR)
    else:
        render_with_pages(filled_docx, OUTPUT_DIR)

    log.info("Done: %s", OUTPUT_DIR / "engagement_letter_filled.pdf")


if __name__ == "__main__":
    main()
