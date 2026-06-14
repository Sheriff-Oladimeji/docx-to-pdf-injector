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
import json
import logging
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document  # bundled with docxtpl
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


def find_soffice() -> str:
    """Locate the LibreOffice binary across macOS / Linux."""
    candidates = [
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            log.info("LibreOffice found: %s", c)
            return c
    sys.exit(
        "LibreOffice not found. Install it:\n"
        "  macOS:  brew install --cask libreoffice\n"
        "  Linux:  sudo apt install libreoffice"
    )


def build_context(data: dict) -> dict:
    """Map the client's fields to template variables, deriving the first name."""
    ctx = dict(data)
    ctx["client_first_name"] = data["client_name"].split()[0]  # "Dear Jonathan,"
    return ctx


def docx_text(path: Path) -> str:
    """Return all paragraph text from a docx, for sanity checks."""
    return "\n".join(p.text for p in Document(path).paragraphs)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data.json", help="path to JSON data file")
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(exist_ok=True)

    # --- Resolve + log inputs -------------------------------------------------
    log.info("Template : %s", TEMPLATE)
    log.info("Data file: %s", ROOT / args.data)

    if not TEMPLATE.exists():
        sys.exit(f"ERROR: template not found at {TEMPLATE}")
    data_path = ROOT / args.data
    if not data_path.exists():
        sys.exit(f"ERROR: data file not found at {data_path}")

    data = json.loads(data_path.read_text())

    # --- Load template and CHECK it actually has merge tags -------------------
    doc = DocxTemplate(TEMPLATE)
    tags = sorted(doc.get_undeclared_template_variables())
    log.info("Merge tags in template: %s", tags or "NONE")

    if not tags:
        log.error("=" * 70)
        log.error("This template has NO {{ }} merge fields.")
        log.error("You're using the ORIGINAL client document, not the TAGGED one.")
        log.error("docxtpl can only replace tags, so the output would just be a")
        log.error("copy of the input (same names, same totals).")
        log.error("FIX: put the tagged engagement_letter_template.docx in templates/")
        log.error("=" * 70)
        sys.exit(1)

    # --- Show exactly what will be injected -----------------------------------
    ctx = build_context(data)
    log.info("Injecting values:")
    for k in sorted(ctx):
        log.info("    %-18s = %s", k, ctx[k])

    # Warn about any tag the data doesn't cover (would render blank)
    missing = [t for t in tags if t not in ctx]
    if missing:
        log.warning("Template tags with NO data (will render blank): %s", missing)

    # --- 1. Inject -------------------------------------------------------------
    doc.render(ctx)
    filled_docx = OUTPUT_DIR / "engagement_letter_filled.docx"
    doc.save(filled_docx)
    log.info("[1/2] Injected data  -> %s", filled_docx.name)

    # --- Verify the injection landed ------------------------------------------
    text = docx_text(filled_docx)
    if data["client_name"] in text:
        log.info("VERIFY ok: '%s' is present in the output.", data["client_name"])
    else:
        log.error(
            "VERIFY FAILED: '%s' not found in output. Check the tags/data.",
            data["client_name"],
        )

    # --- 2. Render PDF ---------------------------------------------------------
    soffice = find_soffice()
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(OUTPUT_DIR),
            str(filled_docx),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log.error("LibreOffice failed:\n%s", result.stderr.strip())
        sys.exit(1)
    log.info("[2/2] Rendered PDF   -> engagement_letter_filled.pdf")
    log.info("Done. Open: %s", OUTPUT_DIR / "engagement_letter_filled.pdf")


if __name__ == "__main__":
    main()
