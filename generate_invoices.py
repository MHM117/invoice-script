"""Command-line invoice generator.

Fills the three date/number fields in a Word template (``template.docx``) and
produces one .docx per invoice, then converts each to PDF with LibreOffice.

The template must contain these three Jinja2 placeholders:

    {{ invoice_number }}    {{ invoice_date }}    {{ due_date }}

Everything else in the template (your address, line items, totals, bank
details) is fixed and edited by hand in Word. See README.md for setup.

Usage:
    python generate_invoices.py
"""

import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

from docxtpl import DocxTemplate

# --- Configuration ---------------------------------------------------------
# Change these if your needs differ; they're collected here on purpose.
DUE_DAYS = 30                       # Net 30. Set to 28, 14, etc. as needed.
WEEK_INTERVAL_DAYS = 7              # Spacing between consecutive invoices.

TEMPLATE_FILE = "template.docx"     # Your personal template (kept out of git).
OUTPUT_DIR = "output"              # Generated files land here.

INPUT_DATE_FORMAT = "%d/%m/%Y"      # How you type dates:   07/06/2026
DISPLAY_DATE_FORMAT = "%d %B %Y"    # How they're rendered: 07 June 2026

# Candidate commands for LibreOffice. The bare names cover a normal PATH
# install; the absolute path covers a standard macOS app install where
# `soffice` isn't on PATH.
LIBREOFFICE_COMMANDS = [
    "soffice",
    "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


# --- Input helpers ---------------------------------------------------------
def prompt_date(message):
    """Ask for a date in DD/MM/YYYY and re-prompt until it's valid."""
    while True:
        raw = input(message).strip()
        try:
            return datetime.strptime(raw, INPUT_DATE_FORMAT).date()
        except ValueError:
            print(f"  '{raw}' isn't a valid date. Use DD/MM/YYYY, e.g. 07/06/2026.")


def prompt_int(message, minimum=None):
    """Ask for a whole number and re-prompt until it's valid."""
    while True:
        raw = input(message).strip()
        try:
            value = int(raw)
        except ValueError:
            print(f"  '{raw}' isn't a whole number. Enter digits only, e.g. 15.")
            continue
        if minimum is not None and value < minimum:
            print(f"  Please enter a number of {minimum} or greater.")
            continue
        return value


# --- Invoice maths ---------------------------------------------------------
def compute_invoice(start_date, start_number, index):
    """Return (number, invoice_date, due_date) for the invoice at ``index``.

    Invoices are spaced one week apart and each is due ``DUE_DAYS`` later.
    """
    invoice_date = start_date + timedelta(days=WEEK_INTERVAL_DAYS * index)
    number = start_number + index
    due_date = invoice_date + timedelta(days=DUE_DAYS)
    return number, invoice_date, due_date


def render_invoice(template_path, output_dir, number, invoice_date, due_date):
    """Render one invoice to a .docx and return its path."""
    context = {
        # Shown inside the document as YEAR-NNN, e.g. 2026-016.
        "invoice_number": f"{invoice_date.year}-{number:03d}",
        "invoice_date": invoice_date.strftime(DISPLAY_DATE_FORMAT),
        "due_date": due_date.strftime(DISPLAY_DATE_FORMAT),
    }
    # A fresh DocxTemplate per render avoids state leaking between invoices.
    doc = DocxTemplate(template_path)
    doc.render(context)

    filename = f"invoice_{number:03d}_{invoice_date:%Y-%m-%d}.docx"
    out_path = output_dir / filename
    doc.save(out_path)
    return out_path


# --- PDF conversion (LibreOffice) ------------------------------------------
def find_libreoffice():
    """Return the path to a usable LibreOffice command, or None if not found."""
    for command in LIBREOFFICE_COMMANDS:
        path = shutil.which(command)
        if path:
            return path
    return None


def convert_to_pdf(soffice, docx_path, pdf_path):
    """Convert ``docx_path`` to ``pdf_path`` using LibreOffice headless.

    LibreOffice always names its output after the input file, so we convert into
    the target folder and then rename to the desired ``pdf_path``. Returns True
    on success, False if the conversion fails.
    """
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(pdf_path.parent), str(docx_path)],
            check=True,
            capture_output=True,
        )
    except subprocess.CalledProcessError:
        return False

    produced = pdf_path.parent / (docx_path.stem + ".pdf")
    if produced != pdf_path and produced.exists():
        produced.replace(pdf_path)
    return pdf_path.exists()


# --- Main ------------------------------------------------------------------
def main():
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        print(f"Template '{TEMPLATE_FILE}' not found.")
        print("Copy the example and edit it with your details first:")
        print(f"  cp template_example.docx {TEMPLATE_FILE}")
        raise SystemExit(1)

    print("Invoice generator\n-----------------")
    start_date = prompt_date("Start date (DD/MM/YYYY): ")
    start_number = prompt_int("Starting invoice number: ", minimum=0)
    count = prompt_int("How many invoices to generate? ", minimum=1)

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    soffice = find_libreoffice()
    if soffice is None:
        print("\nLibreOffice not found — PDFs will be skipped (.docx files still")
        print("created). Install it with:  brew install --cask libreoffice\n")

    results = []  # (number, invoice_date, docx_path, pdf_status)
    for index in range(count):
        number, invoice_date, due_date = compute_invoice(start_date, start_number, index)
        docx_path = render_invoice(template_path, output_dir, number, invoice_date, due_date)

        pdf_status = "docx only"
        if soffice:
            pdf_path = output_dir / f"{number:03d}.pdf"
            if convert_to_pdf(soffice, docx_path, pdf_path):
                pdf_status = pdf_path.name

        results.append((number, invoice_date, docx_path, pdf_status))

    # --- Summary ----------------------------------------------------------
    print(f"\nGenerated {len(results)} invoice(s) in '{output_dir}/':")
    for number, invoice_date, docx_path, pdf_status in results:
        print(f"  {invoice_date.year}-{number:03d}  {invoice_date:{DISPLAY_DATE_FORMAT}}  "
              f"{docx_path.name}  [{pdf_status}]")


if __name__ == "__main__":
    main()
