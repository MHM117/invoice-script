"""Command-line invoice generator.

Fills the three date/number fields in a Word template (``template.docx``) and
produces one .docx per invoice, optionally converting each to PDF.

The template must contain these three Jinja2 placeholders:

    {{ invoice_number }}    {{ invoice_date }}    {{ due_date }}

Everything else in the template (your address, line items, totals, bank
details) is fixed and edited by hand in Word. See README.md for setup.

Usage:
    python generate_invoices.py
"""

import subprocess
import sys
from datetime import date, datetime, timedelta
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

# Candidate commands for the LibreOffice fallback (varies by install/OS).
LIBREOFFICE_COMMANDS = ["soffice", "libreoffice"]


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
        "invoice_number": number,
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


# --- PDF conversion --------------------------------------------------------
def convert_to_pdf(docx_path):
    """Try to convert ``docx_path`` to PDF next to it.

    Tries docx2pdf (Microsoft Word) first, then LibreOffice headless. Returns
    the tool name used on success, or ``None`` if no converter is available.
    Never raises.
    """
    pdf_path = docx_path.with_suffix(".pdf")

    # 1) docx2pdf — uses Microsoft Word (Windows/macOS with Word installed).
    try:
        from docx2pdf import convert

        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists():
            return "docx2pdf"
    except Exception:
        pass  # Word not present or conversion failed; fall through.

    # 2) LibreOffice headless — `soffice --convert-to pdf`.
    for command in LIBREOFFICE_COMMANDS:
        try:
            subprocess.run(
                [command, "--headless", "--convert-to", "pdf",
                 "--outdir", str(docx_path.parent), str(docx_path)],
                check=True,
                capture_output=True,
            )
            if pdf_path.exists():
                return "libreoffice"
        except FileNotFoundError:
            continue  # This command isn't installed; try the next one.
        except subprocess.CalledProcessError:
            break     # LibreOffice exists but failed; no point retrying.

    return None


# --- Main ------------------------------------------------------------------
def main():
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        print(f"Template '{TEMPLATE_FILE}' not found.")
        print("Copy the example and edit it with your details first:")
        print(f'  cp "template.example.docx" {TEMPLATE_FILE}')
        sys.exit(1)

    print("Invoice generator\n-----------------")
    start_date = prompt_date("Start date (DD/MM/YYYY): ")
    start_number = prompt_int("Starting invoice number: ", minimum=0)
    count = prompt_int("How many invoices to generate? ", minimum=1)

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    results = []          # (number, invoice_date, path, pdf_status)
    pdf_available = True  # Flips off once we confirm no converter exists.

    for index in range(count):
        number, invoice_date, due_date = compute_invoice(start_date, start_number, index)
        docx_path = render_invoice(template_path, output_dir, number, invoice_date, due_date)

        pdf_status = "skipped"
        if pdf_available:
            tool = convert_to_pdf(docx_path)
            if tool:
                pdf_status = f"pdf via {tool}"
            else:
                pdf_available = False
                print("\nNo PDF converter found (need Microsoft Word or LibreOffice).")
                print("Skipping PDFs — the .docx files are ready to use.\n")

        results.append((number, invoice_date, docx_path, pdf_status))

    # --- Summary ----------------------------------------------------------
    print(f"\nGenerated {len(results)} invoice(s) in '{output_dir}/':")
    for number, invoice_date, docx_path, pdf_status in results:
        print(f"  #{number:03d}  {invoice_date:{DISPLAY_DATE_FORMAT}}  "
              f"{docx_path.name}  [{pdf_status}]")


if __name__ == "__main__":
    main()
