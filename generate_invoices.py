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

import contextlib
import os
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

# Candidate commands for the LibreOffice fallback. The bare names cover a
# normal PATH install; the absolute path covers a standard macOS app install
# where `soffice` isn't on PATH.
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
# Each converter takes (docx_path, pdf_path) and returns True on success.
def convert_with_docx2pdf(docx_path, pdf_path):
    """Convert using docx2pdf, which drives Microsoft Word.

    docx2pdf calls ``sys.exit()`` on failure, which raises ``SystemExit`` — and
    that is NOT a subclass of ``Exception``, so we must catch it explicitly or
    it would terminate the whole program. ``KeyboardInterrupt`` is deliberately
    left uncaught so Ctrl-C still works during a slow Word call.

    Note: Microsoft's newer Mac Word (16.10x and later) dropped AppleScript
    "save as" support, so docx2pdf simply cannot work there — it'll fail and we
    fall back to LibreOffice.
    """
    try:
        from docx2pdf import convert

        # docx2pdf prints an error dict and a progress bar; silence both so a
        # failure-then-fallback doesn't clutter our output.
        with open(os.devnull, "w") as devnull, \
                contextlib.redirect_stdout(devnull), \
                contextlib.redirect_stderr(devnull):
            convert(str(docx_path), str(pdf_path))
    except (Exception, SystemExit):
        return False
    return pdf_path.exists()


def convert_with_libreoffice(docx_path, pdf_path):
    """Convert using LibreOffice headless, if `soffice`/`libreoffice` is found."""
    for command in LIBREOFFICE_COMMANDS:
        if shutil.which(command) is None:
            continue
        try:
            subprocess.run(
                [command, "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf_path.parent), str(docx_path)],
                check=True,
                capture_output=True,
            )
        except subprocess.CalledProcessError:
            return False
        return pdf_path.exists()
    return False


class PdfConverter:
    """Converts .docx files to PDF, remembering the first tool that works.

    docx2pdf is tried first (it uses Microsoft Word); LibreOffice is the
    fallback. Once a tool succeeds it's reused for the rest of the run, so we
    don't relaunch a slow or broken converter for every invoice. If nothing
    works, the converter gives up quietly and ``gave_up`` becomes True so the
    caller can report it once.
    """

    def __init__(self, methods=None):
        # (name, function) pairs; injectable so the logic can be unit-tested.
        self._methods = methods or [
            ("docx2pdf", convert_with_docx2pdf),
            ("LibreOffice", convert_with_libreoffice),
        ]
        self._working = None   # cached (name, func) after the first success
        self.gave_up = False   # True once we know no converter is available

    def convert(self, docx_path):
        """Convert ``docx_path`` to a PDF beside it.

        Returns the tool name used, or ``None`` if no converter is available.
        """
        if self.gave_up:
            return None
        pdf_path = docx_path.with_suffix(".pdf")
        candidates = [self._working] if self._working else self._methods
        for name, func in candidates:
            if func(docx_path, pdf_path):
                self._working = (name, func)
                return name
        # Only conclude "nothing works" while we were still trying all tools.
        if self._working is None:
            self.gave_up = True
        return None


# --- Main ------------------------------------------------------------------
def main():
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        print(f"Template '{TEMPLATE_FILE}' not found.")
        print("Copy the example and edit it with your details first:")
        print(f'  cp "template.example.docx" {TEMPLATE_FILE}')
        raise SystemExit(1)

    print("Invoice generator\n-----------------")
    start_date = prompt_date("Start date (DD/MM/YYYY): ")
    start_number = prompt_int("Starting invoice number: ", minimum=0)
    count = prompt_int("How many invoices to generate? ", minimum=1)

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    converter = PdfConverter()
    pdf_notice_shown = False
    results = []  # (number, invoice_date, path, pdf_status)

    for index in range(count):
        number, invoice_date, due_date = compute_invoice(start_date, start_number, index)
        docx_path = render_invoice(template_path, output_dir, number, invoice_date, due_date)

        tool = converter.convert(docx_path)
        if tool:
            pdf_status = f"pdf via {tool}"
        else:
            pdf_status = "docx only"
            if converter.gave_up and not pdf_notice_shown:
                print("\nNo PDF converter available — install LibreOffice, or use a")
                print("Word version with working AppleScript support.")
                print("Skipping PDFs; the .docx files are ready to use.\n")
                pdf_notice_shown = True

        results.append((number, invoice_date, docx_path, pdf_status))

    # --- Summary ----------------------------------------------------------
    print(f"\nGenerated {len(results)} invoice(s) in '{output_dir}/':")
    for number, invoice_date, docx_path, pdf_status in results:
        print(f"  #{number:03d}  {invoice_date:{DISPLAY_DATE_FORMAT}}  "
              f"{docx_path.name}  [{pdf_status}]")


if __name__ == "__main__":
    main()
