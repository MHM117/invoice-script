# Invoice Generator

A small command-line tool that fills the date and number fields in a Word
invoice template and produces one `.docx` per invoice, then converts each to PDF
with LibreOffice. Handy for issuing a run of weekly invoices (e.g. Net 30) in one
go.

## What it does

When you run it, it asks for:

1. **Start date** (`DD/MM/YYYY`)
2. **Starting invoice number** (a whole number)
3. **How many invoices** to generate

Then, for each invoice (counting from 0):

| Field          | Value                                  |
| -------------- | -------------------------------------- |
| Invoice date   | start date + (7 days × index)          |
| Invoice number | starting number + index                |
| Due date       | invoice date + `DUE_DAYS` (default 30) |

So a start of `07/06/2026`, number `15`, count `3` produces three invoices for
07 / 14 / 21 June 2026, each due 30 days later.

Inside each document the invoice number is shown as **`YEAR-NNN`** (e.g.
`2026-015`). Output files are written to `output/`:

```
invoice_015_2026-06-07.docx   # intermediate Word file
015.pdf                       # final PDF (named by zero-padded number)
```

`DUE_DAYS` lives in a named constant at the top of `generate_invoices.py` —
change it to `28`, `14`, etc. if your terms differ.

## The template and its placeholders

The script renders a Word template named **`template.docx`** using
[docxtpl](https://docxtpl.readthedocs.io/). Your template must contain these
three [Jinja2](https://jinja.palletsprojects.com/) placeholders where the date
and number should appear:

```
{{ invoice_number }}
{{ invoice_date }}
{{ due_date }}
```

Everything else (your address, line items, totals, bank details) is fixed and
edited by hand in Word — the script only fills those three fields.

### First-time setup

`template.docx` is **deliberately not in this repo** (it's git-ignored) so your
personal details never get committed. A generic starter,
**`template_example.docx`**, is included instead.

```bash
cp template_example.docx template.docx
```

Then open `template.docx` in Word and replace the `[bracketed]` fields with your
real details (name, address, client, bank details, line items).

> **Tip — keep the placeholders intact.** When editing a `{{ ... }}` field,
> select the whole thing, delete it, and retype it in one go. Word can otherwise
> split the text internally so docxtpl no longer recognises the placeholder.
> Type exactly `{{ invoice_number }}` — two braces each side, one space inside.

## Install

Requires Python 3.8+.

```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python generate_invoices.py
```

Follow the prompts. Bad dates or non-numbers are rejected with a friendly
re-prompt, so a typo won't crash it.

## PDF conversion

After saving each `.docx`, the script converts it to PDF using **LibreOffice**
headless (`soffice --headless --convert-to pdf`). Install it once:

```bash
brew install --cask libreoffice        # macOS
```

(On other platforms, install LibreOffice and make sure `soffice` is on your
`PATH`.) If LibreOffice isn't found, the script prints a clear notice and skips
PDFs — the `.docx` files are still generated, so PDFs are never a hard
requirement.

> **Why not Microsoft Word?** An earlier version used `docx2pdf` (which drives
> Word) first. But Microsoft's newer Mac Word (16.10x and later) removed the
> AppleScript `save as` support it relied on, so it can no longer be scripted to
> export PDFs. LibreOffice is reliable, headless, and needs no per-run
> permission prompts, so it's used directly.

## Files

| File                    | Purpose                                              |
| ----------------------- | ---------------------------------------------------- |
| `generate_invoices.py`  | The script.                                          |
| `template_example.docx` | Generic template to copy and personalise.            |
| `template.docx`         | Your personal template (git-ignored; you create it). |
| `requirements.txt`      | Python dependencies.                                 |
| `output/`               | Generated invoices (git-ignored).                    |
