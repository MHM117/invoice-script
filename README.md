# Invoice Generator

A small command-line tool that fills the date and number fields in a Word
invoice template and produces one `.docx` per invoice — optionally converting
each to PDF. Handy for issuing a run of weekly invoices (e.g. Net 30) in one go.

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

So a start of `07/06/2026`, number `15`, count `3` produces invoices
`#015` (07 June 2026), `#016` (14 June 2026), and `#017` (21 June 2026),
each due 30 days later.

Output files are written to `output/`, named like:

```
invoice_015_2026-06-07.docx
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
**`template.example.docx`**, is included instead.

```bash
cp "template.example.docx" template.docx
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

## PDF conversion (optional)

After saving each `.docx`, the script tries to produce a PDF:

1. **[docx2pdf](https://pypi.org/project/docx2pdf/)** — uses Microsoft Word
   (Windows or macOS with Word installed).
2. **LibreOffice** headless (`soffice --convert-to pdf`) — used as a fallback.

If neither is available, it prints a clear notice and skips PDFs — the `.docx`
files are still generated. So PDFs are a bonus, never a requirement.

## Files

| File                    | Purpose                                              |
| ----------------------- | ---------------------------------------------------- |
| `generate_invoices.py`  | The script.                                          |
| `template.example.docx` | Generic template to copy and personalise.            |
| `template.docx`         | Your personal template (git-ignored; you create it). |
| `requirements.txt`      | Python dependencies.                                 |
| `output/`               | Generated invoices (git-ignored).                    |
