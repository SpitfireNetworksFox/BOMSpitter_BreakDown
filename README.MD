# Quote Generator (Excel ➜ HTML + PDF)

Convert a customer quote from a **known Excel format** into a polished **HTML** document and a **print-ready PDF** in one command.

* Reads your quote lines (SKU, Qty, Description, Unit/List price, etc.) and optional notes/category columns.&#x20;
* Renders a branded, page-numbered HTML template.&#x20;
* Exports a matching PDF using Playwright/Chromium.&#x20;
* Auto-fills quote meta (contact, incoterms, currency, dates, quote #, etc.) from defaults.&#x20;
* Appends standard quotation notes if provided.&#x20;

---

## Quick start

```bash
# Recommended: create/activate a virtualenv first

pip install -r requirements.txt

# Playwright needs a one-time browser install:
python -m playwright install

# Generate HTML + PDF from your Excel quote
python fire_a_quote.py --excel "Example Quote.xlsx"
```

> The CLI writes both files to **Generated Quotes/** by default and names them like:
> `Generated Quotes/Example Quote - Q-YYMMDDMMSS.html` and `.pdf`.&#x20;

---

## Usage

**Basic:**

```bash
python fire_a_quote.py --excel "Example Quote.xlsx"
```

**Common options (if supported by your version):**

```bash
# Choose a specific output path (HTML; a PDF with same basename is also created)
--out "out/CustomerA - Q-2509190001.html"

# Control rounding of money values (as rendered in the template)
--round-values True|False

# Show list price column (when round-values=False)
--show-list-price True|False

# Force Excel to recalculate formulas (requires xlwings/Excel installed)
--excel-recalc

# Pick sheet names if your workbook differs
--items-sheet "Quote Lines" --header-sheet "Header"
```

All flags above are parsed by the script’s `argparse` interface. The tool uses the first reasonable “items” sheet automatically if you don’t specify one.&#x20;

---

## Expected Excel format

Your workbook should contain a line-items sheet with headers that **roughly** match one of these candidates (case-insensitive, partial matches allowed):

* **SKU:** `sku`, `part`, `item`, `product`, `pn`, `part number`, `mpn`, `partner sku`
* **Quantity:** `qty`, `quantity`, `qnty`, `q'ty`
* **Description:** `description`, `desc`, `product description`, `item description`
* **Unit/List Price:** `unit price`, `price`, `list price`, `unitprice`, `price ea`, `each`
* **Discount:** `discount`
* **Discounted/Net Price:** `discount price`, `discounted price`, `sell price`, `your price`, `net price`
* **Extended/Line Total:** `extended`, `extended price`, `ext`, `subtotal`, `line total`, `amount`
* **Notes (optional):** `notes`, `note`, `comments`, `line notes`
* **Category (optional):** `product type`, `category`, `product category`, `type`

The parser picks the best matching column names, computes unit/extended values if needed (using cached Excel formula results), and skips “Totals” rows.&#x20;

> **Formula support:** The reader uses Excel’s **cached values** (what Excel saved last time). You can optionally force a fresh recalc with `--excel-recalc` (requires Excel + xlwings installed).&#x20;

---

## Output

* **HTML**: Styled with Manrope, compact margins for printing, fixed quote number on the header, optional list price column, “Unit Price”/“Extended” formatting with 0- or 2-decimal display depending on `--round-values`.&#x20;
* **PDF**: Generated via Playwright (Chromium) with Letter size and print backgrounds enabled.&#x20;

By default, files are written to **`Generated Quotes/`** with a name pattern combining the Excel file stem and the current **quote number**. File-name safety is handled across platforms.

---

## Template & notes

* The HTML layout (header/terms/items/notes/signature) lives in `quote_html_template.py`. It also drives column visibility (List vs Unit price) and currency rounding behavior.&#x20;
* Standard quotation notes can be supplied via `quote_notes.py` and will be injected under “STANDARD QUOTATION NOTES”.&#x20;

---

## Defaults & metadata

The project ships with sensible defaults for quote metadata (customer name/address, payment terms, incoterms, currency, duty/tax wording, **auto-generated quote number**, **date/expiry**, etc.). You can edit these in `client_and_account_rep.py`.&#x20;

---

## How it works (internals)

1. **Read Excel:** Load cached formula values via openpyxl; fall back to pandas if needed. Optional xlwings recalc.&#x20;
2. **Parse items:** Normalize headers, select best-match columns, coerce numbers, derive unit/extended totals if missing, sum grand total.&#x20;
3. **Render HTML:** Fill the Jinja2 template with metadata + line items.
4. **Write PDF:** Use Playwright to render the HTML as a print-styled PDF.&#x20;

---

## Troubleshooting

* **PDF step fails or `base_url` error:** The PDF writer supports two modes—load from an HTML **string** (uses `page.set_content`) or from a **file path** (uses `page.goto`). The CLI uses **file-path mode**, which avoids `base_url` issues and is the most reliable. Ensure you’ve run `python -m playwright install` and have a compatible Playwright version.&#x20;
* **Numbers display as text:** The reader strips `$` and commas and coerces to numbers. If a cell still won’t parse, check for stray text or merged cells in your workbook.&#x20;
* **Wrong sheet chosen:** Pass `--items-sheet "Your Sheet Name"` to override auto-selection.&#x20;
* **File name contains odd characters:** Names are sanitized cross-platform; you can also override with `--out`.&#x20;

---

## Requirements

* Python 3.10+
* Dependencies: `pandas`, `openpyxl`, `jinja2`, `playwright` (plus a one-time `python -m playwright install`)
* Optional (for `--excel-recalc`): `xlwings` + local Excel installation.&#x20;

> If your repo includes `requirements.txt`, install it directly with `pip install -r requirements.txt`.

---

## Examples

```bash
# Basic
python fire_a_quote.py --excel "Example Quote.xlsx"

# Show list price (adds a column) and keep cents
python fire_a_quote.py --excel "Example Quote.xlsx" --show-list-price True --round-values False

# Round to whole dollars in the output
python fire_a_quote.py --excel "Example Quote.xlsx" --round-values True

# Force a fresh Excel formula recalc (requires Excel on your OS)
python fire_a_quote.py --excel "Example Quote.xlsx" --excel-recalc

# Custom output path (PDF is written with the same basename)
python fire_a_quote.py --excel "Example Quote.xlsx" --out "out/CustomerA - Special.html"
```

The generated PDF will be placed beside the HTML using the same name (different extension).&#x20;

---

## Customizing the look

* Edit **`quote_html_template.py`** for colors, fonts, column headings, and page margins. The page footer already prints `QUOTE#  page/total`.&#x20;
* Update **`client_and_account_rep.py`** to change default company/contact/terms/quote number/date rules.&#x20;
* Adjust **notes** in `quote_notes.py`.&#x20;

---

## License

It's a virus. Don't use it.