# Nest Sales History

One flat table of sales lines for reporting, read by Frappe Insights. Requires `erpnext`.

## What it holds

**Sales History**: one row per invoice line.

| Group | Fields |
|---|---|
| Where / when | company (the Sage entity the sale posted to), store, posting date, invoice no, type (Invoice / Credit Note) |
| Who | customer, customer category, sales person |
| What | item code, item name, item entity (the item's own entity), item category, collection, design, colour, size |
| How much | qty, amount excl VAT, VAT, amount incl VAT, cost |
| Provenance | source (ERPNext / Sage history / Spreadsheet), link to the ERPNext invoice and line |

The reporting attributes are stored on the row as they were at the time, so a sold-out
piece or a customer that was never in ERPNext still reports under the right category.

**Report Item** and **Report Customer**: the category lists (product category,
collection, design, colour, size; customer category). Loaded through Data Import from
the previous system's lists, and used to stamp every new row. An item or customer not
listed falls back to the ERPNext Item / Customer record.

## How rows get in

- **Live**: every submitted Sales Invoice (sale or credit note) writes its lines in on
  submit and removes them on cancel. `company` follows the Sage posting rule: all items
  from one entity, that entity; more than one, Home.
- **History from Sage** (the chosen route): `pull_sage_invoices(company, from_date, to_date)`
  pulls one Sage company for one window through Pharoh, about six months a call, and
  stamps the rows (`source` = Sage history). Pulling a window again replaces it. The
  Pharoh endpoint contract is documented at the top of that section in `api.py`.
- **History from spreadsheets** (fallback): `load_customers`, `load_items`, `load_sales`
  read the previous system's exports uploaded as Files (`source` = Spreadsheet);
  `delete_source("Spreadsheet")` removes such a load in one go.
- `nest_sales_history.api.rebuild(from_date)` regenerates the ERPNext rows from the
  submitted invoices; history rows are untouched.

## Reports

Built in Frappe Insights on the Sales History table: filters on company, store, date,
customer category, item category, collection and customer, with click-through from
category to customer to invoice lines.
