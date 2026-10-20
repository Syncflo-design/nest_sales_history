"""Sales History: one flat row per invoice line, for reporting.

Live rows: every submitted Sales Invoice (sale or credit note) writes its lines in
here on submit and removes them on cancel. History rows: loaded once from the
previous system through Data Import. Reports (Frappe Insights) read this table only.

Each row carries the item's and customer's reporting attributes AS STORED AT THE
TIME, so a sold-out piece or a customer that no longer exists in ERPNext still
reports correctly. Report Item and Report Customer hold those attributes; an item or
customer missing from them falls back to the ERPNext Item / Customer record.
"""

import frappe
from frappe.utils import flt

# A sale with items from more than one entity posts in full to this Sage company
# (agreed on site 14 Oct 2026). Mirrors post-taxinvoice-to-sage on the Ardmore site.
MULTI_ENTITY_COMPANY = "Home"


def on_invoice_submit(doc, method=None):
	write_lines(doc)


def on_invoice_cancel(doc, method=None):
	remove_lines(doc.name)


def remove_lines(invoice_name):
	frappe.db.delete("Sales History", {"sales_invoice": invoice_name})


def write_lines(doc):
	"""Replace the invoice's rows in Sales History. Safe to run again."""
	remove_lines(doc.name)
	company = _sage_company(doc)
	store = doc.get("pos_profile") or doc.get("set_warehouse") or ""
	cust_cat = _customer_category(doc.customer, doc.get("customer_name"))
	net_total = flt(doc.net_total)
	ratio = (flt(doc.grand_total) / net_total) if net_total else 1.0
	kind = "Credit Note" if doc.get("is_return") else "Invoice"
	for it in doc.items:
		attrs = _item_attrs(it.item_code, it.item_name)
		excl = flt(it.base_net_amount if it.get("base_net_amount") is not None else it.net_amount)
		incl = round(excl * ratio, 2)
		cost = _line_cost(it)
		frappe.get_doc(
			{
				"doctype": "Sales History",
				"company": company,
				"store": store,
				"posting_date": doc.posting_date,
				"invoice_no": doc.name,
				"invoice_type": kind,
				"customer": doc.get("customer_name") or doc.customer,
				"customer_category": cust_cat,
				"sales_person": doc.get("custom_sales_person") or doc.get("custom_sales_representative") or "",
				"source": "ERPNext",
				"item_code": it.item_code,
				"item_name": attrs["item_name"] or it.item_name,
				"item_entity": attrs["item_entity"],
				"item_category": attrs["item_category"],
				"collection": attrs["collection"],
				"design": attrs["design"],
				"colour": attrs["colour"],
				"size": attrs["size"],
				"qty": flt(it.qty),
				"amount_excl": excl,
				"vat": round(incl - excl, 2),
				"amount_incl": incl,
				"cost": cost,
				"sales_invoice": doc.name,
				"sales_invoice_item": it.name,
			}
		).insert(ignore_permissions=True)


def _sage_company(doc):
	"""Where the sale posts in Sage: the one entity its items belong to, else Home."""
	entities = []
	for it in doc.items:
		e = frappe.db.get_value("Item", it.item_code, "custom_category")
		if e and e not in entities:
			entities.append(e)
	if len(entities) == 1:
		return entities[0]
	if len(entities) > 1:
		return MULTI_ENTITY_COMPANY
	return doc.company


def _item_attrs(item_code, item_name=None):
	r = frappe.db.get_value(
		"Report Item",
		item_code,
		["item_name", "item_entity", "item_category", "collection", "design", "colour", "size"],
		as_dict=True,
	)
	if r:
		return r
	i = frappe.db.get_value(
		"Item",
		item_code,
		["item_name", "custom_category", "custom_type", "item_group", "custom_collection", "custom_design", "custom_colour", "custom_size"],
		as_dict=True,
	) or {}
	category = i.get("custom_type") or ""
	if not category or category == "All Item Groups":
		category = i.get("item_group") or ""
	if category == "All Item Groups":
		category = ""
	return {
		"item_name": i.get("item_name") or item_name or "",
		"item_entity": i.get("custom_category") or "",
		"item_category": category,
		"collection": i.get("custom_collection") or "",
		"design": i.get("custom_design") or "",
		"colour": i.get("custom_colour") or "",
		"size": i.get("custom_size") or "",
	}


def _customer_category(customer, customer_name=None):
	for key in (customer_name, customer):
		if key:
			cat = frappe.db.get_value("Report Customer", key, "customer_category")
			if cat:
				return cat
	return frappe.db.get_value("Customer", customer, "customer_group") or ""


def _line_cost(it):
	rate = flt(it.get("incoming_rate"))
	if not rate:
		rate = flt(frappe.db.get_value("Item", it.item_code, "valuation_rate"))
	return round(rate * flt(it.qty), 2)


@frappe.whitelist()
def rebuild(from_date=None):
	"""Regenerate the ERPNext rows from every submitted invoice (optionally from a date).
	History rows from other sources are left alone."""
	frappe.only_for("System Manager")
	filters = {"docstatus": 1}
	if from_date:
		filters["posting_date"] = [">=", from_date]
	names = frappe.get_all("Sales Invoice", filters=filters, pluck="name", order_by="posting_date asc")
	for name in names:
		write_lines(frappe.get_doc("Sales Invoice", name))
	return {"invoices": len(names)}


# ---- loading the previous system's exports --------------------------------------------
#
# Three spreadsheets, uploaded as Files on the site, then one call each:
#   load_customers(file_url)                  Customer Category.xlsx  -> Report Customer
#   load_items(file_url)                      Product Category.xlsx   -> Report Item
#   load_sales(file_url, company, source)     Sales by Customer.xlsx  -> Sales History
# load_sales runs in the background (135k lines) and stamps every row from the two
# lists, so load the lists first. Rows carry `source`, so one delete removes a load:
#   delete_source("Spreadsheet")


def _sheet_rows(file_url, sheet=None):
	"""Rows of one sheet of an uploaded .xlsx (the active sheet when none is named)."""
	from openpyxl import load_workbook

	path = frappe.get_doc("File", {"file_url": file_url}).get_full_path()
	wb = load_workbook(filename=path, read_only=True, data_only=True)
	ws = wb[sheet] if sheet else wb.active
	rows = [list(r) for r in ws.iter_rows(values_only=True) if any(v not in (None, "") for v in r)]
	wb.close()
	if not rows:
		frappe.throw("The sheet is empty.")
	return rows


def _s(v):
	return str(v).strip() if v is not None else ""


@frappe.whitelist()
def load_customers(file_url):
	"""Customer Category.xlsx: Name, Category, Balance, Contact Name, Tel, Mobile, Active."""
	frappe.only_for("System Manager")
	rows = _sheet_rows(file_url)
	created = updated = 0
	for r in rows[1:]:
		name = _s(r[0])
		if not name:
			continue
		cat = _s(r[1]) if len(r) > 1 else ""
		active = 0 if (len(r) > 6 and _s(r[6]).lower() == "no") else 1
		if frappe.db.exists("Report Customer", name):
			frappe.db.set_value("Report Customer", name, {"customer_category": cat, "active": active}, update_modified=False)
			updated += 1
		else:
			frappe.get_doc({"doctype": "Report Customer", "customer_name": name, "customer_category": cat, "active": active}).insert(ignore_permissions=True)
			created += 1
	frappe.db.commit()
	return {"created": created, "updated": updated}


@frappe.whitelist()
def load_items(file_url):
	"""Product Category.xlsx (Sheet1): Code, Description, Collection, Design, Colour, Size,
	Base Cloth, Category, Price Excl, Price Incl, Avg Cost, Last Cost, Qty On Hand, Active."""
	frappe.only_for("System Manager")
	rows = _sheet_rows(file_url, "Sheet1")
	created = updated = 0
	for r in rows[1:]:
		code = _s(r[0])
		if not code:
			continue
		g = lambda i: _s(r[i]) if len(r) > i else ""  # noqa: E731
		vals = {
			"item_name": g(1), "collection": g(2), "design": g(3), "colour": g(4), "size": g(5),
			"base_cloth": g(6), "item_category": g(7),
			"item_entity": "Ceramics" if g(7) == "Ceramics" else "Home",
			"active": 0 if g(13).lower() == "no" else 1,
		}
		if frappe.db.exists("Report Item", code):
			frappe.db.set_value("Report Item", code, vals, update_modified=False)
			updated += 1
		else:
			vals.update({"doctype": "Report Item", "item_code": code})
			frappe.get_doc(vals).insert(ignore_permissions=True)
			created += 1
	frappe.db.commit()
	return {"created": created, "updated": updated}


@frappe.whitelist()
def load_sales(file_url, company, source="Spreadsheet", sheet="Sheet1", vat_rate=15):
	"""Sales by Customer.xlsx: Customer, Date, Invoice, Code, Quantity, Total Selling Price.
	Runs in the background; check Sales History for rows with this source."""
	frappe.only_for("System Manager")
	frappe.enqueue(
		"nest_sales_history.api._load_sales_job",
		queue="long", timeout=7200,
		file_url=file_url, company=company, source=source, sheet=sheet, vat_rate=float(vat_rate),
	)
	return {"queued": True}


def _load_sales_job(file_url, company, source, sheet, vat_rate):
	from frappe.utils import getdate, now

	rows = _sheet_rows(file_url, sheet)
	items = {d.name: d for d in frappe.get_all("Report Item", fields=["name", "item_name", "item_entity", "item_category", "collection", "design", "colour", "size"])}
	custs = {d.name: d.customer_category for d in frappe.get_all("Report Customer", fields=["name", "customer_category"])}
	stamp = now()
	user = frappe.session.user
	fields = ["name", "owner", "creation", "modified", "modified_by", "docstatus", "idx",
		"company", "store", "posting_date", "invoice_no", "invoice_type", "customer", "customer_category",
		"source", "item_code", "item_name", "item_entity", "item_category", "collection", "design", "colour",
		"size", "qty", "amount_excl", "vat", "amount_incl"]
	batch, done, skipped = [], 0, 0
	for r in rows[1:]:
		inv = _s(r[2]) if len(r) > 2 else ""
		if not inv or r[1] in (None, ""):
			skipped += 1
			continue
		try:
			d = getdate(r[1])
		except Exception:
			skipped += 1
			continue
		code = _s(r[3]) if len(r) > 3 else ""
		qty = flt(r[4]) if len(r) > 4 else 0
		excl = flt(r[5]) if len(r) > 5 else 0
		incl = round(excl * (1 + vat_rate / 100.0), 2)
		it = items.get(code)
		cust = _s(r[0])
		# Russell, 2026-10-20: a line whose product category is Ceramics belongs to the
		# Ceramics entity; everything else in this export is Home (the company passed in).
		row_company = "Ceramics" if (it and it.item_category == "Ceramics") else company
		batch.append([
			frappe.generate_hash(length=10), user, stamp, stamp, user, 0, 0,
			row_company, "", d, inv, "Credit Note" if inv.upper().startswith("CRN") or excl < 0 else "Invoice", cust, custs.get(cust, ""),
			source, code, (it.item_name if it else ""), (it.item_entity if it else row_company), (it.item_category if it else ("Non-stock line" if not it and code else "")),
			(it.collection if it else ""), (it.design if it else ""), (it.colour if it else ""), (it.size if it else ""),
			qty, excl, round(incl - excl, 2), incl,
		])
		if len(batch) >= 2000:
			frappe.db.bulk_insert("Sales History", fields=fields, values=batch)
			frappe.db.commit()
			done += len(batch)
			batch = []
	if batch:
		frappe.db.bulk_insert("Sales History", fields=fields, values=batch)
		frappe.db.commit()
		done += len(batch)
	frappe.log_error(message=f"Loaded {done} rows into Sales History from {file_url} (company {company}, source {source}); skipped {skipped}.", title="Sales History load finished")
	return {"loaded": done, "skipped": skipped}


@frappe.whitelist()
def delete_source(source):
	"""Remove every Sales History row loaded with this source (e.g. Spreadsheet)."""
	frappe.only_for("System Manager")
	n = frappe.db.count("Sales History", {"source": source})
	frappe.db.delete("Sales History", {"source": source})
	frappe.db.commit()
	return {"deleted": n}
