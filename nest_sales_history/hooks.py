app_name        = "nest_sales_history"
app_title       = "Nest Sales History"
app_publisher   = "NestERP / Syncflo"
app_description = "One flat sales-history table for reporting, kept live from ERPNext invoices and loaded once from the previous system."
app_email       = "ops@syncflo.co.za"
app_license     = "MIT"
app_icon        = "octicon octicon-graph"

required_apps = ["erpnext"]

# Every submitted invoice or credit note writes its lines into Sales History; a
# cancel takes them out again. Reports read Sales History only.
doc_events = {
	"Sales Invoice": {
		"on_submit": "nest_sales_history.api.on_invoice_submit",
		"on_cancel": "nest_sales_history.api.on_invoice_cancel",
	},
}
