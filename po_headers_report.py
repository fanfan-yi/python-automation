from pathlib import Path
from openpyxl import Workbook

purchase_orders = [
    {
        "po_number": "41450123",
        "supplier": "Formosa Industrial Supplies",
        "status": "APPROVED",
        "amount": 12500.00,
        "currency_code": "NTD"
    },
    {
        "po_number": "41450131",
        "supplier": "Pacific Components Ltd",
        "status": "IN PROCESS",
        "amount": 8750.50,
        "currency_code": "NTD"
    },
    {
        "po_number": "41450155",
        "supplier": "Southern Cross Logistics",
        "status": "APPROVED",
        "amount": 4200.00,
        "currency_code": "USD"
    },
    {
        "po_number": "41450116",
        "supplier": "Asia Tech Manufacturing",
        "status": "IN PROCESS",
        "amount": 15600.75,
        "currency_code": "NTD"
    },
    {
        "po_number": "41450130",
        "supplier": "Evergreen Office Solutions",
        "status": "APPROVED",
        "amount": 12100.00,
        "currency_code": "USD"
    },
]


summary = {
    "total_po": 0,
    "approved_count": 0,
    "in_process_count": 0,
    "amount": {
        "USD": 0.0,
        "NTD": 0.0
    }
}

def print_po_detail(po):
    print("=" * 40)
    print(f"PO Number : {po['po_number']}")
    print(f"Supplier  : {po['supplier']}")
    print(f"Status    : {po['status']}")
    print(f"Amount    : {po['currency_code']}${po['amount']:,.2f}")
    print("=" * 40)


def print_summary(summary):
    print(f"Total PO : {summary['total_po']}")
    print(f"Approved PO : {summary['approved_count']}")
    print(f"In Process PO : {summary['in_process_count']}")
    print(f"USD Total Amount : USD${summary['amount']['USD']:,.2f}")
    print(f"NTD Total Amount : NTD${summary['amount']['NTD']:,.2f}")

def calculate_summary(po, summary):
    """Update the summary using one purchase order."""
    summary["total_po"] += 1

    if po['status'] == "APPROVED":
        summary['approved_count'] += 1
    elif po['status'] == "IN PROCESS":
        summary['in_process_count'] += 1

    if po['currency_code']  == "USD":
        summary["amount"]["USD"] += po['amount']
    elif po['currency_code'] == "NTD":
        summary["amount"]["NTD"] += po['amount']
    #return summary  函式直接修改 dictionary 內容,不需要return

def export_to_excel(purchase_orders, summary, output_path):
    """Export purchase order details and summary to an Excel workbook."""

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "ERP Procurement Report"

    headers = [
        "PO Number",
        "Supplier",
        "Status",
        "Amount",
        "Currency"
    ]

    worksheet.append(headers)

    for po in purchase_orders:
        worksheet.append([
            po["po_number"],
            po["supplier"],
            po["status"],
            po["amount"],
            po["currency_code"]
        ])

    worksheet.append([])

    worksheet.append(["Summary"])
    worksheet.append(["Total PO", summary["total_po"]])
    worksheet.append(["Approved Count", summary["approved_count"]])
    worksheet.append(["In Process Count", summary["in_process_count"]])

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    workbook.save(output_file)

    print(f"Excel report created: {output_file}")



if __name__ == "__main__":
    for po in purchase_orders:
        print_po_detail(po)
        calculate_summary(po, summary)

    print_summary(summary)

    export_to_excel(
        purchase_orders,
        summary,
        "output/erp_procurement_report.xlsx"
    )
