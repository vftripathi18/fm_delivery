import frappe

from fm_delivery.utils.excel_import import ExcelImporter


@frappe.whitelist()
def import_excel(docname):

    import_doc = frappe.get_doc(
        "FM Delivery Import",
        docname
    )

    importer = ExcelImporter(import_doc)

    importer.load_workbook()

    importer.validate_headers()

    result = importer.import_data()

    import_doc.status = "Completed"
    import_doc.save(ignore_permissions=True)

    return result