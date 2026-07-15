import os
from datetime import datetime

import frappe
from frappe.utils import get_site_path
from openpyxl import load_workbook


class ExcelImporter:

    EXPECTED_HEADERS = [
        "Date",
        "VendorName",
        "Agent_Name",
        "Agent_id",
        "FHRID",
        "Hub_Id",
        "Hub_Name",
        "tripsheetId",
        "sellersAssigned",
        "sellersAttempted",
        "deliveredReverseShipment",
        "undeliveredReverseShipment",
        "assignedForwardShipment",
        "pickedForwardShipment",
        "TripsheetClosureDate",
    ]

    def __init__(self, import_doc):
        self.import_doc = import_doc
        self.workbook = None
        self.sheet = None

    # ------------------------------------------------------------------
    # Validate File
    # ------------------------------------------------------------------

    def validate_file(self):

        if not self.import_doc.import_file:
            frappe.throw("Please attach an Excel file.")

        if not self.import_doc.import_file.lower().endswith(".xlsx"):
            frappe.throw("Only .xlsx files are allowed.")

    # ------------------------------------------------------------------
    # Load Workbook
    # ------------------------------------------------------------------

    def load_workbook(self):

        self.validate_file()

        file_doc = frappe.get_doc(
            "File",
            {"file_url": self.import_doc.import_file}
        )

        file_path = get_site_path(file_doc.file_url.lstrip("/"))

        if not os.path.exists(file_path):
            frappe.throw(f"File not found : {file_path}")

        self.workbook = load_workbook(
            filename=file_path,
            data_only=True
        )

        self.sheet = self.workbook.active

        return self.workbook, self.sheet

    # ------------------------------------------------------------------
    # Validate Headers
    # ------------------------------------------------------------------

    def validate_headers(self):

        headers = [
            cell.value.strip() if isinstance(cell.value, str) else cell.value
            for cell in self.sheet[1]
        ]

        if headers != self.EXPECTED_HEADERS:

            frappe.throw(f"""
                <h4>Excel Header Validation Failed</h4>

                <hr>

                <b>Expected</b>

                <br>{'<br>'.join(self.EXPECTED_HEADERS)}

                <hr>

                <b>Found</b>

                <br>{'<br>'.join([str(i) for i in headers])}
            """)

    # ------------------------------------------------------------------
    # Integer Converter
    # ------------------------------------------------------------------

    def to_int(self, value):

        if value in (None, "", "NULL"):
            return 0

        try:
            return int(float(value))
        except Exception:
            return 0

    # ------------------------------------------------------------------
    # Date Converter
    # ------------------------------------------------------------------

    def to_date(self, value):

        if value in (None, "", "NULL"):
            return None

        if isinstance(value, datetime):
            return value.date()

        return value

    # ------------------------------------------------------------------
    # Datetime Converter
    # ------------------------------------------------------------------

    def to_datetime(self, value):

        if value in (None, "", "NULL"):
            return None

        if isinstance(value, datetime):
            return value

        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%d-%m-%Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %H:%M",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d/%m/%Y",
        ]

        for fmt in formats:
            try:
                return datetime.strptime(str(value), fmt)
            except Exception:
                pass

        return None

    # ------------------------------------------------------------------
    # Import Data
    # ------------------------------------------------------------------

    def import_data(self):

        imported = 0

        for row_no, row in enumerate(
            self.sheet.iter_rows(min_row=2, values_only=True),
            start=2,
        ):

            if not any(row):
                continue

            try:

                doc = frappe.new_doc("FM Delivery Data")

                doc.date = self.to_date(row[0])

                doc.vendor_name = row[1] or ""

                doc.agent_name = row[2] or ""

                doc.agent_id = row[3] or ""

                doc.fhr_id = row[4] or ""

                doc.hub_id = row[5] or ""

                doc.hub_name = row[6] or ""

                doc.tripsheet_id = row[7] or ""

                doc.sellers_assigned = self.to_int(row[8])

                doc.sellers_attempted = self.to_int(row[9])

                doc.delivered_reverse_shipment = self.to_int(row[10])

                doc.undelivered_reverse_shipment = self.to_int(row[11])

                doc.assigned_forward_shipment = self.to_int(row[12])

                doc.picked_forward_shipment = self.to_int(row[13])

                doc.tripsheet_closure_date = self.to_datetime(row[14])

                doc.insert(ignore_permissions=True)

                imported += 1

            except Exception:

                frappe.log_error(
                    frappe.get_traceback(),
                    f"FM Delivery Import Row {row_no}"
                )

        frappe.db.commit()

        return imported

    # ------------------------------------------------------------------
    # Workbook Information
    # ------------------------------------------------------------------

    def workbook_information(self):

        return {
            "sheet_name": self.sheet.title,
            "rows": self.sheet.max_row,
            "columns": self.sheet.max_column,
        }