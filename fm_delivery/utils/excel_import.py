import os
from datetime import datetime

import frappe
from frappe.utils import get_site_path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill


class ExcelImporter:

    EXPECTED_HEADERS = [
        "Date",
        "State",
        "Hub Name",
        "Rider Name",
        "FHRID",
        "Return Shipment",
        "Forward Shipment",
        "Total Shipment",
    ]

    def __init__(self, import_doc):
        self.import_doc = import_doc
        self.workbook = None
        self.sheet = None

    # ------------------------------------------------------------
    # Validate File
    # ------------------------------------------------------------

    def validate_file(self):
        if not self.import_doc.import_file:
            frappe.throw("Please attach an Excel file.")

        if not self.import_doc.import_file.lower().endswith(".xlsx"):
            frappe.throw("Only .xlsx files are allowed.")

    # ------------------------------------------------------------
    # Load Excel
    # ------------------------------------------------------------

    def load_workbook(self):
        self.validate_file()

        file_doc = frappe.get_doc(
            "File",
            {"file_url": self.import_doc.import_file}
        )

        file_path = get_site_path(
            file_doc.file_url.lstrip("/")
        )

        if not os.path.exists(file_path):
            frappe.throw(
                f"File not found: {file_path}"
            )

        self.workbook = load_workbook(
            filename=file_path,
            data_only=True
        )

        self.sheet = self.workbook.active

        return self.workbook, self.sheet

    # ------------------------------------------------------------
    # Validate Headers
    # ------------------------------------------------------------

    def validate_headers(self):
        headers = [
            str(cell.value).strip()
            if cell.value
            else ""
            for cell in self.sheet[1]
        ]

        if headers != self.EXPECTED_HEADERS:
            frappe.throw(
                f"""
                <h4>Excel Header Validation Failed</h4>

                <b>Expected:</b>
                <br>
                {'<br>'.join(self.EXPECTED_HEADERS)}

                <hr>

                <b>Found:</b>
                <br>
                {'<br>'.join(headers)}
                """
            )

    # ------------------------------------------------------------
    # Integer Converter
    # ------------------------------------------------------------

    def to_int(self, value):
        if value in (None, "", "NULL"):
            return 0

        try:
            return int(float(value))
        except Exception:
            return 0

    # ------------------------------------------------------------
    # Date Converter
    # ------------------------------------------------------------

    def to_date(self, value):
        if not value:
            return None

        if isinstance(value, datetime):
            return value.date()

        return value

    # ------------------------------------------------------------
    # Import Data
    # ------------------------------------------------------------

    def import_data(self):
        imported = 0
        skipped = 0
        failed = 0
        failed_rows = []

        for row_no, row in enumerate(
            self.sheet.iter_rows(
                min_row=2,
                values_only=True
            ),
            start=2
        ):
            if not any(row):
                continue

            try:
                date = self.to_date(row[0])
                fhr_id = str(
                    row[4] or ""
                ).strip()

                # Duplicate Check
                if frappe.db.exists(
                    "FM Delivery Data",
                    {
                        "date": date,
                        "fhr_id": fhr_id
                    }
                ):
                    skipped += 1
                    continue

                doc = frappe.new_doc(
                    "FM Delivery Data"
                )

                doc.date = date
                doc.state = (row[1] or "").strip()
                doc.hub_name = (row[2] or "").strip()
                doc.rider_name = (row[3] or "").strip()
                doc.fhr_id = fhr_id

                doc.return_shipment = self.to_int(row[5])
                doc.forward_shipment = self.to_int(row[6])
                doc.total_shipment = self.to_int(row[7])

                doc.insert(
                    ignore_permissions=True
                )
                imported += 1

            except Exception as e:
                failed += 1
                error_msg = frappe.get_traceback() or str(e)
                last_line_error = error_msg.strip().split("\n")[-1]
                
                failed_rows.append({
                    "row_no": row_no,
                    "row_data": row,
                    "error": last_line_error
                })

                frappe.log_error(
                    title=f"FM Delivery Import Row {row_no}",
                    message=error_msg
                )

        frappe.db.commit()

        error_file_url = None
        if failed_rows:
            error_file_url = self.create_error_report(failed_rows)

        # Update Import Document
        self.import_doc.imported_rows = imported
        self.import_doc.skipped_rows = skipped
        self.import_doc.failed_rows = failed
        self.import_doc.status = "Completed"

        if error_file_url:
            self.import_doc.error_file = error_file_url

        self.import_doc.save(
            ignore_permissions=True
        )

        return {
            "imported": imported,
            "skipped": skipped,
            "failed": failed,
            "error_file": error_file_url
        }

    # ------------------------------------------------------------
    # Create Error Report
    # ------------------------------------------------------------

    def create_error_report(self, failed_rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Errors"

        headers = ["Row Number"] + self.EXPECTED_HEADERS + ["Error Message"]
        ws.append(headers)

        # Styles
        bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = bold_font
            cell.fill = red_fill

        for item in failed_rows:
            row_no = item["row_no"]
            row_data = item["row_data"]
            error_msg = item["error"]

            row_to_write = [row_no]
            for val in row_data:
                if isinstance(val, datetime):
                    row_to_write.append(val.strftime("%Y-%m-%d"))
                else:
                    row_to_write.append(val)
            row_to_write.append(error_msg)
            ws.append(row_to_write)

        # Autosize Columns
        for col in ws.columns:
            max_len = 0
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"FM_Import_Errors_{timestamp}.xlsx"
        
        folder_path = get_site_path("private", "files")
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        file_path = os.path.join(folder_path, filename)
        wb.save(file_path)

        # Create DocType File
        file_doc = frappe.new_doc("File")
        file_doc.file_name = filename
        file_doc.is_private = 1
        file_doc.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        file_doc.file_url = f"/private/files/{filename}"
        file_doc.attached_to_doctype = self.import_doc.doctype
        file_doc.attached_to_name = self.import_doc.name
        file_doc.insert(ignore_permissions=True)

        return file_doc.file_url