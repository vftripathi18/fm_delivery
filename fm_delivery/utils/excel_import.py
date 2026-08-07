import os
from datetime import datetime

import frappe
from frappe.utils import get_site_path
from openpyxl import Workbook, load_workbook
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
            data_only=True,
            read_only=True
        )

        self.sheet = self.workbook.active

        return self.workbook, self.sheet


    # ------------------------------------------------------------
    # Validate Headers
    # ------------------------------------------------------------

    def validate_headers(self):

        headers = [
            str(cell.value).strip() if cell.value else ""
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

    def to_float(self, value):

        if value in (None, "", "NULL"):
            return 0.0

        try:
            return float(value)

        except Exception:
            return 0.0


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

        attempted = 0
        imported = 0
        failed = 0

        failed_rows = []


        headers = [
            str(cell.value).strip() if cell.value else ""
            for cell in self.sheet[1]
        ]


        for row_no, values in enumerate(
            self.sheet.iter_rows(
                min_row=2,
                values_only=True
            ),
            start=2
        ):

            # Skip completely empty excel rows
            if not any(values):
                continue

            attempted += 1

            try:

                row = {
                    headers[i]: values[i]
                    if i < len(values)
                    else None

                    for i in range(len(headers))
                }


                doc = frappe.new_doc(
                    "FM Delivery Data"
                )


                doc.update(
                    {
                        "date": self.to_date(
                            row.get("Date")
                        ),

                        "state": str(
                            row.get("State") or ""
                        ).strip(),

                        "hub_name": str(
                            row.get("Hub Name") or ""
                        ).strip(),

                        "rider_name": str(
                            row.get("Rider Name") or ""
                        ).strip(),

                        "fhr_id": str(
                            row.get("FHRID") or ""
                        ).strip(),

                        "return_shipment": self.to_float(
                            row.get("Return Shipment")
                        ),

                        "forward_shipment": self.to_float(
                            row.get("Forward Shipment")
                        ),

                        "total_shipment": self.to_float(
                            row.get("Total Shipment")
                        ),
                    }
                )


                # Direct insert
                # No duplicate validation
                doc.insert(
                    ignore_permissions=True
                )


                imported += 1



            except Exception:


                failed += 1


                error_msg = frappe.get_traceback()


                failed_rows.append(
                    {
                        "row_no": row_no,
                        "row_data": values,
                        "error": error_msg.strip().split("\n")[-1],
                    }
                )


                frappe.log_error(
                    title=f"FM Delivery Import Row {row_no}",
                    message=error_msg
                )


        frappe.db.commit()



        error_file_url = None


        if failed_rows:

            error_file_url = self.create_error_report(
                failed_rows
            )



        self.import_doc.imported_rows = imported
        self.import_doc.failed_rows = failed


        self.import_doc.status = "Completed"



        if error_file_url:

            self.import_doc.error_file = error_file_url



        self.import_doc.save(
            ignore_permissions=True
        )



        return {

            "attempted": attempted,

            "imported": imported,

            "failed": failed,

            "difference": attempted - (imported + failed),

            "error_file": error_file_url,

        }



    # ------------------------------------------------------------
    # Create Error Report
    # ------------------------------------------------------------

    def create_error_report(self, failed_rows):


        wb = Workbook()

        ws = wb.active

        ws.title = "Errors"



        headers = [
            "Row Number"
        ] + self.EXPECTED_HEADERS + [
            "Error Message"
        ]


        ws.append(headers)



        bold_font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF"
        )


        red_fill = PatternFill(
            start_color="FF0000",
            end_color="FF0000",
            fill_type="solid"
        )



        for col_idx in range(
            1,
            len(headers) + 1
        ):

            cell = ws.cell(
                row=1,
                column=col_idx
            )

            cell.font = bold_font

            cell.fill = red_fill



        for item in failed_rows:


            row_to_write = [
                item["row_no"]
            ]


            for val in item["row_data"]:

                if isinstance(val, datetime):

                    row_to_write.append(
                        val.strftime("%Y-%m-%d")
                    )

                else:

                    row_to_write.append(val)



            row_to_write.append(
                item["error"]
            )


            ws.append(
                row_to_write
            )



        for col in ws.columns:


            max_len = 0


            for cell in col:

                value = str(
                    cell.value or ""
                )

                max_len = max(
                    max_len,
                    len(value)
                )


            ws.column_dimensions[
                col[0].column_letter
            ].width = max(
                max_len + 3,
                10
            )



        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )


        filename = (
            f"FM_Import_Errors_{timestamp}.xlsx"
        )



        folder_path = get_site_path(
            "private",
            "files"
        )


        os.makedirs(
            folder_path,
            exist_ok=True
        )


        file_path = os.path.join(
            folder_path,
            filename
        )


        wb.save(
            file_path
        )



        file_doc = frappe.new_doc(
            "File"
        )


        file_doc.file_name = filename

        file_doc.is_private = 1

        file_doc.content_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        file_doc.file_url = (
            f"/private/files/{filename}"
        )

        file_doc.attached_to_doctype = (
            self.import_doc.doctype
        )

        file_doc.attached_to_name = (
            self.import_doc.name
        )


        file_doc.insert(
            ignore_permissions=True
        )


        return file_doc.file_url