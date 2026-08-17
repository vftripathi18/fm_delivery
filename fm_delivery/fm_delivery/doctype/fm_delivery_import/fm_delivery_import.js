frappe.ui.form.on("FM Delivery Import", {
    refresh(frm) {

        if (frm.is_new()) {
            return;
        }

        frm.clear_custom_buttons();

        // ============================================================
        // DOWNLOAD TEMPLATE
        // ============================================================

        frm.add_custom_button(
            __("Download Template"),
            function () {

                window.open(
                    "/private/files/FM%2016.xlsx",
                    "_blank"
                );

            },
            __("Actions")
        );


        // ============================================================
        // IMPORT DATA
        // ============================================================

        frm.add_custom_button(
            __("Import Data"),
            function () {

                frappe.call({

                    method: "fm_delivery.api.import_excel",

                    args: {
                        docname: frm.doc.name
                    },

                    freeze: true,

                    freeze_message: __(
                        "Importing Excel..."
                    ),

                    callback: function (r) {

                        if (r.exc) {
                            return;
                        }

                        if (!r.message) {
                            return;
                        }

                        // ====================================================
                        // IMPORT SUMMARY
                        // ====================================================

                        let message_html = `
                            <div style="
                                font-size:14px;
                                line-height:1.9;
                            ">

                                📄
                                <b>Attempted :</b>
                                ${r.message.attempted || 0}

                                <br>

                                ✅
                                <b>Imported :</b>
                                ${r.message.imported || 0}

                                <br>

                                🔄
                                <b>Updated :</b>
                                ${r.message.updated || 0}

                                <br>

                                ❌
                                <b>Failed :</b>
                                ${r.message.failed || 0}

                                <br>

                                ➖
                                <b>Difference :</b>
                                ${r.message.difference || 0}

                            </div>
                        `;


                        // ====================================================
                        // ERROR REPORT AVAILABLE
                        // ====================================================

                        if (r.message.error_file) {

                            frappe.msgprint({

                                title: __(
                                    "Import Summary"
                                ),

                                indicator: "orange",

                                message: message_html,

                                primary_action: {

                                    action: function () {

                                        window.open(
                                            r.message.error_file,
                                            "_blank"
                                        );

                                    },

                                    label: __(
                                        "Download Error Report"
                                    )

                                }

                            });

                        }

                        // ====================================================
                        // NO ERRORS
                        // ====================================================

                        else {

                            frappe.msgprint({

                                title: __(
                                    "Import Summary"
                                ),

                                indicator: "green",

                                message: message_html

                            });

                        }


                        // ====================================================
                        // RELOAD DOCUMENT
                        // ====================================================

                        frm.reload_doc();

                    }

                });

            },
            __("Actions")
        );


        // ============================================================
        // DOWNLOAD EXISTING ERROR REPORT
        // ============================================================

        if (frm.doc.error_file) {

            frm.add_custom_button(
                __("Download Error Report"),
                function () {

                    window.open(
                        frm.doc.error_file,
                        "_blank"
                    );

                },
                __("Actions")
            );

        }

    }
});