frappe.ui.form.on("FM Delivery Import", {
    refresh(frm) {

        if (frm.is_new()) {
            return;
        }

        // Remove duplicate buttons on refresh
        frm.clear_custom_buttons();

        // Download Template Button
        frm.add_custom_button(__("Download Template"), function () {

            window.open(
                "/private/files/FM%2016.xlsx",
                "_blank"
            );

        }, __("Actions"));

        // Import Data Button
        frm.add_custom_button(__("Import Data"), function () {

            frappe.call({

                method: "fm_delivery.api.import_excel",

                args: {
                    docname: frm.doc.name
                },

                freeze: true,
                freeze_message: __("Importing Excel..."),

                callback: function (r) {

                    if (r.exc) {
                        return;
                    }

                    let message_html = `
                        <div style="font-size:14px;line-height:1.8;">
                            ✅ <b>Imported :</b> ${r.message.imported}<br>
                            ⚠️ <b>Skipped :</b> ${r.message.skipped}<br>
                            ❌ <b>Failed :</b> ${r.message.failed}
                        </div>
                    `;

                    if (r.message.error_file) {
                        frappe.msgprint({
                            title: __("Import Summary"),
                            indicator: "red",
                            message: message_html,
                            primary_action: {
                                action: function () {
                                    window.open(r.message.error_file, "_blank");
                                },
                                label: __("Download Error Report")
                            }
                        });
                    } else {
                        frappe.msgprint({
                            title: __("Import Summary"),
                            indicator: "green",
                            message: message_html
                        });
                    }

                    frm.reload_doc();

                }

            });

        }, __("Actions"));

        // Download Error Report Button (if error file exists on the document)
        if (frm.doc.error_file) {
            frm.add_custom_button(__("Download Error Report"), function () {
                window.open(frm.doc.error_file, "_blank");
            }, __("Actions"));
        }

    }
});