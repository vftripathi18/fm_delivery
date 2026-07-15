frappe.ui.form.on("FM Delivery Import", {
    refresh(frm) {

        if (frm.is_new()) return;

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

                    frappe.msgprint({
                        title: __("Import Successful"),
                        indicator: "green",
                        message: __("Successfully imported <b>{0}</b> records.", [
                            r.message.imported
                        ])
                    });

                    frm.reload_doc();
                }

            });

        });

    }
});