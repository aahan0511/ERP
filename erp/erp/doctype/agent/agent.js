frappe.ui.form.on("Agent", {
	state_name(frm) {
		if (!frm.doc.state_name) return;
		// callback form avoids promise-resolution ambiguity across Frappe versions
		frappe.db.get_value("State", frm.doc.state_name, "state_code", (r) => {
			if (r && r.state_code != null) frm.set_value("state_code", r.state_code);
		});
	},
	state_code(frm) {
		if (frm.doc.state_code && !frm.doc.state_name) {
			frappe.db
				.get_list("State", {
					filters: { state_code: frm.doc.state_code },
					fields: ["name"],
					limit: 1,
				})
				.then((r) => {
					if (r.length) frm.set_value("state_name", r[0].name);
				});
		}
	},
});
