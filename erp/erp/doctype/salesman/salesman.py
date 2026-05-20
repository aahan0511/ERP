import frappe
from frappe import _
from frappe.model.document import Document


class Salesman(Document):
	def validate(self) -> None:
		if not self.is_new():
			before = self.get_doc_before_save()
			if before and before.lock:
				frappe.throw(_("Salesman is locked and cannot be modified."))
