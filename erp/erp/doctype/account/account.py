import frappe
from frappe import _
from frappe.model.document import Document


class Account(Document):
	def validate(self) -> None:
		if not self.is_new():
			before = self.get_doc_before_save()
			if before and before.lock:
				frappe.throw(_("Account is locked and cannot be modified."))
		self._sync_state()

	def _sync_state(self) -> None:
		if self.state_name and not self.state_code:
			self.state_code = frappe.db.get_value("State", self.state_name, "state_code")
		elif self.state_code and not self.state_name:
			result = frappe.db.get_list(
				"State", filters={"state_code": self.state_code}, fields=["name"], limit=1
			)
			if result:
				self.state_name = result[0].name
