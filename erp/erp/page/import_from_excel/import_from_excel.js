frappe.pages["import-from-excel"].on_page_load = function (wrapper) {
	const page = frappe.ui.make_app_page({
		parent: wrapper,
		title: "Import from Excel",
		single_column: true,
	});
	new ImportFromExcel(page);
};

class ImportFromExcel {
	constructor(page) {
		this.page = page;
		this.$main = $(page.main);
		this.state = {
			step: 1,
			file_url: null,
			file_name: null,
			sheet_names: [],
			selected_sheet: null,
			excel_columns: [],
			total_rows: 0,
			doctype: null,
			doctype_fields: [],
			mappings: [],
			selected_x: null,
			selected_f: null,
			sort_mode: "original",
			show_sample: true,
			required_only: false,
			density: "comfortable",
			preview_data: null,
			import_result: null,
			suggested_doctype: null,
		};
		this._inject_styles();
		this.render();
	}

	setState(updates) {
		Object.assign(this.state, updates);
		this.render();
	}

	// ── Styles ──────────────────────────────────────────────────────────────
	_inject_styles() {
		if (document.getElementById("ei-styles")) return;
		const s = document.createElement("style");
		s.id = "ei-styles";
		s.textContent = `
			.ei-wrap { max-width: 1400px; padding: 16px 0 60px; }

			.ei-step-nav {
				display: flex; align-items: center; gap: 6px; flex-wrap: wrap;
				padding: 10px 14px; background: var(--fg-color, #fff);
				border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px); margin-bottom: 18px;
			}
			.ei-sp {
				display: inline-flex; align-items: center; gap: 5px;
				padding: 3px 10px; border-radius: 16px; font-size: 12px;
				color: var(--text-muted, #6b7280); border: 1px solid transparent;
			}
			.ei-sp.active { background: var(--primary-color, #2490ef); color: #fff; }
			.ei-sp.done   { color: var(--green, #2f9a5c); border-color: var(--green, #2f9a5c); }
			.ei-sp-n {
				width: 18px; height: 18px; border-radius: 50%; font-size: 10px; font-weight: 700;
				display: inline-flex; align-items: center; justify-content: center;
				border: 1.5px solid currentColor;
			}
			.ei-sp.active .ei-sp-n { background: rgba(255,255,255,.25); border-color: #fff; }
			.ei-nav-sep { color: var(--text-muted, #9ca3af); font-size: 14px; }

			.ei-card {
				background: var(--fg-color, #fff);
				border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px);
				padding: 22px 24px; margin-bottom: 18px;
			}
			.ei-card-title { font-size: 15px; font-weight: 600; margin-bottom: 14px; }

			/* Upload */
			.ei-drop {
				border: 2px dashed var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px);
				padding: 52px 24px; text-align: center; cursor: pointer;
				transition: border-color .15s, background .15s;
				background: var(--bg-color, #f9fafb); color: var(--text-muted, #6b7280);
			}
			.ei-drop:hover, .ei-drop.dragging {
				border-color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
				color: var(--primary-color, #2490ef);
			}
			.ei-drop-icon { font-size: 28px; margin-bottom: 8px; }
			.ei-drop-text { font-size: 14px; font-weight: 500; margin-bottom: 4px; }
			.ei-drop-sub  { font-size: 12px; }

			/* Sheet picker */
			.ei-sheet-grid { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
			.ei-sheet-btn {
				padding: 8px 16px; border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px); background: var(--fg-color, #fff);
				cursor: pointer; font-size: 13px; display: flex; align-items: center;
				gap: 6px; transition: all .12s;
			}
			.ei-sheet-btn:hover {
				border-color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
			}

			/* Toolbar */
			.ei-toolbar {
				display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
				padding: 8px 12px; background: var(--bg-color, #f9fafb);
				border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px); margin-bottom: 12px;
			}
			.ei-counter {
				display: inline-flex; align-items: center; gap: 5px; font-size: 12px;
				padding: 3px 9px; border: 1px solid var(--border-color, #d1d5db);
				border-radius: 20px; background: var(--fg-color, #fff);
			}
			.ei-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--primary-color, #2490ef); flex-shrink: 0; }
			.ei-spacer { flex: 1; }

			/* DocType grid */
			.ei-dt-grid {
				display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
				gap: 8px; margin-top: 14px;
			}
			.ei-dt-btn {
				padding: 10px 14px; border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px); background: var(--fg-color, #fff);
				cursor: pointer; font-size: 13px; text-align: left; transition: all .12s;
				display: flex; align-items: center; gap: 6px;
			}
			.ei-dt-btn:hover {
				border-color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
			}
			.ei-dt-btn.suggested {
				border-color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
				font-weight: 600;
			}

			/* Two-column mapping */
			.ei-two-col {
				display: grid; grid-template-columns: 1fr 60px 1fr;
				position: relative; min-height: 120px;
			}
			.ei-col-cards { display: flex; flex-direction: column; gap: 7px; }
			.ei-col-head {
				font-size: 12px; font-weight: 600; padding: 3px 10px;
				border: 1px solid var(--primary-color, #2490ef); border-radius: 14px;
				color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
				width: fit-content; margin-bottom: 10px;
			}
			.ei-col-head.right { margin-left: auto; }

			/* Field card */
			.ei-fcard {
				border: 1px solid var(--border-color, #d1d5db);
				border-radius: var(--border-radius, 6px);
				padding: 7px 12px; background: var(--fg-color, #fff);
				position: relative; display: flex; align-items: center; gap: 7px;
				min-height: 40px; cursor: pointer;
				transition: box-shadow .12s, border-color .12s, transform .1s;
				user-select: none;
			}
			.ei-fcard.compact { min-height: 28px; padding: 3px 10px; }
			.ei-fcard:hover { box-shadow: 0 2px 6px rgba(0,0,0,.08); transform: translateY(-1px); }
			.ei-fcard.selected {
				border-color: var(--primary-color, #2490ef);
				background: var(--primary-light-color, #ebf4fd);
				box-shadow: 0 0 0 2px rgba(36,144,239,.15);
			}
			.ei-fcard.mapped { background: #f0faf4; border-color: var(--green, #2f9a5c); }
			.ei-fname { font-size: 13px; font-weight: 500; min-width: 0; }
			.ei-fsample {
				font-family: var(--monospace-font, monospace); font-size: 10px;
				color: var(--text-muted, #6b7280); margin-left: auto;
				background: var(--bg-color, #f9fafb); padding: 1px 5px; border-radius: 3px;
				max-width: 140px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
				flex-shrink: 0;
			}
			.ei-ftype {
				font-size: 9px; color: var(--text-muted, #6b7280);
				border: 1px solid var(--border-color, #d1d5db); border-radius: 3px;
				padding: 0 4px; text-transform: uppercase; white-space: nowrap; flex-shrink: 0;
			}
			.ei-freq { color: var(--red, #e53e3e); font-weight: 700; font-size: 14px; flex-shrink: 0; }

			.ei-dot-l, .ei-dot-r {
				position: absolute; top: 50%; transform: translateY(-50%);
				width: 11px; height: 11px; border-radius: 50%;
				border: 2px solid var(--border-color, #d1d5db); background: var(--fg-color, #fff); z-index: 2;
			}
			.ei-dot-l { left: -16px; }
			.ei-dot-r { right: -16px; }
			.ei-fcard.selected .ei-dot-l, .ei-fcard.selected .ei-dot-r {
				background: var(--primary-color, #2490ef); border-color: var(--primary-color, #2490ef);
				box-shadow: 0 0 0 3px rgba(36,144,239,.2);
			}
			.ei-fcard.mapped .ei-dot-l, .ei-fcard.mapped .ei-dot-r {
				background: var(--green, #2f9a5c); border-color: var(--green, #2f9a5c);
			}

			.ei-rm-pill {
				font-size: 10px; color: var(--text-muted, #6b7280); cursor: pointer;
				padding: 1px 5px; border-radius: 8px; background: var(--bg-color, #f9fafb);
				border: 1px solid var(--border-color, #d1d5db); white-space: nowrap; flex-shrink: 0;
			}
			.ei-rm-pill:hover { border-color: var(--red, #e53e3e); color: var(--red, #e53e3e); }

			/* Center spacer column */
			.ei-center-col { display: flex; flex-direction: column; align-items: center; }

			/* SVG connector layer */
			.ei-svg-layer {
				position: absolute; top: 0; left: 0; pointer-events: none; overflow: visible;
			}

			/* Preview table */
			.ei-tbl { border-collapse: collapse; font-size: 12px; width: 100%; }
			.ei-tbl th, .ei-tbl td {
				border: 1px solid var(--border-color, #d1d5db); padding: 6px 10px; text-align: left;
			}
			.ei-tbl thead th { background: var(--bg-color, #f9fafb); font-weight: 600; font-size: 11px; }

			/* Success */
			.ei-success {
				border: 2px solid var(--green, #2f9a5c); border-radius: var(--border-radius, 6px);
				padding: 36px 24px; background: #f4fbf7; text-align: center; margin-bottom: 16px;
			}
			.ei-s-icon  { font-size: 44px; color: var(--green, #2f9a5c); line-height: 1; }
			.ei-s-title { font-size: 22px; font-weight: 700; color: var(--green, #2f9a5c); margin: 10px 0 4px; }
			.ei-stats   { display: flex; justify-content: center; gap: 28px; margin-top: 18px; }
			.ei-stat-n  { font-size: 26px; font-weight: 700; display: block; }
			.ei-stat-l  { font-size: 11px; color: var(--text-muted, #6b7280); }

			/* Link error section */
			.ei-link-err-card {
				border: 1px solid var(--orange, #f59e0b); border-radius: var(--border-radius, 6px);
				padding: 18px 20px; margin-bottom: 14px; background: #fffbeb;
			}
			.ei-link-err-title { font-size: 14px; font-weight: 600; color: var(--orange, #d97706); margin-bottom: 14px; }
			.ei-link-err-group { margin-bottom: 14px; padding-bottom: 14px; border-bottom: 1px solid var(--border-color, #d1d5db); }
			.ei-link-err-group:last-child { border-bottom: 0; margin-bottom: 0; padding-bottom: 0; }
			.ei-missing-vals {
				font-family: var(--monospace-font, monospace); font-size: 11px;
				color: var(--text-muted, #6b7280); margin: 6px 0 10px;
			}

			.ei-nav { display: flex; gap: 10px; margin-top: 14px; }
		`;
		document.head.appendChild(s);
	}

	// ── Render dispatcher ───────────────────────────────────────────────────
	render() {
		this.$main.empty();
		const $w = $('<div class="ei-wrap"></div>').appendTo(this.$main);
		this._render_step_nav($w);
		[
			this._render_step1,
			this._render_step2,
			this._render_step3,
			this._render_step4,
			this._render_step5,
		][this.state.step - 1].call(this, $w);
	}

	_render_step_nav($w) {
		const labels = ["Upload", "DocType", "Map Columns", "Validate", "Import"];
		const $nav = $('<div class="ei-step-nav"></div>').appendTo($w);
		labels.forEach((lbl, i) => {
			if (i > 0) $('<span class="ei-nav-sep">→</span>').appendTo($nav);
			const n = i + 1;
			const cls =
				n === this.state.step ? "active" : n < this.state.step ? "done" : "";
			$(`<span class="ei-sp ${cls}"><span class="ei-sp-n">${n < this.state.step ? "✓" : n}</span>${lbl}</span>`).appendTo(
				$nav
			);
		});
	}

	// ── Step 1: Upload + Sheet Selection ───────────────────────────────────
	_render_step1($w) {
		const s = this.state;
		const $card = $('<div class="ei-card"></div>').appendTo($w);

		if (s.file_url && s.sheet_names.length > 1) {
			// Sheet picker state
			$('<div class="ei-card-title">Select Sheet to Import</div>').appendTo($card);
			$(
				`<div style="font-size:12px;color:var(--text-muted);margin-bottom:4px;">
					File: <b>${frappe.utils.escape_html(s.file_name)}</b> · ${s.sheet_names.length} sheets found
				</div>`
			).appendTo($card);

			const $grid = $('<div class="ei-sheet-grid"></div>').appendTo($card);
			s.sheet_names.forEach((sheet) => {
				$(`<button class="ei-sheet-btn">📋 ${frappe.utils.escape_html(sheet)}</button>`)
					.appendTo($grid)
					.on("click", () => this._parse_sheet(s.file_url, sheet));
			});

			$('<div class="ei-nav"></div>')
				.appendTo($card)
				.append(
					$('<button class="btn btn-default btn-sm">← Change file</button>').on(
						"click",
						() =>
							this.setState({
								file_url: null,
								file_name: null,
								sheet_names: [],
							})
					)
				);
			return;
		}

		// Drop zone
		$('<div class="ei-card-title">Upload Excel File</div>').appendTo($card);

		if (s.suggested_doctype) {
			$(
				`<div style="font-size:13px;margin-bottom:12px;padding:8px 12px;background:var(--primary-light-color,#ebf4fd);border-radius:4px;color:var(--primary-color,#2490ef);">
					↳ Upload data for: <b>${frappe.utils.escape_html(s.suggested_doctype)}</b>
				</div>`
			).appendTo($card);
		}

		const $drop = $(`
			<div class="ei-drop">
				<div class="ei-drop-icon">⬆</div>
				<div class="ei-drop-text">Drop .xlsx / .xls here</div>
				<div class="ei-drop-sub">or click to browse</div>
			</div>
		`).appendTo($card);

		const $input = $(
			'<input type="file" accept=".xlsx,.xls" style="display:none">'
		).appendTo($card);

		$drop.on("click", () => $input.trigger("click"));
		$drop.on("dragover", (e) => {
			e.preventDefault();
			$drop.addClass("dragging");
		});
		$drop.on("dragleave drop", () => $drop.removeClass("dragging"));
		$drop.on("drop", (e) => {
			e.preventDefault();
			$drop.removeClass("dragging");
			const file = e.originalEvent.dataTransfer.files[0];
			if (file) this._upload_file(file, $drop);
		});
		$input.on("change", (e) => {
			const file = e.target.files[0];
			if (file) this._upload_file(file, $drop);
		});
	}

	_upload_file(file, $drop) {
		$drop.html(
			`<div class="ei-drop-icon">⏳</div><div class="ei-drop-text">Uploading ${frappe.utils.escape_html(file.name)}…</div>`
		);

		const fd = new FormData();
		fd.append("file", file, file.name);
		fd.append("is_private", "1");
		fd.append("folder", "Home");

		$.ajax({
			url: "/api/method/upload_file",
			type: "POST",
			data: fd,
			processData: false,
			contentType: false,
			headers: { "X-Frappe-CSRF-Token": frappe.csrf_token },
			success: (r) => {
				const file_url = r.message.file_url;
				$drop.html(
					`<div class="ei-drop-icon">📋</div><div class="ei-drop-text">${frappe.utils.escape_html(file.name)}</div><div class="ei-drop-sub">Reading sheets…</div>`
				);
				frappe.call({
					method: "erp.erp.page.import_from_excel.import_from_excel.get_sheet_names",
					args: { file_url },
					callback: (res) => {
						const sheet_names = res.message || [];
						if (sheet_names.length <= 1) {
							// Single sheet — parse directly
							const sheet = sheet_names[0] || null;
							$drop.html(
								`<div class="ei-drop-icon">⏳</div><div class="ei-drop-text">Parsing columns…</div>`
							);
							this._parse_sheet(file_url, sheet, file.name);
						} else {
							// Multiple sheets — show picker
							this.setState({
								file_url,
								file_name: file.name,
								sheet_names,
							});
						}
					},
					error: () =>
						$drop.html(
							'<div class="ei-drop-icon">❌</div><div class="ei-drop-text">Failed to read file</div>'
						),
				});
			},
			error: () =>
				$drop.html(
					'<div class="ei-drop-icon">❌</div><div class="ei-drop-text">Upload failed</div>'
				),
		});
	}

	_parse_sheet(file_url, sheet_name, file_name) {
		frappe.call({
			method: "erp.erp.page.import_from_excel.import_from_excel.parse_excel",
			args: { file_url, sheet_name: sheet_name || "" },
			callback: (res) => {
				if (res.message) {
					this.setState({
						file_url,
						file_name: file_name || this.state.file_name,
						selected_sheet: sheet_name,
						excel_columns: res.message.columns,
						total_rows: res.message.total_rows,
						step: 2,
					});
				}
			},
			error: () =>
				frappe.msgprint(__("Failed to parse the selected sheet.")),
		});
	}

	// ── Step 2: DocType Select ──────────────────────────────────────────────
	_render_step2($w) {
		const s = this.state;
		const $card = $('<div class="ei-card"></div>').appendTo($w);
		$('<div class="ei-card-title">Select DocType to Import Into</div>').appendTo(
			$card
		);

		const sheet_info = s.selected_sheet ? ` · sheet <b>${frappe.utils.escape_html(s.selected_sheet)}</b>` : "";
		$(
			`<div style="font-size:12px;color:var(--text-muted);margin-bottom:12px;">
				File: <b>${frappe.utils.escape_html(s.file_name)}</b>${sheet_info}
				· ${s.total_rows} data rows · ${s.excel_columns.length} columns
			</div>`
		).appendTo($card);

		const $search = $(
			'<input type="text" class="form-control" placeholder="Search doctypes…" style="max-width:280px;margin-bottom:10px;">'
		).appendTo($card);

		const $grid = $('<div class="ei-dt-grid"></div>').appendTo($card);

		const render_grid = (filter) => {
			$grid.empty();
			frappe.call({
				method: "erp.erp.page.import_from_excel.import_from_excel.get_all_doctypes",
				callback: (r) => {
					if (!r.message) return;
					const all = filter
						? r.message.filter((dt) =>
								dt.toLowerCase().includes(filter.toLowerCase())
						  )
						: r.message;
					all.forEach((dt) => {
						const is_suggested = s.suggested_doctype === dt;
						const $btn = $(
							`<button class="ei-dt-btn${is_suggested ? " suggested" : ""}">
								<span>📄</span>${frappe.utils.escape_html(dt)}${is_suggested ? ' <small style="color:var(--primary-color)">(suggested)</small>' : ""}
							</button>`
						).appendTo($grid);
						$btn.on("click", () => this._select_doctype(dt));
					});
					if (!all.length)
						$(
							'<div class="text-muted" style="font-size:13px;">No doctypes found.</div>'
						).appendTo($grid);
				},
			});
		};
		render_grid("");
		$search.on(
			"input",
			frappe.utils.debounce(() => render_grid($search.val()), 280)
		);

		const $nav = $('<div class="ei-nav"></div>').appendTo($card);
		$('<button class="btn btn-default btn-sm">← Back</button>')
			.appendTo($nav)
			.on("click", () => this.setState({ step: 1 }));
	}

	_select_doctype(doctype) {
		frappe.call({
			method: "erp.erp.page.import_from_excel.import_from_excel.get_doctype_fields",
			args: { doctype },
			callback: (r) => {
				if (r.message) {
					this.setState({
						doctype,
						doctype_fields: r.message.fields,
						mappings: [],
						selected_x: null,
						selected_f: null,
						suggested_doctype: null,
						step: 3,
					});
				}
			},
		});
	}

	// ── Step 3: Column Mapping ──────────────────────────────────────────────
	_render_step3($w) {
		const s = this.state;
		const mapped_x = new Set(s.mappings.map((m) => m.xid));
		const mapped_f = new Set(s.mappings.map((m) => m.fid));
		const fields_view = s.required_only
			? s.doctype_fields.filter((f) => f.reqd)
			: s.doctype_fields;
		const { left_list, right_list } = this._compute_sort(
			s.excel_columns,
			fields_view,
			s.mappings,
			s.sort_mode
		);

		const unmapped_x = s.excel_columns.length - s.mappings.length;
		const req_missing = s.doctype_fields.filter(
			(f) => f.reqd && !mapped_f.has(f.id)
		).length;

		// Nav — top, before everything else
		const can_continue = req_missing === 0 && s.mappings.length > 0;
		const $nav = $('<div class="ei-nav" style="margin-bottom:12px;"></div>').appendTo($w);
		$('<button class="btn btn-default btn-sm">← Back</button>')
			.appendTo($nav)
			.on("click", () => this.setState({ step: 2 }));
		const $cont = $(`<button class="btn btn-primary btn-sm${can_continue ? "" : " disabled"}">Continue →</button>`).appendTo($nav);
		if (!can_continue) {
			const reason =
				s.mappings.length === 0
					? "Map at least one column to continue"
					: `Map ${req_missing} required field${req_missing !== 1 ? "s" : ""} to continue`;
			$cont.attr("title", reason);
		}
		$cont.on("click", () => { if (can_continue) this._load_preview(); });
		if (req_missing > 0 && s.mappings.length > 0) {
			$('<span style="font-size:12px;color:var(--red,#e53e3e);margin-left:4px;">⚠ Map all required fields (*) to continue.</span>').appendTo($nav);
		}

		// Toolbar
		const $tb = $('<div class="ei-toolbar"></div>').appendTo($w);

		$(`<span style="font-size:12px;font-weight:600;padding:3px 10px;border:1px solid var(--border-color);border-radius:6px;background:var(--fg-color);">
			→ <b>${frappe.utils.escape_html(s.doctype)}</b>
		</span>`).appendTo($tb);

		$(`<span class="ei-counter"><span class="ei-dot"></span>${unmapped_x} excel cols unmapped</span>`).appendTo(
			$tb
		);

		const req_color = req_missing
			? "var(--red, #e53e3e)"
			: "var(--green, #2f9a5c)";
		$(`<span class="ei-counter" style="border-color:${req_color}">
			<span class="ei-dot" style="background:${req_color}"></span>
			${req_missing} required missing
		</span>`).appendTo($tb);

		$('<span class="ei-spacer"></span>').appendTo($tb);

		// Sort
		const $sort = $(
			`<select class="form-control form-control-sm" style="width:auto;">
				<option value="original">↕ Original order</option>
				<option value="untangle">⤺ Untangle arrows</option>
				<option value="mapped-first">⬆ Mapped first</option>
			</select>`
		)
			.val(s.sort_mode)
			.appendTo($tb);
		$sort.on("change", () => this.setState({ sort_mode: $sort.val() }));

		$('<button class="btn btn-default btn-xs">✨ Auto-suggest</button>')
			.appendTo($tb)
			.on("click", () => this._auto_suggest());

		$(`<button class="btn btn-xs ${s.show_sample ? "btn-info" : "btn-default"}">Sample</button>`)
			.appendTo($tb)
			.on("click", () => this.setState({ show_sample: !s.show_sample }));

		$(`<button class="btn btn-xs ${s.required_only ? "btn-info" : "btn-default"}">Req. only</button>`)
			.appendTo($tb)
			.on("click", () =>
				this.setState({ required_only: !s.required_only })
			);

		$(`<button class="btn btn-xs btn-default">${s.density === "compact" ? "⊟ Compact" : "⊞ Cozy"}</button>`)
			.appendTo($tb)
			.on("click", () =>
				this.setState({
					density:
						s.density === "compact" ? "comfortable" : "compact",
				})
			);

		// Two-column
		const $two = $(
			'<div class="ei-two-col" id="ei-two-col"></div>'
		).appendTo($w);

		// ── Left: Excel columns
		const $left = $('<div class="ei-col-cards"></div>').appendTo($two);
		$('<div class="ei-col-head">Excel columns</div>').appendTo($left);

		left_list.forEach((x) => {
			const is_m = mapped_x.has(x.id);
			const is_s = s.selected_x === x.id;
			const cls = ["ei-fcard", s.density === "compact" ? "compact" : "", is_m ? "mapped" : "", is_s ? "selected" : ""]
				.filter(Boolean).join(" ");

			const $card = $(
				`<div class="${cls}" data-xid="${x.id}"></div>`
			).appendTo($left);
			$('<span class="ei-dot-r"></span>').appendTo($card);
			$(`<span class="ei-fname">${frappe.utils.escape_html(x.name)}</span>`).appendTo($card);
			$(`<span class="ei-ftype">${x.type}</span>`).appendTo($card);
			if (s.show_sample && x.sample)
				$(`<span class="ei-fsample" title="${frappe.utils.escape_html(x.sample)}">${frappe.utils.escape_html(x.sample)}</span>`).appendTo($card);
			$card.on("click", () => this._click_x(x.id));
		});

		// ── Center spacer (no content — just gap for connectors)
		$('<div class="ei-center-col"></div>').appendTo($two);

		// ── Right: DocType fields
		const $right = $('<div class="ei-col-cards"></div>').appendTo($two);
		$(`<div class="ei-col-head right">${frappe.utils.escape_html(s.doctype)} fields</div>`).appendTo($right);

		right_list.forEach((f) => {
			const is_m = mapped_f.has(f.id);
			const is_s = s.selected_f === f.id;
			const mapping = s.mappings.find((m) => m.fid === f.id);
			const cls = ["ei-fcard", s.density === "compact" ? "compact" : "", is_m ? "mapped" : "", is_s ? "selected" : ""]
				.filter(Boolean).join(" ");

			const $card = $(
				`<div class="${cls}" data-fid="${f.id}"></div>`
			).appendTo($right);
			$('<span class="ei-dot-l"></span>').appendTo($card);
			$(`<span class="ei-fname">${frappe.utils.escape_html(f.label)}</span>`).appendTo($card);
			if (f.reqd)
				$('<span class="ei-freq" title="required">*</span>').appendTo($card);
			$(`<span class="ei-ftype">${frappe.utils.escape_html(f.fieldtype)}</span>`).appendTo($card);

			if (is_m && mapping) {
				const xcol = s.excel_columns.find((x) => x.id === mapping.xid);
				$(`<span class="ei-rm-pill">← ${frappe.utils.escape_html(xcol?.name || "")} ✕</span>`)
					.appendTo($card)
					.on("click", (e) => {
						e.stopPropagation();
						this._remove_mapping(null, f.id);
					});
			}

			$card.on("click", () => this._click_f(f.id));
		});

		// Draw SVG connectors after layout settles
		requestAnimationFrame(() =>
			requestAnimationFrame(() =>
				this._draw_connectors($two[0], s.mappings)
			)
		);

	}

	_click_x(xid) {
		const s = this.state;
		if (s.selected_f) {
			this._try_pair(xid, s.selected_f);
		} else {
			this.setState({ selected_x: s.selected_x === xid ? null : xid });
		}
	}

	_click_f(fid) {
		const s = this.state;
		if (s.selected_x) {
			this._try_pair(s.selected_x, fid);
		} else {
			this.setState({ selected_f: s.selected_f === fid ? null : fid });
		}
	}

	_try_pair(xid, fid) {
		const mappings = this.state.mappings.filter(
			(m) => m.xid !== xid && m.fid !== fid
		);
		mappings.push({ xid, fid });
		this.setState({ mappings, selected_x: null, selected_f: null });
	}

	_remove_mapping(xid, fid) {
		const mappings = this.state.mappings.filter(
			(m) => (xid ? m.xid !== xid : true) && (fid ? m.fid !== fid : true)
		);
		this.setState({ mappings });
	}

	_auto_suggest() {
		const { excel_columns, doctype_fields } = this.state;
		const mappings = [];
		const used = new Set();
		for (const xcol of excel_columns) {
			let best = null,
				best_score = 0;
			for (const field of doctype_fields) {
				if (used.has(field.id)) continue;
				const score = this._match_score(
					xcol.name,
					field.label || field.fieldname
				);
				if (score > best_score && score > 0.45) {
					best_score = score;
					best = field;
				}
			}
			if (best) {
				mappings.push({ xid: xcol.id, fid: best.id });
				used.add(best.id);
			}
		}
		this.setState({ mappings });
	}

	_match_score(a, b) {
		const norm = (s) => s.toLowerCase().replace(/[_\s\-.]+/g, "");
		a = norm(a);
		b = norm(b);
		if (a === b) return 1;
		if (a.includes(b) || b.includes(a)) return 0.85;
		const bigrams = (s) => {
			const r = new Set();
			for (let i = 0; i < s.length - 1; i++) r.add(s[i] + s[i + 1]);
			return r;
		};
		const bg_a = bigrams(a),
			bg_b = bigrams(b);
		if (!bg_a.size && !bg_b.size) return 0;
		const common = [...bg_a].filter((bg) => bg_b.has(bg)).length;
		return (2 * common) / (bg_a.size + bg_b.size);
	}

	_compute_sort(left_orig, right_orig, mappings, sort_mode) {
		let left_list = [...left_orig];
		let right_list = [...right_orig];

		if (sort_mode === "untangle") {
			const x_idx = Object.fromEntries(
				left_orig.map((x, i) => [x.id, i])
			);
			const map_by_fid = Object.fromEntries(
				mappings.map((m) => [m.fid, m.xid])
			);
			right_list.sort((a, b) => {
				const ax = map_by_fid[a.id],
					bx = map_by_fid[b.id];
				if (ax && bx) return x_idx[ax] - x_idx[bx];
				if (ax) return -1;
				if (bx) return 1;
				return 0;
			});
		} else if (sort_mode === "mapped-first") {
			const x_order = mappings.map((m) => m.xid);
			left_list.sort((a, b) => {
				const ai = x_order.indexOf(a.id),
					bi = x_order.indexOf(b.id);
				if (ai >= 0 && bi >= 0) return ai - bi;
				if (ai >= 0) return -1;
				if (bi >= 0) return 1;
				return 0;
			});
			const map_by_xid = Object.fromEntries(
				mappings.map((m) => [m.xid, m.fid])
			);
			const left_idx_of = (fid) => {
				const e = Object.entries(map_by_xid).find(([, f]) => f === fid);
				return e ? left_list.findIndex((x) => x.id === e[0]) : -1;
			};
			right_list.sort((a, b) => {
				const ai = left_idx_of(a.id),
					bi = left_idx_of(b.id);
				if (ai >= 0 && bi >= 0) return ai - bi;
				if (ai >= 0) return -1;
				if (bi >= 0) return 1;
				return 0;
			});
		}
		return { left_list, right_list };
	}

	// ── SVG connector layer ─────────────────────────────────────────────────
	_draw_connectors($container, mappings) {
		if (!$container) return;
		$container.querySelectorAll(".ei-svg-layer").forEach((el) => el.remove());
		if (!mappings.length) return;

		const root = $container.getBoundingClientRect();
		const paths = [];

		mappings.forEach((m) => {
			const le = $container.querySelector(`[data-xid="${m.xid}"] .ei-dot-r`);
			const re = $container.querySelector(`[data-fid="${m.fid}"] .ei-dot-l`);
			if (!le || !re) return;
			const lr = le.getBoundingClientRect();
			const rr = re.getBoundingClientRect();
			const x1 = lr.left + lr.width / 2 - root.left;
			const y1 = lr.top + lr.height / 2 - root.top;
			const x2 = rr.left + rr.width / 2 - root.left;
			const y2 = rr.top + rr.height / 2 - root.top;
			const dx = (x2 - x1) * 0.55;
			paths.push({
				xid: m.xid,
				fid: m.fid,
				d: `M ${x1} ${y1} C ${x1 + dx} ${y1}, ${x2 - dx} ${y2}, ${x2} ${y2}`,
				mx: (x1 + x2) / 2,
				my: (y1 + y2) / 2,
			});
		});

		if (!paths.length) return;

		const ns = "http://www.w3.org/2000/svg";
		const svg = document.createElementNS(ns, "svg");
		svg.setAttribute("class", "ei-svg-layer");
		svg.setAttribute("width", $container.offsetWidth);
		svg.setAttribute("height", $container.offsetHeight);
		svg.style.cssText =
			"position:absolute;top:0;left:0;pointer-events:none;overflow:visible;";

		const defs = document.createElementNS(ns, "defs");
		const marker = document.createElementNS(ns, "marker");
		marker.setAttribute("id", "ei-arr");
		marker.setAttribute("viewBox", "0 0 10 10");
		marker.setAttribute("refX", "8");
		marker.setAttribute("refY", "5");
		marker.setAttribute("markerWidth", "5");
		marker.setAttribute("markerHeight", "5");
		marker.setAttribute("orient", "auto-start-reverse");
		const ap = document.createElementNS(ns, "path");
		ap.setAttribute("d", "M0,0 L10,5 L0,10 z");
		ap.setAttribute("fill", "var(--primary-color,#2490ef)");
		marker.appendChild(ap);
		defs.appendChild(marker);
		svg.appendChild(defs);

		const primary = "var(--primary-color,#2490ef)";
		paths.forEach((p) => {
			const g = document.createElementNS(ns, "g");
			g.style.pointerEvents = "auto";

			const hit = document.createElementNS(ns, "path");
			hit.setAttribute("d", p.d);
			hit.setAttribute("fill", "none");
			hit.setAttribute("stroke", "transparent");
			hit.setAttribute("stroke-width", "20");
			hit.style.cursor = "pointer";
			hit.addEventListener("click", () =>
				this._remove_mapping(p.xid, p.fid)
			);

			const vis = document.createElementNS(ns, "path");
			vis.setAttribute("d", p.d);
			vis.setAttribute("fill", "none");
			vis.setAttribute("stroke", primary);
			vis.setAttribute("stroke-width", "2.5");
			vis.setAttribute("stroke-linecap", "round");
			vis.setAttribute("marker-end", "url(#ei-arr)");
			vis.style.transition = "stroke-width .1s";

			let hover_g = null;
			g.addEventListener("mouseenter", () => {
				vis.setAttribute("stroke-width", "4");
				hover_g = document.createElementNS(ns, "g");
				hover_g.style.cursor = "pointer";
				hover_g.addEventListener("click", () =>
					this._remove_mapping(p.xid, p.fid)
				);
				const circ = document.createElementNS(ns, "circle");
				circ.setAttribute("cx", p.mx);
				circ.setAttribute("cy", p.my);
				circ.setAttribute("r", "10");
				circ.setAttribute("fill", "white");
				circ.setAttribute("stroke", primary);
				circ.setAttribute("stroke-width", "2");
				const cross = document.createElementNS(ns, "path");
				cross.setAttribute(
					"d",
					`M ${p.mx - 4} ${p.my - 4} L ${p.mx + 4} ${p.my + 4} M ${p.mx + 4} ${p.my - 4} L ${p.mx - 4} ${p.my + 4}`
				);
				cross.setAttribute("stroke", primary);
				cross.setAttribute("stroke-width", "2");
				cross.setAttribute("stroke-linecap", "round");
				hover_g.appendChild(circ);
				hover_g.appendChild(cross);
				g.appendChild(hover_g);
			});
			g.addEventListener("mouseleave", () => {
				vis.setAttribute("stroke-width", "2.5");
				if (hover_g) {
					hover_g.remove();
					hover_g = null;
				}
			});

			g.appendChild(hit);
			g.appendChild(vis);
			svg.appendChild(g);
		});

		$container.style.position = "relative";
		$container.appendChild(svg);
	}

	// ── Step 4: Validate ────────────────────────────────────────────────────
	_load_preview() {
		const s = this.state;
		frappe.call({
			method: "erp.erp.page.import_from_excel.import_from_excel.get_preview_data",
			args: {
				doctype: s.doctype,
				mappings: JSON.stringify(this._build_api_mappings()),
				file_url: s.file_url,
				sheet_name: s.selected_sheet || "",
			},
			callback: (r) => {
				if (r.message) this.setState({ preview_data: r.message, step: 4 });
			},
		});
	}

	_render_step4($w) {
		const s = this.state;
		const pd = s.preview_data;
		const $card = $('<div class="ei-card"></div>').appendTo($w);
		$('<div class="ei-card-title">Preview Mapped Data</div>').appendTo($card);
		$(
			`<div style="font-size:12px;color:var(--text-muted);margin-bottom:12px;">
				Showing first ${pd.rows.length} of <b>${pd.total}</b> rows.
			</div>`
		).appendTo($card);

		const $table = $('<table class="ei-tbl"></table>').appendTo($card);
		const $tr_h = $("<thead><tr></tr></thead>").appendTo($table).find("tr");
		pd.headers.forEach((h) =>
			$(`<th>${frappe.utils.escape_html(h)}</th>`).appendTo($tr_h)
		);

		const $tbody = $("<tbody></tbody>").appendTo($table);
		pd.rows.forEach((row) => {
			const $tr = $("<tr></tr>").appendTo($tbody);
			pd.headers.forEach((h) =>
				$(
					`<td>${frappe.utils.escape_html(
						row[h] != null ? String(row[h]) : ""
					)}</td>`
				).appendTo($tr)
			);
		});

		const $nav = $('<div class="ei-nav"></div>').appendTo($card);
		$('<button class="btn btn-default btn-sm">← Back to mapping</button>')
			.appendTo($nav)
			.on("click", () => this.setState({ step: 3 }));
		$(`<button class="btn btn-primary btn-sm">Import all ${pd.total} rows →</button>`)
			.appendTo($nav)
			.on("click", () => this._do_import());
	}

	_do_import() {
		frappe.show_progress("Importing…", 0, 100, "Please wait");
		frappe.call({
			method: "erp.erp.page.import_from_excel.import_from_excel.import_records",
			type: "POST",
			args: {
				doctype: this.state.doctype,
				mappings: JSON.stringify(this._build_api_mappings()),
				file_url: this.state.file_url,
				sheet_name: this.state.selected_sheet || "",
			},
			callback: (r) => {
				frappe.hide_progress();
				if (r.message) this.setState({ import_result: r.message, step: 5 });
			},
			error: () => {
				frappe.hide_progress();
				frappe.msgprint(__("Import failed. Please try again."));
			},
		});
	}

	_build_api_mappings() {
		return this.state.mappings
			.map((m) => {
				const xcol = this.state.excel_columns.find((x) => x.id === m.xid);
				const field = this.state.doctype_fields.find((f) => f.id === m.fid);
				return { excel_col: xcol?.name, fieldname: field?.fieldname };
			})
			.filter((m) => m.excel_col && m.fieldname);
	}

	// ── Step 5: Success + Link Error Resolution ─────────────────────────────
	_render_step5($w) {
		const s = this.state;
		const r = s.import_result;

		// Success card
		const $card = $('<div class="ei-card"></div>').appendTo($w);
		const $success = $('<div class="ei-success"></div>').appendTo($card);
		$('<div class="ei-s-icon">✓</div>').appendTo($success);
		$(
			`<div class="ei-s-title">${r.imported} ${frappe.utils.escape_html(s.doctype)} record${r.imported !== 1 ? "s" : ""} imported</div>`
		).appendTo($success);
		if (r.skipped > 0) {
			$(
				`<div style="color:var(--text-muted);font-size:13px;">${r.skipped} row${r.skipped !== 1 ? "s" : ""} had errors</div>`
			).appendTo($success);
		}
		const $stats = $('<div class="ei-stats"></div>').appendTo($success);
		$(
			`<div><span class="ei-stat-n" style="color:var(--green,#2f9a5c)">${r.imported}</span><div class="ei-stat-l">imported</div></div>`
		).appendTo($stats);
		if (r.skipped) {
			$(
				`<div><span class="ei-stat-n" style="color:var(--red,#e53e3e)">${r.skipped}</span><div class="ei-stat-l">errors</div></div>`
			).appendTo($stats);
		}
		$(
			`<div><span class="ei-stat-n">${s.mappings.length}</span><div class="ei-stat-l">cols mapped</div></div>`
		).appendTo($stats);

		// Bottom nav
		const $nav = $('<div class="ei-nav" style="justify-content:center;margin-top:16px;"></div>').appendTo(
			$card
		);
		$('<button class="btn btn-default btn-sm">Import another file</button>')
			.appendTo($nav)
			.on("click", () =>
				this.setState({
					step: 1,
					file_url: null, file_name: null,
					sheet_names: [], selected_sheet: null,
					excel_columns: [], doctype: null, doctype_fields: [],
					mappings: [], preview_data: null, import_result: null,
				})
			);
		const slug = s.doctype.toLowerCase().replace(/ /g, "-");
		$(
			`<a class="btn btn-primary btn-sm" href="/app/${slug}">View ${frappe.utils.escape_html(s.doctype)} →</a>`
		).appendTo($nav);

		// ── Link error resolution ──────────────────────────────────────────
		const link_groups = this._parse_link_errors(r.errors || []);
		if (Object.keys(link_groups).length === 0) return;

		const $lerr = $('<div class="ei-link-err-card"></div>').appendTo($w);
		$(
			`<div class="ei-link-err-title">⚠ Unresolved Link Errors</div>`
		).appendTo($lerr);
		$(
			`<div style="font-size:12px;color:var(--text-muted);margin-bottom:14px;">
				These rows were skipped because linked records could not be found.
				Fix the linked data, then retry the import.
			</div>`
		).appendTo($lerr);

		Object.entries(link_groups).forEach(([doctype, values]) => {
			const $grp = $(
				'<div class="ei-link-err-group"></div>'
			).appendTo($lerr);
			$(
				`<div style="font-size:13px;font-weight:600;">Could not find <b>${frappe.utils.escape_html(doctype)}</b> — ${values.length} row(s) affected</div>`
			).appendTo($grp);

			const preview_vals = values
				.slice(0, 6)
				.map((v) => `<code>${frappe.utils.escape_html(v)}</code>`)
				.join(", ");
			$(
				`<div class="ei-missing-vals">Missing: ${preview_vals}${values.length > 6 ? ` <span style="color:var(--text-muted)">+${values.length - 6} more</span>` : ""}</div>`
			).appendTo($grp);

			const $btns = $(
				'<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;"></div>'
			).appendTo($grp);

			// Option A — import that doctype from a fresh file
			$(
				`<button class="btn btn-default btn-sm">📥 Import ${frappe.utils.escape_html(doctype)} from Excel</button>`
			)
				.appendTo($btns)
				.on("click", () => this._reset_for_doctype(doctype));

			// Option B — auto-create (only if eligible, checked async)
			const $auto_slot = $("<span></span>").appendTo($btns);
			frappe.call({
				method: "erp.erp.page.import_from_excel.import_from_excel.check_can_auto_create",
				args: { doctype },
				callback: (res) => {
					if (res.message && res.message.can_auto) {
						$(
							`<button class="btn btn-primary btn-sm">✨ Auto-create ${values.length} missing ${frappe.utils.escape_html(doctype)} record${values.length !== 1 ? "s" : ""}</button>`
						)
							.appendTo($auto_slot)
							.on("click", () =>
								this._auto_create(doctype, values, $grp)
							);
					}
				},
			});
		});

		// Retry import button (shared, appears after any auto-create)
		this._$retry_slot = $('<div style="margin-top:14px;"></div>').appendTo(
			$lerr
		);
	}

	_parse_link_errors(errors) {
		const re = /Could not find ([^:]+): (.+)/;
		const groups = {};
		errors.forEach((e) => {
			const m = String(e.error).match(re);
			if (!m) return;
			const doctype = m[1].trim();
			const value = m[2].trim();
			if (!groups[doctype]) groups[doctype] = [];
			if (!groups[doctype].includes(value)) groups[doctype].push(value);
		});
		return groups;
	}

	_auto_create(doctype, values, $grp) {
		frappe.call({
			method: "erp.erp.page.import_from_excel.import_from_excel.auto_create_linked",
			type: "POST",
			args: { doctype, values: JSON.stringify(values) },
			callback: (r) => {
				if (!r.message) return;
				const { created, errors: errs } = r.message;

				$grp.find(".ei-auto-result").remove();

				if (created.length > 0) {
					$(
						`<div class="ei-auto-result" style="color:var(--green,#2f9a5c);font-size:12px;margin-top:6px;">
							✓ Created ${created.length} ${frappe.utils.escape_html(doctype)} record${created.length !== 1 ? "s" : ""}.
						</div>`
					).appendTo($grp);

					// Show retry button once (in shared slot)
					if (this._$retry_slot && !this._$retry_slot.find(".ei-retry-btn").length) {
						$(
							'<button class="btn btn-primary btn-sm ei-retry-btn">🔄 Retry Import with Fixed Data</button>'
						)
							.appendTo(this._$retry_slot)
							.on("click", () => this._do_import());
					}
				}
				if (errs && errs.length > 0) {
					$(
						`<div class="ei-auto-result" style="color:var(--red,#e53e3e);font-size:12px;margin-top:4px;">
							❌ Could not create: ${errs.map((e) => frappe.utils.escape_html(e.value)).join(", ")}
						</div>`
					).appendTo($grp);
				}
			},
		});
	}

	_reset_for_doctype(doctype) {
		this.setState({
			step: 1,
			file_url: null, file_name: null,
			sheet_names: [], selected_sheet: null,
			excel_columns: [], total_rows: 0,
			doctype: null, doctype_fields: [],
			mappings: [], selected_x: null, selected_f: null,
			preview_data: null, import_result: null,
			suggested_doctype: doctype,
		});
	}
}
