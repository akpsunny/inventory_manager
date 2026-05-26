"""
Inventory Data Browser — separate Toplevel window with filtering, sorting,
and item-level CRUD.

Opened from main.py via a button.  Operates on the user's master .db via
the DatabaseManager passed in by the caller; does not import from main.py
to avoid a circular dependency.  Theme constants are duplicated here to
keep the module self-sufficient until a shared theme module exists.
"""
from __future__ import annotations

import threading
from pathlib import Path
from typing import Callable, Optional

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk


# ─── Theme (duplicated from main.py — keep in sync) ─────────────────────────
DARK_BG     = "#1A1B1E"
CARD_BG     = "#252629"
SURFACE_BG  = "#2D2E32"
THEME_COLOR = "#0F5132"
ACCENT_2    = "#1E7B4A"
TEXT_MAIN   = "#E8E9EB"
TEXT_MUTED  = "#A0A3A8"
ERROR_CLR   = "#DC3545"
WARN_CLR    = "#F0AD4E"
OK_CLR      = "#28A745"

LOW_STOCK_THRESHOLD = 5     # qty ≤ this = "Low Stock"


# ─── Status formatting helper ───────────────────────────────────────────────

def _format_status(available: float) -> tuple[str, str]:
    """Return (label, hex_color) for the status column."""
    if available <= 0:
        return ("Out of Stock", ERROR_CLR)
    if available <= LOW_STOCK_THRESHOLD:
        return ("Low Stock", WARN_CLR)
    return ("In Stock", OK_CLR)


# ─── Item Editor dialog ─────────────────────────────────────────────────────

class ItemDialog(ctk.CTkToplevel):
    """
    Modal dialog for adding a new item or editing an existing one.

    Used in two modes:
      mode='add'  : item_data is None.  Stock-adjustment section hidden.
      mode='edit' : item_data is the snapshot dict.  Stock-adjustment shown.

    On Save: calls on_save(fields_dict, adjust_delta, adjust_note) with
    parameters relevant to the caller's persistence step.  The dialog
    closes only if on_save returns True (the caller signals success).
    """

    # Fields rendered in the form, in display order.
    # Each tuple: (key, label, optional input width)
    _FIELDS = [
        ("item_name",   "Item Name *",  None),
        ("sku_code",    "SKU Code",     None),
        ("hsn_sac",     "HSN / SAC",    None),
        ("gst_rate",    "GST Rate",     120),
        ("rate",        "Rate (₹)",     120),
        ("measurement", "Unit",         120),
    ]

    def __init__(self, master, *, mode: str, item_data: Optional[dict],
                 on_save: Callable[[dict, float, str], bool]):
        super().__init__(master)
        self.mode      = mode
        self.item_data = item_data or {}
        self._on_save  = on_save
        self._entries: dict[str, ctk.CTkEntry] = {}

        # Window chrome
        self.title("Add Item" if mode == "add" else
                   f"Edit Item — {self.item_data.get('item_name', '')[:40]}")
        self.configure(fg_color=DARK_BG)
        self.resizable(False, False)
        self.transient(master)
        # On Linux/X11 a Toplevel sometimes needs a tick before grab_set works
        self.after(50, self._safe_grab)

        # ── Header label ─────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text=("Add Item" if mode == "add" else "Edit Item"),
            font=("Segoe UI", 16, "bold"), text_color=TEXT_MAIN,
        ).pack(padx=24, pady=(18, 4))

        if mode == "edit":
            ctk.CTkLabel(
                self,
                text=f"ID #{self.item_data.get('id', '?')}",
                font=("Segoe UI", 10), text_color=TEXT_MUTED,
            ).pack(padx=24, pady=(0, 8))

        # ── Form fields ──────────────────────────────────────────────────
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=24, pady=(8, 8), fill="x")

        for row_idx, (key, label, width) in enumerate(self._FIELDS):
            ctk.CTkLabel(
                form, text=label, font=("Segoe UI", 11),
                text_color=TEXT_MAIN, anchor="w", width=110,
            ).grid(row=row_idx, column=0, sticky="w", padx=(0, 12), pady=4)

            entry = ctk.CTkEntry(
                form, font=("Segoe UI", 11),
                width=(width or 320), fg_color=SURFACE_BG,
                border_color=THEME_COLOR, text_color=TEXT_MAIN)
            entry.grid(row=row_idx, column=1, sticky="w", pady=4)
            entry.insert(0, self._initial_value(key))
            self._entries[key] = entry

        # Focus into the first empty field
        for key, _, _ in self._FIELDS:
            if not self._entries[key].get():
                self._entries[key].focus_set()
                break
        else:
            self._entries["item_name"].focus_set()

        # ── Stock adjustment (edit mode only) ────────────────────────────
        self._adj_entry: Optional[ctk.CTkEntry] = None
        self._adj_note:  Optional[ctk.CTkEntry] = None
        if mode == "edit":
            self._build_adjustment_section()

        # ── Action buttons ───────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=24, pady=(12, 20), fill="x")

        ctk.CTkButton(
            btn_row, text="Cancel", width=110, height=34,
            font=("Segoe UI", 11), fg_color=SURFACE_BG,
            border_width=1, border_color="#555",
            hover_color="#3A3A3A",
            command=self._on_cancel
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            btn_row, text=("Add" if mode == "add" else "Save"),
            width=110, height=34, font=("Segoe UI", 11, "bold"),
            fg_color=THEME_COLOR, hover_color=ACCENT_2,
            command=self._on_submit
        ).pack(side="right")

        self.bind("<Return>", lambda _e: self._on_submit())
        self.bind("<Escape>", lambda _e: self._on_cancel())
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    # ── helpers ────────────────────────────────────────────────────────

    def _safe_grab(self):
        try:    self.grab_set()
        except tk.TclError: pass

    def _initial_value(self, key: str) -> str:
        v = self.item_data.get(key, "")
        if key == "rate":
            return f"{float(v):.2f}" if v else ""
        return str(v or "")

    def _build_adjustment_section(self):
        sep = ctk.CTkFrame(self, fg_color="#3A3A3A", height=1)
        sep.pack(padx=24, pady=(8, 8), fill="x")

        ctk.CTkLabel(
            self, text="Manual Stock Adjustment",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_MAIN,
        ).pack(padx=24, pady=(0, 4), anchor="w")

        avail = self.item_data.get("available", 0)
        rcv   = self.item_data.get("total_received", 0)
        cons  = self.item_data.get("total_consumed", 0)
        adj   = self.item_data.get("adjustments", 0)
        ctk.CTkLabel(
            self,
            text=(f"Received: {rcv:.0f}   Consumed: {cons:.0f}   "
                  f"Prior adjustments: {adj:+.0f}   "
                  f"Current available: {avail:.0f}"),
            font=("Segoe UI", 10), text_color=TEXT_MUTED,
        ).pack(padx=24, pady=(0, 8), anchor="w")

        adj_form = ctk.CTkFrame(self, fg_color="transparent")
        adj_form.pack(padx=24, pady=(0, 4), fill="x")

        ctk.CTkLabel(
            adj_form, text="Adjust by", font=("Segoe UI", 11),
            text_color=TEXT_MAIN, anchor="w", width=110,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        self._adj_entry = ctk.CTkEntry(
            adj_form, font=("Segoe UI", 11), width=120,
            fg_color=SURFACE_BG, border_color=THEME_COLOR,
            placeholder_text="+5 or -2",
            text_color=TEXT_MAIN)
        self._adj_entry.grid(row=0, column=1, sticky="w", pady=4)
        ctk.CTkLabel(
            adj_form, text="(positive adds, negative removes)",
            font=("Segoe UI", 10), text_color=TEXT_MUTED,
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        ctk.CTkLabel(
            adj_form, text="Note", font=("Segoe UI", 11),
            text_color=TEXT_MAIN, anchor="w", width=110,
        ).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=4)
        self._adj_note = ctk.CTkEntry(
            adj_form, font=("Segoe UI", 11), width=320,
            fg_color=SURFACE_BG, border_color=THEME_COLOR,
            placeholder_text="e.g. 'damaged in transit'",
            text_color=TEXT_MAIN)
        self._adj_note.grid(row=1, column=1, columnspan=2, sticky="w", pady=4)

    # ── actions ────────────────────────────────────────────────────────

    def _collect_fields(self) -> dict:
        return {k: self._entries[k].get().strip()
                for k, _, _ in self._FIELDS}

    def _parse_adjustment(self) -> tuple[float, str]:
        if self._adj_entry is None:
            return (0.0, "")
        raw = self._adj_entry.get().strip()
        if not raw:
            return (0.0, "")
        try:
            delta = float(raw)
        except ValueError:
            raise ValueError(
                f"'{raw}' isn't a valid number. Use a value like +5 or -2.")
        note = self._adj_note.get().strip() if self._adj_note else ""
        return (delta, note)

    def _on_submit(self):
        fields = self._collect_fields()
        if not fields.get("item_name"):
            messagebox.showwarning("Required", "Item Name is required.",
                                   parent=self)
            self._entries["item_name"].focus_set()
            return
        # Validate rate parses
        if fields.get("rate"):
            try:
                fields["rate"] = float(fields["rate"])
            except ValueError:
                messagebox.showwarning(
                    "Bad number", f"Rate '{fields['rate']}' isn't a number.",
                    parent=self)
                self._entries["rate"].focus_set()
                return
        else:
            fields["rate"] = 0.0

        try:
            delta, note = self._parse_adjustment()
        except ValueError as e:
            messagebox.showwarning("Bad number", str(e), parent=self)
            self._adj_entry.focus_set()
            return

        ok = self._on_save(fields, delta, note)
        if ok:
            self.destroy()

    def _on_cancel(self):
        self.destroy()


# ─── Main browser window ────────────────────────────────────────────────────

class DataBrowserWindow(ctk.CTkToplevel):
    """
    Tabular inventory browser.

    Constructor parameters
    ----------------------
    master : the main app window (used as parent + transient).
    excel_manager : DatabaseManager instance.
    get_db_path : callable returning the currently-selected master .db path,
                  re-evaluated on every refresh so the browser stays in sync
                  if the user switches masters.
    log_fn : optional logger (label, level) used for main-window log messages.
    """

    # Columns: (key, header, anchor, width, treeview_id, numeric_for_sort)
    _COLS = [
        ("id",          "ID",          "center",  50,  "id",          True),
        ("item_name",   "Item Name",   "w",      260,  "item_name",   False),
        ("sku_code",    "SKU",         "w",      130,  "sku_code",    False),
        ("hsn_sac",     "HSN/SAC",     "center", 100,  "hsn_sac",     False),
        ("gst_rate",    "GST",         "center",  60,  "gst_rate",    False),
        ("rate",        "Rate",        "e",       80,  "rate",        True),
        ("measurement", "Unit",        "center",  70,  "measurement", False),
        ("total_received", "Recv",     "e",       70,  "received",    True),
        ("total_consumed", "Cons",     "e",       70,  "consumed",    True),
        ("adjustments",    "Adj",      "e",       70,  "adjustments", True),
        ("available",      "Avail",    "e",       80,  "available",   True),
        ("__status__",     "Status",   "center", 110,  "status",      False),
    ]

    def __init__(self, master, *, excel_manager, get_db_path: Callable[[], str],
                 log_fn: Optional[Callable[[str, str], None]] = None):
        super().__init__(master)
        self.dm           = excel_manager
        self._get_db_path = get_db_path
        self._log         = log_fn or (lambda msg, lvl="info": None)

        self.title("Inventory Browser")
        self.configure(fg_color=DARK_BG)
        self.geometry("1280x720")
        self.minsize(960, 520)
        self.transient(master)

        # Snapshot held in memory; filter operates on this without re-querying
        self._snapshot: list[dict] = []
        self._filtered: list[dict] = []
        self._sort_key: str = "id"
        self._sort_reverse: bool = False

        # ── Layout: toolbar (top) | tree (middle) | footer (bottom) ─────
        self._build_toolbar()
        self._build_tree()
        self._build_footer()

        # Load data after window is mapped (Treeview needs sizing first)
        self.after(100, self.refresh)

    # ── toolbar ────────────────────────────────────────────────────────

    def _build_toolbar(self):
        bar = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0, height=110)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # Row 1 — filters
        row1 = ctk.CTkFrame(bar, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=(12, 4))

        ctk.CTkLabel(row1, text="🔍",
                     font=("Segoe UI", 13), text_color=TEXT_MUTED
                     ).pack(side="left", padx=(0, 4))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filter())
        ctk.CTkEntry(
            row1, textvariable=self._search_var,
            placeholder_text="Search by name or SKU…",
            font=("Segoe UI", 11), width=320, height=30,
            fg_color=SURFACE_BG, border_color=THEME_COLOR,
            text_color=TEXT_MAIN,
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(row1, text="Status:",
                     font=("Segoe UI", 11), text_color=TEXT_MAIN
                     ).pack(side="left", padx=(0, 6))
        self._status_var = tk.StringVar(value="All")
        ctk.CTkOptionMenu(
            row1, variable=self._status_var,
            values=["All", "In Stock", "Low Stock", "Out of Stock"],
            font=("Segoe UI", 11), width=130, height=30,
            fg_color=SURFACE_BG, button_color=THEME_COLOR,
            button_hover_color=ACCENT_2, text_color=TEXT_MAIN,
            command=lambda _v: self._apply_filter(),
        ).pack(side="left", padx=(0, 16))

        self._empty_sku_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            row1, text="Empty SKU only",
            variable=self._empty_sku_var,
            command=self._apply_filter,
            font=("Segoe UI", 11), text_color=TEXT_MAIN,
            fg_color=THEME_COLOR, hover_color=ACCENT_2,
            border_color=TEXT_MUTED,
        ).pack(side="left")

        # Row 2 — action buttons
        row2 = ctk.CTkFrame(bar, fg_color="transparent")
        row2.pack(fill="x", padx=16, pady=(8, 12))

        ctk.CTkButton(
            row2, text="➕  Add Item", width=120, height=32,
            font=("Segoe UI", 11), fg_color=THEME_COLOR,
            hover_color=ACCENT_2, command=self._on_add,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            row2, text="✏  Edit Selected", width=140, height=32,
            font=("Segoe UI", 11), fg_color=SURFACE_BG,
            border_width=1, border_color=THEME_COLOR,
            hover_color="#1A3A2A", command=self._on_edit,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            row2, text="🗑  Delete Selected", width=150, height=32,
            font=("Segoe UI", 11), fg_color=SURFACE_BG,
            border_width=1, border_color=ERROR_CLR,
            hover_color="#3A1A1A", text_color=ERROR_CLR,
            command=self._on_delete,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            row2, text="🔄  Refresh", width=110, height=32,
            font=("Segoe UI", 11), fg_color=SURFACE_BG,
            border_width=1, border_color=THEME_COLOR,
            hover_color="#1A3A2A", command=self.refresh,
        ).pack(side="left")

        ctk.CTkButton(
            row2, text="Close", width=90, height=32,
            font=("Segoe UI", 11), fg_color=SURFACE_BG,
            border_width=1, border_color="#555",
            hover_color="#3A3A3A", command=self.destroy,
        ).pack(side="right")

    # ── tree ───────────────────────────────────────────────────────────

    def _build_tree(self):
        self._configure_tree_style()

        wrap = tk.Frame(self, bg=DARK_BG, highlightthickness=0)
        wrap.pack(fill="both", expand=True, padx=16, pady=(12, 8))

        cols_ids = [c[4] for c in self._COLS]
        self.tree = ttk.Treeview(
            wrap, columns=cols_ids, show="headings",
            style="Inv.Treeview")
        for key, header, anchor, width, col_id, _ in self._COLS:
            # Click header → sort by that key
            self.tree.heading(
                col_id, text=header,
                command=lambda k=key: self._sort_by(k))
            self.tree.column(col_id, anchor=anchor, width=width,
                             stretch=(col_id == "item_name"))

        vsb = ttk.Scrollbar(wrap, orient="vertical",
                            command=self.tree.yview,
                            style="Inv.Vertical.TScrollbar")
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Alternating row backgrounds
        self.tree.tag_configure("odd",  background=SURFACE_BG)
        self.tree.tag_configure("even", background=CARD_BG)
        # Status tags
        self.tree.tag_configure("status_out",  foreground=ERROR_CLR)
        self.tree.tag_configure("status_low",  foreground=WARN_CLR)
        self.tree.tag_configure("status_in",   foreground=OK_CLR)

        # Double-click row → edit
        self.tree.bind("<Double-1>", lambda _e: self._on_edit())

    def _configure_tree_style(self):
        style = ttk.Style(self)
        try:    style.theme_use("clam")
        except tk.TclError: pass
        style.configure("Inv.Treeview",
                        background=SURFACE_BG, foreground=TEXT_MAIN,
                        fieldbackground=SURFACE_BG,
                        bordercolor=DARK_BG, borderwidth=0,
                        rowheight=26, font=("Segoe UI", 10))
        style.configure("Inv.Treeview.Heading",
                        background=THEME_COLOR, foreground="white",
                        relief="flat", font=("Segoe UI", 10, "bold"))
        style.map("Inv.Treeview",
                  background=[("selected", ACCENT_2)],
                  foreground=[("selected", "white")])
        style.map("Inv.Treeview.Heading",
                  background=[("active", ACCENT_2)])
        style.configure("Inv.Vertical.TScrollbar",
                        background=SURFACE_BG, troughcolor=DARK_BG,
                        bordercolor=DARK_BG, arrowcolor=TEXT_MUTED)

    # ── footer ─────────────────────────────────────────────────────────

    def _build_footer(self):
        foot = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0, height=36)
        foot.pack(fill="x", side="bottom")
        foot.pack_propagate(False)

        self._status_lbl = ctk.CTkLabel(
            foot, text="Loading…", font=("Segoe UI", 11),
            text_color=TEXT_MUTED, anchor="w")
        self._status_lbl.pack(side="left", padx=16, pady=8)

        ctk.CTkLabel(
            foot, text=f"Threshold: qty ≤ {LOW_STOCK_THRESHOLD} = Low Stock",
            font=("Segoe UI", 10), text_color=TEXT_MUTED,
        ).pack(side="right", padx=16, pady=8)

    # ── data loading & rendering ───────────────────────────────────────

    def refresh(self):
        """Pull a fresh snapshot from the DB (off-thread) and re-render."""
        db = self._get_db_path() if self._get_db_path else ""
        if not db or not Path(db).exists():
            self._snapshot = []
            self._render()
            self._status_lbl.configure(
                text="No master database selected.",
                text_color=WARN_CLR)
            return

        self._status_lbl.configure(text="Loading…", text_color=TEXT_MUTED)

        def _work():
            try:
                snap = self.dm.get_inventory_snapshot(db)
                self.after(0, lambda s=snap: self._on_loaded(s))
            except Exception as e:
                self.after(0, lambda err=str(e): self._on_load_error(err))

        threading.Thread(target=_work, daemon=True).start()

    def _on_loaded(self, snap: list[dict]):
        self._snapshot = snap
        self._apply_filter()

    def _on_load_error(self, msg: str):
        self._snapshot = []
        self._render()
        self._status_lbl.configure(
            text=f"Could not load data: {msg}", text_color=ERROR_CLR)

    def _apply_filter(self):
        q = self._search_var.get().strip().lower()
        status_pick = self._status_var.get()
        only_empty  = self._empty_sku_var.get()

        rows = self._snapshot
        if q:
            rows = [r for r in rows
                    if q in r["item_name"].lower() or q in r["sku_code"].lower()]
        if status_pick != "All":
            rows = [r for r in rows
                    if _format_status(r["available"])[0] == status_pick]
        if only_empty:
            rows = [r for r in rows if not r["sku_code"].strip()]

        self._filtered = rows
        self._sort_in_place()
        self._render()

    def _sort_by(self, key: str):
        if self._sort_key == key:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_key = key
            self._sort_reverse = False
        self._sort_in_place()
        self._render(preserve_selection=False)

    def _sort_in_place(self):
        col_meta = next((c for c in self._COLS if c[0] == self._sort_key), None)
        numeric  = col_meta[5] if col_meta else False

        def key_fn(row):
            v = row.get(self._sort_key)
            if numeric:
                try:    return float(v or 0)
                except (TypeError, ValueError): return 0.0
            return str(v or "").lower()

        self._filtered.sort(key=key_fn, reverse=self._sort_reverse)

    def _render(self, preserve_selection: bool = True):
        prev_sel_ids = set()
        if preserve_selection:
            for iid in self.tree.selection():
                vals = self.tree.item(iid)["values"]
                if vals:
                    prev_sel_ids.add(int(vals[0]))

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for idx, row in enumerate(self._filtered):
            status_lbl, _ = _format_status(row["available"])
            stripe = "even" if idx % 2 == 0 else "odd"
            status_tag = {
                "Out of Stock": "status_out",
                "Low Stock":    "status_low",
                "In Stock":     "status_in",
            }.get(status_lbl, "")

            values = [
                row["id"],
                row["item_name"],
                row["sku_code"] or "—",
                row["hsn_sac"]  or "—",
                row["gst_rate"] or "—",
                f"{row['rate']:.2f}"        if row["rate"]        else "—",
                row["measurement"] or "—",
                f"{row['total_received']:.0f}",
                f"{row['total_consumed']:.0f}",
                f"{row['adjustments']:+.0f}" if row["adjustments"] else "0",
                f"{row['available']:.0f}",
                status_lbl,
            ]
            tags = (stripe,) + ((status_tag,) if status_tag else ())
            iid = self.tree.insert("", "end", values=values, tags=tags)
            if int(row["id"]) in prev_sel_ids:
                self.tree.selection_add(iid)

        total = len(self._snapshot)
        showing = len(self._filtered)
        if total == 0:
            self._status_lbl.configure(
                text="No items in this database yet.", text_color=TEXT_MUTED)
        elif showing == total:
            self._status_lbl.configure(
                text=f"Showing all {total} item(s).", text_color=TEXT_MUTED)
        else:
            self._status_lbl.configure(
                text=f"Showing {showing} of {total} item(s) (filtered).",
                text_color=TEXT_MUTED)

    # ── selection helpers ──────────────────────────────────────────────

    def _selected_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            return None
        vals = self.tree.item(sel[0])["values"]
        if not vals:
            return None
        try:    return int(vals[0])
        except (TypeError, ValueError): return None

    def _selected_row(self) -> Optional[dict]:
        item_id = self._selected_id()
        if item_id is None:
            return None
        return next((r for r in self._snapshot if r["id"] == item_id), None)

    # ── action handlers ────────────────────────────────────────────────

    def _on_add(self):
        db = self._get_db_path()
        if not db or not Path(db).exists():
            messagebox.showwarning(
                "No Master", "Select a master database first.", parent=self)
            return

        def save(fields, delta, note):  # delta/note are 0/'' in add mode
            try:
                new_id = self.dm.add_item(db, **fields)
                self._log(f"Added item id={new_id}: {fields['item_name']!r}",
                          "success")
                self.refresh()
                return True
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=self)
                return False

        ItemDialog(self, mode="add", item_data=None, on_save=save)

    def _on_edit(self):
        row = self._selected_row()
        if not row:
            messagebox.showinfo(
                "No selection", "Select a row to edit (or double-click it).",
                parent=self)
            return
        db = self._get_db_path()

        def save(fields, delta, note):
            try:
                self.dm.update_item(db, row["id"], **fields)
                if abs(delta) > 1e-9:
                    self.dm.add_manual_adjustment(db, row["id"], delta, note)
                self._log(
                    f"Updated item id={row['id']}"
                    + (f" (adjustment {delta:+g})" if abs(delta) > 1e-9 else ""),
                    "success")
                self.refresh()
                return True
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=self)
                return False

        ItemDialog(self, mode="edit", item_data=row, on_save=save)

    def _on_delete(self):
        row = self._selected_row()
        if not row:
            messagebox.showinfo(
                "No selection", "Select a row to delete.", parent=self)
            return
        if not messagebox.askyesno(
                "Delete item?",
                f"Permanently delete this item and ALL its receipts, "
                f"consumption history, and adjustments?\n\n"
                f"  ID:   {row['id']}\n"
                f"  Name: {row['item_name']}\n"
                f"  SKU:  {row['sku_code'] or '(none)'}\n"
                f"  Available: {row['available']:.0f}\n\n"
                f"This cannot be undone.",
                parent=self):
            return
        db = self._get_db_path()
        try:
            self.dm.delete_item(db, row["id"])
            self._log(f"Deleted item id={row['id']}: "
                      f"{row['item_name']!r}", "warning")
            self.refresh()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
