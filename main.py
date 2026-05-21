"""
Inventory Manager — Main Application Entry Point
A local Windows desktop application for automated inventory management.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import json
from pathlib import Path
from datetime import datetime

# Ensure core modules are importable when running as .exe
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent

sys.path.insert(0, str(BASE_DIR))

from core.invoice_parser import InvoiceParser
from core.database_manager import DatabaseManager
from core.consumption_processor import ConsumptionProcessor

# ─── App Configuration ────────────────────────────────────────────────────────
APP_TITLE   = "Inventory Manager Pro"
APP_VERSION = "1.0.0"
THEME_COLOR = "#1A6B3C"       # Forest green – primary accent
ACCENT_2    = "#2E9E5E"       # Lighter green for hover
DARK_BG     = "#121212"
CARD_BG     = "#1E1E1E"
SURFACE_BG  = "#252525"
TEXT_MAIN   = "#F0F0F0"
TEXT_MUTED  = "#888888"
SUCCESS_CLR = "#2ECC71"
WARNING_CLR = "#F39C12"
ERROR_CLR   = "#E74C3C"

CONFIG_FILE = Path.home() / ".inventory_manager_config.json"


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text())
        except Exception:
            pass
    return {"master_path": "", "output_dir": ""}


def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))


# ─── Reusable UI Components ───────────────────────────────────────────────────

class StatusLog(ctk.CTkFrame):
    """Scrollable log console at the bottom of the window."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, corner_radius=8, fg_color=DARK_BG, **kwargs)
        self._build()

    def _build(self):
        header = ctk.CTkLabel(self, text="  Activity Log", font=("Consolas", 11),
                              text_color=TEXT_MUTED, anchor="w")
        header.pack(fill="x", padx=8, pady=(6, 0))

        self.textbox = ctk.CTkTextbox(self, font=("Consolas", 11),
                                      fg_color=DARK_BG, text_color="#CCCCCC",
                                      wrap="word", state="disabled")
        self.textbox.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def log(self, message: str, level: str = "info"):
        colours = {"info": TEXT_MAIN, "success": SUCCESS_CLR,
                   "warning": WARNING_CLR, "error": ERROR_CLR}
        prefix  = {"info": "ℹ ", "success": "✔ ", "warning": "⚠ ", "error": "✖ "}
        ts  = datetime.now().strftime("%H:%M:%S")
        tag = level
        line = f"[{ts}]  {prefix.get(level, '')} {message}\n"

        self.textbox.configure(state="normal")
        self.textbox.tag_config(tag, foreground=colours.get(level, TEXT_MAIN))
        self.textbox.insert("end", line, tag)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def clear(self):
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")


class FilePickerRow(ctk.CTkFrame):
    """A labelled filepath entry + Browse button row."""

    def __init__(self, parent, label: str, placeholder: str = "No file selected",
                 filetypes=None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._filetypes = filetypes or [("All Files", "*.*")]
        self._var = tk.StringVar(value="")
        self._build(label, placeholder)

    def _build(self, label: str, placeholder: str):
        ctk.CTkLabel(self, text=label, font=("Segoe UI", 12, "bold"),
                     text_color=TEXT_MAIN, anchor="w", width=180).pack(side="left")

        self.entry = ctk.CTkEntry(self, textvariable=self._var,
                                  placeholder_text=placeholder,
                                  font=("Segoe UI", 11), height=34,
                                  fg_color=SURFACE_BG, border_color="#444",
                                  text_color=TEXT_MAIN)
        self.entry.pack(side="left", fill="x", expand=True, padx=(8, 6))

        self.btn = ctk.CTkButton(self, text="Browse", width=80, height=34,
                                 fg_color=THEME_COLOR, hover_color=ACCENT_2,
                                 font=("Segoe UI", 11),
                                 command=self._browse)
        self.btn.pack(side="left")

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=self._filetypes)
        if path:
            self._var.set(path)

    def get(self) -> str:
        return self._var.get().strip()

    def set(self, value: str):
        self._var.set(value)


class DirPickerRow(ctk.CTkFrame):
    """A labelled directory entry + Browse button row."""

    def __init__(self, parent, label: str, placeholder: str = "Choose folder…", **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._var = tk.StringVar(value="")
        self._build(label, placeholder)

    def _build(self, label: str, placeholder: str):
        ctk.CTkLabel(self, text=label, font=("Segoe UI", 12, "bold"),
                     text_color=TEXT_MAIN, anchor="w", width=180).pack(side="left")

        self.entry = ctk.CTkEntry(self, textvariable=self._var,
                                  placeholder_text=placeholder,
                                  font=("Segoe UI", 11), height=34,
                                  fg_color=SURFACE_BG, border_color="#444",
                                  text_color=TEXT_MAIN)
        self.entry.pack(side="left", fill="x", expand=True, padx=(8, 6))

        self.btn = ctk.CTkButton(self, text="Browse", width=80, height=34,
                                 fg_color=THEME_COLOR, hover_color=ACCENT_2,
                                 font=("Segoe UI", 11), command=self._browse)
        self.btn.pack(side="left")

    def _browse(self):
        path = filedialog.askdirectory()
        if path:
            self._var.set(path)

    def get(self) -> str:
        return self._var.get().strip()

    def set(self, value: str):
        self._var.set(value)


class SectionCard(ctk.CTkFrame):
    """Raised card container with a title bar."""

    def __init__(self, parent, title: str, icon: str = "", **kwargs):
        super().__init__(parent, corner_radius=10, fg_color=CARD_BG,
                         border_width=1, border_color="#333", **kwargs)
        self._build(title, icon)

    def _build(self, title: str, icon: str):
        title_bar = ctk.CTkFrame(self, fg_color=SURFACE_BG, corner_radius=0,
                                 height=40)
        title_bar.pack(fill="x", padx=0, pady=0)
        title_bar.pack_propagate(False)

        ctk.CTkLabel(title_bar, text=f"  {icon}  {title}",
                     font=("Segoe UI", 13, "bold"), text_color=TEXT_MAIN,
                     anchor="w").pack(side="left", fill="y", padx=8)

        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.pack(fill="both", expand=True, padx=16, pady=12)


class ActionButton(ctk.CTkButton):
    """Styled primary action button with optional spinner state."""

    def __init__(self, parent, text: str, command=None, icon: str = "", **kwargs):
        self._base_text = f"{icon}  {text}" if icon else text
        super().__init__(parent, text=self._base_text,
                         font=("Segoe UI", 13, "bold"),
                         height=44, corner_radius=8,
                         fg_color=THEME_COLOR, hover_color=ACCENT_2,
                         text_color="white", command=command, **kwargs)

    def set_busy(self, busy: bool):
        if busy:
            self.configure(text="⏳  Processing…", state="disabled",
                           fg_color="#555")
        else:
            self.configure(text=self._base_text, state="normal",
                           fg_color=THEME_COLOR)


# ─── Main Application Window ──────────────────────────────────────────────────

class InventoryApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.geometry("1050x800")
        self.minsize(900, 680)

        self.cfg = load_config()
        self._setup_services()
        self._build_ui()
        self._restore_paths()

    # ── Service Initialisation ─────────────────────────────────────────────

    def _setup_services(self):
        self.invoice_parser = InvoiceParser()
        self.excel_manager  = DatabaseManager()   # attribute kept for compatibility
        self.cons_processor = ConsumptionProcessor()

    # ── UI Construction ────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar ──────────────────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color=SURFACE_BG, height=56, corner_radius=0)
        top.pack(fill="x", side="top")
        top.pack_propagate(False)

        ctk.CTkLabel(top,
                     text="📦  Inventory Manager Pro",
                     font=("Segoe UI", 16, "bold"),
                     text_color=TEXT_MAIN).pack(side="left", padx=20, pady=10)

        ctk.CTkLabel(top, text=f"v{APP_VERSION}",
                     font=("Segoe UI", 10), text_color=TEXT_MUTED).pack(
            side="left", pady=10)

        # ── Status log (bottom) ───────────────────────────────────────────
        self.log = StatusLog(self, height=190)
        self.log.pack(fill="x", side="bottom", padx=10, pady=(0, 10))

        # ── Main scrollable content area ──────────────────────────────────
        scroll_container = ctk.CTkScrollableFrame(self, fg_color=DARK_BG,
                                                  scrollbar_button_color="#333",
                                                  scrollbar_button_hover_color="#555")
        scroll_container.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_config_section(scroll_container)
        self._build_invoice_section(scroll_container)
        self._build_consumption_section(scroll_container)
        self._build_stock_section(scroll_container)

    # ── Section: Configuration ─────────────────────────────────────────────

    def _build_config_section(self, parent):
        card = SectionCard(parent, "Configuration", "⚙")
        card.pack(fill="x", pady=(0, 12))
        c = card.content

        self.master_picker = FilePickerRow(
            c, "Master Database:",
            placeholder="Select your Master Inventory .db file",
            filetypes=[("Inventory Database",   "*.db"),
                       ("Legacy Excel Master",  "*.xlsx *.xls"),
                       ("All Files",            "*.*")])
        self.master_picker.pack(fill="x", pady=(0, 8))

        self.output_dir_picker = DirPickerRow(
            c, "Output Folder:",
            placeholder="Folder where reports will be saved")
        self.output_dir_picker.pack(fill="x", pady=(0, 4))

        row = ctk.CTkFrame(c, fg_color="transparent")
        row.pack(fill="x", pady=(8, 0))

        ctk.CTkButton(row, text="💾  Save Paths", width=130, height=32,
                      font=("Segoe UI", 11), fg_color=SURFACE_BG,
                      border_width=1, border_color=THEME_COLOR,
                      hover_color="#1A3A2A",
                      command=self._save_config).pack(side="left")

        ctk.CTkButton(row, text="📂  Create New Master",
                      width=160, height=32,
                      font=("Segoe UI", 11), fg_color=SURFACE_BG,
                      border_width=1, border_color=THEME_COLOR,
                      hover_color="#1A3A2A",
                      command=self._create_new_master).pack(side="left", padx=(8, 0))

        ctk.CTkButton(row, text="🔄  Migrate from Excel",
                      width=170, height=32,
                      font=("Segoe UI", 11), fg_color=SURFACE_BG,
                      border_width=1, border_color=THEME_COLOR,
                      hover_color="#1A3A2A",
                      command=self._migrate_from_excel).pack(side="left", padx=(8, 0))

    # ── Section: Invoice Upload ─────────────────────────────────────────────

    def _build_invoice_section(self, parent):
        card = SectionCard(parent, "Step 1 — Upload & Process Invoice", "📄")
        card.pack(fill="x", pady=(0, 12))
        c = card.content

        info = ctk.CTkLabel(
            c,
            text="Accepts PDF or image invoices. Extracts line items and updates the Master Database with date-wise stock.",
            font=("Segoe UI", 11), text_color=TEXT_MUTED, wraplength=800, justify="left")
        info.pack(anchor="w", pady=(0, 10))

        self.invoice_picker = FilePickerRow(
            c, "Invoice File:",
            placeholder="Select invoice PDF or image (PNG, JPG, TIFF)…",
            filetypes=[
                ("Supported Files", "*.pdf *.png *.jpg *.jpeg *.tiff *.bmp *.webp"),
                ("PDF Files", "*.pdf"),
                ("Images", "*.png *.jpg *.jpeg *.tiff *.bmp"),
                ("All Files", "*.*"),
            ])
        self.invoice_picker.pack(fill="x", pady=(0, 10))

        row = ctk.CTkFrame(c, fg_color="transparent")
        row.pack(fill="x")

        self.inv_date_label = ctk.CTkLabel(row, text="Invoice Date (auto-detected):",
                                           font=("Segoe UI", 11), text_color=TEXT_MUTED,
                                           width=200, anchor="w")
        self.inv_date_label.pack(side="left")

        self.inv_date_var = tk.StringVar(value="")
        self.inv_date_entry = ctk.CTkEntry(row, textvariable=self.inv_date_var,
                                           placeholder_text="DD/MM/YYYY (override if needed)",
                                           font=("Segoe UI", 11), height=34, width=200,
                                           fg_color=SURFACE_BG, border_color="#444",
                                           text_color=TEXT_MAIN)
        self.inv_date_entry.pack(side="left", padx=(8, 0))

        self.upload_btn = ActionButton(c, "Process Invoice & Update Master",
                                      icon="🚀", command=self._run_invoice)
        self.upload_btn.pack(pady=(14, 0))

    # ── Section: Consumption Report ──────────────────────────────────────────

    def _build_consumption_section(self, parent):
        card = SectionCard(parent, "Step 2 — Upload Consumption Report", "📊")
        card.pack(fill="x", pady=(0, 12))
        c = card.content

        info = ctk.CTkLabel(
            c,
            text="Upload a PDF or Excel file listing items consumed. "
                 "The system will match items against the Master and track deductions.",
            font=("Segoe UI", 11), text_color=TEXT_MUTED, wraplength=800, justify="left")
        info.pack(anchor="w", pady=(0, 10))

        self.cons_picker = FilePickerRow(
            c, "Consumption File:",
            placeholder="Select consumption report (PDF or .xlsx)…",
            filetypes=[
                ("Supported Files", "*.pdf *.xlsx *.xls *.csv"),
                ("PDF Files",       "*.pdf"),
                ("Excel Files",     "*.xlsx *.xls"),
                ("CSV Files",       "*.csv"),
                ("All Files",       "*.*"),
            ])
        self.cons_picker.pack(fill="x", pady=(0, 10))

        self.cons_btn = ActionButton(c, "Upload Consumption Report",
                                     icon="📥", command=self._run_consumption)
        self.cons_btn.pack(pady=(4, 0))

    # ── Section: Generate Stock Report ──────────────────────────────────────

    def _build_stock_section(self, parent):
        card = SectionCard(parent, "Step 3 — Generate Current Stock Report", "📈")
        card.pack(fill="x", pady=(0, 4))
        c = card.content

        info = ctk.CTkLabel(
            c,
            text="Computes available stock = Σ(all date purchases) − total consumed. "
                 "Saves a final 'Current Stock' Excel report to the output folder.",
            font=("Segoe UI", 11), text_color=TEXT_MUTED, wraplength=800, justify="left")
        info.pack(anchor="w", pady=(0, 10))

        self.stock_btn = ActionButton(c, "Generate Current Stock Report",
                                      icon="📋", command=self._run_stock_report)
        self.stock_btn.pack(pady=(4, 0))

        # Quick-stats bar
        stats_bar = ctk.CTkFrame(c, fg_color=SURFACE_BG, corner_radius=6)
        stats_bar.pack(fill="x", pady=(16, 0))

        self._stat_items    = self._stat_label(stats_bar, "Total SKUs", "—")
        self._stat_invoices = self._stat_label(stats_bar, "Invoice Dates", "—")
        self._stat_low      = self._stat_label(stats_bar, "Low-Stock Items", "—")
        self._stat_output   = self._stat_label(stats_bar, "Last Report", "—")

    def _stat_label(self, parent, title: str, value: str):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=20, pady=8, expand=True)
        ctk.CTkLabel(frame, text=title, font=("Segoe UI", 10),
                     text_color=TEXT_MUTED).pack()
        val_lbl = ctk.CTkLabel(frame, text=value, font=("Segoe UI", 18, "bold"),
                                text_color=THEME_COLOR)
        val_lbl.pack()
        return val_lbl

    # ── Path Restore & Save ────────────────────────────────────────────────

    def _restore_paths(self):
        if self.cfg.get("master_path"):
            self.master_picker.set(self.cfg["master_path"])
        if self.cfg.get("output_dir"):
            self.output_dir_picker.set(self.cfg["output_dir"])

    def _save_config(self):
        self.cfg["master_path"] = self.master_picker.get()
        self.cfg["output_dir"]  = self.output_dir_picker.get()
        save_config(self.cfg)
        self.log.log("Configuration saved.", "success")

    # ── Create New Master ──────────────────────────────────────────────────

    def _create_new_master(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Inventory Database", "*.db")],
            title="Create New Inventory Database",
            initialfile="Master_Inventory.db")
        if not path:
            return
        try:
            final = self.excel_manager.create_empty_master(path)
            self.master_picker.set(final)
            self.log.log(f"New database created: {final}", "success")
        except Exception as e:
            self.log.log(f"Failed to create database: {e}", "error")

    # ── Migrate Excel → SQLite ─────────────────────────────────────────────

    def _migrate_from_excel(self):
        xlsx = filedialog.askopenfilename(
            title="Select existing Master_Inventory.xlsx",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if not xlsx:
            return
        db_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Inventory Database", "*.db")],
            title="Save migrated database as…",
            initialfile="Master_Inventory.db")
        if not db_path:
            return

        self.log.log(f"Migrating {Path(xlsx).name} → SQLite…", "info")

        def _work():
            try:
                stats = self.excel_manager.migrate_from_excel(xlsx, db_path)
                self.after(0, lambda: (
                    self.master_picker.set(db_path),
                    self.log.log(
                        f"Migration complete ✔  |  {stats['items']} items  |  "
                        f"{stats['receipts']} receipt rows  |  "
                        f"{stats['consumption']} consumption rows",
                        "success")))
            except Exception as e:
                self.after(0, lambda err=str(e):
                           self.log.log(f"Migration error: {err}", "error"))

        threading.Thread(target=_work, daemon=True).start()

    # ── Workflow: Invoice ──────────────────────────────────────────────────

    def _run_invoice(self):
        inv_path    = self.invoice_picker.get()
        master_path = self.master_picker.get()
        inv_date    = self.inv_date_var.get().strip()

        if not inv_path:
            messagebox.showwarning("Missing File", "Please select an invoice file.")
            return
        if not master_path or not Path(master_path).exists():
            messagebox.showwarning("Missing Master", "Please select a valid Master Database file.")
            return

        self.upload_btn.set_busy(True)
        self.log.log(f"Processing invoice: {Path(inv_path).name}", "info")

        def _work():
            try:
                items, detected_date = self.invoice_parser.parse(inv_path)
                final_date = inv_date or detected_date

                if not final_date:
                    self.after(0, lambda: messagebox.showwarning(
                        "Date Required",
                        "Could not auto-detect invoice date.\n"
                        "Please enter it manually in the date field (DD/MM/YYYY)."))
                    return

                # Update date field in UI
                self.after(0, lambda d=final_date: self.inv_date_var.set(d))

                stats = self.excel_manager.update_with_invoice(master_path, items, final_date)
                msg = (f"Invoice processed ✔  |  "
                       f"{stats['new_rows']} new items  |  "
                       f"{stats['updated_rows']} updated  |  "
                       f"Date column: {final_date}")
                self.after(0, lambda m=msg: self.log.log(m, "success"))

            except Exception as e:
                self.after(0, lambda err=str(e): self.log.log(f"Invoice error: {err}", "error"))
            finally:
                self.after(0, lambda: self.upload_btn.set_busy(False))

        threading.Thread(target=_work, daemon=True).start()

    # ── Workflow: Consumption ──────────────────────────────────────────────

    def _run_consumption(self):
        cons_path   = self.cons_picker.get()
        master_path = self.master_picker.get()

        if not cons_path:
            messagebox.showwarning("Missing File", "Please select a consumption report file.")
            return
        if not master_path or not Path(master_path).exists():
            messagebox.showwarning("Missing Master", "Please select a valid Master Database file.")
            return

        self.cons_btn.set_busy(True)
        self.log.log(f"Processing consumption report: {Path(cons_path).name}", "info")

        def _work():
            try:
                consumed_items = self.cons_processor.parse(cons_path)
                matched, unmatched = self.excel_manager.record_consumption(
                    master_path, consumed_items)

                self.after(0, lambda: self.log.log(
                    f"Consumption recorded ✔  |  {matched} items matched  |  "
                    f"{unmatched} unmatched (check log)", "success"))

                if unmatched > 0:
                    self.after(0, lambda: self.log.log(
                        f"{unmatched} item(s) from the consumption report "
                        "could not be matched – verify names/SKUs in the master.", "warning"))

            except Exception as e:
                self.after(0, lambda err=str(e): self.log.log(
                    f"Consumption error: {err}", "error"))
            finally:
                self.after(0, lambda: self.cons_btn.set_busy(False))

        threading.Thread(target=_work, daemon=True).start()

    # ── Workflow: Stock Report ─────────────────────────────────────────────

    def _run_stock_report(self):
        master_path = self.master_picker.get()
        output_dir  = self.output_dir_picker.get()

        if not master_path or not Path(master_path).exists():
            messagebox.showwarning("Missing Master", "Please select a valid Master Database file.")
            return
        if not output_dir:
            messagebox.showwarning("Missing Output", "Please choose an output folder.")
            return

        self.stock_btn.set_busy(True)
        self.log.log("Generating Current Stock report…", "info")

        def _work():
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_path = Path(output_dir) / f"Current_Stock_{ts}.xlsx"
                stats = self.excel_manager.generate_stock_report(
                    master_path, str(out_path))

                self.after(0, lambda: (
                    self._stat_items.configure(text=str(stats["total_skus"])),
                    self._stat_invoices.configure(text=str(stats["date_columns"])),
                    self._stat_low.configure(
                        text=str(stats["low_stock"]),
                        text_color=WARNING_CLR if stats["low_stock"] else SUCCESS_CLR),
                    self._stat_output.configure(
                        text=datetime.now().strftime("%H:%M")),
                    self.log.log(
                        f"Report saved: {out_path.name}  |  "
                        f"{stats['total_skus']} SKUs  |  "
                        f"{stats['low_stock']} low-stock items", "success")
                ))

                if messagebox.askyesno("Report Ready",
                                       f"Stock report saved!\n\n{out_path}\n\nOpen now?"):
                    os.startfile(str(out_path))

            except Exception as e:
                self.after(0, lambda err=str(e): self.log.log(
                    f"Report error: {err}", "error"))
            finally:
                self.after(0, lambda: self.stock_btn.set_busy(False))

        threading.Thread(target=_work, daemon=True).start()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()
