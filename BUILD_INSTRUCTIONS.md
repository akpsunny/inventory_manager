# 📦 Inventory Manager Pro — Setup & Build Guide

---

## Table of Contents
1. [Project Structure](#1-project-structure)
2. [Prerequisites](#2-prerequisites)
3. [Installation](#3-installation)
4. [Tesseract OCR Setup (for image invoices)](#4-tesseract-ocr-setup)
5. [Running the App](#5-running-the-app)
6. [Using the Application](#6-using-the-application)
7. [Building a Standalone .exe](#7-building-a-standalone-exe)
8. [Troubleshooting](#8-troubleshooting)
9. [Data Schema Reference](#9-data-schema-reference)
10. [Edge Cases & Tips](#10-edge-cases--tips)

---

## 1. Project Structure

```
inventory_manager/
│
├── main.py                      ← Entry point (run this)
├── requirements.txt             ← Python dependencies
├── BUILD_INSTRUCTIONS.md        ← This file
│
└── core/
    ├── __init__.py
    ├── invoice_parser.py        ← PDF / image invoice parser
    ├── database_manager.py      ← SQLite storage + stock report writer
    ├── excel_manager.py         ← Legacy Excel storage (kept as fallback)
    ├── consumption_processor.py ← Consumption report parser
    └── fuzzy_matcher.py         ← Fuzzy string matching utility
```

---

## 2. Prerequisites

| Requirement | Version | Download |
|---|---|---|
| Python | 3.10, 3.11, or 3.12 | https://www.python.org/downloads/ |
| pip | (bundled with Python) | — |
| Tesseract OCR | 5.x | https://github.com/UB-Mannheim/tesseract/wiki |

> ⚠️ **Python 3.13 is NOT recommended** — some dependencies (openpyxl, PyMuPDF)
> may not yet have stable wheels for it.

---

## 3. Installation

Open **Command Prompt** or **PowerShell** in the project folder and run:

```bat
:: Step 1 — Create a virtual environment (strongly recommended)
python -m venv venv

:: Step 2 — Activate it
venv\Scripts\activate

:: Step 3 — Upgrade pip
python -m pip install --upgrade pip

:: Step 4 — Install all dependencies
pip install -r requirements.txt
```

---

## 4. Tesseract OCR Setup

Tesseract is only needed for **image-based invoices** (PNG, JPG, TIFF).
PDF invoices do not require it.

### Install Steps

1. Download the installer from:
   **https://github.com/UB-Mannheim/tesseract/wiki**
   (Choose the Windows 64-bit installer, e.g. `tesseract-ocr-w64-setup-5.x.x.exe`)

2. Run the installer. Accept defaults. The default install path is:
   ```
   C:\Program Files\Tesseract-OCR\
   ```

3. Add Tesseract to your Windows PATH:
   - Open **Start → "Edit the system environment variables"**
   - Click **Environment Variables**
   - Under **System Variables**, select **Path** → **Edit**
   - Click **New** and add:
     ```
     C:\Program Files\Tesseract-OCR
     ```
   - Click **OK** on all dialogs

4. Verify the installation by opening a new Command Prompt and running:
   ```bat
   tesseract --version
   ```
   You should see `tesseract 5.x.x` printed.

### If You Skip the PATH Step

If you prefer not to modify the PATH, open `core/invoice_parser.py` and
add this line near the top (after the imports):

```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

---

## 5. Running the App

With the virtual environment active:

```bat
python main.py
```

The GUI window will open. No browser, no server — fully local.

---

## 6. Using the Application

### First-Time Setup
1. Click **Browse** next to **"Master Database"** and either:
   - Select an existing `Master_Inventory.db`, OR
   - Click **"📂 Create New Master"** to generate a blank `.db` with the SQLite schema, OR
   - Click **"🔄 Migrate from Excel"** to import an existing `Master_Inventory.xlsx`
     (all items + date receipts + the running Consumed Qty) into a fresh `.db`.
2. Click **Browse** next to **"Output Folder"** and choose where reports will be saved.
3. Click **💾 Save Paths** — these paths are remembered across sessions.

### Workflow: Processing an Invoice (Step 1)
1. Click **Browse** next to **"Invoice File"** and select your PDF or image invoice.
2. The system will auto-detect the invoice date. If it cannot, type it manually in `DD/MM/YYYY` format.
3. Click **🚀 Process Invoice & Update Master**.
4. Watch the Activity Log for confirmation. The master `.db` is updated instantly (atomic SQLite transaction).

### Workflow: Uploading a Consumption Report (Step 2)
1. Click **Browse** next to **"Consumption File"** and select a PDF, Excel, or CSV file.
2. Click **📥 Upload Consumption Report**.
3. The system fuzzy-matches item names/SKUs against the master and fills the **Consumed Qty** column.

### Workflow: Generating the Stock Report (Step 3)
1. Click **📋 Generate Current Stock Report**.
2. A timestamped `Current_Stock_YYYYMMDD_HHMMSS.xlsx` file is saved to your output folder.
3. A dialog appears asking if you want to open the file immediately.

---

## 7. Building a Standalone .exe

This converts the Python app into a double-clickable `.exe` that runs on any
Windows machine — **no Python installation required on the target machine**.

### Step 1 — Install PyInstaller

```bat
pip install pyinstaller==6.8.0
```

### Step 2 — Build the executable

From the `inventory_manager/` folder, run:

```bat
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "InventoryManagerPro" ^
  --add-data "core;core" ^
  --hidden-import "pdfplumber" ^
  --hidden-import "fitz" ^
  --hidden-import "pytesseract" ^
  --hidden-import "openpyxl" ^
  --hidden-import "rapidfuzz" ^
  --hidden-import "dateutil" ^
  --hidden-import "pandas" ^
  --hidden-import "customtkinter" ^
  main.py
```

> **Tip:** On PowerShell, replace the `^` line-continuation characters with `` ` ``
>
> **SQLite note:** `sqlite3` is part of Python's standard library, so it's
> bundled into the .exe automatically — no `--hidden-import` needed.

### Step 3 — Find your .exe

After the build completes (takes 1–3 minutes), the executable is at:
```
inventory_manager\dist\InventoryManagerPro.exe
```

### Step 4 — Deploy

Copy **`InventoryManagerPro.exe`** to any Windows 10/11 machine and double-click it.

> ⚠️ **Tesseract still needs to be installed separately** on the target machine
> if image-based invoice OCR is required. Tesseract cannot be bundled into the .exe.

### Optional: Add an Icon

1. Create or download a `.ico` file (e.g., `icon.ico`).
2. Place it in the project folder.
3. Add `--icon icon.ico` to the PyInstaller command above.

### Optional: One-Folder Build (Faster startup, easier to debug)

Replace `--onefile` with `--onedir` for a folder-based distribution.
The folder `dist\InventoryManagerPro\` will contain the `.exe` plus all supporting files.

---

## 8. Troubleshooting

| Problem | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named 'customtkinter'` | venv not activated | Run `venv\Scripts\activate` |
| `TesseractNotFoundError` | Tesseract not in PATH | Add `C:\Program Files\Tesseract-OCR` to PATH |
| `No items extracted from invoice` | Invoice is a scanned image inside a PDF | Use Image upload mode, or ensure Tesseract is installed |
| Excel stock report shows "file is locked" | Report file open in Excel | Close the previous report in Excel before generating a new one |
| `database is locked` error | Another process / second copy of the app has the `.db` open | Close other instances; on Windows, check Task Manager for stray `InventoryManagerPro.exe` |
| `database disk image is malformed` | Power loss while writing | Restore from your most recent backup `.db` |
| Fuzzy matching gives wrong item | Names too different | Add the item manually to the master first; use consistent SKU codes |
| .exe crashes on launch | Missing hidden import | Rerun PyInstaller and add `--hidden-import <module_name>` |
| Date not auto-detected | Non-standard date format | Enter date manually in DD/MM/YYYY format |
| `xlrd.XLRDError` on .xls consumption files | Old-format Excel | Convert to .xlsx in Excel and re-upload |

---

## 9. Data Schema Reference

The master file is now a **SQLite database** (`.db`). The Stock Report
remains an Excel deliverable (`.xlsx`) with columns described at the
bottom of this section.

### SQLite Tables

#### `items` — master record per SKU
| Column | Type | Description |
|---|---|---|
| id | INTEGER PK | Auto-incremented |
| item_name | TEXT | Full product name (required) |
| sku_code | TEXT | Unique product identifier (may be empty) |
| hsn_sac | TEXT | Harmonised System Nomenclature code |
| gst_rate | TEXT | e.g., `"18%"` |
| rate | REAL | Unit price (₹) |
| measurement | TEXT | Unit of measure (kg, pcs, box, etc.) |
| created_at / updated_at | TEXT | ISO timestamps |

#### `invoice_receipts` — one row per (item, date)
| Column | Type | Description |
|---|---|---|
| id | INTEGER PK | Auto-incremented |
| item_id | INTEGER FK → items.id | Cascade-deleted with parent item |
| receipt_date | TEXT | `DD/MM/YYYY` |
| quantity | REAL | Quantity received |
| source_file | TEXT | Optional, name of the invoice file |

`UNIQUE(item_id, receipt_date)` — multiple invoices on the same day
accumulate via UPSERT, so quantities add up safely.

#### `consumption` — append-only consumption log
| Column | Type | Description |
|---|---|---|
| id | INTEGER PK | Auto-incremented |
| item_id | INTEGER FK → items.id | Cascade-deleted with parent item |
| quantity | REAL | Quantity consumed |
| source_file | TEXT | Optional, name of the consumption report |
| recorded_at | TEXT | ISO timestamp |

#### `meta` — key/value housekeeping
Stores `schema_version`, `created_at`, and other settings.

### Current Stock Report Columns (Excel output)

| Column | Source |
|---|---|
| S.No. | Output row number |
| Item Name … Measurement | From `items` |
| Total Received | `Σ invoice_receipts.quantity` per item |
| Consumed Qty | `Σ consumption.quantity` per item |
| Available Qty | `max(0, Total Received − Consumed Qty)` |
| Status | `In Stock` / `Low Stock` (≤5) / `Out of Stock` (0) |

### Consumption Report Columns (minimum required)
- **Item Name** or **SKU Code** — at least one identifier
- **Quantity** — consumed quantity

---

## 10. Edge Cases & Tips

### Fuzzy Matching Thresholds
- **90%** similarity → near-certain (SKU/code match)
- **78%** similarity → probable (name match, minor differences)
- **65%** similarity → possible match (flagged in log)

Items below 65% similarity are logged as **unmatched** and require manual review.

### Recommended Practices
- Always use **SKU Codes** on your invoices — they are the most reliable matching key.
- Keep item names consistent across suppliers to maximise fuzzy-match accuracy.
- Process invoices in **date order** — the system handles out-of-order dates but
  the database will be cleaner if dates are sequential.
- Back up `Master_Inventory.db` regularly (e.g., weekly copy to a dated folder).
  A single `.db` file copy is a complete backup.

### SQLite Sidecar Files
SQLite runs in **WAL mode** for safer concurrent reads/writes. While the app
is open you may see two extra files next to your `.db`:

  - `Master_Inventory.db-wal`  — write-ahead log
  - `Master_Inventory.db-shm`  — shared memory file

These are **normal**. They merge back into the main `.db` on clean close.
Don't delete them while the app is running.

### Multiple Invoices per Day
The system **adds** to the existing receipt if the same `(item, date)` already
exists in `invoice_receipts` (enforced by a `UNIQUE` constraint + UPSERT).
You can safely upload multiple invoices from the same date — quantities
accumulate correctly.

### Negative Stock
The Current Stock report shows **0** (not negative) as the minimum for Available Qty.
A negative result means your consumption records exceed your recorded purchases —
check for missing invoices.
