# FILES.md — File Index for Claude

This file lists every source file in the `inventory_manager` repository with
its raw URL on GitHub. It exists so Claude (and other AI assistants) can fetch
any file in this repo without the user having to paste URLs one at a time.

**For Claude:** When the user asks about any file in this repo, fetch its raw
URL from the tables below before proposing changes. These URLs are the source
of truth — prefer them over any cached snapshot.

**For humans:** This file is maintained manually. After adding, renaming, or
deleting a file, update the relevant table in the same commit.

---

## Root

| File                    | Description                                                                   | Raw URL                                                                                          |
| ----------------------- | ----------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `main.py`               | Application entry point — CustomTkinter GUI, threading, workflow orchestration | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/main.py                        |
| `requirements.txt`      | Pinned Python dependencies                                                    | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/requirements.txt               |
| `BUILD_INSTRUCTIONS.md` | Setup, build, and troubleshooting guide                                       | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/BUILD_INSTRUCTIONS.md          |
| `.gitignore`            | Git ignore patterns                                                           | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/.gitignore                     |
| `FILES.md`              | This file — index of all source files for AI assistants                       | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/FILES.md                       |

## core/

| File                           | Description                                                              | Raw URL                                                                                                |
| ------------------------------ | ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------ |
| `core/__init__.py`             | Package marker                                                           | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/__init__.py                     |
| `core/invoice_parser.py`       | PDF/image invoice parser using pdfplumber + Tesseract OCR fallback       | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/invoice_parser.py               |
| `core/database_manager.py`     | SQLite storage, Excel-to-SQLite migration, stock report generator        | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/database_manager.py             |
| `core/excel_manager.py`        | Legacy Excel storage (kept as fallback during transition)                | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/excel_manager.py                |
| `core/consumption_processor.py` | Consumption report parser (PDF, Excel, CSV)                              | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/consumption_processor.py        |
| `core/fuzzy_matcher.py`        | rapidfuzz-based item name / SKU matching utility                         | https://raw.githubusercontent.com/akpsunny/inventory_manager/main/core/fuzzy_matcher.py                |

---

## How this file is used

At the start of a Claude conversation about this project, Claude fetches
`FILES.md` first. Once that fetch completes, every URL in the tables above
is in Claude's context and becomes individually fetchable on demand.

A useful first-message pattern:

> "First, fetch https://raw.githubusercontent.com/akpsunny/inventory_manager/main/FILES.md
> to see the file index. Then look at core/invoice_parser.py and tell me…"

Or, if you've added an instruction in the project's system prompt to always
fetch `FILES.md` first, you can skip this and just ask the question directly.

## Maintenance

Update this table whenever a file is added, renamed, or deleted. The file
list is small and changes infrequently, so a manual update in the same
commit as the change is sufficient. A pre-commit hook could automate it
later if the project grows.

## Descriptions disclaimer

The descriptions in the tables above are short summaries inferred from
file names, the `main.py` imports, and `BUILD_INSTRUCTIONS.md` §1. They
may drift from reality if the code changes substantially — refresh them
when updating the table.
