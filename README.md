# 📊 Google Sheets Advanced Audit Logger

This Google Apps Script provides **industry-ready auditing** inside Google Sheets.  
It creates a dedicated `Logs` sheet and records **all important user actions** in clear, human-readable sentences.

---

## 🚀 Features

### ✅ Cell Activity
- Tracks **single cell edits**
  - Captures **old value → new value**
  - Works with **formulas** (records formula text instead of just result)
- Tracks **bulk edits**
  - Pasting, dragging, or clearing multiple cells
  - Records range (`A1:C10`) and sample of first value/formula

### ✅ Structural Changes
- **Rows**
  - Inserted (with index)
  - Deleted (with index)
- **Columns**
  - Inserted (with index)
  - Deleted (with index)
- **Sheets**
  - Added (with name)
  - Deleted (with name)
  - Renamed (detected, though does not show old→new yet)

### ✅ Logging Details
- Every log entry contains:
  - **Timestamp** (yyyy-MM-dd HH:mm:ss)
  - **User** (email address of editor, if available)
  - **Action type** (EDIT, BULK_EDIT, INSERT_ROW, REMOVE_COLUMN, etc.)
  - **Readable description** of the action

### ✅ Example Logs

| Timestamp | User | Action Type | Details |
|---------------------|---------------------------|---------------|---------|
| 2025-09-02 13:40:11 | gilbert.keter@gmail.com | EDIT | gilbert.keter@gmail.com edited B4 on 'Sales' from '34' to '45' at 2025-09-02 13:40:11 |
| 2025-09-02 13:42:55 | gilbert.keter@gmail.com | REMOVE_ROW | gilbert.keter@gmail.com deleted row(s) (approx. at index 5) in 'Inventory' at 2025-09-02 13:42:55 |
| 2025-09-02 13:43:20 | gilbert.keter@gmail.com | INSERT_COLUMN | gilbert.keter@gmail.com inserted column(s) (approx. at index 3) in 'Expenses' at 2025-09-02 13:43:20 |
| 2025-09-02 13:44:00 | gilbert.keter@gmail.com | REMOVE_GRID | gilbert.keter@gmail.com deleted sheet 'Archive2024' at 2025-09-02 13:44:00 |

---

## ⚙️ Installation

1. Open your Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Copy-paste the provided `audit_logger.js` script.
4. Save and close the editor.
5. Refresh your spreadsheet.

A `Logs` sheet will be created automatically on first use.

---

## ⚠️ Limitations

- **Formula recalculations** (caused by dependencies) are not logged, only user edits.  
- **Script-based edits** (from other scripts) are not logged — only manual user actions.  
- **Row/column insert/delete indices** are approximate, based on current selection.  
- **Renames** are detected but currently only logged as “renamed a sheet” (no old→new mapping).  

---

## 📌 Roadmap / Possible Improvements
- Detect **sheet rename old → new** for more detail.  
- More precise row/column change detection (index diffs).  
- Option to export logs to an external database for compliance.  

---

## 👨‍💻 Author
Developed for **advanced auditing** in collaborative spreadsheets.  
Maintainer: *Gilbert Keter*

---
