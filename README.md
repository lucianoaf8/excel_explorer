# Excel Explorer

A lightweight desktop utility for quickly inspecting Excel workbooks and exporting interactive HTML reports.

## Current Features

* **Cross-platform GUI (Tkinter)** – open any `.xlsx` file via a friendly interface.
* **Asynchronous Analysis** – the UI stays responsive while the workbook is processed.
* **Progress Feedback** – circular progress indicator, live status messages, and elapsed timer.
* **Automatic Tab Switching** – jumps to the report tab when analysis finishes.
* **Auto-exported HTML Report** – saved with a timestamped filename inside a dedicated `reports/` folder.
* **Collapsible HTML Report** –
  * File overview (full path, size, created & modified dates).
  * Ordered sheet list.
  * Per-sheet collapsible sections showing range, data vs. empty cells, and a table of columns (letter, range, dominant data type).
  * Module execution status and high-level metrics.
* **Open Last Report Button** – instantly launches the most recent HTML report in your default browser.

## Installation

```
python -m venv venv
venv\Scripts\activate  # or source venv/bin/activate on macOS/Linux
pip install -r requirements.txt
```

## Usage

```
python excel_explorer_gui.py
```

1. Click **Select Excel File** and choose a workbook.
2. Press **Analyze** – watch the progress indicator.
3. When complete, view the embedded preview or click **Open HTML Report**.

## Dependencies

* Python ≥ 3.9
* `openpyxl`
* `tkinter` (bundled with standard Python on Windows/macOS; install `python3-tk` on some Linux distros)

## Project Structure (key files)

```
excel_explorer/
├─ excel_explorer_gui.py   # Main GUI application
├─ analyzer.py             # Workbook analysis logic
├─ report_generator.py     # HTML / JSON report creation
├─ reports/                # Auto-generated reports
└─ config.yaml             # Optional runtime configuration
```

---
*Everything in this README reflects the **currently implemented and working** functionality.*
