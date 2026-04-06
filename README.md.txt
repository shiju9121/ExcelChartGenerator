# Excel Chart Generator

A Python program that generates charts directly inside Excel files using **pandas** and **xlsxwriter**.  
It is designed to be interactive, user-friendly, and safe — preventing accidental overwrites and handling invalid inputs gracefully.

---

## ? Features
- Choose an input Excel file (with automatic quote stripping if pasted from File Explorer).
- Select X-axis and Y-axis columns by **number, name, or ranges** (e.g., `3-5`).
- Supports **multiple Y-axis series** in one chart.
- Chart types: **bar, pie, line**.
- Prevents overwriting existing files unless confirmed.
- Handles **PermissionError** if the output file is open in Excel.
- Interactive loop: generate multiple charts in multiple output files.

---

## ?? Requirements
- Python 3.x
- pandas
- xlsxwriter

Install dependencies:
```bash
pip install pandas xlsxwriter
