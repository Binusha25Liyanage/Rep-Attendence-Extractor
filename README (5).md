# Rep Attendance Image Maker

A small desktop app that turns a sales-rep attendance Excel export from an SFA (Sales Force Automation) system into clean, cropped **PNG images** and a **PDF**, ready to share in a chat or report.

Pick the Excel file, choose which columns you want, and click **Generate**.

## Features

- Load the attendance `.xlsx` export with a file picker
- Choose any columns for the table and set their left-to-right order
- Filter rows: not logged out, logged out, both (two images) or everyone who logged in
- Sort by Excel order, territory A-Z or login time
- Editable title, with the report date filled in automatically
- Save the two images and the PDF into **three different folders**, remembered for next time
- Live preview inside the app
- Output is tightly cropped at 300 dpi

## Requirements

- Python 3.9 or newer
- `tkinter` (included with the standard Windows and macOS Python installers)

```bash
python -m pip install openpyxl reportlab pillow
```

## Usage

```bash
python rep_attendance_app.py
```

1. Click **Choose Excel file…** and select the attendance export.
2. Double-click columns on the left to add them. Use **Move up / Move down** to reorder.
3. Choose which rows to show, the sort order and the title.
4. Set the folder for Image 1, Image 2 and the PDF.
5. Click **Generate images + PDF**.

Output files:

| File | Contents |
|---|---|
| `rep_attendance_not_logged_out.png` | Reps who logged in but have no logout time |
| `rep_attendance_logged_out.png` | Reps who logged in and logged out |
| `rep_attendance.pdf` | The same tables, one page per image, cropped to the table |

### No window needed

`rep_attendance_extractor.py` is a command-line version with the default layout (territory, logged in, logged out):

```bash
python rep_attendance_extractor.py "path\to\RepAttendance_Reports.xlsx"
```

## How it works

The app finds the header row in the export, labels duplicate column names by their section (for example `LOGGED IN (ATTENDANCE)` and `LOGGED IN (ASM ATTENDANCE)`), and keeps only reps who have a login time. A rep counts as "not logged out" when the LOGGED OUT cell is empty. The table is drawn once with ReportLab (PDF) and once with Pillow (PNG) from the same layout.

## Notes

- The export layout is expected to match the SFA "RepAttendance Reports" format. A different layout may need small changes in `read_excel()`.
- Attendance data can contain names and phone numbers. Don't commit real exports or generated images to a public repository.

## Project files

```
rep_attendance_app.py         # desktop app (main)
rep_attendance_extractor.py   # command-line version
README.md
```
