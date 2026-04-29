# Student Information System Notes

## Overview

This is a Tkinter desktop app for managing student scholarship records. The main application entry point is `main.py`.

The app stores student data locally in JSON files and exports scholarship/payroll reports into Excel and Word templates.

## Environment

Use the project virtual environment:

```powershell
. .\env\Scripts\Activate.ps1
python main.py
```

Install dependencies with:

```powershell
pip install -r requirements.txt
```

Required packages:

- `openpyxl` for Excel generation.
- `pywin32` for Microsoft Word COM automation used by `.doc` export.

The Word export requires Microsoft Word installed on Windows.

## Important Files

- `main.py`: Tkinter app and all export logic.
- `student_data.json`: active student records.
- `deleted_students.json`: trash records, created when needed.
- `school_course_options.json`: saved school/course dropdown options.
- `PAYROLL_TEMPLATE.xlsx`: Excel payroll template.
- `PAYROLL_WORD_TEMPLATE.doc`: Word payroll template.
- `requirements.txt`: Python dependencies.
- `.gitignore`: excludes `env/`, `venv/`, caches, and temporary Office lock files.

Do not commit the virtual environment folder.

## Student Data Model

Each student record is a JSON object with fields like:

- `full_name`
- `barangay`
- `address`
- `contact_number`
- `school`
- `course`
- `school_year`
- `batch`
- `status`
- `documents`
- `registration_date`

Renewal state is tracked by the presence of:

- `renewal_date`
- `renewal_requirements`

If `renewal_date` exists, the student is considered renewed. If it does not exist, the student is considered unrenewed/new.

## Regular Excel Export

The regular student-list export is handled by `export_to_excel`.

It can export:

- all students
- renewed students
- unrenewed students

It creates:

- one sheet per batch
- an `All Students` sheet

This export is a generated workbook and does not use `PAYROLL_TEMPLATE.xlsx`.

## Payroll Export Bundle

The `Print to Payroll` flow exports three files into one generated folder:

- `student_list_<filter>.xlsx`
- `payroll_<filter>.xlsx`
- `payroll_<filter>.doc`

The app asks the user to select a parent folder. It then creates a timestamped folder:

```text
student_exports_<filter>_<YYYYMMDD_HHMMSS>
```

The export runs asynchronously on a background thread. A modal loading dialog with an indeterminate progress bar is shown while files are being created.

## Payroll Excel Template Rules

Template file:

```text
PAYROLL_TEMPLATE.xlsx
```

The payroll Excel export creates one worksheet per 15 students.

Per sheet:

- Student names go in `B10:B24`.
- `E10:E24` gets `5000` for each real student row.
- `J10:J24` also gets `5000` for each real student row.
- `X-X-X-X` is written after the last student entry.

Marker behavior:

- If the page has fewer than 15 students, the marker goes in the next available name row.
- If the page has exactly 15 students, the marker goes in `B25`.

Do not manually write the total cell. The template handles totals with its own formula.

## Payroll Word Template Rules

Template file:

```text
PAYROLL_WORD_TEMPLATE.doc
```

This is a legacy Word `.doc` file, so it is edited through Microsoft Word COM automation using `pywin32`.

The Word export creates one `.doc` file with multiple copied template pages. It does not create multiple Word files.

Per 15-student section:

- The full template page is copied, not just the table.
- This preserves signatory/signature areas.
- One table/page is filled per 15 students.

Word table mapping:

- Column 2: student name
- Column 3: year level
- Column 4: school

Word COM reports the SCHOOL field as column 4 for student rows. Do not write school names to column 5, because Word raises `The requested member of the collection does not exist`.

Word formatting rules:

- Names are ALL CAPS.
- School names are ALL CAPS.
- `San Jose City` is removed from school names before writing.
- Year levels use ordinal format:
  - `1ST YEAR`
  - `2ND YEAR`
  - `3RD YEAR`
  - `4TH YEAR`
- Nonnumeric year values become `<VALUE> YEAR`.

Word sorting rules:

- Sort by year level first.
- Within each year level, sort alphabetically by last name.

Word marker behavior:

- `X-X-X-X` is always written after the last student entry on every page.
- If the page has fewer than 15 students, it goes in the next available row.
- If the page has exactly 15 students, it goes in row 17 of the Word table.

## Sorting

Payroll Excel `all` export currently sorts by last name before sectioning into 15-person sheets.

Payroll Word export sorts independently by:

1. year level
2. last name

Do not assume the Excel payroll and Word payroll order are identical unless the requirements are changed.

## Async Export Notes

The export flow is split into:

- `print_to_payroll`: gathers UI inputs and starts the worker thread.
- `export_payroll_bundle`: creates all three files in the background.
- `show_export_loading`: shows the loading window.
- `poll_export_result`: reports success or failure on the Tk main thread.

Tkinter UI calls must stay on the main thread. Worker threads should not call `messagebox` directly.

## Common Gotchas

- Use `env\Scripts\python.exe`, not the global Python, because `pywin32` is installed in the project env.
- If Word export fails, close any open copies of `PAYROLL_WORD_TEMPLATE.doc` or generated payroll docs.
- If Excel template save fails, close any open copies of generated `.xlsx` files.
- Office temporary lock files like `~$PAYROLL_TEMPLATE.xlsx` should not be committed.
- The project has no git repo by default in this workspace; initialize one separately if needed.

