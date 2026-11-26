# PH Roster Planner

This is an Excel workbook with macros that plans the staff roster for public holidays. A build script creates the workbook from the files in this repository, and the roster logic is written in VBA with a Python twin used for checking.

## Requirements

- Windows with desktop Excel (version 16, that is Excel 2016 or later, including Microsoft 365).
- Excel must trust access to the VBA project. In Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings, tick "Trust access to the VBA project object model", click OK, then close Excel.

## Build

Open PowerShell in the repository folder and run:

    powershell -ExecutionPolicy Bypass -File build\build.ps1

The script opens Excel in the background, builds the workbook, saves it as `dist\PHRosterPlanner.xlsm` and runs the built-in self-test. What the result means:

| Exit code | Meaning |
| --- | --- |
| 0 | The self-test printed PASS. The workbook is ready to use. |
| 2 | The self-test ran and printed FAIL lines. The workbook is saved, but a rule does not behave as designed. Send the printed lines to the developer. |
| 3 | Excel could not run the self-test, usually because the VBA has a compile error. Open `dist\PHRosterPlanner.xlsm`, press Alt+F11, then choose Debug > Compile VBAProject to find the line. |

Exit code 1 means the trust setting above is switched off. Exit code 4 means something else went wrong while building; the message printed above it says what. In every case where the workbook was saved, it is kept.

To see the exit code after the script ends, run:

    echo $LASTEXITCODE

If the script seems stuck, look for an Excel or VBA dialog.

## First run

1. Open `dist\PHRosterPlanner.xlsm` and click Enable Content (enable macros).
2. The workbook comes with fictional sample data. Replace it with your own on the Staff, Holidays, LastYear and Leave sheets, and check the Settings sheet (year, number of slots, rule switches). Keep the header in row 1 of every sheet.
3. On the Holidays sheet click "Expand long weekends" to add the neighbouring Saturday, Sunday or Monday around holidays that fall on a Friday, Saturday, Sunday or Monday (a Saturday or Sunday holiday adds the following Monday).
4. On the Settings sheet click "Generate roster". Problems in the input are shown in a message box, or listed on the Review sheet from cell I1 when there are many. Row numbers in these messages count data rows: row 1 is the first row under the header, and blank rows are not counted.
5. Open the Roster sheet. Slots that could not be filled show GAP. Work through the Review sheet: each row says which rule was not met and who could be considered. Type a staff name in the Approve column for each row you want to fill.
6. On the Review sheet click "Apply approvals". Each approval asks for confirmation and lists the rules it would break. Later dates are regenerated around your choices. When at least one approval is applied, the Approve column is cleared, so entries that were skipped or declined must be typed again.

The "Run self-test" button on the Settings sheet repeats the built-in check at any time.

## Starting a new year

When the year is finished, click "Start new year" on the Settings sheet and confirm. The current roster becomes the LastYear sheet, Holidays, Leave, Roster, Review, Summary and Approvals are cleared, and the year in Settings goes up by one. Then enter the new holidays and leave and generate again.

## Developing

Run these from the repository root:

    python -m pytest python/tests -q
    python tools/check_vba.py src/vba
    python python/roster/gen_fixtures.py

The first runs the Python twin of the engine, the second checks the VBA source files, and the third regenerates the scenario files in `fixtures/`.

## Known limits

- The VBA and the build script had not been executed before the first Windows build. Expect at least one round of fixes after that first run.
- The 2027 sample holiday dates have not been checked against the official list. Verify them (the Holidays sheet says "Sample dates - verify against MOM") before relying on the roster.
- Each date's search stops after 200,000 placements; the Review sheet says when that happened.
