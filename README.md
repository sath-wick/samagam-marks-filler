# MARKS FILLER - CHROME EXTENSION

## What it does

1. Loads an .xlsx/.xlsm/.xlsb/.xls workbook locally in the extension.
2. Uses the "Student Name" column as the identity column.
3. Shows every other non-empty header as an exam selector (e.g. MT (Apr), MT (JULY), MT (Aug), MT (SEPT)).
4. Scans the current webpage's #stageMarksGrid.
5. Matches students by normalized name, NOT by row order.
6. Blocks filling if any student is unmatched, duplicated, or has an invalid mark.
7. Fills only the selected exam's marks. It does NOT click Save/Submit.

## Install in Chrome

1. Unzip this folder.
2. Open chrome://extensions
3. Enable Developer mode.
4. Click "Load unpacked".
5. Select the unzipped marks_extension folder.
6. Open the marks page.
7. Click the extension icon.
8. Choose the Excel file and exam.
9. Click "Scan & verify names".
10. Only after the page reports "Verified" click "Fill marks".
11. Review the entered fields and use the website's normal Save/Submit control.

## Safety behavior

The extension intentionally refuses to fill when a page student cannot be matched to exactly one Excel student. This matters because spreadsheet row order can differ from the website.

The extension never uploads the workbook to a server. The workbook is parsed locally in the popup.
