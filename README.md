# NABFID Offense Processor

Simple browser-based Excel processor for NABFID offense reports.

## V1 workflow
1. Upload an Excel file.
2. Process the first worksheet.
3. Generate Classification and Organization fields.
4. Show client-wise status summary.
5. Preview or edit processed data.
6. Download a new Excel workbook.

## Rules
- Classification: extract the value after `Domain:` and before `|`. Empty Domain becomes `NABFID DC`.
- Organization: if Subject contains `KPMG` (case-insensitive), use `KPMG Assurance and Consulting Services LLP`; otherwise use `Edgeverve Systems Limited- NaBFID`.

The app uses SheetJS in the browser, so the file is processed client-side.
