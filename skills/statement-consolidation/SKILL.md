---
name: statement-consolidation
description: Extract bank-statement transactions from PDF, CSV, Excel, or OFX files into a template-driven master Excel table, classify them using editable workbook rules, reconcile source totals, and generate debit, P&L, and consolidated statement views. Use when the user supplies bank statements plus an Excel template, asks to consolidate statements, or asks to design, validate, or change such a workbook.
---

# Statement consolidation

## Safety and scope

- Treat every bank statement as sensitive financial information. Process it only in the user-authorized workspace. Do not paste account numbers, identifiers, or transaction detail into chat unless it is necessary to answer the request. Do not upload statements to a third-party service without explicit user approval.
- Do not promise support for every bank, scan quality, language, or document layout. Inspect the supplied material and state how each file was read.
- Never silently infer a debit/credit direction, a currency, a date year, a transaction category, an opening balance, or a duplicate decision. Record the uncertainty as an exception and ask for review when it materially affects the result.
- Do not use the resulting workbook for payment execution, tax filing, or other regulated decisions without human review.

## Tools

| Tool | Use it for |
| --- | --- |
| `bank_statement_create_template` | Making a fresh workbook when the user has none. |
| `bank_statement_inspect_template` | Checking a workbook before importing into it. Always run this first. |
| `bank_statement_consolidate` | The main import. Reads CSV, XLSX/XLS, OFX/QFX, and text-layer PDF, from files or a folder. |
| `bank_statement_add_transactions` | Fallback for scanned/image PDFs only, after you have read the document yourself. |
| `bank_statement_update_template` | Changing category rules and report settings on request, instead of asking the user to edit Excel. |

## Inputs

Accept an Excel workbook plus either selected statement files or a folder containing them. `statementPaths` takes folders directly; a folder expands to the statement files immediately inside it.

`bank_statement_consolidate` machine-reads CSV, XLSX/XLS, OFX/QFX, and PDFs that carry a text layer. Prefer structured formats when the user has a choice: OFX/QFX, then CSV, then XLSX/XLS, then PDF. A PDF is reconstructed from text positions, so its rows are recorded as `Parsed (PDF text layer)` and must be checked against the statement's own totals.

Never write over the input workbook. `outputPath` must differ from `templatePath`; the tool refuses otherwise.

## When a PDF cannot be machine-read

A scanned or image-only PDF has no text layer, and `bank_statement_consolidate` will report it as an exception rather than importing nothing silently. Only then:

1. Read the document yourself and transcribe the transactions you can actually see.
2. Call `bank_statement_add_transactions` with those rows and the source filename.
3. Say plainly, in your summary, which rows came from your own reading rather than from a parser. They are written as `Parsed (assistant-extracted)` and flagged on the Audit sheet.

Do not infer, complete, or average a row you could not read. Leave it out and report it as unread.

## Workbook contract

The workbook must contain these sheets, found by exact name: `Start Here`, `Template Columns`, `Category Rules`, `Report Setup`, `Master Transactions`, `Debit statement`, `P&L statement`, `Consolidated statement`, `Audit`. If the user's workbook differs, offer `bank_statement_create_template` rather than hard-coding around it.

`Master Transactions` holds fourteen ordered columns: `Transaction date`, `Description`, `Debit`, `Credit`, `Amount`, `Currency`, `Account ID`, `Source file`, `Source row`, `Category`, `Subcategory`, `P&L group`, `Transaction fingerprint`, `Parse status`.

`Category Rules` applies active rules from the lowest `Priority` upward. Supported match types are `equals`, `contains`, `starts_with`, `regex`, and `amount_range` (written `low..high`). Two active rules at the same priority that both match are an exception, not a tie-break.

`Report Setup` drives the report sheets. Each row refreshes the sheet whose name matches `Report name`. The value columns accept a fixed vocabulary; an unrecognised value causes that report to be skipped and logged, never guessed.

`Transaction fingerprint` is a SHA-256 over account, date, amount, currency, description, source file, and source row. Rows are appended only when the fingerprint is new. Never overwrite an existing master row because its narrative looks similar.

## Required workflow

1. Inventory each source: filename, type, masked account identifier, stated period, opening and closing balances when shown, row or page count.
2. Run `bank_statement_inspect_template` and resolve what it flags before importing.
3. Run `bank_statement_consolidate`. For any file it reports as unreadable, decide whether the scanned-PDF fallback above applies.
4. Reconcile. Where the statement shows opening and closing balances, check the movement matches. Otherwise reconcile the record count and totals against whatever control the source provides. Report any gap; do not close it yourself.
5. Read the `Audit` sheet and report every exception on it.
6. If rows came back `Uncategorized`, propose rules and, once the user approves the categories, apply them with `bank_statement_update_template`, then re-run the import to reclassify. Do not invent a P&L group.
7. Summarize the results and the open exceptions, and say where the output file was written.

## Changing the workbook on request

When the user asks for a new category, a different grouping, or a report turned off, use `bank_statement_update_template` rather than telling them to edit Excel by hand. It validates against the same vocabulary the dropdowns offer and rejects anything outside it.

Rules only affect rows at import time, so after changing rules re-run `bank_statement_consolidate` to reclassify. Editing the sheets in Excel remains equally valid; both paths lead to the same workbook.

## Completion criteria

Report the source files processed and how each was read, records extracted, rows appended, rows skipped as duplicates, uncategorized rows, and every reconciliation or parsing exception. Do not call the import complete while any source cannot be parsed or reconciled well enough for the requested use.
