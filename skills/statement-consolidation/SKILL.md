---
name: statement-consolidation
description: Extract bank-statement transactions into a template-driven master Excel table, classify them using editable workbook rules, reconcile source totals, and generate debit, P&L, and consolidated statement views. Use when the user supplies bank statements plus an Excel template, or asks to design, validate, or refresh such a workflow.
---

# Statement consolidation

## Safety and scope

- Treat every bank statement as sensitive financial information. Process it only in the user-authorized workspace. Do not paste account numbers, identifiers, or transaction detail into chat unless it is necessary to answer the request. Do not upload statements to a third-party service without explicit user approval.
- Do not promise support for every bank, scan quality, language, or document layout. Inspect the supplied material and state the parse method and its exceptions.
- Never silently infer a debit/credit direction, a currency, a date year, a transaction category, an opening balance, or a duplicate decision. Record the uncertainty as an exception and ask for review when it materially affects the result.
- Do not use the resulting workbook for payment execution, tax filing, or other regulated decisions without human review.

## Inputs

Accept either an Excel template workbook plus one or more selected statement files, or an Excel template workbook plus a user-selected folder that contains only the in-scope statements.

The `bank_statement_consolidate` tool reads CSV, XLSX/XLS, and OFX/QFX directly. Prefer structured formats in this order: OFX/QFX, CSV, XLSX/XLS. PDFs are not handled by the tool — extract a text PDF yourself for review, or ask the user to supply a structured export. Say which files were machine-read and which were not; never let an unread statement become a silent gap.

Preserve each source file name and its original row or page reference on every transaction.

## Workbook contract

Before processing, call `bank_statement_inspect_template`. The workbook must contain these sheets, found by exact name:

`Start Here`, `Template Columns`, `Category Rules`, `Report Setup`, `Master Transactions`, `Debit statement`, `P&L statement`, `Consolidated statement`, `Audit`.

If the user's workbook uses different names, ask them to map it or offer to generate a fresh template with `bank_statement_create_template`. Do not hard-code columns, categories, or report layouts that the workbook should control.

`Master Transactions` holds fourteen ordered columns: `Transaction date`, `Description`, `Debit`, `Credit`, `Amount`, `Currency`, `Account ID`, `Source file`, `Source row`, `Category`, `Subcategory`, `P&L group`, `Transaction fingerprint`, `Parse status`.

`Category Rules` defines `Priority`, `Look in this field`, `Match type`, `Text or amount to match`, `Category`, `Subcategory`, `P&L group`, and `Use this rule?`. Active rules apply from the lowest `Priority` upward. Only `equals`, `contains`, `starts_with`, `regex`, and `amount_range` are supported. Two active rules at the same priority that both match are an exception, not a tie-break.

`Report Setup` defines `Report name`, `Include transactions when`, `Group by`, `Measure`, and `Show this report?`. Each row refreshes the sheet whose name matches `Report name`. The three value columns accept a fixed vocabulary — see the README table. An unrecognised value causes the report to be skipped and logged, never guessed.

`Transaction fingerprint` is a SHA-256 over account, date, amount, currency, description, source file, and source row. Preserve a source-specific transaction identifier when one is available. Never overwrite an existing master row because its narrative looks similar.

## Required workflow

1. Inventory each source: filename, type, masked account identifier, stated period, opening and closing balances when shown, row or page count, and parse method.
2. Run `bank_statement_inspect_template` and resolve anything it flags before importing.
3. Run `bank_statement_consolidate`. It extracts transactions, applies the workbook's rules, appends only rows whose fingerprint is new, rebuilds the enabled reports from the entire master table, and writes the audit.
4. Reconcile. Where the statement shows opening and closing balances, check that the movement matches. Otherwise reconcile record count and totals against whatever control the source provides. Report any gap; do not close it yourself.
5. Read the `Audit` sheet and report every exception it lists: unreadable files, rows needing review, conflicting rules, and uncategorized rows.
6. If rows came back `Uncategorized`, offer to add rules to `Category Rules` and re-run. Propose the rule; let the user approve the category. Do not invent a P&L group.
7. Recalculate, check for formula errors, and summarize only the results and the open exceptions.

## Designing a new template

When no template exists, call `bank_statement_create_template` rather than hand-building a workbook. It produces the correct sheet names, the fourteen master columns, dropdown-validated rule and report sheets, and shaded cells marking what is safe to edit.

## Completion criteria

Report the source files processed, records extracted, rows appended, rows skipped as duplicates, uncategorized rows, and every reconciliation or parsing exception. Do not call the import complete while any source cannot be parsed or reconciled well enough for the requested use.
