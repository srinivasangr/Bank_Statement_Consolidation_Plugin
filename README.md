# Bank Statement Consolidator

A Codex plugin that turns a set of bank statements and an editable Excel workbook into a reconciled, consolidated workbook. Everything runs locally over stdio: the service never uploads a statement.

The point of the design is that a non-technical user controls the output by editing the workbook, not by changing code. Categories, reporting groups, and which reports get built all come from sheets in the template.

## Install

Requires Node.js 20 or later.

```bash
npm install
```

`npm install` runs the build automatically and produces `dist/`, which `.mcp.json` starts.

## Quick start

1. Copy `assets/BankStatementTemplate.xlsx` and open your copy. The `Start Here` sheet explains the rest.
2. Edit the shaded sheets so they match your client: `Template Columns`, `Category Rules`, `Report Setup`.
3. Give Codex the workbook and the statement files, and ask: `Consolidate these statements using this template.`
4. Read the `Audit` sheet before you use the output.

## How the workbook works

| Sheet | Who fills it | Purpose |
| --- | --- | --- |
| `Start Here` | — | Plain-language instructions. |
| `Template Columns` | You | Documents what each master column means. Reference for humans. |
| `Category Rules` | You | Rules that assign Category, Subcategory, and P&L group. |
| `Report Setup` | You | Controls which reports are built, and how each one groups and measures. |
| `Master Transactions` | Service | The source of truth. Rows are appended, never rewritten or deleted. |
| `Debit statement` | Service | Rebuilt from the master table on every import. |
| `P&L statement` | Service | Rebuilt from the master table on every import. |
| `Consolidated statement` | Service | Rebuilt from the master table on every import. |
| `Audit` | Service | Counts, reconciliation, and every exception. Terminal — nothing reads from it. |

Sheets are found by these exact names. `Master Transactions` must keep its fourteen columns in order: Transaction date, Description, Debit, Credit, Amount, Currency, Account ID, Source file, Source row, Category, Subcategory, P&L group, Transaction fingerprint, Parse status.

### Category Rules

Active rules are applied from the lowest `Priority` number upward, and the first match wins. If two active rules at the *same* priority both match, that is an exception: the row keeps `Uncategorized` and is flagged for review rather than being assigned arbitrarily.

`Match type` accepts `contains`, `equals`, `starts_with`, `regex`, and `amount_range` (written as `low..high`). `Look in this field` accepts `Description` or `Amount`. `Use this rule?` accepts `Yes` or `No`. All of these are dropdowns in the shipped template.

### Report Setup

Each row refreshes the sheet whose name matches `Report name`. The other columns accept a fixed vocabulary, offered as dropdowns:

| Column | Accepted values |
| --- | --- |
| `Include transactions when` | `All transactions`, `Debit > 0`, `Credit > 0`, `P&L group is not blank`, `Category is not blank` |
| `Group by` | `Category`, `Subcategory`, `P&L group`, `Currency`, `Account ID`, `Source file`, `Transaction date` |
| `Measure` | `Sum of Amount`, `Sum of Debit`, `Sum of Credit`, `Count of transactions` |
| `Show this report?` | `Yes`, `No` |

A value outside this list is not guessed at. The report is skipped and the reason is written to `Audit`.

## Tools

- `bank_statement_create_template` — writes a fresh template workbook to a path you choose.
- `bank_statement_inspect_template` — checks a workbook for the required sheets, the master columns, at least one active rule, and any unrecognised `Report Setup` values.
- `bank_statement_consolidate` — reads the statements, applies the rules, appends only non-duplicate rows, rebuilds the enabled reports from the whole master table, and writes the audit.

## Duplicate handling

Every row gets a SHA-256 `Transaction fingerprint` over account, date, amount, currency, description, source file, and source row. On import, a row is skipped if its fingerprint already exists in `Master Transactions` or appeared earlier in the same import. Re-running the same statements against the output appends nothing.

## Supported inputs and limits

CSV, XLSX/XLS, and OFX/QFX are read automatically. Header names are matched flexibly, so `Narration`, `Particulars`, `Withdrawal`, and `Deposit` are understood alongside the obvious ones.

PDFs are **not** parsed by this service. A text or scanned PDF needs Codex to extract it for review, or a bank-specific parser/OCR adapter. A statement the service cannot read becomes an exception on the `Audit` sheet; it never becomes a silent gap in the numbers.

Statement layouts vary by institution and period. Do not use unreviewed output for tax filings, payments, or regulated reporting.

## Development

```bash
npm run check           # typecheck
npm run build           # compile src/ to dist/
npm test                # end-to-end test against the real MCP service over stdio
npm run build:template  # regenerate assets/BankStatementTemplate.xlsx
npm run verify:template # confirm the shipped template is readable and complete
```

The shipped template is generated by the same `createTemplateWorkbook()` that `bank_statement_create_template` calls, so the asset and a freshly created workbook cannot drift apart. Regenerate and re-verify it after any change to `src/template.ts`.

Test fixtures are synthetic. `.gitignore` blocks `*.csv`, `*.xlsx`, `*.pdf`, `*.ofx`, and `*.qfx` by default and allowlists only the template and the fixtures, so a real statement cannot be committed by accident.

## Before distributing outside a trusted team

1. Add public privacy-policy and terms URLs to `.codex-plugin/plugin.json`.
2. Add fixtures for each institution and layout you intend to support. Synthetic or irreversibly redacted statements only.
3. Add a parser/OCR service if you need dependable PDF extraction. Keep documents local, or document any processor you use.
4. Validate the plugin and test it in a fresh Codex task before publishing it to a marketplace or team channel.

## License

MIT. See [LICENSE](LICENSE).
