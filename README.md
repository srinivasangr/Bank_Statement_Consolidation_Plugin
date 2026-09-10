# Bank Statement Consolidator

A Codex plugin that turns a pile of bank statements — PDF, CSV, Excel, or OFX — into one reconciled Excel workbook. Everything runs locally over stdio: the service never uploads a statement.

The design goal is that a non-technical user controls the output by editing a workbook or by asking in plain English, never by changing code.

## What the interface actually is

There is no window and no upload button. This is a Codex plugin, so the interface is a conversation:

- **You talk to Codex.** The bundled `statement-consolidation` skill teaches it the workflow and the safety rules.
- **Codex calls the tools.** Five local MCP tools do the real work on files on your disk.
- **Files go in and out by path.** Codex already runs on your machine, so "uploading" is just telling it where the statements are.

A session looks like this:

> **You:** Consolidate the statements in `C:\Clients\Acme\2026-09` using `C:\Clients\Acme\Acme-Master.xlsx`. Write the result to `C:\Clients\Acme\Acme-2026-09.xlsx`.
>
> **Codex:** Read 4 files. Appended 212 transactions, skipped 6 duplicates. 3 rows are uncategorized and 41 came from a PDF text layer — check those against the statement totals. Audit sheet has the detail.
>
> **You:** Anything with STRIPE is revenue. Redo it.
>
> **Codex:** *(adds the rule, re-runs the import)* STRIPE rows are now Revenue. 0 uncategorized.

## Install

### From GitHub (what you give other people)

```bash
codex plugin marketplace add srinivasangr/Bank_Statement_Consolidation_Plugin
codex plugin add bank-statement-consolidator@bank-statement-tools
```

Nothing else is needed. `bundle/server.mjs` is committed, so the plugin runs
without `npm install` and without a build step. Then start a **new Codex thread**.

### From a local checkout (what you use while developing)

Requires Node.js 20 or later.

```bash
npm install     # builds dist/ and refreshes bundle/
```

Then register the folder as a local marketplace. The plugin must sit inside the
marketplace root as `<root>/plugins/<plugin-name>`:

```bash
codex plugin marketplace add <root>
codex plugin add bank-statement-consolidator@<marketplace-name>
codex plugin list                            # confirm: installed, enabled
```

Codex copies the folder into its plugin cache, so after changing code run
`npm run build && npm run bundle`, re-run `codex plugin add`, and start a **new
thread** — tools and skills are read at thread start.

Do not keep both the GitHub and the local install enabled at once; they register
the same five tool names.

## Testing without Codex

`npm test` drives the real MCP service over stdio and asserts on the workbooks it produces — 39 checks covering all five tools, all four input formats, rule matching, report contents, duplicate handling, and refusal of bad input. This is the fast loop; use it for anything that isn't about how Codex phrases things.

```bash
npm run check           # typecheck
npm run build           # compile src/ to dist/
npm run bundle          # rebuild the committed bundle/server.mjs
npm test                # end-to-end against the real service
npm run build:template  # regenerate assets/BankStatementTemplate.xlsx
npm run verify:template # confirm the shipped template is readable and complete
```

## Getting statements in

Point at files, or at a folder:

```text
statementPaths: ["C:\\Clients\\Acme\\2026-09"]                    # every statement in the folder
statementPaths: ["C:\\...\\sept.pdf", "C:\\...\\amex.csv"]        # specific files
```

A folder expands to the statement files directly inside it (not subfolders). Recognised extensions: `.pdf`, `.csv`, `.xlsx`, `.xls`, `.ofx`, `.qfx`.

| Format | How it is read | Trust level |
| --- | --- | --- |
| OFX / QFX | Tagged fields, exact | `Parsed` |
| CSV | Header matched by name | `Parsed` |
| XLSX / XLS | First worksheet, header matched by name | `Parsed` |
| PDF with a text layer | Text positions clustered into rows and columns | `Parsed (PDF text layer)` — verify against statement totals |
| Scanned / image-only PDF | Not machine-readable. Codex reads it and posts rows via `bank_statement_add_transactions` | `Parsed (assistant-extracted)` — verify every row |

Every row records how it was read, so the Audit sheet can tell you exactly how much of the output is machine-certain. Nothing is silently upgraded from one level to another. **No OCR is bundled** — a scan goes down the assistant-extracted path, or you supply a text export.

## Getting the output out

The result is a new `.xlsx` at the `outputPath` you name. The input workbook is never modified — the tool refuses if `outputPath` equals `templatePath`.

Because the output workbook contains all nine sheets plus the accumulated master table, **it is itself a valid template**. That is the intended monthly rhythm:

```text
Acme-Master.xlsx  + September statements  ->  Acme-2026-09.xlsx
Acme-2026-09.xlsx + October statements    ->  Acme-2026-10.xlsx   (running ledger, duplicates skipped)
```

Sharing and storage are yours to choose — the plugin only writes a file. Write the output straight into a OneDrive, SharePoint, Dropbox, or Google Drive folder and it syncs and shares like any other workbook. Keep the input statements outside your source repository.

## Changing the template

Two paths, same workbook. Use whichever suits the moment.

**In Excel.** Open the workbook and edit the shaded cells. `Category Rules`, `Report Setup`, and the value columns are dropdown-validated, so a non-technical user cannot enter a value the service will reject.

**By asking Codex.** `bank_statement_update_template` edits the same sheets from a request:

> "Add a rule: anything containing STRIPE is Revenue / Card settlements, P&L group Revenue."
> "Group the consolidated statement by P&L group instead of Category."
> "Turn off the debit statement for this client."
> "Drop the UBER rule."

Values are validated against the same vocabulary the dropdowns offer; an unrecognised value is rejected rather than written. Rules apply at import time, so after changing them re-run the consolidation to reclassify existing rows.

### Category Rules

Active rules apply from the lowest `Priority` upward and the first match wins. Two active rules at the *same* priority that both match are an exception: the row stays `Uncategorized` and is flagged, rather than being assigned arbitrarily.

`Match type` accepts `contains`, `equals`, `starts_with`, `regex`, and `amount_range` (as `low..high`). `Look in this field` accepts `Description` or `Amount`.

### Report Setup

Each row refreshes the sheet whose name matches `Report name`:

| Column | Accepted values |
| --- | --- |
| `Include transactions when` | `All transactions`, `Debit > 0`, `Credit > 0`, `P&L group is not blank`, `Category is not blank` |
| `Group by` | `Category`, `Subcategory`, `P&L group`, `Currency`, `Account ID`, `Source file`, `Transaction date` |
| `Measure` | `Sum of Amount`, `Sum of Debit`, `Sum of Credit`, `Count of transactions` |
| `Show this report?` | `Yes`, `No` |

A value outside these lists is not guessed at — the report is skipped and the reason written to `Audit`.

## The workbook

| Sheet | Who fills it | Purpose |
| --- | --- | --- |
| `Start Here` | — | Plain-language instructions. |
| `Template Columns` | You | Documents what each master column means. |
| `Category Rules` | You | Assigns Category, Subcategory, and P&L group. |
| `Report Setup` | You | Controls which reports are built, and how. |
| `Master Transactions` | Service | Source of truth. Appended to, never rewritten. |
| `Debit statement` | Service | Rebuilt from the master table on every import. |
| `P&L statement` | Service | Rebuilt from the master table on every import. |
| `Consolidated statement` | Service | Rebuilt from the master table on every import. |
| `Audit` | Service | Counts, provenance, and every exception. Terminal — nothing reads from it. |

## Duplicate handling

Every row carries a SHA-256 `Transaction fingerprint` over account, date, amount, currency, description, source file, and source row. A row is skipped if that fingerprint already exists in `Master Transactions` or appeared earlier in the same import. Re-running the same statements appends nothing.

## Limits

Statement layouts vary by institution and period. PDF column detection depends on the header row being present and the text layer being sane; a heavily styled or multi-column statement can still confuse it, which is why PDF rows are marked for verification. Do not use unreviewed output for tax filings, payments, or regulated reporting.

## Development

The shipped template is generated by the same `createTemplateWorkbook()` that `bank_statement_create_template` calls, so the asset and a freshly created workbook cannot drift apart. Regenerate and re-verify after any change to `src/template.ts`.

Test fixtures are synthetic. `.gitignore` blocks `*.pdf`, `*.csv`, `*.xlsx`, `*.ofx`, and `*.qfx` by default and allowlists only the template and the fixtures, so a real statement cannot be committed by accident.

## Before distributing outside a trusted team

1. Add public privacy-policy and terms URLs to `.codex-plugin/plugin.json`.
2. Add fixtures for each institution and layout you intend to support. Synthetic or irreversibly redacted statements only.
3. Add an OCR path if you need dependable scanned-PDF extraction without a human in the loop.
4. Validate the plugin and test it in a fresh Codex thread before publishing.

## License

MIT. See [LICENSE](LICENSE).
