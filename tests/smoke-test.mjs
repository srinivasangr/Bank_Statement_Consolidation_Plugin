// End-to-end check against the real MCP service over stdio.
// Run with: npm test
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const fixtures = join(root, "tests", "fixtures");
const fixture = (name) => join(fixtures, name);
const shipped = join(root, "assets", "BankStatementTemplate.xlsx");

let passed = 0;
function check(condition, label) {
  if (!condition) throw new Error(`FAILED: ${label}`);
  passed += 1;
  console.log(`  ok  ${label}`);
}

async function readWorkbook(path) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path);
  return workbook;
}

function textOf(result) {
  const text = result.content.map((item) => (item.type === "text" ? item.text : "")).join("\n");
  if (result.isError) throw new Error(text);
  return text;
}

function columnValues(sheet, column) {
  const values = [];
  sheet.eachRow({ includeEmpty: false }, (row, number) => { if (number > 1) values.push(String(row.getCell(column).value ?? "")); });
  return values;
}

const scratch = await mkdtemp(join(tmpdir(), "bank-statement-test-"));
const client = new Client({ name: "bank-statement-consolidator-smoke-test", version: "0.1.0" });
const transport = new StdioClientTransport({ command: process.execPath, args: [join(root, "dist", "index.js")], cwd: root, stderr: "inherit" });

try {
  await client.connect(transport);

  console.log("tools");
  const { tools } = await client.listTools();
  const names = tools.map((tool) => tool.name).sort();
  check(names.length === 5, `five tools registered (${names.join(", ")})`);
  check(names.includes("bank_statement_update_template"), "template can be edited by prompt");
  check(names.includes("bank_statement_add_transactions"), "assistant-extraction fallback is available");

  console.log("bank_statement_create_template");
  const createdPath = join(scratch, "Created.xlsx");
  await client.callTool({ name: "bank_statement_create_template", arguments: { outputPath: createdPath } });
  const created = await readWorkbook(createdPath);
  check(created.getWorksheet("Master Transactions") !== undefined, "created workbook is readable and has Master Transactions");

  console.log("bank_statement_inspect_template");
  const inspected = textOf(await client.callTool({ name: "bank_statement_inspect_template", arguments: { templatePath: shipped } }));
  check(inspected.includes("ready for consolidation"), "shipped template passes inspection");
  check(inspected.includes("3 active category rule(s)"), "reads the three default category rules");

  console.log("bank_statement_consolidate: a folder of CSV, OFX and PDF statements");
  const firstPath = join(scratch, "Consolidated.xlsx");
  const first = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: shipped, statementPaths: [fixtures], outputPath: firstPath }
  }));
  check(first.includes("Read 3 file(s)"), `a folder path expands to its statement files (${first.split(".")[1]})`);
  check(first.includes("appended 11"), `appends all eleven transactions across three formats (${first})`);

  const workbook = await readWorkbook(firstPath);
  const master = workbook.getWorksheet("Master Transactions");
  check(master.rowCount === 12, `master holds a header plus eleven rows (found ${master.rowCount})`);

  const categories = columnValues(master, 10);
  check(categories.filter((value) => value === "Travel").length === 3, "UBER rows matched Travel in all three formats");
  check(categories.filter((value) => value === "Revenue").length === 3, "CLIENT PAYMENT rows matched Revenue in all three formats");
  check(categories.filter((value) => value === "Payroll").length === 2, "PAYROLL rows matched Payroll");
  check(categories.filter((value) => value === "Uncategorized").length === 3, "unmatched rows stay Uncategorized rather than being guessed");

  const statuses = columnValues(master, 14);
  check(statuses.filter((value) => value === "Parsed (PDF text layer)").length === 5, "PDF rows are recorded with their provenance, not as plain Parsed");
  check(statuses.filter((value) => value === "Parsed").length === 6, "CSV and OFX rows are recorded as fully parsed");

  const descriptions = columnValues(master, 2);
  check(descriptions.includes("CITY WATER UTILITY"), "PDF description column was reconstructed from text positions");
  const pdfRow = descriptions.indexOf("CITY WATER UTILITY");
  check(columnValues(master, 3)[pdfRow] === "88.4", `PDF withdrawal column landed in Debit (got ${columnValues(master, 3)[pdfRow]})`);
  check(columnValues(master, 4)[pdfRow] === "0", "PDF balance column was not mistaken for a credit");

  const dates = columnValues(master, 1);
  check(dates.includes("2026-09-04"), `OFX YYYYMMDD dates are normalised to ISO (got ${dates.filter((d) => d.startsWith("2026-09-0")).join(", ")})`);
  check(!dates.some((value) => /^\d{8}$/.test(value)), "no raw 8-digit OFX date reaches the master table");

  const debitSheet = workbook.getWorksheet("Debit statement");
  const debitTotal = debitSheet.getRow(debitSheet.rowCount).getCell(2).value;
  check(debitTotal === 3011.8, `report totals are rounded to cents, not floating-point dust (got ${debitTotal})`);

  const audit = workbook.getWorksheet("Audit");
  const auditText = columnValues(audit, 1).join(" | ");
  check(auditText.includes("Read from PDF text layer"), "audit breaks out how many rows came from a PDF");
  check(auditText.includes("reconstructed from a PDF text layer"), "audit warns that PDF rows need checking against statement totals");

  console.log("bank_statement_update_template: edit rules by prompt");
  const retunedPath = join(scratch, "Retuned.xlsx");
  const updated = textOf(await client.callTool({
    name: "bank_statement_update_template",
    arguments: {
      templatePath: shipped,
      outputPath: retunedPath,
      addRules: [
        { match: "CITY WATER", category: "Utilities", subcategory: "Water", pnlGroup: "Operating expense" },
        { match: "AMAZON", category: "Supplies", pnlGroup: "Operating expense" }
      ],
      updateReports: [{ name: "Consolidated statement", groupBy: "P&L group" }]
    }
  }));
  check(updated.includes("Added 2 rule(s)"), "adds rules from a prompt");
  check(updated.includes("5 active rule(s) now"), `rule count grows from three to five (${updated})`);

  const reappliedPath = join(scratch, "Reapplied.xlsx");
  textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: retunedPath, statementPaths: [fixtures], outputPath: reappliedPath }
  }));
  const reapplied = await readWorkbook(reappliedPath);
  const newCategories = columnValues(reapplied.getWorksheet("Master Transactions"), 10);
  check(newCategories.includes("Utilities"), "the prompt-added Utilities rule classified the PDF water bill");
  check(newCategories.includes("Supplies"), "the prompt-added Supplies rule classified the PDF Amazon row");
  check(newCategories.filter((value) => value === "Uncategorized").length === 1, "only the coffee shop remains uncategorized");
  check(String(reapplied.getWorksheet("Consolidated statement").getRow(4).getCell(1).value) === "p&l group", "the report now groups by the field set through the prompt");

  console.log("bank_statement_update_template: bad values are refused");
  let refusedRule = "";
  try {
    refusedRule = textOf(await client.callTool({
      name: "bank_statement_update_template",
      arguments: { templatePath: retunedPath, outputPath: join(scratch, "Bad.xlsx"), addRules: [{ match: "not-a-range", matchType: "amount_range", category: "X" }] }
    }));
  } catch (error) { refusedRule = String(error); }
  check(refusedRule.includes("low..high"), "an amount_range rule without low..high is rejected, not written");

  console.log("bank_statement_add_transactions: scanned-PDF fallback");
  const manualPath = join(scratch, "Manual.xlsx");
  const manual = textOf(await client.callTool({
    name: "bank_statement_add_transactions",
    arguments: {
      templatePath: firstPath,
      outputPath: manualPath,
      sourceFile: "scanned-march.pdf",
      transactions: [
        { transactionDate: "2026-03-04", description: "UBER TRIP SCAN", debit: 31.4, currency: "USD", accountId: "XXXX4417", sourceRow: "page 2 line 6" },
        { transactionDate: "2026-03-06", description: "CLIENT PAYMENT SCAN", credit: 900, currency: "USD", accountId: "XXXX4417", sourceRow: "page 2 line 7" }
      ]
    }
  }));
  check(manual.includes("appended 2"), `assistant-supplied rows are appended (${manual})`);
  const manualBook = await readWorkbook(manualPath);
  const manualStatuses = columnValues(manualBook.getWorksheet("Master Transactions"), 14);
  check(manualStatuses.filter((value) => value === "Parsed (assistant-extracted)").length === 2, "assistant rows are labelled as assistant-extracted");
  check(columnValues(manualBook.getWorksheet("Audit"), 1).join(" ").includes("supplied by the assistant"), "audit flags assistant rows for verification");
  check(columnValues(manualBook.getWorksheet("Master Transactions"), 10).filter((value) => value === "Travel").length === 4, "assistant rows still go through the category rules");

  console.log("re-import is idempotent");
  const secondPath = join(scratch, "Consolidated2.xlsx");
  const second = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: firstPath, statementPaths: [fixtures], outputPath: secondPath }
  }));
  check(second.includes("appended 0"), `re-importing the same statements appends nothing (${second})`);
  check(second.includes("skipped 11 duplicate(s)"), "all eleven rows are recognised as duplicates");

  console.log("bad input is refused, not guessed");
  const refusedPath = join(scratch, "Refused.xlsx");
  const refused = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: shipped, statementPaths: [join(root, "README.md")], outputPath: refusedPath }
  }));
  check(refused.includes("appended 0"), "nothing is appended from an unsupported file");
  check(columnValues((await readWorkbook(refusedPath)).getWorksheet("Audit"), 1).join(" ").includes("unsupported file type"), "the unsupported file is named on the audit sheet");

  let sameFile = "";
  try {
    sameFile = textOf(await client.callTool({ name: "bank_statement_consolidate", arguments: { templatePath: shipped, statementPaths: [fixture("sample-statement.csv")], outputPath: shipped } }));
  } catch (error) { sameFile = String(error); }
  check(sameFile.includes("must differ"), "writing over the input template is refused");

  console.log(`\nAll ${passed} checks passed.`);
} finally {
  await client.close();
  await rm(scratch, { recursive: true, force: true });
}
