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
const fixture = (name) => join(root, "tests", "fixtures", name);

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
  if (result.isError) throw new Error(result.content.map((item) => (item.type === "text" ? item.text : "")).join("\n"));
  return result.content.map((item) => (item.type === "text" ? item.text : "")).join("\n");
}

const scratch = await mkdtemp(join(tmpdir(), "bank-statement-test-"));
const client = new Client({ name: "bank-statement-consolidator-smoke-test", version: "0.1.0" });
const transport = new StdioClientTransport({ command: process.execPath, args: [join(root, "dist", "index.js")], cwd: root, stderr: "inherit" });

try {
  await client.connect(transport);

  console.log("tools");
  const { tools } = await client.listTools();
  const names = tools.map((tool) => tool.name).sort();
  check(names.join(",") === "bank_statement_consolidate,bank_statement_create_template,bank_statement_inspect_template", `three tools registered (${names.join(", ")})`);

  console.log("bank_statement_create_template");
  const createdPath = join(scratch, "Created.xlsx");
  await client.callTool({ name: "bank_statement_create_template", arguments: { outputPath: createdPath } });
  check(existsSync(createdPath), "creates a workbook file");
  const created = await readWorkbook(createdPath);
  check(created.getWorksheet("Master Transactions") !== undefined, "created workbook is readable and has Master Transactions");

  console.log("bank_statement_inspect_template");
  const shipped = join(root, "assets", "BankStatementTemplate.xlsx");
  const inspected = textOf(await client.callTool({ name: "bank_statement_inspect_template", arguments: { templatePath: shipped } }));
  check(inspected.includes("ready for structured statement consolidation"), "shipped template passes inspection");
  check(inspected.includes("3 active category rule(s)"), "reads the three default category rules");

  console.log("bank_statement_consolidate: first import");
  const firstPath = join(scratch, "Consolidated.xlsx");
  const first = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: shipped, statementPaths: [fixture("sample-statement.csv"), fixture("sample-statement.ofx")], outputPath: firstPath }
  }));
  check(first.includes("appended 6"), `appends all six transactions (${first})`);
  check(first.includes("No exceptions recorded."), "no exceptions on a clean import");

  const workbook = await readWorkbook(firstPath);
  const master = workbook.getWorksheet("Master Transactions");
  check(master.rowCount === 7, `master holds a header plus six rows (found ${master.rowCount})`);

  const categories = [];
  master.eachRow({ includeEmpty: false }, (row, number) => { if (number > 1) categories.push(String(row.getCell(10).value ?? "")); });
  check(categories.filter((value) => value === "Travel").length === 2, "UBER rows in both formats matched the Travel rule");
  check(categories.filter((value) => value === "Revenue").length === 2, "CLIENT PAYMENT rows matched the Revenue rule");
  check(categories.includes("Payroll"), "PAYROLL row matched the Payroll rule");
  check(categories.includes("Uncategorized"), "an unmatched row stays Uncategorized rather than being guessed");

  const debit = workbook.getWorksheet("Debit statement");
  const debitLabels = [];
  debit.eachRow({ includeEmpty: false }, (row) => debitLabels.push(String(row.getCell(1).value ?? "")));
  check(debitLabels[0] === "Debit statement", "debit report is titled");
  check(debitLabels.includes("Total"), "debit report has a total row");
  check(!debitLabels.includes("Revenue"), "debit report excludes credit-only categories");

  const audit = workbook.getWorksheet("Audit");
  const auditLabels = [];
  audit.eachRow({ includeEmpty: false }, (row) => auditLabels.push(String(row.getCell(1).value ?? "")));
  check(auditLabels.includes("Rows in master table"), "audit reports the master row count");

  console.log("bank_statement_consolidate: re-import is idempotent");
  const secondPath = join(scratch, "Consolidated2.xlsx");
  const second = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: firstPath, statementPaths: [fixture("sample-statement.csv"), fixture("sample-statement.ofx")], outputPath: secondPath }
  }));
  check(second.includes("appended 0"), `re-importing the same statements appends nothing (${second})`);
  check(second.includes("skipped 6 duplicate(s)"), "all six rows are recognised as duplicates");
  const reimported = await readWorkbook(secondPath);
  check(reimported.getWorksheet("Master Transactions").rowCount === 7, "master table did not grow on re-import");

  console.log("unsupported input is refused, not guessed");
  const refusedPath = join(scratch, "Refused.xlsx");
  const refused = textOf(await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: shipped, statementPaths: [join(root, "README.md")], outputPath: refusedPath }
  }));
  check(refused.includes("exception(s) recorded"), "an unsupported file is recorded as an exception");
  check(refused.includes("appended 0"), "nothing is appended from an unsupported file");

  console.log(`\nAll ${passed} checks passed.`);
} finally {
  await client.close();
  await rm(scratch, { recursive: true, force: true });
}
