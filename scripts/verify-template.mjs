// Confirms the shipped template is readable by the same library the MCP service
// uses, and that every sheet, master column, rule, and report it needs is present.
// The first release shipped a workbook this check would have rejected.
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";

const templatePath = fileURLToPath(new URL("../assets/BankStatementTemplate.xlsx", import.meta.url));
const { MASTER_COLUMNS, SHEET_NAMES } = await import(new URL("../dist/template.js", import.meta.url));

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(templatePath);

const problems = [];
const missingSheets = SHEET_NAMES.filter((name) => !workbook.getWorksheet(name));
if (missingSheets.length) problems.push(`Missing sheets: ${missingSheets.join(", ")}`);

const master = workbook.getWorksheet("Master Transactions");
if (master) {
  const headers = [];
  master.getRow(1).eachCell({ includeEmpty: true }, (cell, column) => { headers[column - 1] = String(cell.value ?? "").trim(); });
  const missingColumns = MASTER_COLUMNS.filter((column) => !headers.includes(column));
  if (missingColumns.length) problems.push(`Master Transactions is missing: ${missingColumns.join(", ")}`);
}

const rules = workbook.getWorksheet("Category Rules");
if (rules && rules.rowCount < 2) problems.push("Category Rules has no example rows.");

const reports = workbook.getWorksheet("Report Setup");
if (reports && reports.rowCount < 4) problems.push("Report Setup does not define the three default reports.");

console.log(`Sheets: ${workbook.worksheets.map((sheet) => sheet.name).join(" | ")}`);
if (problems.length) {
  console.error(`\nTemplate verification failed:\n- ${problems.join("\n- ")}`);
  process.exit(1);
}
console.log("\nTemplate verification passed.");
