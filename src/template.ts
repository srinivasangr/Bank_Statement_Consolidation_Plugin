import ExcelJS from "exceljs";

export const SHEET_NAMES = [
  "Start Here",
  "Template Columns",
  "Category Rules",
  "Report Setup",
  "Master Transactions",
  "Debit statement",
  "P&L statement",
  "Consolidated statement",
  "Audit"
] as const;

export const MASTER_COLUMNS = [
  "Transaction date", "Description", "Debit", "Credit", "Amount", "Currency", "Account ID",
  "Source file", "Source row", "Category", "Subcategory", "P&L group", "Transaction fingerprint", "Parse status"
] as const;

export const FINGERPRINT_COLUMN = MASTER_COLUMNS.indexOf("Transaction fingerprint") + 1;

const TEAL = "FF0F766E";
const AMBER = "FFFEF3C7";
const INK = "FF12302E";
const MONEY_FORMAT = "#,##0.00;(#,##0.00);-";

// Excel table objects are deliberately not used. The consolidation tool reads this
// workbook, rewrites the report sheets, and appends master rows; ExcelJS does not carry
// table definitions through that round trip reliably. Styled headers plus an autofilter
// give the same usability without risking a corrupt file on the second import.
function headerRow(sheet: ExcelJS.Worksheet, labels: readonly string[]): void {
  const row = sheet.addRow([...labels]);
  row.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TEAL } };
  row.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  row.height = 30;
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: labels.length } };
}

function editable(sheet: ExcelJS.Worksheet, firstRow: number, lastRow: number, lastColumn: number): void {
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      sheet.getRow(rowNumber).getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMBER } };
    }
  }
}

function dropdown(sheet: ExcelJS.Worksheet, column: number, options: string[], lastRow: number): void {
  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber += 1) {
    sheet.getRow(rowNumber).getCell(column).dataValidation = {
      type: "list", allowBlank: true, formulae: [`"${options.join(",")}"`],
      showErrorMessage: true, errorTitle: "Not an allowed value", error: `Choose one of: ${options.join(", ")}`
    };
  }
}

function widths(sheet: ExcelJS.Worksheet, values: number[]): void {
  values.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
}

export function createTemplateWorkbook(): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Bank Statement Consolidator";
  workbook.created = new Date();

  const start = workbook.addWorksheet("Start Here", { views: [{ showGridLines: false }] });
  start.addRow(["Bank Statement Consolidator"]).font = { name: "Arial", size: 16, bold: true, color: { argb: INK } };
  start.addRow([]);
  [
    "1. Save a copy of this workbook for the client.",
    "2. Edit the shaded sheets so they match the client's categories and reporting needs: Template Columns, Category Rules, Report Setup.",
    "3. Give this workbook and the statement files to your assistant, and ask: Consolidate these statements using this template.",
    "4. Review the Audit sheet before using the output.",
    "",
    "Shaded cells are yours to edit. The unshaded sheets are filled in for you on each import.",
    "Master Transactions is the source of truth; the three statement sheets are rebuilt from it every time.",
    "Keep real bank statements out of the plugin source repository."
  ].forEach((line) => {
    const row = start.addRow([line]);
    row.alignment = { wrapText: true, vertical: "top" };
    row.font = { name: "Arial", size: 10 };
  });
  widths(start, [110]);

  const columns = workbook.addWorksheet("Template Columns");
  headerRow(columns, ["Output column", "What it means", "Required?", "Example", "How it is filled"]);
  columns.addRows([
    ["Transaction date", "Date the bank posted the transaction", "Yes", "2026-09-10", "From statement"],
    ["Description", "Bank transaction description", "Yes", "CARD PURCHASE", "From statement"],
    ["Debit", "Money leaving the account", "Yes", 125.5, "From statement"],
    ["Credit", "Money entering the account", "Yes", 0, "From statement"],
    ["Amount", "Credit minus debit", "Yes", -125.5, "Calculated or from statement"],
    ["Currency", "Transaction currency", "Yes", "USD", "From statement or default"],
    ["Account ID", "Masked account identifier", "Yes", "XXXX1234", "From statement"],
    ["Source file", "Original statement filename", "Yes", "September.csv", "Filled automatically"],
    ["Source row", "Row or page location in the source", "Yes", "27", "Filled automatically"],
    ["Category", "Category assigned by your rules", "Yes", "Travel", "Category Rules"],
    ["Subcategory", "Optional detailed category", "No", "Airfare", "Category Rules"],
    ["P&L group", "Profit-and-loss reporting group", "No", "Operating expense", "Category Rules"],
    ["Transaction fingerprint", "Duplicate-protection identifier", "Yes", "SHA-256", "Filled automatically"],
    ["Parse status", "Parsed, or the reason a row needs review", "Yes", "Parsed", "Filled automatically"]
  ]);
  editable(columns, 2, columns.rowCount, 5);
  widths(columns, [25, 40, 12, 20, 28]);

  const rules = workbook.addWorksheet("Category Rules");
  headerRow(rules, ["Priority", "Look in this field", "Match type", "Text or amount to match", "Category", "Subcategory", "P&L group", "Use this rule?"]);
  rules.addRows([
    [10, "Description", "contains", "UBER", "Travel", "Rideshare", "Operating expense", "Yes"],
    [20, "Description", "contains", "PAYROLL", "Payroll", "Wages", "Operating expense", "Yes"],
    [30, "Description", "contains", "CLIENT PAYMENT", "Revenue", "Client receipts", "Revenue", "Yes"]
  ]);
  editable(rules, 2, 100, 8);
  dropdown(rules, 2, ["Description", "Amount"], 100);
  dropdown(rules, 3, ["contains", "equals", "starts_with", "regex", "amount_range"], 100);
  dropdown(rules, 8, ["Yes", "No"], 100);
  widths(rules, [10, 20, 18, 30, 22, 22, 22, 15]);

  const reports = workbook.addWorksheet("Report Setup");
  headerRow(reports, ["Report name", "Include transactions when", "Group by", "Measure", "Show this report?"]);
  reports.addRows([
    ["Debit statement", "Debit > 0", "Category", "Sum of Debit", "Yes"],
    ["P&L statement", "P&L group is not blank", "P&L group", "Sum of Amount", "Yes"],
    ["Consolidated statement", "All transactions", "Category", "Sum of Amount", "Yes"]
  ]);
  editable(reports, 2, 4, 5);
  dropdown(reports, 2, ["All transactions", "Debit > 0", "Credit > 0", "P&L group is not blank", "Category is not blank"], 4);
  dropdown(reports, 3, ["Category", "Subcategory", "P&L group", "Currency", "Account ID", "Source file", "Transaction date"], 4);
  dropdown(reports, 4, ["Sum of Amount", "Sum of Debit", "Sum of Credit", "Count of transactions"], 4);
  dropdown(reports, 5, ["Yes", "No"], 4);
  widths(reports, [24, 30, 20, 24, 18]);

  const master = workbook.addWorksheet("Master Transactions");
  headerRow(master, MASTER_COLUMNS);
  widths(master, [15, 40, 13, 13, 13, 10, 15, 22, 12, 20, 20, 22, 36, 26]);
  [3, 4, 5].forEach((column) => { master.getColumn(column).numFmt = MONEY_FORMAT; });

  for (const name of ["Debit statement", "P&L statement", "Consolidated statement", "Audit"] as const) {
    const sheet = workbook.addWorksheet(name, { views: [{ showGridLines: false }] });
    sheet.addRow([name]).font = { name: "Arial", size: 14, bold: true, color: { argb: INK } };
    sheet.addRow([]);
    sheet.addRow(["This sheet is rebuilt from Master Transactions on every import. Do not edit it by hand."]).font = { name: "Arial", size: 10, italic: true };
    widths(sheet, [46, 18]);
  }

  return workbook;
}
