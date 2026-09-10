import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import ExcelJS from "exceljs";
import { McpServer } from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import * as z from "zod/v4";
import { createTemplateWorkbook, FINGERPRINT_COLUMN, MASTER_COLUMNS, SHEET_NAMES } from "./template.js";

type Transaction = {
  transactionDate: string;
  description: string;
  debit: number;
  credit: number;
  amount: number;
  currency: string;
  accountId: string;
  sourceFile: string;
  sourceRow: string;
  category: string;
  subcategory: string;
  pnlGroup: string;
  fingerprint: string;
  parseStatus: string;
};

type Rule = {
  priority: number;
  field: string;
  matchType: string;
  matchValue: string;
  category: string;
  subcategory: string;
  pnlGroup: string;
  active: boolean;
};

type ReportSetup = {
  name: string;
  includeWhen: string;
  groupBy: string;
  measure: string;
  show: boolean;
};

const server = new McpServer({ name: "bank-statement-consolidator", version: "0.1.0" });

function clean(value: unknown): string {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value && typeof value === "object") {
    const cell = value as { text?: unknown; result?: unknown; richText?: { text: string }[] };
    if (Array.isArray(cell.richText)) return clean(cell.richText.map((part) => part.text).join(""));
    if (cell.text !== undefined) return clean(cell.text);
    if (cell.result !== undefined) return clean(cell.result);
  }
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function numberValue(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = clean(value).replace(/[$₹£€]/g, "").replace(/,/g, "");
  if (!text) return 0;
  if (/^\(.*\)$/.test(text)) return -Number(text.slice(1, -1));
  const result = Number(text);
  return Number.isFinite(result) ? result : 0;
}

function normalHeader(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function excelRowValues(row: ExcelJS.Row): string[] {
  const values: string[] = [];
  row.eachCell({ includeEmpty: true }, (cell, columnNumber) => { values[columnNumber - 1] = clean(cell.value); });
  return values;
}

function fingerprint(record: Omit<Transaction, "fingerprint">): string {
  return createHash("sha256")
    .update([record.accountId, record.transactionDate, record.amount, record.currency, record.description, record.sourceFile, record.sourceRow].join("|"))
    .digest("hex");
}

function findColumn(headers: string[], aliases: string[]): number {
  return headers.findIndex((header) => aliases.includes(normalHeader(header)));
}

function csvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"' && text[index + 1] === '"' && quoted) { cell += '"'; index += 1; continue; }
    if (char === '"') { quoted = !quoted; continue; }
    if (char === "," && !quoted) { row.push(cell); cell = ""; continue; }
    if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && text[index + 1] === "\n") index += 1;
      row.push(cell); if (row.some((item) => clean(item))) rows.push(row); row = []; cell = ""; continue;
    }
    cell += char;
  }
  row.push(cell); if (row.some((item) => clean(item))) rows.push(row);
  return rows;
}

function normalizeRows(rows: string[][], sourceFile: string): Transaction[] {
  const headerAt = rows.findIndex((row) => {
    const values = row.map(normalHeader);
    return values.some((item) => ["date", "transactiondate", "valuedate", "postingdate"].includes(item))
      && values.some((item) => ["description", "narration", "details", "particulars", "memo"].includes(item));
  });
  if (headerAt < 0) throw new Error(`${sourceFile}: no Date and Description header row was found.`);
  const headers = rows[headerAt];
  const dateIndex = findColumn(headers, ["date", "transactiondate", "valuedate", "postingdate"]);
  const descriptionIndex = findColumn(headers, ["description", "narration", "details", "particulars", "memo"]);
  const debitIndex = findColumn(headers, ["debit", "withdrawal", "withdrawals"]);
  const creditIndex = findColumn(headers, ["credit", "deposit", "deposits"]);
  const amountIndex = findColumn(headers, ["amount", "transactionamount"]);
  const currencyIndex = findColumn(headers, ["currency", "curr"]);
  const accountIndex = findColumn(headers, ["accountid", "accountnumber", "account"]);
  if (debitIndex < 0 && creditIndex < 0 && amountIndex < 0) {
    throw new Error(`${sourceFile}: no Debit, Credit, or Amount column was found, so transaction values cannot be read.`);
  }
  return rows.slice(headerAt + 1).map((row, offset) => {
    const debit = debitIndex >= 0 ? Math.abs(numberValue(row[debitIndex])) : 0;
    const credit = creditIndex >= 0 ? Math.abs(numberValue(row[creditIndex])) : 0;
    const suppliedAmount = amountIndex >= 0 ? numberValue(row[amountIndex]) : 0;
    // A signed Amount column wins when the source has no explicit debit/credit split.
    const amount = debitIndex >= 0 || creditIndex >= 0 ? credit - debit : suppliedAmount;
    const derivedDebit = debitIndex >= 0 || creditIndex >= 0 ? debit : Math.max(0, -suppliedAmount);
    const derivedCredit = debitIndex >= 0 || creditIndex >= 0 ? credit : Math.max(0, suppliedAmount);
    const partial = {
      transactionDate: clean(row[dateIndex]),
      description: clean(row[descriptionIndex]),
      debit: derivedDebit,
      credit: derivedCredit,
      amount,
      currency: currencyIndex >= 0 ? clean(row[currencyIndex]) || "Unknown" : "Unknown",
      accountId: accountIndex >= 0 ? clean(row[accountIndex]) : "",
      sourceFile,
      sourceRow: String(headerAt + offset + 2),
      category: "Uncategorized",
      subcategory: "",
      pnlGroup: "",
      parseStatus: amount === 0 && !suppliedAmount ? "Review: no debit or credit value was read" : "Parsed"
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  }).filter((row) => row.transactionDate || row.description || row.amount !== 0);
}

function parseOfx(text: string, sourceFile: string): Transaction[] {
  const accountId = clean(text.match(/<ACCTID>([^<\r\n]+)/i)?.[1]);
  const currency = clean(text.match(/<CURDEF>([^<\r\n]+)/i)?.[1]) || "Unknown";
  const blocks = text.match(/<STMTTRN>[\s\S]*?<\/STMTTRN>/gi) ?? [];
  if (!blocks.length) throw new Error(`${sourceFile}: no OFX transaction blocks were found.`);
  return blocks.map((block, index) => {
    const value = (tag: string) => clean(block.match(new RegExp(`<${tag}>([^<\\r\\n]+)`, "i"))?.[1]);
    const amount = numberValue(value("TRNAMT"));
    const partial = {
      transactionDate: value("DTPOSTED").slice(0, 8),
      description: value("NAME") || value("MEMO"),
      debit: amount < 0 ? Math.abs(amount) : 0,
      credit: amount > 0 ? amount : 0,
      amount,
      currency,
      accountId,
      sourceFile,
      sourceRow: String(index + 1),
      category: "Uncategorized",
      subcategory: "",
      pnlGroup: "",
      parseStatus: "Parsed"
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  });
}

async function parseStructuredFile(filePath: string): Promise<Transaction[]> {
  if (!existsSync(filePath)) throw new Error(`File does not exist: ${filePath}`);
  const extension = extname(filePath).toLowerCase();
  if (extension === ".csv") return normalizeRows(csvRows(await readFile(filePath, "utf8")), basename(filePath));
  if ([".xlsx", ".xls"].includes(extension)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error(`${basename(filePath)}: the workbook has no worksheets.`);
    const rows: string[][] = [];
    sheet.eachRow({ includeEmpty: false }, (row) => rows.push(excelRowValues(row)));
    return normalizeRows(rows, basename(filePath));
  }
  if ([".ofx", ".qfx"].includes(extension)) return parseOfx(await readFile(filePath, "utf8"), basename(filePath));
  throw new Error(`${basename(filePath)}: unsupported for automatic import. Use CSV, XLSX/XLS, or OFX/QFX, or ask your assistant to extract the PDF for review.`);
}

function sheetRows(sheet: ExcelJS.Worksheet): Record<string, string>[] {
  const headers = excelRowValues(sheet.getRow(1));
  const result: Record<string, string>[] = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const record: Record<string, string> = {};
    headers.forEach((header, index) => { record[header] = clean(row.getCell(index + 1).value); });
    if (Object.values(record).some(Boolean)) result.push(record);
  });
  return result;
}

function getSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`Template is missing the '${name}' sheet. Start from BankStatementTemplate.xlsx or add this sheet.`);
  return sheet;
}

function isYes(value: string | undefined): boolean {
  return /^(yes|true|y|1)$/i.test((value ?? "").trim());
}

function readRules(workbook: ExcelJS.Workbook): Rule[] {
  return sheetRows(getSheet(workbook, "Category Rules")).map((row) => ({
    priority: Number(row["Priority"]) || 9999,
    field: row["Look in this field"] || "Description",
    matchType: (row["Match type"] || "contains").toLowerCase(),
    matchValue: row["Text or amount to match"] || "",
    category: row["Category"] || "Uncategorized",
    subcategory: row["Subcategory"] || "",
    pnlGroup: row["P&L group"] || "",
    active: isYes(row["Use this rule?"])
  })).filter((rule) => rule.active && rule.matchValue).sort((left, right) => left.priority - right.priority);
}

function readReportSetups(workbook: ExcelJS.Workbook): ReportSetup[] {
  return sheetRows(getSheet(workbook, "Report Setup")).map((row) => ({
    name: row["Report name"] || "",
    includeWhen: (row["Include transactions when"] || "All transactions").trim().toLowerCase(),
    groupBy: (row["Group by"] || "Category").trim().toLowerCase(),
    measure: (row["Measure"] || "Sum of Amount").trim().toLowerCase(),
    show: isYes(row["Show this report?"])
  })).filter((setup) => setup.name);
}

function applyRules(row: Transaction, rules: Rule[]): Transaction {
  const matches = rules.filter((rule) => {
    const fieldValue = rule.field.toLowerCase().includes("amount") ? String(row.amount) : row.description;
    if (rule.matchType === "equals") return fieldValue.toLowerCase() === rule.matchValue.toLowerCase();
    if (rule.matchType === "starts_with") return fieldValue.toLowerCase().startsWith(rule.matchValue.toLowerCase());
    if (rule.matchType === "regex") { try { return new RegExp(rule.matchValue, "i").test(fieldValue); } catch { return false; } }
    if (rule.matchType === "amount_range") { const [low, high] = rule.matchValue.split("..").map(numberValue); return row.amount >= low && row.amount <= high; }
    return fieldValue.toLowerCase().includes(rule.matchValue.toLowerCase());
  });
  if (!matches.length) return row;
  const first = matches[0];
  // Two active rules at the same priority are an exception, not a coin flip.
  if (matches.filter((rule) => rule.priority === first.priority).length > 1) {
    return { ...row, parseStatus: `Review: rules at priority ${first.priority} both match` };
  }
  return { ...row, category: first.category, subcategory: first.subcategory, pnlGroup: first.pnlGroup };
}

const INCLUDE_FILTERS: Record<string, (row: Transaction) => boolean> = {
  "all transactions": () => true,
  "debit > 0": (row) => row.debit > 0,
  "credit > 0": (row) => row.credit > 0,
  "p&l group is not blank": (row) => Boolean(row.pnlGroup),
  "category is not blank": (row) => Boolean(row.category)
};

const GROUP_FIELDS: Record<string, (row: Transaction) => string> = {
  "category": (row) => row.category || "Uncategorized",
  "subcategory": (row) => row.subcategory || "(blank)",
  "p&l group": (row) => row.pnlGroup || "(blank)",
  "currency": (row) => row.currency || "Unknown",
  "account id": (row) => row.accountId || "(blank)",
  "source file": (row) => row.sourceFile || "(blank)",
  "transaction date": (row) => row.transactionDate || "(blank)"
};

const MEASURES: Record<string, (row: Transaction) => number> = {
  "sum of amount": (row) => row.amount,
  "sum of debit": (row) => row.debit,
  "sum of credit": (row) => row.credit,
  "count of transactions": () => 1
};

function readMasterRows(sheet: ExcelJS.Worksheet): Transaction[] {
  const rows: Transaction[] = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const at = (column: number) => clean(row.getCell(column).value);
    const record: Transaction = {
      transactionDate: at(1), description: at(2),
      debit: numberValue(row.getCell(3).value), credit: numberValue(row.getCell(4).value), amount: numberValue(row.getCell(5).value),
      currency: at(6) || "Unknown", accountId: at(7), sourceFile: at(8), sourceRow: at(9),
      category: at(10) || "Uncategorized", subcategory: at(11), pnlGroup: at(12),
      fingerprint: at(FINGERPRINT_COLUMN), parseStatus: at(14) || "Parsed"
    };
    if (record.fingerprint || record.description || record.transactionDate) rows.push(record);
  });
  return rows;
}

function writeMasterRow(sheet: ExcelJS.Worksheet, transaction: Transaction): void {
  sheet.addRow([
    transaction.transactionDate, transaction.description, transaction.debit, transaction.credit, transaction.amount,
    transaction.currency, transaction.accountId, transaction.sourceFile, transaction.sourceRow,
    transaction.category, transaction.subcategory, transaction.pnlGroup, transaction.fingerprint, transaction.parseStatus
  ]);
}

function buildReport(sheet: ExcelJS.Worksheet, setup: ReportSetup, rows: Transaction[]): number {
  const include = INCLUDE_FILTERS[setup.includeWhen];
  const groupBy = GROUP_FIELDS[setup.groupBy];
  const measure = MEASURES[setup.measure];
  const selected = rows.filter(include);
  const groups = new Map<string, number>();
  selected.forEach((row) => groups.set(groupBy(row), (groups.get(groupBy(row)) ?? 0) + measure(row)));

  sheet.spliceRows(1, sheet.rowCount);
  sheet.addRow([setup.name]).font = { name: "Arial", size: 14, bold: true, color: { argb: "FF12302E" } };
  sheet.addRow([`Includes: ${setup.includeWhen}. Grouped by: ${setup.groupBy}. Measure: ${setup.measure}.`]).font = { name: "Arial", size: 9, italic: true };
  sheet.addRow([]);
  const header = sheet.addRow([setup.groupBy, setup.measure]);
  header.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
  [...groups.entries()].sort(([left], [right]) => left.localeCompare(right)).forEach(([group, value]) => sheet.addRow([group, value]));
  const total = sheet.addRow(["Total", [...groups.values()].reduce((sum, value) => sum + value, 0)]);
  total.font = { name: "Arial", size: 10, bold: true };
  total.border = { top: { style: "thin" } };
  sheet.getColumn(1).width = 46;
  sheet.getColumn(2).width = 18;
  sheet.getColumn(2).numFmt = setup.measure === "count of transactions" ? "#,##0" : "#,##0.00;(#,##0.00);-";
  sheet.views = [{ showGridLines: false }];
  return selected.length;
}

function writeAudit(sheet: ExcelJS.Worksheet, lines: [string, string | number][], exceptions: string[]): void {
  sheet.spliceRows(1, sheet.rowCount);
  sheet.addRow(["Import audit"]).font = { name: "Arial", size: 14, bold: true, color: { argb: "FF12302E" } };
  sheet.addRow([`Generated ${new Date().toISOString()}`]).font = { name: "Arial", size: 9, italic: true };
  sheet.addRow([]);
  lines.forEach(([label, value]) => {
    const row = sheet.addRow([label, value]);
    row.getCell(1).font = { name: "Arial", size: 10, bold: true };
  });
  sheet.addRow([]);
  sheet.addRow(["Exceptions and notes"]).font = { name: "Arial", size: 11, bold: true };
  (exceptions.length ? exceptions : ["None."]).forEach((line) => sheet.addRow([line]));
  sheet.getColumn(1).width = 62;
  sheet.getColumn(2).width = 18;
  sheet.views = [{ showGridLines: false }];
}

server.registerTool("bank_statement_create_template", {
  description: "Create an editable, plain-language bank-statement template workbook.",
  inputSchema: z.object({ outputPath: z.string().describe("Absolute path for the new .xlsx workbook.") })
}, async ({ outputPath }) => {
  const target = resolve(outputPath);
  await createTemplateWorkbook().xlsx.writeFile(target);
  return { content: [{ type: "text", text: `Created template: ${target}` }] };
});

server.registerTool("bank_statement_inspect_template", {
  description: "Check that a client workbook has the sheets, master columns, and rules needed for template-driven consolidation.",
  inputSchema: z.object({ templatePath: z.string().describe("Absolute path to the client template workbook.") })
}, async ({ templatePath }) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolve(templatePath));
  const missingSheets = SHEET_NAMES.filter((name) => !workbook.getWorksheet(name));
  if (missingSheets.length) {
    return { content: [{ type: "text", text: `Template needs these sheets: ${missingSheets.join(", ")}` }] };
  }
  const headers = excelRowValues(getSheet(workbook, "Master Transactions").getRow(1)).map(normalHeader);
  const missingColumns = MASTER_COLUMNS.filter((column) => !headers.includes(normalHeader(column)));
  const rules = readRules(workbook);
  const reports = readReportSetups(workbook);
  const unknown = reports.filter((setup) => setup.show).flatMap((setup) => [
    ...(INCLUDE_FILTERS[setup.includeWhen] ? [] : [`${setup.name}: unknown "Include transactions when" value "${setup.includeWhen}"`]),
    ...(GROUP_FIELDS[setup.groupBy] ? [] : [`${setup.name}: unknown "Group by" value "${setup.groupBy}"`]),
    ...(MEASURES[setup.measure] ? [] : [`${setup.name}: unknown "Measure" value "${setup.measure}"`])
  ]);
  const problems = [
    ...(missingColumns.length ? [`Master Transactions is missing: ${missingColumns.join(", ")}`] : []),
    ...(rules.length ? [] : ["Category Rules has no active rule, so every transaction will be Uncategorized."]),
    ...unknown
  ];
  const summary = `Sheets present. ${rules.length} active category rule(s), ${reports.filter((setup) => setup.show).length} report(s) turned on.`;
  return { content: [{ type: "text", text: problems.length ? `${summary}\nNeeds attention:\n- ${problems.join("\n- ")}` : `${summary} Template is ready for structured statement consolidation.` }] };
});

server.registerTool("bank_statement_consolidate", {
  description: "Read CSV, XLSX/XLS, or OFX/QFX statements, apply the workbook's category rules, append non-duplicate transactions to the master table, refresh the reports named in Report Setup, and save a consolidated workbook.",
  inputSchema: z.object({
    templatePath: z.string().describe("Absolute path to the client template workbook."),
    statementPaths: z.array(z.string()).min(1).describe("Absolute paths to CSV, XLSX/XLS, OFX, or QFX statement files."),
    outputPath: z.string().describe("Absolute path for the resulting consolidated .xlsx workbook.")
  })
}, async ({ templatePath, statementPaths, outputPath }) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolve(templatePath));
  const master = getSheet(workbook, "Master Transactions");
  const rules = readRules(workbook);
  const reports = readReportSetups(workbook);
  const exceptions: string[] = [];

  const existingRows = readMasterRows(master);
  const seen = new Set(existingRows.map((row) => row.fingerprint).filter(Boolean));

  const parsed: Transaction[] = [];
  for (const statementPath of statementPaths) {
    try {
      parsed.push(...(await parseStructuredFile(resolve(statementPath))).map((row) => applyRules(row, rules)));
    } catch (error) {
      exceptions.push(`Not imported - ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  // Deduplicate within this import as well as against the existing master table.
  const added: Transaction[] = [];
  for (const row of parsed) {
    if (seen.has(row.fingerprint)) continue;
    seen.add(row.fingerprint);
    added.push(row);
    writeMasterRow(master, row);
  }

  // Reports always describe the whole master table, not just this import.
  const allRows = [...existingRows, ...added];
  for (const setup of reports) {
    const sheet = workbook.getWorksheet(setup.name);
    if (!sheet) { exceptions.push(`Report "${setup.name}" has no sheet with that name, so it was skipped.`); continue; }
    if (!setup.show) { exceptions.push(`Report "${setup.name}" is turned off in Report Setup, so it was left unchanged.`); continue; }
    if (!INCLUDE_FILTERS[setup.includeWhen] || !GROUP_FIELDS[setup.groupBy] || !MEASURES[setup.measure]) {
      exceptions.push(`Report "${setup.name}" uses a value this version does not understand, so it was skipped. Pick from the dropdowns in Report Setup.`);
      continue;
    }
    buildReport(sheet, setup, allRows);
  }

  const needsReview = added.filter((row) => row.parseStatus !== "Parsed");
  needsReview.forEach((row) => exceptions.push(`${row.sourceFile} row ${row.sourceRow}: ${row.parseStatus}`));
  const uncategorized = added.filter((row) => row.category === "Uncategorized").length;

  writeAudit(getSheet(workbook, "Audit"), [
    ["Statements requested", statementPaths.length],
    ["Statements read", statementPaths.length - exceptions.filter((line) => line.startsWith("Not imported")).length],
    ["Transactions extracted", parsed.length],
    ["Rows appended", added.length],
    ["Duplicates skipped", parsed.length - added.length],
    ["Rows in master table", allRows.length],
    ["Uncategorized in this import", uncategorized],
    ["Rows needing review", needsReview.length],
    ["Active category rules", rules.length]
  ], exceptions);

  const target = resolve(outputPath);
  await workbook.xlsx.writeFile(target);
  const text = [
    `Created ${target}.`,
    `Extracted ${parsed.length}; appended ${added.length}; skipped ${parsed.length - added.length} duplicate(s).`,
    `Master table now holds ${allRows.length} transaction(s). ${uncategorized} uncategorized, ${needsReview.length} needing review.`,
    exceptions.length ? `${exceptions.length} exception(s) recorded on the Audit sheet. Review it before using the output.` : "No exceptions recorded."
  ].join(" ");
  return { content: [{ type: "text", text }] };
});

await server.connect(new StdioServerTransport());
