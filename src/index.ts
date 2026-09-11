import { createHash } from "node:crypto";
import { existsSync, statSync } from "node:fs";
import { readdir, readFile } from "node:fs/promises";
import { basename, extname, join, resolve } from "node:path";
import ExcelJS from "exceljs";
import { McpServer } from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import * as z from "zod/v4";
import { looksLikeBoaBusinessCsv, parseBoaBusinessCsv } from "./boa.js";
import { extractPdfTable } from "./pdf.js";
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

const STATEMENT_EXTENSIONS = [".csv", ".xlsx", ".xls", ".ofx", ".qfx", ".pdf"];

// Parse status carries provenance. Anything starting with "Review" is an exception
// that a human must clear; the rest records how the row was read, because a row
// recovered from a PDF deserves less trust than one from an OFX export.
const STATUS_PARSED = "Parsed";
const STATUS_PDF = "Parsed (PDF text layer)";
const STATUS_ASSISTANT = "Parsed (assistant-extracted)";

function needsReview(status: string): boolean {
  return status.startsWith("Review");
}

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

function normalizeRows(rows: string[][], sourceFile: string, parseStatus: string): Transaction[] {
  const headerAt = rows.findIndex((row) => {
    const values = row.map(normalHeader);
    return values.some((item) => ["date", "transactiondate", "valuedate", "postingdate"].includes(item))
      && values.some((item) => ["description", "narration", "details", "particulars", "memo"].includes(item));
  });
  if (headerAt < 0) throw new Error(`${sourceFile}: no Date and Description header row was found.`);
  const headers = rows[headerAt];
  const dateIndex = findColumn(headers, ["date", "transactiondate", "valuedate", "postingdate"]);
  const descriptionIndex = findColumn(headers, ["description", "narration", "details", "particulars", "memo"]);
  const debitIndex = findColumn(headers, ["debit", "withdrawal", "withdrawals", "paidout", "moneyout"]);
  const creditIndex = findColumn(headers, ["credit", "deposit", "deposits", "paidin", "moneyin"]);
  const amountIndex = findColumn(headers, ["amount", "transactionamount"]);
  const currencyIndex = findColumn(headers, ["currency", "curr"]);
  const accountIndex = findColumn(headers, ["accountid", "accountnumber", "account"]);
  if (debitIndex < 0 && creditIndex < 0 && amountIndex < 0) {
    throw new Error(`${sourceFile}: no Debit, Credit, or Amount column was found, so transaction values cannot be read.`);
  }
  const split = debitIndex >= 0 || creditIndex >= 0;
  return rows.slice(headerAt + 1).map((row, offset) => {
    const debit = debitIndex >= 0 ? Math.abs(numberValue(row[debitIndex])) : 0;
    const credit = creditIndex >= 0 ? Math.abs(numberValue(row[creditIndex])) : 0;
    const suppliedAmount = amountIndex >= 0 ? numberValue(row[amountIndex]) : 0;
    // A signed Amount column is only trusted when the source has no debit/credit split.
    const amount = split ? credit - debit : suppliedAmount;
    const partial = {
      transactionDate: clean(row[dateIndex]),
      description: clean(row[descriptionIndex]),
      debit: split ? debit : Math.max(0, -suppliedAmount),
      credit: split ? credit : Math.max(0, suppliedAmount),
      amount,
      currency: currencyIndex >= 0 ? clean(row[currencyIndex]) || "Unknown" : "Unknown",
      accountId: accountIndex >= 0 ? clean(row[accountIndex]) : "",
      sourceFile,
      sourceRow: String(headerAt + offset + 2),
      category: "Uncategorized",
      subcategory: "",
      pnlGroup: "",
      parseStatus: amount === 0 ? "Review: no debit or credit value was read" : parseStatus
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  }).filter((row) => row.transactionDate || row.description || row.amount !== 0);
}

/**
 * OFX DTPOSTED is YYYYMMDD[HHMMSS], so the date part is unambiguous and safe to
 * reformat. Dates from CSV, Excel, and PDF are left exactly as the bank wrote
 * them: 03/04/2026 could be March or April, and guessing is not this tool's job.
 */
function ofxDate(value: string): string {
  const match = value.match(/^(\d{4})(\d{2})(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : value;
}

/** Money is summed in cents-precision to keep totals off floating-point dust. */
function toCents(value: number): number {
  return Math.round(value * 100) / 100;
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
      transactionDate: ofxDate(value("DTPOSTED")),
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
      parseStatus: STATUS_PARSED
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  });
}

/** Converts a recognised Bank of America business export into master rows. */
function fromBoa(rows: string[][], sourceFile: string): ParsedFile {
  const statement = parseBoaBusinessCsv(rows, sourceFile);
  const transactions = statement.rows.map((row) => {
    const partial = {
      transactionDate: row.transactionDate,
      description: row.description,
      debit: Math.max(0, -row.amount),
      credit: Math.max(0, row.amount),
      amount: row.amount,
      // The export states no currency, and this tool does not invent one.
      currency: "Unknown",
      accountId: statement.accountId,
      sourceFile,
      sourceRow: row.sourceRow,
      category: "Uncategorized",
      subcategory: "",
      pnlGroup: "",
      parseStatus: row.parseNote || STATUS_PARSED
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  });
  return { transactions, notes: statement.notes };
}

type ParsedFile = { transactions: Transaction[]; notes: string[] };

async function parseStatementFile(filePath: string): Promise<ParsedFile> {
  if (!existsSync(filePath)) throw new Error(`File does not exist: ${filePath}`);
  const extension = extname(filePath).toLowerCase();
  const name = basename(filePath);
  if (extension === ".csv") {
    const rows = csvRows(await readFile(filePath, "utf8"));
    // Bank-specific layouts are tried before the generic header matcher.
    if (looksLikeBoaBusinessCsv(rows)) return fromBoa(rows, name);
    return { transactions: normalizeRows(rows, name, STATUS_PARSED), notes: [] };
  }
  if ([".xlsx", ".xls"].includes(extension)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error(`${name}: the workbook has no worksheets.`);
    const rows: string[][] = [];
    sheet.eachRow({ includeEmpty: false }, (row) => rows.push(excelRowValues(row)));
    if (looksLikeBoaBusinessCsv(rows)) return fromBoa(rows, name);
    return { transactions: normalizeRows(rows, name, STATUS_PARSED), notes: [] };
  }
  if ([".ofx", ".qfx"].includes(extension)) return { transactions: parseOfx(await readFile(filePath, "utf8"), name), notes: [] };
  if (extension === ".pdf") return { transactions: normalizeRows((await extractPdfTable(filePath)).rows, name, STATUS_PDF), notes: [] };
  throw new Error(`${name}: unsupported file type. Use CSV, XLSX/XLS, OFX/QFX, or PDF.`);
}

/** A folder is expanded to the statement files directly inside it, so a user can point at a month's folder. */
async function expandStatementPaths(paths: string[]): Promise<{ files: string[]; notes: string[] }> {
  const files: string[] = [];
  const notes: string[] = [];
  for (const path of paths) {
    const target = resolve(path);
    if (!existsSync(target)) { notes.push(`Skipped - path does not exist: ${target}`); continue; }
    if (!statSync(target).isDirectory()) { files.push(target); continue; }
    const entries = (await readdir(target, { withFileTypes: true }))
      .filter((entry) => entry.isFile() && STATEMENT_EXTENSIONS.includes(extname(entry.name).toLowerCase()))
      .map((entry) => join(target, entry.name))
      .sort();
    if (!entries.length) notes.push(`Skipped - no statement files found in folder: ${target}`);
    files.push(...entries);
  }
  return { files, notes };
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
  if (!sheet) throw new Error(`Template is missing the '${name}' sheet. Start from BankStatementTemplate.xlsx or create one with bank_statement_create_template.`);
  return sheet;
}

function isYes(value: string | undefined): boolean {
  return /^(yes|true|y|1)$/i.test((value ?? "").trim());
}

const MATCH_TYPES = ["contains", "equals", "starts_with", "regex", "amount_range"];
const RULE_FIELDS = ["description", "amount"];

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
      fingerprint: at(FINGERPRINT_COLUMN), parseStatus: at(14) || STATUS_PARSED
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

/**
 * Empties a sheet before it is rewritten. ExcelJS ignores spliceRows(1, rowCount)
 * on a sheet loaded from a file — the rows survive and new content lands beneath
 * the old, so report sheets have to be cleared one row at a time from the bottom.
 */
function clearSheet(sheet: ExcelJS.Worksheet): void {
  for (let rowNumber = sheet.rowCount; rowNumber >= 1; rowNumber -= 1) sheet.spliceRows(rowNumber, 1);
}

function buildReport(sheet: ExcelJS.Worksheet, setup: ReportSetup, rows: Transaction[]): void {
  const include = INCLUDE_FILTERS[setup.includeWhen];
  const groupBy = GROUP_FIELDS[setup.groupBy];
  const measure = MEASURES[setup.measure];
  const groups = new Map<string, number>();
  rows.filter(include).forEach((row) => groups.set(groupBy(row), toCents((groups.get(groupBy(row)) ?? 0) + measure(row))));

  clearSheet(sheet);
  sheet.addRow([setup.name]).font = { name: "Arial", size: 14, bold: true, color: { argb: "FF12302E" } };
  sheet.addRow([`Includes: ${setup.includeWhen}. Grouped by: ${setup.groupBy}. Measure: ${setup.measure}.`]).font = { name: "Arial", size: 9, italic: true };
  sheet.addRow([]);
  const header = sheet.addRow([setup.groupBy, setup.measure]);
  header.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
  [...groups.entries()].sort(([left], [right]) => left.localeCompare(right)).forEach(([group, value]) => sheet.addRow([group, value]));
  const total = sheet.addRow(["Total", toCents([...groups.values()].reduce((sum, value) => sum + value, 0))]);
  total.font = { name: "Arial", size: 10, bold: true };
  total.border = { top: { style: "thin" } };
  sheet.getColumn(1).width = 46;
  sheet.getColumn(2).width = 18;
  sheet.getColumn(2).numFmt = setup.measure === "count of transactions" ? "#,##0" : "#,##0.00;(#,##0.00);-";
  sheet.views = [{ showGridLines: false }];
}

function writeAudit(sheet: ExcelJS.Worksheet, lines: [string, string | number][], exceptions: string[]): void {
  clearSheet(sheet);
  sheet.addRow(["Import audit"]).font = { name: "Arial", size: 14, bold: true, color: { argb: "FF12302E" } };
  sheet.addRow([`Generated ${new Date().toISOString()}`]).font = { name: "Arial", size: 9, italic: true };
  sheet.addRow([]);
  lines.forEach(([label, value]) => {
    sheet.addRow([label, value]).getCell(1).font = { name: "Arial", size: 10, bold: true };
  });
  sheet.addRow([]);
  sheet.addRow(["Exceptions and notes"]).font = { name: "Arial", size: 11, bold: true };
  (exceptions.length ? exceptions : ["None."]).forEach((line) => sheet.addRow([line]));
  sheet.getColumn(1).width = 70;
  sheet.getColumn(2).width = 18;
  sheet.views = [{ showGridLines: false }];
}

/**
 * The shared tail of every import: classify, deduplicate, append, refresh the
 * enabled reports over the whole master table, and write the audit.
 */
function finishImport(workbook: ExcelJS.Workbook, incoming: Transaction[], exceptions: string[], sourceCount: number): { added: Transaction[]; allRows: Transaction[]; summary: string } {
  const master = getSheet(workbook, "Master Transactions");
  const rules = readRules(workbook);
  const existingRows = readMasterRows(master);
  const seen = new Set(existingRows.map((row) => row.fingerprint).filter(Boolean));

  const classified = incoming.map((row) => applyRules(row, rules));
  const added: Transaction[] = [];
  for (const row of classified) {
    // Deduplicates against the master table and within this import.
    if (seen.has(row.fingerprint)) continue;
    seen.add(row.fingerprint);
    added.push(row);
    writeMasterRow(master, row);
  }

  const allRows = [...existingRows, ...added];
  for (const setup of readReportSetups(workbook)) {
    const sheet = workbook.getWorksheet(setup.name);
    if (!sheet) { exceptions.push(`Report "${setup.name}" has no sheet with that name, so it was skipped.`); continue; }
    if (!setup.show) { exceptions.push(`Report "${setup.name}" is turned off in Report Setup, so it was left unchanged.`); continue; }
    if (!INCLUDE_FILTERS[setup.includeWhen] || !GROUP_FIELDS[setup.groupBy] || !MEASURES[setup.measure]) {
      exceptions.push(`Report "${setup.name}" uses a value this version does not understand, so it was skipped. Pick from the dropdowns in Report Setup.`);
      continue;
    }
    buildReport(sheet, setup, allRows);
  }

  const review = added.filter((row) => needsReview(row.parseStatus));
  review.forEach((row) => exceptions.push(`${row.sourceFile} row ${row.sourceRow}: ${row.parseStatus}`));
  const uncategorized = added.filter((row) => row.category === "Uncategorized").length;
  const fromPdf = added.filter((row) => row.parseStatus === STATUS_PDF).length;
  const fromAssistant = added.filter((row) => row.parseStatus === STATUS_ASSISTANT).length;
  if (fromPdf) exceptions.push(`${fromPdf} row(s) were reconstructed from a PDF text layer. Check them against the statement totals before relying on them.`);
  if (fromAssistant) exceptions.push(`${fromAssistant} row(s) were supplied by the assistant rather than machine-parsed. Verify each one against the source document.`);

  writeAudit(getSheet(workbook, "Audit"), [
    ["Sources in this import", sourceCount],
    ["Transactions extracted", incoming.length],
    ["Rows appended", added.length],
    ["Duplicates skipped", incoming.length - added.length],
    ["Rows in master table", allRows.length],
    ["Read from PDF text layer", fromPdf],
    ["Supplied by assistant", fromAssistant],
    ["Uncategorized in this import", uncategorized],
    ["Rows needing review", review.length],
    ["Active category rules", rules.length]
  ], exceptions);

  const summary = [
    `Extracted ${incoming.length}; appended ${added.length}; skipped ${incoming.length - added.length} duplicate(s).`,
    `Master table now holds ${allRows.length} transaction(s). ${uncategorized} uncategorized, ${review.length} needing review.`,
    exceptions.length ? `${exceptions.length} note(s) on the Audit sheet. Read it before using the output.` : "No exceptions recorded."
  ].join(" ");
  return { added, allRows, summary };
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
  description: "Check that a client workbook has the sheets, master columns, category rules, and report settings needed for consolidation.",
  inputSchema: z.object({ templatePath: z.string().describe("Absolute path to the client template workbook.") })
}, async ({ templatePath }) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(resolve(templatePath));
  const missingSheets = SHEET_NAMES.filter((name) => !workbook.getWorksheet(name));
  if (missingSheets.length) return { content: [{ type: "text", text: `Template needs these sheets: ${missingSheets.join(", ")}` }] };

  const headers = excelRowValues(getSheet(workbook, "Master Transactions").getRow(1)).map(normalHeader);
  const missingColumns = MASTER_COLUMNS.filter((column) => !headers.includes(normalHeader(column)));
  const rules = readRules(workbook);
  const reports = readReportSetups(workbook);
  const problems = [
    ...(missingColumns.length ? [`Master Transactions is missing: ${missingColumns.join(", ")}`] : []),
    ...(rules.length ? [] : ["Category Rules has no active rule, so every transaction will be Uncategorized."]),
    ...reports.filter((setup) => setup.show).flatMap((setup) => [
      ...(INCLUDE_FILTERS[setup.includeWhen] ? [] : [`${setup.name}: unknown "Include transactions when" value "${setup.includeWhen}"`]),
      ...(GROUP_FIELDS[setup.groupBy] ? [] : [`${setup.name}: unknown "Group by" value "${setup.groupBy}"`]),
      ...(MEASURES[setup.measure] ? [] : [`${setup.name}: unknown "Measure" value "${setup.measure}"`])
    ]),
    ...reports.filter((setup) => setup.show && !workbook.getWorksheet(setup.name)).map((setup) => `Report "${setup.name}" has no sheet with that name.`)
  ];
  const summary = `Sheets present. ${rules.length} active category rule(s), ${reports.filter((setup) => setup.show).length} report(s) turned on. Master table holds ${Math.max(0, getSheet(workbook, "Master Transactions").rowCount - 1)} transaction(s).`;
  return { content: [{ type: "text", text: problems.length ? `${summary}\nNeeds attention:\n- ${problems.join("\n- ")}` : `${summary} Template is ready for consolidation.` }] };
});

server.registerTool("bank_statement_consolidate", {
  description: "Read CSV, XLSX/XLS, OFX/QFX, and text-layer PDF statements (files or a folder of them), apply the workbook's category rules, append non-duplicate transactions, refresh the reports named in Report Setup, and save a consolidated workbook.",
  inputSchema: z.object({
    templatePath: z.string().describe("Absolute path to the client template workbook. Pass last month's consolidated output to keep a running ledger."),
    statementPaths: z.array(z.string()).min(1).describe("Absolute paths to statement files, or to folders containing them."),
    outputPath: z.string().describe("Absolute path for the resulting consolidated .xlsx workbook. Must differ from templatePath to keep the input intact.")
  })
}, async ({ templatePath, statementPaths, outputPath }) => {
  const source = resolve(templatePath);
  const target = resolve(outputPath);
  if (source === target) throw new Error("outputPath must differ from templatePath so the original workbook is preserved.");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(source);
  const { files, notes } = await expandStatementPaths(statementPaths);
  const exceptions = [...notes];
  if (!files.length) throw new Error("No readable statement files were found at the paths given.");

  const incoming: Transaction[] = [];
  for (const file of files) {
    try {
      const parsed = await parseStatementFile(file);
      incoming.push(...parsed.transactions);
      exceptions.push(...parsed.notes);
    } catch (error) {
      exceptions.push(`Not imported - ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  const { summary } = finishImport(workbook, incoming, exceptions, files.length);
  await workbook.xlsx.writeFile(target);
  return { content: [{ type: "text", text: `Created ${target}. Read ${files.length} file(s). ${summary}` }] };
});

server.registerTool("bank_statement_add_transactions", {
  description: "Append transactions you extracted yourself, for a scanned or image-only PDF that has no text layer. Rows are marked as assistant-extracted so the audit shows they need verification. Use bank_statement_consolidate first; only fall back to this when the file cannot be machine-parsed.",
  inputSchema: z.object({
    templatePath: z.string().describe("Absolute path to the client template workbook."),
    outputPath: z.string().describe("Absolute path for the resulting workbook."),
    sourceFile: z.string().describe("Name of the document these rows were read from, recorded on every row."),
    transactions: z.array(z.object({
      transactionDate: z.string().describe("Date exactly as printed on the statement."),
      description: z.string(),
      debit: z.number().optional().describe("Money out, as a positive number."),
      credit: z.number().optional().describe("Money in, as a positive number."),
      currency: z.string().optional(),
      accountId: z.string().optional().describe("Masked account identifier."),
      sourceRow: z.string().optional().describe("Page or line reference in the source document.")
    })).min(1).describe("Only transactions you actually read. Never infer or complete a row you could not see.")
  })
}, async ({ templatePath, outputPath, sourceFile, transactions }) => {
  const source = resolve(templatePath);
  const target = resolve(outputPath);
  if (source === target) throw new Error("outputPath must differ from templatePath so the original workbook is preserved.");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(source);
  const incoming = transactions.map((row, index) => {
    const debit = Math.abs(row.debit ?? 0);
    const credit = Math.abs(row.credit ?? 0);
    const partial = {
      transactionDate: row.transactionDate, description: row.description, debit, credit, amount: credit - debit,
      currency: row.currency || "Unknown", accountId: row.accountId ?? "", sourceFile,
      sourceRow: row.sourceRow ?? String(index + 1), category: "Uncategorized", subcategory: "", pnlGroup: "",
      parseStatus: debit === 0 && credit === 0 ? "Review: no debit or credit value was supplied" : STATUS_ASSISTANT
    };
    return { ...partial, fingerprint: fingerprint(partial) };
  });

  const { summary } = finishImport(workbook, incoming, [], 1);
  await workbook.xlsx.writeFile(target);
  return { content: [{ type: "text", text: `Created ${target}. ${summary}` }] };
});

server.registerTool("bank_statement_update_template", {
  description: "Change a workbook's category rules and report settings from a request, instead of editing the sheets by hand. Values are validated against the same vocabulary the dropdowns offer; an unrecognised value is rejected rather than written.",
  inputSchema: z.object({
    templatePath: z.string().describe("Absolute path to the workbook to change."),
    outputPath: z.string().optional().describe("Where to write the result. Omit to update the workbook in place."),
    addRules: z.array(z.object({
      priority: z.number().optional().describe("Lower numbers win. Defaults to the end of the list."),
      lookIn: z.enum(["Description", "Amount"]).optional(),
      matchType: z.enum(["contains", "equals", "starts_with", "regex", "amount_range"]).optional(),
      match: z.string().describe("Text to look for, or low..high for amount_range."),
      category: z.string(),
      subcategory: z.string().optional(),
      pnlGroup: z.string().optional(),
      active: z.boolean().optional()
    })).optional().describe("Category rules to add."),
    removeRulesMatching: z.array(z.string()).optional().describe("Remove rules whose match text equals one of these, case-insensitively."),
    updateReports: z.array(z.object({
      name: z.string().describe("Must equal an existing Report name."),
      includeWhen: z.enum(["All transactions", "Debit > 0", "Credit > 0", "P&L group is not blank", "Category is not blank"]).optional(),
      groupBy: z.enum(["Category", "Subcategory", "P&L group", "Currency", "Account ID", "Source file", "Transaction date"]).optional(),
      measure: z.enum(["Sum of Amount", "Sum of Debit", "Sum of Credit", "Count of transactions"]).optional(),
      show: z.boolean().optional()
    })).optional().describe("Report settings to change.")
  })
}, async ({ templatePath, outputPath, addRules, removeRulesMatching, updateReports }) => {
  const source = resolve(templatePath);
  const target = outputPath ? resolve(outputPath) : source;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(source);
  const changes: string[] = [];

  const rulesSheet = getSheet(workbook, "Category Rules");
  if (removeRulesMatching?.length) {
    const wanted = removeRulesMatching.map((value) => value.trim().toLowerCase());
    // Collected first, then removed bottom-up, so earlier deletions do not shift later row numbers.
    const doomed: number[] = [];
    rulesSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      if (wanted.includes(clean(row.getCell(4).value).toLowerCase())) doomed.push(rowNumber);
    });
    doomed.reverse().forEach((rowNumber) => rulesSheet.spliceRows(rowNumber, 1));
    changes.push(doomed.length ? `Removed ${doomed.length} rule(s).` : "No rule matched the text given for removal.");
  }

  if (addRules?.length) {
    const existing = sheetRows(rulesSheet);
    const highest = existing.reduce((max, row) => Math.max(max, Number(row["Priority"]) || 0), 0);
    addRules.forEach((rule, index) => {
      if (rule.matchType === "amount_range" && !/^-?[\d.]+\.\.-?[\d.]+$/.test(rule.match.trim())) {
        throw new Error(`Rule "${rule.match}" uses amount_range, which needs the form low..high, for example 100..500.`);
      }
      if (rule.matchType === "regex") { try { new RegExp(rule.match); } catch { throw new Error(`Rule "${rule.match}" is not a valid regular expression.`); } }
      rulesSheet.addRow([
        rule.priority ?? highest + (index + 1) * 10,
        rule.lookIn ?? "Description",
        rule.matchType ?? "contains",
        rule.match,
        rule.category,
        rule.subcategory ?? "",
        rule.pnlGroup ?? "",
        rule.active === false ? "No" : "Yes"
      ]);
    });
    changes.push(`Added ${addRules.length} rule(s).`);
  }

  if (updateReports?.length) {
    const reportSheet = getSheet(workbook, "Report Setup");
    for (const change of updateReports) {
      let found = false;
      reportSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1 || clean(row.getCell(1).value).toLowerCase() !== change.name.trim().toLowerCase()) return;
        found = true;
        if (change.includeWhen) row.getCell(2).value = change.includeWhen;
        if (change.groupBy) row.getCell(3).value = change.groupBy;
        if (change.measure) row.getCell(4).value = change.measure;
        if (change.show !== undefined) row.getCell(5).value = change.show ? "Yes" : "No";
      });
      if (!found) throw new Error(`Report Setup has no report named "${change.name}". Existing reports: ${readReportSetups(workbook).map((setup) => setup.name).join(", ")}.`);
      changes.push(`Updated report "${change.name}".`);
    }
  }

  if (!changes.length) return { content: [{ type: "text", text: "Nothing to change. Pass addRules, removeRulesMatching, or updateReports." }] };
  await workbook.xlsx.writeFile(target);
  const active = readRules(workbook).length;
  return { content: [{ type: "text", text: `${changes.join(" ")} Saved ${target}. ${active} active rule(s) now. Re-run bank_statement_consolidate to apply them to existing rows.` }] };
});

await server.connect(new StdioServerTransport());
