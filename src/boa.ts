/**
 * Bank of America business-account CSV exports.
 *
 * These have no header row. Column 0 is a section label, and the section decides
 * what the row means:
 *
 *   Statement Information   account number, period, company, address
 *   Account Summary         beginning balance, totals, ending balance, counts
 *   Deposits and other credits   [section, date, check, amount, description, reference]
 *   Withdrawals and other Debits same shape, amount negative
 *   Service fees                 same shape, no reference column
 *   Checks                       same shape, description may be blank
 *   Daily Ledger Balances        [section, date, balance] - not transactions
 *
 * Amounts are signed, so the section is not needed to decide direction, but it
 * is needed to tell transactions from balances and headers.
 */

const TRANSACTION_SECTIONS = new Set([
  "deposits and other credits",
  "withdrawals and other debits",
  "service fees",
  "checks"
]);

const NON_TRANSACTION_SECTIONS = new Set([
  "statement information",
  "account summary",
  "daily ledger balances"
]);

export type BoaRow = {
  transactionDate: string;
  description: string;
  amount: number;
  sourceRow: string;
  parseNote: string;
};

export type BoaStatement = {
  rows: BoaRow[];
  accountId: string;
  period: string;
  notes: string[];
};

function label(row: string[]): string {
  return (row[0] ?? "").trim().toLowerCase();
}

function money(value: string | undefined): number {
  const text = String(value ?? "").replace(/[$,]/g, "").trim();
  if (!text) return 0;
  const result = Number(text);
  return Number.isFinite(result) ? result : 0;
}

/** Recognises the format by its section labels rather than by filename. */
export function looksLikeBoaBusinessCsv(rows: string[][]): boolean {
  return rows.some((row) => label(row) === "account summary")
    && rows.some((row) => TRANSACTION_SECTIONS.has(label(row)));
}

const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

export type Period = { start: string; end: string };

/** Reads "January 1, 2025 to January 31, 2025" from the statement header. */
export function parsePeriod(period: string): Period | null {
  const match = period.match(/([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})\s+to\s+([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})/i);
  if (!match) return null;
  const toIso = (monthName: string, day: string, year: string) => {
    const monthIndex = MONTHS.indexOf(monthName.toLowerCase());
    if (monthIndex < 0) return null;
    return `${year}-${String(monthIndex + 1).padStart(2, "0")}-${day.padStart(2, "0")}`;
  };
  const start = toIso(match[1], match[2], match[3]);
  const end = toIso(match[4], match[5], match[6]);
  return start && end ? { start, end } : null;
}

/**
 * BOA writes MM/DD/YY for most rows, but the Checks section writes MM/DD with no
 * year at all. The missing year is taken from the statement period rather than
 * guessed: whichever of the period's start or end year places the date inside
 * the period wins. Every date, however it was written, is then checked against
 * the period, so a misread cannot pass silently.
 */
function isoDate(value: string, period: Period | null): { date: string; note: string } {
  const text = String(value ?? "").trim();
  const outside = (date: string) => period && (date < period.start || date > period.end)
    ? `Review: date ${date} falls outside the statement period ${period.start} to ${period.end}`
    : "";

  const full = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (full) {
    const [, month, day, year] = full;
    if (Number(month) < 1 || Number(month) > 12 || Number(day) < 1 || Number(day) > 31) {
      return { date: text, note: "Review: the date is out of range" };
    }
    const date = `${year.length === 4 ? year : `20${year}`}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    return { date, note: outside(date) };
  }

  const short = text.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (short) {
    const [, month, day] = short;
    if (Number(month) < 1 || Number(month) > 12 || Number(day) < 1 || Number(day) > 31) {
      return { date: text, note: "Review: the date is out of range" };
    }
    if (!period) return { date: text, note: "Review: the date has no year and the statement period is unknown" };
    const suffix = `-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    for (const year of new Set([period.start.slice(0, 4), period.end.slice(0, 4)])) {
      const candidate = `${year}${suffix}`;
      if (candidate >= period.start && candidate <= period.end) return { date: candidate, note: "" };
    }
    return { date: `${period.start.slice(0, 4)}${suffix}`, note: `Review: date ${month}/${day} has no year and does not fall inside the statement period` };
  }

  return { date: text, note: "Review: the date could not be read" };
}

export function parseBoaBusinessCsv(rows: string[][], sourceFile: string): BoaStatement {
  const notes: string[] = [];
  const information = rows.find((row) => label(row) === "statement information");
  const summary = rows.find((row) => label(row) === "account summary");
  const accountId = (information?.[1] ?? "").trim();
  const period = (information?.[2] ?? summary?.[1] ?? "").trim();
  const periodRange = parsePeriod(period);

  const unknown = new Set<string>();
  const parsed: BoaRow[] = [];

  rows.forEach((row, index) => {
    const section = label(row);
    if (!section || NON_TRANSACTION_SECTIONS.has(section)) return;
    if (!TRANSACTION_SECTIONS.has(section)) { unknown.add(row[0]); return; }

    const { date, note } = isoDate(row[1] ?? "", periodRange);
    const checkNumber = (row[2] ?? "").trim();
    const amount = money(row[3]);
    const description = (row[4] ?? "").trim() || (checkNumber ? `Check ${checkNumber}` : "");
    parsed.push({
      transactionDate: date,
      description,
      amount,
      sourceRow: String(index + 1),
      parseNote: note || (amount === 0 ? "Review: the amount was zero or unreadable" : "")
    });
  });

  // An unrecognised section is reported, never dropped in silence.
  for (const section of unknown) {
    notes.push(`${sourceFile}: section "${section}" was not recognised, so its rows were not imported.`);
  }

  if (summary) {
    const beginning = money(summary[2]);
    const ending = money(summary[7]);
    const expected = ending - beginning;
    const actual = parsed.reduce((sum, row) => sum + row.amount, 0);
    const difference = Math.round((expected - actual) * 100) / 100;
    notes.push(difference === 0
      ? `${sourceFile}: reconciled. ${parsed.length} transactions total ${actual.toFixed(2)}, matching the stated balance change from ${beginning.toFixed(2)} to ${ending.toFixed(2)}.`
      : `${sourceFile}: DOES NOT RECONCILE. Transactions total ${actual.toFixed(2)} but the statement's balance moved ${expected.toFixed(2)} (difference ${difference.toFixed(2)}). Check for missing rows before using this data.`);
  } else {
    notes.push(`${sourceFile}: no Account Summary row, so the import could not be reconciled against a stated balance.`);
  }

  return { rows: parsed, accountId, period, notes };
}
