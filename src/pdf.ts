import { readFile } from "node:fs/promises";
import { basename } from "node:path";

type TextItem = { text: string; x: number; y: number; width: number; page: number };

// This service speaks MCP over stdout. pdf.js reports warnings with console.log,
// which would inject text straight into the JSON-RPC stream and break the session.
// Everything pdf.js prints is redirected to stderr for the duration of the parse.
async function withQuietConsole<T>(work: () => Promise<T>): Promise<T> {
  const { log, info, warn, debug } = console;
  console.log = console.info = console.warn = console.debug = (...args: unknown[]) => { console.error(...args); };
  try {
    return await work();
  } finally {
    Object.assign(console, { log, info, warn, debug });
  }
}

async function readTextItems(filePath: string): Promise<{ items: TextItem[]; pages: number }> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const data = new Uint8Array(await readFile(filePath));
  const document = await pdfjs.getDocument({ data, useSystemFonts: true, verbosity: 0 }).promise;
  const items: TextItem[] = [];
  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const content = await page.getTextContent();
    for (const item of content.items as { str?: string; transform?: number[]; width?: number }[]) {
      const text = (item.str ?? "").trim();
      if (!text || !item.transform) continue;
      items.push({ text, x: item.transform[4], y: item.transform[5], width: item.width ?? text.length * 4, page: pageNumber });
    }
    page.cleanup();
  }
  await document.destroy();
  return { items, pages: document.numPages };
}

// Text arrives as scattered fragments. Group by baseline to rebuild visual lines.
function groupIntoLines(items: TextItem[]): TextItem[][] {
  const sorted = [...items].sort((left, right) => left.page - right.page || right.y - left.y || left.x - right.x);
  const lines: TextItem[][] = [];
  for (const item of sorted) {
    const current = lines[lines.length - 1];
    const anchor = current?.[0];
    if (current && anchor && anchor.page === item.page && Math.abs(anchor.y - item.y) <= 3) current.push(item);
    else lines.push([item]);
  }
  return lines.map((line) => line.sort((left, right) => left.x - right.x));
}

const DATE_WORDS = ["date", "transactiondate", "valuedate", "postingdate", "posted"];
const DESCRIPTION_WORDS = ["description", "narration", "details", "particulars", "memo", "transaction"];

function normalWord(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function isHeaderLine(line: TextItem[]): boolean {
  const words = line.map((item) => normalWord(item.text));
  return words.some((word) => DATE_WORDS.includes(word)) && words.some((word) => DESCRIPTION_WORDS.includes(word));
}

/**
 * Finds the vertical whitespace gutters that separate columns.
 *
 * Header positions cannot be trusted for this. On a real HDFC statement the
 * "Narration" heading sits at x=144 while its text starts at x=68, so splitting
 * on the midpoint between headings pulls the description into the date column.
 * Amount columns are right-aligned, so their left edge moves with the number of
 * digits. What is stable is the empty space between columns, across all rows.
 *
 * Occupancy is counted per line so one full-width line cannot erase a gutter.
 */
function columnBoundaries(lines: TextItem[][], minGap = 4): number[] {
  const items = lines.flat();
  if (!items.length) return [];
  const maxX = Math.ceil(Math.max(...items.map((item) => item.x + item.width))) + 2;
  const counts = new Int32Array(maxX + 2);
  for (const line of lines) {
    const occupied = new Uint8Array(maxX + 2);
    for (const item of line) {
      const from = Math.max(0, Math.floor(item.x));
      const to = Math.min(maxX, Math.ceil(item.x + item.width));
      occupied.fill(1, from, to + 1);
    }
    for (let x = 0; x <= maxX; x += 1) if (occupied[x]) counts[x] += 1;
  }

  const firstX = Math.floor(Math.min(...items.map((item) => item.x)));
  const boundaries: number[] = [];
  let runStart = -1;
  for (let x = firstX; x <= maxX; x += 1) {
    const empty = counts[x] === 0;
    if (empty && runStart < 0) runStart = x;
    if (!empty && runStart >= 0) {
      if (x - runStart >= minGap) boundaries.push((runStart + x - 1) / 2);
      runStart = -1;
    }
  }
  return boundaries;
}

function columnize(line: TextItem[], boundaries: number[]): string[] {
  const cells: string[][] = Array.from({ length: boundaries.length + 1 }, () => []);
  for (const item of line) {
    // An item belongs to the column its own span starts in.
    let column = 0;
    while (column < boundaries.length && item.x >= boundaries[column]) column += 1;
    cells[column].push(item.text);
  }
  return cells.map((parts) => parts.join(" ").trim());
}

/**
 * Some statements render the date and the description with no space between
 * them, so no amount of column geometry can separate them: on a real HDFC
 * statement the date ends at x=68 and the narration begins at x=68. When a
 * leading date is followed by other text and the next cell is empty, the two
 * are split apart. A cell holding only a date is left alone.
 */
function splitLeadingDate(row: string[]): string[] {
  // The separator is optional: pdf.js sometimes emits the date and the text that
  // abuts it as one run, with no space at all ("24/01/26UPI-AMAZON").
  const match = (row[0] ?? "").match(/^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*(\S.*)$/);
  if (!match || row.length < 2 || row[1]) return row;
  const split = [...row];
  split[0] = match[1];
  split[1] = match[2];
  return split;
}

const LOOKS_NUMERIC = /^[(-]?[\d,]+\.?\d*\)?$/;
const LOOKS_LIKE_DATE = /\d{1,4}[/-]\d{1,2}([/-]\d{2,4})?/;
// A transaction's date cell holds a date and nothing else. A footer that merely
// mentions a date ("Statement From : 01/01/2026 ...") must not qualify.
const IS_ONLY_A_DATE = /^\d{1,4}[/-]\d{1,2}(?:[/-]\d{2,4})?$/;

type PlacedRow = { cells: string[]; y: number; page: number };

/**
 * A long description wraps onto its own line, carrying no date and no amount.
 * Those lines belong to the transaction above, so their text is folded into its
 * description cell - not into whichever cell happens to be longest, which on a
 * real statement put the wrapped narration into the cheque-number column.
 *
 * Proximity decides what counts as a continuation. A wrapped line sits one line
 * height below its transaction, while a summary line like "Opening Balance" is
 * separated by a larger gap and must stay its own row.
 */
function mergeWrappedRows(rows: PlacedRow[], descriptionColumn: number, maxGap: number): string[][] {
  const merged: PlacedRow[] = [];
  for (const row of rows) {
    const filled = row.cells.filter(Boolean);
    const previous = merged[merged.length - 1];
    const adjacent = previous !== undefined
      && previous.page === row.page
      && previous.y - row.y <= maxGap;
    const isContinuation = merged.length > 1
      && adjacent
      && filled.length === 1
      && !LOOKS_NUMERIC.test(filled[0])
      && !LOOKS_LIKE_DATE.test(filled[0]);
    if (!isContinuation || !previous) { merged.push({ ...row, cells: [...row.cells] }); continue; }
    const target = Math.min(descriptionColumn, previous.cells.length - 1);
    previous.cells[target] = `${previous.cells[target] ?? ""} ${filled[0]}`.trim();
  }
  return merged.map((row) => row.cells);
}

/** Typical line spacing, used to tell a wrapped line from a new block. */
function medianLineGap(rows: PlacedRow[]): number {
  const gaps: number[] = [];
  for (let index = 1; index < rows.length; index += 1) {
    if (rows[index].page !== rows[index - 1].page) continue;
    const gap = rows[index - 1].y - rows[index].y;
    if (gap > 0) gaps.push(gap);
  }
  if (!gaps.length) return 14;
  gaps.sort((left, right) => left - right);
  return gaps[Math.floor(gaps.length / 2)];
}

export type PdfExtraction = {
  rows: string[][];
  pages: number;
  headerFound: boolean;
  lineCount: number;
  discardedRows: number;
};

/**
 * Reconstructs a table from a text-layer PDF. Returns rows shaped like a CSV so
 * the shared header-matching and normalization path can consume them.
 * A PDF with no text layer (a scan) yields no items; the caller must treat that
 * as an exception rather than as an empty statement.
 */
export async function extractPdfTable(filePath: string): Promise<PdfExtraction> {
  const { items, pages } = await withQuietConsole(() => readTextItems(filePath));
  if (!items.length) {
    throw new Error(`${basename(filePath)}: no text layer was found. This looks like a scanned image PDF, which needs OCR or assistant-led extraction; it was not imported.`);
  }
  const lines = groupIntoLines(items);
  const headerAt = lines.findIndex(isHeaderLine);
  if (headerAt < 0) {
    throw new Error(`${basename(filePath)}: text was read, but no row with both a date and a description heading was found, so the columns could not be identified.`);
  }
  // The grid is measured from the header line down, and covers continuation
  // pages that repeat no headings of their own.
  const body = lines.slice(headerAt);
  const header = lines[headerAt];
  // Gutters are measured from rows that actually look tabular; address blocks
  // and footers would otherwise fill the gaps and hide every column break.
  const tabular = body.filter((line) => line.length >= Math.max(3, Math.ceil(header.length / 2)));
  const gutters = columnBoundaries(tabular.length >= 3 ? tabular : body);
  // The header says how many columns there are. Gutters are only trusted when
  // they resolve all of them; when neighbouring columns touch with no whitespace
  // there is no gutter to find, and header midpoints are the better guide.
  const anchors = header.map((item) => item.x);
  const midpoints = anchors.slice(1).map((anchor, index) => (anchor + anchors[index]) / 2);
  const boundaries = gutters.length >= midpoints.length ? gutters : midpoints;

  const placed: PlacedRow[] = body
    .map((line) => ({ cells: splitLeadingDate(columnize(line, boundaries)), y: line[0].y, page: line[0].page }))
    .filter((row) => row.cells.some(Boolean));

  const headerRow = placed[0]?.cells ?? [];
  const descriptionColumn = Math.max(1, headerRow.findIndex((cell) => DESCRIPTION_WORDS.includes(normalWord(cell))));
  const merged = mergeWrappedRows(placed, descriptionColumn, medianLineGap(placed) * 1.4);

  // Everything below the header is swept up, including the address block and the
  // legal footer repeated on later pages, plus summary lines like "Opening
  // Balance" and the closing totals. A transaction must carry a date in the date
  // column; anything else is discarded and counted, never silently dropped.
  const headerCells = merged[0] ?? [];
  const dateColumn = Math.max(0, headerCells.findIndex((cell) => DATE_WORDS.includes(normalWord(cell))));
  const kept = merged.filter((row, index) => index === 0 || IS_ONLY_A_DATE.test((row[dateColumn] ?? "").trim()));
  return { rows: kept, pages, headerFound: true, lineCount: lines.length, discardedRows: merged.length - kept.length };
}
