import { readFile } from "node:fs/promises";
import { basename } from "node:path";

type TextItem = { text: string; x: number; y: number; page: number };

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
    for (const item of content.items as { str?: string; transform?: number[] }[]) {
      const text = (item.str ?? "").trim();
      if (!text || !item.transform) continue;
      items.push({ text, x: item.transform[4], y: item.transform[5], page: pageNumber });
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

// Column anchors come from the header line's x positions; every other line is
// bucketed against the midpoints between them. This keeps left-aligned text and
// right-aligned amounts in their own columns.
function columnize(line: TextItem[], anchors: number[]): string[] {
  const boundaries = anchors.slice(1).map((anchor, index) => (anchor + anchors[index]) / 2);
  const cells: string[][] = anchors.map(() => []);
  for (const item of line) {
    let column = 0;
    while (column < boundaries.length && item.x >= boundaries[column]) column += 1;
    cells[column].push(item.text);
  }
  return cells.map((parts) => parts.join(" ").trim());
}

export type PdfExtraction = {
  rows: string[][];
  pages: number;
  headerFound: boolean;
  lineCount: number;
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
  // The header sets the column grid for every page, including continuation pages
  // that repeat no headings of their own.
  const anchors = lines[headerAt].map((item) => item.x);
  const rows = lines.slice(headerAt).map((line) => columnize(line, anchors)).filter((row) => row.some(Boolean));
  return { rows, pages, headerFound: true, lineCount: lines.length };
}
