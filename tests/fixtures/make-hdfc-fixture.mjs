// Generates tests/fixtures/hdfc/hdfc-style.pdf: a synthetic statement that
// reproduces the awkward parts of a real HDFC PDF export, all of which broke the
// first version of the PDF reader.
//
//   - the date and the narration touch, leaving no gutter to split on
//   - the "Narration" heading sits far right of the text beneath it
//   - amounts are right-aligned, so their left edge moves with the digit count
//   - a long narration wraps onto its own line
//   - a summary line carries no date, and a footer merely mentions one
//
// Regenerate with: node tests/fixtures/make-hdfc-fixture.mjs
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

// [text, x, y] - positions copied from the shape of a real statement.
const PLACED = [
  ["Statement of account", 40, 700],
  ["MR EXAMPLE CUSTOMER", 40, 686],
  ["Date", 40, 660], ["Narration", 144, 660], ["Chq./Ref.No.", 284, 660],
  ["Withdrawal Amt.", 405, 660], ["Deposit Amt.", 491, 660], ["Closing Balance", 564, 660],

  // Date ends exactly where the narration begins: no whitespace between them.
  ["24/01/26", 34, 640], ["UPI-AMAZON", 68, 640], ["0000123456789", 293, 640],
  ["1,499.00", 442, 640], ["98,501.00", 580, 640],
  ["INDIA-AMAZON@RAPL-RATN001RAPL", 68, 626],

  ["09/02/26", 34, 606], ["ACH C- WIPRO LIMITED", 68, 606], ["0000123456790", 293, 606],
  ["564.00", 505, 606], ["99,065.00", 580, 606],

  ["25/02/26", 34, 586], ["POS PURCHASE GROCERY", 68, 586], ["0000123456791", 293, 586],
  ["12,345.67", 436, 586], ["86,719.33", 580, 586],

  ["Opening Balance", 40, 528],
  ["Statement From : 01/01/2026 This is a computer generated statement.", 40, 506]
];

function escapeText(value) {
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

const content = PLACED.map(([text, x, y]) => `BT /F1 9 Tf ${x} ${y} Td (${escapeText(text)}) Tj ET`).join("\n");

const objects = [
  "<< /Type /Catalog /Pages 2 0 R >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 700 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
  "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
  `<< /Length ${Buffer.byteLength(content, "latin1")} >>\nstream\n${content}\nendstream`
];

let pdf = "%PDF-1.4\n";
const offsets = [];
objects.forEach((body, index) => {
  offsets.push(Buffer.byteLength(pdf, "latin1"));
  pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
});
const xrefAt = Buffer.byteLength(pdf, "latin1");
pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
offsets.forEach((offset) => { pdf += `${String(offset).padStart(10, "0")} 00000 n \n`; });
pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`;

const directory = fileURLToPath(new URL("./hdfc/", import.meta.url));
await mkdir(directory, { recursive: true });
const outputPath = `${directory}hdfc-style.pdf`;
await writeFile(outputPath, Buffer.from(pdf, "latin1"));
console.log(`Wrote ${outputPath}`);
