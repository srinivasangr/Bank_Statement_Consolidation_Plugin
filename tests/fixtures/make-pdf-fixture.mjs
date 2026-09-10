// Generates tests/fixtures/sample-statement.pdf: a synthetic, text-layer bank
// statement laid out in columns, the way a real bank PDF export is.
// Regenerate with: node tests/fixtures/make-pdf-fixture.mjs
import { writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const COLUMNS = [56, 150, 380, 460, 540];
const ROWS = [
  ["Date", "Description", "Withdrawal", "Deposit", "Balance"],
  ["2026-09-11", "AMAZON MARKETPLACE", "64.30", "", "3,120.55"],
  ["2026-09-12", "UBER TRIP 8842", "23.75", "", "3,096.80"],
  ["2026-09-15", "CLIENT PAYMENT INV 5013", "", "2,400.00", "5,496.80"],
  ["2026-09-18", "PAYROLL SEPTEMBER", "1,910.00", "", "3,586.80"],
  ["2026-09-21", "CITY WATER UTILITY", "88.40", "", "3,498.40"]
];

function escapeText(value) {
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

const lines = ["BT /F1 14 Tf 56 742 Td (Northwind Bank - Statement of Account) Tj ET", "BT /F1 9 Tf 56 726 Td (Account XXXX4417   Period 01 Sep 2026 to 30 Sep 2026) Tj ET"];
ROWS.forEach((row, rowIndex) => {
  const y = 690 - rowIndex * 22;
  row.forEach((cell, columnIndex) => {
    if (!cell) return;
    lines.push(`BT /F1 ${rowIndex === 0 ? 10 : 9} Tf ${COLUMNS[columnIndex]} ${y} Td (${escapeText(cell)}) Tj ET`);
  });
});
const content = lines.join("\n");

const objects = [
  "<< /Type /Catalog /Pages 2 0 R >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
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

const outputPath = fileURLToPath(new URL("./sample-statement.pdf", import.meta.url));
await writeFile(outputPath, Buffer.from(pdf, "latin1"));
console.log(`Wrote ${outputPath}`);
