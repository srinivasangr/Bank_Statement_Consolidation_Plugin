// Regenerates assets/BankStatementTemplate.xlsx from the same builder the
// bank_statement_create_template tool uses, so the shipped workbook and a
// freshly created one never drift apart.
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
const builderPath = `${root}dist/template.js`;

if (!existsSync(builderPath)) {
  console.error("dist/template.js is missing. Run: npm install && npm run build");
  process.exit(1);
}

const { createTemplateWorkbook } = await import(new URL("../dist/template.js", import.meta.url));
const outputPath = `${root}assets/BankStatementTemplate.xlsx`;
await createTemplateWorkbook().xlsx.writeFile(outputPath);
console.log(`Wrote ${outputPath}`);
