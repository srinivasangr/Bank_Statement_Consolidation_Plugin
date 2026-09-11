// Run a consolidation without an agent, against the real MCP service.
//
//   node scripts/consolidate.mjs <template.xlsx> <statements folder or file> <output.xlsx>
//
// Use it to try a real bank statement locally. Nothing leaves this machine.
import { existsSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

const [templateArg, statementsArg, outputArg] = process.argv.slice(2);
if (!templateArg || !statementsArg || !outputArg) {
  console.error("Usage: node scripts/consolidate.mjs <template.xlsx> <statements folder or file> <output.xlsx>");
  process.exit(1);
}

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const template = resolve(templateArg);
const statements = resolve(statementsArg);
const output = resolve(outputArg);

for (const [label, path] of [["Template", template], ["Statements", statements]]) {
  if (!existsSync(path)) { console.error(`${label} not found: ${path}`); process.exit(1); }
}

const server = existsSync(`${root}/dist/index.js`) ? `${root}/dist/index.js` : `${root}/bundle/server.mjs`;
const client = new Client({ name: "bank-statement-cli", version: "0.1.0" });
const transport = new StdioClientTransport({ command: process.execPath, args: [server], cwd: root, stderr: "inherit" });

try {
  await client.connect(transport);

  const inspected = await client.callTool({ name: "bank_statement_inspect_template", arguments: { templatePath: template } });
  console.log(`\nTemplate check:\n  ${inspected.content.map((item) => item.text).join("\n  ")}`);
  if (inspected.isError) process.exitCode = 1;

  const result = await client.callTool({
    name: "bank_statement_consolidate",
    arguments: { templatePath: template, statementPaths: [statements], outputPath: output }
  });
  const text = result.content.map((item) => item.text).join("\n");
  console.log(`\n${result.isError ? "Failed" : "Result"}:\n  ${text.split(". ").join(".\n  ")}`);
  if (result.isError) process.exitCode = 1;
  else console.log(`\nOpen ${output} and read the Audit sheet before trusting the numbers.`);
} finally {
  await client.close();
}
