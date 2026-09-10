// Started by .mcp.json. Paths resolve against this file, not the caller's
// working directory, so the service starts the same way from any host.
//
// dist/ is the TypeScript build and wins when present, so a developer's latest
// compile is what runs. bundle/server.mjs is the committed, dependency-free
// build used when the plugin was installed straight from Git, where neither
// dist/ nor node_modules/ exists.
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const candidates = [
  { path: fileURLToPath(new URL("../dist/index.js", import.meta.url)), label: "dist build" },
  { path: fileURLToPath(new URL("../bundle/server.mjs", import.meta.url)), label: "bundled build" }
];

const chosen = candidates.find((candidate) => existsSync(candidate.path));
if (!chosen) {
  console.error("Bank Statement Consolidator has no runnable build. Run npm install in the plugin folder, or reinstall a release that includes bundle/server.mjs.");
  process.exit(1);
}

const child = spawn(process.execPath, [chosen.path], { stdio: "inherit" });
child.on("exit", (code, signal) => process.exit(signal ? 1 : code ?? 1));
child.on("error", (error) => { console.error(`Could not start the service: ${error.message}`); process.exit(1); });
