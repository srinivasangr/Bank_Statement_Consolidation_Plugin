// Started by .mcp.json. Paths resolve against this file, not the caller's
// working directory, so the service starts the same way from any host.
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const serverPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));

if (!existsSync(serverPath)) {
  console.error("Bank Statement Consolidator is not built. Run npm install in the plugin folder, which builds it automatically.");
  process.exit(1);
}

const child = spawn(process.execPath, [serverPath], { stdio: "inherit" });
child.on("exit", (code, signal) => process.exit(signal ? 1 : code ?? 1));
child.on("error", (error) => { console.error(`Could not start the service: ${error.message}`); process.exit(1); });
