// Produces bundle/server.mjs: the whole service and its dependencies in one
// file, so the plugin runs after a plain `git clone` with no npm install.
// Committed to the repository; regenerate with `npm run bundle`.
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const root = fileURLToPath(new URL("../", import.meta.url));

const shared = {
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node20",
  // exceljs and pdf.js are CommonJS and call require() for Node builtins.
  // esbuild's ESM output stubs require and throws, so give them a real one.
  banner: { js: "import { createRequire as __createRequire } from 'node:module';\nconst require = __createRequire(import.meta.url);" },
  // Optional native renderer that pdf.js probes for and never needs here.
  external: ["canvas"],
  logLevel: "warning",
  metafile: true
};

const outputs = [
  { entry: `${root}src/index.ts`, out: `${root}bundle/server.mjs` },
  // pdf.js runs its parser in a "fake worker" it imports at runtime, resolved
  // as ./pdf.worker.mjs next to whatever loaded it. Bundling it to that exact
  // name is what lets PDF extraction work from the bundle.
  { entry: `${root}node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs`, out: `${root}bundle/pdf.worker.mjs` }
];

for (const { entry, out } of outputs) {
  const result = await build({ ...shared, entryPoints: [entry], outfile: out });
  const bytes = Object.values(result.metafile.outputs)[0].bytes;
  console.log(`Wrote ${out} (${(bytes / 1024 / 1024).toFixed(1)} MB)`);
}
