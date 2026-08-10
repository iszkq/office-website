import { readdir, readFile, stat } from "node:fs/promises";
import { extname, join, relative, resolve } from "node:path";

const root = resolve(process.argv[2] || "out");
const forbidden = [
  "office-editor.ziziyi.com",
  "office-plugins.ziziyi.com",
  "googletagmanager.com",
  "google-analytics.com",
  "api.producthunt.com",
  "chromewebstore.google.com",
];
const textExtensions = new Set([
  ".css",
  ".html",
  ".js",
  ".json",
  ".mjs",
  ".svg",
  ".txt",
  ".xml",
]);

const matches = [];

async function scan(path) {
  const info = await stat(path);
  if (info.isDirectory()) {
    for (const entry of await readdir(path)) await scan(join(path, entry));
    return;
  }
  if (!textExtensions.has(extname(path).toLowerCase())) return;

  const content = await readFile(path, "utf8");
  for (const hostname of forbidden) {
    if (content.includes(hostname)) {
      matches.push(`${relative(root, path)} -> ${hostname}`);
    }
  }
}

await scan(root);

if (matches.length > 0) {
  console.error("Forbidden third-party runtime references found:\n" + matches.join("\n"));
  process.exit(1);
}

console.log("Self-hosted runtime verification passed.");
