import { readdir, readFile, writeFile } from "node:fs/promises";
import { extname, join } from "node:path";

const EDITOR_ORIGIN = "https://office-editor.ziziyi.com";
const EDITOR_ROOT = join(process.cwd(), "out", "v9.3.0.24-1");
const TEXT_EXTENSIONS = new Set([".html", ".htm"]);
const CACHE_VERSION_SUFFIX = "-xinghuo-2";

const rewriteDirectory = async (directory) => {
  const entries = await readdir(directory, { withFileTypes: true });

  await Promise.all(
    entries.map(async (entry) => {
      const path = join(directory, entry.name);
      if (entry.isDirectory()) {
        await rewriteDirectory(path);
        return;
      }
      if (!TEXT_EXTENSIONS.has(extname(entry.name).toLowerCase())) return;

      const source = await readFile(path, "utf8");
      const rewritten = source.replaceAll(EDITOR_ORIGIN, "");
      if (rewritten !== source) await writeFile(path, rewritten, "utf8");
    })
  );
};

await rewriteDirectory(EDITOR_ROOT);

// Force an already-open deployment to discard editor HTML cached with the old
// cross-origin <base> URL. The service worker derives its main cache key only
// from the ONLYOFFICE version, which otherwise stays unchanged across deploys.
const serviceWorkerPath = join(EDITOR_ROOT, "document_editor_service_worker.js");
const serviceWorker = await readFile(serviceWorkerPath, "utf8");
const cacheKey = "g_cacheName=g_cacheNamePrefix+g_version";
if (!serviceWorker.includes(cacheKey)) {
  throw new Error("Unable to find the ONLYOFFICE service-worker cache key.");
}
await writeFile(
  serviceWorkerPath,
  serviceWorker.replace(cacheKey, `${cacheKey}+\"${CACHE_VERSION_SUFFIX}\"`),
  "utf8"
);
