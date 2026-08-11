import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { brotliDecompress, constants, gzip } from "node:zlib";

const decompress = promisify(brotliDecompress);
const compressGzip = promisify(gzip);
const publicRoot = resolve("public");
const sourceRoot = join(publicRoot, "x2t-1");
const outputRoot = join(publicRoot, "x2t-compat");

await rm(outputRoot, { recursive: true, force: true });
await mkdir(outputRoot, { recursive: true });

const sourceScript = await decompress(await readFile(join(sourceRoot, "x2t.js")));
const sourceWasm = await decompress(await readFile(join(sourceRoot, "x2t.wasm")));
if (
  !sourceScript.toString("utf8").includes("var Module") ||
  !sourceWasm.subarray(0, 4).equals(Buffer.from([0, 97, 115, 109]))
) {
  throw new Error("Generated x2t compatibility assets are invalid");
}

const [script, wasm] = await Promise.all(
  [sourceScript, sourceWasm].map((content) =>
    compressGzip(content, { level: constants.Z_BEST_COMPRESSION }),
  ),
);
await Promise.all([
  writeFile(join(outputRoot, "x2t.js"), script),
  writeFile(join(outputRoot, "x2t.wasm"), wasm),
]);

console.log(
  `Prepared Android WebView x2t gzip assets (${script.length} byte script, ${wasm.length} byte wasm)`,
);
