#!/bin/sh
# Tests the built LocalDocuments.mcpb in ISOLATION: unpacked into a temp directory with
# no node_modules in any parent, so module resolution cannot fall back to this project's
# development node_modules (which would hide missing runtime dependencies).
set -eu
cd "$(dirname "$0")/.."
ROOT=$(pwd)
NODE_BIN="${NODE_BIN:-node}"
T=$(mktemp -d "${TMPDIR:-/tmp}/ld-bundle-test.XXXXXX")
trap 'rm -rf "$T"' EXIT
npx mcpb unpack "$ROOT/LocalDocuments.mcpb" "$T/ext" > /dev/null

# 1. Isolation: no node_modules above the unpacked bundle.
d="$T"
while [ "$d" != "/" ]; do
  if [ -d "$d/node_modules" ]; then echo "Not isolated: $d/node_modules exists"; exit 1; fi
  d=$(dirname "$d")
done
echo "Isolated bundle at $T/ext ($("$NODE_BIN" -v), $(uname -s)-$(uname -m))"

# 2. Every engine loads from the bundle alone; pdf.js globals come from @napi-rs/canvas.
(cd "$T" && NODE_NO_WARNINGS=1 "$NODE_BIN" --input-type=module -e '
const ext = (await import("node:fs")).realpathSync(process.argv[1]);
const poly = await import(ext + "/server/pdfjs-polyfills.js");
for (const [plat, arch] of [["win32", "x64"], ["darwin", "arm64"]]) {
  const p = poly.canvasBinaryPath(plat, arch);
  if (!p || !p.startsWith(ext)) throw new Error(`bundled canvas binary for ${plat}-${arch} not found inside the bundle: ${p}`);
  console.log(`canvas binary for ${plat}-${arch}: ${p.slice(ext.length + 1)}`);
}
const pdf = await import(ext + "/server/pdf.js");
for (const g of ["DOMMatrix", "ImageData", "Path2D"]) if (typeof globalThis[g] !== "function") throw new Error(g + " missing");
await import(ext + "/server/xlsx.js");
await import(ext + "/server/docx.js");
console.log("engines load in isolation: pdf (DOMMatrix/ImageData/Path2D installed), xlsx, docx");
' "$T/ext")

# 3. Full test suite against the isolated bundle.
"$NODE_BIN" test/make-fixtures.mjs > /dev/null 2>&1
SERVER_ENTRY="$T/ext/server/index.js" "$NODE_BIN" --test test/e2e.test.mjs test/negotiation.test.mjs test/pdfjs-paths.test.mjs test/xlsx-edit.test.mjs
