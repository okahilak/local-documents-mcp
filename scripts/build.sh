#!/bin/sh
# Builds LocalDocuments.mcpb from a clean staging directory.
set -eu
cd "$(dirname "$0")/.."
ROOT=$(pwd)
STAGE="$ROOT/build/bundle"
OUT="$ROOT/LocalDocuments.mcpb"

rm -rf "$STAGE" "$OUT"
mkdir -p "$STAGE"
cp -R manifest.json icon.png package.json package-lock.json server "$STAGE/"
cp README.md LICENSE "$STAGE/" 2>/dev/null || true

# Production deps only. --omit=optional skips whatever native binaries npm would pick for
# the *build* machine; --ignore-scripts guarantees nothing compiles or downloads at install.
(cd "$STAGE" && npm ci --omit=dev --omit=optional --ignore-scripts --no-audit --no-fund)

# @napi-rs/canvas (DOMMatrix/ImageData/Path2D for pdf.js) ships its native code in
# per-platform packages. Add exactly the platforms we support, regardless of the build
# machine, each verified against the sha512 integrity recorded in package-lock.json.
CANVAS_PLATFORMS="win32-x64-msvc darwin-arm64"
CANVAS_VERSION=$(node -p 'require(process.argv[1]).version' "$STAGE/node_modules/@napi-rs/canvas/package.json")
PACKTMP=$(mktemp -d "${TMPDIR:-/tmp}/ld-canvas.XXXXXX")
for plat in $CANVAS_PLATFORMS; do
  pkg="@napi-rs/canvas-$plat"
  expected=$(node -p 'const l=require(process.argv[1]); const e=l.packages["node_modules/"+process.argv[2]]; if(!e||e.version!==process.argv[3]) throw new Error("lockfile has no "+process.argv[2]+"@"+process.argv[3]); e.integrity' \
    "$ROOT/package-lock.json" "$pkg" "$CANVAS_VERSION")
  actual=$(cd "$PACKTMP" && npm pack "$pkg@$CANVAS_VERSION" --json --silent | node -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>console.log(JSON.parse(s)[0].integrity))')
  if [ "$expected" != "$actual" ]; then echo "Integrity mismatch for $pkg: lockfile $expected, registry $actual"; exit 1; fi
  mkdir -p "$STAGE/node_modules/$pkg"
  tar -xzf "$PACKTMP/napi-rs-canvas-$plat-$CANVAS_VERSION.tgz" -C "$STAGE/node_modules/$pkg" --strip-components=1
  echo "Bundled $pkg@$CANVAS_VERSION ($actual)"
done
rm -rf "$PACKTMP"

# Prune files not needed at runtime.
NM="$STAGE/node_modules"
rm -rf "$NM/pdfjs-dist/build" "$NM/pdfjs-dist/web" "$NM/pdfjs-dist/image_decoders" "$NM/pdfjs-dist/types" \
       "$NM/pdfjs-dist/legacy/web" "$NM/pdfjs-dist/legacy/image_decoders" \
       "$NM/pdfjs-dist/legacy/build/pdf.sandbox"* "$NM/pdfjs-dist/legacy/build/"*.min.mjs \
       "$NM/@hyzyla/pdfium/dist/worker" "$NM/@hyzyla/pdfium/dist/worker.bundle.js" \
       "$NM/@hyzyla/pdfium/dist/index.esm.worker.js" "$NM/@hyzyla/pdfium/dist/index.esm.cdn.js" \
       "$NM/@hyzyla/pdfium/dist/index.esm.browser.js" "$NM/@hyzyla/pdfium/dist/index.esm.base64.js" \
       "$NM/@hyzyla/pdfium/dist/pdfium.wasm.base64-"*.js \
       "$NM/exceljs/dist" \
       "$NM/officeparser/dist/officeparser.browser"* "$NM/officeparser/dist/cli.js"

# officeparser loads tesseract.js (OCR) only via dynamic import when OCR is enabled.
# We never enable OCR, so drop it and the packages reachable only through it
# (about 45 MB, including a postinstall script).
for p in tesseract.js tesseract.js-core bmp-js idb-keyval is-url node-fetch encoding whatwg-url tr46 \
         webidl-conversions opencollective-postinstall regenerator-runtime wasm-feature-detect zlibjs; do
  rm -rf "$NM/$p"
done
find "$NM" \( -name "*.map" -o -name "*.d.ts" -o -name "*.d.mts" -o -name "*.d.cts" -o -name "*.md" -o -name "*.markdown" \
       -o -name ".github" -o -name "test" -o -name "tests" -o -name "__tests__" -o -name "docs" -o -name "example" -o -name "examples" \) \
       -prune -exec rm -rf {} +

CANVAS_PLATFORMS="$CANVAS_PLATFORMS" sh scripts/check-portable.sh "$STAGE"

npx mcpb validate "$STAGE/manifest.json"
npx mcpb pack "$STAGE" "$OUT"
echo "Built: $OUT"
