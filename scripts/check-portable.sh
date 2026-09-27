#!/bin/sh
# Verifies the bundle's native code is exactly what we intend to ship:
#   - REQUIRED: one @napi-rs/canvas binary per supported platform (CANVAS_PLATFORMS),
#     with the right binary format and a version matching the canvas loader package.
#   - FORBIDDEN: any other native binary, OS/CPU-restricted package, or install script.
set -eu
DIR="$1"
PLATFORMS="${CANVAS_PLATFORMS:-win32-x64-msvc darwin-arm64}"
NM="$DIR/node_modules"
FAIL=0

CANVAS_VERSION=$(node -p 'require(process.argv[1]).version' "$NM/@napi-rs/canvas/package.json" 2>/dev/null || echo "")
if [ -z "$CANVAS_VERSION" ]; then echo "MISSING: @napi-rs/canvas loader package"; FAIL=1; fi

ALLOWED=""
for plat in $PLATFORMS; do
  pkgdir="$NM/@napi-rs/canvas-$plat"
  bin="$pkgdir/skia.$plat.node"
  if [ ! -f "$bin" ]; then echo "MISSING: $bin"; FAIL=1; continue; fi
  v=$(node -p 'require(process.argv[1]).version' "$pkgdir/package.json")
  [ "$v" = "$CANVAS_VERSION" ] || { echo "VERSION MISMATCH: canvas-$plat@$v vs @napi-rs/canvas@$CANVAS_VERSION"; FAIL=1; }
  desc=$(file -b "$bin")
  case "$plat" in
    win32-x64-msvc) echo "$desc" | grep -q "PE32+ executable.*DLL.*x86-64" || { echo "BAD FORMAT for $plat: $desc"; FAIL=1; } ;;
    win32-arm64-msvc) echo "$desc" | grep -q "PE32+ executable.*DLL.*Aarch64" || { echo "BAD FORMAT for $plat: $desc"; FAIL=1; } ;;
    darwin-arm64) echo "$desc" | grep -q "Mach-O 64-bit.*arm64" || { echo "BAD FORMAT for $plat: $desc"; FAIL=1; } ;;
    darwin-x64) echo "$desc" | grep -q "Mach-O 64-bit.*x86_64" || { echo "BAD FORMAT for $plat: $desc"; FAIL=1; } ;;
    *) echo "UNKNOWN PLATFORM $plat"; FAIL=1 ;;
  esac
  echo "OK  @napi-rs/canvas-$plat@$v  $(du -h "$bin" | cut -f1)  $desc"
  ALLOWED="$ALLOWED $bin"
done

is_allowed() { for a in $ALLOWED; do [ "$1" = "$a" ] && return 0; done; return 1; }

# Any other native file (by extension or by binary format) is an error.
find "$DIR" -type f \( -name "*.node" -o -name "*.dylib" -o -name "*.so" -o -name "*.so.*" -o -name "*.dll" \
  -o -name "*.exe" -o -name "*.a" -o -name "*.lib" -o -name "binding.gyp" \) > "${TMPDIR:-/tmp}/ld-native.$$"
find "$DIR" -type f -size +1k -exec file {} + | grep -E "Mach-O|ELF|PE32|MS-DOS executable" | cut -d: -f1 >> "${TMPDIR:-/tmp}/ld-native.$$" || true
for f in $(sort -u "${TMPDIR:-/tmp}/ld-native.$$"); do
  is_allowed "$f" || { echo "UNEXPECTED NATIVE FILE: $f"; FAIL=1; }
done
rm -f "${TMPDIR:-/tmp}/ld-native.$$"

# Other @napi-rs/canvas platform packages must not sneak in.
for d in "$NM"/@napi-rs/canvas-*; do
  [ -d "$d" ] || continue
  plat=${d##*/canvas-}
  case " $PLATFORMS " in *" $plat "*) ;; *) echo "UNEXPECTED PLATFORM PACKAGE: $d"; FAIL=1 ;; esac
done

# OS/CPU-restricted packages (other than the allowed canvas platforms) and install scripts.
PLATFORMS="$PLATFORMS" node --input-type=module -e '
import fs from "node:fs"; import path from "node:path";
const root = process.argv[1]; let bad = 0;
const allowed = new Set(process.env.PLATFORMS.split(/\s+/).filter(Boolean).map((p) => "@napi-rs/canvas-" + p));
const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) {
  const p = path.join(d, e.name);
  if (e.isDirectory()) { walk(p); continue; }
  if (e.name !== "package.json" || !p.includes("node_modules")) continue;
  let j; try { j = JSON.parse(fs.readFileSync(p, "utf8")); } catch { continue; }
  if (!j.name) continue;
  const s = j.scripts || {};
  if ((j.os || j.cpu) && !allowed.has(j.name)) { console.log(`UNEXPECTED platform-restricted package: ${j.name} os=${j.os} cpu=${j.cpu}`); bad = 1; }
  for (const k of ["preinstall", "install", "postinstall"]) if (s[k]) { console.log(`INSTALL SCRIPT in ${j.name}: ${k}=${s[k]}`); bad = 1; }
  if (j.gypfile) { console.log(`GYP package: ${j.name}`); bad = 1; }
}};
walk(root); process.exit(bad);
' "$DIR" || FAIL=1

echo "WebAssembly modules: $(find "$DIR" -name "*.wasm" | sed "s|$DIR/||" | tr '\n' ' ')"
if [ "$FAIL" -ne 0 ]; then echo "Native-code check FAILED"; exit 1; fi
echo "Native-code check passed: only the intended @napi-rs/canvas binaries ($PLATFORMS)."
