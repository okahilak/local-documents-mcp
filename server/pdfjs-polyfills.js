// pdf.js (legacy build) needs DOMMatrix, ImageData and Path2D in Node. It evaluates
// `new DOMMatrix()` at module top level, so without them importing pdf.js throws
// "DOMMatrix is not defined". This module installs them from @napi-rs/canvas and must be
// imported BEFORE pdfjs-dist.
//
// The canvas package picks its native binary with a heuristic that, on Windows x64,
// switches to a "-gnu" build (which is not published) whenever Node reports itself as a
// shared library, as embedded runtimes can. We therefore point the loader at the exact
// bundled binary for this platform via NAPI_RS_NATIVE_LIBRARY_PATH.

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

/** Native binary packages shipped in the bundle, by `${platform}-${arch}`. */
export const CANVAS_BINARIES = {
  "win32-x64": { pkg: "@napi-rs/canvas-win32-x64-msvc", file: "skia.win32-x64-msvc.node" },
  "darwin-arm64": { pkg: "@napi-rs/canvas-darwin-arm64", file: "skia.darwin-arm64.node" },
};

/** Absolute path of the bundled canvas binary for a platform/arch, or null if not shipped. */
export function canvasBinaryPath(platform = process.platform, arch = process.arch) {
  const entry = CANVAS_BINARIES[`${platform}-${arch}`];
  if (!entry) return null;
  try {
    const p = path.join(path.dirname(require.resolve(`${entry.pkg}/package.json`)), entry.file);
    return fs.existsSync(p) ? p : null;
  } catch {
    return null;
  }
}

function loadCanvas() {
  const explicit = canvasBinaryPath();
  const hadEnv = Object.prototype.hasOwnProperty.call(process.env, "NAPI_RS_NATIVE_LIBRARY_PATH");
  const prevEnv = process.env.NAPI_RS_NATIVE_LIBRARY_PATH;
  if (explicit && !hadEnv) process.env.NAPI_RS_NATIVE_LIBRARY_PATH = explicit;
  try {
    return require("@napi-rs/canvas");
  } catch (err) {
    const shipped = Object.keys(CANVAS_BINARIES).join(", ");
    throw new Error(
      `Could not load @napi-rs/canvas for ${process.platform}-${process.arch} (bundled binaries: ${shipped}; ` +
        `binary for this platform ${explicit ? `at ${explicit}` : "not found"}): ${err?.message || err}`
    );
  } finally {
    if (!hadEnv) delete process.env.NAPI_RS_NATIVE_LIBRARY_PATH;
    else process.env.NAPI_RS_NATIVE_LIBRARY_PATH = prevEnv;
  }
}

const canvas = loadCanvas();
for (const name of ["DOMMatrix", "ImageData", "Path2D"]) {
  if (!globalThis[name]) {
    if (!canvas[name]) throw new Error(`@napi-rs/canvas does not provide ${name}`);
    globalThis[name] = canvas[name];
  }
}
