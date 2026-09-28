// Allowed-directory handling and safe path resolution.
//
// Every tool path goes through resolveAllowedFile(), which:
//   1. expands "~", resolves relative paths against the allowed directories,
//   2. normalises ".." segments (path.resolve),
//   3. resolves symlinks/junctions with realpath, and
//   4. checks the *real* path is inside the *real* path of an allowed directory.
// Because the check runs on the fully resolved target, neither "../" tricks nor
// symlinks that point outside an allowed directory can escape.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const CASE_INSENSITIVE = process.platform === "win32" || process.platform === "darwin";
const MAX_FILE_BYTES = 512 * 1024 * 1024;

let allowedRoots = [];

function expandHome(p) {
  if (p === "~") return os.homedir();
  if (p.startsWith("~/") || p.startsWith("~\\")) return path.join(os.homedir(), p.slice(2));
  return p;
}

function norm(p) {
  return CASE_INSENSITIVE ? p.toLowerCase() : p;
}

function realpath(p) {
  // .native uses the OS resolver (handles Windows junctions / 8.3 names / subst drives).
  let r = fs.realpathSync.native(p);
  if (process.platform === "win32" && r.startsWith("\\\\?\\")) {
    r = r.startsWith("\\\\?\\UNC\\") ? "\\\\" + r.slice(8) : r.slice(4);
  }
  return r;
}

export function isInside(root, target) {
  const rel = path.relative(norm(root), norm(target));
  if (rel === "") return true;
  if (path.isAbsolute(rel)) return false; // different drive on Windows
  return rel !== ".." && !rel.startsWith(".." + path.sep);
}

/**
 * Configure allowed directories. Unexpanded "${user_config...}" placeholders and
 * non-existent directories are skipped (with a warning on stderr).
 */
export function setAllowedDirectories(dirs) {
  const roots = [];
  for (const raw of dirs) {
    if (!raw || typeof raw !== "string") continue;
    const trimmed = raw.trim();
    if (!trimmed || trimmed.startsWith("${")) continue;
    try {
      const real = realpath(path.resolve(expandHome(trimmed)));
      if (!fs.statSync(real).isDirectory()) {
        console.error(`[local-documents] Not a directory, ignoring: ${trimmed}`);
        continue;
      }
      if (!roots.some((r) => norm(r) === norm(real))) roots.push(real);
    } catch (err) {
      console.error(`[local-documents] Cannot access allowed directory, ignoring: ${trimmed} (${err.code || err.message})`);
    }
  }
  allowedRoots = roots;
  return roots;
}

export function getAllowedDirectories() {
  return [...allowedRoots];
}

export class PathError extends Error {}

/** Common checks for a user-supplied path; returns it with "~" expanded. */
function checkInput(input, label) {
  if (typeof input !== "string" || input.trim() === "") throw new PathError(`${label} must be a non-empty string.`);
  if (input.includes("\0")) throw new PathError(`${label} contains a NUL byte.`);
  const p = expandHome(input.trim());
  if (process.platform === "win32") {
    // Block NTFS alternate data streams ("file.pdf:stream") and device paths.
    const withoutDrive = p.replace(/^[a-zA-Z]:/, "");
    if (withoutDrive.includes(":")) throw new PathError(`${label} may not contain ':' outside the drive letter.`);
    if (p.startsWith("\\\\?\\") || p.startsWith("\\\\.\\")) throw new PathError("Device/namespace paths are not allowed.");
  }
  return p;
}

const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(\..*)?$/i;

// Names that are invalid or dangerous on Windows are rejected everywhere, so a
// file written on one OS can always be opened on the other.
export function assertPortableName(base, label) {
  if (!base || base === "." || base === ".." || WINDOWS_RESERVED.test(base) || /[<>:"|?*\x00-\x1f\\/]/.test(base) || /[. ]$/.test(base)) {
    throw new PathError(`Invalid ${label}: ${base}`);
  }
}

function deniedMessage(input) {
  return `Access denied: ${input} resolves outside the allowed directories.\nAllowed directories:\n${allowedRoots.map((r) => "  " + r).join("\n")}`;
}

/**
 * Resolve a user-supplied path to a real, readable regular file inside an allowed
 * directory. Returns { realPath, size, mtimeMs }.
 * @param {string} input
 * @param {string[]} extensions lower-case extensions including the dot
 */
export function resolveAllowedFile(input, extensions) {
  if (allowedRoots.length === 0) {
    throw new PathError(
      "No allowed directories are configured. Add at least one folder in the Local Documents extension settings."
    );
  }
  const p = checkInput(input, "Path");

  // Candidate absolute paths: absolute input as-is, or relative input under each allowed root.
  const candidates = path.isAbsolute(p) ? [path.resolve(p)] : allowedRoots.map((r) => path.resolve(r, p));

  let lastErr = null;
  for (const candidate of candidates) {
    let real;
    try {
      real = realpath(candidate);
    } catch (err) {
      lastErr = err;
      continue;
    }
    if (!allowedRoots.some((root) => isInside(root, real))) {
      throw new PathError(deniedMessage(input));
    }
    const st = fs.statSync(real);
    if (!st.isFile()) throw new PathError(`Not a regular file: ${input}`);
    const ext = path.extname(real).toLowerCase();
    if (!extensions.includes(ext)) {
      throw new PathError(`Unsupported file type "${ext || "(none)"}". Expected: ${extensions.join(", ")}`);
    }
    if (st.size > MAX_FILE_BYTES) throw new PathError(`File is too large (${st.size} bytes).`);
    return { realPath: real, size: st.size, mtimeMs: st.mtimeMs };
  }

  if (lastErr && lastErr.code === "ENOENT") throw new PathError(`File not found: ${input}`);
  if (lastErr && (lastErr.code === "EACCES" || lastErr.code === "EPERM")) throw new PathError(`Permission denied: ${input}`);
  throw new PathError(`Cannot access ${input}${lastErr ? ` (${lastErr.code || lastErr.message})` : ""}`);
}

// ---------- output paths (for tools that write files) ----------

export function samePath(a, b) {
  return norm(a) === norm(b);
}

/**
 * Resolve a user-supplied OUTPUT path for a file that will be written.
 *
 * - Relative paths resolve against `baseDir` (the source file's directory).
 * - The parent directory must already exist; it is resolved with realpath and must be
 *   inside an allowed directory (so ".." and symlinked directories cannot escape).
 * - The target itself must not be a symlink or a non-file. If it exists, it is only
 *   accepted when `overwrite` is true.
 * Returns { path, exists }.
 */
export function resolveAllowedOutput(input, { extensions, baseDir, overwrite = false }) {
  if (allowedRoots.length === 0) throw new PathError("No allowed directories are configured.");
  const p = checkInput(input, "Output path");
  const abs = path.isAbsolute(p) ? path.resolve(p) : path.resolve(baseDir, p);
  const base = path.basename(abs);
  const ext = path.extname(base).toLowerCase();
  if (!extensions.includes(ext)) throw new PathError(`Output file must have one of these extensions: ${extensions.join(", ")}`);
  assertPortableName(base, "output file name");

  let realParent;
  try {
    realParent = realpath(path.dirname(abs));
  } catch (err) {
    throw new PathError(`Output folder does not exist: ${path.dirname(abs)} (${err.code || err.message})`);
  }
  if (!fs.statSync(realParent).isDirectory()) throw new PathError(`Output folder is not a directory: ${path.dirname(abs)}`);
  if (!allowedRoots.some((root) => isInside(root, realParent))) {
    throw new PathError(
      `Access denied: output path ${input} is outside the allowed directories.\nAllowed directories:\n${allowedRoots.map((r) => "  " + r).join("\n")}`
    );
  }

  const target = path.join(realParent, base);
  let st = null;
  try {
    st = fs.lstatSync(target);
  } catch (err) {
    if (err.code !== "ENOENT") throw new PathError(`Cannot access output path ${input} (${err.code || err.message})`);
  }
  if (st) {
    if (st.isSymbolicLink()) throw new PathError(`Refusing to write through a symbolic link: ${input}`);
    if (!st.isFile()) throw new PathError(`Output path exists and is not a regular file: ${input}`);
    if (!overwrite) throw new PathError(`Output file already exists: ${target}. Pass overwrite=true to replace it, or choose another output_path.`);
  }
  return { path: target, exists: Boolean(st) };
}

/**
 * Re-check, immediately before the final rename, that the output's parent directory
 * still resolves inside an allowed directory and the target is not a symlink.
 */
export function assertOutputStillAllowed(target) {
  const realParent = realpath(path.dirname(target));
  if (!allowedRoots.some((root) => isInside(root, realParent)) || !samePath(path.join(realParent, path.basename(target)), target)) {
    throw new PathError(`Output location changed while writing and is no longer allowed: ${target}`);
  }
  try {
    if (fs.lstatSync(target).isSymbolicLink()) throw new PathError(`Refusing to write through a symbolic link: ${target}`);
  } catch (err) {
    if (err instanceof PathError) throw err;
  }
}

// ---------- files and folders (for the file-organising tools) ----------

function requireRoots() {
  if (allowedRoots.length === 0) {
    throw new PathError("No allowed directories are configured. Add at least one folder in the Local Documents extension settings.");
  }
}

function lstatOrNull(p) {
  try {
    return fs.lstatSync(p);
  } catch (err) {
    if (err.code === "ENOENT" || err.code === "ENOTDIR") return null;
    throw new PathError(`Cannot access ${p} (${err.code || err.message})`);
  }
}

function rootOf(realPath) {
  return allowedRoots.find((root) => isInside(root, realPath)) ?? null;
}

/**
 * Resolve an EXISTING file or folder to move or rename. The parent directory is
 * resolved with realpath and must be inside an allowed directory; the entry itself
 * is not followed, so a symbolic link is handled as the link, never its target.
 * An allowed directory itself is never accepted.
 * Returns { path, stat (lstat), root }.
 */
export function resolveAllowedEntry(input) {
  requireRoots();
  const p = checkInput(input, "Path");
  const candidates = path.isAbsolute(p) ? [path.resolve(p)] : allowedRoots.map((r) => path.resolve(r, p));

  let lastErr = null;
  for (const candidate of candidates) {
    const base = path.basename(candidate);
    if (!base) throw new PathError(`Not a file or folder: ${input}`);
    let target, stat;
    try {
      target = path.join(realpath(path.dirname(candidate)), base);
      stat = fs.lstatSync(target);
    } catch (err) {
      lastErr = err;
      continue;
    }
    if (allowedRoots.some((r) => samePath(r, target))) throw new PathError(`An allowed directory itself cannot be moved or renamed: ${target}`);
    const root = rootOf(path.dirname(target));
    if (!root) throw new PathError(deniedMessage(input));
    return { path: target, stat, root };
  }

  if (lastErr && (lastErr.code === "ENOENT" || lastErr.code === "ENOTDIR")) throw new PathError(`File or folder not found: ${input}`);
  if (lastErr && (lastErr.code === "EACCES" || lastErr.code === "EPERM")) throw new PathError(`Permission denied: ${input}`);
  throw new PathError(`Cannot access ${input}${lastErr ? ` (${lastErr.code || lastErr.message})` : ""}`);
}

/**
 * Resolve a NEW path (a destination or a folder to create). Relative input resolves
 * against `baseDir`. The nearest existing ancestor is resolved with realpath and must be
 * a directory inside an allowed directory; every name below it must be portable.
 * Nothing is created here: `missing` lists the absolute folders, outermost first, that
 * would have to be created for the parent to exist.
 * Returns { path, missing, stat } where `stat` is the lstat of an existing entry or null.
 */
export function resolveAllowedNewPath(input, { baseDir, label = "destination" }) {
  requireRoots();
  const p = checkInput(input, label[0].toUpperCase() + label.slice(1));
  const abs = path.isAbsolute(p) ? path.resolve(p) : path.resolve(baseDir, p);

  const names = [];
  let ancestor = abs;
  let realAncestor = null;
  while (realAncestor === null) {
    try {
      realAncestor = realpath(ancestor);
    } catch (err) {
      if (err.code !== "ENOENT") throw new PathError(`Cannot access ${ancestor} (${err.code || err.message})`);
      const up = path.dirname(ancestor);
      if (up === ancestor) throw new PathError(`Cannot resolve ${label}: ${input}`);
      names.unshift(path.basename(ancestor));
      ancestor = up;
    }
  }

  if (names.length === 0) {
    // The path exists: resolve its parent, but not the entry itself (it may be a link).
    const base = path.basename(abs);
    if (!base) throw new PathError(`Invalid ${label}: ${input}`);
    const parent = realpath(path.dirname(abs));
    const target = path.join(parent, base);
    if (!rootOf(parent)) throw new PathError(deniedMessage(input));
    return { path: target, missing: [], stat: fs.lstatSync(target) };
  }

  if (!rootOf(realAncestor)) throw new PathError(deniedMessage(input));
  if (!fs.statSync(realAncestor).isDirectory()) throw new PathError(`Not a folder: ${realAncestor}`);
  for (const n of names) assertPortableName(n, "name");
  const missing = [];
  let dir = realAncestor;
  for (const n of names.slice(0, -1)) missing.push((dir = path.join(dir, n)));
  const target = path.join(dir, names.at(-1));
  return { path: target, missing, stat: lstatOrNull(target) };
}

/** Re-check just before a write that `dir` still resolves to itself inside an allowed directory. */
export function assertFolderStillAllowed(dir) {
  const real = realpath(dir);
  if (!samePath(real, dir) || !rootOf(real)) throw new PathError(`Folder location changed and is no longer allowed: ${dir}`);
}
