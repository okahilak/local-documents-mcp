// File organisation: list folders, create folders, and move/rename files and folders
// inside the allowed directories.
//
// Safety rules:
//   - Both ends of every move are resolved through paths.js, so neither ".." nor
//     symbolic links can reach outside the allowed directories. A link being moved is
//     moved as a link; its target is never touched.
//   - Nothing is ever overwritten or deleted. An existing destination is an error
//     (except a case-only rename of the same entry).
//   - A batch is all-or-nothing: moves run in order, and if one fails, the moves
//     already done are undone in reverse and the folders created for them removed.
// Everything here is synchronous, so two tool calls can never interleave.

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import {
  PathError,
  allowedRootOf,
  assertFolderStillAllowed,
  assertPortableName,
  getAllowedDirectories,
  isInside,
  resolveAllowedDirectory,
  resolveAllowedEntry,
  resolveAllowedNewPath,
  samePath,
} from "./paths.js";

export const MAX_BATCH = 1000;

const endsWithSeparator = (p) => /[\\/]$/.test(p.trim());

function lstatOrNull(p) {
  try {
    return fs.lstatSync(p);
  } catch (err) {
    if (err.code === "ENOENT" || err.code === "ENOTDIR") return null;
    throw err;
  }
}

function kind(stat) {
  if (stat.isSymbolicLink()) return "link";
  return stat.isDirectory() ? "folder" : "file";
}

// ---------- list_directory ----------

export const MAX_LIST_ENTRIES = 10000;
const JUNK = new Set(["thumbs.db", "desktop.ini"]);
const isHidden = (name) => name.startsWith(".") || JUNK.has(name.toLowerCase());
const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

function formatSize(n) {
  if (n < 1024) return `${n} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let i = -1;
  do (n /= 1024), i++;
  while (n >= 1024 && i < units.length - 1);
  return `${n < 10 ? n.toFixed(1) : Math.round(n)} ${units[i]}`;
}

function formatDate(ms) {
  const d = new Date(ms);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Sort key compare: path segment by segment, so a folder's contents follow it. */
function comparePaths(a, b) {
  const x = a.split("/");
  const y = b.split("/");
  for (let i = 0; i < Math.min(x.length, y.length); i++) {
    const c = collator.compare(x[i], y[i]) || (x[i] < y[i] ? -1 : x[i] > y[i] ? 1 : 0);
    if (c) return c;
  }
  return x.length - y.length;
}

function readNames(dir, includeHidden) {
  return fs.readdirSync(dir).filter((n) => includeHidden || !isHidden(n));
}

/**
 * List a folder (or, with no path, every allowed directory). Entries are shown with
 * paths relative to the allowed directory they are in, so they can be passed to move,
 * rename and move_batch as they are. Symbolic links are listed but never followed.
 * Recursive listings go breadth-first, so a truncated listing still shows the top levels.
 */
export function listDirectory(input, { recursive = false, maxDepth = 5, includeHidden = false, maxEntries = 1000 } = {}) {
  const dirs = input === undefined || input.trim() === "" ? getAllowedDirectories() : [resolveAllowedDirectory(input)];
  if (dirs.length === 0) throw new PathError("No allowed directories are configured. Add at least one folder in the Local Documents extension settings.");
  const depthLimit = recursive ? maxDepth : 1;
  let budget = maxEntries;
  const sections = [];

  for (const dir of dirs) {
    const root = allowedRootOf(dir);
    const rel = (p) => path.relative(root, p).split(path.sep).join("/");
    const entries = [];
    let files = 0, folders = 0, bytes = 0, truncated = false;
    const queue = [{ dir, depth: 1 }];
    const folderLine = new Map(); // folder path -> entry, to note unlisted contents

    while (queue.length) {
      const { dir: d, depth } = queue.shift();
      let names;
      try {
        names = readNames(d, includeHidden).sort(collator.compare);
      } catch (err) {
        const e = folderLine.get(d);
        if (e) e.note = `cannot read: ${err.code || err.message}`;
        else throw new PathError(`Cannot read folder ${d} (${err.code || err.message})`);
        continue;
      }
      for (const name of names) {
        if (budget === 0) {
          truncated = true;
          break;
        }
        const p = path.join(d, name);
        let st;
        try {
          st = fs.lstatSync(p);
        } catch (err) {
          entries.push({ path: rel(p), line: `${rel(p)}  (cannot read: ${err.code || err.message})` });
          budget--;
          continue;
        }
        budget--;
        if (st.isSymbolicLink()) {
          entries.push({ path: rel(p), line: `${rel(p)}  (link, not followed)` });
        } else if (st.isDirectory()) {
          folders++;
          const e = { path: rel(p), dir: p, mtime: st.mtimeMs };
          entries.push(e);
          folderLine.set(p, e);
          if (depth < depthLimit) queue.push({ dir: p, depth: depth + 1 });
          else e.unlisted = true;
        } else {
          files++;
          bytes += st.size;
          entries.push({ path: rel(p), line: `${rel(p)}  ${formatSize(st.size)}  ${formatDate(st.mtimeMs)}` });
        }
      }
      if (truncated) break;
    }
    if (truncated) for (const e of folderLine.values()) if (!e.unlisted && queue.some((q) => q.dir === e.dir)) e.unlisted = true;

    for (const e of entries) {
      if (e.line) continue;
      let note = e.note;
      if (!note && e.unlisted) {
        try {
          const n = readNames(e.dir, includeHidden).length;
          note = n ? `${n} item${n === 1 ? "" : "s"}, not listed` : "empty";
        } catch (err) {
          note = `cannot read: ${err.code || err.message}`;
        }
      }
      e.line = `${e.path}/  ${note ? `(${note})  ` : ""}${formatDate(e.mtime)}`;
    }
    entries.sort((a, b) => comparePaths(a.path, b.path));

    const where = dir === root ? root : `${dir} (in allowed directory ${root})`;
    let head = `${where}\n${folders} folder${folders === 1 ? "" : "s"}, ${files} file${files === 1 ? "" : "s"} (${formatSize(bytes)})`;
    if (recursive) head += `, up to ${depthLimit} level${depthLimit === 1 ? "" : "s"} deep`;
    head += `. Paths are relative to ${root}.`;
    const body = entries.length ? entries.map((e) => e.line).join("\n") : "(empty)";
    let section = `${head}\n${body}`;
    if (truncated) section += `\n[Listing stopped at ${maxEntries} entries. List a subfolder, lower max_depth, or raise max_entries (up to ${MAX_LIST_ENTRIES}).]`;
    sections.push(section);
    if (truncated) break;
  }
  if (sections.length < dirs.length) sections.push(`[${dirs.length - sections.length} more allowed director${dirs.length - sections.length === 1 ? "y" : "ies"} not listed.]`);
  return sections.join("\n\n");
}

// ---------- create_folder ----------

export function createFolder(input, { parents = false } = {}) {
  // Relative folder paths resolve against the first allowed directory.
  const target = resolveAllowedNewPath(input, { baseDir: getAllowedDirectories()[0] ?? ".", label: "folder path" });
  if (target.stat) {
    if (target.stat.isDirectory()) return `Folder already exists: ${target.path}`;
    throw new PathError(`A ${kind(target.stat)} already exists at ${target.path}.`);
  }
  assertPortableName(path.basename(target.path), "folder name");
  if (target.missing.length && !parents) {
    throw new PathError(`Parent folder does not exist: ${target.missing[0]}. Pass parents=true to create it too.`);
  }
  const created = makeFolders([...target.missing, target.path]);
  return `Created folder${created.length > 1 ? "s" : ""}:\n${created.map((d) => "- " + d).join("\n")}`;
}

/** Create folders in order; returns the ones actually created (for undo). */
function makeFolders(dirs) {
  const created = [];
  try {
    for (const dir of dirs) {
      assertFolderStillAllowed(path.dirname(dir));
      try {
        fs.mkdirSync(dir);
        created.push(dir);
      } catch (err) {
        if (err.code !== "EEXIST" || !fs.lstatSync(dir).isDirectory()) throw err;
      }
    }
  } catch (err) {
    removeFolders(created);
    throw err;
  }
  return created;
}

function removeFolders(dirs) {
  const left = [];
  for (const dir of [...dirs].reverse()) {
    try {
      fs.rmdirSync(dir); // only succeeds while empty
    } catch {
      left.push(dir);
    }
  }
  return left;
}

// ---------- move / rename / move_batch ----------

/**
 * Validate one move against the disk.
 * `to` ending in "/" or "\" means "into this folder, keeping the name".
 */
function planMove(from, to, { createFolders }) {
  const src = resolveAllowedEntry(from);
  const baseDir = src.root;
  let dest, missing, destStat;

  if (endsWithSeparator(to)) {
    const folder = resolveAllowedNewPath(to, { baseDir, label: "destination folder" });
    if (folder.stat && (folder.stat.isSymbolicLink() || !folder.stat.isDirectory())) {
      throw new PathError(`Destination is not a folder: ${folder.path}`);
    }
    dest = path.join(folder.path, path.basename(src.path));
    missing = folder.stat ? [] : [...folder.missing, folder.path];
    destStat = folder.stat ? lstatOrNull(dest) : null;
  } else {
    const d = resolveAllowedNewPath(to, { baseDir });
    dest = d.path;
    missing = d.missing;
    destStat = d.stat;
    assertPortableName(path.basename(dest), "destination name");
  }

  if (src.path === dest) throw new PathError(`Source and destination are the same: ${src.path}`);
  const caseOnly = samePath(src.path, dest) && destStat && destStat.ino === src.stat.ino && destStat.dev === src.stat.dev;
  if (destStat && !caseOnly) {
    const hint = destStat.isDirectory() && !destStat.isSymbolicLink() ? ` To move into that folder, end the destination with "/".` : "";
    throw new PathError(`Destination already exists (${kind(destStat)}): ${dest}. Nothing is overwritten.${hint}`);
  }
  if (src.stat.isDirectory() && isInside(src.path, dest)) throw new PathError(`Cannot move a folder into itself: ${src.path} → ${dest}`);
  if (missing.length && !createFolders) {
    throw new PathError(`Destination folder does not exist: ${missing[0]}. Create it with create_folder, or pass create_folders=true.`);
  }
  return { from: src.path, to: dest, kind: kind(src.stat), missing, caseOnly };
}

/** rename(2), falling back to copy + delete between volumes. */
function relocate(from, to, { caseOnly = false } = {}) {
  assertFolderStillAllowed(path.dirname(to));
  if (!caseOnly && lstatOrNull(to)) throw new PathError(`Destination appeared while moving: ${to}. Nothing is overwritten.`);
  try {
    fs.renameSync(from, to);
    return;
  } catch (err) {
    if (err.code !== "EXDEV") throw err;
  }
  // Different volume: copy to a temporary name next to the destination, move that into
  // place (same volume, atomic), and only then remove the original.
  const tmp = path.join(path.dirname(to), `.~move-${crypto.randomBytes(6).toString("hex")}`);
  try {
    fs.cpSync(from, tmp, { recursive: true, errorOnExist: true, force: false, preserveTimestamps: true, verbatimSymlinks: true });
    if (lstatOrNull(to)) throw new PathError(`Destination appeared while moving: ${to}. Nothing is overwritten.`);
    fs.renameSync(tmp, to);
  } catch (err) {
    fs.rmSync(tmp, { recursive: true, force: true });
    throw err;
  }
  try {
    fs.rmSync(from, { recursive: true });
  } catch (err) {
    throw Object.assign(new PathError(`Copied ${from} to ${to}, but the original could not be fully removed (${err.code || err.message}). Both may now exist.`), {
      partial: true,
    });
  }
}

/**
 * Run moves in order as one transaction. Each move is validated against the disk as
 * left by the moves before it, so a batch may, e.g., move a file into a folder that an
 * earlier move renamed. On any failure, completed moves are undone.
 */
export function moveEntries(moves, { createFolders = false } = {}) {
  if (!Array.isArray(moves) || moves.length === 0) throw new PathError("No moves given.");
  if (moves.length > MAX_BATCH) throw new PathError(`Too many moves (${moves.length}); the limit is ${MAX_BATCH} per call.`);

  const done = [];
  for (const [i, m] of moves.entries()) {
    let created = [];
    try {
      const plan = planMove(m.from, m.to, { createFolders });
      created = makeFolders(plan.missing);
      relocate(plan.from, plan.to, plan);
      done.push({ ...plan, created });
    } catch (err) {
      if (!err.partial) removeFolders(created);
      const label = moves.length > 1 ? `Move ${i + 1} of ${moves.length} failed` : "Move failed";
      const problems = err.partial ? [] : undo(done);
      let msg = `${label} (${m.from} → ${m.to}): ${err.message}`;
      if (err.partial) msg += done.length ? `\nThe ${done.length} earlier move(s) were kept.` : "";
      else if (problems.length) msg += `\nUndoing the earlier moves was incomplete:\n${problems.map((p) => "- " + p).join("\n")}`;
      else msg += done.length ? `\nThe ${done.length} earlier move(s) were undone; nothing was changed.` : "\nNothing was changed.";
      throw new Error(msg);
    }
  }

  return `Moved ${describe(done, done.flatMap((d) => d.created), "Created")}`;
}

function describe(plans, folders, verb) {
  const s = (n, one, many) => `${n} ${n === 1 ? one : many}`;
  let out = `${s(plans.length, "item", "items")}:\n${plans.map((d) => `- ${d.from} → ${d.to}${d.kind === "file" ? "" : ` (${d.kind})`}`).join("\n")}`;
  if (folders.length) out += `\n${verb} folder${folders.length === 1 ? "" : "s"}:\n${folders.map((d) => "- " + d).join("\n")}`;
  return out;
}

function undo(done) {
  const problems = [];
  for (const d of [...done].reverse()) {
    try {
      relocate(d.to, d.from, d);
    } catch (err) {
      problems.push(`could not move ${d.to} back to ${d.from}: ${err.message}`);
      continue;
    }
    for (const dir of removeFolders(d.created)) problems.push(`created folder left in place: ${dir}`);
  }
  return problems;
}

export function renameEntry(input, newName) {
  if (typeof newName !== "string" || newName.trim() !== newName || /[\\/]/.test(newName)) {
    throw new PathError(`new_name must be a plain name without folders or surrounding spaces: ${JSON.stringify(newName)}. Use move to change folders.`);
  }
  assertPortableName(newName, "new name");
  const src = resolveAllowedEntry(input);
  return moveEntries([{ from: src.path, to: path.join(path.dirname(src.path), newName) }]);
}
