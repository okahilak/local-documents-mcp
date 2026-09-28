// File organisation tools: list_directory, create_folder, rename, move, move_batch.
// Proves moves stay inside the allowed directories, never overwrite, and that a
// failing batch leaves the tree exactly as it was.
// SERVER_ENTRY env var can point at an unpacked bundle's server/index.js.
import { test, before, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { connectLikeClaude } from "./claude-like-client.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const entry = path.resolve(process.env.SERVER_ENTRY || path.join(here, "..", "server", "index.js"));
fs.mkdirSync(path.join(here, "fixtures", "files"), { recursive: true });
const root = fs.realpathSync(path.join(here, "fixtures", "files"));
const allowed = path.join(root, "allowed");
const outside = path.join(root, "outside");

const text = (r) => r.content.filter((c) => c.type === "text").map((c) => c.text).join("\n");
const at = (...p) => path.join(allowed, ...p);
const exists = (p) => fs.existsSync(p);

/** Every path under dir (relative), with file contents, so trees can be compared exactly. */
function snapshot(dir) {
  const out = {};
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      const rel = path.relative(dir, p);
      if (e.isSymbolicLink()) out[rel] = "-> " + fs.readlinkSync(p);
      else if (e.isDirectory()) (out[rel + "/"] = true), walk(p);
      else out[rel] = fs.readFileSync(p, "utf8");
    }
  };
  walk(dir);
  return out;
}

let client;
const call = async (name, args) => {
  const r = await client.callTool({ name, arguments: args });
  return { ok: !r.isError, text: text(r) };
};
const ok = async (name, args) => {
  const r = await call(name, args);
  assert.ok(r.ok, r.text);
  return r.text;
};
const fails = async (name, args, pattern) => {
  const r = await call(name, args);
  assert.ok(!r.ok, `expected ${name} to fail: ${r.text}`);
  assert.match(r.text, pattern);
  return r.text;
};

before(async () => {
  fs.mkdirSync(allowed, { recursive: true });
  ({ client } = await connectLikeClaude(process.execPath, [entry, allowed]));
});
after(async () => client?.close());

beforeEach(() => {
  // Empty the allowed directory in place: the server resolved it once at startup.
  for (const e of fs.readdirSync(allowed)) fs.rmSync(path.join(allowed, e), { recursive: true, force: true });
  fs.rmSync(outside, { recursive: true, force: true });
  fs.mkdirSync(at("Inbox", "deep"), { recursive: true });
  fs.mkdirSync(at("Archive"));
  fs.mkdirSync(outside);
  fs.writeFileSync(at("Inbox", "a.pdf"), "A");
  fs.writeFileSync(at("Inbox", "b.docx"), "B");
  fs.writeFileSync(at("Inbox", "notes.txt"), "N");
  fs.writeFileSync(at("Inbox", "deep", "c.xlsx"), "C");
  fs.writeFileSync(at("Archive", "a.pdf"), "old A");
  fs.writeFileSync(path.join(outside, "secret.txt"), "S");
});

test("the organising tools are listed as writable but not destructive", async () => {
  const { tools } = await client.listTools();
  assert.equal(tools.find((t) => t.name === "list_directory").annotations.readOnlyHint, true);
  for (const name of ["create_folder", "rename", "move", "move_batch"]) {
    const t = tools.find((x) => x.name === name);
    assert.ok(t, name);
    assert.equal(t.annotations.readOnlyHint, false, name);
    assert.equal(t.annotations.destructiveHint, false, name);
  }
});

test("list_directory: flat, recursive, hidden files, links and limits", async () => {
  fs.writeFileSync(at(".DS_Store"), "");
  fs.writeFileSync(at("Inbox", "file10.txt"), "1234567890");
  fs.writeFileSync(at("Inbox", "file2.txt"), "");
  fs.symlinkSync(outside, at("escape"), "dir");

  // No path: every allowed directory, one level, hidden files skipped.
  let t = await ok("list_directory", {});
  assert.ok(t.startsWith(allowed + "\n"), t);
  assert.match(t, /^Archive\/  \(1 item, not listed\)/m);
  assert.match(t, /^Inbox\/  \(6 items, not listed\)/m);
  assert.match(t, /^escape  \(link, not followed\)$/m);
  assert.doesNotMatch(t, /DS_Store|secret/);
  assert.match(await ok("list_directory", { include_hidden: true }), /^\.DS_Store  0 B/m);

  // A subfolder: paths stay relative to the allowed directory, ready for move.
  t = await ok("list_directory", { path: "Inbox" });
  assert.match(t, /in allowed directory/);
  const lines = t.split("\n").slice(2);
  assert.deepEqual(lines.map((l) => l.split("  ")[0]), ["Inbox/a.pdf", "Inbox/b.docx", "Inbox/deep/", "Inbox/file2.txt", "Inbox/file10.txt", "Inbox/notes.txt"]);
  assert.match(t, /^Inbox\/file10\.txt  10 B  \d{4}-\d\d-\d\d \d\d:\d\d$/m);

  // Recursive, with depth limit and truncation.
  t = await ok("list_directory", { recursive: true });
  assert.match(t, /^Inbox\/deep\/c\.xlsx  1 B/m);
  assert.match(t, /^Inbox\/deep\/  \d{4}/m);
  t = await ok("list_directory", { recursive: true, max_depth: 1 });
  assert.doesNotMatch(t, /c\.xlsx/);
  t = await ok("list_directory", { recursive: true, max_entries: 3 });
  assert.match(t, /Listing stopped at 3 entries/);
  assert.equal(t.split("\n").filter((l) => /^\S/.test(l) && !l.startsWith("[") && !l.startsWith(allowed) && !/^\d+ folder/.test(l)).length, 3);

  // Same boundaries as every other tool.
  await fails("list_directory", { path: outside }, /Access denied/);
  await fails("list_directory", { path: "escape" }, /Access denied/);
  await fails("list_directory", { path: "../outside" }, /Access denied/);
  await fails("list_directory", { path: "Inbox/a.pdf" }, /Not a folder/);
  await fails("list_directory", { path: "Nope" }, /Folder not found/);
});

test("create_folder: absolute, relative, existing, parents", async () => {
  await ok("create_folder", { path: at("Projects") });
  assert.ok(fs.statSync(at("Projects")).isDirectory());
  assert.match(await ok("create_folder", { path: "Projects" }), /already exists/);
  await fails("create_folder", { path: "Clients/Acme/2024" }, /Parent folder does not exist.*Clients/);
  assert.ok(!exists(at("Clients")));
  const t = await ok("create_folder", { path: "Clients/Acme/2024", parents: true });
  assert.match(t, /Created folders/);
  assert.ok(fs.statSync(at("Clients", "Acme", "2024")).isDirectory());
  await fails("create_folder", { path: "Inbox/a.pdf" }, /file already exists/);
  await fails("create_folder", { path: "bad:name" }, /Invalid|may not contain/);
  await fails("create_folder", { path: "CON" }, /Invalid name: CON/);
});

test("create_folder and move refuse to leave the allowed directories", async () => {
  await fails("create_folder", { path: path.join(outside, "x") }, /Access denied/);
  await fails("create_folder", { path: "../outside/x" }, /Access denied/);
  await fails("move", { from: "Inbox/a.pdf", to: "../outside/a.pdf" }, /Access denied/);
  await fails("move", { from: path.join(outside, "secret.txt"), to: "secret.txt" }, /Access denied/);
  // A symlinked folder that points outside is not a way out.
  fs.symlinkSync(outside, at("escape"), "dir");
  await fails("move", { from: "Inbox/a.pdf", to: "escape/" }, /Destination is not a folder|Access denied/);
  await fails("move", { from: "Inbox/a.pdf", to: "escape/a.pdf" }, /Access denied/);
  await fails("move", { from: "escape/secret.txt", to: "secret.txt" }, /Access denied/);
  await fails("create_folder", { path: "escape/x" }, /Access denied/);
  assert.deepEqual(fs.readdirSync(outside), ["secret.txt"]);
  // The allowed directory itself cannot be moved.
  await fails("move", { from: allowed, to: "moved" }, /allowed directory itself/);
});

test("rename: in place, case-only, conflicts and bad names", async () => {
  await ok("rename", { path: "Inbox/b.docx", new_name: "Brief.docx" });
  assert.equal(fs.readFileSync(at("Inbox", "Brief.docx"), "utf8"), "B");
  assert.ok(!exists(at("Inbox", "b.docx")));
  await ok("rename", { path: "Inbox/notes.txt", new_name: "NOTES.txt" });
  assert.ok(fs.readdirSync(at("Inbox")).includes("NOTES.txt"));
  await fails("rename", { path: "Inbox/a.pdf", new_name: "Brief.docx" }, /already exists.*Nothing is overwritten/);
  await fails("rename", { path: "Inbox/a.pdf", new_name: "../a.pdf" }, /plain name/);
  await fails("rename", { path: "Inbox/a.pdf", new_name: "a?.pdf" }, /Invalid new name/);
  await fails("rename", { path: "Inbox/missing.pdf", new_name: "x.pdf" }, /not found/);
  await ok("rename", { path: "Inbox/deep", new_name: "Deeper" });
  assert.equal(fs.readFileSync(at("Inbox", "Deeper", "c.xlsx"), "utf8"), "C");
});

test("move: into a folder, to a new name, folders, and never overwriting", async () => {
  await ok("move", { from: "Inbox/b.docx", to: "Archive/" });
  assert.equal(fs.readFileSync(at("Archive", "b.docx"), "utf8"), "B");
  await ok("move", { from: at("Inbox", "notes.txt"), to: "Archive/2024 notes.txt" });
  assert.ok(exists(at("Archive", "2024 notes.txt")));
  await ok("move", { from: "Inbox/deep", to: "Archive/" });
  assert.ok(exists(at("Archive", "deep", "c.xlsx")));

  const t = await fails("move", { from: "Inbox/a.pdf", to: "Archive/" }, /already exists.*Nothing is overwritten/);
  assert.doesNotMatch(t, /end the destination/);
  assert.equal(fs.readFileSync(at("Archive", "a.pdf"), "utf8"), "old A");
  await fails("move", { from: "Inbox/a.pdf", to: "Archive" }, /already exists \(folder\).*end the destination with "\/"/);
  await fails("move", { from: "Inbox/a.pdf", to: "Inbox/a.pdf" }, /the same/);
  await fails("move", { from: "Archive", to: "Archive/deep/Archive" }, /into itself/);
  await fails("move", { from: "Inbox/a.pdf", to: "Archive/a.pdf/" }, /not a folder/);

  await fails("move", { from: "Inbox/a.pdf", to: "New/Sub/" }, /Destination folder does not exist.*New/);
  assert.ok(!exists(at("New")));
  const m = await ok("move", { from: "Inbox/a.pdf", to: "New/Sub/", create_folders: true });
  assert.match(m, /Created folders/);
  assert.equal(fs.readFileSync(at("New", "Sub", "a.pdf"), "utf8"), "A");
});

test("move: a symbolic link is moved as a link; its target is untouched", async () => {
  fs.symlinkSync(path.join(outside, "secret.txt"), at("Inbox", "link.txt"));
  await ok("move", { from: "Inbox/link.txt", to: "Archive/" });
  assert.equal(fs.readlinkSync(at("Archive", "link.txt")), path.join(outside, "secret.txt"));
  assert.equal(fs.readFileSync(path.join(outside, "secret.txt"), "utf8"), "S");
  // A dangling link at the destination counts as existing.
  fs.symlinkSync(path.join(root, "nowhere"), at("Archive", "ghost.pdf"));
  await fails("move", { from: "Inbox/a.pdf", to: "Archive/ghost.pdf" }, /already exists \(link\)/);
  assert.ok(exists(at("Inbox", "a.pdf")));
});

test("move_batch: ordered moves see earlier results", async () => {
  const t = await ok("move_batch", {
    create_folders: true,
    moves: [
      { from: "Inbox/deep", to: "Sheets" },
      { from: "Sheets/c.xlsx", to: "Finance/2024/Budget.xlsx" },
      { from: "Inbox/b.docx", to: "Finance/2024/" },
      { from: "Archive/a.pdf", to: "Archive/a (old).pdf" },
      { from: "Inbox/a.pdf", to: "Archive/" },
    ],
  });
  assert.match(t, /Moved 5 items/);
  assert.match(t, /Created folders:\n- .*Finance\n- .*Finance[\\/]2024/);
  assert.deepEqual(snapshot(allowed), {
    "Archive/": true,
    [path.join("Archive", "a (old).pdf")]: "old A",
    [path.join("Archive", "a.pdf")]: "A",
    "Finance/": true,
    [path.join("Finance", "2024") + "/"]: true,
    [path.join("Finance", "2024", "Budget.xlsx")]: "C",
    [path.join("Finance", "2024", "b.docx")]: "B",
    "Inbox/": true,
    [path.join("Inbox", "notes.txt")]: "N",
    "Sheets/": true,
  });
});

test("move_batch: a case-only rename and renaming a folder filled by earlier moves", async () => {
  const t = await ok("move_batch", {
    create_folders: true,
    moves: [
      { from: "Inbox/deep/c.xlsx", to: "Finance/2024/Budget.xlsx" },
      { from: "Inbox/b.docx", to: "Finance/2024/" },
      { from: "Inbox/notes.txt", to: "Inbox/NOTES.txt" },
      { from: "Finance", to: "Money" },
    ],
  });
  assert.match(t, /Moved 4 items/);
  assert.match(t, /Created folders:\n- .*Finance\n- .*Finance[\\/]2024/);
  assert.deepEqual(snapshot(allowed), {
    "Archive/": true,
    [path.join("Archive", "a.pdf")]: "old A",
    "Inbox/": true,
    [path.join("Inbox", "a.pdf")]: "A",
    [path.join("Inbox", "deep") + "/"]: true,
    [path.join("Inbox", "NOTES.txt")]: "N",
    "Money/": true,
    [path.join("Money", "2024") + "/"]: true,
    [path.join("Money", "2024", "Budget.xlsx")]: "C",
    [path.join("Money", "2024", "b.docx")]: "B",
  });
});

test("move_batch: a move that an earlier move made invalid fails, and the batch is undone", async () => {
  fs.symlinkSync(outside, at("escape"), "dir");
  const before = snapshot(allowed);
  const cases = [
    // The source was moved away by an earlier move.
    [[{ from: "Inbox/deep", to: "Archive/deep" }, { from: "Inbox/deep/c.xlsx", to: "Inbox/c.xlsx" }], /Move 2 of 2 failed .*not found/],
    // The destination was taken by an earlier move.
    [[{ from: "Inbox/a.pdf", to: "Inbox/c.pdf" }, { from: "Inbox/b.docx", to: "Inbox/c.pdf" }], /Move 2 of 2 failed .*already exists/],
    // A link moved inside is still resolved: moving through it would leave the allowed directory.
    [[{ from: "escape", to: "Archive/escape" }, { from: "Inbox/notes.txt", to: "Archive/escape/notes.txt" }], /Move 2 of 2 failed .*Access denied/],
  ];
  for (const [moves, re] of cases) {
    const t = await fails("move_batch", { moves }, re);
    assert.match(t, /The 1 earlier move\(s\) were undone; nothing was changed\./);
    assert.deepEqual(snapshot(allowed), before, JSON.stringify(moves));
  }
  assert.ok(exists(path.join(outside, "secret.txt")));
  assert.ok(!exists(path.join(outside, "notes.txt")));
});

test("move_batch: a failure undoes every earlier move and created folder", async () => {
  const before = snapshot(allowed);
  const t = await fails(
    "move_batch",
    {
      create_folders: true,
      moves: [
        { from: "Inbox/b.docx", to: "Work/Docs/" },
        { from: "Inbox/deep", to: "Work/Sheets" },
        { from: "Inbox/notes.txt", to: "Inbox/renamed.txt" },
        { from: "Inbox/a.pdf", to: "Archive/" }, // conflicts with Archive/a.pdf
        { from: "Inbox/renamed.txt", to: "Work/" },
      ],
    },
    /Move 4 of 5 failed .*already exists/
  );
  assert.match(t, /The 3 earlier move\(s\) were undone; nothing was changed\./);
  assert.deepEqual(snapshot(allowed), before);
});

test("move_batch: rejects malformed input before touching anything", async () => {
  const before = snapshot(allowed);
  await fails("move_batch", { moves: [] }, /./);
  await fails("move_batch", { moves: [{ from: "Inbox/a.pdf" }] }, /./);
  await fails("move_batch", { moves: [{ from: "Inbox/a.pdf", to: "x.pdf", overwrite: true }] }, /./);
  assert.deepEqual(snapshot(allowed), before);
});

test("moves between volumes (EXDEV) copy, then remove the original, and undo the same way", async () => {
  const serverDir = path.dirname(entry);
  const { setAllowedDirectories } = await import(path.join(serverDir, "paths.js"));
  const { moveEntries } = await import(path.join(serverDir, "files.js"));
  setAllowedDirectories([allowed]);
  const rename = fs.renameSync;
  // Only moves of the original items cross "volumes"; the final rename of the temp copy does not.
  fs.renameSync = (from, to) => {
    if (!path.basename(from).startsWith(".~move-")) throw Object.assign(new Error("cross-device link"), { code: "EXDEV" });
    return rename(from, to);
  };
  try {
    const before = snapshot(allowed);
    assert.throws(
      () => moveEntries([{ from: "Inbox/deep", to: "Archive/" }, { from: "Inbox/a.pdf", to: "Archive/" }]),
      /Move 2 of 2 failed[\s\S]*were undone/
    );
    assert.deepEqual(snapshot(allowed), before);
    moveEntries([{ from: "Inbox/deep", to: "Archive/" }]);
    assert.equal(fs.readFileSync(at("Archive", "deep", "c.xlsx"), "utf8"), "C");
    assert.ok(!exists(at("Inbox", "deep")));
    assert.deepEqual(fs.readdirSync(at("Archive")).filter((f) => f.startsWith(".~")), [], "no temp copies left");
  } finally {
    fs.renameSync = rename;
  }
});
