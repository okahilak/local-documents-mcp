// Protocol version negotiation tests (MCP 2026-07-28 `server/discover` probe vs the
// 2025-era `initialize` handshake), at the raw JSON-RPC level and through real clients.
// SERVER_ENTRY env var can point at an unpacked bundle's server/index.js.
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { Client as ClientV2 } from "@modelcontextprotocol/client";
import { StdioClientTransport as StdioV2 } from "@modelcontextprotocol/client/stdio";
import { Client as ClientV1 } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport as StdioV1 } from "@modelcontextprotocol/sdk/client/stdio.js";
import { connectLikeClaude } from "./claude-like-client.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const allowed = path.join(here, "fixtures", "allowed");
const entry = path.resolve(process.env.SERVER_ENTRY || path.join(here, "..", "server", "index.js"));

// The exact probe the SDK 2.x client (and Claude Desktop) sends first.
const PROBE = {
  jsonrpc: "2.0",
  id: "server-discover-probe-1",
  method: "server/discover",
  params: {
    _meta: {
      "io.modelcontextprotocol/protocolVersion": "2026-07-28",
      "io.modelcontextprotocol/clientInfo": { name: "probe-test", version: "1.0.0" },
      "io.modelcontextprotocol/clientCapabilities": {},
    },
  },
};
const INITIALIZE = (id) => ({
  jsonrpc: "2.0",
  id,
  method: "initialize",
  params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "raw", version: "1" } },
});

/** Minimal raw JSON-RPC-over-stdio driver. */
function rawServer(serverEntry = entry, args = [allowed]) {
  const child = spawn(process.execPath, [serverEntry, ...args], { stdio: ["pipe", "pipe", "pipe"] });
  let buf = "";
  let stderr = "";
  const waiting = new Map();
  const received = [];
  let exited = null;
  child.on("exit", (code, signal) => (exited = { code, signal }));
  child.stderr.on("data", (d) => (stderr += d));
  child.stdout.on("data", (d) => {
    buf += d;
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      const msg = JSON.parse(line); // any non-JSON on stdout fails the test
      received.push(msg);
      if (msg.id !== undefined && waiting.has(msg.id)) {
        waiting.get(msg.id)(msg);
        waiting.delete(msg.id);
      }
    }
  });
  return {
    child,
    received,
    stderr: () => stderr,
    exited: () => exited,
    send: (m) => child.stdin.write(JSON.stringify(m) + "\n"),
    request(m, timeoutMs = 10000) {
      return new Promise((resolve, reject) => {
        const t = setTimeout(() => reject(new Error(`timeout waiting for ${m.method}; stderr:\n${stderr}`)), timeoutMs);
        waiting.set(m.id, (msg) => (clearTimeout(t), resolve(msg)));
        child.stdin.write(JSON.stringify(m) + "\n");
      });
    },
    async close() {
      child.stdin.end();
      if (exited) return;
      await new Promise((r) => {
        const t = setTimeout(() => (child.kill(), r()), 3000);
        child.on("exit", () => (clearTimeout(t), r()));
      });
    },
  };
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

test("raw: server/discover before initialize gets a modern discover result and the process stays up", async () => {
  const s = rawServer();
  try {
    const res = await s.request(PROBE);
    assert.equal(res.error, undefined, JSON.stringify(res.error));
    assert.ok(Array.isArray(res.result.supportedVersions), JSON.stringify(res.result));
    assert.ok(res.result.supportedVersions.includes("2026-07-28"));
    assert.equal(res.result._meta?.["io.modelcontextprotocol/serverInfo"]?.name, "local-documents");
    assert.ok(res.result.capabilities?.tools, "advertises tools capability");
    await sleep(300);
    assert.equal(s.exited(), null, "server must not exit after the probe");

    // Modern-era follow-up on the same connection: tools/list with the per-request envelope.
    const list = await s.request({ jsonrpc: "2.0", id: 2, method: "tools/list", params: { _meta: PROBE.params._meta } });
    assert.equal(list.error, undefined, JSON.stringify(list.error));
    assert.equal(list.result.tools.length, 15);
    const call = await s.request({
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: { _meta: PROBE.params._meta, name: "xlsx_search", arguments: { path: "book.xlsx", query: "needle" } },
    });
    assert.equal(call.error, undefined, JSON.stringify(call.error));
    assert.match(call.result.content[0].text, /Hidden Notes!B3/);
  } finally {
    await s.close();
  }
});

test("raw: envelope-less server/discover is answered with an error (not a close), then legacy initialize works", async () => {
  const s = rawServer();
  try {
    const res = await s.request({ jsonrpc: "2.0", id: 1, method: "server/discover", params: {} });
    assert.ok(res.error, "expected a JSON-RPC error reply");
    assert.equal(res.error.code, -32601);
    await sleep(300);
    assert.equal(s.exited(), null, "server must not exit after an unrecognised probe");

    const init = await s.request(INITIALIZE(2));
    assert.equal(init.error, undefined, JSON.stringify(init.error));
    assert.equal(init.result.serverInfo.name, "local-documents");
    s.send({ jsonrpc: "2.0", method: "notifications/initialized" });
    const list = await s.request({ jsonrpc: "2.0", id: 3, method: "tools/list" });
    assert.equal(list.result.tools.length, 15);
  } finally {
    await s.close();
  }
});

test("raw: plain legacy initialize handshake (no probe) still works", async () => {
  const s = rawServer();
  try {
    const init = await s.request(INITIALIZE(1));
    assert.equal(init.result.protocolVersion, "2025-06-18");
    s.send({ jsonrpc: "2.0", method: "notifications/initialized" });
    const call = await s.request({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: "pdf_info", arguments: { path: "report.pdf" } } });
    assert.equal(JSON.parse(call.result.content[0].text).pages, 3);
  } finally {
    await s.close();
  }
});

test("raw: stdout carries only JSON-RPC, and garbage input does not kill the server", async () => {
  const s = rawServer();
  try {
    s.child.stdin.write("this is not json\n");
    s.child.stdin.write("{}\n");
    await sleep(200);
    assert.equal(s.exited(), null);
    const res = await s.request(PROBE);
    assert.ok(res.result?.supportedVersions);
    // Trigger PDF engine load (pdf.js is chatty) and make sure nothing non-JSON hit stdout.
    await s.request({ jsonrpc: "2.0", id: 9, method: "tools/call", params: { _meta: PROBE.params._meta, name: "pdf_read_text", arguments: { path: "report.pdf" } } });
  } finally {
    await s.close();
  }
});

test("raw: probe is answered even with no usable allowed directories", async () => {
  const s = rawServer(entry, ["${user_config.allowed_directories}"]);
  try {
    const res = await s.request(PROBE);
    assert.ok(res.result?.supportedVersions);
  } finally {
    await s.close();
  }
});

test("client: Claude-like in-place probe negotiates the modern protocol", async () => {
  const { client } = await connectLikeClaude(process.execPath, [entry, allowed]);
  try {
    assert.equal(client.getNegotiatedProtocolVersion(), "2026-07-28");
    const { tools } = await client.listTools();
    assert.equal(tools.length, 15);
    const r = await client.callTool({ name: "pdf_info", arguments: { path: "report.pdf" } });
    assert.equal(JSON.parse(r.content[0].text).pages, 3);
  } finally {
    await client.close();
  }
});

test("client: SDK 2.x base transport (sibling probe) and legacy mode both connect", async () => {
  for (const mode of ["auto", "legacy"]) {
    const client = new ClientV2({ name: "t", version: "1" }, { versionNegotiation: { mode } });
    await client.connect(new StdioV2({ command: process.execPath, args: [entry, allowed], stderr: "ignore" }));
    try {
      const v = client.getNegotiatedProtocolVersion();
      if (mode === "auto") assert.equal(v, "2026-07-28");
      else assert.match(v, /^2025-/);
      assert.equal((await client.listTools()).tools.length, 15);
    } finally {
      await client.close();
    }
  }
});

test("client: SDK 1.x legacy client connects", async () => {
  const client = new ClientV1({ name: "t", version: "1" });
  await client.connect(new StdioV1({ command: process.execPath, args: [entry, allowed], stderr: "ignore" }));
  try {
    assert.match(client.getServerVersion().name, /local-documents/);
    assert.equal((await client.listTools()).tools.length, 15);
  } finally {
    await client.close();
  }
});

test("resilience: if the PDF engine cannot load, the probe still succeeds and XLSX/DOCX tools keep working", async () => {
  // Build a copy of the server whose node_modules lacks pdfjs-dist.
  const serverDir = path.dirname(entry);
  const pkgRoot = path.dirname(serverDir);
  const nmSrc = fs.existsSync(path.join(pkgRoot, "node_modules")) ? path.join(pkgRoot, "node_modules") : null;
  assert.ok(nmSrc, "node_modules next to the server is required for this test");
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ld-broken-"));
  try {
    fs.cpSync(serverDir, path.join(tmp, "server"), { recursive: true });
    fs.writeFileSync(path.join(tmp, "package.json"), JSON.stringify({ type: "module" }));
    const nm = path.join(tmp, "node_modules");
    fs.mkdirSync(nm);
    for (const name of fs.readdirSync(nmSrc)) {
      if (name === "pdfjs-dist" || name === ".bin") continue;
      fs.symlinkSync(path.join(nmSrc, name), path.join(nm, name), "junction");
    }
    const s = rawServer(path.join(tmp, "server", "index.js"), [allowed]);
    try {
      const res = await s.request(PROBE);
      assert.ok(res.result?.supportedVersions, "probe must succeed even though the PDF engine is broken");
      await sleep(500); // background warm-up has failed by now
      assert.equal(s.exited(), null, "server must survive an engine load failure");
      const meta = { _meta: PROBE.params._meta };
      const pdf = await s.request({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { ...meta, name: "pdf_info", arguments: { path: "report.pdf" } } });
      assert.equal(pdf.result.isError, true);
      assert.match(pdf.result.content[0].text, /PDF engine failed to load/);
      const x = await s.request({ jsonrpc: "2.0", id: 3, method: "tools/call", params: { ...meta, name: "xlsx_info", arguments: { path: "book.xlsx" } } });
      assert.equal(x.result.isError, undefined, x.result.content[0].text);
      assert.equal(JSON.parse(x.result.content[0].text).sheet_count, 2);
      const d = await s.request({ jsonrpc: "2.0", id: 4, method: "tools/call", params: { ...meta, name: "docx_info", arguments: { path: "memo.docx" } } });
      assert.equal(d.result.isError, undefined, d.result.content[0].text);
      assert.equal(JSON.parse(d.result.content[0].text).title, "Project Memo");
    } finally {
      await s.close();
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
});
