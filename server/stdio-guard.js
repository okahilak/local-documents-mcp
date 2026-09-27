// MCP uses stdout for JSON-RPC. Some libraries (pdf.js) log warnings with
// console.log, which would corrupt the stream, so route all console output to stderr.
// This module must be imported before anything else.
for (const method of ["log", "info", "warn", "debug", "trace"]) {
  console[method] = (...args) => console.error(...args);
}
