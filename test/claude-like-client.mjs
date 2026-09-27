// Connects the way Claude Desktop does: SDK v2 client with versionNegotiation 'auto'
// on a *subclassed* stdio transport, so the server/discover probe runs in place on
// the same process (a mid-probe close is fatal, exactly as in Claude Desktop).
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

export class InPlaceStdioTransport extends StdioClientTransport {}

export async function connectLikeClaude(command, args, { stderr = "pipe", timeoutMs = 10000, env } = {}) {
  const transport = new InPlaceStdioTransport({ command, args, stderr, ...(env ? { env: { ...process.env, ...env } } : {}) });
  let stderrText = "";
  transport.stderr?.on("data", (d) => (stderrText += d));
  const client = new Client(
    { name: "claude-like", version: "1.0.0" },
    { versionNegotiation: { mode: "auto", probe: { timeoutMs } } }
  );
  await client.connect(transport);
  return { client, transport, stderr: () => stderrText };
}
