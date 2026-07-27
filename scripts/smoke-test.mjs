import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const pluginRoot = path.dirname(scriptDir);
const server = spawn(
  "powershell.exe",
  [
    "-NoLogo",
    "-NoProfile",
    "-NonInteractive",
    "-STA",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(scriptDir, "world-mcp.ps1"),
  ],
  { cwd: pluginRoot, stdio: ["pipe", "pipe", "pipe"] },
);

const requests = [
  { jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "world-smoke-test", version: "1.0.0" } } },
  { jsonrpc: "2.0", method: "notifications/initialized", params: {} },
  { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} },
  { jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "word_status", arguments: {} } },
  { jsonrpc: "2.0", id: 4, method: "tools/call", params: { name: "word_read", arguments: { scope: "selection", max_chars: 2000 } } },
  { jsonrpc: "2.0", id: 5, method: "tools/call", params: { name: "word_table", arguments: { action: "list" } } },
  { jsonrpc: "2.0", id: 6, method: "tools/call", params: { name: "word_comment", arguments: { action: "list" } } },
  { jsonrpc: "2.0", id: 7, method: "tools/call", params: { name: "word_review_changes", arguments: { action: "list", limit: 10 } } },
  { jsonrpc: "2.0", id: 8, method: "tools/call", params: { name: "word_track_changes", arguments: { action: "get" } } },
  { jsonrpc: "2.0", id: 9, method: "tools/call", params: { name: "word_com", arguments: { root: "document", operation: "get", path: "Name" } } },
  { jsonrpc: "2.0", id: 10, method: "tools/call", params: { name: "word_com", arguments: { root: "document", operation: "get", path: "PageSetup" } } },
  { jsonrpc: "2.0", id: 11, method: "tools/call", params: { name: "word_com", arguments: { operation: "release_all" } } },
];

let stdout = "";
let stderr = "";
server.stdout.setEncoding("utf8");
server.stderr.setEncoding("utf8");
server.stdout.on("data", (chunk) => { stdout += chunk; });
server.stderr.on("data", (chunk) => { stderr += chunk; });

for (const request of requests) {
  server.stdin.write(`${JSON.stringify(request)}\n`);
}
server.stdin.end();

const exitCode = await new Promise((resolve) => server.on("close", resolve));
if (exitCode !== 0) {
  throw new Error(`World MCP exited with ${exitCode}: ${stderr}`);
}

const responses = stdout
  .split(/\r?\n/)
  .filter(Boolean)
  .map((line) => JSON.parse(line));
const byId = new Map(responses.filter((item) => item.id !== undefined).map((item) => [item.id, item]));

for (const id of [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]) {
  if (!byId.has(id)) throw new Error(`Missing JSON-RPC response ${id}`);
  if (byId.get(id).error) throw new Error(`JSON-RPC ${id} failed: ${JSON.stringify(byId.get(id).error)}`);
  if (id >= 3 && byId.get(id).result.isError) {
    throw new Error(`Tool call ${id} failed: ${byId.get(id).result.content[0].text}`);
  }
}

const tools = byId.get(2).result.tools;
const statusCall = byId.get(3).result;
const readCall = byId.get(4).result;
if (tools.length < 10) throw new Error(`Expected broad Word coverage, received only ${tools.length} tools`);

const status = statusCall.structuredContent;
const selection = readCall.structuredContent;
process.stdout.write(`${JSON.stringify({
  ok: true,
  server: byId.get(1).result.serverInfo,
  toolCount: tools.length,
  toolNames: tools.map((tool) => tool.name),
  activeDocument: status.active_document,
  openDocuments: status.documents.map((document) => document.full_name),
  selection: { start: selection.start, end: selection.end, textLength: selection.text.length },
  genericComName: byId.get(9).result.structuredContent.result,
  genericComObject: byId.get(10).result.structuredContent.result.kind,
}, null, 2)}\n`);
