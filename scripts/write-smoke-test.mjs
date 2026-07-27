import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const pluginRoot = path.dirname(scriptDir);

async function run(requests) {
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
  let stdout = "";
  let stderr = "";
  server.stdout.setEncoding("utf8");
  server.stderr.setEncoding("utf8");
  server.stdout.on("data", (chunk) => { stdout += chunk; });
  server.stderr.on("data", (chunk) => { stderr += chunk; });
  server.stdin.write(`${JSON.stringify({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "world-write-smoke-test", version: "1.0.0" } },
  })}\n`);
  server.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized", params: {} })}\n`);
  for (const request of requests) server.stdin.write(`${JSON.stringify(request)}\n`);
  server.stdin.end();
  const exitCode = await new Promise((resolve) => server.on("close", resolve));
  if (exitCode !== 0) throw new Error(`World MCP exited with ${exitCode}: ${stderr}`);
  const messages = stdout.split(/\r?\n/).filter(Boolean).map((line) => JSON.parse(line));
  const byId = new Map(messages.filter((message) => message.id !== undefined).map((message) => [message.id, message]));
  for (const request of requests) {
    const response = byId.get(request.id);
    if (!response) throw new Error(`Missing response ${request.id}`);
    if (response.error) throw new Error(`RPC ${request.id} failed: ${JSON.stringify(response.error)}`);
    if (response.result?.isError) throw new Error(`Tool ${request.params.name} failed: ${response.result.content[0].text}`);
  }
  return byId;
}

const initial = await run([
  { jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: "word_status", arguments: {} } },
  { jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "word_read", arguments: { scope: "document", max_chars: 200000 } } },
]);
const status = initial.get(2).result.structuredContent;
const before = initial.get(3).result.structuredContent;
const document = status.active_document;
if (!document) throw new Error("No active Word document");
if (before.truncated) throw new Error("Document is too large for a reversible smoke test");

const marker = `[[WORLD-MCP-SMOKE-A-${Date.now()}]]`;
const replacement = marker.replace("-A-", "-B-");
const insertAt = Math.max(0, before.end - 1);
let changed;
try {
  changed = await run([
    {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/call",
      params: { name: "word_edit", arguments: { document, scope: "range", start: insertAt, end: insertAt, mode: "replace", text: marker } },
    },
    {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: { name: "word_read", arguments: { document, scope: "range", start: insertAt, end: insertAt + marker.length, max_chars: marker.length + 10 } },
    },
    {
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: { name: "word_find_replace", arguments: { document, find_text: marker, replace_text: replacement, replace: "one" } },
    },
    {
      jsonrpc: "2.0",
      id: 5,
      method: "tools/call",
      params: { name: "word_read", arguments: { document, scope: "range", start: insertAt, end: insertAt + replacement.length, max_chars: replacement.length + 10 } },
    },
    {
      jsonrpc: "2.0",
      id: 6,
      method: "tools/call",
      params: { name: "word_undo_redo", arguments: { document, action: "undo", steps: 2 } },
    },
    {
      jsonrpc: "2.0",
      id: 7,
      method: "tools/call",
      params: { name: "word_read", arguments: { document, scope: "document", max_chars: 200000 } },
    },
  ]);
} catch (error) {
  await run([
    {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/call",
      params: { name: "word_find_replace", arguments: { document, find_text: marker, replace_text: "", replace: "all" } },
    },
    {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: { name: "word_find_replace", arguments: { document, find_text: replacement, replace_text: "", replace: "all" } },
    },
  ]).catch(() => {});
  throw error;
}

const inserted = changed.get(3).result.structuredContent.text;
const replaced = changed.get(5).result.structuredContent.text;
const findReplace = changed.get(4).result.structuredContent;
const undo = changed.get(6).result.structuredContent;
const after = changed.get(7).result.structuredContent;
if (inserted !== marker) throw new Error("Inserted marker could not be read back exactly");
if (!findReplace.performed || replaced !== replacement) throw new Error("Word-native find/replace could not be verified");
if (after.text !== before.text) {
  await run([
    {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/call",
      params: { name: "word_find_replace", arguments: { document, find_text: marker, replace_text: "", replace: "all" } },
    },
    {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: { name: "word_find_replace", arguments: { document, find_text: replacement, replace_text: "", replace: "all" } },
    },
  ]);
  throw new Error("Undo did not restore the original document text; the marker was removed by recovery");
}

process.stdout.write(`${JSON.stringify({
  ok: true,
  document,
  insertedAndReadBack: true,
  findReplaceVerified: true,
  undoCompleted: undo.completed,
  originalTextRestored: true,
  savedInPlace: undo.save.saved,
  residualMarker: false,
}, null, 2)}\n`);
