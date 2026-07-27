import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
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
  { cwd: path.dirname(scriptDir), stdio: ["pipe", "pipe", "pipe"] },
);

let nextId = 1;
const pending = new Map();
let stdoutBuffer = "";
let stderr = "";

server.stdout.setEncoding("utf8");
server.stderr.setEncoding("utf8");
server.stderr.on("data", (chunk) => { stderr += chunk; });
server.stdout.on("data", (chunk) => {
  stdoutBuffer += chunk;
  const lines = stdoutBuffer.split(/\r?\n/);
  stdoutBuffer = lines.pop() ?? "";
  for (const line of lines) {
    if (!line) continue;
    const message = JSON.parse(line);
    const waiter = pending.get(message.id);
    if (!waiter) continue;
    pending.delete(message.id);
    if (message.error) waiter.reject(new Error(JSON.stringify(message.error)));
    else if (message.result?.isError) waiter.reject(new Error(message.result.content?.[0]?.text ?? "Tool failed"));
    else waiter.resolve(message.result);
  }
});

function request(method, params = {}) {
  const id = nextId++;
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    server.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
  });
}

async function tool(name, args) {
  return request("tools/call", { name, arguments: args });
}

let scratchDocument;
try {
  await request("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "world-context-format-smoke-test", version: "1.0.0" },
  });
  server.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized", params: {} })}\n`);

  scratchDocument = (await tool("word_document", { action: "new" })).structuredContent.full_name;
  const sourceInsert = (await tool("word_table", {
    action: "insert",
    scope: "range",
    start: 0,
    end: 0,
    rows: 2,
    columns: 3,
    text: "A\tB\tC\n1\t2\t3",
    match_context: false,
    save_after: false,
  })).structuredContent;
  if (sourceInsert.cells_written !== 6) throw new Error(`Expected 6 populated source cells, got ${sourceInsert.cells_written}`);

  const tableObjectsResult = (await tool("word_com", {
    root: "document",
    operation: "get",
    path: "Tables",
    save_after: false,
  })).structuredContent.result;
  const tableObjects = Array.isArray(tableObjectsResult) ? tableObjectsResult : [tableObjectsResult];
  const sourceTableHandle = tableObjects[0].handle;
  for (let row = 1; row <= 2; row++) {
    for (let column = 1; column <= 3; column++) {
      const cellHandle = (await tool("word_com", {
        root: "handle",
        handle: sourceTableHandle,
        operation: "call",
        path: "Cell",
        arguments: [row, column],
        save_after: false,
      })).structuredContent.result.handle;
      for (const [pathName, value] of row === 1
        ? [
            ["Shading.BackgroundPatternColor", 7884319],
            ["Range.Font.Color", 16777215],
            ["Range.Font.Bold", -1],
          ]
        : [
            ["Shading.BackgroundPatternColor", 16315114],
            ["Range.Font.Color", -16777216],
            ["Range.Font.Bold", 0],
          ]) {
        await tool("word_com", {
          root: "handle",
          handle: cellHandle,
          operation: "set",
          path: pathName,
          value,
          save_after: false,
        });
      }
    }
  }
  const sourceHeaderCell = (await tool("word_com", {
    root: "handle",
    handle: sourceTableHandle,
    operation: "call",
    path: "Cell",
    arguments: [1, 1],
    save_after: false,
  })).structuredContent.result.handle;
  const sourceHeaderColor = (await tool("word_com", {
    root: "handle",
    handle: sourceHeaderCell,
    operation: "get",
    path: "Shading.BackgroundPatternColor",
    save_after: false,
  })).structuredContent.result;
  if (sourceHeaderColor !== 7884319) throw new Error(`Source header setup failed: ${sourceHeaderColor}`);

  await tool("word_edit", {
    scope: "document",
    mode: "insert_after",
    text: "\r\r",
    save_after: false,
  });
  const documentRead = (await tool("word_read", {
    scope: "document",
    max_chars: 1000,
  })).structuredContent;
  const insertAt = Math.max(0, documentRead.end - 1);
  const targetInsert = (await tool("word_table", {
    action: "insert",
    scope: "range",
    start: insertAt,
    end: insertAt,
    rows: 3,
    columns: 2,
    text: "H1\tH2\n\t\n\t",
    template_table_index: 1,
    save_after: false,
  })).structuredContent;
  if (!targetInsert.context_matched) throw new Error("Target table did not match context");
  if (targetInsert.template_table_index !== 1) throw new Error("Wrong template table was used");
  if (targetInsert.cells_written !== 6) throw new Error(`Expected 6 populated target cells, got ${targetInsert.cells_written}`);
  if (targetInsert.format.header_color !== 7884319) throw new Error(`Header color mismatch: ${JSON.stringify(targetInsert.format)}`);
  if (targetInsert.format.body_color !== 16315114) throw new Error(`Body color mismatch: ${targetInsert.format.body_color}`);

  process.stdout.write(`${JSON.stringify({
    ok: true,
    populatedCells: targetInsert.cells_written,
    contextMatched: targetInsert.context_matched,
    templateTableIndex: targetInsert.template_table_index,
    headerColor: targetInsert.format.header_color,
    bodyColor: targetInsert.format.body_color,
  }, null, 2)}\n`);
} finally {
  if (scratchDocument) {
    await tool("word_document", {
      action: "close",
      save_changes: false,
      save_after: false,
    }).catch(() => {});
  }
  await tool("word_com", { operation: "release_all", save_after: false }).catch(() => {});
  server.stdin.end();
  const exitCode = await new Promise((resolve) => server.on("close", resolve));
  if (exitCode !== 0) throw new Error(`World MCP exited with ${exitCode}: ${stderr}`);
}
