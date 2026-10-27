import http from "node:http";
import { File } from "node:buffer";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createChecker, processWorkbook } from "./src/weeklyCheck.js";

const root = dirname(fileURLToPath(import.meta.url));
const checker = createChecker(join(root, "data", "http-cache.json"));
await checker.load();

// The page posts the workbook here. Steps are written back as they happen.
const server = http.createServer(async (req, res) => {
  if (req.method !== "POST" || req.url !== "/api/run") {
    res.writeHead(404);
    res.end();
    return;
  }

  try {
    const chunks = [];
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    const file = new File([Buffer.concat(chunks)], "week.xlsx");
    console.log("workbook received", file.size, "bytes");
    res.writeHead(200, {
      "Content-Type": "application/x-ndjson; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
    });
    res.flushHeaders();

    const result = await processWorkbook(file, (message) => {
      res.write(JSON.stringify({ step: message }) + "\n");
    }, {
      checkUrl: (url) => checker.check(url),
    });

    await checker.save();
    res.write(JSON.stringify({ done: true, findings: result.findings }) + "\n");
    res.end();
  } catch (err) {
    if (!res.headersSent) {
      res.writeHead(500, { "Content-Type": "application/json" });
    }
    res.end(JSON.stringify({ error: err.message || "The workbook could not be processed." }));
  }
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    console.error("Already running. Open http://localhost:5173");
    process.exit(0);
  }
  throw err;
});

server.requestTimeout = 0;
server.headersTimeout = 0;
server.timeout = 0;
server.listen(8787, () => {
  console.log("API http://localhost:8787");
});
