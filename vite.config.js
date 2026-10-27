import { File } from "node:buffer";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { createChecker, processWorkbook } from "./src/weeklyCheck.js";

const root = dirname(fileURLToPath(import.meta.url));
const checker = createChecker(join(root, "data", "http-cache.json"));
const ready = checker.load();

// The React page and this checker are one JavaScript app.
// A browser visit is a GET. The upload from the page is a POST.
function runApi() {
  return {
    name: "run-api",
    configureServer(server) {
      server.httpServer?.on("checkContinue", (req, res) => {
        res.writeContinue();
        server.httpServer.emit("request", req, res);
      });

      server.middlewares.use("/api/run", async (req, res) => {
        if (req.method === "GET") {
          res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({
            ok: true,
            message: "The checker is running. Open http://localhost:5173 and upload the workbook there.",
          }));
          return;
        }
        if (req.method !== "POST") {
          res.writeHead(405, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({ error: "Upload the workbook from http://localhost:5173" }));
          return;
        }

        try {
          await ready;
          const chunks = [];
          for await (const chunk of req) {
            chunks.push(chunk);
          }
          const file = new File([Buffer.concat(chunks)], "week.xlsx");
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
          res.end(JSON.stringify({ done: true, findings: result.findings }) + "\n");
        } catch (err) {
          const body = JSON.stringify({ error: err.message || "The workbook could not be processed." });
          if (!res.headersSent) {
            res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
          }
          res.end(body);
        }
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), runApi()],
  server: {
    host: true,
    port: 5173,
    strictPort: true,
  },
});
