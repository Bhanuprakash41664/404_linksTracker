import { spawn } from "node:child_process";

// One JavaScript process serves the React page and answers /api/run.
const page = spawn(process.execPath, ["node_modules/vite/bin/vite.js"], { stdio: "inherit" });

function stop() {
  page.kill();
}

page.on("exit", (code) => {
  process.exit(code ?? 0);
});

process.on("SIGINT", stop);
process.on("SIGTERM", stop);
