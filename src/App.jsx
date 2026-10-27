import { useState } from "react";
import * as XLSX from "xlsx";

// The weekly check itself is in weeklyCheck.js. This only builds the download in the browser.
function findingsBlob(rows) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Findings");
  const bytes = XLSX.write(book, { bookType: "xlsx", type: "array" });
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function splitStep(text) {
  const match = String(text).match(/^(Step \d+[a-z]?):\s*(.*)$/);
  if (!match) {
    return { label: "", body: text };
  }
  return { label: match[1], body: match[2] };
}

function addStep(current, text) {
  const checking = text.includes("Checking page");
  const previous = current.at(-1);
  if (checking && previous && previous.text.includes("Checking page")) {
    return [...current.slice(0, -1), { text, state: "active" }];
  }
  const finished = current.map((step) => ({ ...step, state: "done" }));
  return [...finished, { text, state: "active" }];
}

export default function App() {
  const [steps, setSteps] = useState([]);
  const [findings, setFindings] = useState([]);
  const [blob, setBlob] = useState(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function onFile(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) {
      return;
    }

    // start a fresh run and show a step before the checker answers
    setError("");
    setFindings([]);
    setBlob(null);
    setSteps([{ text: "Reading the workbook in the browser", state: "active" }]);
    setBusy(true);

    try {
      // Chrome cannot upload a workbook that Excel still has open.
      // Read the bytes here so that failure is shown immediately.
      const bytes = await file.arrayBuffer();
      setSteps([{ text: "Uploading the workbook", state: "active" }]);

      // the JavaScript server reads the workbook and checks the pages
      const response = await fetch("/api/run", {
        method: "POST",
        headers: { "Content-Type": "application/octet-stream" },
        body: bytes,
      });
      if (!response.ok || !response.body) {
        throw new Error("The server could not read the workbook.");
      }
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let pending = "";
      let report = null;

      while (true) {
        const chunk = await reader.read();
        if (chunk.done) {
          break;
        }
        pending += decoder.decode(chunk.value, { stream: true });
        const lines = pending.split("\n");
        pending = lines.pop();
        lines.forEach((line) => {
          if (!line.trim()) {
            return;
          }
          const message = JSON.parse(line);
          if (message.error) {
            throw new Error(message.error);
          }
          if (message.done) {
            report = message.findings;
            return;
          }
          // replace the running page-check line instead of adding hundreds of rows
          setSteps((current) => addStep(current, message.step));
        });
      }

      setSteps((current) => current.map((step) => ({ ...step, state: "done" })));
      setFindings(report || []);
      setBlob(findingsBlob(report || []));
    } catch (err) {
      const locked = err && (err.name === "NotReadableError" || err.name === "NotFoundError");
      setError(locked ? "Close the workbook in Excel, then choose the file again." : (err.message || "The workbook could not be read."));
    } finally {
      setBusy(false);
    }
  }

  function download() {
    if (!blob) {
      return;
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "ca_404_findings.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  return (
    <main>
      <header className="hero">
        <p className="eyebrow">Weekly check · changiairport.com</p>
        <h1>CA 404 finder</h1>
        <p>Upload the weekly workbook. Each step says what the check is doing. A full file keeps running for a few minutes.</p>
      </header>

      <label className="upload">
        <span>Weekly workbook</span>
        <input type="file" accept=".xlsx" onChange={onFile} disabled={busy} />
      </label>

      {steps.length > 0 && (
        <ol className="steps">
          {steps.map((step) => {
            const parts = splitStep(step.text);
            return (
              <li key={step.text} className={step.state}>
                {parts.label && <span className="step-label">{parts.label}</span>}
                <span>{parts.body}</span>
              </li>
            );
          })}
        </ol>
      )}

      {error && <p className="error">{error}</p>}

      {findings.length > 0 && (
        <section>
          <div className="toolbar">
            <button type="button" onClick={download}>
              Download Excel
            </button>
          </div>

          <p>{findings.length} rows. Only broken links and unknown sources are listed.</p>

          <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>URL</th>
                <th>Views</th>
                <th>Status</th>
                <th>Action</th>
                <th>Evidence</th>
              </tr>
            </thead>
            <tbody>
              {findings.slice(0, 200).map((row) => (
                <tr key={row.item_id + row.url}>
                  <td>
                    {row.url ? (
                      <a href={row.url} target="_blank" rel="noreferrer">
                        {row.url}
                      </a>
                    ) : (
                      ""
                    )}
                  </td>
                  <td>{row.page_views}</td>
                  <td><span className={"badge status-" + row.status}>{row.status}</span></td>
                  <td><span className="badge action">{row.recommended_action}</span></td>
                  <td>{row.evidence}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
          {findings.length > 200 && <p>Showing the first 200 rows. The Excel file has all of them.</p>}
        </section>
      )}
    </main>
  );
}
