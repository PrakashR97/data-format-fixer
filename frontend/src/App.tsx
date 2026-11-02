import { useRef, useState } from "react";
import { discard, regenerate, uploadFile } from "./api";
import type { ColType, DateOrder, OutputFormat, Overrides, Result } from "./types";

const TYPE_CHOICES: ColType[] = ["Integer", "Decimal", "Date", "DateTime", "Text"];

type Status = "idle" | "processing" | "done" | "error";

export default function App() {
  const [status, setStatus] = useState<Status>("idle");
  const [error, setError] = useState("");
  const [result, setResult] = useState<Result | null>(null);
  const [dateOrder, setDateOrder] = useState<DateOrder>("auto");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("yyyy-mm-dd");
  const [overrides, setOverrides] = useState<Overrides>({});
  const [activeSheet, setActiveSheet] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [staged, setStaged] = useState<File | null>(null);
  const input = useRef<HTMLInputElement>(null);

  function finish(r: Result) {
    setResult(r);
    setStatus("done");
    setActiveSheet((i) => Math.min(i, r.sheets.length - 1));
  }

  function chooseFile(file: File | undefined) {
    if (!file) return;
    setStaged(file);
    setError("");
    if (status === "error") setStatus("idle");
  }

  async function handleUpload() {
    if (!staged) return;
    setStatus("processing");
    setError("");
    setOverrides({});
    setActiveSheet(0);
    try {
      finish(await uploadFile(staged, dateOrder, outputFormat));
      setStaged(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Upload failed.");
      setResult(null);
      setStatus("error");
    }
  }

  async function handleRegenerate() {
    if (!result) return;
    setStatus("processing");
    setError("");
    try {
      finish(await regenerate(result.job_id, dateOrder, outputFormat, overrides));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not regenerate the file.");
      setStatus("error");
    }
  }

  function reset() {
    if (result) discard(result.job_id);
    setResult(null);
    setStaged(null);
    setStatus("idle");
    setError("");
    setOverrides({});
  }

  function setOverride(sheet: string, col: number, value: ColType | "Auto") {
    setOverrides((prev) => ({ ...prev, [sheet]: { ...prev[sheet], [String(col)]: value } }));
  }

  const busy = status === "processing";
  const sheet = result?.sheets[activeSheet];

  return (
    <div className="page">
      <header className="top">
        <div className="brand">
          <span className="dot" /> DATA FORMAT FIXER
        </div>
        <div className="tag">XLSX / CSV → CORRECTED XLSX</div>
      </header>

      <main>
        {!result && (<>
        <h1>Fix numbers and dates stored as text.</h1>
        <p className="lead">
          Upload a spreadsheet export. Values are checked with plain rules (no AI, no key). Only
          confident conversions are made; everything else is kept as is and listed below.
        </p>
        </>)}

        <section className="panel settings">
          <label>
            <span className="label">INPUT DATE ORDER</span>
            <select value={dateOrder} onChange={(e) => setDateOrder(e.target.value as DateOrder)} disabled={busy}>
              <option value="auto">Auto (never guess)</option>
              <option value="dmy">Day-first (31/12/2026)</option>
              <option value="mdy">Month-first (12/31/2026)</option>
            </select>
          </label>
          <label>
            <span className="label">OUTPUT DATE FORMAT</span>
            <select value={outputFormat} onChange={(e) => setOutputFormat(e.target.value as OutputFormat)} disabled={busy}>
              <option value="yyyy-mm-dd">yyyy-mm-dd</option>
              <option value="mm/dd/yyyy">mm/dd/yyyy</option>
              <option value="dd/mm/yyyy">dd/mm/yyyy</option>
            </select>
          </label>
        </section>

        {!result && (
        <div
          className={`drop ${dragging ? "over" : ""} ${busy ? "busy" : ""}`}
          onClick={() => !busy && input.current?.click()}
          onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && !busy && input.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            if (!busy) chooseFile(e.dataTransfer.files[0]);
          }}
          role="button"
          tabIndex={0}
          aria-label="Upload a file"
        >
          <input
            ref={input}
            type="file"
            accept=".xlsx,.csv"
            hidden
            onChange={(e) => {
              chooseFile(e.target.files?.[0]);
              e.target.value = "";
            }}
          />
          <div className="drop-title">{busy ? "Processing…" : staged ? staged.name : "Drop a .xlsx or .csv file here"}</div>
          <div className="muted">{busy ? "Analysing columns and converting values" : staged ? "Click to choose a different file" : "or click to choose a file"}</div>
          {busy && <div className="bar"><i /></div>}
        </div>
        )}
        {!result && (
          <div className="regen">
            <button className="btn primary" onClick={handleUpload} disabled={!staged || busy}>
              Upload
            </button>
            <span className="muted small">{staged ? "Ready. Press Upload to analyse it." : "Choose a file first."}</span>
          </div>
        )}

        <div className="status" aria-live="polite">
          <span className={`pip ${status}`} />
          {status === "idle" && "Waiting for a file"}
          {status === "processing" && "Processing"}
          {status === "done" && result && `Done · ${result.converted_total.toLocaleString()} cells converted · ${result.warning_total.toLocaleString()} kept with warnings`}
          {status === "error" && "Something went wrong"}
        </div>

        {error && <div className="error" role="alert">{error}</div>}

        {result && sheet && (
          <>
            <section className="panel actions">
              <div>
                <div className="label">OUTPUT</div>
                <div className="file">{result.download_name}</div>
              </div>
              <div className="buttons">
                <a className="btn primary" href={result.download_url} download={result.download_name}>
                  Download
                </a>
                <button className="btn" onClick={reset} disabled={busy}>Start over</button>
              </div>
            </section>

            {result.csv && (
              <div className="note">
                <b>CSV file</b> · {result.csv.encoding}, {result.csv.delimiter}-separated. {result.csv.note}
              </div>
            )}
            {!!result.ambiguous_total && (
              <div className="note warn">
                <b>{result.ambiguous_total} ambiguous date{result.ambiguous_total === 1 ? "" : "s"}</b> kept
                as text (could be day-first or month-first). Choose a date order above and press
                Apply changes &amp; regenerate.
              </div>
            )}
            {result.notices.map((n) => (
              <div className="note warn" key={n}>{n}</div>
            ))}

            <section>
              <div className="tabs" role="tablist">
                {result.sheets.map((s, i) => (
                  <button key={s.name + i} role="tab" aria-selected={i === activeSheet}
                    className={i === activeSheet ? "tab on" : "tab"} onClick={() => setActiveSheet(i)}>
                    {s.name}
                  </button>
                ))}
              </div>
              {sheet.columns.length === 0 ? (
                <p className="muted pad">This sheet is empty.</p>
              ) : (
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>COL</th><th>HEADER</th><th>DETECTED</th><th>USE TYPE</th><th>CONVERTED</th><th>WARNINGS</th><th>SAMPLE</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sheet.columns.map((c) => {
                        const chosen = overrides[sheet.name]?.[String(c.index)] ?? (c.overridden ? c.effective_type : "Auto");
                        return (
                          <tr key={c.index}>
                            <td className="mono">{c.letter}</td>
                            <td>{c.header ?? <span className="muted">(none)</span>}</td>
                            <td>
                              <span className={`chip ${c.detected_type === "Mixed" ? "red" : ""}`}>{c.detected_type}</span>
                              {c.notes.map((n) => <div className="muted small" key={n}>{n}</div>)}
                            </td>
                            <td>
                              <select value={chosen} disabled={busy} aria-label={`Type for column ${c.letter}`}
                                onChange={(e) => setOverride(sheet.name, c.index, e.target.value as ColType | "Auto")}>
                                <option value="Auto">Auto</option>
                                {TYPE_CHOICES.map((t) => <option key={t} value={t}>{t}</option>)}
                              </select>
                            </td>
                            <td className="mono">{c.converted}</td>
                            <td className={`mono ${c.warning_count ? "redtext" : ""}`}>{c.warning_count}</td>
                            <td className="muted small">{c.samples.join(" · ")}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
              <div className="regen">
                <button className="btn primary" onClick={handleRegenerate} disabled={busy}>
                  Apply changes &amp; regenerate
                </button>
                <span className="muted small">Uses the date settings above and any column types you changed.</span>
              </div>
            </section>

            <section>
              <h2>Warnings <span className="count">{result.warning_total}</span></h2>
              {result.warnings.length === 0 ? (
                <p className="muted">No problems found. Nothing was left unconverted.</p>
              ) : (
                <ul className="warnings">
                  {result.warnings.map((w, i) => (
                    <li key={i}>
                      <div>
                        <b>{w.sheet}</b> · column {w.column}{w.header ? ` (${w.header})` : ""} · <span className="redtext">{w.count}×</span>{" "}
                        {w.ambiguous && <span className="chip red">AMBIGUOUS</span>} {w.reason}
                      </div>
                      <div className="locs mono small">
                        {w.locations.map((l) => `${l.cell}: "${l.value}"`).join("   ")}
                        {w.count > w.locations.length && `   …and ${w.count - w.locations.length} more`}
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </section>
          </>
        )}
      </main>
      <footer>Files are deleted automatically and your original is never changed. Contact: Prakash · <a href="mailto:prak.cs97@gmail.com">prak.cs97@gmail.com</a></footer>
    </div>
  );
}
