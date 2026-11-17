import { useRef, useState } from "react";
import { ArrowDownToLine, Check, FileSpreadsheet, FileUp, LoaderCircle, RefreshCw, ShieldCheck, Sparkles, X } from "lucide-react";
import API_BASE from "./api";

function FileCard({ title, description, file, onFile, inputRef, kind }) {
  const [dragging, setDragging] = useState(false);
  const acceptFile = (candidate) => {
    if (!candidate) return;
    if (!candidate.name.toLowerCase().endsWith(".xlsx")) {
      onFile(null, "Please choose an .xlsx workbook.");
      return;
    }
    onFile(candidate, "");
  };
  return (
    <section className={`upload-card ${dragging ? "is-dragging" : ""} ${file ? "has-file" : ""}`}>
      <div className="upload-card-top">
        <span className={`upload-icon ${kind}`}><FileSpreadsheet size={20} strokeWidth={1.8} /></span>
        <span className="step-number">{kind === "data" ? "01" : "02"}</span>
      </div>
      <div className="upload-copy">
        <h3>{title}</h3>
        <p>{description}</p>
      </div>
      <input ref={inputRef} type="file" accept=".xlsx" hidden onChange={(event) => acceptFile(event.target.files?.[0])} />
      <button
        type="button"
        className="drop-zone"
        onClick={() => inputRef.current?.click()}
        onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(event) => { event.preventDefault(); setDragging(false); acceptFile(event.dataTransfer.files?.[0]); }}
        aria-label={`Choose ${title.toLowerCase()}`}
      >
        {file ? <>
          <span className="file-check"><Check size={16} /></span>
          <span className="file-meta"><strong>{file.name}</strong><small>{(file.size / 1024).toFixed(file.size < 1024 * 1024 ? 0 : 1)} {file.size < 1024 * 1024 ? "KB" : "MB"} · Ready to use</small></span>
          <span className="change-file">Change</span>
        </> : <>
          <span className="upload-action-icon"><FileUp size={18} /></span>
          <span className="drop-label"><strong>Choose a file</strong> or drag it here</span>
          <span className="file-format">XLSX</span>
        </>}
      </button>
    </section>
  );
}

function App() {
  const [dataFile, setDataFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const dataRef = useRef(null);
  const templateRef = useRef(null);
  const ready = Boolean(dataFile && templateFile);

  const updateFile = (setter) => (file, message) => {
    setter(file);
    setError(message);
    setResult(null);
  };

  const generate = async () => {
    if (!ready || loading) return;
    setLoading(true);
    setError("");
    setResult(null);
    const body = new FormData();
    body.append("data_file", dataFile);
    body.append("template_file", templateFile);
    try {
      const response = await fetch(`${API_BASE}/generate-pods`, { method: "POST", body });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.detail || payload.message || "Generation failed. Please check your files and try again.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "generated-pods.zip";
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setResult({ generated: Number(response.headers.get("X-POD-Generated") || 0), failed: Number(response.headers.get("X-POD-Failed") || 0) });
    } catch (exception) {
      setError(exception.message || "Could not connect to the POD generator. Make sure the backend is running.");
    } finally {
      setLoading(false);
    }
  };

  const reset = () => {
    setDataFile(null);
    setTemplateFile(null);
    setError("");
    setResult(null);
    if (dataRef.current) dataRef.current.value = "";
    if (templateRef.current) templateRef.current.value = "";
  };

  return (
    <main className="app-shell">
      <div className="ambient ambient-one" />
      <div className="ambient ambient-two" />
      <header className="topbar">
        <a className="brand" href="#top" aria-label="POD Studio home">
          <span className="brand-mark"><FileSpreadsheet size={19} /></span>
          <span>POD<span className="brand-light">studio</span></span>
        </a>
        <span className="topbar-note"><span className="live-dot" /> Your files stay on your device</span>
      </header>

      <div className="main-content" id="top">
        <section className="intro">
          <div className="eyebrow"><Sparkles size={14} /> DELIVERY DOCUMENTS, SIMPLIFIED</div>
          <h1>Proof of delivery,<br /><span>without the busywork.</span></h1>
          <p className="intro-text">Turn your delivery spreadsheet into ready-to-send PODs in a few clicks. Upload your data, add your template, and we’ll handle the rest.</p>
        </section>

        <section className="generator-panel" aria-label="Generate proof of delivery documents">
          <div className="panel-heading">
            <div><h2>Build your PODs</h2><p>Two files in. Your documents out.</p></div>
            <span className="secure-note"><ShieldCheck size={15} /> Secure processing</span>
          </div>
          <div className="upload-grid">
            <FileCard title="Delivery data" description="The Excel sheet with your shipment details." file={dataFile} onFile={updateFile(setDataFile)} inputRef={dataRef} kind="data" />
            <FileCard title="POD template" description="Your existing Excel layout for each delivery." file={templateFile} onFile={updateFile(setTemplateFile)} inputRef={templateRef} kind="template" />
          </div>
          {error && <div className="notice notice-error" role="alert"><span className="notice-symbol"><X size={14} /></span><span>{error}</span></div>}
          {result && <div className={`notice ${result.failed ? "notice-warning" : "notice-success"}`} role="status"><span className="notice-symbol"><Check size={14} /></span><span><strong>{result.generated} POD{result.generated === 1 ? "" : "s"} ready.</strong> Your ZIP download has started.{result.failed ? ` ${result.failed} row${result.failed === 1 ? "" : "s"} could not be generated; see generation-report.json in the ZIP.` : ""}</span></div>}
          <div className="action-row">
            <div className="readiness">{ready ? <><span className="ready-check"><Check size={12} /></span> Both workbooks are ready</> : "Add both workbooks to continue"}</div>
            <button className="generate-button" type="button" onClick={generate} disabled={!ready || loading}>
              {loading ? <><LoaderCircle className="spin" size={17} /> Creating your PODs…</> : <><span>Generate PODs</span><ArrowDownToLine size={17} /></>}
            </button>
          </div>
          {result && <button className="reset-button" onClick={reset} type="button"><RefreshCw size={13} /> Start another batch</button>}
        </section>

        <div className="process-note"><span>01</span> Upload your spreadsheet <i /> <span>02</span> Add your template <i /> <span>03</span> Download your PODs</div>
        <footer className="footer"><span>Made for the work behind every delivery.</span><span>Excel files · Private by design</span></footer>
      </div>
    </main>
  );
}

export default App;
