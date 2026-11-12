import { useState } from "react";
import axios from "axios";
import API_BASE from "./api";

function App() {
  const [dataFile, setDataFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [failedRows, setFailedRows] = useState([]);

  const uploadFiles = async () => {
    if (!dataFile || !templateFile) {
      alert("Please select both Excel files");
      return;
    }

    const formData = new FormData();
    formData.append("data_file", dataFile);
    formData.append("template_file", templateFile);

    try {
      setLoading(true);
      setMessage("Generating PODs... Please wait ⏳");
      setFailedRows([]);

      const res = await axios.post(
        `${API_BASE}/generate-pods`,
        formData,
        { headers: { "Content-Type": "multipart/form-data" } }
      );

      const data = res.data;

      if (data.status === "SUCCESS") {
        setMessage(
          `✅ ${data.generated_excels} PODs Generated\nSaved in: ${data.folder}` +
            (data.failed_count
              ? `\n⚠️ ${data.failed_count} rows failed`
              : "")
        );
      } else {
        setMessage(`❌ ${data.message || "Generation failed"}`);
      }

      if (data.failed_rows && data.failed_rows.length) {
        setFailedRows(data.failed_rows);
      }
    } catch (err) {
      console.error(err);
      const errMsg =
        err.response?.data?.message ||
        err.response?.data?.detail ||
        err.message ||
        "Unknown error";
      setMessage(`❌ Error generating PODs\n${errMsg}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h2>POD Excel Generator</h2>

        <label style={styles.label}>Upload Data Excel</label>
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setDataFile(e.target.files[0])}
        />

        <label style={styles.label}>Upload POD Template</label>
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setTemplateFile(e.target.files[0])}
        />

        <button onClick={uploadFiles} disabled={loading} style={styles.button}>
          {loading ? "Processing..." : "Generate PODs"}
        </button>

        {message && <pre style={styles.message}>{message}</pre>}

        {failedRows.length > 0 && (
          <div style={styles.failedBox}>
            <strong>Failed rows:</strong>
            <ul style={styles.failedList}>
              {failedRows.map((r, i) => (
                <li key={i}>
                  POD <b>{r.pod}</b>: {r.error}
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
}

const styles = {
  container: {
    height: "100vh",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    background: "#f4f6f8",
    fontFamily: "Arial",
  },
  card: {
    background: "white",
    padding: "30px",
    borderRadius: "12px",
    width: "400px",
    boxShadow: "0 5px 20px rgba(0,0,0,0.15)",
    textAlign: "center",
  },
  label: {
    display: "block",
    marginTop: "15px",
    marginBottom: "5px",
    fontWeight: "bold",
  },
  button: {
    marginTop: "20px",
    padding: "12px",
    width: "100%",
    background: "#007bff",
    color: "white",
    border: "none",
    borderRadius: "8px",
    fontSize: "16px",
    cursor: "pointer",
  },
  message: {
    marginTop: "20px",
    background: "#eef2f7",
    padding: "10px",
    borderRadius: "8px",
    whiteSpace: "pre-wrap",
    textAlign: "left",
    fontSize: "13px",
  },
  failedBox: {
    marginTop: "15px",
    background: "#ffe8e8",
    padding: "10px",
    borderRadius: "8px",
    textAlign: "left",
    fontSize: "12px",
  },
  failedList: {
    margin: "6px 0 0 0",
    paddingLeft: "18px",
  },
};

export default App;