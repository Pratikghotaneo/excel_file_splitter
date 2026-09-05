"use client";

import { ChangeEvent, DragEvent, useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);

  const [dragging, setDragging] = useState(false);

  const [processing, setProcessing] = useState(false);

  const [error, setError] = useState("");

  const [success, setSuccess] = useState("");

  function selectFile(selected?: File) {
    setError("");
    setSuccess("");

    if (!selected) {
      return;
    }

    const name = selected.name.toLowerCase();

    if (!name.endsWith(".xlsx") && !name.endsWith(".xlsm")) {
      setError("Please select an .xlsx or .xlsm file.");

      return;
    }

    setFile(selected);
  }

  function handleInput(event: ChangeEvent<HTMLInputElement>) {
    selectFile(event.target.files?.[0]);
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();

    setDragging(false);

    selectFile(event.dataTransfer.files?.[0]);
  }

  async function splitFile() {
    if (!file) {
      setError("Please select an Excel file.");

      return;
    }

    try {
      setProcessing(true);
      setError("");
      setSuccess("");

      const formData = new FormData();

      formData.append("file", file);

      const response = await fetch("/api/split", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const data = await response.json().catch(() => null);

        throw new Error(data?.error ?? "Unable to split the workbook.");
      }

      const blob = await response.blob();

      const url = URL.createObjectURL(blob);

      const link = document.createElement("a");

      link.href = url;
      link.download = "split-excel-files.zip";

      document.body.appendChild(link);

      link.click();

      link.remove();

      URL.revokeObjectURL(url);

      setSuccess(
        "Done! Your individual Excel files have been downloaded as a ZIP.",
      );
    } catch (error) {
      setError(
        error instanceof Error ? error.message : "Something went wrong.",
      );
    } finally {
      setProcessing(false);
    }
  }

  return (
    <main className="page">
      <div className="container">
        <header className="header">
          <div className="logo">X</div>

          <div>
            <h1>Excel Sheet Splitter</h1>

            <p>
              Split a workbook into individual Excel files while preserving the
              original Excel package structure.
            </p>
          </div>
        </header>

        <section
          className={`dropzone ${dragging ? "dragging" : ""} ${
            file ? "hasFile" : ""
          }`}
          onDragOver={(event) => {
            event.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
        >
          {!file ? (
            <>
              <div className="uploadIcon">↑</div>

              <h2>Drop your Excel file</h2>

              <p>XLSX and XLSM are supported</p>

              <label className="selectButton">
                Select Excel File
                <input
                  type="file"
                  accept=".xlsx,.xlsm"
                  hidden
                  onChange={handleInput}
                />
              </label>
            </>
          ) : (
            <div className="selectedFile">
              <div className="excelIcon">X</div>

              <div className="fileInfo">
                <strong>{file.name}</strong>

                <span>{formatBytes(file.size)}</span>
              </div>

              <button
                type="button"
                className="removeButton"
                onClick={() => setFile(null)}
              >
                Remove
              </button>
            </div>
          )}
        </section>

        {error && <div className="message error">{error}</div>}

        {success && <div className="message success">{success}</div>}

        <button
          className="splitButton"
          disabled={!file || processing}
          onClick={splitFile}
        >
          {processing ? (
            <>
              <span className="spinner" />
              Splitting workbook...
            </>
          ) : (
            <>
              Split Excel File
              <span>→</span>
            </>
          )}
        </button>

        <section className="features">
          <Feature
            title="Original XML"
            text="The workbook is copied at the Open XML package level instead of rebuilding cells."
          />

          <Feature
            title="Rich Excel content"
            text="Worksheet relationships such as drawings, images, tables and comments are retained."
          />

          <Feature
            title="XLSM support"
            text="The original VBA project binary is retained when processing macro-enabled workbooks."
          />
        </section>

        <p className="note">
          The generated files contain one worksheet each. The original worksheet
          XML and related package parts are retained rather than recreating the
          workbook through a spreadsheet library.
        </p>
        <div className="w-full flex justify-center items-center mt-5 bg-gray-800 p-10">
           <p className="text-xl text-white">Developed by Pratik S. Ghotane</p>
        </div>
      </div>
    </main>
  );
}

function Feature({ title, text }: { title: string; text: string }) {
  return (
    <div className="feature">
      <div className="check">✓</div>

      <div>
        <strong>{title}</strong>

        <p>{text}</p>
      </div>
    </div>
  );
}

function formatBytes(bytes: number) {
  if (bytes === 0) {
    return "0 Bytes";
  }

  const units = ["Bytes", "KB", "MB", "GB"];

  const index = Math.floor(Math.log(bytes) / Math.log(1024));

  return `${(bytes / Math.pow(1024, index)).toFixed(2)} ${units[index]}`;
}
