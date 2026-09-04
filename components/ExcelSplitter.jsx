"use client"

import React, { useState } from "react";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { splitExcelWorkbook } from "@/utils/excelSplitter";

const ExcelSplitter = () => {
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheets, setSelectedSheets] = useState([]);
  const [processing, setProcessing] = useState(false);

  const handleFile = async (selectedFile) => {
    if (!selectedFile) return;

    const extension = selectedFile.name
      .split(".")
      .pop()
      .toLowerCase();

    if (!["xlsx"].includes(extension)) {
      alert(
        "Please upload an .xlsx file."
      );
      return;
    }

    try {
      const buffer =
        await selectedFile.arrayBuffer();

      const zip = await JSZip.loadAsync(buffer);

      const workbookXml =
        await zip
          .file("xl/workbook.xml")
          .async("string");

      /*
       * Simple extraction of sheet names.
       * We use regex here only for the UI preview.
       */
      const matches = [
        ...workbookXml.matchAll(
          /<sheet\b[^>]*name="([^"]+)"/g
        ),
      ];

      const sheetNames = matches.map(
        (match) =>
          decodeXml(match[1])
      );

      setFile(selectedFile);
      setSheets(sheetNames);
      setSelectedSheets(sheetNames);
    } catch (error) {
      console.error(error);

      alert(
        "Unable to read this Excel file."
      );
    }
  };

  const decodeXml = (value) => {
    return value
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">");
  };

  const toggleSheet = (sheet) => {
    setSelectedSheets((current) =>
      current.includes(sheet)
        ? current.filter(
            (item) => item !== sheet
          )
        : [...current, sheet]
    );
  };

  const selectAll = () => {
    setSelectedSheets([...sheets]);
  };

  const clearAll = () => {
    setSelectedSheets([]);
  };

  const splitWorkbook = async () => {
    if (!file) return;

    if (selectedSheets.length === 0) {
      alert(
        "Please select at least one sheet."
      );
      return;
    }

    try {
      setProcessing(true);

      const files =
        await splitExcelWorkbook(
          file,
          selectedSheets
        );

      /*
       * Put all XLSX files into one ZIP.
       */
      const zip = new JSZip();

      files.forEach((item) => {
        zip.file(
          item.name,
          item.blob
        );
      });

      const output =
        await zip.generateAsync({
          type: "blob",
          compression: "DEFLATE",
        });

      const originalName =
        file.name.replace(
          /\.[^/.]+$/,
          ""
        );

      saveAs(
        output,
        `${originalName}-split.zip`
      );
    } catch (error) {
      console.error(error);

      alert(
        "Failed to split the workbook."
      );
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div className="excel-splitter">

      <h1>
        Excel Sheet Splitter
      </h1>

      {!file && (
        <label className="upload-area">

          <input
            type="file"
            accept=".xlsx"
            hidden
            onChange={(event) =>
              handleFile(
                event.target.files?.[0]
              )
            }
          />

          <div className="upload-icon">
            📊
          </div>

          <h3>
            Drop your Excel file here
          </h3>

          <p>
            or click to browse
          </p>

          <span>
            Supports .xlsx
          </span>

        </label>
      )}

      {file && (
        <div className="workspace">

          <div className="file-info">
            <strong>
              {file.name}
            </strong>

            <span>
              {sheets.length} sheets
            </span>
          </div>

          <div className="sheet-header">

            <h3>
              Select sheets
            </h3>

            <div>
              <button
                onClick={selectAll}
              >
                Select all
              </button>

              <button
                onClick={clearAll}
              >
                Clear
              </button>
            </div>

          </div>

          <div className="sheet-list">

            {sheets.map((sheet) => (
              <label
                key={sheet}
                className="sheet-item"
              >

                <input
                  type="checkbox"
                  checked={selectedSheets.includes(
                    sheet
                  )}
                  onChange={() =>
                    toggleSheet(sheet)
                  }
                />

                <span>
                  {sheet}
                </span>

              </label>
            ))}

          </div>

          <div className="footer">

            <span>
              {selectedSheets.length} of{" "}
              {sheets.length} selected
            </span>

            <button
              className="split-button"
              disabled={
                processing ||
                selectedSheets.length === 0
              }
              onClick={splitWorkbook}
            >
              {processing
                ? "Creating files..."
                : "Split & Download ZIP"}
            </button>

          </div>

        </div>
      )}

    </div>
  );
};

export default ExcelSplitter;

