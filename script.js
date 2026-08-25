"use strict";

/* =========================================================
   0. CONFIGURATION
   Update these two values to point at your real backend.
   ========================================================= */

// Replace with your real conversion endpoint, e.g. "https://api.yourapp.com/convert"
const API_ENDPOINT = "http://127.0.0.1:8000/convert";

// 25 MB — adjust to whatever your backend can comfortably handle.
const MAX_FILE_SIZE = 25 * 1024 * 1024;

const DEFAULT_ADDITIONAL_ROWS = 10;

/* =========================================================
   1. APPLICATION STATE
   A single source of truth for what the interface is doing.
   Exactly one on-screen state ever reflects appState.current.
   ========================================================= */

const appState = {
  current: "upload", // upload | selected | converting | success | error
  file: null,        // the File chosen by the user (kept even though it isn't displayed)
  additionalRows: DEFAULT_ADDITIONAL_ROWS,
  result: null,       // { blob, filename, url } once a conversion succeeds
};

/* =========================================================
   2. ELEMENT REFERENCES
   ========================================================= */

const stateSections = Array.from(document.querySelectorAll(".state"));

const dropzone = document.getElementById("dropzone");
const chooseBtn = document.getElementById("chooseBtn");
const fileInput = document.getElementById("fileInput");
const uploadError = document.getElementById("uploadError");

const additionalRowsInput = document.getElementById("additionalRows");
const convertBtn = document.getElementById("convertBtn");

const successDetail = document.getElementById("successDetail");
const downloadBtn = document.getElementById("downloadBtn");
const convertAnotherBtn = document.getElementById("convertAnotherBtn");

const errorDetail = document.getElementById("errorDetail");
const retryBtn = document.getElementById("retryBtn");

let dragCounter = 0; // tracks nested dragenter/dragleave events

/* =========================================================
   3. UI STATE MANAGEMENT
   The only place that decides what's visible. Every state section
   starts as display:none (see style.css `.state`); adding
   "is-active" is the sole mechanism that reveals one of them.
   ========================================================= */

function showState(state) {
  appState.current = state;

  stateSections.forEach((section) => {
    section.classList.toggle("is-active", section.dataset.state === state);
  });
}

function resetApp() {
  appState.file = null;
  appState.additionalRows = DEFAULT_ADDITIONAL_ROWS;

  if (appState.result && appState.result.url) {
    URL.revokeObjectURL(appState.result.url);
  }
  appState.result = null;

  fileInput.value = "";
  additionalRowsInput.value = String(DEFAULT_ADDITIONAL_ROWS);
  hideUploadError();

  showState("upload");
}

/* =========================================================
   4. FILE VALIDATION
   ========================================================= */

function showUploadError(message) {
  uploadError.textContent = message;
  uploadError.hidden = false;
}

function hideUploadError() {
  uploadError.hidden = true;
  uploadError.textContent = "";
}

/**
 * Confirms a File is a PDF within the configured size limit.
 * @param {File} file
 * @returns {{ valid: boolean, message?: string }}
 */
function validateFile(file) {
  if (!file) {
    return { valid: false, message: "Please choose a PDF file." };
  }

  const looksLikePdfType = file.type === "application/pdf";
  const looksLikePdfName = file.name.toLowerCase().endsWith(".pdf");

  if (!looksLikePdfType && !looksLikePdfName) {
    return { valid: false, message: "Please choose a PDF file." };
  }

  if (file.size > MAX_FILE_SIZE) {
    return {
      valid: false,
      message: "This PDF is too large to process. Please choose a smaller file.",
    };
  }

  return { valid: true };
}

/* =========================================================
   5. FILE SELECTION
   ========================================================= */

function handleFile(file) {
  const result = validateFile(file);

  if (!result.valid) {
    showUploadError(result.message);
    return;
  }

  hideUploadError();
  appState.file = file;

  showState("selected");
}

/* =========================================================
   6. DRAG & DROP + FILE PICKER WIRING
   ========================================================= */

function openFilePicker() {
  fileInput.click();
}

dropzone.addEventListener("click", (event) => {
  // The "Choose PDF" button handles its own click; avoid opening the
  // dialog twice when the click bubbles up from the button.
  if (event.target === chooseBtn) return;
  openFilePicker();
});

dropzone.addEventListener("keydown", (event) => {
  if (event.target !== dropzone) return;
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    openFilePicker();
  }
});

chooseBtn.addEventListener("click", (event) => {
  event.stopPropagation();
  openFilePicker();
});

fileInput.addEventListener("change", () => {
  const file = fileInput.files && fileInput.files[0];
  if (file) handleFile(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (eventName === "dragenter") dragCounter++;
    dropzone.classList.add("is-dragover");
  });
});

dropzone.addEventListener("dragleave", (event) => {
  event.preventDefault();
  event.stopPropagation();
  dragCounter = Math.max(0, dragCounter - 1);
  if (dragCounter === 0) dropzone.classList.remove("is-dragover");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  event.stopPropagation();
  dragCounter = 0;
  dropzone.classList.remove("is-dragover");

  const file = event.dataTransfer.files && event.dataTransfer.files[0];
  if (file) handleFile(file);
});

/* =========================================================
   7. CONVERSION OPTIONS
   The only remaining option is "additional page rows".
   ========================================================= */

additionalRowsInput.addEventListener("blur", () => {
  const min = Number(additionalRowsInput.min) || 0;
  const max = Number(additionalRowsInput.max) || 200;
  let value = parseInt(additionalRowsInput.value, 10);

  if (Number.isNaN(value)) value = min;
  value = Math.min(Math.max(value, min), max);

  additionalRowsInput.value = String(value);
  appState.additionalRows = value;
});

additionalRowsInput.addEventListener("input", () => {
  const parsed = parseInt(additionalRowsInput.value, 10);
  appState.additionalRows = Number.isNaN(parsed) ? DEFAULT_ADDITIONAL_ROWS : parsed;
});

/* =========================================================
   8. API COMMUNICATION
   ========================================================= */

/**
 * Sends the PDF and the additional-rows setting to the backend and
 * returns the generated Word document as a blob.
 * @param {File} file
 * @param {number} additionalRows
 * @returns {Promise<{ blob: Blob, filename: string }>}
 */
async function convertPDF(file, additionalRows) {
  const formData = new FormData();
  formData.append("pdf", file);
  formData.append("additional_rows", String(additionalRows));

  // Note: do not set a Content-Type header manually — the browser sets the
  // correct multipart boundary automatically when body is a FormData.
  const response = await fetch(API_ENDPOINT, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`Conversion failed (${response.status}): ${errorText}`);
  }

  const blob = await response.blob();

  const disposition = response.headers.get("Content-Disposition") || "";
  const filenameMatch = disposition.match(/filename="?([^"]+)"?/);
  const filename = filenameMatch
    ? filenameMatch[1]
    : file.name.replace(/\.pdf$/i, "") + ".docx";

  return { blob, filename };
}

/* =========================================================
   9. CONVERSION FLOW
   ========================================================= */

convertBtn.addEventListener("click", async () => {
  if (!appState.file) return;

  showState("converting");

  try {
    const { blob, filename } = await convertPDF(appState.file, appState.additionalRows);

    appState.result = {
      blob,
      filename,
      url: URL.createObjectURL(blob),
    };

    successDetail.textContent = filename;
    showState("success");
  } catch (err) {
    // Technical detail stays in the console; the UI stays calm and vague.
    console.error("Doc Converter: conversion request failed.", err);
    errorDetail.textContent = "We couldn't convert this document. Please try again.";
    showState("error");
  }
});

/* =========================================================
   10. DOWNLOAD HANDLING
   ========================================================= */

downloadBtn.addEventListener("click", () => {
  if (!appState.result) return;

  const link = document.createElement("a");
  link.href = appState.result.url;
  link.download = appState.result.filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});

convertAnotherBtn.addEventListener("click", () => {
  resetApp();
});

retryBtn.addEventListener("click", () => {
  // Return to the selection/options state — the file is still held
  // in appState.file, so the user doesn't need to re-upload.
  if (appState.file) {
    showState("selected");
  } else {
    showState("upload");
  }
});

/* =========================================================
   11. INITIAL STATE
   ========================================================= */

showState("upload");
