"use strict";

/**
 * Optional native Office conversion bridge.
 *
 * The pure-JavaScript converters in processing.js are good enough for simple
 * documents, but they cannot reproduce the full layout engine of Word, Excel
 * and PowerPoint. When the machine running PDF Machan has LibreOffice or the
 * Microsoft Office desktop apps installed, we hand the document to that
 * application instead and use its own PDF exporter. This is the same strategy
 * used by Stirling-PDF / Gotenberg for Office documents.
 *
 * The bridge is best-effort: if the native application is missing, takes too
 * long, or fails, the caller falls back to the JavaScript converter. Temporary
 * files are created in the OS temp folder and always removed afterwards.
 */

const fs = require("fs");
const fsp = fs.promises;
const os = require("os");
const path = require("path");
const { execFile } = require("child_process");

const IS_WINDOWS = process.platform === "win32";

const OFFICE_DIR =
  process.env.MSOFFICE_DIR ||
  path.join(process.env.ProgramFiles || "C:\\Program Files", "Microsoft Office", "root", "Office16");

const OFFICE_EXE = {
  word: path.join(OFFICE_DIR, "WINWORD.EXE"),
  excel: path.join(OFFICE_DIR, "EXCEL.EXE"),
  powerpoint: path.join(OFFICE_DIR, "POWERPNT.EXE"),
};

const LIBREOFFICE_CANDIDATES = [
  process.env.SOFFICE_PATH,
  IS_WINDOWS ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe" : null,
  IS_WINDOWS ? "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe" : null,
  "/usr/bin/soffice",
  "/usr/local/bin/soffice",
  "/opt/libreoffice/program/soffice",
].filter(Boolean);

const POWERSHELL = IS_WINDOWS
  ? path.join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
  : null;

const disabled = new Set();

function isDisabled(engine, kind) {
  return disabled.has(`${engine}:${kind}`);
}

function disable(engine, kind) {
  disabled.add(`${engine}:${kind}`);
}

function exists(file) {
  try {
    return Boolean(file) && fs.existsSync(file);
  } catch (error) {
    return false;
  }
}

function librewriter() {
  return LIBREOFFICE_CANDIDATES.find(exists) || null;
}

function engineFor(kind) {
  const configured = String(process.env.OFFICE_ENGINE || "auto").toLowerCase();

  // Microsoft Word's PDF reflow is useful but can sit behind a hidden dialog in
  // some Office installs. Keep it opt-in so a normal conversion never waits on
  // it; the pure-JS text extractor is the default.
  if (kind === "pdf2word" && process.env.OFFICE_PDF2WORD !== "1") return null;

  if (configured === "none") return null;
  if (configured === "libreoffice") {
    return !isDisabled("libreoffice", kind) && librewriter() ? "libreoffice" : null;
  }
  if (configured === "msoffice") {
    const app = officeAppFor(kind);
    return IS_WINDOWS && !isDisabled("msoffice", kind) && app && exists(OFFICE_EXE[app]) && exists(POWERSHELL)
      ? "msoffice"
      : null;
  }

  const app = officeAppFor(kind);
  if (IS_WINDOWS && !isDisabled("msoffice", kind) && app && exists(OFFICE_EXE[app]) && exists(POWERSHELL)) {
    return "msoffice";
  }
  if (!isDisabled("libreoffice", kind) && librewriter()) return "libreoffice";
  return null;
}

function officeAppFor(kind) {
  if (kind === "word2pdf" || kind === "pdf2word") return "word";
  if (kind === "xls2pdf") return "excel";
  if (kind === "ppt2pdf") return "powerpoint";
  return null;
}

function extensionFor(kind, originalName) {
  const ext = (path.extname(originalName || "") || "").toLowerCase();
  if (kind === "word2pdf") return ext === ".doc" ? ".doc" : ".docx";
  if (kind === "xls2pdf") return [".xls", ".xlsm", ".xlsb"].includes(ext) ? ext : ".xlsx";
  if (kind === "ppt2pdf") return ext === ".ppt" ? ".ppt" : ".pptx";
  if (kind === "pdf2word") return ".pdf";
  return ".bin";
}

function outputExtension(kind) {
  return kind === "pdf2word" ? ".docx" : ".pdf";
}

function runProcess(file, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      args,
      {
        timeout: options.timeout || 120000,
        windowsHide: true,
        maxBuffer: 16 * 1024 * 1024,
        env: options.env || process.env,
      },
      (error, stdout, stderr) => {
        if (error) {
          error.stdout = String(stdout || "");
          error.stderr = String(stderr || "");
          reject(error);
          return;
        }
        resolve({ stdout: String(stdout || ""), stderr: String(stderr || "") });
      }
    );
  });
}

function powershellScript(kind) {
  if (kind === "word2pdf") {
    return `
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0
try {
  $doc = $app.Documents.Open($env:PM_INPUT, $false, $true, $false)
  try { $doc.ExportAsFixedFormat($env:PM_OUTPUT, 17) } finally { $doc.Close(0) }
} finally { $app.Quit() }
`.trim();
  }

  if (kind === "xls2pdf") {
    return `
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$app.ScreenUpdating = $false
try {
  $wb = $app.Workbooks.Open($env:PM_INPUT, 0, $true)
  try {
    foreach ($ws in $wb.Worksheets) {
      $ws.PageSetup.Zoom = $false
      $ws.PageSetup.FitToPagesWide = 1
      $ws.PageSetup.FitToPagesTall = $false
    }
    $wb.ExportAsFixedFormat(0, $env:PM_OUTPUT)
  } finally { $wb.Close($false) }
} finally { $app.Quit() }
`.trim();
  }

  if (kind === "ppt2pdf") {
    return `
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($env:PM_INPUT, $true, $false, $false)
  try { $pres.SaveAs($env:PM_OUTPUT, 32) } finally { $pres.Close() }
} finally { $app.Quit() }
`.trim();
  }

  if (kind === "pdf2word") {
    return `
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0
try {
  $doc = $app.Documents.Open($env:PM_INPUT, $false, $true, $false)
  try { $doc.SaveAs2($env:PM_OUTPUT, 16) } finally { $doc.Close(0) }
} finally { $app.Quit() }
`.trim();
  }

  return "";
}

async function convertWithPowerShell(kind, inputPath, outputPath) {
  const script = powershellScript(kind);
  if (!script) throw new Error(`No PowerShell converter for ${kind}`);
  const encoded = Buffer.from(script, "utf16le").toString("base64");
  await runProcess(
    POWERSHELL,
    ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
    {
      timeout: kind === "pdf2word" ? 45000 : 150000,
      env: { ...process.env, PM_INPUT: inputPath, PM_OUTPUT: outputPath },
    }
  );
}

async function convertWithLibreOffice(kind, inputPath, outputDir) {
  const soffice = librewriter();
  if (!soffice) throw new Error("LibreOffice is not installed");

  let filter = "pdf";
  if (kind === "pdf2word") filter = 'docx:MS Word 2007 XML';
  const args = ["--headless", "--norestore", "--nolockcheck", "--nodefault", "--convert-to", filter, "--outdir", outputDir];
  if (kind === "pdf2word") args.unshift('--infilter=writer_pdf_import');
  args.push(inputPath);
  await runProcess(soffice, args, { timeout: 150000 });

  const base = path.basename(inputPath, path.extname(inputPath));
  const outputPath = path.join(outputDir, `${base}.${kind === "pdf2word" ? "docx" : "pdf"}`);
  const buffer = await fsp.readFile(outputPath);
  return buffer;
}

function looksValid(kind, buffer) {
  if (!buffer || buffer.length < 8) return false;
  if (kind === "pdf2word") {
    return buffer[0] === 0x50 && buffer[1] === 0x4b; // PK zip / docx
  }
  return buffer.subarray(0, 5).toString("latin1") === "%PDF-";
}

/**
 * Convert a buffer with the native Office application when one is available.
 * Resolves to a Buffer, or null when no native converter is usable. Never
 * throws: the caller is expected to use the JavaScript fallback instead.
 */
async function convertOfficeDocument(kind, buffer, originalName) {
  const engine = engineFor(kind);
  if (!engine) return null;

  let tempDir;
  try {
    tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "pdfmachan-office-"));
    const inputPath = path.join(tempDir, `input${extensionFor(kind, originalName)}`);
    const outputPath = path.join(tempDir, `output${outputExtension(kind)}`);
    await fsp.writeFile(inputPath, buffer);

    if (engine === "msoffice") {
      await convertWithPowerShell(kind, inputPath, outputPath);
    } else {
      return validate(kind, await convertWithLibreOffice(kind, inputPath, tempDir), engine);
    }

    const output = await fsp.readFile(outputPath);
    return validate(kind, output, engine);
  } catch (error) {
    // Do not retry a broken engine for every request in this process.
    disable(engine, kind);
    if (process.env.DEBUG_OFFICE === "1") {
      console.warn(`[office] ${engine} ${kind} failed, using JS fallback:`, error.message);
    }
    return null;
  } finally {
    if (tempDir) {
      try {
        await fsp.rm(tempDir, { recursive: true, force: true });
      } catch (error) {
        /* best effort cleanup */
      }
    }
  }
}

function validate(kind, buffer, engine) {
  if (!looksValid(kind, buffer)) {
    disable(engine, kind);
    if (process.env.DEBUG_OFFICE === "1") {
      console.warn(`[office] ${engine} returned an invalid ${kind} output`);
    }
    return null;
  }
  return buffer;
}

module.exports = { convertOfficeDocument, engineFor };
