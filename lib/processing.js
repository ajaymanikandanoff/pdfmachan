"use strict";

/**
 * PDF Machan / Image Machan — server side processing engine.
 *
 * Every handler receives `{ files, meta }` and returns
 * `{ buffer, filename, mime }`.  Files are kept in memory only: nothing is
 * ever written to disk and nothing is retained after the response is sent.
 */

const path = require("path");
// @cantoo/pdf-lib is an API-compatible fork of pdf-lib that also implements
// password encryption (needed by the "Protect PDF" tool). The low-level
// stream classes are used by the PDF compressor to rewrite embedded images.
const {
  PDFDocument,
  rgb,
  degrees,
  StandardFonts,
  PDFName,
  PDFNumber,
  PDFDict,
  PDFArray,
  PDFRawStream,
  decodePDFRawStream,
} = require("@cantoo/pdf-lib");
const JSZip = require("jszip");
const sharp = require("sharp");
const mammoth = require("mammoth");
const XLSX = require("xlsx");
const PDFKit = require("pdfkit");
const { Document, Packer, Paragraph, TextRun, PageBreak } = require("docx");
const PptxGenJS = require("pptxgenjs");
const napi = require("@napi-rs/canvas");
const { createCanvas } = napi;
const { convertOfficeDocument } = require("./office");

// pdf.js expects a couple of DOM primitives while rendering in Node. These
// have to be installed *before* the pdf.js module is required.
globalThis.DOMMatrix = globalThis.DOMMatrix || napi.DOMMatrix;
globalThis.Path2D = globalThis.Path2D || napi.Path2D;
globalThis.ImageData = globalThis.ImageData || napi.ImageData;

const pdfjs = require("pdfjs-dist/legacy/build/pdf.js");

const PDFJS_ROOT = path.dirname(require.resolve("pdfjs-dist/legacy/build/pdf.js"));
const STANDARD_FONTS_DIR = path.join(PDFJS_ROOT, "..", "..", "standard_fonts") + path.sep;
const CMAPS_DIR = path.join(PDFJS_ROOT, "..", "..", "cmaps") + path.sep;

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

/* ------------------------------------------------------------------ */
/* generic helpers                                                     */
/* ------------------------------------------------------------------ */

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function baseName(filename, fallback = "document") {
  const raw = path.basename(String(filename || fallback));
  const noExt = raw.replace(/\.[^./\\]+$/, "");
  const clean = (noExt || fallback).replace(/[^\w\-.() ]+/g, "_").slice(0, 80);
  return clean || fallback;
}

function getOne(files, label = "file") {
  if (!files || !files.length) throw new HttpError(400, `Please select a ${label}.`);
  return files[0];
}

function parseJsonField(value, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  if (typeof value === "object") return value;
  try {
    return JSON.parse(value);
  } catch (error) {
    throw new HttpError(400, "The request data was not valid JSON.");
  }
}

function pdfResult(doc, filename) {
  return doc.save().then((bytes) => ({
    buffer: Buffer.from(bytes),
    filename,
    mime: "application/pdf",
  }));
}

/* ------------------------------------------------------------------ */
/* pdf.js helpers                                                      */
/* ------------------------------------------------------------------ */

class NapiCanvasFactory {
  create(width, height) {
    const canvas = createCanvas(width, height);
    return { canvas, context: canvas.getContext("2d") };
  }

  reset(target, width, height) {
    target.canvas.width = width;
    target.canvas.height = height;
    target.context = target.canvas.getContext("2d");
  }

  destroy(target) {
    target.canvas.width = 0;
    target.canvas.height = 0;
    target.canvas = null;
    target.context = null;
  }
}

async function openPdf(buffer) {
  try {
    return await pdfjs.getDocument({
      data: new Uint8Array(buffer),
      canvasFactory: new NapiCanvasFactory(),
      standardFontDataUrl: STANDARD_FONTS_DIR,
      cMapUrl: CMAPS_DIR,
      cMapPacked: true,
      useWorkerFetch: false,
      isEvalSupported: false,
      disableFontFace: true,
      verbosity: 0,
    }).promise;
  } catch (error) {
    throw new HttpError(400, "That PDF could not be opened. It may be damaged or password protected.");
  }
}

async function renderPage(page, scale = 2) {
  const viewport = page.getViewport({ scale });
  const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({
    canvasContext: context,
    viewport,
    background: "rgb(255,255,255)",
  }).promise;
  return { canvas, viewport };
}

/**
 * Group the text items of a page into visual lines (top to bottom).
 */
async function extractPageLines(page) {
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const items = content.items
    .filter((item) => item.str && item.str.trim())
    .map((item) => {
      const transform = pdfjs.Util.transform(viewport.transform, item.transform);
      return {
        x: transform[4],
        y: transform[5],
        str: item.str,
        width: item.width || 0,
      };
    });

  const rows = [];
  for (const item of items) {
    let row = rows.find((candidate) => Math.abs(candidate.y - item.y) < 5);
    if (!row) {
      row = { y: item.y, items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }

  rows.sort((a, b) => b.y - a.y);
  return rows
    .map((row) =>
      row.items
        .sort((a, b) => a.x - b.x)
        .map((item) => item.str)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim()
    )
    .filter(Boolean);
}

/**
 * Group text items into a table-like matrix, using clustered x positions as
 * columns. Used by the PDF → Excel converter.
 */
async function extractPageRows(page) {
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const items = content.items
    .filter((item) => item.str && item.str.trim())
    .map((item) => {
      const transform = pdfjs.Util.transform(viewport.transform, item.transform);
      return { x: transform[4], y: transform[5], str: item.str };
    });

  if (!items.length) return [];

  const anchors = [];
  for (const item of [...items].sort((a, b) => a.x - b.x)) {
    const existing = anchors.find((anchor) => Math.abs(anchor - item.x) < 20);
    if (existing === undefined) anchors.push(item.x);
  }
  anchors.sort((a, b) => a - b);

  const rows = [];
  for (const item of items) {
    let row = rows.find((candidate) => Math.abs(candidate.y - item.y) < 5);
    if (!row) {
      row = { y: item.y, items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }
  rows.sort((a, b) => b.y - a.y);

  return rows.map((row) => {
    const cells = new Array(anchors.length).fill("");
    for (const item of row.items.sort((a, b) => a.x - b.x)) {
      let bestIndex = 0;
      let bestDistance = Infinity;
      anchors.forEach((anchor, index) => {
        const distance = Math.abs(anchor - item.x);
        if (distance < bestDistance) {
          bestDistance = distance;
          bestIndex = index;
        }
      });
      cells[bestIndex] = cells[bestIndex] ? `${cells[bestIndex]} ${item.str}` : item.str;
    }
    return cells;
  });
}

async function withPdfjs(buffer, callback) {
  const doc = await openPdf(rootBuffer(buffer));
  try {
    return await callback(doc);
  } finally {
    try {
      await doc.destroy();
    } catch (error) {
      /* ignore */
    }
  }
}

function rootBuffer(file) {
  // pdf.js transfers the ArrayBuffer it receives, so always hand it a copy.
  return Buffer.from(file.buffer);
}

/* ------------------------------------------------------------------ */
/* PDF helper                                                          */
/* ------------------------------------------------------------------ */

async function loadPdf(file) {
  try {
    return await PDFDocument.load(file.buffer, { ignoreEncryption: true, updateMetadata: false });
  } catch (error) {
    throw new HttpError(400, `"${file.originalname}" is not a valid PDF file.`);
  }
}

function sanitizePdfText(text) {
  // pdf-lib's standard fonts are WinAnsi only. Replace anything they cannot
  // encode instead of letting the whole conversion fail.
  return String(text == null ? "" : text)
    .replace(/\r\n?/g, "\n")
    .replace(/[^\x09\x0A\x0D\x20-\xFF]/g, "?");
}

function colorFromHex(hex, fallback = { r: 0.91, g: 0.28, b: 0.23 }) {
  if (typeof hex === "string" && /^#[0-9a-f]{6}$/i.test(hex)) {
    return rgb(
      parseInt(hex.slice(1, 3), 16) / 255,
      parseInt(hex.slice(3, 5), 16) / 255,
      parseInt(hex.slice(5, 7), 16) / 255
    );
  }
  return rgb(fallback.r, fallback.g, fallback.b);
}

/* ------------------------------------------------------------------ */
/* PDFKit helpers                                                      */
/* ------------------------------------------------------------------ */

const FONT_CANDIDATES = {
  regular: [
    "C:\\Windows\\Fonts\\Nirmala.ttf",
    "C:\\Windows\\Fonts\\arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
  ],
  bold: [
    "C:\\Windows\\Fonts\\NirmalaB.ttf",
    "C:\\Windows\\Fonts\\arialbd.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
  ],
};

let resolvedFonts;
function findFonts() {
  if (resolvedFonts) return resolvedFonts;
  const fs = require("fs");
  const pick = (list) => list.find((candidate) => fs.existsSync(candidate)) || null;
  const regular = pick(FONT_CANDIDATES.regular);
  const bold = pick(FONT_CANDIDATES.bold) || regular;
  resolvedFonts = regular
    ? {
        regular: "Body",
        bold: bold === regular ? "Body" : "Body-Bold",
        regularPath: regular,
        boldPath: bold,
      }
    : { regular: "Helvetica", bold: "Helvetica-Bold", regularPath: null, boldPath: null };
  return resolvedFonts;
}

function newPdfDocument(options = {}) {
  const doc = new PDFKit({
    size: options.size || "A4",
    layout: options.layout || "portrait",
    margin: options.margin == null ? 50 : options.margin,
    bufferPages: true,
    info: { Title: options.title || "PDF Machan", Author: "PDF Machan" },
  });
  const fonts = findFonts();
  if (fonts.regularPath) {
    try {
      doc.registerFont("Body", fonts.regularPath);
      if (fonts.boldPath && fonts.boldPath !== fonts.regularPath) {
        doc.registerFont("Body-Bold", fonts.boldPath);
      }
    } catch (error) {
      /* fall back to Helvetica below */
    }
  }
  return doc;
}

function collectPdf(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });
}

function safePdfKitText(text) {
  return String(text == null ? "" : text)
    .replace(/\r\n?/g, "\n")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, " ");
}

function decodeEntities(text) {
  const named = {
    amp: "&",
    lt: "<",
    gt: ">",
    quot: '"',
    apos: "'",
    nbsp: " ",
    hellip: "…",
    mdash: "—",
    ndash: "–",
    rsquo: "’",
    lsquo: "‘",
    ldquo: "“",
    rdquo: "”",
    copy: "©",
    reg: "®",
    trade: "™",
  };
  return String(text).replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (match, entity) => {
    const key = entity.toLowerCase();
    if (named[key]) return named[key];
    if (entity[0] === "#") {
      const code = key[1] === "x" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
      if (Number.isFinite(code) && code > 0 && code <= 0x10ffff) {
        try {
          return String.fromCodePoint(code);
        } catch (error) {
          return match;
        }
      }
    }
    return match;
  });
}

function dataUrlToBuffer(dataUrl) {
  const match = /^data:([^;,]+)?(;base64)?,([\s\S]*)$/.exec(String(dataUrl || ""));
  if (!match) throw new HttpError(400, "A watermark image could not be read.");
  if (match[2]) return Buffer.from(match[3], "base64");
  return Buffer.from(decodeURIComponent(match[3]), "utf8");
}

function imageKind(buffer) {
  if (buffer.length > 8 && buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) {
    return "png";
  }
  if (buffer.length > 3 && buffer[0] === 0xff && buffer[1] === 0xd8) return "jpeg";
  return null;
}

/* ------------------------------------------------------------------ */
/* PDF compression helpers                                             */
/* ------------------------------------------------------------------ */

function pdfNameString(value) {
  if (!value) return "";
  try {
    return String(value.toString()).replace(/^\/+/, "");
  } catch (error) {
    return "";
  }
}

function pdfNumberValue(value, fallback = 0) {
  if (!value) return fallback;
  try {
    if (typeof value.asNumber === "function") {
      const number = value.asNumber();
      if (Number.isFinite(number)) return number;
    }
    const parsed = Number(String(value));
    return Number.isFinite(parsed) ? parsed : fallback;
  } catch (error) {
    return fallback;
  }
}

function pdfFilterNames(dict) {
  let filter;
  try {
    filter = dict.lookup(PDFName.of("Filter"));
  } catch (error) {
    filter = dict.get(PDFName.of("Filter"));
  }
  if (!filter) return [];
  if (filter instanceof PDFName) return [pdfNameString(filter)];
  if (filter instanceof PDFArray) {
    const names = [];
    for (let index = 0; index < filter.size(); index += 1) {
      try {
        names.push(pdfNameString(filter.lookup(index, PDFName)));
      } catch (error) {
        /* ignore malformed array entries */
      }
    }
    return names;
  }
  return [];
}

function pdfColorSpaceChannels(dict, context) {
  let colorSpace;
  try {
    colorSpace = dict.lookup(PDFName.of("ColorSpace"));
  } catch (error) {
    colorSpace = dict.get(PDFName.of("ColorSpace"));
  }
  if (!colorSpace) return null;

  if (colorSpace instanceof PDFName) {
    const name = pdfNameString(colorSpace);
    if (name === "DeviceRGB") return 3;
    if (name === "DeviceGray") return 1;
    if (name === "DeviceCMYK") return 4;
    return null;
  }

  if (colorSpace instanceof PDFArray) {
    let family = "";
    try {
      family = pdfNameString(colorSpace.lookup(0, PDFName));
    } catch (error) {
      family = "";
    }
    if (family === "CalRGB") return 3;
    if (family === "CalGray") return 1;
    if (family === "ICCBased") {
      try {
        const profile = context.lookup(colorSpace.get(1));
        const components = profile && profile.dict && profile.dict.get(PDFName.of("N"));
        const channels = pdfNumberValue(components, 0);
        if (channels === 1 || channels === 3 || channels === 4) return channels;
      } catch (error) {
        /* fall through – unknown ICC profile */
      }
    }
  }
  return null;
}

function pdfHasPredictor(dict) {
  let parms;
  try {
    parms = dict.lookup(PDFName.of("DecodeParms"));
  } catch (error) {
    return false;
  }
  const check = (entry) => {
    if (!(entry instanceof PDFDict)) return false;
    const predictor = entry.lookup(PDFName.of("Predictor"));
    return Boolean(predictor) && pdfNumberValue(predictor, 1) !== 1;
  };
  if (parms instanceof PDFDict) return check(parms);
  if (parms instanceof PDFArray) {
    for (let index = 0; index < parms.size(); index += 1) {
      try {
        if (check(parms.lookup(index, PDFDict))) return true;
      } catch (error) {
        /* ignore malformed entry */
      }
    }
  }
  return false;
}

function targetImageDimension(quality) {
  if (quality >= 0.8) return 2400;
  if (quality >= 0.6) return 1900;
  if (quality >= 0.42) return 1500;
  return 1150;
}

/**
 * Recompress every safe embedded raster image in a PDF. Images are decoded,
 * optionally downscaled, re-encoded as JPEG and only replaced when the new
 * stream is genuinely smaller. This keeps text as vectors (sharp) while still
 * shrinking scan-heavy documents dramatically.
 */
async function optimizePdfImages(doc, quality) {
  const context = doc.context;
  const targetDimension = targetImageDimension(quality);
  const jpegQuality = Math.round(clamp(quality, 0.2, 0.95) * 100);
  const stats = { changed: 0, before: 0, after: 0, skipped: 0, imagesSeen: 0 };

  for (const [ref, object] of context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFRawStream)) continue;
    const dict = object.dict;

    let subtype = "";
    try {
      subtype = pdfNameString(dict.get(PDFName.of("Subtype")));
    } catch (error) {
      subtype = "";
    }
    if (subtype !== "Image") continue;
    stats.imagesSeen += 1;

    // Masks and explicitly decoded images must be left untouched: rewriting
    // them without their companion mask would corrupt transparency.
    if (dict.get(PDFName.of("SMask")) || dict.get(PDFName.of("Mask")) || dict.get(PDFName.of("Decode"))) {
      stats.skipped += 1;
      continue;
    }
    const imageMask = dict.get(PDFName.of("ImageMask"));
    if (imageMask && String(imageMask) === "true") {
      stats.skipped += 1;
      continue;
    }

    const filters = pdfFilterNames(dict);
    if (filters.length !== 1 || !["DCTDecode", "FlateDecode"].includes(filters[0])) {
      stats.skipped += 1;
      continue;
    }
    if (pdfHasPredictor(dict)) {
      stats.skipped += 1;
      continue;
    }

    const sourceWidth = Math.round(pdfNumberValue(dict.get(PDFName.of("Width")), 0));
    const sourceHeight = Math.round(pdfNumberValue(dict.get(PDFName.of("Height")), 0));
    const bits = pdfNumberValue(dict.get(PDFName.of("BitsPerComponent")), 8);
    if (!sourceWidth || !sourceHeight || bits !== 8) {
      stats.skipped += 1;
      continue;
    }

    const declaredChannels = pdfColorSpaceChannels(dict, context);
    if (declaredChannels === 4) {
      stats.skipped += 1;
      continue;
    }

    const original = Buffer.from(object.contents);
    const originalSize = original.length;

    let pipeline;
    let width = sourceWidth;
    let height = sourceHeight;
    let channels = declaredChannels || 3;

    try {
      if (filters[0] === "DCTDecode") {
        const metadata = await sharp(original, { failOn: "none" }).metadata();
        if (!metadata.width || !metadata.height) throw new Error("no jpeg dimensions");
        pipeline = sharp(original, { failOn: "none" });
        width = metadata.width;
        height = metadata.height;
        channels = metadata.channels || declaredChannels || 3;
      } else {
        let decoded;
        try {
          decoded = decodePDFRawStream(object).decode();
        } catch (error) {
          stats.skipped += 1;
          continue;
        }
        const samples = Buffer.from(decoded);
        if (declaredChannels) {
          channels = declaredChannels;
        } else if (samples.length === width * height * 3) {
          channels = 3;
        } else if (samples.length === width * height) {
          channels = 1;
        } else {
          stats.skipped += 1;
          continue;
        }
        const needed = width * height * channels;
        if (samples.length < needed) {
          stats.skipped += 1;
          continue;
        }
        pipeline = sharp(samples.subarray(0, needed), {
          raw: { width, height, channels },
          failOn: "none",
        });
      }

      if (channels === 4) {
        stats.skipped += 1;
        continue;
      }

      const maxDimension = Math.max(width, height);
      const scale = Math.min(1, targetDimension / maxDimension);
      const outWidth = Math.max(1, Math.round(width * scale));
      const outHeight = Math.max(1, Math.round(height * scale));

      // Nothing to gain on tiny, cheap images unless the user asked for very
      // aggressive compression.
      if (scale === 1 && quality > 0.9) {
        stats.skipped += 1;
        continue;
      }

      pipeline = pipeline.resize({ width: outWidth, height: outHeight, fit: "fill" });
      if (channels === 1) pipeline = pipeline.grayscale();
      else pipeline = pipeline.flatten({ background: "#ffffff" });
      pipeline = pipeline.jpeg({ quality: jpegQuality, mozjpeg: true, chromaSubsampling: jpegQuality < 88 ? "4:2:0" : "4:4:4" });

      const jpeg = await pipeline.toBuffer();
      if (jpeg.length >= originalSize * 0.97) {
        stats.skipped += 1;
        continue;
      }

      dict.set(PDFName.of("Filter"), PDFName.of("DCTDecode"));
      dict.delete(PDFName.of("DecodeParms"));
      dict.delete(PDFName.of("Decode"));
      dict.set(PDFName.of("Width"), PDFNumber.of(outWidth));
      dict.set(PDFName.of("Height"), PDFNumber.of(outHeight));
      dict.set(PDFName.of("ColorSpace"), PDFName.of(channels === 1 ? "DeviceGray" : "DeviceRGB"));
      dict.set(PDFName.of("BitsPerComponent"), PDFNumber.of(8));
      dict.delete(PDFName.of("Interpolate"));

      context.assign(ref, PDFRawStream.of(dict, new Uint8Array(jpeg)));
      stats.changed += 1;
      stats.before += originalSize;
      stats.after += jpeg.length;
    } catch (error) {
      stats.skipped += 1;
    }
  }

  return stats;
}

async function rasterizePdf(file, quality) {
  const scale = quality >= 0.75 ? 2 : quality >= 0.5 ? 1.6 : quality >= 0.32 ? 1.3 : 1.1;
  const jpegQuality = Math.round(clamp(quality, 0.25, 0.95) * 100);
  const out = await PDFDocument.create();

  await withPdfjs(file, async (proxy) => {
    for (let index = 1; index <= proxy.numPages; index += 1) {
      const page = await proxy.getPage(index);
      const original = page.getViewport({ scale: 1 });
      const { canvas } = await renderPage(page, scale);
      const jpeg = canvas.toBuffer("image/jpeg", jpegQuality);
      const image = await out.embedJpg(jpeg);
      const outputPage = out.addPage([original.width, original.height]);
      outputPage.drawImage(image, {
        x: 0,
        y: 0,
        width: original.width,
        height: original.height,
      });
    }
  });

  return Buffer.from(await out.save());
}

/* ------------------------------------------------------------------ */
/* individual tools                                                    */
/* ------------------------------------------------------------------ */

async function mergePdfs({ files, meta }) {
  if (!files || !files.length) throw new HttpError(400, "Please select at least one PDF file.");

  const order = Array.isArray(meta.order) ? meta.order : null;
  const out = await PDFDocument.create();

  if (order) {
    if (!order.length) throw new HttpError(400, "Keep at least one page before merging.");
    const sources = new Map();
    for (const entry of order) {
      const fileIndex = Array.isArray(entry) ? parseInt(entry[0], 10) : parseInt(entry && entry.fileIndex, 10);
      const pageIndex = Array.isArray(entry) ? parseInt(entry[1], 10) : parseInt(entry && entry.pageIndex, 10);
      if (!Number.isFinite(fileIndex) || !Number.isFinite(pageIndex)) continue;
      if (fileIndex < 0 || fileIndex >= files.length) continue;
      if (!sources.has(fileIndex)) sources.set(fileIndex, await loadPdf(files[fileIndex]));
      const source = sources.get(fileIndex);
      if (pageIndex < 0 || pageIndex >= source.getPageCount()) continue;
      const [page] = await out.copyPages(source, [pageIndex]);
      out.addPage(page);
    }
  } else {
    for (const file of files) {
      const source = await loadPdf(file);
      const pages = await out.copyPages(source, source.getPageIndices());
      pages.forEach((page) => out.addPage(page));
    }
  }

  if (out.getPageCount() === 0) throw new HttpError(400, "The selected PDFs do not contain any pages.");
  return pdfResult(out, "merged.pdf");
}

async function splitPdf({ files, meta }) {
  const source = await loadPdf(getOne(files));
  const total = source.getPageCount();
  if (!total) throw new HttpError(400, "That PDF does not contain any pages.");

  const mode = ["each", "select", "range"].includes(meta.mode) ? meta.mode : "range";
  const requestedPages = Array.isArray(meta.pages)
    ? [...new Set(meta.pages.map((value) => parseInt(value, 10)).filter((value) => Number.isFinite(value) && value >= 1 && value <= total))]
    : [];

  if (mode === "each") {
    const indices = requestedPages.length ? requestedPages.map((page) => page - 1) : Array.from({ length: total }, (_, index) => index);
    const zip = new JSZip();
    for (const index of indices) {
      const out = await PDFDocument.create();
      const [page] = await out.copyPages(source, [index]);
      out.addPage(page);
      zip.file(`page_${index + 1}.pdf`, await out.save());
    }
    return {
      buffer: await zip.generateAsync({ type: "nodebuffer" }),
      filename: `${baseName(files[0].originalname)}_pages.zip`,
      mime: "application/zip",
    };
  }

  if (mode === "select") {
    if (!requestedPages.length) throw new HttpError(400, "Select at least one page to extract.");
    const out = await PDFDocument.create();
    const pages = await out.copyPages(source, requestedPages.map((page) => page - 1));
    pages.forEach((page) => out.addPage(page));
    return pdfResult(out, `${baseName(files[0].originalname)}_selected_pages.pdf`);
  }

  const start = clamp(parseInt(meta.start, 10) || 1, 1, total);
  let end = clamp(parseInt(meta.end, 10) || total, 1, total);
  if (start > end) end = start;
  const indices = [];
  for (let index = start - 1; index < end; index += 1) indices.push(index);

  const out = await PDFDocument.create();
  const pages = await out.copyPages(source, indices);
  pages.forEach((page) => out.addPage(page));
  return pdfResult(out, `${baseName(files[0].originalname)}_pages_${start}-${end}.pdf`);
}

async function compressPdf({ files, meta }) {
  const file = getOne(files);
  if (!file.buffer || !file.buffer.length) throw new HttpError(400, "That PDF is empty.");
  const quality = clamp(Number(meta.quality) || 0.75, 0.2, 0.95);
  const candidates = [{ buffer: file.buffer, strategy: "original", note: "already optimised" }];
  let opened = false;
  let imageStats = null;
  let optimizedBuffer = null;

  // 1. Rewrite the embedded images. Text and vectors stay sharp.
  try {
    const doc = await PDFDocument.load(file.buffer, { updateMetadata: false });
    opened = doc.getPageCount() > 0;
    if (opened) {
      imageStats = await optimizePdfImages(doc, quality);
      if (imageStats.changed > 0) {
        optimizedBuffer = Buffer.from(await doc.save());
        if (optimizedBuffer.length < file.buffer.length) {
          candidates.push({ buffer: optimizedBuffer, strategy: "image-optimization", note: `${imageStats.changed} image(s) optimised` });
        }
      }
    }
  } catch (error) {
    /* try the raster path below */
  }

  // 2. Raster fallback for image-only scans where re-encoding alone is not
  // enough. Skip it when the document has no raster images (rasterising a
  // text PDF can only make it bigger) or when image rewriting already worked.
  const imageRewriteWorked =
    Boolean(optimizedBuffer) && imageStats && imageStats.changed > 0 && optimizedBuffer.length < file.buffer.length * 0.9;
  const shouldTryRaster = !imageStats ? true : imageStats.imagesSeen === 0 ? false : !imageRewriteWorked;

  if (shouldTryRaster) {
    try {
      const raster = await rasterizePdf(file, quality);
      opened = opened || raster.length > 0;
      if (raster.length < file.buffer.length) {
        candidates.push({ buffer: raster, strategy: "raster", note: "page raster compression" });
      }
    } catch (error) {
      /* the original document may be encrypted or malformed */
    }
  }

  if (!opened) {
    throw new HttpError(400, "That PDF could not be opened. It may be damaged or password protected.");
  }

  candidates.sort((left, right) => left.buffer.length - right.buffer.length);
  if (process.env.DEBUG_COMPRESS === "1") {
    console.log("[compress] candidates", candidates.map((candidate) => `${candidate.strategy}:${candidate.buffer.length}`).join(", "));
  }
  const optimized = candidates.find((candidate) => candidate.strategy === "image-optimization");
  const raster = candidates.find((candidate) => candidate.strategy === "raster");

  // Prefer image optimisation because it keeps text and vector art sharp.
  // The raster path is only used when it is dramatically smaller (e.g. the
  // PDF uses an image filter we cannot rewrite) or when nothing else helped.
  let best;
  if (optimized && raster) {
    best = raster.buffer.length < optimized.buffer.length * 0.55 ? raster : optimized;
  } else if (optimized) {
    best = optimized;
  } else if (raster) {
    best = raster;
  } else {
    best = candidates[0];
  }

  return {
    buffer: best.buffer,
    filename: `compressed_${baseName(file.originalname)}.pdf`,
    mime: "application/pdf",
    info: {
      originalSize: file.buffer.length,
      outputSize: best.buffer.length,
      strategy: best.strategy,
      note: best.note,
    },
  };
}

async function rotatePdf({ files, meta }) {
  const file = getOne(files);
  const doc = await loadPdf(file);
  const pages = doc.getPages();
  if (!pages.length) throw new HttpError(400, "That PDF does not contain any pages.");
  const degreesToAdd = ((Number(meta.degrees) || 90) % 360 + 360) % 360;
  const requested = Array.isArray(meta.pages)
    ? meta.pages.map((value) => parseInt(value, 10)).filter((value) => Number.isFinite(value))
    : [];
  const targets = requested.length ? requested : pages.map((page, index) => index + 1);

  for (const pageNumber of targets) {
    const page = pages[pageNumber - 1];
    if (!page) continue;
    const current = page.getRotation().angle || 0;
    page.setRotation(degrees(((current + degreesToAdd) % 360 + 360) % 360));
  }

  return pdfResult(doc, `rotated_${baseName(file.originalname)}.pdf`);
}

async function protectPdf({ files, meta }) {
  const file = getOne(files);
  const password = String(meta.password || "");
  if (!password) throw new HttpError(400, "Please enter a password.");
  if (password.length < 4) throw new HttpError(400, "Use a password with at least 4 characters.");

  const doc = await loadPdf(file);
  doc.encrypt({
    userPassword: password,
    ownerPassword: `${password}::pdfmachan`,
    permissions: {
      printing: "highResolution",
      modifying: false,
      copying: false,
      annotating: false,
      fillingForms: false,
      contentAccessibility: true,
      documentAssembly: false,
    },
  });
  return pdfResult(doc, `protected_${baseName(file.originalname)}.pdf`);
}

async function editPdf({ files, meta }) {
  const file = getOne(files);
  const edits = Array.isArray(meta.edits) ? meta.edits : [];
  if (!edits.length) throw new HttpError(400, "No text changes were submitted. Edit some text first.");

  const doc = await loadPdf(file);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const pages = doc.getPages();

  for (const block of edits) {
    const page = pages[(parseInt(block.page, 10) || 0) - 1];
    if (!page) continue;

    const rawTransform = Array.isArray(block.rawTransform) ? block.rawTransform : [1, 0, 0, 1, 0, 0];
    const originalText = String(block.originalText || "");
    const nextText = String(block.text == null ? originalText : block.text);
    const fontSize = Math.max(1, Number(block.fontSize) || 12);
    const blockWidth = Number(block.blockWidth) || 0;
    const origX = Number(block.origX) || 0;
    const origY = Number(block.origY) || 0;
    const rotation = Number(block.customRotation) || 0;

    const originallyVertical = Math.abs(rawTransform[1]) > 0.1 || Math.abs(rawTransform[2]) > 0.1;
    const baseWidth = blockWidth > 0 ? blockWidth : originalText.length * fontSize * 0.55;
    const baseHeight = fontSize;

    let maskWidth;
    let maskHeight;
    let maskXOffset;
    let maskYOffset;
    if (originallyVertical) {
      maskWidth = baseHeight + 16;
      maskHeight = baseWidth + 30;
      maskXOffset = maskWidth / 2;
      maskYOffset = 10;
    } else {
      maskWidth = baseWidth + 32;
      maskHeight = baseHeight + 16;
      maskXOffset = 16;
      maskYOffset = 8;
    }

    page.drawRectangle({
      x: origX - maskXOffset,
      y: origY - maskYOffset,
      width: maskWidth,
      height: maskHeight,
      color: rgb(1, 1, 1),
      opacity: 1,
    });

    const text = sanitizePdfText(nextText);
    if (text.trim().length) {
      try {
        page.drawText(text, {
          x: origX,
          y: origY,
          size: fontSize,
          font,
          color: rgb(0, 0, 0),
          rotate: degrees(rotation),
        });
      } catch (error) {
        /* a single pathological glyph should not abort the whole export */
      }
    }
  }

  return pdfResult(doc, `edited_${baseName(file.originalname)}.pdf`);
}

async function watermarkPdf({ files, meta }) {
  const file = getOne(files);
  const layers = Array.isArray(meta.layers) ? meta.layers : [];
  if (!layers.length) throw new HttpError(400, "Add at least one watermark or stamp first.");

  const doc = await loadPdf(file);
  const pages = doc.getPages();
  const brand = meta.brand === "image" ? "image" : "pdf";
  const fallback = brand === "image" ? { r: 0.16, g: 0.5, b: 0.89 } : { r: 0.91, g: 0.28, b: 0.23 };
  const scaleFactor = 1.35;

  for (const layer of layers) {
    const page = pages[(parseInt(layer.page, 10) || 1) - 1];
    if (!page) continue;
    const { height: pageHeight } = page.getSize();
    const absoluteX = (Number(layer.scaleX) || 0) / scaleFactor;
    const absoluteY = (pageHeight * scaleFactor - (Number(layer.scaleY) || 0)) / scaleFactor;

    if (layer.type === "text") {
      const family = layer.fontNameFamily;
      const fontKey =
        family === "Courier"
          ? StandardFonts.CourierBold
          : family === "Times"
            ? StandardFonts.TimesRomanBold
            : StandardFonts.HelveticaBold;
      const embeddedFont = await doc.embedFont(fontKey);
      const size = Math.max(4, Number(layer.fontSize) || 24);
      const text = sanitizePdfText(layer.text);
      if (!text.trim()) continue;
      try {
        page.drawText(text, {
          x: absoluteX,
          y: absoluteY - size * 0.8,
          size,
          font: embeddedFont,
          color: colorFromHex(layer.fontColorHex, fallback),
          opacity: 0.75,
        });
      } catch (error) {
        /* ignore unencodable text */
      }
    } else if (layer.type === "image" && layer.dataUrl) {
      const raw = dataUrlToBuffer(layer.dataUrl);
      let pngBuffer;
      try {
        pngBuffer = await sharp(raw).rotate().png().toBuffer();
      } catch (error) {
        throw new HttpError(400, "The watermark image could not be processed.");
      }
      const image = await doc.embedPng(pngBuffer);
      const width = Math.max(4, Number(layer.width) || 140) / scaleFactor;
      const height = Math.max(4, Number(layer.height) || 70) / scaleFactor;
      page.drawImage(image, {
        x: absoluteX,
        y: absoluteY - height,
        width,
        height,
        opacity: 1,
      });
    }
  }

  return pdfResult(doc, `watermarked_${baseName(file.originalname)}.pdf`);
}

async function addPageNumbers({ files, meta }) {
  const file = getOne(files);
  const doc = await loadPdf(file);
  const pages = doc.getPages();
  if (!pages.length) throw new HttpError(400, "That PDF does not contain any pages.");

  const font = await doc.embedFont(meta.bold ? StandardFonts.HelveticaBold : StandardFonts.Helvetica);
  const total = pages.length;
  const position = String(meta.position || "bottom-center");
  const format = String(meta.format || "number");
  const startNumber = Math.max(0, parseInt(meta.startNumber, 10) || 1);
  const fontSize = clamp(Number(meta.fontSize) || 12, 6, 72);
  const margin = clamp(Number(meta.margin) || 28, 4, 150);
  const color = colorFromHex(meta.colorHex, { r: 0.12, g: 0.13, b: 0.16 });

  const requested = Array.isArray(meta.pages)
    ? new Set(meta.pages.map((value) => parseInt(value, 10)).filter((value) => Number.isFinite(value) && value >= 1 && value <= total))
    : null;

  for (let index = 0; index < total; index += 1) {
    const pageNumber = index + 1;
    if (requested && requested.size && !requested.has(pageNumber)) continue;
    const page = pages[index];
    const { width, height } = page.getSize();
    const value = startNumber + index;

    let label;
    if (format === "number-of-total") label = `${value} / ${total}`;
    else if (format === "page-number") label = `Page ${value}`;
    else if (format === "page-number-of-total") label = `Page ${value} of ${total}`;
    else label = String(value);

    const textWidth = font.widthOfTextAtSize(label, fontSize);
    const vertical = position.startsWith("top") ? "top" : "bottom";
    const horizontal = position.endsWith("left") ? "left" : position.endsWith("right") ? "right" : "center";

    let x;
    if (horizontal === "left") x = margin;
    else if (horizontal === "right") x = width - margin - textWidth;
    else x = (width - textWidth) / 2;

    const y = vertical === "top" ? height - margin - fontSize : margin;

    page.drawText(label, {
      x: Math.max(2, Math.min(x, width - textWidth - 2)),
      y: Math.max(2, y),
      size: fontSize,
      font,
      color,
    });
  }

  return pdfResult(doc, `numbered_${baseName(file.originalname)}.pdf`);
}

/* ------------------------------------------------------------------ */
/* Word                                                                */
/* ------------------------------------------------------------------ */

function htmlToBlocks(html) {
  const images = [];
  let working = String(html || "");

  working = working.replace(/<img\b[^>]*src=["']([^"']+)["'][^>]*>/gi, (match, src) => {
    if (/^data:/i.test(src)) {
      try {
        images.push({ type: "image", buffer: dataUrlToBuffer(src) });
        return `\n[[IMG:${images.length - 1}]]\n`;
      } catch (error) {
        return "\n";
      }
    }
    return "\n";
  });

  working = working.replace(/<h([1-6])[^>]*>([\s\S]*?)<\/h\1>/gi, (match, level, inner) => {
    return `\n[[H${level}]]${inner.replace(/<[^>]+>/g, "")}\n`;
  });
  working = working.replace(/<li[^>]*>/gi, "\n• ");
  working = working.replace(/<br\s*\/?>/gi, "\n");
  working = working.replace(/<\/(p|div|tr|blockquote|ul|ol|table)>/gi, "\n\n");
  working = working.replace(/<[^>]+>/g, "");
  working = decodeEntities(working);

  const blocks = [];
  for (const rawLine of working.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    const imageMatch = /^\[\[IMG:(\d+)\]\]$/.exec(line);
    if (imageMatch) {
      const image = images[Number(imageMatch[1])];
      if (image) blocks.push(image);
      continue;
    }
    const headingMatch = /^\[\[H([1-6])\]\]([\s\S]*)$/.exec(line);
    if (headingMatch) {
      const text = headingMatch[2].trim();
      if (text) blocks.push({ type: "text", text, heading: Number(headingMatch[1]) });
      continue;
    }
    blocks.push({ type: "text", text: line });
  }
  return blocks;
}

async function wordToPdf({ files }) {
  const file = getOne(files);

  // Prefer Word / LibreOffice when installed: their own PDF exporter keeps
  // styles, tables, headers, footers and images exactly as the author sees them.
  const native = await convertOfficeDocument("word2pdf", file.buffer, file.originalname);
  if (native) {
    return {
      buffer: native,
      filename: `${baseName(file.originalname)}.pdf`,
      mime: "application/pdf",
    };
  }

  let result;
  try {
    result = await mammoth.convertToHtml({ buffer: file.buffer });
  } catch (error) {
    throw new HttpError(400, "That Word document could not be read. Please use a .docx file.");
  }

  const blocks = htmlToBlocks(result.value);
  const doc = newPdfDocument({ title: baseName(file.originalname) });
  const done = collectPdf(doc);
  const fonts = findFonts();
  const usableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

  if (!blocks.length) {
    doc.font(fonts.regular).fontSize(12).text("(This document did not contain any readable text.)");
  }

  for (const block of blocks) {
    if (block.type === "image") {
      try {
        doc.image(block.buffer, { fit: [usableWidth, 620], align: "center" });
        doc.moveDown(0.5);
      } catch (error) {
        /* skip unreadable embedded images */
      }
      continue;
    }

    const heading = block.heading;
    const size = heading ? [0, 22, 19, 17, 15, 13, 12][heading] : 11;
    doc
      .font(heading ? fonts.bold : fonts.regular)
      .fontSize(size)
      .fillColor("#111111")
      .text(safePdfKitText(block.text), { lineGap: 2, paragraphGap: heading ? 8 : 3 });
    doc.moveDown(heading ? 0.25 : 0.35);
  }

  doc.end();
  return {
    buffer: await done,
    filename: `${baseName(file.originalname)}.pdf`,
    mime: "application/pdf",
  };
}

async function pdfToWord({ files }) {
  const file = getOne(files);

  // Word can reflow a PDF into a real .docx with its layout engine. This is
  // much closer to the original than plain text extraction, so use it first.
  const native = await convertOfficeDocument("pdf2word", file.buffer, file.originalname);
  if (native) {
    return {
      buffer: native,
      filename: `${baseName(file.originalname)}.docx`,
      mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };
  }

  const children = [];

  await withPdfjs(file, async (proxy) => {
    for (let index = 1; index <= proxy.numPages; index += 1) {
      if (index > 1) children.push(new Paragraph({ children: [new PageBreak()] }));
      const page = await proxy.getPage(index);
      const lines = await extractPageLines(page);
      if (!lines.length) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `[Page ${index} does not contain extractable text — it may be a scanned image.]`,
                italics: true,
                color: "888888",
              }),
            ],
          })
        );
        continue;
      }
      for (const line of lines) {
        children.push(new Paragraph({ children: [new TextRun(safePdfKitText(line))] }));
      }
    }
  });

  if (!children.length) children.push(new Paragraph({ children: [new TextRun("(Empty document)")] }));

  const document = new Document({
    creator: "PDF Machan",
    title: baseName(file.originalname),
    sections: [{ children }],
  });
  const buffer = await Packer.toBuffer(document);
  return {
    buffer,
    filename: `${baseName(file.originalname)}.docx`,
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  };
}

/* ------------------------------------------------------------------ */
/* Excel                                                               */
/* ------------------------------------------------------------------ */

async function xlsToPdf({ files }) {
  const file = getOne(files);

  const native = await convertOfficeDocument("xls2pdf", file.buffer, file.originalname);
  if (native) {
    return {
      buffer: native,
      filename: `${baseName(file.originalname)}.pdf`,
      mime: "application/pdf",
    };
  }

  let workbook;
  try {
    workbook = XLSX.read(file.buffer, { type: "buffer" });
  } catch (error) {
    throw new HttpError(400, "That spreadsheet could not be read. Please use a valid .xlsx/.xls file.");
  }

  const doc = newPdfDocument({
    size: "A4",
    layout: "landscape",
    margin: 30,
    title: baseName(file.originalname),
  });
  const done = collectPdf(doc);
  const fonts = findFonts();
  const left = 30;
  const right = doc.page.width - 30;
  const usableWidth = right - left;
  const bottom = doc.page.height - 30;
  const rowHeight = 20;
  let hasSheet = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    let rows;
    try {
      rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "", blankrows: false });
    } catch (error) {
      rows = [];
    }
    if (!rows.length) continue;

    if (hasSheet) doc.addPage();
    hasSheet = true;

    const colCount = Math.min(Math.max(...rows.map((row) => (Array.isArray(row) ? row.length : 0)), 1), 18);
    const rawWidths = [];
    for (let column = 0; column < colCount; column += 1) {
      let widest = 7;
      for (const row of rows) {
        const value = String((row && row[column]) || "");
        widest = Math.max(widest, Math.min(value.length, 42));
      }
      rawWidths.push(widest);
    }
    const totalWidth = rawWidths.reduce((sum, value) => sum + value, 0) || 1;
    const widths = rawWidths.map((value) => Math.max(42, (value / totalWidth) * usableWidth));
    const widthScale = usableWidth / widths.reduce((sum, value) => sum + value, 0);
    for (let index = 0; index < widths.length; index += 1) widths[index] *= widthScale;

    doc.font(fonts.bold).fontSize(14).fillColor("#111111").text(safePdfKitText(sheetName), left, 28, {
      width: usableWidth,
    });

    let y = 52;
    const drawRow = (row, header) => {
      if (y + rowHeight > bottom) {
        doc.addPage();
        y = 30;
        drawHeader();
      }
      let x = left;
      for (let column = 0; column < colCount; column += 1) {
        const width = widths[column];
        const value = safePdfKitText(String((row && row[column]) || "").replace(/\s+/g, " "));
        if (header) {
          doc.save();
          doc.rect(x, y, width, rowHeight).fill("#EFEFEF");
          doc.restore();
        }
        doc
          .rect(x, y, width, rowHeight)
          .lineWidth(0.4)
          .strokeColor("#C8C8C8")
          .stroke();
        doc
          .font(header ? fonts.bold : fonts.regular)
          .fontSize(8.5)
          .fillColor("#111111")
          .text(value, x + 3, y + 5.5, {
            width: Math.max(1, width - 6),
            height: rowHeight - 4,
            lineBreak: false,
            ellipsis: true,
          });
        x += width;
      }
      y += rowHeight;
    };

    const drawHeader = () => {
      const first = rows[0] || [];
      drawRow(first, true);
    };

    drawHeader();
    for (let index = 1; index < rows.length; index += 1) {
      if (!Array.isArray(rows[index])) continue;
      drawRow(rows[index], false);
    }
  }

  if (!hasSheet) {
    doc.font(fonts.regular).fontSize(12).text("(This workbook did not contain any data.)");
  }

  doc.end();
  return {
    buffer: await done,
    filename: `${baseName(file.originalname)}.pdf`,
    mime: "application/pdf",
  };
}

async function pdfToXls({ files }) {
  const file = getOne(files);
  const workbook = XLSX.utils.book_new();
  let sheetIndex = 0;

  await withPdfjs(file, async (proxy) => {
    for (let index = 1; index <= proxy.numPages; index += 1) {
      const page = await proxy.getPage(index);
      const rows = await extractPageRows(page);
      const data = rows.length ? rows : [["(No extractable text on this page)"]];
      const sheet = XLSX.utils.aoa_to_sheet(data);
      sheet["!cols"] = new Array(Math.max(...data.map((row) => row.length), 1))
        .fill(null)
        .map(() => ({ wch: 24 }));
      sheetIndex += 1;
      XLSX.utils.book_append_sheet(workbook, sheet, `Page ${sheetIndex}`.slice(0, 31));
    }
  });

  if (!sheetIndex) throw new HttpError(400, "That PDF does not contain any pages.");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  return {
    buffer,
    filename: `${baseName(file.originalname)}.xlsx`,
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  };
}

/* ------------------------------------------------------------------ */
/* PowerPoint                                                          */
/* ------------------------------------------------------------------ */

function slideNumber(name) {
  const match = /slide(\d+)\.xml$/.exec(name);
  return match ? parseInt(match[1], 10) : 0;
}

function parseSlideParagraphs(xml) {
  const paragraphs = [];
  const paragraphRegex = /<a:p\b[\s\S]*?<\/a:p>/g;
  let match;
  while ((match = paragraphRegex.exec(xml))) {
    const runs = [];
    const runRegex = /<a:t>([\s\S]*?)<\/a:t>/g;
    let runMatch;
    while ((runMatch = runRegex.exec(match[0]))) runs.push(decodeEntities(runMatch[1]));
    const text = runs.join("").replace(/\s+/g, " ").trim();
    if (text) paragraphs.push(text);
  }
  return paragraphs;
}

async function findSlideImage(zip, slidePath) {
  const number = slideNumber(slidePath);
  const relationshipsPath = `ppt/slides/_rels/slide${number}.xml.rels`;
  const relationshipsFile = zip.file(relationshipsPath);
  if (!relationshipsFile) return null;
  const xml = await relationshipsFile.async("string");
  const candidates = [];
  const relationshipRegex = /<Relationship\b[^>]*\/?>/g;
  let match;
  while ((match = relationshipRegex.exec(xml))) {
    const tag = match[0];
    if (!/Type="[^"]*\/image"/i.test(tag)) continue;
    const targetMatch = /Target="([^"]+)"/i.exec(tag);
    if (!targetMatch) continue;
    const target = targetMatch[1];
    if (/\.(emf|wmf|tiff?|svg)$/i.test(target)) continue;
    let resolved = target.startsWith("/")
      ? target.replace(/^\/+/, "")
      : path.posix.normalize(path.posix.join("ppt/slides", target));
    const imageFile = zip.file(resolved);
    if (!imageFile) continue;
    const buffer = await imageFile.async("nodebuffer");
    candidates.push(buffer);
  }
  if (!candidates.length) return null;
  return candidates.sort((a, b) => b.length - a.length)[0];
}

async function pptToPdf({ files }) {
  const file = getOne(files);

  const native = await convertOfficeDocument("ppt2pdf", file.buffer, file.originalname);
  if (native) {
    return {
      buffer: native,
      filename: `${baseName(file.originalname)}.pdf`,
      mime: "application/pdf",
    };
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(file.buffer);
  } catch (error) {
    throw new HttpError(400, "That PowerPoint file could not be read. Please use a .pptx file.");
  }

  const slides = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));

  if (!slides.length) throw new HttpError(400, "No slides were found in that presentation.");

  const pageWidth = 960;
  const pageHeight = 540;
  const doc = newPdfDocument({ size: [pageWidth, pageHeight], margin: 0, title: baseName(file.originalname) });
  const done = collectPdf(doc);
  const fonts = findFonts();

  for (let index = 0; index < slides.length; index += 1) {
    if (index > 0) doc.addPage({ size: [pageWidth, pageHeight], margin: 0 });
    doc.rect(0, 0, pageWidth, pageHeight).fill("#ffffff");

    let imageBuffer = null;
    try {
      imageBuffer = await findSlideImage(zip, slides[index]);
    } catch (error) {
      imageBuffer = null;
    }

    let paragraphs = [];
    try {
      const xml = await zip.file(slides[index]).async("string");
      paragraphs = parseSlideParagraphs(xml);
    } catch (error) {
      paragraphs = [];
    }

    const hasImage = Boolean(imageBuffer);
    const textAreaHeight = hasImage ? 205 : 460;
    let y = 34;

    if (hasImage) {
      try {
        const png = await sharp(imageBuffer).rotate().png().toBuffer();
        const metadata = await sharp(png).metadata();
        const maxWidth = 860;
        const maxHeight = 285;
        const scale = Math.min(maxWidth / metadata.width, maxHeight / metadata.height, 1);
        const width = metadata.width * scale;
        const height = metadata.height * scale;
        doc.image(png, (pageWidth - width) / 2, 240, { width, height });
      } catch (error) {
        /* skip unsupported slide pictures */
      }
    }

    if (!paragraphs.length) {
      doc
        .font(fonts.regular)
        .fontSize(16)
        .fillColor("#666666")
        .text(`Slide ${index + 1}`, 40, pageHeight / 2 - 10, { width: pageWidth - 80, align: "center" });
      continue;
    }

    paragraphs.forEach((paragraph, paragraphIndex) => {
      if (y > textAreaHeight) return;
      const isTitle = paragraphIndex === 0;
      const size = isTitle ? 22 : 14;
      const remaining = Math.max(18, textAreaHeight - y);
      const height = Math.min(doc.heightOfString(paragraph, { width: pageWidth - 100, lineGap: 2 }), remaining);
      doc
        .font(isTitle ? fonts.bold : fonts.regular)
        .fontSize(size)
        .fillColor(isTitle ? "#111111" : "#333333")
        .text(paragraph, 50, y, {
          width: pageWidth - 100,
          height,
          ellipsis: true,
          lineGap: 2,
        });
      y += height + (isTitle ? 10 : 6);
    });
  }

  doc.end();
  return {
    buffer: await done,
    filename: `${baseName(file.originalname)}.pdf`,
    mime: "application/pdf",
  };
}

async function pdfToPpt({ files }) {
  const file = getOne(files);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "PDF Machan";
  pptx.title = baseName(file.originalname);

  await withPdfjs(file, async (proxy) => {
    for (let index = 1; index <= proxy.numPages; index += 1) {
      const page = await proxy.getPage(index);
      const lines = await extractPageLines(page);
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };

      if (lines.length) {
        const textRuns = lines.slice(0, 60).map((line) => ({
          text: line,
          options: { breakLine: true },
        }));
        slide.addText(textRuns, {
          x: 0.5,
          y: 0.4,
          w: 12.3,
          h: 6.7,
          fontSize: lines.length > 30 ? 11 : 14,
          color: "222222",
          valign: "top",
          margin: 4,
        });
      } else {
        const { canvas } = await renderPage(page, 2);
        const png = canvas.toBuffer("image/png");
        slide.addImage({
          data: `image/png;base64,${png.toString("base64")}`,
          x: 0.35,
          y: 0.3,
          w: 12.63,
          h: 6.9,
          sizing: { type: "contain", w: 12.63, h: 6.9 },
        });
      }
    }
  });

  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return {
    buffer: Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer),
    filename: `${baseName(file.originalname)}.pptx`,
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  };
}

/* ------------------------------------------------------------------ */
/* Images                                                              */
/* ------------------------------------------------------------------ */

async function pdfToImages({ files }) {
  const file = getOne(files);
  const zip = new JSZip();

  await withPdfjs(file, async (proxy) => {
    for (let index = 1; index <= proxy.numPages; index += 1) {
      const page = await proxy.getPage(index);
      const { canvas } = await renderPage(page, 2);
      zip.file(`page_${index}.png`, canvas.toBuffer("image/png"));
    }
  });

  return {
    buffer: await zip.generateAsync({ type: "nodebuffer" }),
    filename: `${baseName(file.originalname)}_images.zip`,
    mime: "application/zip",
  };
}

async function imagesToPdf({ files }) {
  if (!files || !files.length) throw new HttpError(400, "Please select at least one image.");
  const out = await PDFDocument.create();

  for (const file of files) {
    let buffer;
    let width;
    let height;
    try {
      const source = await sharp(file.buffer).metadata();
      if (!source.width || !source.height) throw new Error("no dimensions");
      buffer =
        source.format === "png"
          ? await sharp(file.buffer).rotate().png().toBuffer()
          : await sharp(file.buffer).rotate().jpeg({ quality: 92 }).toBuffer();
      // Read the *converted* dimensions so EXIF-rotated phone photos keep the
      // correct portrait/landscape orientation on the page.
      const converted = await sharp(buffer).metadata();
      width = converted.width;
      height = converted.height;
    } catch (error) {
      throw new HttpError(400, `"${file.originalname}" is not a supported image.`);
    }

    const kind = imageKind(buffer) || "png";
    const image = kind === "png" ? await out.embedPng(buffer) : await out.embedJpg(buffer);

    const scale = Math.min(1, 1600 / Math.max(width, height));
    const pageWidth = Math.max(50, Math.round(width * scale));
    const pageHeight = Math.max(50, Math.round(height * scale));
    const page = out.addPage([pageWidth, pageHeight]);
    page.drawImage(image, { x: 0, y: 0, width: pageWidth, height: pageHeight });
  }

  return {
    buffer: Buffer.from(await out.save()),
    filename: "images_compiled.pdf",
    mime: "application/pdf",
  };
}

async function jpgToPng({ files }) {
  const file = getOne(files);
  let buffer;
  try {
    buffer = await sharp(file.buffer).rotate().png().toBuffer();
  } catch (error) {
    throw new HttpError(400, "That image could not be converted. Please select a valid JPG.");
  }
  return {
    buffer,
    filename: `${baseName(file.originalname)}.png`,
    mime: "image/png",
  };
}

async function pngToJpg({ files }) {
  const file = getOne(files);
  let buffer;
  try {
    buffer = await sharp(file.buffer)
      .rotate()
      .flatten({ background: "#ffffff" })
      .jpeg({ quality: 92, mozjpeg: true })
      .toBuffer();
  } catch (error) {
    throw new HttpError(400, "That image could not be converted. Please select a valid PNG.");
  }
  return {
    buffer,
    filename: `${baseName(file.originalname)}.jpg`,
    mime: "image/jpeg",
  };
}

async function compressImage({ files, meta }) {
  const file = getOne(files);
  const quality = Math.round(clamp((Number(meta.quality) || 0.75) * 100, 20, 95));
  let metadata;
  let buffer;
  let filename;
  let mime;
  try {
    metadata = await sharp(file.buffer).metadata();
    if (metadata.format === "png") {
      buffer = await sharp(file.buffer)
        .rotate()
        .png({ compressionLevel: 9, quality, effort: 8, palette: true })
        .toBuffer();
      filename = `${baseName(file.originalname)}.png`;
      mime = "image/png";
    } else {
      buffer = await sharp(file.buffer)
        .rotate()
        .jpeg({ quality, mozjpeg: true })
        .toBuffer();
      filename = `${baseName(file.originalname)}.jpg`;
      mime = "image/jpeg";
    }
  } catch (error) {
    throw new HttpError(400, "That image could not be compressed. Please select a valid image file.");
  }
  return { buffer, filename, mime };
}

/* ------------------------------------------------------------------ */
/* registry                                                            */
/* ------------------------------------------------------------------ */

const HANDLERS = {
  merge: mergePdfs,
  split: splitPdf,
  compress: compressPdf,
  editor: editPdf,
  watermark: watermarkPdf,
  pagenum: addPageNumbers,
  rotate: rotatePdf,
  protect: protectPdf,
  word2pdf: wordToPdf,
  pdf2word: pdfToWord,
  xls2pdf: xlsToPdf,
  pdf2xls: pdfToXls,
  ppt2pdf: pptToPdf,
  pdf2ppt: pdfToPpt,
  pdf2img: pdfToImages,
  img2pdf: imagesToPdf,
  jpg2png: jpgToPng,
  png2jpg: pngToJpg,
  imgcompress: compressImage,
};

const TOOL_NAMES = Object.keys(HANDLERS);

async function processTool(tool, files, meta) {
  const handler = HANDLERS[tool];
  if (!handler) throw new HttpError(404, "Unknown tool.");
  if (!files || !files.length) throw new HttpError(400, "Please select a file first.");
  return handler({ files, meta: meta || {} });
}

module.exports = {
  HttpError,
  TOOL_NAMES,
  processTool,
  parseJsonField,
  baseName,
  // exported for tests
  htmlToBlocks,
  extractPageLines,
  extractPageRows,
  parseSlideParagraphs,
};
