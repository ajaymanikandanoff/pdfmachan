"use strict";

const path = require("path");
const express = require("express");
const multer = require("multer");
const { HttpError, TOOL_NAMES, processTool, parseJsonField } = require("./lib/processing");

const ROOT = __dirname;
const PORT = Number(process.env.PORT) || 3000;
const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB per file
const MAX_FILES = 20;

const app = express();
app.disable("x-powered-by");
app.set("trust proxy", true);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: MAX_FILE_SIZE,
    files: MAX_FILES,
    fields: 20,
    fieldSize: 25 * 1024 * 1024, // watermark image stamps travel inside the meta JSON
    parts: MAX_FILES + 25,
  },
});

function decodeOriginalName(name) {
  if (!name) return "file";
  // Busboy hands UTF-8 filenames back as latin1; recover the real name.
  if (/[\u0080-\u00ff]/.test(name)) {
    try {
      const decoded = Buffer.from(name, "latin1").toString("utf8");
      if (!decoded.includes("\uFFFD")) return decoded;
    } catch (error) {
      /* keep the original */
    }
  }
  return name;
}

function asciiFallback(name) {
  return String(name).replace(/[^\x20-\x7e]/g, "_").replace(/["\\]/g, "_");
}

app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    service: "pdfmachan",
    version: "2.0.0",
    tools: TOOL_NAMES,
    limits: { maxFileSizeMb: MAX_FILE_SIZE / (1024 * 1024), maxFiles: MAX_FILES },
    uptimeSeconds: Math.round(process.uptime()),
  });
});

app.post("/api/tools/:tool", (req, res, next) => {
  const tool = req.params.tool;
  if (!TOOL_NAMES.includes(tool)) {
    return res.status(404).json({ error: "Unknown tool." });
  }

  upload.any()(req, res, async (error) => {
    if (error) return next(error);
    try {
      const meta = parseJsonField(req.body && req.body.meta, {});
      const files = (req.files || [])
        .filter((file) => file.fieldname === "files")
        .map((file) => ({
          buffer: file.buffer,
          originalname: decodeOriginalName(file.originalname),
          mimetype: file.mimetype,
          size: file.size,
        }));

      const result = await processTool(tool, files, meta);
      const filename = result.filename || "output";

      res.setHeader("Content-Type", result.mime || "application/octet-stream");
      res.setHeader("Content-Length", result.buffer.length);
      res.setHeader("X-Output-Filename", encodeURIComponent(filename));
      if (result.info) {
        if (Number.isFinite(result.info.originalSize)) res.setHeader("X-Original-Size", String(result.info.originalSize));
        if (Number.isFinite(result.info.outputSize)) res.setHeader("X-Output-Size", String(result.info.outputSize));
        if (result.info.strategy) res.setHeader("X-Compression-Strategy", String(result.info.strategy));
      }
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${asciiFallback(filename)}"; filename*=UTF-8''${encodeURIComponent(filename)}`
      );
      res.setHeader("Cache-Control", "no-store, max-age=0");
      res.end(result.buffer);
    } catch (processingError) {
      next(processingError);
    }
  });
});

// Only expose the public site files — never the server source or node_modules.
app.get("/", (req, res) => {
  res.setHeader("Cache-Control", "no-cache");
  res.sendFile(path.join(ROOT, "index.html"));
});
app.get("/index.html", (req, res) => {
  res.setHeader("Cache-Control", "no-cache");
  res.sendFile(path.join(ROOT, "index.html"));
});
app.get("/logo.png", (req, res) => {
  res.setHeader("Cache-Control", "public, max-age=86400");
  res.sendFile(path.join(ROOT, "logo.png"));
});
app.get("/logo1.png", (req, res) => {
  res.setHeader("Cache-Control", "public, max-age=86400");
  res.sendFile(path.join(ROOT, "logo1.png"));
});
app.get("/favicon.ico", (req, res) => {
  res.redirect(301, "/logo1.png");
});

app.use((req, res) => {
  if (req.path.startsWith("/api/")) {
    return res.status(404).json({ error: "Not found." });
  }
  res.status(404).type("text/plain").send("Not found");
});

app.use((error, req, res, next) => {
  if (res.headersSent) return next(error);

  if (error instanceof multer.MulterError) {
    if (error.code === "LIMIT_FILE_SIZE") {
      return res.status(413).json({
        error: `File is too large. The limit is ${MAX_FILE_SIZE / (1024 * 1024)} MB per file.`,
      });
    }
    if (error.code === "LIMIT_FILE_COUNT" || error.code === "LIMIT_PART_COUNT") {
      return res.status(400).json({ error: `Too many files. You can upload up to ${MAX_FILES} at once.` });
    }
    return res.status(400).json({ error: error.message });
  }

  const status = error instanceof HttpError ? error.status : error.status || 500;
  if (status >= 500) {
    console.error("[pdfmachan] processing error:", error);
  }
  res.status(status).json({
    error: status >= 500 ? "Something went wrong while processing the file. Please try again." : error.message,
  });
});

let server;
if (require.main === module) {
  server = app.listen(PORT, "0.0.0.0", () => {
    console.log(`PDF Machan server running on http://localhost:${PORT}`);
  });
  server.requestTimeout = 10 * 60 * 1000;
  server.headersTimeout = 60 * 1000;
}

module.exports = app;
