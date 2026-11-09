"use strict";

/* End-to-end API tests for every PDF Machan / Image Machan tool. */

const assert = require("assert");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const JSZip = require("jszip");
const XLSX = require("xlsx");
const sharp = require("sharp");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const PptxGenJS = require("pptxgenjs");
const napi = require("@napi-rs/canvas");

globalThis.DOMMatrix = globalThis.DOMMatrix || napi.DOMMatrix;
globalThis.Path2D = globalThis.Path2D || napi.Path2D;
globalThis.ImageData = globalThis.ImageData || napi.ImageData;
const pdfjs = require("pdfjs-dist/legacy/build/pdf.js");

const app = require("../server");

let server;
let baseUrl;

async function makePdf(lines = ["Hello PDF Machan", "Second line with more words"]) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let pageIndex = 0; pageIndex < 2; pageIndex += 1) {
    const page = doc.addPage([595, 842]);
    lines.forEach((line, index) => {
      page.drawText(index === 0 ? `${line} (page ${pageIndex + 1})` : line, {
        x: 50,
        y: 760 - index * 30,
        size: index === 0 ? 18 : 12,
        font,
        color: rgb(0, 0, 0),
      });
    });
  }
  return Buffer.from(await doc.save());
}

async function makeDocx() {
  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun({ text: "Hello from Word", bold: true })] }),
          new Paragraph({ children: [new TextRun("Second paragraph for the converter.")] }),
        ],
      },
    ],
  });
  return Packer.toBuffer(document);
}

function makeXlsx() {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Name", "Amount", "Status"],
    ["Alpha", "120", "Paid"],
    ["Beta", "75", "Pending"],
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, "Report");
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}

async function makePptx() {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText("Quarterly Update", { x: 0.5, y: 0.4, w: 8, h: 1, fontSize: 24, bold: true });
  slide.addText("Revenue grew 24% this quarter.", { x: 0.5, y: 1.6, w: 8, h: 1, fontSize: 16 });
  const second = pptx.addSlide();
  second.addText("Thank you", { x: 1, y: 1, w: 6, h: 1, fontSize: 22 });
  return Buffer.from(await pptx.write({ outputType: "nodebuffer" }));
}

async function makeJpeg() {
  return sharp({ create: { width: 320, height: 200, channels: 3, background: "#d94f3d" } })
    .jpeg({ quality: 95 })
    .toBuffer();
}

async function makePng() {
  return sharp({ create: { width: 240, height: 180, channels: 4, background: "#2a7fe2" } })
    .png()
    .toBuffer();
}

async function makeImageHeavyPdf() {
  const doc = await PDFDocument.create();
  for (let pageIndex = 0; pageIndex < 2; pageIndex += 1) {
    const width = 1200;
    const height = 900;
    const raw = Buffer.alloc(width * height * 3);
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const offset = (y * width + x) * 3;
        raw[offset] = (x + pageIndex * 60) % 256;
        raw[offset + 1] = (y * 2 + x) % 256;
        raw[offset + 2] = (x * y + pageIndex * 40) % 256;
      }
    }
    const jpeg = await sharp(raw, { raw: { width, height, channels: 3 } }).jpeg({ quality: 96 }).toBuffer();
    const image = await doc.embedJpg(jpeg);
    doc.addPage([600, 450]).drawImage(image, { x: 0, y: 0, width: 600, height: 450 });
  }
  return Buffer.from(await doc.save());
}

async function post(tool, files, meta) {
  const form = new FormData();
  for (const file of files) {
    form.append("files", new Blob([new Uint8Array(file.buffer)], { type: file.type || "application/octet-stream" }), file.name);
  }
  form.append("meta", JSON.stringify(meta || {}));
  const response = await fetch(`${baseUrl}/api/tools/${tool}`, { method: "POST", body: form });
  const buffer = Buffer.from(await response.arrayBuffer());
  return {
    status: response.status,
    buffer,
    filename: decodeURIComponent(response.headers.get("x-output-filename") || ""),
    contentType: response.headers.get("content-type") || "",
  };
}

function expectOk(result, tool) {
  let message = "";
  try {
    message = result.buffer.toString("utf8").slice(0, 300);
  } catch (error) {
    message = "";
  }
  assert.strictEqual(result.status, 200, `${tool} failed: ${result.status} ${message}`);
  assert.ok(result.buffer.length > 0, `${tool} returned an empty file`);
  return result.buffer;
}

async function expectPdf(buffer, tool, pages) {
  const doc = await PDFDocument.load(buffer);
  if (pages != null) assert.strictEqual(doc.getPageCount(), pages, `${tool} page count`);
  return doc;
}

async function extractText(buffer, password) {
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(buffer),
    password,
    useWorkerFetch: false,
    isEvalSupported: false,
    verbosity: 0,
  }).promise;
  const pages = [];
  for (let index = 1; index <= doc.numPages; index += 1) {
    const content = await (await doc.getPage(index)).getTextContent();
    pages.push(content.items.map((item) => item.str).join(" "));
  }
  await doc.destroy();
  return pages.join("\n");
}

async function main() {
  server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  baseUrl = `http://127.0.0.1:${server.address().port}`;

  const pdf = await makePdf();
  const secondPdf = await makePdf(["Only one page"]);
  const onePagePdf = await (async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    doc.addPage([595, 842]).drawText("Single page document", { x: 50, y: 700, size: 16, font });
    return Buffer.from(await doc.save());
  })();
  const docx = await makeDocx();
  const xlsx = makeXlsx();
  const pptx = await makePptx();
  const jpeg = await makeJpeg();
  const png = await makePng();
  const imageHeavyPdf = await makeImageHeavyPdf();

  const pdfFile = { name: "sample.pdf", buffer: pdf, type: "application/pdf" };
  const onePageFile = { name: "one.pdf", buffer: onePagePdf, type: "application/pdf" };

  const results = [];
  const run = async (tool, fn) => {
    await fn();
    results.push(tool);
    console.log(`  ✓ ${tool}`);
  };

  console.log("Running API tests…");

  await run("merge", async () => {
    const out = expectOk(await post("merge", [pdfFile, onePageFile]), "merge");
    await expectPdf(out, "merge", 3);
  });

  await run("merge (custom page order)", async () => {
    const order = [
      { fileIndex: 1, pageIndex: 0 },
      { fileIndex: 0, pageIndex: 0 },
      { fileIndex: 0, pageIndex: 1 },
    ];
    const out = expectOk(await post("merge", [pdfFile, onePageFile], { order }), "merge order");
    await expectPdf(out, "merge order", 3);
    const text = await extractText(out);
    const singleIndex = text.indexOf("Single page document");
    const firstIndex = text.indexOf("Hello PDF Machan (page 1)");
    const secondIndex = text.indexOf("(page 2)");
    assert.ok(singleIndex >= 0, "merged first page should be the single-page PDF");
    assert.ok(singleIndex < firstIndex && firstIndex < secondIndex, "page order should follow the requested map");
  });

  await run("split (range)", async () => {
    const out = expectOk(await post("split", [pdfFile], { mode: "range", start: 1, end: 1 }), "split range");
    await expectPdf(out, "split range", 1);
  });

  await run("split (selected pages)", async () => {
    const out = expectOk(await post("split", [pdfFile], { mode: "select", pages: [2, 1] }), "split select");
    await expectPdf(out, "split select", 2);
    const text = await extractText(out);
    assert.ok(text.indexOf("(page 2)") < text.indexOf("(page 1)"), "selected pages should keep the requested order");
  });

  await run("split (each)", async () => {
    const out = expectOk(await post("split", [pdfFile], { mode: "each" }), "split each");
    const zip = await JSZip.loadAsync(out);
    const entries = Object.keys(zip.files).filter((name) => name.endsWith(".pdf"));
    assert.strictEqual(entries.length, 2, "split each should return 2 PDFs");
  });

  await run("split (each selected)", async () => {
    const out = expectOk(await post("split", [pdfFile], { mode: "each", pages: [2] }), "split each selected");
    const zip = await JSZip.loadAsync(out);
    const entries = Object.keys(zip.files).filter((name) => name.endsWith(".pdf"));
    assert.strictEqual(entries.length, 1, "only the selected page should be split");
    assert.strictEqual(entries[0], "page_2.pdf");
  });

  await run("compress", async () => {
    const out = expectOk(await post("compress", [pdfFile], { quality: 0.4 }), "compress");
    await expectPdf(out, "compress", 2);
    assert.ok(out.length <= pdf.length, "a text PDF must never become larger");
  });

  await run("compress (image heavy)", async () => {
    const out = expectOk(await post("compress", [{ name: "scan.pdf", buffer: imageHeavyPdf }], { quality: 0.55 }), "compress scan");
    await expectPdf(out, "compress scan", 2);
    assert.ok(out.length < imageHeavyPdf.length * 0.7, `image-heavy PDF should shrink well (${imageHeavyPdf.length} -> ${out.length})`);
  });

  await run("page numbers", async () => {
    const out = expectOk(
      await post("pagenum", [pdfFile], { position: "bottom-center", format: "page-number-of-total", startNumber: 5, fontSize: 12 }),
      "pagenum"
    );
    await expectPdf(out, "pagenum", 2);
    const text = await extractText(out);
    assert.ok(text.includes("Page 5 of 2"), "first page should show the custom starting number");
    assert.ok(text.includes("Page 6 of 2"), "second page should continue numbering");
  });

  await run("page numbers (selected page)", async () => {
    const out = expectOk(await post("pagenum", [pdfFile], { format: "number", pages: [2], startNumber: 9 }), "pagenum selected");
    await expectPdf(out, "pagenum selected", 2);
    const pages = (await extractText(out)).split("\n");
    assert.ok(!pages[0].includes("10"), "unselected first page should stay untouched");
    assert.ok(pages[1].includes("10"), "selected second page should get its number");
  });

  await run("rotate", async () => {
    const out = expectOk(await post("rotate", [pdfFile], { pages: [1], degrees: 90 }), "rotate");
    const doc = await expectPdf(out, "rotate", 2);
    assert.strictEqual(doc.getPages()[0].getRotation().angle, 90);
  });

  await run("protect", async () => {
    const out = expectOk(await post("protect", [pdfFile], { password: "test1234" }), "protect");
    assert.ok(out.includes("/Encrypt"), "protected PDF should contain an /Encrypt dictionary");
    const text = await extractText(out, "test1234");
    assert.ok(text.includes("Hello PDF Machan"), "the password should unlock the protected PDF");
    let refusedWrongPassword = false;
    try {
      await extractText(out, "wrong-password");
    } catch (error) {
      refusedWrongPassword = true;
    }
    assert.ok(refusedWrongPassword, "a wrong password must not open the PDF");
  });

  await run("editor", async () => {
    const edits = [
      {
        page: 1,
        origX: 50,
        origY: 760,
        fontSize: 18,
        originalText: "Hello PDF Machan (page 1)",
        text: "EDITED TEXT",
        blockWidth: 200,
        customRotation: 0,
        rawTransform: [18, 0, 0, 18, 50, 760],
      },
    ];
    const out = expectOk(await post("editor", [pdfFile], { edits }), "editor");
    await expectPdf(out, "editor", 2);
    const text = await extractText(out);
    assert.ok(text.includes("EDITED TEXT"), "the edited text should be written into the PDF");
  });

  await run("watermark", async () => {
    const layers = [
      { type: "text", page: 1, text: "CONFIDENTIAL", scaleX: 80, scaleY: 120, fontSize: 24, fontColorHex: "#E8483A", fontNameFamily: "Helvetica" },
      { type: "image", page: 1, scaleX: 120, scaleY: 180, width: 140, height: 70, dataUrl: `data:image/png;base64,${png.toString("base64")}` },
    ];
    const out = expectOk(await post("watermark", [pdfFile], { brand: "pdf", layers }), "watermark");
    await expectPdf(out, "watermark", 2);
    const text = await extractText(out);
    assert.ok(text.includes("CONFIDENTIAL"), "the text watermark should be present in the PDF");
  });

  await run("word2pdf", async () => {
    const out = expectOk(await post("word2pdf", [{ name: "sample.docx", buffer: docx }]), "word2pdf");
    const pdfDoc = await expectPdf(out, "word2pdf");
    assert.ok(pdfDoc.getPageCount() >= 1);
    const text = await extractText(out);
    assert.ok(/Hello from Word/i.test(text), "converted PDF should keep the Word text");
  });

  await run("pdf2word", async () => {
    const out = expectOk(await post("pdf2word", [pdfFile]), "pdf2word");
    const zip = await JSZip.loadAsync(out);
    const documentXml = await zip.file("word/document.xml").async("string");
    assert.ok(documentXml.includes("Hello PDF Machan"), "converted DOCX should contain extracted text");
  });

  await run("xls2pdf", async () => {
    const out = expectOk(await post("xls2pdf", [{ name: "sample.xlsx", buffer: xlsx }]), "xls2pdf");
    const doc = await expectPdf(out, "xls2pdf");
    assert.ok(doc.getPageCount() >= 1);
    const text = await extractText(out);
    assert.ok(/Alpha|Name/i.test(text), "converted PDF should keep the spreadsheet cells");
  });

  await run("pdf2xls", async () => {
    const out = expectOk(await post("pdf2xls", [pdfFile]), "pdf2xls");
    const workbook = XLSX.read(out, { type: "buffer" });
    assert.strictEqual(workbook.SheetNames.length, 2, "pdf2xls should create one sheet per page");
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]);
    assert.ok(csv.includes("Hello PDF Machan"), "sheet should contain the extracted PDF text");
  });

  await run("ppt2pdf", async () => {
    const out = expectOk(await post("ppt2pdf", [{ name: "sample.pptx", buffer: pptx }]), "ppt2pdf");
    const doc = await expectPdf(out, "ppt2pdf");
    assert.strictEqual(doc.getPageCount(), 2, "ppt2pdf should create one page per slide");
    const text = await extractText(out);
    assert.ok(/Quarterly Update/i.test(text), "converted PDF should keep the slide text");
  });

  await run("pdf2ppt", async () => {
    const out = expectOk(await post("pdf2ppt", [pdfFile]), "pdf2ppt");
    const zip = await JSZip.loadAsync(out);
    const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    assert.strictEqual(slides.length, 2, "pdf2ppt should create one slide per page");
  });

  await run("pdf2img", async () => {
    const out = expectOk(await post("pdf2img", [pdfFile]), "pdf2img");
    const zip = await JSZip.loadAsync(out);
    const images = Object.keys(zip.files).filter((name) => name.endsWith(".png"));
    assert.strictEqual(images.length, 2);
    const metadata = await sharp(await zip.file(images[0]).async("nodebuffer")).metadata();
    assert.strictEqual(metadata.format, "png");
  });

  await run("img2pdf", async () => {
    const out = expectOk(
      await post("img2pdf", [
        { name: "photo.jpg", buffer: jpeg, type: "image/jpeg" },
        { name: "graphic.png", buffer: png, type: "image/png" },
      ]),
      "img2pdf"
    );
    await expectPdf(out, "img2pdf", 2);
  });

  await run("jpg2png", async () => {
    const out = expectOk(await post("jpg2png", [{ name: "photo.jpg", buffer: jpeg, type: "image/jpeg" }]), "jpg2png");
    const metadata = await sharp(out).metadata();
    assert.strictEqual(metadata.format, "png");
  });

  await run("png2jpg", async () => {
    const out = expectOk(await post("png2jpg", [{ name: "graphic.png", buffer: png, type: "image/png" }]), "png2jpg");
    const metadata = await sharp(out).metadata();
    assert.strictEqual(metadata.format, "jpeg");
  });

  await run("imgcompress", async () => {
    const out = expectOk(
      await post("imgcompress", [{ name: "photo.jpg", buffer: jpeg, type: "image/jpeg" }], { quality: 0.4 }),
      "imgcompress"
    );
    const metadata = await sharp(out).metadata();
    assert.strictEqual(metadata.format, "jpeg");
    assert.ok(out.length < jpeg.length, "compressed image should be smaller");
  });

  await run("error handling", async () => {
    const unknown = await post("not-a-tool", [{ name: "x.pdf", buffer: pdf }]);
    assert.strictEqual(unknown.status, 404);
    const missing = await post("merge", []);
    assert.strictEqual(missing.status, 400);
    const badSplit = await post("split", [pdfFile], { mode: "range", start: 5, end: 1 });
    assert.strictEqual(badSplit.status, 200, "an inverted range is clamped to the start page");
    const badPassword = await post("protect", [pdfFile], { password: "123" });
    assert.strictEqual(badPassword.status, 400);
    const emptyOrder = await post("merge", [pdfFile], { order: [] });
    assert.strictEqual(emptyOrder.status, 400, "merging zero pages should be rejected");
    const noSplitPages = await post("split", [pdfFile], { mode: "select", pages: [] });
    assert.strictEqual(noSplitPages.status, 400, "splitting with no selected pages should be rejected");
  });

  console.log(`\nAll ${results.length} test groups passed.`);
}

main()
  .catch((error) => {
    console.error("\nTest failure:");
    console.error(error);
    process.exitCode = 1;
  })
  .finally(() => {
    if (server) server.close();
  });
