"use strict";

const fs = require("fs");
const path = require("path");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const XLSX = require("xlsx");
const sharp = require("sharp");
const PptxGenJS = require("pptxgenjs");

const outDir = path.join(__dirname, "fixtures");
fs.mkdirSync(outDir, { recursive: true });
const write = (name, buffer) => fs.writeFileSync(path.join(outDir, name), buffer);

async function pdf(name, label) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([595, 842]);
  page.drawText(`${label} — PDF Machan test`, { x: 50, y: 760, size: 20, font, color: rgb(0.1, 0.1, 0.1) });
  page.drawText("This file is used by the automated browser check.", { x: 50, y: 720, size: 12, font });
  page.drawText("Second line of sample content.", { x: 50, y: 695, size: 12, font });
  write(name, Buffer.from(await doc.save()));
}

(async () => {
  await pdf("sample-a.pdf", "Document A");
  await pdf("sample-b.pdf", "Document B");

  const word = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun({ text: "Word document heading", bold: true })] }),
          new Paragraph({ children: [new TextRun("This paragraph comes from a .docx file.")] }),
        ],
      },
    ],
  });
  write("sample.docx", await Packer.toBuffer(word));

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Item", "Price", "Qty"],
    ["Notebook", "120", "2"],
    ["Pen", "15", "10"],
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  write("sample.xlsx", XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));

  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText("Presentation title", { x: 0.5, y: 0.4, w: 8, h: 1, fontSize: 24, bold: true });
  slide.addText("Bullet point from PowerPoint", { x: 0.5, y: 1.6, w: 8, h: 1, fontSize: 16 });
  write("sample.pptx", Buffer.from(await pptx.write({ outputType: "nodebuffer" })));

  write(
    "sample.jpg",
    await sharp({ create: { width: 400, height: 260, channels: 3, background: "#e8483a" } }).jpeg({ quality: 90 }).toBuffer()
  );
  write(
    "sample.png",
    await sharp({ create: { width: 400, height: 260, channels: 4, background: "#2a7fe2" } }).png().toBuffer()
  );

  console.log("Fixtures written to", outDir);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
