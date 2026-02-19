import express from "express";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { exec } from "child_process";
import { fileURLToPath } from "url";
import os from "os";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 10000;

app.use(express.json({ limit: "10mb" }));

// -----------------------------
// Find LibreOffice (for Docker/Render)
// -----------------------------
const findSoffice = () => {
  const envPath = process.env.SOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;

  const candidates = [
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
  ];

  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }

  return null;
};

// -----------------------------
// Generate PDF from DOCX Template
// -----------------------------
const generatePDF = (data, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);

  if (!fs.existsSync(templatePath)) {
    return res.status(404).send("Template not found.");
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res.status(500).send("LibreOffice not found. Install LibreOffice.");
  }

  const stamp = Date.now();
  const outDir = path.join(os.tmpdir(), "pdf_output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outputDocx = path.join(outDir, `${filenameBase}-${stamp}.docx`);

  try {
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => "",
    });

    doc.render(data);

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    fs.writeFileSync(outputDocx, buf);

    const command = `"${soffice}" --headless --nologo --nolockcheck --norestore --convert-to pdf "${outputDocx}" --outdir "${outDir}"`;

    exec(command, (err) => {
      if (err) {
        console.log("LibreOffice error:", err);
        return res.status(500).send("PDF conversion failed.");
      }

      const pdfFile = outputDocx.replace(".docx", ".pdf");

      if (!fs.existsSync(pdfFile)) {
        return res.status(500).send("PDF file not created.");
      }

      res.download(pdfFile, () => {
        try { fs.unlinkSync(outputDocx); } catch {}
        try { fs.unlinkSync(pdfFile); } catch {}
      });
    });

  } catch (e) {
    console.log("PDF generation error:", e);
    return res.status(500).send("PDF generation failed.");
  }
};

// -----------------------------
// ENDPOINTS
// -----------------------------

// FSIC Certificate
app.post("/generate/fsic", (req, res) => {
  generatePDF(req.body, "fsic-owner.docx", "fsic", res);
});

// IO
app.post("/generate/io", (req, res) => {
  generatePDF(req.body, "officers.docx", "io", res);
});

// Reinspection
app.post("/generate/reinspection", (req, res) => {
  generatePDF(req.body, "reinspection.docx", "reinspection", res);
});

// NFSI
app.post("/generate/nfsi", (req, res) => {
  generatePDF(req.body, "nfsi-form.docx", "nfsi", res);
});

// Health check
app.get("/", (req, res) => {
  res.send("PDF Generator Backend Running");
});

// -----------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log("PDF Backend running on port", PORT);
});
