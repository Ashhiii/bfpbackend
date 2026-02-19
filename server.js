import express from "express";
import cors from "cors";
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

// ✅ CORS FIX (works for localhost + deployed frontends)
app.use(
  cors({
    origin: true,
    credentials: false,
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());

app.use(express.json({ limit: "10mb" }));

// -----------------------------
// Locate LibreOffice
// -----------------------------
const findSoffice = () => {
  const paths = [
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
  ];
  for (const p of paths) {
    if (fs.existsSync(p)) return p;
  }
  return null;
};

// -----------------------------
// Health check (IMPORTANT)
// -----------------------------
app.get("/health", (req, res) => {
  const soffice = findSoffice();
  const templatesDir = path.join(__dirname, "templates");

  const listTemplates = fs.existsSync(templatesDir)
    ? fs.readdirSync(templatesDir)
    : [];

  res.json({
    ok: true,
    sofficeFound: Boolean(soffice),
    sofficePath: soffice,
    templatesDirExists: fs.existsSync(templatesDir),
    templates: listTemplates,
    cwd: process.cwd(),
    dirname: __dirname,
  });
});

// -----------------------------
// Core PDF generator
// -----------------------------
const generatePDF = (data, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);

  if (!fs.existsSync(templatePath)) {
    return res
      .status(404)
      .send(`Template not found: ${templateFile} (path=${templatePath})`);
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res
      .status(500)
      .send("LibreOffice not installed / soffice not found on server.");
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

    const command = `"${soffice}" --headless --nologo --nofirststartwizard --convert-to pdf "${outputDocx}" --outdir "${outDir}"`;

    exec(command, (err, stdout, stderr) => {
      if (err) {
        console.error("LibreOffice error:", err);
        console.error("stdout:", stdout);
        console.error("stderr:", stderr);
        return res
          .status(500)
          .send(`PDF conversion failed. ${stderr || err.message}`);
      }

      const pdfFile = outputDocx.replace(".docx", ".pdf");

      if (!fs.existsSync(pdfFile)) {
        return res.status(500).send("PDF not created after conversion.");
      }

      // ✅ Ensure correct headers
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${filenameBase}-${stamp}.pdf"`
      );

      res.download(pdfFile, () => {
        try {
          fs.unlinkSync(outputDocx);
        } catch {}
        try {
          fs.unlinkSync(pdfFile);
        } catch {}
      });
    });
  } catch (e) {
    console.error("PDF generation error:", e);
    return res.status(500).send(`PDF generation error. ${e.message}`);
  }
};

// -----------------------------
// Routes
// -----------------------------
app.post("/generate/fsic", (req, res) => {
  generatePDF(req.body, "fsic-owner.docx", "fsic", res);
});

app.post("/generate/io", (req, res) => {
  generatePDF(req.body, "officers.docx", "io", res);
});

app.post("/generate/reinspection", (req, res) => {
  generatePDF(req.body, "reinspection.docx", "reinspection", res);
});

app.post("/generate/nfsi", (req, res) => {
  generatePDF(req.body, "nfsi-form.docx", "nfsi", res);
});

app.get("/", (req, res) => res.send("PDF Backend Running ✅"));

app.listen(PORT, "0.0.0.0", () => {
  console.log("Server running on port", PORT);
});
