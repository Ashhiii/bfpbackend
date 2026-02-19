import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import os from "os";
import { exec } from "child_process";
import { fileURLToPath } from "url";

import admin from "firebase-admin";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

const PORT = Number(process.env.PORT) || 5000;

// -------------------------
// ✅ Firebase Admin
// -------------------------
// Option A (recommended on Render): use env var FIREBASE_SERVICE_ACCOUNT_JSON
// Put the full json string in env (Render > Environment))
if (!admin.apps.length) {
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT_JSON;
  if (!raw) {
    console.warn("⚠️ Missing FIREBASE_SERVICE_ACCOUNT_JSON env var");
  } else {
    admin.initializeApp({
      credential: admin.credential.cert(JSON.parse(raw)),
    });
  }
}
const db = admin.firestore();

// -------------------------
// ✅ Templates folder
// -------------------------
const TPL_DIR = path.join(__dirname, "templates"); // make sure exists
const TPL = {
  io: path.join(TPL_DIR, "io.docx"),
  reinspection: path.join(TPL_DIR, "reinspection.docx"),
  nfsi: path.join(TPL_DIR, "nfsi.docx"),
};

// -------------------------
// Helpers
// -------------------------
const safe = (s) => String(s || "").replace(/[^\w\-]+/g, "_").slice(0, 80);

const readTemplate = (filePath) => {
  if (!fs.existsSync(filePath)) throw new Error(`Template not found: ${filePath}`);
  return fs.readFileSync(filePath, "binary");
};

const renderDocx = (templatePath, data) => {
  const content = readTemplate(templatePath);
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // ✅ Fill template variables
  doc.setData(data);

  try {
    doc.render();
  } catch (e) {
    console.error("DOCX render error:", e);
    throw new Error("DOCX template render failed. Check placeholders.");
  }

  return doc.getZip().generate({ type: "nodebuffer" });
};

const convertDocxToPdf = async (docxBuffer, outName = "output") => {
  const tmp = os.tmpdir();
  const inPath = path.join(tmp, `${outName}.docx`);
  const outDir = tmp;

  fs.writeFileSync(inPath, docxBuffer);

  // ✅ LibreOffice conversion
  // Windows: soffice.exe available if installed
  // Linux/Docker: libreoffice installed
  const cmd = `soffice --headless --nologo --nofirststartwizard --convert-to pdf --outdir "${outDir}" "${inPath}"`;

  await new Promise((resolve, reject) => {
    exec(cmd, (err, stdout, stderr) => {
      if (err) {
        console.error("LibreOffice error:", stderr || stdout || err);
        return reject(new Error("PDF conversion failed (LibreOffice)."));
      }
      resolve();
    });
  });

  const pdfPath = path.join(outDir, `${outName}.pdf`);
  if (!fs.existsSync(pdfPath)) throw new Error("PDF not created.");

  const pdfBuffer = fs.readFileSync(pdfPath);

  // cleanup
  try {
    fs.unlinkSync(inPath);
    fs.unlinkSync(pdfPath);
  } catch {}

  return pdfBuffer;
};

const getDocumentData = async (id) => {
  const snap = await db.collection("documents").doc(String(id)).get();
  if (!snap.exists) throw new Error("Document not found in Firestore.");
  const data = snap.data() || {};

  // ✅ Return data object for docx placeholders
  // Make sure placeholders in docx match these keys!
  return {
    id,
    fsicAppNo: data.fsicAppNo || "",
    ownerName: data.ownerName || "",
    establishmentName: data.establishmentName || "",
    businessAddress: data.businessAddress || "",
    contactNumber: data.contactNumber || "",
    ioNumber: data.ioNumber || "",
    ioDate: data.ioDate || "",
    nfsiNumber: data.nfsiNumber || "",
    nfsiDate: data.nfsiDate || "",
    inspectors: data.inspectors || "",
    teamLeader: data.teamLeader || "",
    chiefName: data.chiefName || "",
    marshalName: data.marshalName || "",
  };
};

const sendPdf = (res, pdfBuffer, filename) => {
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(pdfBuffer);
};

// -------------------------
// ✅ PDF Endpoints (match your frontend URLs)
// -------------------------
app.get("/documents/:id/io/pdf", async (req, res) => {
  try {
    const { id } = req.params;
    const data = await getDocumentData(id);
    const baseName = `IO_${safe(data.fsicAppNo || id)}`;

    const docx = renderDocx(TPL.io, data);
    const pdf = await convertDocxToPdf(docx, baseName);

    return sendPdf(res, pdf, `${baseName}.pdf`);
  } catch (e) {
    console.error(e);
    return res.status(500).send(e.message || "Failed to generate IO PDF");
  }
});

app.get("/documents/:id/reinspection/pdf", async (req, res) => {
  try {
    const { id } = req.params;
    const data = await getDocumentData(id);
    const baseName = `REINSPECTION_${safe(data.fsicAppNo || id)}`;

    const docx = renderDocx(TPL.reinspection, data);
    const pdf = await convertDocxToPdf(docx, baseName);

    return sendPdf(res, pdf, `${baseName}.pdf`);
  } catch (e) {
    console.error(e);
    return res.status(500).send(e.message || "Failed to generate Reinspection PDF");
  }
});

app.get("/documents/:id/nfsi/pdf", async (req, res) => {
  try {
    const { id } = req.params;
    const data = await getDocumentData(id);
    const baseName = `NFSI_${safe(data.fsicAppNo || id)}`;

    const docx = renderDocx(TPL.nfsi, data);
    const pdf = await convertDocxToPdf(docx, baseName);

    return sendPdf(res, pdf, `${baseName}.pdf`);
  } catch (e) {
    console.error(e);
    return res.status(500).send(e.message || "Failed to generate NFSI PDF");
  }
});

app.get("/", (_, res) => res.send("OK"));

app.listen(PORT, () => console.log("Server running on", PORT));
