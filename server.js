// server.js — BFP PDF Backend (Firestore Option B + LibreOffice)
// ESM: "type": "module"

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { exec } from "child_process";
import { fileURLToPath } from "url";
import os from "os";

// ✅ Firestore Admin
import admin from "firebase-admin";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 10000;

// -----------------------------
// CORS (safe for dev + prod)
// -----------------------------
app.use(
  cors({
    origin: true, // allows any origin (ok for internal system). You can restrict later.
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());

// -----------------------------
// Body parsing
// -----------------------------
app.use(express.json({ limit: "10mb" }));

// -----------------------------
// Firebase Admin init
// -----------------------------
function requireEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

function initFirebaseAdmin() {
  if (admin.apps.length) return;

  // Option 1: Provide as separate env vars
  const projectId = requireEnv("bfp-system-dd589");
  const clientEmail = requireEnv("firebase-adminsdk-fbsvc@bfp-system-dd589.iam.gserviceaccount.com");
  let privateKey = requireEnv("-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC0E3/dJmQxKzFe\nPkgIsjYUTn9vk3Y1aIfnY64qrUDSCZxPx1WrAJw1wG1YF7wE4OqD7a4b/6GMyXn+\ndQmyVOP5JCRPsncRyV3ytKrzImIhU3QkZiNt+rSS0iy1dU4mZwBDw+X0PEJjTOvo\nSL/exN39TdFL63TC4mjERkwHv7tq2ChbLepWH0lSnbSm7CrpWNLWMp7DSlqsUxE4\n8AdcVKZLbQNTayUYAv3ESazIPnqFmeD8AIZ0dxU8GvznxVsKlg/BUy8B0WDZMIsa\nEksGYb78a+eo/feMnsTQPmVe38Rm8oeHg7KVfQ5c6ZR/t1fCFUx1590uwG/RLz0h\nXS1NEB/5AgMBAAECggEAK5A+EhOcEHMJDeoWRhhK0e3NkYIEilYf7p/kPgnM0SMn\nD/vCo0xR/gTvzuxq/WkmAsWIpdhZTjRdkWvh7STxQ8/1AtzDRmfHFRSU+J28TKCi\nS73z20LX1V8IFtvqjIrFex9UrZm4W8p7laWn2S4aIBHnzKyJAi6Att7kvMkq/2Uk\nj5855AEz2J5Kkg3cC3WUVu71cEYsIGYcHhRakbKHUVEDvS5Q4KxfQtcf7hEDKCsS\n7oFZoKsPrinywAQ9Zc7mQX9gM71KL5yepicvgHf6qQNQAmBJzwolBFfz5qI0IBPJ\nsRA34fJ49dxwRRym5f8vqTnrpzkefs3MHoe8Ay4wqQKBgQDi1NL9OssMRS5EdOfH\nzloD5yKS7SZbIF2CGqR1poCrVSjmFLvHuwmsTG6+cV5u1cPJCkuFLLGefct3HL2P\ni7si2eUPdM4EG/ovGp05yh/qaAbNXfgg535glyOplFyRKNx27AfloTUr3fwVI9xG\n0QSC2vaFR9Ys3Oc+QcABaBg7ZQKBgQDLO4Sd+8jiuM65y477XW3zQ5Yn5VtytG1s\nwTu+LYXgudK20HMeso5g+UpCUvPSaT8mu79zKeeb2F6a6/ptPNNyyYQam6iPzhjU\nppRn5JimXAbab/ZDCI87BDtrO60ITbWsB/eyNl9EUOceDJFlZTE3RJSAQjqnMezU\n8WTb0tkrBQKBgEj8D5NZOoINNHDkaLuags4krbQ2yBOkG1JEfkw4scfvE8TMQQtn\n/BnLUgrG8YqaD2/9dwU3t0L22OPPhJLOc/0nvRM3Ex6xXdOQLaytg41pwV6M8Myt\nj7VGjDwBN7OhQSamlXfmYfRmbbjjm/nRQTFVgOKrlZdD0MfwYfZR5xIdAoGBAI76\nuYtYJ3VyZWYLUvV/Pqplg09gKh0h2NYpgDmE12pJ8XfSYRtYrZiF6fuNVqJSjy5V\n6hRAiAR5jd1aZZ3SrERWh80cc9NwQMHizAfnqIlptxyYpl10uSfhC1ie7FYqLcTw\n+g9z90Qd0uuuFeCK97wVUJCHh7I3CpKpnbaVuVltAoGBAKl8cJMyChBCwb3yW/1d\nySq+XeUO5bYWVNnkhBgYMDWdR510O+VAFzK1bggT/n234o7PHEBd1yvsVrWMuHB1\nofh1Pokh4YoYJIRtPcxpScxGDJ9czfSn7x1Rae2cBPPDS7/waBXKt9welgtakgJy\nD7O5fvYHdbYCOzeLHxtkrYG1\n-----END PRIVATE KEY-----\n");

  // Render env usually stores \n as literal text → convert to real newlines
  privateKey = privateKey.replace(/\\n/g, "\n");

  admin.initializeApp({
    credential: admin.credential.cert({
      projectId,
      clientEmail,
      privateKey,
    }),
  });
}

initFirebaseAdmin();
const db = admin.firestore();

// -----------------------------
// Helpers
// -----------------------------
const normalize = (v) => String(v ?? "").trim();

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

const templatesDir = path.join(__dirname, "templates");
const mustExistTemplate = (filename) => {
  const p = path.join(templatesDir, filename);
  if (!fs.existsSync(p)) {
    throw new Error(`Template not found: ${filename} (expected at ${p})`);
  }
  return p;
};

// Firestore fetch
async function getDocData(collectionName, id) {
  const snap = await db.collection(collectionName).doc(String(id)).get();
  if (!snap.exists) return null;

  // include id too, para sure
  return { id: snap.id, ...snap.data() };
}

// Build template variables (matches your docx placeholders)
function buildTemplateView(record = {}) {
  return {
    // FSIC
    FSIC_NUMBER: record.FSIC_NUMBER || record.FSIC_APP_NO || record.fsicAppNo || "",
    DATE_INSPECTED: record.DATE_INSPECTED || record.dateInspected || "",
    NAME_OF_ESTABLISHMENT:
      record.NAME_OF_ESTABLISHMENT || record.ESTABLISHMENT_NAME || record.establishmentName || "",
    NAME_OF_OWNER: record.NAME_OF_OWNER || record.OWNERS_NAME || record.ownerName || "",
    ADDRESS: record.ADDRESS || record.BUSSINESS_ADDRESS || record.businessAddress || "",
    FLOOR_AREA: record.FLOOR_AREA || record.floorArea || "",
    BLDG_DESCRIPTION: record.BLDG_DESCRIPTION || record.BUILDING_DESC || record.buildingDesc || "",
    FSIC_VALIDITY: record.FSIC_VALIDITY || record.fsicValidity || "",
    OR_NUMBER: record.OR_NUMBER || record.orNumber || "",
    OR_DATE: record.OR_DATE || record.orDate || "",
    OR_AMOUNT: record.OR_AMOUNT || record.orAmount || "",

    // IO
    IO_NUMBER: record.IO_NUMBER || record.ioNumber || "",
    IO_DATE: record.IO_DATE || record.ioDate || "",
    TAXPAYER: record.TAXPAYER || record.OWNERS_NAME || record.ownerName || "",
    TRADE_NAME: record.TRADE_NAME || record.ESTABLISHMENT_NAME || record.establishmentName || "",
    CONTACT_: record.CONTACT_ || record.CONTACT_NUMBER || record.contactNumber || "",

    // NFSI
    NFSI_NUMBER: record.NFSI_NUMBER || record.nfsiNumber || "",
    NFSI_DATE: record.NFSI_DATE || record.nfsiDate || "",

    // Signatures / Team
    OWNER: record.OWNER || record.OWNERS_NAME || record.ownerName || "",
    INSPECTORS: record.INSPECTORS || record.inspectors || "",
    TEAM_LEADER: record.TEAM_LEADER || record.teamLeader || "",

    DATE: new Date().toLocaleDateString(),

    CHIEF: record.CHIEF || record.chiefName || "",
    MARSHAL: record.MARSHAL || record.marshalName || "",
  };
}

// PDF generator
async function generatePDF(data, templateFile, filenameBase, res) {
  try {
    const templatePath = mustExistTemplate(templateFile);

    const soffice = findSoffice();
    if (!soffice) {
      return res
        .status(500)
        .send("LibreOffice not found. Install LibreOffice in Docker or set SOFFICE_PATH.");
    }

    const stamp = `${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
    const outDir = path.join(os.tmpdir(), "bfp_pdf_out");
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const outputDocx = path.join(outDir, `${filenameBase}-${stamp}.docx`);

    // Fill template
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => "",
    });

    doc.render(buildTemplateView(data));
    const buf = doc.getZip().generate({ type: "nodebuffer" });
    fs.writeFileSync(outputDocx, buf);

    // Convert to PDF
    const command = `"${soffice}" --headless --nologo --nolockcheck --norestore --convert-to pdf "${outputDocx}" --outdir "${outDir}"`;

    exec(command, (err, stdout, stderr) => {
      if (err) {
        console.log("LibreOffice ERROR:", err);
        console.log("stdout:", stdout);
        console.log("stderr:", stderr);

        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        return res.status(500).send("PDF conversion failed. Check backend logs.");
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");
      if (!fs.existsSync(expectedPdf)) {
        console.log("PDF not found after conversion.");
        console.log("expectedPdf:", expectedPdf);

        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        return res.status(500).send("PDF file not produced. Check LibreOffice conversion.");
      }

      // Download
      res.download(expectedPdf, `${filenameBase}.pdf`, () => {
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        try { if (fs.existsSync(expectedPdf)) fs.unlinkSync(expectedPdf); } catch {}
      });
    });
  } catch (e) {
    console.log("PDF generation error:", e);
    return res.status(500).send(`PDF generation error: ${e.message}`);
  }
}

// -----------------------------
// Routes
// -----------------------------
app.get("/", (req, res) => res.send("✅ BFP PDF Backend Running (Firestore Option B)"));

// FSIC Certificate from records
// /records/:id/certificate/owner/pdf
// /records/:id/certificate/bfp/pdf
app.get("/records/:id/certificate/:type/pdf", async (req, res) => {
  const id = req.params.id;
  const type = normalize(req.params.type).toLowerCase();

  const record = await getDocData("records", id);
  if (!record) return res.status(404).send("Record not found in Firestore (records).");

  const templateFile =
    type === "owner" ? "fsic-owner.docx" :
    type === "bfp" ? "fsic-bfp.docx" :
    "";

  if (!templateFile) return res.status(400).send("Invalid certificate type. Use owner or bfp.");

  return generatePDF(record, templateFile, `fsic-${type}-${id}`, res);
});

// IO / REINSPECTION / NFSI for records
// /records/:id/io/pdf
// /records/:id/reinspection/pdf
// /records/:id/nfsi/pdf
app.get("/records/:id/:docType/pdf", async (req, res) => {
  const id = req.params.id;
  const docType = normalize(req.params.docType).toLowerCase();

  const record = await getDocData("records", id);
  if (!record) return res.status(404).send("Record not found in Firestore (records).");

  let templateFile = "";
  if (docType === "io") templateFile = "officers.docx";
  else if (docType === "reinspection") templateFile = "reinspection.docx";
  else if (docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid docType. Use io, reinspection, nfsi.");

  return generatePDF(record, templateFile, `${docType}-${id}`, res);
});

// DOCUMENTS PDF generation
// /documents/:id/io/pdf
// /documents/:id/reinspection/pdf
// /documents/:id/nfsi/pdf
app.get("/documents/:id/:docType/pdf", async (req, res) => {
  const id = req.params.id;
  const docType = normalize(req.params.docType).toLowerCase();

  const doc = await getDocData("documents", id);
  if (!doc) return res.status(404).send("Document not found in Firestore (documents).");

  let templateFile = "";
  if (docType === "io") templateFile = "officers.docx";
  else if (docType === "reinspection") templateFile = "reinspection.docx";
  else if (docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid docType. Use io, reinspection, nfsi.");

  return generatePDF(doc, templateFile, `doc-${docType}-${id}`, res);
});

// -----------------------------
// Start
// -----------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
  console.log("Firestore Project:", process.env.VITE_FB_PROJECT_ID || "(missing)");
});
