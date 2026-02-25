// server.js — BFP System Backend (FULL FIX + FIRESTORE OPTION B)
// ✅ CORS + OPTIONS
// ✅ /health
// ✅ PDF generation (LibreOffice) + clear errors
// ✅ Firestore lookup for Records + Archive + Documents (NO MORE records.json mismatch)
// ✅ Excel import endpoints still kept (optional) BUT PDF uses Firestore

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { exec } from "child_process";
import { fileURLToPath } from "url";
import os from "os";
import multer from "multer";
import xlsx from "xlsx";
import crypto from "crypto";
import admin from "firebase-admin";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 10000;

// -----------------------------
// MIDDLEWARE
// -----------------------------
app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "DELETE", "PUT", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());
app.use(express.json({ limit: "10mb" }));

// multer memory upload
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
});

// -----------------------------
// FIREBASE ADMIN (Option B)
// -----------------------------
// Put these in backend .env or Render env vars:
// FIREBASE_PROJECT_ID=xxx
// FIREBASE_CLIENT_EMAIL=xxx@xxx.iam.gserviceaccount.com
// FIREBASE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
if (!admin.apps.length) {
  const projectId = process.env.FIREBASE_PROJECT_ID;
  const clientEmail = process.env.FIREBASE_CLIENT_EMAIL;
  const privateKey = (process.env.FIREBASE_PRIVATE_KEY || "").replace(/\\n/g, "\n");

  if (!projectId || !clientEmail || !privateKey) {
    console.warn("⚠️ Firebase Admin env vars missing. Firestore lookups will fail.");
  } else {
    admin.initializeApp({
      credential: admin.credential.cert({
        projectId,
        clientEmail,
        privateKey,
      }),
    });
  }
}
const fdb = admin.apps.length ? admin.firestore() : null;

// -----------------------------
// OPTIONAL JSON FILES (kept for imports/export if you still want local fallback)
// -----------------------------
const DATA_FILE = path.join(__dirname, "records.json");
const ARCHIVE_FILE = path.join(__dirname, "archive.json");
const DOCUMENTS_FILE = path.join(__dirname, "documents.json");
const HISTORY_FILE = path.join(__dirname, "history.json");

const ensureFile = (file, defaultData) => {
  if (!fs.existsSync(file)) fs.writeFileSync(file, JSON.stringify(defaultData, null, 2));
};
ensureFile(DATA_FILE, []);
ensureFile(ARCHIVE_FILE, {});
ensureFile(DOCUMENTS_FILE, []);
ensureFile(HISTORY_FILE, []);

const readJSON = (file) => JSON.parse(fs.readFileSync(file, "utf-8"));
const writeJSON = (file, data) => fs.writeFileSync(file, JSON.stringify(data, null, 2));

// -----------------------------
// HELPERS
// -----------------------------
const normalize = (v) => String(v ?? "").trim();

const ensureEntityKey = (r) => {
  if (!r) return r;
  if (r.entityKey) return r;
  const fsic = normalize(r.fsicAppNo || r.FSIC_APP_NO || r.FSIC_NUMBER);
  return { ...r, entityKey: fsic ? `fsic:${fsic}` : `rec:${r.id || Date.now()}` };
};

const normalizeEntityKey = (entityKey) => normalize(entityKey);

const pickAllowedRecordFields = (obj = {}) => ({
  appno: obj.appno ?? obj.APPLICATION_NO ?? "",
  fsicAppNo: obj.fsicAppNo ?? obj.FSIC_APP_NO ?? obj.FSIC_NUMBER ?? "",
  natureOfInspection: obj.natureOfInspection ?? obj.NATURE_OF_INSPECTION ?? "",
  ownerName: obj.ownerName ?? obj.OWNERS_NAME ?? "",
  establishmentName: obj.establishmentName ?? obj.ESTABLISHMENT_NAME ?? "",
  businessAddress: obj.businessAddress ?? obj.BUSSINESS_ADDRESS ?? obj.ADDRESS ?? "",
  contactNumber: obj.contactNumber ?? obj.CONTACT_NUMBER ?? "",
  dateInspected: obj.dateInspected ?? obj.DATE_INSPECTED ?? "",

  ioNumber: obj.ioNumber ?? obj.IO_NUMBER ?? "",
  ioDate: obj.ioDate ?? obj.IO_DATE ?? "",

  nfsiNumber: obj.nfsiNumber ?? obj.NFSI_NUMBER ?? "",
  nfsiDate: obj.nfsiDate ?? obj.NFSI_DATE ?? "",

  ntcNumber: obj.ntcNumber ?? obj.NTC_NUMBER ?? "",
  ntcDate: obj.ntcDate ?? obj.NTC_DATE ?? "",

  fsicValidity: obj.fsicValidity ?? obj.FSIC_VALIDITY ?? "",
  defects: obj.defects ?? obj.DEFECTS ?? "",
  inspectors: obj.inspectors ?? obj.INSPECTORS ?? "",
  occupancyType: obj.occupancyType ?? obj.OCCUPANCY_TYPE ?? "",
  buildingDesc: obj.buildingDesc ?? obj.BUILDING_DESC ?? obj.BLDG_DESCRIPTION ?? "",
  floorArea: obj.floorArea ?? obj.FLOOR_AREA ?? "",
  buildingHeight: obj.buildingHeight ?? obj.BUILDING_HEIGHT ?? "",
  storeyCount: obj.storeyCount ?? obj.STOREY_COUNT ?? "",
  highRise: obj.highRise ?? obj.HIGH_RISE ?? "",
  fsmr: obj.fsmr ?? obj.FSMR ?? "",
  remarks: obj.remarks ?? obj.REMARKS ?? "",

  orNumber: obj.orNumber ?? obj.OR_NUMBER ?? "",
  orAmount: obj.orAmount ?? obj.OR_AMOUNT ?? "",
  orDate: obj.orDate ?? obj.OR_DATE ?? "",

  chiefName: obj.chiefName ?? obj.CHIEF ?? "",
  marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
});

const pickAllowedDocumentFields = (obj = {}) => ({
  fsicAppNo: obj.fsicAppNo ?? obj.FSIC_APP_NO ?? obj.FSIC_NUMBER ?? "",
  ownerName: obj.ownerName ?? obj.OWNERS_NAME ?? "",
  establishmentName: obj.establishmentName ?? obj.ESTABLISHMENT_NAME ?? "",
  businessAddress: obj.businessAddress ?? obj.BUSSINESS_ADDRESS ?? obj.ADDRESS ?? "",
  contactNumber: obj.contactNumber ?? obj.CONTACT_NUMBER ?? "",

  ioNumber: obj.ioNumber ?? obj.IO_NUMBER ?? "",
  ioDate: obj.ioDate ?? obj.IO_DATE ?? "",

  ntcNumber: obj.ntcNumber ?? obj.NTC_NUMBER ?? "",
  ntcDate: obj.ntcDate ?? obj.NTC_DATE ?? "",

  nfsiNumber: obj.nfsiNumber ?? obj.NFSI_NUMBER ?? "",
  nfsiDate: obj.nfsiDate ?? obj.NFSI_DATE ?? "",

  inspectors: obj.inspectors ?? obj.INSPECTORS ?? "",
  
  teamLeader: obj.teamLeader ?? obj.TEAM_LEADER ?? "",
  teamLeaderSerial: obj.teamLeaderSerial ?? obj.TEAM_LEADER_SERIAL ?? "",

  inspector1: obj.inspector1 ?? obj.INSPECTOR_1 ?? "",
  inspector1Serial: obj.inspector1Serial ?? obj.INSPECTOR_1_SERIAL ?? "",

  inspector2: obj.inspector2 ?? obj.INSPECTOR_2 ?? "",
  inspector2Serial: obj.inspector2Serial ?? obj.INSPECTOR_2_SERIAL ?? "",

  inspector3: obj.inspector3 ?? obj.INSPECTOR_3 ?? "",
  inspector3Serial: obj.inspector3Serial ?? obj.INSPECTOR_3_SERIAL ?? "",

  chiefName: obj.chiefName ?? obj.CHIEF ?? "",
  marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
});

// -----------------------------
// 🔥 FIRESTORE LOOKUPS (THIS FIXES "Record not found")
// Assumed structure:
// - current records: collection "records" (doc id = record id)
// - documents: collection "documents" (doc id = document id)
// - archive: collection "archive" -> doc {YYYY-MM} -> subcollection "records" -> doc id = record id
// If lahi imo structure, ingna ko para i-adjust.
// -----------------------------
const findRecordById = async (id) => {
  if (!fdb) return null;

  // 1) current records
  const snap = await fdb.collection("records").doc(String(id)).get();
  if (snap.exists) return ensureEntityKey({ id: snap.id, ...snap.data() });

  // 2) archive/{YYYY-MM}/records/{id}
  const monthsSnap = await fdb.collection("archive").get();
  for (const m of monthsSnap.docs) {
    const recSnap = await fdb
      .collection("archive")
      .doc(m.id)
      .collection("records")
      .doc(String(id))
      .get();

    if (recSnap.exists) return ensureEntityKey({ id: recSnap.id, ...recSnap.data() });
  }

  return null;
};

const findDocumentById = async (id) => {
  if (!fdb) return null;
  const snap = await fdb.collection("documents").doc(String(id)).get();
  if (snap.exists) return { id: snap.id, ...snap.data() };
  return null;
};

// -----------------------------
// PDF (LibreOffice)
// -----------------------------
const findSoffice = () => {
  const envPath = process.env.SOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;

  const candidates = [
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
  ];
  for (const p of candidates) if (fs.existsSync(p)) return p;
  return null;
};

const generatePDF = (record, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);

  if (!fs.existsSync(templatePath)) {
    return res.status(404).send(`Template not found: ${templateFile} (path=${templatePath})`);
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res.status(500).send("LibreOffice not found (soffice). Use Docker install libreoffice.");
  }

  const stamp = `${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
  const outDir = path.join(os.tmpdir(), "bfp_pdf_out");
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

    const view = {
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

      IO_NUMBER: record.IO_NUMBER || record.ioNumber || "",
      IO_DATE: record.IO_DATE || record.ioDate || "",
      TAXPAYER: record.TAXPAYER || record.OWNERS_NAME || record.ownerName || "",
      TRADE_NAME: record.TRADE_NAME || record.ESTABLISHMENT_NAME || record.establishmentName || "",
      CONTACT_: record.CONTACT_ || record.CONTACT_NUMBER || record.contactNumber || "",

      NFSI_NUMBER: record.NFSI_NUMBER || record.nfsiNumber || "",
      NFSI_DATE: record.NFSI_DATE || record.nfsiDate || "",

      OWNER: record.OWNER || record.OWNERS_NAME || record.ownerName || "",
      TEAM_LEADER: record.teamLeader || record.TEAM_LEADER || "",
      TEAM_LEADER_SERIAL: record.teamLeaderSerial || record.TEAM_LEADER_SERIAL || "",
      INSPECTORS: record.INSPECTORS || record.inspectors || "",

      INSPECTOR_1: record.inspector1 || record.INSPECTOR_1 || "",
      INSPECTOR_1_SERIAL: record.inspector1Serial || record.INSPECTOR_1_SERIAL || "",

      INSPECTOR_2: record.inspector2 || record.INSPECTOR_2 || "",
      INSPECTOR_2_SERIAL: record.inspector2Serial || record.INSPECTOR_2_SERIAL || "",

      INSPECTOR_3: record.inspector3 || record.INSPECTOR_3 || "",
      INSPECTOR_3_SERIAL: record.inspector3Serial || record.INSPECTOR_3_SERIAL || "",

      NTC_NUMBER: record.ntcNumber || record.NTC_NUMBER || "",
      NTC_DATE: record.ntcDate || record.NTC_DATE || "",

      DATE: new Date().toLocaleDateString(),
      CHIEF: record.CHIEF || record.chiefName || "",
      MARSHAL: record.MARSHAL || record.marshalName || "",
    };

    doc.render(view);

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    fs.writeFileSync(outputDocx, buf);

    const command = `"${soffice}" --headless --nologo --nolockcheck --norestore --convert-to pdf "${outputDocx}" --outdir "${outDir}"`;

    exec(command, (err, stdout, stderr) => {
      if (err) {
        console.log("LibreOffice ERROR:", err);
        console.log("stdout:", stdout);
        console.log("stderr:", stderr);
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        return res.status(500).send(`PDF conversion failed. ${String(stderr || err?.message || err)}`);
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");
      if (!fs.existsSync(expectedPdf)) {
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        return res.status(500).send("PDF file not produced after conversion.");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `attachment; filename="${filenameBase}.pdf"`);

      res.download(expectedPdf, () => {
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        try {
          if (fs.existsSync(expectedPdf)) fs.unlinkSync(expectedPdf);
        } catch {}
      });
    });
  } catch (e) {
    console.log("PDF generation failed:", e);
    try {
      if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
    } catch {}
    return res.status(500).send(`PDF generation failed (templater). ${e.message}`);
  }
};

// ✅ health check
app.get("/health", async (req, res) => {
  const soffice = findSoffice();
  const templatesDir = path.join(__dirname, "templates");

  let firestoreOk = false;
  try {
    if (fdb) {
      await fdb.collection("_health").doc("ping").get();
      firestoreOk = true;
    }
  } catch {
    firestoreOk = false;
  }

  res.json({
    ok: true,
    firestoreConnected: Boolean(fdb) && firestoreOk,
    sofficeFound: Boolean(soffice),
    sofficePath: soffice,
    templatesDirExists: fs.existsSync(templatesDir),
    templates: fs.existsSync(templatesDir) ? fs.readdirSync(templatesDir) : [],
  });
});

// -----------------------------
// EXCEL IMPORT HELPERS (optional, unchanged)
// -----------------------------
const pickAny = (obj, keys = []) => {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return "";
};

const normHeader = (s) =>
  String(s ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/_/g, "")
    .replace(/-/g, "");

const makeId = () => {
  try {
    return crypto.randomUUID();
  } catch {
    return String(Date.now() + Math.floor(Math.random() * 100000));
  }
};

const excelDateToISO = (v) => {
  if (typeof v === "string") return v;
  if (typeof v === "number") {
    const dt = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(dt.getTime())) {
      const y = dt.getFullYear();
      const m = String(dt.getMonth() + 1).padStart(2, "0");
      const d = String(dt.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
  }
  return String(v ?? "");
};

const mapExcelRowToRecord = (row = {}) => {
  const headerMap = {};
  for (const k of Object.keys(row)) headerMap[normHeader(k)] = row[k];
  const get = (...variants) => pickAny(headerMap, variants.map(normHeader));

  const rec = {
    id: makeId(),
    createdAt: new Date().toISOString(),

    appno: String(get("appno", "applicationno", "application#", "applicationnumber") || ""),
    fsicAppNo: String(get("fsicappno", "fsicno", "fsicnumber", "fsicapp#", "fsicapp") || ""),
    natureOfInspection: String(get("natureofinspection", "inspection", "nature") || ""),
    ownerName: String(get("owner", "ownername", "ownersname", "taxpayer") || ""),
    establishmentName: String(
      get("establishment", "establishmentname", "tradename", "nameofestablishment") || ""
    ),
    businessAddress: String(get("businessaddress", "address", "bussinessaddress") || ""),
    contactNumber: String(get("contactnumber", "contact", "mobile") || ""),
    dateInspected: excelDateToISO(get("dateinspected", "date") || ""),

    ioNumber: String(get("ionumber", "io#", "io") || ""),
    ioDate: excelDateToISO(get("iodate") || ""),

    nfsiNumber: String(get("nfsinumber", "nfsi#", "nfsi") || ""),
    nfsiDate: excelDateToISO(get("nfsidate") || ""),

    fsicValidity: String(get("fsicvalidity", "validity") || ""),
    defects: String(get("defects", "violations") || ""),
    inspectors: String(get("inspectors", "inspector") || ""),

    chiefName: String(get("chiefname", "chief") || ""),
    marshalName: String(get("marshalname", "marshal") || ""),
  };

  rec.fsicAppNo = String(rec.fsicAppNo || "").toUpperCase().trim();
  rec.ownerName = String(rec.ownerName || "").toUpperCase().trim();
  rec.establishmentName = String(rec.establishmentName || "").toUpperCase().trim();
  rec.businessAddress = String(rec.businessAddress || "").toUpperCase().trim();

  return ensureEntityKey(rec);
};

const readExcelRows = (req) => {
  if (!req.file) return { error: "No file uploaded." };

  const filename = String(req.file.originalname || "").toLowerCase();
  if (!filename.endsWith(".xlsx") && !filename.endsWith(".xls")) {
    return { error: "Invalid file. Upload .xlsx or .xls only." };
  }

  const wb = xlsx.read(req.file.buffer, { type: "buffer" });
  const sheetName = wb.SheetNames?.[0];
  if (!sheetName) return { error: "Excel file has no sheets." };

  const ws = wb.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(ws, { defval: "" });

  if (!rows.length) return { error: "Excel sheet is empty." };
  return { rows, sheetName };
};

// -----------------------------
// (OPTIONAL) EXCEL IMPORT ENDPOINTS (still writes to JSON)
// If you want Firestore import instead, tell me and I’ll convert these to Firestore writes.
// -----------------------------
app.post("/import/records", upload.single("file"), (req, res) => {
  try {
    const { rows, sheetName, error } = readExcelRows(req);
    if (error) return res.status(400).json({ success: false, message: error });

    const current = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    const mapped = rows.map(mapExcelRowToRecord);
    const toAdd = mapped.filter((r) => normalize(r.fsicAppNo) && normalize(r.ownerName));

    writeJSON(DATA_FILE, [...current, ...toAdd]);

    res.json({
      success: true,
      imported: toAdd.length,
      skipped: mapped.length - toAdd.length,
      sheet: sheetName,
    });
  } catch (e) {
    console.error("POST /import/records error:", e);
    res.status(500).json({ success: false, message: "Failed to import records Excel." });
  }
});

// -----------------------------
// PIN auth
// -----------------------------
const CORRECT_PIN = String(process.env.PIN || "1234").trim();
app.post("/auth/pin", (req, res) => {
  const pin = String(req.body?.pin || "").trim();
  if (!pin) return res.status(400).json({ ok: false, message: "Missing PIN" });
  if (pin === CORRECT_PIN) return res.json({ ok: true });
  return res.status(401).json({ ok: false, message: "Incorrect PIN" });
});

// -----------------------------
// PDF ROUTES (Firestore based)
// -----------------------------
// FSIC Certificate
app.get("/records/:id/certificate/:type/pdf", async (req, res) => {
  const record = await findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const type = String(req.params.type || "").toLowerCase();
  const templateFile = type === "owner" ? "fsic-owner.docx" : "fsic-bfp.docx";

  generatePDF(record, templateFile, `fsic-${type}-${record.id}`, res);
});

// IO / REINSPECTION / NFSI for records
app.get("/records/:id/:docType/pdf", async (req, res) => {
  const record = await findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const dt = String(req.params.docType || "").toLowerCase();
  let templateFile = "";
  if (dt === "io") templateFile = "officers.docx";
  else if (dt === "reinspection") templateFile = "reinspection.docx";
  else if (dt === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(record, templateFile, `${dt}-${record.id}`, res);
});

// DOCUMENTS PDF generation (Firestore)
app.get("/documents/:id/:docType/pdf", async (req, res) => {
  const doc = await findDocumentById(req.params.id);
  if (!doc) return res.status(404).send("Document not found");

  const dt = String(req.params.docType || "").toLowerCase();
  let templateFile = "";
  if (dt === "io") templateFile = "officers.docx";
  else if (dt === "reinspection") templateFile = "reinspection.docx";
  else if (dt === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(doc, templateFile, `doc-${dt}-${doc.id}`, res);
});

app.get("/", (req, res) => res.send("✅ BFP Backend Running (Firestore Option B)"));

app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
  console.log("PIN:", process.env.PIN ? "(set)" : "(default 1234)");
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
  console.log("FIREBASE:", fdb ? "connected (check /health)" : "NOT initialized");
});