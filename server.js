// server.js — BFP System Backend (FULL FIX)
// ✅ CORS + OPTIONS
// ✅ /health
// ✅ PDF generation (LibreOffice) + clear errors
// ✅ Records + Archive + Documents + Renew + Manager + Export
// ✅ Excel import endpoints (/import/records, /import/documents, /import/excel)

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
app.options("*", cors()); // ✅ IMPORTANT FOR PREFLIGHT
app.use(express.json({ limit: "10mb" }));

// multer memory upload
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
});

// -----------------------------
// FILES
// -----------------------------
const DATA_FILE = path.join(__dirname, "records.json");
const ARCHIVE_FILE = path.join(__dirname, "archive.json");
const DOCUMENTS_FILE = path.join(__dirname, "documents.json");
const HISTORY_FILE = path.join(__dirname, "history.json");

// -----------------------------
// HELPERS
// -----------------------------
const ensureFile = (file, defaultData) => {
  if (!fs.existsSync(file)) fs.writeFileSync(file, JSON.stringify(defaultData, null, 2));
};
ensureFile(DATA_FILE, []);
ensureFile(ARCHIVE_FILE, {});
ensureFile(DOCUMENTS_FILE, []);
ensureFile(HISTORY_FILE, []);

const readJSON = (file) => JSON.parse(fs.readFileSync(file, "utf-8"));
const writeJSON = (file, data) => fs.writeFileSync(file, JSON.stringify(data, null, 2));
const normalize = (v) => String(v ?? "").trim();

const ensureEntityKey = (r) => {
  if (!r) return r;
  if (r.entityKey) return r;
  const fsic = normalize(r.fsicAppNo || r.FSIC_APP_NO || r.FSIC_NUMBER);
  return { ...r, entityKey: fsic ? `fsic:${fsic}` : `rec:${r.id || Date.now()}` };
};

const normalizeEntityKey = (entityKey) => normalize(entityKey);

const findRecordById = (id) => {
  let record = (readJSON(DATA_FILE) || []).find((r) => String(r.id) == String(id));
  if (record) return ensureEntityKey(record);

  const archive = readJSON(ARCHIVE_FILE) || {};
  for (const month of Object.keys(archive)) {
    record = (archive[month] || []).find((r) => String(r.id) == String(id));
    if (record) return ensureEntityKey(record);
  }
  return null;
};

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

  nfsiNumber: obj.nfsiNumber ?? obj.NFSI_NUMBER ?? "",
  nfsiDate: obj.nfsiDate ?? obj.NFSI_DATE ?? "",

  inspectors: obj.inspectors ?? obj.INSPECTORS ?? "",
  teamLeader: obj.teamLeader ?? obj.TEAM_LEADER ?? "",

  chiefName: obj.chiefName ?? obj.CHIEF ?? "",
  marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
});

const buildRenewedRecord = ({ entityKey, updatedRecord }) => {
  const base = pickAllowedRecordFields(updatedRecord || {});
  const now = new Date().toISOString();

  const newRecord = ensureEntityKey({
    id: Date.now(),
    entityKey,
    ...base,
    teamLeader: updatedRecord?.teamLeader || "",
    renewedAt: now,
    createdAt: now,
  });

  return newRecord;
};

const getLatestRenewedByEntityKey = (entityKey) => {
  const ek = normalizeEntityKey(entityKey);
  if (!ek) return null;

  const history = readJSON(HISTORY_FILE) || [];
  const renewed = history
    .filter(
      (h) =>
        normalizeEntityKey(h.entityKey) === ek &&
        String(h.action || "").toUpperCase() === "RENEWED"
    )
    .sort((a, b) => String(a.changedAt || "").localeCompare(String(b.changedAt || "")));

  if (!renewed.length) return null;
  const last = renewed[renewed.length - 1];
  return ensureEntityKey(last.data);
};

// -----------------------------
// EXCEL IMPORT HELPERS
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
    establishmentName: String(get("establishment", "establishmentname", "tradename", "nameofestablishment") || ""),
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

const mapExcelRowToDocument = (row = {}) => {
  const headerMap = {};
  for (const k of Object.keys(row)) headerMap[normHeader(k)] = row[k];
  const get = (...variants) => pickAny(headerMap, variants.map(normHeader));

  const doc = {
    id: makeId(),
    createdAt: new Date().toISOString(),

    fsicAppNo: String(get("fsicappno", "fsicno", "fsicnumber", "fsicapp#", "fsicapp") || ""),
    ownerName: String(get("owner", "ownername", "ownersname", "taxpayer") || ""),
    establishmentName: String(get("establishment", "establishmentname", "tradename", "nameofestablishment") || ""),
    businessAddress: String(get("businessaddress", "address", "bussinessaddress") || ""),
    contactNumber: String(get("contactnumber", "contact", "mobile") || ""),

    ioNumber: String(get("ionumber", "io#", "io") || ""),
    ioDate: excelDateToISO(get("iodate") || ""),

    nfsiNumber: String(get("nfsinumber", "nfsi#", "nfsi") || ""),
    nfsiDate: excelDateToISO(get("nfsidate") || ""),

    inspectors: String(get("inspectors", "inspector") || ""),
    teamLeader: String(get("teamleader", "team_leader") || ""),

    chiefName: String(get("chiefname", "chief") || ""),
    marshalName: String(get("marshalname", "marshal") || ""),
  };

  doc.fsicAppNo = String(doc.fsicAppNo || "").toUpperCase().trim();
  doc.ownerName = String(doc.ownerName || "").toUpperCase().trim();
  doc.establishmentName = String(doc.establishmentName || "").toUpperCase().trim();
  doc.businessAddress = String(doc.businessAddress || "").toUpperCase().trim();

  return doc;
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

    // map fields for templates (works with old caps + camelCase)
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
      INSPECTORS: record.INSPECTORS || record.inspectors || "",
      TEAM_LEADER: record.TEAM_LEADER || record.teamLeader || "",

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
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        return res
          .status(500)
          .send(`PDF conversion failed. ${String(stderr || err?.message || err)}`);
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");
      if (!fs.existsSync(expectedPdf)) {
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        return res.status(500).send("PDF file not produced after conversion.");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `attachment; filename="${filenameBase}.pdf"`);

      res.download(expectedPdf, () => {
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        try { if (fs.existsSync(expectedPdf)) fs.unlinkSync(expectedPdf); } catch {}
      });
    });
  } catch (e) {
    console.log("PDF generation failed:", e);
    try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
    return res.status(500).send(`PDF generation failed (templater). ${e.message}`);
  }
};

// ✅ health check (see if soffice + templates exist)
app.get("/health", (req, res) => {
  const soffice = findSoffice();
  const templatesDir = path.join(__dirname, "templates");
  res.json({
    ok: true,
    sofficeFound: Boolean(soffice),
    sofficePath: soffice,
    templatesDirExists: fs.existsSync(templatesDir),
    templates: fs.existsSync(templatesDir) ? fs.readdirSync(templatesDir) : [],
  });
});

// -----------------------------
// EXCEL IMPORT ENDPOINTS
// -----------------------------
app.post("/import/records", upload.single("file"), (req, res) => {
  try {
    const { rows, sheetName, error } = readExcelRows(req);
    if (error) return res.status(400).json({ success: false, message: error });

    const current = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    const mapped = rows.map(mapExcelRowToRecord);
    const toAdd = mapped.filter((r) => normalize(r.fsicAppNo) && normalize(r.ownerName));

    writeJSON(DATA_FILE, [...current, ...toAdd]);

    res.json({ success: true, imported: toAdd.length, skipped: mapped.length - toAdd.length, sheet: sheetName });
  } catch (e) {
    console.error("POST /import/records error:", e);
    res.status(500).json({ success: false, message: "Failed to import records Excel." });
  }
});

app.post("/import/documents", upload.single("file"), (req, res) => {
  try {
    const { rows, sheetName, error } = readExcelRows(req);
    if (error) return res.status(400).json({ success: false, message: error });

    const current = readJSON(DOCUMENTS_FILE) || [];
    const mapped = rows.map(mapExcelRowToDocument);
    const toAdd = mapped.filter((d) => normalize(d.fsicAppNo) && (normalize(d.establishmentName) || normalize(d.ownerName)));

    writeJSON(DOCUMENTS_FILE, [...current, ...toAdd]);

    res.json({ success: true, imported: toAdd.length, skipped: mapped.length - toAdd.length, sheet: sheetName });
  } catch (e) {
    console.error("POST /import/documents error:", e);
    res.status(500).json({ success: false, message: "Failed to import documents Excel." });
  }
});

app.post("/import/excel", upload.single("file"), (req, res) => {
  try {
    const { rows, sheetName, error } = readExcelRows(req);
    if (error) return res.status(400).json({ success: false, message: error });

    const current = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    const mapped = rows.map(mapExcelRowToRecord);
    const toAdd = mapped.filter((r) => normalize(r.fsicAppNo) && normalize(r.ownerName));

    writeJSON(DATA_FILE, [...current, ...toAdd]);

    res.json({ success: true, imported: toAdd.length, skipped: mapped.length - toAdd.length, sheet: sheetName });
  } catch (e) {
    console.error("POST /import/excel error:", e);
    res.status(500).json({ success: false, message: "Failed to import Excel." });
  }
});

// -----------------------------
// RECORDS (CURRENT)
// -----------------------------
app.get("/records", (req, res) => res.json((readJSON(DATA_FILE) || []).map(ensureEntityKey)));

app.post("/records", (req, res) => {
  try {
    const records = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    let newRecord = {
      id: Date.now(),
      createdAt: new Date().toISOString(),
      ...pickAllowedRecordFields(req.body || {}),
      entityKey: req.body?.entityKey,
    };
    newRecord = ensureEntityKey(newRecord);
    records.push(newRecord);
    writeJSON(DATA_FILE, records);
    res.json({ success: true, data: newRecord });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Add record failed" });
  }
});

app.put("/records/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const records = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    const idx = records.findIndex((r) => Number(r.id) === id);
    if (idx === -1) return res.status(404).json({ success: false, message: "Record not found" });

    const allowed = pickAllowedRecordFields(req.body || {});
    const teamLeader = req.body?.teamLeader ?? records[idx]?.teamLeader ?? "";

    records[idx] = ensureEntityKey({
      ...records[idx],
      ...allowed,
      teamLeader,
      updatedAt: new Date().toISOString(),
    });

    writeJSON(DATA_FILE, records);
    res.json({ success: true, data: records[idx] });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Update record failed" });
  }
});

app.delete("/records/:id", (req, res) => {
  try {
    const id = String(req.params.id);
    const list = readJSON(DATA_FILE) || [];
    const before = list.length;
    const after = list.filter((r) => String(r.id) !== id);
    writeJSON(DATA_FILE, after);
    res.json({ success: true, deleted: before - after.length });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// close month
app.post("/records/close-month", (req, res) => {
  try {
    const records = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    if (!records.length) return res.json({ success: false, message: "No records" });

    const now = new Date();
    const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

    const archive = readJSON(ARCHIVE_FILE) || {};
    archive[monthKey] = [...(archive[monthKey] || []), ...records];

    writeJSON(ARCHIVE_FILE, archive);
    writeJSON(DATA_FILE, []);

    res.json({ success: true, month: monthKey, archivedCount: records.length });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Close month failed" });
  }
});

// archive
app.get("/archive/months", (req, res) => res.json(Object.keys(readJSON(ARCHIVE_FILE) || {})));
app.get("/archive/:month", (req, res) => {
  const archive = readJSON(ARCHIVE_FILE) || {};
  res.json((archive[req.params.month] || []).map(ensureEntityKey));
});
app.delete("/archive/:month/:id", (req, res) => {
  try {
    const month = String(req.params.month || "");
    const id = String(req.params.id);

    const archive = readJSON(ARCHIVE_FILE) || {};
    const list = archive[month] || [];
    const before = list.length;

    archive[month] = list.filter((r) => String(r.id) !== id);
    writeJSON(ARCHIVE_FILE, archive);

    res.json({ success: true, deleted: before - (archive[month] || []).length });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// documents
app.get("/documents", (req, res) => res.json(readJSON(DOCUMENTS_FILE) || []));
app.post("/documents", (req, res) => {
  try {
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const newDoc = { id: Date.now(), createdAt: new Date().toISOString(), ...pickAllowedDocumentFields(req.body || {}) };
    docs.push(newDoc);
    writeJSON(DOCUMENTS_FILE, docs);
    res.json({ success: true, data: newDoc });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Add document failed" });
  }
});
app.put("/documents/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const idx = docs.findIndex((d) => Number(d.id) === id);
    if (idx === -1) return res.status(404).json({ success: false, message: "Document not found" });

    docs[idx] = { ...docs[idx], ...pickAllowedDocumentFields({ ...docs[idx], ...(req.body || {}) }), updatedAt: new Date().toISOString() };
    writeJSON(DOCUMENTS_FILE, docs);
    res.json({ success: true, data: docs[idx] });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Update document failed" });
  }
});
app.delete("/documents/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const before = docs.length;
    const after = docs.filter((d) => Number(d.id) !== id);
    writeJSON(DOCUMENTS_FILE, after);
    res.json({ success: true, deleted: before - after.length });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// renewed
app.get("/records/renewed/:entityKey", (req, res) => {
  try {
    const ek = normalizeEntityKey(decodeURIComponent(req.params.entityKey || ""));
    if (!ek) return res.json({ success: true, record: null });
    const latest = getLatestRenewedByEntityKey(ek);
    res.json({ success: true, record: latest || null });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Failed to load renewed record" });
  }
});

app.get("/records/renewed", (req, res) => {
  try {
    const history = readJSON(HISTORY_FILE) || [];
    const renewedRecords = history
      .filter((h) => String(h.action || "").toUpperCase() === "RENEWED")
      .map((h) => {
        const rec = ensureEntityKey(h.data || {});
        return { ...rec, entityKey: h.entityKey || rec.entityKey, renewedAt: h.changedAt || rec.renewedAt || "", source: h.source || "Renewed" };
      })
      .sort((a, b) => String(b.renewedAt || "").localeCompare(String(a.renewedAt || "")));
    res.json({ success: true, records: renewedRecords });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, records: [], message: "Failed to load renewed" });
  }
});

app.post("/records/renew", (req, res) => {
  try {
    let { entityKey, source, oldRecord, updatedRecord } = req.body;

    oldRecord = ensureEntityKey(oldRecord);
    updatedRecord = ensureEntityKey(updatedRecord);

    entityKey =
      normalizeEntityKey(entityKey) ||
      normalizeEntityKey(oldRecord?.entityKey) ||
      normalizeEntityKey(updatedRecord?.entityKey);

    if (!entityKey || !oldRecord || !updatedRecord) {
      return res.status(400).json({ success: false, message: "Missing payload (entityKey/oldRecord/updatedRecord)" });
    }

    const now = new Date().toISOString();
    const history = readJSON(HISTORY_FILE) || [];

    history.push({ entityKey, source: source || "Unknown", changedAt: now, action: "PREVIOUS", data: oldRecord });

    const newRecord = buildRenewedRecord({ entityKey, updatedRecord });

    history.push({ entityKey, source: "Renewed", changedAt: now, action: "RENEWED", data: newRecord });

    writeJSON(HISTORY_FILE, history);
    res.json({ success: true, newRecord });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Renew failed" });
  }
});

// export (latest renewed replacement)
app.get("/records/export", (req, res) => {
  try {
    const month = normalize(req.query.month);
    if (!month) return res.status(400).json({ success: false, message: "Missing month" });

    const archive = readJSON(ARCHIVE_FILE) || {};
    const list = (archive[month] || []).map(ensureEntityKey);

    const replaced = list.map((r) => {
      const ek = normalizeEntityKey(r.entityKey);
      const latest = ek ? getLatestRenewedByEntityKey(ek) : null;
      return latest || r;
    });

    res.json({ success: true, records: replaced });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, message: "Export failed", records: [] });
  }
});

// manager (simple)
app.get("/manager/items", (req, res) => {
  try {
    const scope = normalize(req.query.scope || "all");
    const month = normalize(req.query.month || "");
    const items = [];
    const pushItem = ({ kind, id, source, createdAt, changedAt, entityKey, data, month }) =>
      items.push({ kind, id, source, createdAt: createdAt || "", changedAt: changedAt || "", entityKey: entityKey || "", month: month || "", data: data || {} });

    if (scope === "all" || scope === "current") {
      (readJSON(DATA_FILE) || []).map(ensureEntityKey).forEach((r) =>
        pushItem({ kind: "current", id: r.id, source: "Current", createdAt: r.createdAt, entityKey: r.entityKey, data: r })
      );
    }

    if (scope === "all" || scope === "archive") {
      const archive = readJSON(ARCHIVE_FILE) || {};
      const months = month ? [month] : Object.keys(archive);
      months.forEach((m) => {
        (archive[m] || []).map(ensureEntityKey).forEach((r) =>
          pushItem({ kind: "archive", id: r.id, source: `Archive:${m}`, createdAt: r.createdAt, entityKey: r.entityKey, data: r, month: m })
        );
      });
    }

    if (scope === "all" || scope === "documents") {
      (readJSON(DOCUMENTS_FILE) || []).forEach((d) =>
        pushItem({ kind: "documents", id: d.id, source: "Documents", createdAt: d.createdAt, data: d })
      );
    }

    if (scope === "all" || scope === "renewed") {
      const history = readJSON(HISTORY_FILE) || [];
      history
        .filter((h) => String(h.action || "").toUpperCase() === "RENEWED")
        .forEach((h) => {
          const rec = ensureEntityKey(h.data || {});
          pushItem({ kind: "renewed", id: rec.id, source: "Renewed", changedAt: h.changedAt, entityKey: h.entityKey || rec.entityKey, data: rec });
        });
    }

    items.sort((a, b) => String(b.changedAt || b.createdAt).localeCompare(String(a.changedAt || a.createdAt)));
    res.json({ success: true, items });
  } catch (e) {
    console.log(e);
    res.status(500).json({ success: false, items: [] });
  }
});

// PIN auth
const CORRECT_PIN = String(process.env.PIN || "1234").trim();
app.post("/auth/pin", (req, res) => {
  const pin = String(req.body?.pin || "").trim();
  if (!pin) return res.status(400).json({ ok: false, message: "Missing PIN" });
  if (pin === CORRECT_PIN) return res.json({ ok: true });
  return res.status(401).json({ ok: false, message: "Incorrect PIN" });
});

// -----------------------------
// PDF ROUTES (GET - open in new tab)
// -----------------------------
// FSIC Certificate
app.get("/records/:id/certificate/:type/pdf", (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const type = String(req.params.type || "").toLowerCase();
  const templateFile = type === "owner" ? "fsic-owner.docx" : "fsic-bfp.docx";

  generatePDF(record, templateFile, `fsic-${type}-${record.id}`, res);
});

// IO / REINSPECTION / NFSI for records
app.get("/records/:id/:docType/pdf", (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const dt = String(req.params.docType || "").toLowerCase();
  let templateFile = "";
  if (dt === "io") templateFile = "officers.docx";
  else if (dt === "reinspection") templateFile = "reinspection.docx";
  else if (dt === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(record, templateFile, `${dt}-${record.id}`, res);
});

// DOCUMENTS PDF generation
app.get("/documents/:id/:docType/pdf", (req, res) => {
  const doc = (readJSON(DOCUMENTS_FILE) || []).find((r) => String(r.id) == String(req.params.id));
  if (!doc) return res.status(404).send("Document not found");

  const dt = String(req.params.docType || "").toLowerCase();
  let templateFile = "";
  if (dt === "io") templateFile = "officers.docx";
  else if (dt === "reinspection") templateFile = "reinspection.docx";
  else if (dt === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(doc, templateFile, `doc-${dt}-${doc.id}`, res);
});

app.get("/", (req, res) => res.send("✅ BFP Backend Running"));

app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
  console.log("PIN:", process.env.PIN ? "(set)" : "(default 1234)");
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
});
