// server.js (UPDATED FULL) — BFP System Backend (Docker + Render Ready)
// ✅ Current + Archive + Documents
// ✅ Renew logs (history.json) + get latest renewed per entityKey
// ✅ Data Manager endpoints (list + delete)
// ✅ Export endpoint (returns JSON list with latest renewed per entityKey for a month)
// ✅ PDF generation endpoints (LibreOffice in Docker)

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { exec } from "child_process";
import { fileURLToPath } from "url";
import os from "os";

// -----------------------------
// PATHS / APP
// -----------------------------
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
app.use(express.json({ limit: "10mb" }));

// -----------------------------
// FILES
// -----------------------------
const DATA_FILE = path.join(__dirname, "records.json"); // current records (array)
const ARCHIVE_FILE = path.join(__dirname, "archive.json"); // { "YYYY-MM": [records...] }
const DOCUMENTS_FILE = path.join(__dirname, "documents.json"); // documents (array)
const HISTORY_FILE = path.join(__dirname, "history.json"); // renew logs (array)

// -----------------------------
// HELPERS
// -----------------------------
const ensureFile = (file, defaultData) => {
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, JSON.stringify(defaultData, null, 2));
  }
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
  // 1) current
  let record = (readJSON(DATA_FILE) || []).find((r) => String(r.id) == String(id));
  if (record) return ensureEntityKey(record);

  // 2) archive
  const archive = readJSON(ARCHIVE_FILE) || {};
  for (const month of Object.keys(archive)) {
    record = (archive[month] || []).find((r) => String(r.id) == String(id));
    if (record) return ensureEntityKey(record);
  }
  return null;
};

// ✅ accepts BOTH:
// - old ALL CAPS (FSIC_APP_NO, OWNERS_NAME...)
// - new camelCase (fsicAppNo, ownerName...)
const pickAllowedRecordFields = (obj = {}) => {
  return {
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
  };
};

// ✅ Documents fields (store what templates need for IO/NFSI/Reinspection too)
const pickAllowedDocumentFields = (obj = {}) => {
  return {
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

    // signatures
    chiefName: obj.chiefName ?? obj.CHIEF ?? "",
    marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
  };
};

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
// RECORDS (CURRENT)
// -----------------------------
app.get("/records", (req, res) => {
  const records = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
  res.json(records);
});

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

    return res.json({ success: true, data: newRecord });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Add record failed" });
  }
});

app.delete("/records/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const list = readJSON(DATA_FILE) || [];
    const before = list.length;
    const after = list.filter((r) => Number(r.id) !== id);
    writeJSON(DATA_FILE, after);
    return res.json({ success: true, deleted: before - after.length });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// -----------------------------
// CLOSE MONTH => MOVE CURRENT TO ARCHIVE
// -----------------------------
app.post("/records/close-month", (req, res) => {
  try {
    const records = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    if (records.length === 0) {
      return res.json({ success: false, message: "No records" });
    }

    const now = new Date();
    const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

    const archive = readJSON(ARCHIVE_FILE) || {};
    if (!archive[monthKey]) archive[monthKey] = [];

    archive[monthKey] = [...(archive[monthKey] || []), ...records];

    writeJSON(ARCHIVE_FILE, archive);
    writeJSON(DATA_FILE, []);

    return res.json({ success: true, month: monthKey, archivedCount: records.length });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Close month failed" });
  }
});

// -----------------------------
// ARCHIVE
// -----------------------------
app.get("/archive/months", (req, res) => {
  const archive = readJSON(ARCHIVE_FILE) || {};
  res.json(Object.keys(archive));
});

app.get("/archive/:month", (req, res) => {
  const archive = readJSON(ARCHIVE_FILE) || {};
  const list = (archive[req.params.month] || []).map(ensureEntityKey);
  res.json(list);
});

app.delete("/archive/:month/:id", (req, res) => {
  try {
    const month = String(req.params.month || "");
    const id = Number(req.params.id);

    const archive = readJSON(ARCHIVE_FILE) || {};
    const list = archive[month] || [];
    const before = list.length;

    archive[month] = list.filter((r) => Number(r.id) !== id);
    writeJSON(ARCHIVE_FILE, archive);

    return res.json({ success: true, deleted: before - (archive[month] || []).length });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// -----------------------------
// DOCUMENTS
// -----------------------------
app.get("/documents", (req, res) => {
  res.json(readJSON(DOCUMENTS_FILE) || []);
});

app.post("/documents", (req, res) => {
  try {
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const newDoc = {
      id: Date.now(),
      createdAt: new Date().toISOString(),
      ...pickAllowedDocumentFields(req.body || {}),
    };
    docs.push(newDoc);
    writeJSON(DOCUMENTS_FILE, docs);
    return res.json({ success: true, data: newDoc });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Add document failed" });
  }
});

app.put("/documents/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const idx = docs.findIndex((d) => Number(d.id) === id);

    if (idx === -1) return res.status(404).json({ success: false, message: "Document not found" });

    docs[idx] = {
      ...docs[idx],
      ...pickAllowedDocumentFields({ ...docs[idx], ...(req.body || {}) }),
      updatedAt: new Date().toISOString(),
    };

    writeJSON(DOCUMENTS_FILE, docs);
    return res.json({ success: true, data: docs[idx] });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Update document failed" });
  }
});

app.delete("/documents/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const docs = readJSON(DOCUMENTS_FILE) || [];
    const before = docs.length;
    const after = docs.filter((d) => Number(d.id) !== id);
    writeJSON(DOCUMENTS_FILE, after);
    return res.json({ success: true, deleted: before - after.length });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Delete failed" });
  }
});

// -----------------------------
// RENEW (HISTORY)
// -----------------------------
app.get("/records/renewed/:entityKey", (req, res) => {
  try {
    const ek = normalizeEntityKey(decodeURIComponent(req.params.entityKey || ""));
    if (!ek) return res.json({ success: true, record: null });

    const latest = getLatestRenewedByEntityKey(ek);
    return res.json({ success: true, record: latest || null });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Failed to load renewed record" });
  }
});

app.get("/records/renewed", (req, res) => {
  try {
    const history = readJSON(HISTORY_FILE) || [];

    const renewedRecords = history
      .filter((h) => String(h.action || "").toUpperCase() === "RENEWED")
      .map((h) => {
        const rec = ensureEntityKey(h.data || {});
        return {
          ...rec,
          entityKey: h.entityKey || rec.entityKey,
          renewedAt: h.changedAt || rec.renewedAt || "",
          source: h.source || "Renewed",
        };
      })
      .sort((a, b) => String(b.renewedAt || "").localeCompare(String(a.renewedAt || "")));

    return res.json({ success: true, records: renewedRecords });
  } catch (e) {
    console.error("GET /records/renewed error:", e);
    return res.status(500).json({ success: false, records: [], message: "Failed to load renewed" });
  }
});

app.delete("/records/renewed/:id", (req, res) => {
  try {
    const id = Number(req.params.id);
    const history = readJSON(HISTORY_FILE) || [];
    const before = history.length;

    const after = history.filter((h) => {
      if (String(h.action || "").toUpperCase() !== "RENEWED") return true;
      const rec = h.data || {};
      return Number(rec.id) !== id;
    });

    writeJSON(HISTORY_FILE, after);
    return res.json({ success: true, deleted: before - after.length });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Delete renewed failed" });
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
      return res.status(400).json({
        success: false,
        message: "Missing payload (entityKey/oldRecord/updatedRecord)",
      });
    }

    const now = new Date().toISOString();
    const history = readJSON(HISTORY_FILE) || [];

    history.push({
      entityKey,
      source: source || "Unknown",
      changedAt: now,
      action: "PREVIOUS",
      data: oldRecord,
    });

    const newRecord = buildRenewedRecord({ entityKey, updatedRecord });

    history.push({
      entityKey,
      source: "Renewed",
      changedAt: now,
      action: "RENEWED",
      data: newRecord,
    });

    writeJSON(HISTORY_FILE, history);

    return res.json({ success: true, newRecord });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Renew failed" });
  }
});

// -----------------------------
// EXPORT
// -----------------------------
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

    return res.json({ success: true, records: replaced });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Export failed", records: [] });
  }
});

// -----------------------------
// DATA MANAGER
// -----------------------------
app.get("/manager/items", (req, res) => {
  try {
    const scope = normalize(req.query.scope || "all");
    const month = normalize(req.query.month || "");

    const items = [];

    const pushItem = ({ kind, id, source, createdAt, changedAt, entityKey, data, month }) => {
      items.push({
        kind,
        id,
        source,
        createdAt: createdAt || "",
        changedAt: changedAt || "",
        entityKey: entityKey || "",
        month: month || "",
        data: data || {},
      });
    };

    if (scope === "all" || scope === "current") {
      const current = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
      current.forEach((r) =>
        pushItem({
          kind: "current",
          id: r.id,
          source: "Current",
          createdAt: r.createdAt,
          entityKey: r.entityKey,
          data: r,
        })
      );
    }

    if (scope === "all" || scope === "archive") {
      const archive = readJSON(ARCHIVE_FILE) || {};
      const months = month ? [month] : Object.keys(archive);

      months.forEach((m) => {
        (archive[m] || []).map(ensureEntityKey).forEach((r) =>
          pushItem({
            kind: "archive",
            id: r.id,
            source: `Archive:${m}`,
            createdAt: r.createdAt,
            entityKey: r.entityKey,
            data: r,
            month: m,
          })
        );
      });
    }

    if (scope === "all" || scope === "documents") {
      const docs = readJSON(DOCUMENTS_FILE) || [];
      docs.forEach((d) =>
        pushItem({
          kind: "documents",
          id: d.id,
          source: "Documents",
          createdAt: d.createdAt,
          data: d,
        })
      );
    }

    if (scope === "all" || scope === "renewed") {
      const history = readJSON(HISTORY_FILE) || [];
      history
        .filter((h) => String(h.action || "").toUpperCase() === "RENEWED")
        .forEach((h) => {
          const rec = ensureEntityKey(h.data || {});
          pushItem({
            kind: "renewed",
            id: rec.id,
            source: "Renewed",
            changedAt: h.changedAt,
            entityKey: h.entityKey || rec.entityKey,
            data: rec,
          });
        });
    }

    items.sort((a, b) =>
      String(b.changedAt || b.createdAt).localeCompare(String(a.changedAt || a.createdAt))
    );

    return res.json({ success: true, items });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, items: [] });
  }
});

// -----------------------------
// PIN AUTH
// -----------------------------
const CORRECT_PIN = String(process.env.PIN || "1234").trim();

app.post("/auth/pin", (req, res) => {
  const pin = String(req.body?.pin || "").trim();
  if (!pin) return res.status(400).json({ ok: false, message: "Missing PIN" });
  if (pin === CORRECT_PIN) return res.json({ ok: true });
  return res.status(401).json({ ok: false, message: "Incorrect PIN" });
});

// -----------------------------
// PDF GENERATION (LibreOffice)
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

const generatePDF = (record, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);
  if (!fs.existsSync(templatePath)) {
    return res.status(404).send(`Template not found: ${templateFile}`);
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res
      .status(500)
      .send("LibreOffice not found. Install LibreOffice or fix SOFFICE_PATH.");
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
      // ✅ if tag not provided, return empty string (so it won't crash)
      nullGetter: () => "",
    });

    // ✅ Build a single "view" object with ALL tags used by your templates:
    // FSIC templates use {CHIEF} {MARSHAL} etc. 
    // IO template uses {IO_NUMBER} {IO_DATE} {TAXPAYER} ... :contentReference[oaicite:3]{index=3}
    // Reinspection template uses same tags :contentReference[oaicite:4]{index=4}
    // NFSI uses {NFSI_NUMBER} {NFSI_DATE} {MARSHAL} :contentReference[oaicite:5]{index=5}

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
        return res.status(500).send("PDF conversion failed. Check backend logs.");
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");

      if (!fs.existsSync(expectedPdf)) {
        console.log("PDF not found after conversion.");
        console.log("expectedPdf:", expectedPdf);
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        return res.status(500).send("PDF file not produced. Check LibreOffice conversion.");
      }

      res.download(expectedPdf, () => {
        try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
        try { if (fs.existsSync(expectedPdf)) fs.unlinkSync(expectedPdf); } catch {}
      });
    });
  } catch (e) {
    console.log("PDF generation failed:", e);
    try { if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx); } catch {}
    return res.status(500).send("PDF generation failed (docxtemplater/render). Check logs.");
  }
};

// FSIC Certificate
app.get("/records/:id/certificate/:type/pdf", (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const templateFile = req.params.type === "owner" ? "fsic-owner.docx" : "fsic-bfp.docx";
  generatePDF(record, templateFile, `fsic-${record.id}`, res);
});

// IO / REINSPECTION / NFSI for records
app.get("/records/:id/:docType/pdf", (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  let templateFile = "";
  if (req.params.docType === "io") templateFile = "officers.docx";
  else if (req.params.docType === "reinspection") templateFile = "reinspection.docx";
  else if (req.params.docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(record, templateFile, `${req.params.docType}-${record.id}`, res);
});

// DOCUMENTS PDF generation
app.get("/documents/:id/:docType/pdf", (req, res) => {
  const doc = (readJSON(DOCUMENTS_FILE) || []).find((r) => String(r.id) == String(req.params.id));
  if (!doc) return res.status(404).send("Document not found");

  let templateFile = "";
  if (req.params.docType === "io") templateFile = "officers.docx";
  else if (req.params.docType === "reinspection") templateFile = "reinspection.docx";
  else if (req.params.docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  generatePDF(doc, templateFile, `doc-${req.params.docType}-${doc.id}`, res);
});

// -----------------------------
// START
// -----------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
  console.log("PIN:", process.env.PIN ? "(set)" : "(default 1234)");
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
});
