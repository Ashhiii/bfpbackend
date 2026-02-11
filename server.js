// server.js (FULL) — BFP System Backend
// ✅ Current + Archive + Documents
// ✅ Renew logs (history.json) + get latest renewed per entityKey
// ✅ Data Manager endpoints (list + delete)
// ✅ Export endpoint (returns JSON list with latest renewed per entityKey for a month)
// ✅ PDF generation endpoints (LibreOffice via SOFFICE_PATH or auto-detect)

// -----------------------------
// LOAD .env FIRST
// -----------------------------
import "dotenv/config";

// -----------------------------
// IMPORTS
// -----------------------------
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import os from "os";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { exec, execSync } from "child_process";
import { fileURLToPath } from "url";

// -----------------------------
// PATHS / APP
// -----------------------------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

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
const writeJSON = (file, data) =>
  fs.writeFileSync(file, JSON.stringify(data, null, 2));

const normalize = (v) => String(v ?? "").trim();

const ensureEntityKey = (r) => {
  if (!r) return r;
  if (r.entityKey) return r;
  const fsic = normalize(r.fsicAppNo);
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

const pickAllowedRecordFields = (obj = {}) => {
  // ✅ keeps only record fields (prevents weird extra keys)
  return {
    no: obj.no ?? "",
    fsicAppNo: obj.fsicAppNo ?? "",
    natureOfInspection: obj.natureOfInspection ?? "",
    ownerName: obj.ownerName ?? "",
    establishmentName: obj.establishmentName ?? "",
    businessAddress: obj.businessAddress ?? "",
    contactNumber: obj.contactNumber ?? "",
    dateInspected: obj.dateInspected ?? "",
    ioNumber: obj.ioNumber ?? "",
    ioDate: obj.ioDate ?? "",
    nfsiNumber: obj.nfsiNumber ?? "",
    nfsiDate: obj.nfsiDate ?? "",
    fsicValidity: obj.fsicValidity ?? "",
    defects: obj.defects ?? "",
    inspectors: obj.inspectors ?? "",
    occupancyType: obj.occupancyType ?? "",
    buildingDesc: obj.buildingDesc ?? "",
    floorArea: obj.floorArea ?? "",
    buildingHeight: obj.buildingHeight ?? "",
    storeyCount: obj.storeyCount ?? "",
    highRise: obj.highRise ?? "",
    fsmr: obj.fsmr ?? "",
    remarks: obj.remarks ?? "",
    orNumber: obj.orNumber ?? "",
    orAmount: obj.orAmount ?? "",
    orDate: obj.orDate ?? "",
  };
};

const buildRenewedRecord = ({ entityKey, updatedRecord }) => {
  // ✅ teamLeader only appears on renewed record
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
// ✅ AUTH PIN
// -----------------------------
app.post("/auth/pin", (req, res) => {
  const { pin } = req.body || {};
  const OK_PIN = process.env.PIN || "1234";

  if (String(pin) === String(OK_PIN)) {
    return res.json({ ok: true });
  }
  return res.status(401).json({ ok: false });
});

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
      // keep entityKey if passed; else generate from FSIC
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

// DELETE current record
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

// DELETE archived record by month
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
      ...req.body,
    };
    docs.push(newDoc);
    writeJSON(DOCUMENTS_FILE, docs);
    return res.json({ success: true, data: newDoc });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Add document failed" });
  }
});

// DELETE document
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

// Get latest renewed record for this entityKey
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

// List all renewed (simple)
app.get("/records/renewed-all", (req, res) => {
  try {
    const history = readJSON(HISTORY_FILE) || [];
    const renewed = history
      .filter((h) => String(h.action || "").toUpperCase() === "RENEWED")
      .map((h) => ({
        entityKey: h.entityKey,
        changedAt: h.changedAt,
        data: ensureEntityKey(h.data),
      }));
    return res.json({ success: true, renewed });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, renewed: [] });
  }
});

// Delete a renewed record log by record id (removes matching RENEWED entries)
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

// Renew endpoint
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

    // log PREVIOUS snapshot
    history.push({
      entityKey,
      source: source || "Unknown",
      changedAt: now,
      action: "PREVIOUS",
      data: oldRecord,
    });

    // build RENEWED record
    const newRecord = buildRenewedRecord({ entityKey, updatedRecord });

    // log RENEWED
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
// EXPORT endpoint (for archive month)
// returns JSON array where each record is replaced by latest renewed (if exists)
// -----------------------------
app.get("/records/export", (req, res) => {
  try {
    const month = normalize(req.query.month);
    if (!month) {
      return res.status(400).json({ success: false, message: "Missing month" });
    }

    const archive = readJSON(ARCHIVE_FILE) || {};
    const list = (archive[month] || []).map(ensureEntityKey);

    // replace by latest renewed if exists
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
// DATA MANAGER endpoint (list combined items)
// -----------------------------
app.get("/manager/items", (req, res) => {
  try {
    const scope = normalize(req.query.scope || "all"); // all|current|archive|documents|renewed
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

    // current
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

    // archive
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

    // documents
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

    // renewed logs
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

    // newest first
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
// PDF GENERATION (LibreOffice)
// -----------------------------
const findSoffice = () => {
  // 1) Use env first
  const envPath = process.env.SOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;

  // 2) Common paths
  const candidates = [
    // Windows
    "C:\\\\Program Files\\\\LibreOffice\\\\program\\\\soffice.exe",
    "C:\\\\Program Files (x86)\\\\LibreOffice\\\\program\\\\soffice.exe",
    // Linux
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/snap/bin/libreoffice",
    "/snap/bin/soffice",
    // Mac
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }

  // 3) Try which (Linux/Mac)
  try {
    const which = (cmd) =>
      execSync(`which ${cmd}`, { stdio: ["ignore", "pipe", "ignore"] })
        .toString()
        .trim();
    const p1 = which("soffice");
    if (p1 && fs.existsSync(p1)) return p1;
    const p2 = which("libreoffice");
    if (p2 && fs.existsSync(p2)) return p2;
  } catch {}

  return null;
};

const generatePDF = (record, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);
  if (!fs.existsSync(templatePath)) {
    return res.status(404).send(`Template not found: ${templateFile}`);
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res.status(500).send(
      "LibreOffice not found. Install LibreOffice OR set SOFFICE_PATH in .env to your soffice.exe path."
    );
  }

  // ✅ unique per request
  const stamp = `${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
  const outDir = path.join(os.tmpdir(), "bfp_pdf_out");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outputDocx = path.join(outDir, `${filenameBase}-${stamp}.docx`);

  try {
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    doc.render({
      FSIC_NUMBER: record.fsicAppNo || "",
      DATE_INSPECTED: record.dateInspected || "",
      NAME_OF_ESTABLISHMENT: record.establishmentName || "",
      NAME_OF_OWNER: record.ownerName || "",
      ADDRESS: record.businessAddress || "",
      FLOOR_AREA: record.floorArea || "",
      BLDG_DESCRIPTION: record.buildingDesc || "",
      FSIC_VALIDITY: record.fsicValidity || "",
      OR_NUMBER: record.orNumber || "",
      OR_DATE: record.orDate || "",
      OR_AMOUNT: record.orAmount || "",

      IO_NUMBER: record.ioNumber || "",
      IO_DATE: record.ioDate ? String(record.ioDate) : "",
      TAXPAYER: record.ownerName || "",
      TRADE_NAME: record.establishmentName || "",
      CONTACT_: record.contactNumber || "",
      NFSI_NUMBER: record.nfsiNumber || "",
      NFSI_DATE: record.nfsiDate || "",
      OWNER: record.ownerName || "",
      INSPECTORS: record.inspectors || "",
      TEAM_LEADER: record.teamLeader || "",
      DATE: new Date().toLocaleDateString(),
    });

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    fs.writeFileSync(outputDocx, buf);

    // ✅ Use quotes (spaces safe)
    const command = `"${soffice}" --headless --nologo --nolockcheck --norestore --convert-to pdf "${outputDocx}" --outdir "${outDir}"`;

    exec(command, (err, stdout, stderr) => {
      if (err) {
        console.log("LibreOffice ERROR:", err);
        console.log("stdout:", stdout);
        console.log("stderr:", stderr);
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        return res.status(500).send("PDF conversion failed. Check backend terminal logs.");
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");

      if (!fs.existsSync(expectedPdf)) {
        console.log("PDF not found after conversion.");
        console.log("expectedPdf:", expectedPdf);
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        return res.status(500).send("PDF file not produced. Check LibreOffice conversion.");
      }

      res.download(expectedPdf, () => {
        // cleanup both
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
    return res.status(500).send("PDF generation failed (docxtemplater/render). Check terminal.");
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
  const doc = (readJSON(DOCUMENTS_FILE) || []).find(
    (r) => String(r.id) == String(req.params.id)
  );
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
app.listen(PORT, () => {
  console.log(`✅ Backend running at http://localhost:${PORT}`);
  console.log("PIN:", process.env.PIN ? "(set)" : "(default 1234)");
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
});
