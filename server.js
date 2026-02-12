// server.js (FULL) — BFP System Backend
// ✅ Current + Archive + Documents
// ✅ Renew logs (history.json) + get latest renewed per entityKey
// ✅ Data Manager endpoints (list + delete)
// ✅ Export endpoint (returns JSON list with latest renewed per entityKey for a month)
// ✅ PDF generation endpoints (NO LibreOffice) using mammoth + puppeteer

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { fileURLToPath } from "url";
import os from "os";
import mammoth from "mammoth";
import puppeteer from "puppeteer";

// -----------------------------
// PATHS / APP
// -----------------------------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json({ limit: "10mb" }));

// -----------------------------
// FILES
// -----------------------------
const DATA_FILE = path.join(__dirname, "records.json");        // current records (array)
const ARCHIVE_FILE = path.join(__dirname, "archive.json");     // { "YYYY-MM": [records...] }
const DOCUMENTS_FILE = path.join(__dirname, "documents.json"); // documents (array)
const HISTORY_FILE = path.join(__dirname, "history.json");     // renew logs (array)

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

const pickAllowedRecordFields = (obj = {}) => ({
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
});

const buildRenewedRecord = ({ entityKey, updatedRecord }) => {
  const base = pickAllowedRecordFields(updatedRecord || {});
  const now = new Date().toISOString();

  return ensureEntityKey({
    id: Date.now(),
    entityKey,
    ...base,
    teamLeader: updatedRecord?.teamLeader || "",
    renewedAt: now,
    createdAt: now,
  });
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
  return ensureEntityKey(renewed[renewed.length - 1].data);
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
    if (records.length === 0) return res.json({ success: false, message: "No records" });

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
    const newDoc = { id: Date.now(), createdAt: new Date().toISOString(), ...req.body };
    docs.push(newDoc);
    writeJSON(DOCUMENTS_FILE, docs);
    return res.json({ success: true, data: newDoc });
  } catch (e) {
    console.log(e);
    return res.status(500).json({ success: false, message: "Add document failed" });
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

// Get all renewed (for dashboard / other usage)
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
// EXPORT endpoint (archive month)
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
// DATA MANAGER endpoint
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
// AUTH PIN
// -----------------------------
const CORRECT_PIN = process.env.PIN || "1234";

app.post("/auth/pin", (req, res) => {
  const pin = String(req.body?.pin || "").trim();
  if (!pin) return res.status(400).json({ ok: false, message: "Missing PIN" });
  if (pin === CORRECT_PIN) return res.json({ ok: true });
  return res.status(401).json({ ok: false, message: "Incorrect PIN" });
});

// -----------------------------
// PDF GENERATION (NO LibreOffice)
// -----------------------------
const tempDir = path.join(os.tmpdir(), "bfp_pdf_out");
if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

const generatePDF = async (record, templateFile, filenameBase, res) => {
  const templatePath = path.join(__dirname, "templates", templateFile);
  if (!fs.existsSync(templatePath)) {
    return res.status(404).send(`Template not found: ${templateFile}`);
  }

  // unique per request
  const stamp = `${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
  const outputDocx = path.join(tempDir, `${filenameBase}-${stamp}.docx`);

  try {
    // 1) Render DOCX from template
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

    // 2) DOCX -> HTML
    const { value: htmlBody } = await mammoth.convertToHtml({ path: outputDocx });

    const html = `
<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<style>
  body { font-family: Arial, sans-serif; font-size: 12px; padding: 28px; }
  table { width: 100%; border-collapse: collapse; }
  td, th { vertical-align: top; }
</style>
</head>
<body>
${htmlBody}
</body>
</html>`.trim();

    // 3) HTML -> PDF (Puppeteer)
    const browser = await puppeteer.launch({
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: { top: "15mm", right: "12mm", bottom: "15mm", left: "12mm" },
    });

    await browser.close();

    // cleanup docx
    try {
      if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
    } catch {}

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filenameBase}.pdf"`);
    return res.send(pdfBuffer);
  } catch (e) {
    console.log("PDF generation failed:", e);

    try {
      if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
    } catch {}

    return res.status(500).send("PDF generation failed. Check backend logs.");
  }
};

// FSIC Certificate
app.get("/records/:id/certificate/:type/pdf", async (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  const templateFile = req.params.type === "owner" ? "fsic-owner.docx" : "fsic-bfp.docx";
  return generatePDF(record, templateFile, `fsic-${record.id}`, res);
});

// IO / REINSPECTION / NFSI for records
app.get("/records/:id/:docType/pdf", async (req, res) => {
  const record = findRecordById(req.params.id);
  if (!record) return res.status(404).send("Record not found");

  let templateFile = "";
  if (req.params.docType === "io") templateFile = "officers.docx";
  else if (req.params.docType === "reinspection") templateFile = "reinspection.docx";
  else if (req.params.docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  return generatePDF(record, templateFile, `${req.params.docType}-${record.id}`, res);
});

// DOCUMENTS PDF generation
app.get("/documents/:id/:docType/pdf", async (req, res) => {
  const doc = (readJSON(DOCUMENTS_FILE) || []).find((r) => String(r.id) == String(req.params.id));
  if (!doc) return res.status(404).send("Document not found");

  let templateFile = "";
  if (req.params.docType === "io") templateFile = "officers.docx";
  else if (req.params.docType === "reinspection") templateFile = "reinspection.docx";
  else if (req.params.docType === "nfsi") templateFile = "nfsi-form.docx";
  else return res.status(400).send("Invalid type");

  return generatePDF(doc, templateFile, `doc-${req.params.docType}-${doc.id}`, res);
});

// -----------------------------
// START
// -----------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
});
