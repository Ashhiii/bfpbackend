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

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
});

// -----------------------------
// FIREBASE ADMIN
// -----------------------------
if (!admin.apps.length) {
  const projectId = process.env.FIREBASE_PROJECT_ID;
  const clientEmail = process.env.FIREBASE_CLIENT_EMAIL;
  const privateKey = (process.env.FIREBASE_PRIVATE_KEY || "").replace(
    /\\n/g,
    "\n"
  );

  if (!projectId || !clientEmail || !privateKey) {
    console.warn(
      "⚠️ Firebase Admin env vars missing. Firestore lookups will fail."
    );
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
// OPTIONAL JSON FILES
// -----------------------------
const DATA_FILE = path.join(__dirname, "records.json");
const ARCHIVE_FILE = path.join(__dirname, "archive.json");
const DOCUMENTS_FILE = path.join(__dirname, "documents.json");
const HISTORY_FILE = path.join(__dirname, "history.json");
const CLEARANCES_FILE = path.join(__dirname, "clearances.json");

const ensureFile = (file, defaultData) => {
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, JSON.stringify(defaultData, null, 2));
  }
};

ensureFile(DATA_FILE, []);
ensureFile(ARCHIVE_FILE, {});
ensureFile(DOCUMENTS_FILE, []);
ensureFile(HISTORY_FILE, []);
ensureFile(CLEARANCES_FILE, []);

const readJSON = (file) => JSON.parse(fs.readFileSync(file, "utf-8"));
const writeJSON = (file, data) =>
  fs.writeFileSync(file, JSON.stringify(data, null, 2));

// -----------------------------
// HELPERS
// -----------------------------
const normalize = (v) => String(v ?? "").trim();

const ensureEntityKey = (r) => {
  if (!r) return r;
  if (r.entityKey) return r;
  const fsic = normalize(r.fsicAppNo || r.FSIC_APP_NO || r.FSIC_NUMBER);
  return {
    ...r,
    entityKey: fsic ? `fsic:${fsic}` : `rec:${r.id || Date.now()}`,
  };
};

const normalizeEntityKey = (entityKey) => normalize(entityKey);

const toLongDate = (v) => {
  if (!v) return "";

  if (typeof v === "object" && typeof v.toDate === "function") {
    v = v.toDate();
  }

  let d = null;

  if (v instanceof Date) d = v;

  if (!d && typeof v === "string") {
    const s = v.trim();

    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    } else {
      const tmp = new Date(s);
      if (!Number.isNaN(tmp.getTime())) d = tmp;
    }
  }

  if (!d && typeof v === "number") {
    const tmp = new Date(v);
    if (!Number.isNaN(tmp.getTime())) d = tmp;
  }

  if (!d || Number.isNaN(d.getTime())) return String(v);

  const months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
};

const formatIssuedDay = (day) => {
  const n = Number(day);
  if (!Number.isFinite(n)) return String(day || "");

  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${n}th`;

  const mod10 = n % 10;
  if (mod10 === 1) return `${n}st`;
  if (mod10 === 2) return `${n}nd`;
  if (mod10 === 3) return `${n}rd`;
  return `${n}th`;
};


const pickAllowedRecordFields = (obj = {}) => {
  const syncedFsicNo = String(
    obj.fsicNo ?? obj.FSIC_NUMBER ?? obj.fsicNumber ?? ""
  ).trim();

  const syncedFsicAppNo = String(
    obj.fsicAppNo ?? obj.FSIC_APP_NO ?? ""
  ).trim();

  return {
    fsicNo: syncedFsicNo,
    FSIC_NUMBER: syncedFsicNo,

    fsicAppNo: syncedFsicAppNo,
    FSIC_APP_NO: syncedFsicAppNo,

    natureOfInspection: obj.natureOfInspection ?? obj.NATURE_OF_INSPECTION ?? "",
    ownerName: obj.ownerName ?? obj.OWNERS_NAME ?? "",
    establishmentName: obj.establishmentName ?? obj.ESTABLISHMENT_NAME ?? "",
    businessAddress:
      obj.businessAddress ?? obj.BUSSINESS_ADDRESS ?? obj.ADDRESS ?? "",
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
    buildingDesc:
      obj.buildingDesc ?? obj.BUILDING_DESC ?? obj.BLDG_DESCRIPTION ?? "",
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
    chiefPosition: obj.chiefPosition ?? obj.CHIEF_POSITION ?? "",
    marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
  };
};

const pickAllowedDocumentFields = (obj = {}) => {
  const syncedFsicNo = String(
    obj.fsicNo ?? obj.FSIC_NUMBER ?? obj.fsicNumber ?? ""
  ).trim();

  const syncedFsicAppNo = String(
    obj.fsicAppNo ?? obj.FSIC_APP_NO ?? ""
  ).trim();

  return {
    fsicNo: syncedFsicNo,
    FSIC_NUMBER: syncedFsicNo,

    fsicAppNo: syncedFsicAppNo,
    FSIC_APP_NO: syncedFsicAppNo,

    ownerName: obj.ownerName ?? obj.OWNERS_NAME ?? "",
    establishmentName: obj.establishmentName ?? obj.ESTABLISHMENT_NAME ?? "",
    businessAddress:
      obj.businessAddress ?? obj.BUSSINESS_ADDRESS ?? obj.ADDRESS ?? "",
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

    inspector4: obj.inspector4 ?? obj.INSPECTOR_4 ?? "",
    inspector4Serial: obj.inspector4Serial ?? obj.INSPECTOR_4_SERIAL ?? "",

    inspector5: obj.inspector5 ?? obj.INSPECTOR_5 ?? "",
    inspector5Serial: obj.inspector5Serial ?? obj.INSPECTOR_5_SERIAL ?? "",

    chiefName: obj.chiefName ?? obj.CHIEF ?? "",
    chiefPosition: obj.chiefPosition ?? obj.CHIEF_POSITION ?? "",
    marshalName: obj.marshalName ?? obj.MARSHAL ?? "",
  };
};

const makeId = () => {
  try {
    return crypto.randomUUID();
  } catch {
    return String(Date.now() + Math.floor(Math.random() * 100000));
  }
};

const pickAllowedClearanceFields = (obj = {}) => ({
  id: obj.id || makeId(),
  entityKey: normalizeEntityKey(obj.entityKey || ""),
  recordId: obj.recordId || "",
  type: String(
    obj.type || obj.clearanceType || obj.templateType || ""
  ).toLowerCase().trim(),

  FSIC_NUMBER: obj.FSIC_NUMBER ?? obj.fsicNumber ?? "",
  FSIC_APP_NO: obj.FSIC_APP_NO ?? obj.fsicAppNo ?? "",

  ownerName: obj.ownerName ?? obj.OWNERS_NAME ?? "",
  establishmentName:
    obj.establishmentName ??
    obj.ESTABLISHMENT_NAME ??
    obj.NAME_OF_BUILDING ??
    "",
  businessAddress:
    obj.businessAddress ?? obj.BUSSINESS_ADDRESS ?? obj.ADDRESS ?? "",
  contactNumber: obj.contactNumber ?? obj.CONTACT_NUMBER ?? "",

  // NEW COMMON CLEARANCE FIELDS
  clearanceDate: obj.clearanceDate ?? obj.CLEARANCE_DATE ?? "",
  clearanceValidity:
    obj.clearanceValidity ?? obj.CLEARANCE_VALIDITY ?? obj.validUntil ?? obj.VALID_UNTIL ?? "",
  validUntil:
    obj.validUntil ?? obj.VALID_UNTIL ?? obj.clearanceValidity ?? obj.CLEARANCE_VALIDITY ?? "",

  orNumber: obj.orNumber ?? obj.OR_NUMBER ?? "",
  orAmount: obj.orAmount ?? obj.OR_AMOUNT ?? obj.AMOUNT_PAID ?? "",
  orDate: obj.orDate ?? obj.OR_DATE ?? "",

  chiefName: obj.chiefName ?? obj.CHIEF ?? obj.CHIEF_FSES ?? "",
  chiefPosition: obj.chiefPosition ?? obj.CHIEF_POSITION ?? "",
  marshalName: obj.marshalName ?? obj.MARSHAL ?? obj.FIRE_MARSHAL ?? "",

  amountPaid: obj.amountPaid ?? obj.AMOUNT_PAID ?? obj.OR_AMOUNT ?? "",

  // Conveyance
  plateNumber: obj.plateNumber ?? obj.PLATE_NUMBER ?? "",
  typeOfVehicle:
    obj.typeOfVehicle ?? obj.vehicleType ?? obj.TYPE_OF_VEHICLE ?? "",
  vehicleType:
    obj.vehicleType ?? obj.typeOfVehicle ?? obj.TYPE_OF_VEHICLE ?? "",
  chassisNumber: obj.chassisNumber ?? obj.CHASSIS_NUMBER ?? "",
  motorNumber: obj.motorNumber ?? obj.MOTOR_NUMBER ?? "",
  licenseNumber: obj.licenseNumber ?? obj.LICENSE_NUMBER ?? "",
  nameOfDriver:
    obj.nameOfDriver ?? obj.driverName ?? obj.NAME_OF_DRIVER ?? "",
  driverName:
    obj.driverName ?? obj.nameOfDriver ?? obj.NAME_OF_DRIVER ?? "",
  trailerNumber: obj.trailerNumber ?? obj.TRAILER_NUMBER ?? "",
  capacity: obj.capacity ?? obj.CAPACITY ?? "",

  // Storage
  storageAddress: obj.storageAddress ?? obj.STORAGE_ADDRESS ?? "",
  flammable1: obj.flammable1 ?? obj.FLAMMABLE_1 ?? "",
  capacity1: obj.capacity1 ?? obj.CAPACITY_1 ?? "",
  flammable2: obj.flammable2 ?? obj.FLAMMABLE_2 ?? "",
  capacity2: obj.capacity2 ?? obj.CAPACITY_2 ?? "",
  flammable3: obj.flammable3 ?? obj.FLAMMABLE_3 ?? "",
  capacity3: obj.capacity3 ?? obj.CAPACITY_3 ?? "",
  flammable4: obj.flammable4 ?? obj.FLAMMABLE_4 ?? "",
  capacity4: obj.capacity4 ?? obj.CAPACITY_4 ?? "",

  // Hot Works
  companyName: obj.companyName ?? obj.COMPANY_NAME ?? "",
  jobOrderNumber: obj.jobOrderNumber ?? obj.JOB_ORDER_NUMBER ?? "",
  natureOfJob: obj.natureOfJob ?? obj.NATURE_OF_JOB ?? "",
  permitAuthorizingIndividual:
    obj.permitAuthorizingIndividual ??
    obj.PERMIT_AUTHORIZING_INDIVIDUAL ??
    "",
  hotworkOperator:
    obj.hotworkOperator ?? obj.hotWorkOperator ?? obj.HOTWORK_OPERATOR ?? "",
  hotWorkOperator:
    obj.hotWorkOperator ?? obj.hotworkOperator ?? obj.HOTWORK_OPERATOR ?? "",
  fireWatch: obj.fireWatch ?? obj.fireWatchName ?? obj.FIRE_WATCH ?? "",
  fireWatchName: obj.fireWatchName ?? obj.fireWatch ?? obj.FIRE_WATCH ?? "",

  // Fire Drill
  dateConducted:
    obj.dateConducted ?? obj.fireDrillDate ?? obj.DATE_CONDUCTED ?? "",
  fireDrillDate:
    obj.fireDrillDate ?? obj.dateConducted ?? obj.DATE_CONDUCTED ?? "",
  issuedDay: obj.issuedDay ?? obj.ISSUED_DAY ?? "",
  issuedMonth: obj.issuedMonth ?? obj.ISSUED_MONTH ?? "",

// Fumigation
operatorName: obj.operatorName ?? obj.OPERATOR_NAME ?? "",
operationTime: obj.operationTime ?? obj.OPERATION_TIME ?? "",
operationDate: obj.operationDate ?? obj.OPERATION_DATE ?? "",
operationDuration: obj.operationDuration ?? obj.OPERATION_DURATION ?? "",
foggingAddress: obj.foggingAddress ?? obj.FOGGING_ADDRESS ?? "",
conductedBy: obj.conductedBy ?? obj.CONDUCTED_BY ?? "",

//Seminar
    fireDrillDate:
      obj.fireDrillDate ??
      obj.dateConducted ??
      obj.DATE_CONDUCTED ??
      "",
    issuedDay: obj.issuedDay ?? obj.ISSUED_DAY ?? "",
    issuedMonth: obj.issuedMonth ?? obj.ISSUED_MONTH ?? "",
    
    //Fire Safety
    plateNumber: obj.plateNumber ?? obj.PLATE_NUMBER ?? "",
    vehicleType: obj.vehicleType ?? obj.TYPE_OF_VEHICLE ?? "",
    brandOfVehicle: obj.brandOfVehicle ?? obj.BRAND_OF_VEHICLE ?? "",
    engineNumber: obj.engineNumber ?? obj.ENGINE_NUMBER ?? "",
    chassisNumber: obj.chassisNumber ?? obj.CHASSIS_NUMBER ?? "",
    permitNumber: obj.permitNumber ?? obj.PERMIT_NUMBER ?? "",
    fsicIssued: obj.fsicIssued ?? obj.FSIC_ISSUED ?? "",
    cageSize: obj.cageSize ?? obj.CAGE_SIZE ?? "",
    capacity: obj.capacity ?? obj.CAPACITY ?? "",


  createdAt: obj.createdAt || new Date().toISOString(),
});

// -----------------------------
// FIRESTORE LOOKUPS
// -----------------------------
const findRecordById = async (id) => {
  if (!fdb) return null;

  const snap = await fdb.collection("records").doc(String(id)).get();
  if (snap.exists) return ensureEntityKey({ id: snap.id, ...snap.data() });

  const archiveCollections = ["archive", "archives"];

  for (const colName of archiveCollections) {
    try {
      const monthsSnap = await fdb.collection(colName).get();

      for (const m of monthsSnap.docs) {
        const recSnap = await fdb
          .collection(colName)
          .doc(m.id)
          .collection("records")
          .doc(String(id))
          .get();

        if (recSnap.exists) {
          return ensureEntityKey({ id: recSnap.id, ...recSnap.data() });
        }
      }
    } catch {
      // skip if collection does not exist
    }
  }

  return null;
};

const findDocumentById = async (id) => {
  if (!fdb) return null;
  const snap = await fdb.collection("documents").doc(String(id)).get();
  if (snap.exists) return { id: snap.id, ...snap.data() };
  return null;
};

const findClearanceById = async (id) => {
  if (!fdb) return null;

  const snap = await fdb.collection("clearances").doc(String(id)).get();
  if (!snap.exists) return null;

  return { id: snap.id, ...snap.data() };
};

const getAllClearances = async () => {
  if (!fdb) return [];

  const snap = await fdb.collection("clearances").get();
  return snap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
};

// -----------------------------
// PDF
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
    return res
      .status(404)
      .send(`Template not found: ${templateFile} (path=${templatePath})`);
  }

  const soffice = findSoffice();
  if (!soffice) {
    return res
      .status(500)
      .send("LibreOffice not found (soffice). Use Docker install libreoffice.");
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
      // record/document fields
FSIC_NUMBER:
    record.fsicNo ||
    record.FSIC_NUMBER ||
    record.fsicNumber ||
    "",

  FSIC_APP_NO:
    record.fsicAppNo ||
    record.FSIC_APP_NO ||
    "",
    
      DATE_INSPECTED: toLongDate(
        record.DATE_INSPECTED || record.dateInspected || ""
      ),

      NAME_OF_ESTABLISHMENT:
        record.NAME_OF_ESTABLISHMENT ||
        record.ESTABLISHMENT_NAME ||
        record.establishmentName ||
        "",

      NAME_OF_OWNER:
        record.NAME_OF_OWNER || record.OWNERS_NAME || record.ownerName || "",

      ADDRESS:
        record.ADDRESS || record.BUSSINESS_ADDRESS || record.businessAddress || "",

      FLOOR_AREA: record.FLOOR_AREA || record.floorArea || "",
      BLDG_DESCRIPTION:
        record.BLDG_DESCRIPTION || record.BUILDING_DESC || record.buildingDesc || "",
      FSIC_VALIDITY: toLongDate(
        record.FSIC_VALIDITY || record.fsicValidity || ""
      ),
      OR_NUMBER: record.OR_NUMBER || record.orNumber || "",
      OR_DATE: toLongDate(record.OR_DATE || record.orDate || ""),
      OR_AMOUNT: record.OR_AMOUNT || record.orAmount || "",

      IO_NUMBER: record.IO_NUMBER || record.ioNumber || "",
      IO_DATE: toLongDate(record.IO_DATE || record.ioDate || ""),
      TAXPAYER: record.TAXPAYER || record.OWNERS_NAME || record.ownerName || "",
      TRADE_NAME:
        record.TRADE_NAME ||
        record.ESTABLISHMENT_NAME ||
        record.establishmentName ||
        "",
      CONTACT_:
        record.CONTACT_ || record.CONTACT_NUMBER || record.contactNumber || "",

      NFSI_NUMBER: record.NFSI_NUMBER || record.nfsiNumber || "",
      NFSI_DATE: toLongDate(record.NFSI_DATE || record.nfsiDate || ""),

      NTC_NUMBER: record.NTC_NUMBER || record.ntcNumber || "",
      NTC_DATE: toLongDate(record.NTC_DATE || record.ntcDate || ""),

      OWNER: record.OWNER || record.OWNERS_NAME || record.ownerName || "",
      TEAM_LEADER: record.teamLeader || record.TEAM_LEADER || "",
      TEAM_LEADER_SERIAL:
        record.teamLeaderSerial || record.TEAM_LEADER_SERIAL || "",
      INSPECTORS: record.INSPECTORS || record.inspectors || "",

      INSPECTOR_1: record.inspector1 || record.INSPECTOR_1 || "",
      INSPECTOR_1_SERIAL:
        record.inspector1Serial || record.INSPECTOR_1_SERIAL || "",

      INSPECTOR_2: record.inspector2 || record.INSPECTOR_2 || "",
      INSPECTOR_2_SERIAL:
        record.inspector2Serial || record.INSPECTOR_2_SERIAL || "",

      INSPECTOR_3: record.inspector3 || record.INSPECTOR_3 || "",
      INSPECTOR_3_SERIAL:
        record.inspector3Serial || record.INSPECTOR_3_SERIAL || "",

      INSPECTOR_4: record.inspector4 || record.INSPECTOR_4 || "",
      INSPECTOR_4_SERIAL:
        record.inspector4Serial || record.INSPECTOR_4_SERIAL || "",

      INSPECTOR_5: record.inspector5 || record.INSPECTOR_5 || "",
      INSPECTOR_5_SERIAL:
        record.inspector5Serial || record.INSPECTOR_5_SERIAL || "",

      DATE: toLongDate(new Date()),
      CHIEF: record.CHIEF || record.chiefName || "",
      CHIEF_POSITION: record.CHIEF_POSITION || record.chiefPosition || "",
      MARSHAL: record.MARSHAL || record.marshalName || "",

      // clearance fields
      NAME_OF_BUILDING:
        record.NAME_OF_BUILDING ||
        record.establishmentName ||
        record.ESTABLISHMENT_NAME ||
        "",

      AMOUNT_PAID:
        record.AMOUNT_PAID ||
        record.amountPaid ||
        record.OR_AMOUNT ||
        record.orAmount ||
        "",

      CHIEF_FSES: record.CHIEF_FSES || record.chiefName || "",
      FIRE_MARSHAL: record.FIRE_MARSHAL || record.marshalName || "",

      PLATE_NUMBER: record.PLATE_NUMBER || record.plateNumber || "",
      TYPE_OF_VEHICLE: record.TYPE_OF_VEHICLE || record.typeOfVehicle || "",
      CHASSIS_NUMBER: record.CHASSIS_NUMBER || record.chassisNumber || "",
      MOTOR_NUMBER: record.MOTOR_NUMBER || record.motorNumber || "",
      LICENSE_NUMBER: record.LICENSE_NUMBER || record.licenseNumber || "",
      NAME_OF_DRIVER: record.NAME_OF_DRIVER || record.nameOfDriver || "",
      TRAILER_NUMBER: record.TRAILER_NUMBER || record.trailerNumber || "",
      CAPACITY: record.CAPACITY || record.capacity || "",

      FLAMMABLE_1: record.FLAMMABLE_1 || record.flammable1 || "",
      CAPACITY_1: record.CAPACITY_1 || record.capacity1 || "",
      FLAMMABLE_2: record.FLAMMABLE_2 || record.flammable2 || "",
      CAPACITY_2: record.CAPACITY_2 || record.capacity2 || "",
      FLAMMABLE_3: record.FLAMMABLE_3 || record.flammable3 || "",
      CAPACITY_3: record.CAPACITY_3 || record.capacity3 || "",
      FLAMMABLE_4: record.FLAMMABLE_4 || record.flammable4 || "",
      CAPACITY_4: record.CAPACITY_4 || record.capacity4 || "",

      COMPANY_NAME: record.COMPANY_NAME || record.companyName || "",
      JOB_ORDER_NUMBER: record.JOB_ORDER_NUMBER || record.jobOrderNumber || "",
      NATURE_OF_JOB: record.NATURE_OF_JOB || record.natureOfJob || "",
      PERMIT_AUTHORIZING_INDIVIDUAL:
        record.PERMIT_AUTHORIZING_INDIVIDUAL ||
        record.permitAuthorizingIndividual ||
        "",
      HOTWORK_OPERATOR: record.HOTWORK_OPERATOR || record.hotworkOperator || "",
      FIRE_WATCH: record.FIRE_WATCH || record.fireWatch || "",

      DATE_CONDUCTED: toLongDate(
        record.DATE_CONDUCTED || record.dateConducted || ""
      ),

      OPERATOR_NAME: record.OPERATOR_NAME || record.operatorName || "",
      OPERATION_TIME: record.OPERATION_TIME || record.operationTime || "",
      OPERATION_DATE: toLongDate(
        record.OPERATION_DATE || record.operationDate || ""
      ),
      VALID_UNTIL: toLongDate(record.VALID_UNTIL || record.validUntil || ""),
      CONTROL_NUMBER: record.CONTROL_NUMBER || record.controlNumber || "",
      CLEARANCE_DATE: toLongDate(
        record.CLEARANCE_DATE || record.clearanceDate || ""
      ),
      CLEARANCE_VALIDITY: toLongDate(
        record.CLEARANCE_VALIDITY || record.clearanceValidity || record.VALID_UNTIL || record.validUntil || ""
      ),
      ISSUED_DAY: formatIssuedDay(record.ISSUED_DAY || record.issuedDay || ""),
      ISSUED_MONTH: record.ISSUED_MONTH || record.issuedMonth || "",
      STORAGE_ADDRESS: record.STORAGE_ADDRESS || record.storageAddress || "",
      OPERATION_DURATION: record.OPERATION_DURATION || record.operationDuration || "",
      FOGGING_ADDRESS:
  record.FOGGING_ADDRESS || record.foggingAddress || record.fogging_address || "",
CONDUCTED_BY:
  record.CONDUCTED_BY || record.conductedBy || record.conducted_by || "",

  BRAND_OF_VEHICLE:
  record.BRAND_OF_VEHICLE || record.brandOfVehicle || "",

ENGINE_NUMBER:
  record.ENGINE_NUMBER || record.engineNumber || "",

PERMIT_NUMBER:
  record.PERMIT_NUMBER || record.permitNumber || "",

FSIC_ISSUED:
  record.FSIC_ISSUED || record.fsicIssued || "",

CAGE_SIZE:
  record.CAGE_SIZE || record.cageSize || "",

  OR_NUMBER:
  record.OR_NUMBER || record.orNumber || "",

OR_AMOUNT:
  record.OR_AMOUNT || record.orAmount || record.amountPaid || "",

OR_DATE:
  toLongDate(record.OR_DATE || record.orDate || ""),
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
        return res
          .status(500)
          .send(
            `PDF conversion failed. ${String(stderr || err?.message || err)}`
          );
      }

      const expectedPdf = outputDocx.replace(/\.docx$/i, ".pdf");
      if (!fs.existsSync(expectedPdf)) {
        try {
          if (fs.existsSync(outputDocx)) fs.unlinkSync(outputDocx);
        } catch {}
        return res.status(500).send("PDF file not produced after conversion.");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${filenameBase}.pdf"`
      );

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
    return res
      .status(500)
      .send(`PDF generation failed (templater). ${e.message}`);
  }
};

// -----------------------------
// HEALTH
// -----------------------------
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
  for (const k of Object.keys(row)) {
    headerMap[normHeader(k)] = row[k];
  }

  const get = (...variants) => pickAny(headerMap, variants.map(normHeader));

  const rec = {
    id: makeId(),
    createdAt: new Date().toISOString(),

    appno: String(
      get("appno", "applicationno", "application#", "applicationnumber") || ""
    ),
    fsicAppNo: String(
      get("fsicappno", "fsicno", "fsicnumber", "fsicapp#", "fsicapp") || ""
    ),
    natureOfInspection: String(
      get("natureofinspection", "inspection", "nature") || ""
    ),
    ownerName: String(get("owner", "ownername", "ownersname", "taxpayer") || ""),
    establishmentName: String(
      get("establishment", "establishmentname", "tradename", "nameofestablishment") ||
        ""
    ),
    businessAddress: String(
      get("businessaddress", "address", "bussinessaddress") || ""
    ),
    contactNumber: String(get("contactnumber", "contact", "mobile") || ""),
    dateInspected: excelDateToISO(get("dateinspected", "date") || ""),

    ioNumber: String(get("ionumber", "io#", "io") || ""),
    ioDate: excelDateToISO(get("iodate") || ""),

    nfsiNumber: String(get("nfsinumber", "nfsi#", "nfsi") || ""),
    nfsiDate: excelDateToISO(get("nfsidate") || ""),

    ntcNumber: String(get("ntcnumber", "ntc#", "ntc") || ""),
    ntcDate: excelDateToISO(get("ntcdate") || ""),

    fsicValidity: String(get("fsicvalidity", "validity") || ""),
    defects: String(get("defects", "violations") || ""),
    inspectors: String(get("inspectors", "inspector") || ""),

    chiefName: String(get("chiefname", "chief") || ""),
    chiefPosition: String(get("chiefposition", "position") || ""),
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
// IMPORT
// -----------------------------
app.post("/import/records", upload.single("file"), (req, res) => {
  try {
    const { rows, sheetName, error } = readExcelRows(req);
    if (error) return res.status(400).json({ success: false, message: error });

    const current = (readJSON(DATA_FILE) || []).map(ensureEntityKey);
    const mapped = rows.map(mapExcelRowToRecord);
    const toAdd = mapped.filter(
      (r) => normalize(r.fsicAppNo) && normalize(r.ownerName)
    );

    writeJSON(DATA_FILE, [...current, ...toAdd]);

    res.json({
      success: true,
      imported: toAdd.length,
      skipped: mapped.length - toAdd.length,
      sheet: sheetName,
    });
  } catch (e) {
    console.error("POST /import/records error:", e);
    res
      .status(500)
      .json({ success: false, message: "Failed to import records Excel." });
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
// CLEARANCES ROUTES
// -----------------------------
app.post("/clearances", async (req, res) => {
  try {
    if (!fdb) {
      return res.status(500).json({
        success: false,
        message: "Firestore is not initialized.",
      });
    }

    let payload = pickAllowedClearanceFields(req.body);

    if (!payload.type) {
      return res.status(400).json({
        success: false,
        message: "Clearance type is required.",
      });
    }

    if (payload.recordId && !payload.entityKey) {
      const record = await findRecordById(payload.recordId);

      if (record?.entityKey) {
        payload.entityKey = record.entityKey;
      }

      if (!payload.FSIC_APP_NO && record?.FSIC_APP_NO) {
        payload.FSIC_APP_NO = record.FSIC_APP_NO;
      }

      if (!payload.FSIC_NUMBER && record?.FSIC_NUMBER) {
        payload.FSIC_NUMBER = record.FSIC_NUMBER;
      }

      if (!payload.ownerName && (record?.ownerName || record?.OWNERS_NAME)) {
        payload.ownerName = record.ownerName || record.OWNERS_NAME || "";
      }

      if (
        !payload.establishmentName &&
        (record?.establishmentName || record?.ESTABLISHMENT_NAME)
      ) {
        payload.establishmentName =
          record.establishmentName || record.ESTABLISHMENT_NAME || "";
      }

      if (
        !payload.businessAddress &&
        (record?.businessAddress || record?.BUSSINESS_ADDRESS || record?.ADDRESS)
      ) {
        payload.businessAddress =
          record.businessAddress ||
          record.BUSSINESS_ADDRESS ||
          record.ADDRESS ||
          "";
      }

      if (
        !payload.contactNumber &&
        (record?.contactNumber || record?.CONTACT_NUMBER)
      ) {
        payload.contactNumber = record.contactNumber || record.CONTACT_NUMBER || "";
      }
    }

    await fdb.collection("clearances").doc(String(payload.id)).set(payload);

    res.json({ success: true, data: payload });
  } catch (e) {
    console.error("POST /clearances error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to save clearance.",
    });
  }
});

app.get("/clearances", async (req, res) => {
  try {
    if (!fdb) {
      return res.status(500).json({
        success: false,
        message: "Firestore is not initialized.",
      });
    }

    const snap = await fdb.collection("clearances").orderBy("createdAt", "desc").get();
    const items = snap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));

    res.json(items);
  } catch (e) {
    console.error("GET /clearances error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to fetch clearances.",
    });
  }
});

app.get("/clearances/:id", async (req, res) => {
  try {
    const item = await findClearanceById(req.params.id);

    if (!item) {
      return res.status(404).json({
        success: false,
        message: "Clearance not found.",
      });
    }

    res.json(item);
  } catch (e) {
    console.error("GET /clearances/:id error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to fetch clearance.",
    });
  }
});

app.put("/clearances/:id", async (req, res) => {
  try {
    if (!fdb) {
      return res.status(500).json({
        success: false,
        message: "Firestore is not initialized.",
      });
    }

    const existing = await findClearanceById(req.params.id);

    if (!existing) {
      return res.status(404).json({
        success: false,
        message: "Clearance not found.",
      });
    }

    const merged = {
      ...existing,
      ...pickAllowedClearanceFields({
        ...existing,
        ...req.body,
        id: existing.id,
        createdAt: existing.createdAt,
      }),
    };

    await fdb.collection("clearances").doc(String(existing.id)).set(merged);

    res.json({ success: true, data: merged });
  } catch (e) {
    console.error("PUT /clearances/:id error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to update clearance.",
    });
  }
});

app.delete("/clearances/:id", async (req, res) => {
  try {
    if (!fdb) {
      return res.status(500).json({
        success: false,
        message: "Firestore is not initialized.",
      });
    }

    const existing = await findClearanceById(req.params.id);

    if (!existing) {
      return res.status(404).json({
        success: false,
        message: "Clearance not found.",
      });
    }

    await fdb.collection("clearances").doc(String(req.params.id)).delete();

    res.json({ success: true });
  } catch (e) {
    console.error("DELETE /clearances/:id error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to delete clearance.",
    });
  }
});

app.get("/records/:id/clearances", async (req, res) => {
  try {
    if (!fdb) {
      return res.status(500).json({
        success: false,
        message: "Firestore is not initialized.",
      });
    }

    const record = await findRecordById(req.params.id);
    if (!record) {
      return res.status(404).json({
        success: false,
        message: "Record not found",
      });
    }

    const entityKey = normalize(record.entityKey);

    const snap = await fdb.collection("clearances").get();
    const items = snap.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter(
        (c) =>
          normalize(c.entityKey) === entityKey ||
          String(c.recordId) === String(record.id)
      );

    res.json(items);
  } catch (e) {
    console.error("GET /records/:id/clearances error:", e);
    res.status(500).json({
      success: false,
      message: "Failed to fetch record clearances.",
    });
  }
});

app.get("/clearances/:id/certificate/:type/pdf", async (req, res) => {
  try {
    console.log("PDF REQUEST ID:", req.params.id);
    console.log("PDF REQUEST TYPE:", req.params.type);

    const clearance = await findClearanceById(req.params.id);
    if (!clearance) {
      console.log("Clearance not found in Firestore");
      return res.status(404).send("Clearance not found");
    }

    const type = String(req.params.type || "").toLowerCase().trim();

    let templateFile = "";
    if (type === "conveyance") templateFile = "FSED-38F-Conveyance.docx";
    else if (type === "storage") templateFile = "FSED-37F-Storage.docx";
    else if (type === "hotworks") templateFile = "FSED-34F-Hot-Works.docx";
    else if (type === "firedrill") templateFile = "FSED-44F-Fire-Drill-Rev02.docx";
    else if (type === "fumigation") templateFile = "FSED-41F-Fumigation.docx";
    else if (type === "seminar") templateFile = "Seminar.docx";
    else if (type === "firesafety") templateFile = "Fire-Safety.docx";

    else return res.status(400).send("Invalid clearance certificate type");

    console.log("USING TEMPLATE:", templateFile);

    generatePDF(
      clearance,
      templateFile,
      `clearance-${type}-${clearance.id}`,
      res
    );
  } catch (e) {
    console.error("GET /clearances/:id/certificate/:type/pdf error:", e);
    res.status(500).send("Failed to generate clearance PDF.");
  }
});

app.get("/clearances/:id/pdf", async (req, res) => {
  try {
    const clearance = await findClearanceById(req.params.id);
    if (!clearance) return res.status(404).send("Clearance not found");

    const type = String(clearance.type || "").toLowerCase().trim();

    let templateFile = "";
    if (type === "conveyance") {
      templateFile = "FSED-38F-Conveyance.docx";
    } else if (type === "storage") {
      templateFile = "FSED-37F-Storage.docx";
    } else if (type === "hotworks") {
      templateFile = "FSED-34F-Hot-Works.docx";
    } else if (type === "firedrill") {
      templateFile = "FSED-44F-Fire-Drill-Rev02.docx";
    } else if (type === "fumigation") {
      templateFile = "FSED-41F-Fumigation.docx";
    } else if (type === "seminar") {
      templateFile = "FSED-Seminar.docx";
    } else if (type === "firesafety") {
      templateFile = "Fire-Safety.docx";

    } else {
      return res.status(400).send("Invalid clearance type");
    }

    generatePDF(
      clearance,
      templateFile,
      `clearance-${type}-${clearance.id}`,
      res
    );
  } catch (e) {
    console.error("GET /clearances/:id/pdf error:", e);
    res.status(500).send("Failed to generate clearance PDF.");
  }
});

// -----------------------------
// PDF ROUTES
// -----------------------------
app.get("/records/:id/certificate/:type/pdf", async (req, res) => {
  try {
    const record = await findRecordById(req.params.id);
    if (!record) return res.status(404).send("Record not found");

    const type = String(req.params.type || "").toLowerCase();

    let templateFile = "";
    if (type === "owner") {
      templateFile = "fsic-owner.docx";
    } else if (type === "bfp") {
      templateFile = "fsic-bfp.docx";
    } else if (type === "owner-new") {
      templateFile = "fsic-owner-new.docx";
    } else if (type === "bfp-new") {
      templateFile = "fsic-bfp-new.docx";
    } else {
      return res.status(400).send("Invalid certificate type");
    }

    generatePDF(record, templateFile, `fsic-${type}-${record.id}`, res);
  } catch (e) {
    console.error("GET /records/:id/certificate/:type/pdf error:", e);
    res.status(500).send("Failed to generate certificate PDF.");
  }
});

app.get("/records/:id/:docType/pdf", async (req, res) => {
  try {
    const record = await findRecordById(req.params.id);
    if (!record) return res.status(404).send("Record not found");

    const dt = String(req.params.docType || "").toLowerCase();

    let templateFile = "";
    if (dt === "io") templateFile = "officers.docx";
    else if (dt === "reinspection") templateFile = "reinspection.docx";
    else if (dt === "nfsi") templateFile = "nfsi-form.docx";
    else return res.status(400).send("Invalid type");

    generatePDF(record, templateFile, `${dt}-${record.id}`, res);
  } catch (e) {
    console.error("GET /records/:id/:docType/pdf error:", e);
    res.status(500).send("Failed to generate document PDF.");
  }
});

app.get("/documents/:id/:docType/pdf", async (req, res) => {
  try {
    const docu = await findDocumentById(req.params.id);
    if (!docu) return res.status(404).send("Document not found");

    const dt = String(req.params.docType || "").toLowerCase();

    let templateFile = "";
    if (dt === "io") templateFile = "officers.docx";
    else if (dt === "reinspection") templateFile = "reinspection.docx";
    else if (dt === "nfsi") templateFile = "nfsi-form.docx";
    else return res.status(400).send("Invalid type");

    generatePDF(docu, templateFile, `doc-${dt}-${docu.id}`, res);
  } catch (e) {
    console.error("GET /documents/:id/:docType/pdf error:", e);
    res.status(500).send("Failed to generate document PDF.");
  }
});

// -----------------------------
// ROOT
// -----------------------------
app.get("/", (req, res) => {
  res.send("✅ BFP Backend Running (Firestore Option B)");
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend running on port ${PORT}`);
  console.log("PIN:", process.env.PIN ? "(set)" : "(default 1234)");
  console.log("SOFFICE_PATH:", process.env.SOFFICE_PATH || "(not set)");
  console.log("FIREBASE:", fdb ? "connected (check /health)" : "NOT initialized");
});