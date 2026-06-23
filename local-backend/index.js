require("dotenv").config({ path: require("path").join(__dirname, ".env") });

const express = require("express");
const cors = require("cors");
const sql = require("mssql");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
} = require("docx");

const rateLimit = require("express-rate-limit");

const app = express();

// ── CORS — only allow the invoice frontend and local dev ───────────────────
app.use(cors({
  origin: [
    "https://okan.invoice.baikgroup.com",
    "http://localhost:3000",
    "http://localhost:3001",
  ],
}));
app.use(express.json());

// ── Rate limiting — 60 requests per IP per minute across all endpoints ─────
app.use(rateLimit({
  windowMs: 60 * 1000,
  max: 60,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many requests. Try again in a minute." },
}));

// ── API key guard — all sensitive endpoints require X-API-Key header ────────
function requireApiKey(req, res, next) {
  if (req.headers["x-api-key"] !== process.env.API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

// ── DB config — all values come from the .env file on this machine ──────────
// Never hardcode credentials here. See .env.example for required keys.
const dbConfig = {
  server:   process.env.DB_SERVER,
  port:     parseInt(process.env.DB_PORT || "1433", 10),
  user:     process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
  options: { trustServerCertificate: true, encrypt: false },
};

// Fail fast at startup if any required env var is missing
const REQUIRED_ENV = ["DB_SERVER", "DB_USER", "DB_PASSWORD", "DB_NAME", "ADMIN_PASS", "API_KEY"];
const missing = REQUIRED_ENV.filter(k => !process.env[k]);
if (missing.length) {
  console.error(`\n[FATAL] Missing required environment variables: ${missing.join(", ")}`);
  console.error("Create a .env file in local-backend/ — see .env.example for the required keys.\n");
  process.exit(1);
}

let pool;
async function getPool() {
  if (!pool) pool = await sql.connect(dbConfig);
  return pool;
}

// ── Admin password — loaded from .env, never hardcoded ────────────────────
const ADMIN_PASS = process.env.ADMIN_PASS;

// ── Admin: lookup student (returns full raw row) ───────────────────────────
app.post("/admin-lookup", requireApiKey, async (req, res) => {
  const { password, studentNo } = req.body;
  if (!password || password !== ADMIN_PASS) {
    return res.status(401).json({ error: "Unauthorized: wrong password" });
  }
  if (!studentNo) {
    return res.status(400).json({ error: "studentNo is required" });
  }
  try {
    const db = await getPool();
    const result = await db.request()
      .input("id", sql.VarChar, String(studentNo).trim())
      .query(`SELECT TOP 1 HesapKodu, Unvan1, Bolum, Fakulte, EgitimYil, EgitimUcreti
              FROM TEMP_OgrenciKayitListesi_2
              WHERE HesapKodu = @id`);
    if (result.recordset.length === 0) {
      return res.status(404).json({ error: "Student not found" });
    }
    res.json(result.recordset[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Admin: update all editable fields for a student ────────────────────────
app.post("/admin-update-all", requireApiKey, async (req, res) => {
  const { password, studentNo, Unvan1, Bolum, Fakulte, EgitimYil, EgitimUcreti } = req.body;

  if (!password || password !== ADMIN_PASS) {
    return res.status(401).json({ error: "Unauthorized: wrong password" });
  }
  if (!studentNo) {
    return res.status(400).json({ error: "studentNo is required" });
  }

  try {
    const db = await getPool();
    const result = await db.request()
      // WHERE key — identifies the row, never updated
      .input("studentNo",    sql.VarChar(50),      String(studentNo).trim())
      // Fields being updated — all parameterized, no string concatenation
      .input("Unvan1",       sql.NVarChar(255),    Unvan1       ?? null)
      .input("Bolum",        sql.NVarChar(255),    Bolum        ?? null)
      .input("Fakulte",      sql.NVarChar(255),    Fakulte      ?? null)
      .input("EgitimYil",    sql.NVarChar(50),     EgitimYil    ?? null)
      .input("EgitimUcreti", sql.Decimal(18, 2),   EgitimUcreti != null ? parseFloat(EgitimUcreti) : null)
      .query(`
        UPDATE TEMP_OgrenciKayitListesi_2
        SET
          Unvan1       = @Unvan1,
          Bolum        = @Bolum,
          Fakulte      = @Fakulte,
          EgitimYil    = @EgitimYil,
          EgitimUcreti = @EgitimUcreti
        WHERE HesapKodu = @studentNo
      `);

    res.json({ ok: true, rowsAffected: result.rowsAffected[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Student lookup ─────────────────────────────────────────────────────────
app.get("/student/:id", requireApiKey, async (req, res) => {
  try {
    const db = await getPool();

    // Search all TEMP_OgrenciKayitListesi tables that have data for this student number
    // Use FORNET user with most data (FORNET kisiID=2 has 20268 rows — most complete)
    const tableName = `TEMP_OgrenciKayitListesi_2`;

    const result = await db.request()
      .input("id", sql.VarChar, req.params.id)
      .query(`SELECT TOP 1 HesapKodu, Unvan1, Bolum, Fakulte, EgitimYil, EgitimUcreti
              FROM ${tableName}
              WHERE HesapKodu = @id`);

    if (result.recordset.length === 0) {
      return res.status(404).json({ error: "Student not found" });
    }

    const row = result.recordset[0];
    res.json({
      studentNo: row.HesapKodu,
      fullName: row.Unvan1,
      program: row.Bolum,
      faculty: row.Fakulte,
      acYear: row.EgitimYil,
      tuition: row.EgitimUcreti,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Debug: check user KisiID and TEMP table row count ─────────────────────
app.get("/debug", requireApiKey, async (req, res) => {
  try {
    const db = await getPool();

    // Get all users
    const users = await db.request().query(`SELECT LoginName, KisiID FROM FORNET_Kullanici`);

    // For each user, check if their TEMP table has rows
    const results = [];
    for (const u of users.recordset) {
      try {
        const count = await db.request().query(
          `SELECT COUNT(*) as cnt FROM TEMP_OgrenciKayitListesi_${u.KisiID}`
        );
        results.push({ login: u.LoginName, kisiID: u.KisiID, rows: count.recordset[0].cnt });
      } catch (e) {
        results.push({ login: u.LoginName, kisiID: u.KisiID, rows: "table missing" });
      }
    }
    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Schema discovery (use once to find correct table/column names) ──────────
app.get("/schema", requireApiKey, async (req, res) => {
  try {
    const db = await getPool();
    const tables = await db.request().query(`
      SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
      WHERE TABLE_TYPE = 'BASE TABLE'
      ORDER BY TABLE_NAME
    `);
    res.json(tables.recordset.map(r => r.TABLE_NAME));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/schema/:table", requireApiKey, async (req, res) => {
  try {
    const db = await getPool();
    const cols = await db.request()
      .input("t", sql.VarChar, req.params.table)
      .query(`
        SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @t ORDER BY ORDINAL_POSITION
      `);
    res.json(cols.recordset);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Invoice generation (same as Render backend) ────────────────────────────
function fmtAmt(n) {
  return parseFloat(n).toLocaleString("tr-TR", {
    minimumFractionDigits: 2, maximumFractionDigits: 2,
  });
}

function sp(n) {
  return new Paragraph({ children: [new TextRun("")], spacing: { after: n } });
}

function buildDoc({ firstName, lastName, date, acYear, invAmount, totalAmount, program, invNo, descType, customDesc, feeType }) {
  const fullName = `${firstName} ${lastName}`;
  const acYearStr = acYear || "2024-2025";
  const d = new Date(date);
  const dateStr = `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
  const invAmt = fmtAmt(invAmount);
  const totAmt = totalAmount ? fmtAmt(totalAmount) : null;

  const feeLabels = { tuition: "tuition fee", dormitory: "dormitory fee", summer: "summer course fee" };
  const feeText = feeLabels[feeType] || "tuition fee";

  let descText;
  if (customDesc && customDesc.trim()) {
    descText = customDesc.trim();
  } else if (descType === "debt") {
    descText = `${program} program at İstanbul Okan University ${acYearStr} tuition debt payment`;
  } else if (descType === "dorm") {
    descText = `${program} program at İstanbul Okan University ${acYearStr} dormitory fee`;
  } else {
    descText = `${program} program at İstanbul Okan University ${acYearStr} registration fee`;
  }

  // ---- styling extracted from reference/amna.docx (see DESIGN_SPEC.md) ----
  const HEAD_FONT = "Tahoma";          // letterhead + signatory
  const ITEM_SHADE = "F8F9FA";         // faint tint on the description cell
  const line = { style: BorderStyle.SINGLE, size: 4, color: "000000" }; // 0.5pt black
  const noLine = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
  const fullB = { top: line, bottom: line, left: line, right: line };
  const totalLabelB = { top: line, bottom: noLine, left: noLine, right: line }; // open TOTAL label
  const noB = { top: noLine, bottom: noLine, left: noLine, right: noLine };
  const cellMargins = { top: 80, bottom: 80, left: 100, right: 100 };

  // ---- letterhead (Tahoma 11pt, no logo, right column positioned via borderless table) ----
  const lh = (text) => new Paragraph({ children: [new TextRun({ text, font: HEAD_FONT, size: 22 })] });
  const rightChildren = [
    new Paragraph({ children: [new TextRun({ text: "PROFORMA INVOICE", font: HEAD_FONT, size: 28, bold: true })] }),
    lh(`DATE: ${dateStr}`),
  ];
  if (invNo) rightChildren.push(lh(`NO: ${invNo}`));

  const headerTable = new Table({
    width: { size: 9225, type: WidthType.DXA },
    columnWidths: [6160, 3065],
    borders: { top: noLine, bottom: noLine, left: noLine, right: noLine, insideH: noLine, insideV: noLine },
    rows: [new TableRow({ children: [
      new TableCell({
        borders: noB, width: { size: 6160, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 120 },
        verticalAlign: VerticalAlign.TOP,
        children: [
          lh("İstanbul Okan Üniversitesi Tuzla Kampüsü"),
          lh("Tuzla Kampusu, Akfırat- Tuzla"),
          lh("34959 İstanbul/Turkey"),
          lh("Tel: +90 216 677 1630"),
          lh(""),
          lh("www.okan.edu.tr"),
        ],
      }),
      new TableCell({
        borders: noB, width: { size: 3065, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 120, right: 0 },
        verticalAlign: VerticalAlign.TOP,
        children: rightChildren,
      }),
    ]})],
  });

  // ---- line-item table (12pt, plain white header, black borders, true right-aligned amounts) ----
  const tblTxt = (text) => new TextRun({ text, size: 24 });

  const invoiceTable = new Table({
    width: { size: 9225, type: WidthType.DXA },
    columnWidths: [6285, 2940],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: fullB, width: { size: 6285, type: WidthType.DXA }, margins: cellMargins, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [tblTxt("DESCRIPTION")] })] }),
        new TableCell({ borders: fullB, width: { size: 2940, type: WidthType.DXA }, margins: cellMargins, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [tblTxt("AMOUNT")] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: fullB, width: { size: 6285, type: WidthType.DXA }, shading: { fill: ITEM_SHADE, type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [tblTxt(descText)] })] }),
        new TableCell({ borders: fullB, width: { size: 2940, type: WidthType.DXA }, margins: cellMargins, verticalAlign: VerticalAlign.TOP, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [tblTxt(`${invAmt} USD`)] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: totalLabelB, width: { size: 6285, type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [tblTxt("TOTAL")] })] }),
        new TableCell({ borders: fullB, width: { size: 2940, type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [tblTxt(`${totAmt || invAmt} USD`)] })] }),
      ]}),
    ],
  });

  // ---- payment terms (label 12pt; sentences 11pt, student name bold) ----
  const paymentLines = [
    new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "PAYMENT TERMS:", size: 24, bold: true })] }),
  ];
  if (totAmt) {
    paymentLines.push(new Paragraph({ spacing: { after: 80 }, children: [
      new TextRun({ text: fullName, bold: true, size: 22 }),
      new TextRun({ text: `'s ${feeText} is `, size: 22 }),
      new TextRun({ text: `${totAmt} USD`, size: 22, bold: true }),
      new TextRun({ text: ` for ${acYearStr}.`, size: 22 }),
    ]}));
  }
  paymentLines.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "This amount should be paid at once to the University's Bank Account information below.", size: 22 })] }));

  // ---- bank information (plain Calibri 11pt paragraphs, not a table) ----
  const bankRows = [
    ["BENEFICIARY NAME", "İSTANBUL OKAN UNİVERSİTESİ"],
    ["NAME OF BANK", "FIBA BANKA"],
    ["BRANCH", "MERKEZ ISTANBUL BRANCH"],
    null,
    ["SWIFT", "FBHLTRISXXX"],
    ["BANK NO", "103"],
    ["IBAN", "TR73 0010 3000 0000 0029 9568 81"],
  ];
  const bankLines = bankRows.map((r) =>
    new Paragraph({ spacing: { after: 40 }, children: r
      ? [new TextRun({ text: `${r[0]}: `, size: 22, bold: true }), new TextRun({ text: r[1], size: 22 })]
      : [new TextRun({ text: "", size: 22 })] })
  );

  return new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 2880, right: 1417, bottom: 2400, left: 993 } }
      },
      children: [
        headerTable, sp(240), invoiceTable, sp(240),
        ...paymentLines, sp(160),
        new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "BANK INFORMATION:", bold: true, size: 24 })] }),
        ...bankLines, sp(800),
        new Paragraph({ children: [new TextRun({ text: "Elif Tuğçe Dağ", font: HEAD_FONT, size: 22 })] }),
        new Paragraph({ children: [new TextRun({ text: "Student Operations Specialist", font: HEAD_FONT, size: 22 })] }),
      ]
    }]
  });
}

app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;
    data.descType = data.descType || "registration";
    data.feeType = data.feeType || "tuition";
    if (!data.firstName || !data.lastName || !data.invAmount || !data.program) {
      return res.status(400).json({ error: "Missing required fields" });
    }
    const doc = buildDoc(data);
    const buffer = await Packer.toBuffer(doc);
    const filename = `invoice_${data.firstName}_${data.lastName}_${(data.acYear || "2024-2025").replace("-", "_")}${data.invNo ? "_" + data.invNo : ""}.docx`;
    const asciiName = filename.normalize("NFKD").replace(/[^\x20-\x7E]/g, "");
    res.setHeader("Content-Disposition", `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = 3001;
app.listen(PORT, () => console.log(`Local invoice server running on http://localhost:${PORT}`));
