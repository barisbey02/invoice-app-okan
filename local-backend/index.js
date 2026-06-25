require("dotenv").config({ path: require("path").join(__dirname, ".env") });

const express = require("express");
const cors = require("cors");
const sql = require("mssql");
// Shared invoice generator — single source of truth, also used by the Render
// backend. docx is injected so this module needs no deps of its own.
const { generateInvoiceBuffer } = require("../frontend/backend/invoice-doc")(require("docx"));

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
const ADMIN_PASS = (process.env.ADMIN_PASS || "").trim();

// ── Admin: lookup student (returns full raw row) ───────────────────────────
app.post("/admin-lookup", requireApiKey, async (req, res) => {
  const { password, studentNo } = req.body;
  if (!password || password.trim() !== ADMIN_PASS) {
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

  if (!password || password.trim() !== ADMIN_PASS) {
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

// ── Invoice generation — buildDoc/fmtAmt/sp live in the shared module ──────
// (../frontend/backend/invoice-doc.js, imported at the top of this file).

// ── Env diagnostic — no auth required, safe (values never exposed) ──────────
app.get("/debug-env", (req, res) => {
  const vars = ["DB_SERVER", "DB_USER", "DB_PASSWORD", "DB_NAME", "ADMIN_PASS", "API_KEY"];
  res.json({
    __dirname,
    cwd: process.cwd(),
    env: Object.fromEntries(
      vars.map(k => [k, process.env[k] ? `set (${process.env[k].trim().length} chars)` : "MISSING"])
    ),
  });
});

app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;
    data.descType = data.descType || "registration";
    data.feeType = data.feeType || "tuition";
    if (!data.firstName || !data.lastName || !data.invAmount || !data.program) {
      return res.status(400).json({ error: "Missing required fields" });
    }
    const { buffer, contentDisposition } = await generateInvoiceBuffer(data);
    res.setHeader("Content-Disposition", contentDisposition);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = 3001;
app.listen(PORT, () => console.log(`Local invoice server running on http://localhost:${PORT}`));
