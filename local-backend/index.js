const express = require("express");
const cors = require("cors");
const sql = require("mssql");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
} = require("docx");

const app = express();
app.use(cors());
app.use(express.json());

// ── DB config ──────────────────────────────────────────────────────────────
const dbConfig = {
  server: "192.168.5.9",
  user: "SA",
  password: "Wsad6aeqKu23s",
  database: "ForNetOkan",
  options: { trustServerCertificate: true, encrypt: false },
};

let pool;
async function getPool() {
  if (!pool) pool = await sql.connect(dbConfig);
  return pool;
}

// ── Student lookup ─────────────────────────────────────────────────────────
app.get("/student/:id", async (req, res) => {
  try {
    const db = await getPool();

    // Find the current user's KisiID to determine which TEMP table to query
    const userResult = await db.request()
      .input("login", sql.VarChar, "maliogrencilocal1")
      .query(`SELECT KisiID FROM FORNET_Kullanici WHERE LoginName = @login`);

    if (userResult.recordset.length === 0) {
      return res.status(500).json({ error: "Could not find user in FORNET_Kullanici" });
    }

    const kisiID = userResult.recordset[0].KisiID;
    const tableName = `TEMP_OgrenciKayitListesi_${kisiID}`;

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

// ── Schema discovery (use once to find correct table/column names) ──────────
app.get("/schema", async (req, res) => {
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

app.get("/schema/:table", async (req, res) => {
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

function buildDoc({ firstName, lastName, date, acYear, invAmount, totalAmount, program, invNo, descType, customDesc }) {
  const fullName = `${firstName} ${lastName}`;
  const acYearStr = acYear || "2024-2025";
  const d = new Date(date);
  const dateStr = `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
  const invAmt = fmtAmt(invAmount);
  const totAmt = totalAmount ? fmtAmt(totalAmount) : null;

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

  const thin = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
  const none = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
  const cb = { top: thin, bottom: thin, left: thin, right: thin };
  const nb = { top: none, bottom: none, left: none, right: none };

  const rightChildren = [
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PROFORMA INVOICE", bold: true, size: 28 })] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `DATE: ${dateStr}`, size: 18 })] }),
  ];
  if (invNo) rightChildren.push(
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `NO: ${invNo}`, size: 18 })] })
  );

  const headerTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [5400, 3960],
    borders: { top: none, bottom: none, left: none, right: none, insideH: none, insideV: none },
    rows: [new TableRow({ children: [
      new TableCell({
        borders: nb, width: { size: 5400, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: "İstanbul Okan Üniversitesi Tuzla Kampüsü", bold: true, size: 20 })] }),
          new Paragraph({ children: [new TextRun({ text: "Tuzla Kampusu, Akfırat- Tuzla", size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: "34959 İstanbul/Turkey", size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: "Tel: +90 216 677 1630", size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: "www.okan.edu.tr", size: 18, color: "1155CC" })] }),
        ],
      }),
      new TableCell({
        borders: nb, width: { size: 3960, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 120, right: 0 },
        verticalAlign: VerticalAlign.TOP,
        children: rightChildren,
      }),
    ]})]
  });

  const invoiceTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [7200, 2160],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: cb, width: { size: 7200, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "DESCRIPTION", bold: true, size: 20, color: "FFFFFF" })] })] }),
        new TableCell({ borders: cb, width: { size: 2160, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "AMOUNT", bold: true, size: 20, color: "FFFFFF" })] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: cb, width: { size: 7200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: descText, size: 20 })] })] }),
        new TableCell({ borders: cb, width: { size: 2160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `${invAmt} USD`, size: 20 })] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: cb, width: { size: 7200, type: WidthType.DXA }, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "TOTAL", bold: true, size: 20 })] })] }),
        new TableCell({ borders: cb, width: { size: 2160, type: WidthType.DXA }, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `${totAmt || invAmt} USD`, bold: true, size: 20 })] })] }),
      ]}),
    ]
  });

  const paymentLines = [
    new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "PAYMENT TERMS:", bold: true, size: 20 })] }),
  ];
  if (totAmt) {
    paymentLines.push(new Paragraph({ spacing: { after: 80 }, children: [
      new TextRun({ text: fullName, bold: true, size: 20 }),
      new TextRun({ text: "'s tuition fee is ", size: 20 }),
      new TextRun({ text: `${totAmt} USD`, bold: true, size: 20 }),
      new TextRun({ text: ` for ${acYearStr}.`, size: 20 }),
    ]}));
  }
  paymentLines.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "This amount should be paid at once to the University's Bank Account information below.", size: 20 })] }));

  const bankTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 6560],
    rows: [
      ["BENEFICIARY NAME", "İSTANBUL OKAN UNİVERSİTESİ"],
      ["NAME OF BANK", "FIBA BANKA"],
      ["BRANCH", "MERKEZ ISTANBUL BRANCH"],
      ["SWIFT", "FBHLTRISXXX"],
      ["BANK NO", "103"],
      ["IBAN", "TR73 0010 3000 0000 0029 9568 81"],
    ].map(([label, val]) => new TableRow({ children: [
      new TableCell({ borders: cb, width: { size: 2800, type: WidthType.DXA }, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18 })] })] }),
      new TableCell({ borders: cb, width: { size: 6560, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: val, size: 18 })] })] }),
    ]}))
  });

  return new Document({
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1800, right: 1080, bottom: 1080, left: 1080 } }
      },
      children: [
        headerTable, sp(240), invoiceTable, sp(240),
        ...paymentLines, sp(200),
        new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "BANK INFORMATION:", bold: true, size: 20 })] }),
        bankTable, sp(3200),
        new Paragraph({ children: [new TextRun({ text: "Elif Tuğçe Dağ", bold: true, size: 20 })] }),
        new Paragraph({ children: [new TextRun({ text: "Student Operations Specialist", size: 18 })] }),
      ]
    }]
  });
}

app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;
    data.descType = data.descType || "registration";
    if (!data.firstName || !data.lastName || !data.invAmount || !data.program) {
      return res.status(400).json({ error: "Missing required fields" });
    }
    const doc = buildDoc(data);
    const buffer = await Packer.toBuffer(doc);
    const filename = `invoice_${data.firstName}_${data.lastName}_${(data.acYear || "2024-2025").replace("-", "_")}${data.invNo ? "_" + data.invNo : ""}.docx`;
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = 3001;
app.listen(PORT, () => console.log(`Local invoice server running on http://localhost:${PORT}`));
