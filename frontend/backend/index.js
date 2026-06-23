const express = require("express");
const cors = require("cors");
const { generateInvoiceBuffer } = require("./invoice-doc")(require("docx"));

const app = express();
app.use(cors());
app.use(express.json());

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

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Invoice server running on port ${PORT}`));
