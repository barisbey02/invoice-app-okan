import { useState } from "react";

const API_URL = process.env.REACT_APP_API_URL || "http://localhost:3001";

const DEPTS = [
  { name: "Medicine", typical: 19475 },
  { name: "Dentistry", typical: 18000 },
  { name: "Pharmacy", typical: 14000 },
  { name: "Computer Engineering", typical: 9000 },
  { name: "Software Engineering", typical: 8500 },
  { name: "Electrical Engineering", typical: 8000 },
  { name: "Business Administration", typical: 7500 },
  { name: "Architecture", typical: 9500 },
];

function getSuggestion(amt) {
  if (!amt) return null;
  const n = parseFloat(amt);
  if (n >= 6000) return "Medicine";
  if (n >= 5500) return "Dentistry";
  if (n >= 4500) return "Pharmacy";
  if (n >= 3000) return "Computer Engineering";
  if (n >= 2000) return "Software Engineering";
  if (n > 0) return "Business Administration";
  return null;
}

function fmt(n) {
  return parseFloat(n).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default function App() {
  const [form, setForm] = useState({
    firstName: "", lastName: "",
    date: new Date().toISOString().split("T")[0],
    acYear: "2025-2026",
    invAmount: "", totalAmount: "",
    invNo: "", customDept: "", customDesc: "",
  });
  const [selectedDept, setSelectedDept] = useState("");
  const [descType, setDescType] = useState("registration");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);

  const set = (k) => (e) => {
    setForm((f) => ({ ...f, [k]: e.target.value }));
    setSuccess(false);
    setError("");
  };

  const suggestion = getSuggestion(form.invAmount);
  const program = form.customDept.trim() || selectedDept;

  const handleDeptClick = (name) => {
    setSelectedDept(name);
    setForm((f) => ({ ...f, customDept: "" }));
  };

  const handleGenerate = async () => {
    if (!form.firstName || !form.lastName) return setError("Enter first and last name.");
    if (!form.invAmount) return setError("Enter invoice amount.");
    if (!program) return setError("Select or type a program.");
    if (!form.date) return setError("Select a date.");

    setLoading(true);
    setError("");
    setSuccess(false);

    try {
      const res = await fetch(`${API_URL}/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          firstName: form.firstName,
          lastName: form.lastName,
          date: form.date,
          acYear: form.acYear || "2025-2026",
          invAmount: form.invAmount,
          totalAmount: form.totalAmount || null,
          program,
          invNo: form.invNo || null,
          descType,
          customDesc: form.customDesc || null,
        }),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Server error");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const fn = `invoice_${form.firstName}_${form.lastName}_${(form.acYear || "2025-2026").replace("-", "_")}${form.invNo ? "_" + form.invNo : ""}.docx`;
      a.download = fn;
      a.click();
      URL.revokeObjectURL(url);
      setSuccess(true);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={s.page}>
      <div style={s.container}>

        {/* Header */}
        <div style={s.header}>
          <div style={s.headerBadge}>Istanbul Okan University</div>
          <h1 style={s.title}>Invoice Generator</h1>
          <p style={s.subtitle}>Fill in the student details and download a ready-to-send proforma invoice.</p>
        </div>

        {/* Student Info */}
        <section style={s.card}>
          <p style={s.cardLabel}>Student info</p>
          <div style={s.row2}>
            <Field label="First name" value={form.firstName} onChange={set("firstName")} placeholder="Amna" />
            <Field label="Last name" value={form.lastName} onChange={set("lastName")} placeholder="Atiq" />
          </div>
          <div style={s.row3}>
            <Field label="Invoice date" type="date" value={form.date} onChange={set("date")} />
            <Field label="Academic year" value={form.acYear} onChange={set("acYear")} placeholder="2025-2026" />
            <Field label={<>Invoice no <span style={s.opt}>optional</span></>} value={form.invNo} onChange={set("invNo")} placeholder="INV-001" />
          </div>
        </section>

        {/* Amount */}
        <section style={s.card}>
          <p style={s.cardLabel}>Invoice amount (USD)</p>
          <div style={s.row2}>
            <Field label="Amount on this invoice" type="number" value={form.invAmount} onChange={set("invAmount")} placeholder="7,000" />
            <Field label={<>Total annual tuition <span style={s.opt}>optional</span></>} type="number" value={form.totalAmount} onChange={set("totalAmount")} placeholder="19,475" />
          </div>
          {form.invAmount && fmt(form.invAmount) && (
            <div style={s.amountPreview}>
              <span style={s.amountValue}>{fmt(form.invAmount)} USD</span>
              {form.totalAmount && <span style={s.amountTotal}> · Total: {fmt(form.totalAmount)} USD</span>}
            </div>
          )}
        </section>

        {/* Description type */}
        <section style={s.card}>
          <p style={s.cardLabel}>Invoice description</p>
          <div style={s.deptGrid}>
            {[
              { key: "registration", label: "Registration fee" },
              { key: "debt", label: "Debt payment" },
              { key: "dorm", label: "Dormitory fee" },
            ].map((d) => (
              <button
                key={d.key}
                style={{ ...s.deptBtn, ...(descType === d.key && !form.customDesc ? s.deptBtnActive : {}) }}
                onClick={() => { setDescType(d.key); setForm(f => ({ ...f, customDesc: "" })); }}
              >
                <div style={s.deptName}>{d.label}</div>
              </button>
            ))}
          </div>
          <Field
            label={<>Or type custom description <span style={s.opt}>overrides selection</span></>}
            value={form.customDesc}
            onChange={(e) => { set("customDesc")(e); }}
            placeholder="e.g. 2025-2026 exchange program fee..."
          />
        </section>

        {/* Program */}
        <section style={s.card}>
          <p style={s.cardLabel}>Department / program</p>

          {suggestion && !selectedDept && !form.customDept && (
            <div style={s.suggestion}>
              Suggested based on amount: <strong>{suggestion}</strong>
              <button style={s.applyBtn} onClick={() => { setSelectedDept(suggestion); setForm(f => ({ ...f, customDept: "" })); }}>
                Apply
              </button>
            </div>
          )}

          <div style={s.deptGrid}>
            {DEPTS.map((d) => (
              <button
                key={d.name}
                style={{ ...s.deptBtn, ...(selectedDept === d.name && !form.customDept ? s.deptBtnActive : {}) }}
                onClick={() => handleDeptClick(d.name)}
              >
                <div style={s.deptName}>{d.name}</div>
                <div style={s.deptFee}>~${d.typical.toLocaleString()} / yr</div>
              </button>
            ))}
          </div>

          <Field
            label={<>Or type custom program <span style={s.opt}>overrides selection</span></>}
            value={form.customDept}
            onChange={(e) => { set("customDept")(e); setSelectedDept(""); }}
            placeholder="e.g. Law, Nursing..."
          />
        </section>

        {/* Preview strip */}
        {(form.firstName || form.lastName || program) && (
          <div style={s.preview}>
            <span style={s.previewItem}><span style={s.previewLabel}>Student</span> {[form.firstName, form.lastName].filter(Boolean).join(" ") || "—"}</span>
            <span style={s.previewDot}>·</span>
            <span style={s.previewItem}><span style={s.previewLabel}>Program</span> {program || "—"}</span>
            <span style={s.previewDot}>·</span>
            <span style={s.previewItem}><span style={s.previewLabel}>Year</span> {form.acYear || "—"}</span>
            {form.invNo && <><span style={s.previewDot}>·</span><span style={s.previewItem}><span style={s.previewLabel}>No.</span> {form.invNo}</span></>}
          </div>
        )}

        {/* Generate */}
        {error && <div style={s.errorBox}>{error}</div>}
        {success && <div style={s.successBox}>Invoice downloaded successfully.</div>}

        <button style={{ ...s.genBtn, ...(loading ? s.genBtnLoading : {}) }} onClick={handleGenerate} disabled={loading}>
          {loading ? "Generating..." : "Download invoice .docx"}
        </button>

      </div>
    </div>
  );
}

function Field({ label, value, onChange, placeholder, type = "text" }) {
  return (
    <div style={s.field}>
      <label style={s.label}>{label}</label>
      <input
        type={type}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        style={s.input}
        onFocus={e => e.target.style.borderColor = "#2563eb"}
        onBlur={e => e.target.style.borderColor = "#e2e0d8"}
      />
    </div>
  );
}

const s = {
  page: {
    minHeight: "100vh",
    background: "#f5f3ee",
    fontFamily: "'DM Sans', sans-serif",
    padding: "2rem 1rem 4rem",
  },
  container: {
    maxWidth: 640,
    margin: "0 auto",
  },
  header: {
    marginBottom: "2rem",
    paddingBottom: "1.5rem",
    borderBottom: "1px solid #ddd9ce",
  },
  headerBadge: {
    display: "inline-block",
    fontSize: 11,
    fontFamily: "'DM Mono', monospace",
    letterSpacing: "0.08em",
    color: "#888",
    textTransform: "uppercase",
    marginBottom: 10,
  },
  title: {
    fontSize: 28,
    fontWeight: 300,
    color: "#1a1814",
    margin: "0 0 6px",
    letterSpacing: "-0.02em",
  },
  subtitle: {
    fontSize: 14,
    color: "#888",
    margin: 0,
    fontWeight: 300,
  },
  card: {
    background: "#fff",
    border: "1px solid #e8e4db",
    borderRadius: 10,
    padding: "1.25rem 1.5rem",
    marginBottom: "1rem",
  },
  cardLabel: {
    fontSize: 11,
    fontFamily: "'DM Mono', monospace",
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    color: "#aaa",
    margin: "0 0 1rem",
  },
  row2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  row3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 },
  field: { marginBottom: 12 },
  label: { display: "block", fontSize: 12, color: "#888", marginBottom: 4, fontWeight: 400 },
  input: {
    width: "100%",
    padding: "8px 10px",
    fontSize: 14,
    border: "1px solid #e2e0d8",
    borderRadius: 6,
    background: "#faf9f7",
    color: "#1a1814",
    outline: "none",
    fontFamily: "'DM Sans', sans-serif",
    boxSizing: "border-box",
    transition: "border-color 0.15s",
  },
  opt: { fontSize: 11, color: "#bbb", fontWeight: 300 },
  amountPreview: {
    marginTop: 4,
    padding: "8px 12px",
    background: "#f5f3ee",
    borderRadius: 6,
    fontSize: 13,
  },
  amountValue: { fontFamily: "'DM Mono', monospace", fontWeight: 500, color: "#1a1814" },
  amountTotal: { color: "#999", marginLeft: 4 },
  suggestion: {
    background: "#fffbeb",
    border: "1px solid #fde68a",
    borderRadius: 6,
    padding: "8px 12px",
    fontSize: 13,
    color: "#92400e",
    marginBottom: 12,
    display: "flex",
    alignItems: "center",
    gap: 8,
  },
  applyBtn: {
    marginLeft: "auto",
    padding: "3px 10px",
    fontSize: 12,
    background: "#1a1814",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    cursor: "pointer",
    fontFamily: "'DM Sans', sans-serif",
  },
  deptGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 8,
    marginBottom: 12,
  },
  deptBtn: {
    border: "1px solid #e8e4db",
    borderRadius: 7,
    padding: "9px 12px",
    textAlign: "left",
    cursor: "pointer",
    background: "#faf9f7",
    transition: "all 0.12s",
    fontFamily: "'DM Sans', sans-serif",
  },
  deptBtnActive: {
    background: "#1a1814",
    borderColor: "#1a1814",
    color: "#fff",
  },
  deptName: { fontSize: 13, fontWeight: 500, color: "inherit" },
  deptFee: { fontSize: 11, color: "#aaa", marginTop: 2 },
  preview: {
    display: "flex",
    flexWrap: "wrap",
    gap: 6,
    alignItems: "center",
    padding: "10px 14px",
    background: "#fff",
    border: "1px solid #e8e4db",
    borderRadius: 8,
    marginBottom: 12,
    fontSize: 13,
    color: "#1a1814",
  },
  previewItem: { display: "flex", gap: 5, alignItems: "center" },
  previewLabel: { fontSize: 11, color: "#aaa", fontFamily: "'DM Mono', monospace" },
  previewDot: { color: "#ddd" },
  errorBox: {
    background: "#fff1f0",
    border: "1px solid #fca5a5",
    borderRadius: 6,
    padding: "8px 12px",
    fontSize: 13,
    color: "#b91c1c",
    marginBottom: 10,
  },
  successBox: {
    background: "#f0fdf4",
    border: "1px solid #86efac",
    borderRadius: 6,
    padding: "8px 12px",
    fontSize: 13,
    color: "#15803d",
    marginBottom: 10,
  },
  genBtn: {
    width: "100%",
    padding: "12px",
    fontSize: 14,
    fontWeight: 500,
    cursor: "pointer",
    background: "#1a1814",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontFamily: "'DM Sans', sans-serif",
    letterSpacing: "0.01em",
    transition: "opacity 0.15s",
  },
  genBtnLoading: { opacity: 0.5 },
};
