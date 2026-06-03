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

function fmt(n) {
  return parseFloat(n).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default function App() {
  const [form, setForm] = useState({
    firstName: "", lastName: "",
    date: (() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; })(),
    acYear: "2025-2026",
    invAmount: "", totalAmount: "",
    invNo: "", customDept: "",
  });
  const [selectedDept, setSelectedDept] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const [showContact, setShowContact] = useState(false);

  const set = (k) => (e) => {
    setForm((f) => ({ ...f, [k]: e.target.value }));
    setSuccess(false);
    setError("");
  };

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
      a.download = `invoice_${form.firstName}_${form.lastName}_${(form.acYear || "2025-2026").replace("-", "_")}${form.invNo ? "_" + form.invNo : ""}.docx`;
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
      {/* Background blobs */}
      <div style={s.blob1} />
      <div style={s.blob2} />

      <div style={s.container}>

        {/* Header */}
        <div style={s.header}>
          <div style={s.logoRow}>
            <div style={s.logoBadge}>
              <span style={s.logoIcon}>◈</span>
              <span style={s.logoText}>Okan University</span>
            </div>
            <a href="https://baikgroup.com" target="_blank" rel="noreferrer" style={s.madeBy}>
              Made by baikgroup.com
            </a>
          </div>
          <h1 style={s.title}>Invoice Generator</h1>
          <p style={s.subtitle}>Generate proforma invoices for student tuition fees</p>
        </div>

        {/* Student Info */}
        <div style={s.card}>
          <div style={s.cardHeader}>
            <span style={s.cardIcon}>👤</span>
            <span style={s.cardTitle}>Student Information</span>
          </div>
          <div style={s.row2}>
            <Field label="First Name" value={form.firstName} onChange={set("firstName")} placeholder="Amna" />
            <Field label="Last Name" value={form.lastName} onChange={set("lastName")} placeholder="Atiq" />
          </div>
          <div style={s.row3}>
            <Field label="Invoice Date" type="date" value={form.date} onChange={set("date")} />
            <Field label="Academic Year" value={form.acYear} onChange={set("acYear")} placeholder="2025-2026" />
            <Field label="Invoice No" value={form.invNo} onChange={set("invNo")} placeholder="INV-001" optional />
          </div>
        </div>

        {/* Amount */}
        <div style={s.card}>
          <div style={s.cardHeader}>
            <span style={s.cardIcon}>💵</span>
            <span style={s.cardTitle}>Invoice Amount (USD)</span>
          </div>
          <div style={s.row2}>
            <Field label="Amount on This Invoice" type="number" value={form.invAmount} onChange={set("invAmount")} placeholder="7,000" />
            <Field label="Total Annual Tuition" type="number" value={form.totalAmount} onChange={set("totalAmount")} placeholder="19,475" optional />
          </div>
          {form.invAmount && (
            <div style={s.amountPreview}>
              <div style={s.amountLeft}>
                <div style={s.amountLabel}>Invoice Amount</div>
                <div style={s.amountValue}>{fmt(form.invAmount)} <span style={s.amountCurrency}>USD</span></div>
              </div>
              {form.totalAmount && (
                <div style={s.amountRight}>
                  <div style={s.amountLabel}>Annual Total</div>
                  <div style={s.amountValue}>{fmt(form.totalAmount)} <span style={s.amountCurrency}>USD</span></div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Program */}
        <div style={s.card}>
          <div style={s.cardHeader}>
            <span style={s.cardIcon}>🎓</span>
            <span style={s.cardTitle}>Department / Program</span>
          </div>
          <div style={s.deptGrid}>
            {DEPTS.map((d) => (
              <button
                key={d.name}
                style={{ ...s.deptBtn, ...(selectedDept === d.name && !form.customDept ? s.deptBtnActive : {}) }}
                onClick={() => handleDeptClick(d.name)}
              >
                <div style={{ ...s.deptName, ...(selectedDept === d.name && !form.customDept ? { color: "#fff" } : {}) }}>{d.name}</div>
                <div style={{ ...s.deptFee, ...(selectedDept === d.name && !form.customDept ? { color: "rgba(255,255,255,0.6)" } : {}) }}>${d.typical.toLocaleString()} / yr</div>
              </button>
            ))}
          </div>
          <Field
            label="Custom Program"
            value={form.customDept}
            onChange={(e) => { set("customDept")(e); setSelectedDept(""); }}
            placeholder="e.g. Law, Nursing, Physiotherapy..."
            optional
          />
        </div>

        {/* Preview */}
        {(form.firstName || form.lastName || program) && (
          <div style={s.preview}>
            <div style={s.previewTitle}>Invoice Preview</div>
            <div style={s.previewRow}>
              {form.firstName || form.lastName ? (
                <div style={s.previewChip}>
                  <span style={s.previewChipLabel}>Student</span>
                  <span style={s.previewChipValue}>{[form.firstName, form.lastName].filter(Boolean).join(" ")}</span>
                </div>
              ) : null}
              {program ? (
                <div style={s.previewChip}>
                  <span style={s.previewChipLabel}>Program</span>
                  <span style={s.previewChipValue}>{program}</span>
                </div>
              ) : null}
              {form.acYear ? (
                <div style={s.previewChip}>
                  <span style={s.previewChipLabel}>Year</span>
                  <span style={s.previewChipValue}>{form.acYear}</span>
                </div>
              ) : null}
              {form.invNo ? (
                <div style={s.previewChip}>
                  <span style={s.previewChipLabel}>No.</span>
                  <span style={s.previewChipValue}>{form.invNo}</span>
                </div>
              ) : null}
            </div>
          </div>
        )}

        {/* Error / Success */}
        {error && (
          <div style={s.errorBox}>
            <span style={s.alertIcon}>⚠</span> {error}
          </div>
        )}
        {success && (
          <div style={s.successBox}>
            <span style={s.alertIcon}>✓</span> Invoice downloaded successfully.
          </div>
        )}

        {/* Generate Button */}
        <button style={{ ...s.genBtn, ...(loading ? s.genBtnLoading : {}) }} onClick={handleGenerate} disabled={loading}>
          {loading ? (
            <span style={s.genBtnInner}><span style={s.spinner} /> Generating...</span>
          ) : (
            <span style={s.genBtnInner}>⬇ &nbsp;Download Invoice (.docx)</span>
          )}
        </button>

      </div>

      {/* Contact Card */}
      {showContact && (
        <div style={s.contactCard}>
          <button style={s.contactClose} onClick={() => setShowContact(false)}>✕</button>
          <div style={s.contactName}>Salah Al Baik</div>
          <div style={s.contactCompany}>Baik Group</div>
          <div style={s.contactDivider} />
          <a href="https://baikgroup.com" target="_blank" rel="noreferrer" style={s.contactLink}>
            🌐 baikgroup.com
          </a>
          <a href="mailto:contact@baikgroup.com" style={s.contactLink}>
            ✉ contact@baikgroup.com
          </a>
        </div>
      )}

      <button style={s.contactBtn} onClick={() => setShowContact((v) => !v)} title="Contact">
        SB
      </button>
    </div>
  );
}

function Field({ label, value, onChange, placeholder, type = "text", optional }) {
  return (
    <div style={s.field}>
      <label style={s.label}>
        {label}
        {optional && <span style={s.opt}>optional</span>}
      </label>
      <input
        type={type}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        style={s.input}
        onFocus={e => { e.target.style.borderColor = "#6366f1"; e.target.style.boxShadow = "0 0 0 3px rgba(99,102,241,0.12)"; }}
        onBlur={e => { e.target.style.borderColor = "#e5e7eb"; e.target.style.boxShadow = "none"; }}
      />
    </div>
  );
}

const s = {
  page: {
    minHeight: "100vh",
    background: "#f8f9ff",
    fontFamily: "'Inter', 'Segoe UI', sans-serif",
    padding: "2.5rem 1rem 5rem",
    position: "relative",
    overflow: "hidden",
  },
  blob1: {
    position: "fixed", top: -180, right: -180,
    width: 500, height: 500, borderRadius: "50%",
    background: "radial-gradient(circle, rgba(99,102,241,0.12) 0%, transparent 70%)",
    pointerEvents: "none",
  },
  blob2: {
    position: "fixed", bottom: -200, left: -150,
    width: 450, height: 450, borderRadius: "50%",
    background: "radial-gradient(circle, rgba(139,92,246,0.10) 0%, transparent 70%)",
    pointerEvents: "none",
  },
  container: {
    maxWidth: 660,
    margin: "0 auto",
    position: "relative",
    zIndex: 1,
  },
  header: {
    marginBottom: "2rem",
  },
  logoRow: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    marginBottom: "1.25rem",
  },
  logoBadge: {
    display: "flex", alignItems: "center", gap: 8,
    background: "linear-gradient(135deg, #6366f1, #8b5cf6)",
    color: "#fff",
    padding: "6px 14px",
    borderRadius: 100,
    fontSize: 13,
    fontWeight: 500,
  },
  logoIcon: { fontSize: 14 },
  logoText: { letterSpacing: "0.01em" },
  madeBy: {
    fontSize: 12,
    color: "#9ca3af",
    background: "#fff",
    border: "1px solid #e5e7eb",
    borderRadius: 100,
    padding: "4px 12px",
    textDecoration: "none",
    transition: "color 0.15s",
  },
  title: {
    fontSize: 34,
    fontWeight: 700,
    color: "#0f172a",
    margin: "0 0 8px",
    letterSpacing: "-0.03em",
  },
  subtitle: {
    fontSize: 15,
    color: "#6b7280",
    margin: 0,
    fontWeight: 400,
  },
  card: {
    background: "#fff",
    border: "1px solid #e5e7eb",
    borderRadius: 16,
    padding: "1.5rem",
    marginBottom: "1rem",
    boxShadow: "0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03)",
  },
  cardHeader: {
    display: "flex", alignItems: "center", gap: 8,
    marginBottom: "1.25rem",
  },
  cardIcon: { fontSize: 16 },
  cardTitle: {
    fontSize: 14,
    fontWeight: 600,
    color: "#374151",
    letterSpacing: "-0.01em",
  },
  row2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  row3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 },
  field: { marginBottom: 12 },
  label: {
    display: "flex", alignItems: "center", gap: 6,
    fontSize: 12, fontWeight: 500, color: "#6b7280",
    marginBottom: 5,
  },
  input: {
    width: "100%",
    padding: "9px 12px",
    fontSize: 14,
    border: "1px solid #e5e7eb",
    borderRadius: 9,
    background: "#fafafa",
    color: "#111827",
    outline: "none",
    fontFamily: "inherit",
    boxSizing: "border-box",
    transition: "border-color 0.15s, box-shadow 0.15s",
  },
  opt: {
    fontSize: 10, color: "#d1d5db",
    background: "#f3f4f6", borderRadius: 4,
    padding: "1px 5px", fontWeight: 400,
  },
  amountPreview: {
    display: "flex", gap: 12,
    marginTop: 4,
    padding: "14px 16px",
    background: "linear-gradient(135deg, #f5f3ff, #ede9fe)",
    borderRadius: 10,
    border: "1px solid #ddd6fe",
  },
  amountLeft: { flex: 1 },
  amountRight: { flex: 1 },
  amountLabel: { fontSize: 11, color: "#8b5cf6", fontWeight: 500, marginBottom: 2, textTransform: "uppercase", letterSpacing: "0.05em" },
  amountValue: { fontSize: 20, fontWeight: 700, color: "#1e1b4b", fontVariantNumeric: "tabular-nums" },
  amountCurrency: { fontSize: 13, fontWeight: 400, color: "#7c3aed" },
  deptGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 8,
    marginBottom: 14,
  },
  deptBtn: {
    border: "1.5px solid #e5e7eb",
    borderRadius: 10,
    padding: "10px 13px",
    textAlign: "left",
    cursor: "pointer",
    background: "#fafafa",
    transition: "all 0.15s",
    fontFamily: "inherit",
  },
  deptBtnActive: {
    background: "linear-gradient(135deg, #6366f1, #8b5cf6)",
    borderColor: "transparent",
    boxShadow: "0 4px 12px rgba(99,102,241,0.3)",
  },
  deptName: { fontSize: 13, fontWeight: 500, color: "#1f2937" },
  deptFee: { fontSize: 11, color: "#9ca3af", marginTop: 2 },
  preview: {
    background: "#fff",
    border: "1px solid #e5e7eb",
    borderRadius: 12,
    padding: "14px 16px",
    marginBottom: 12,
    boxShadow: "0 1px 3px rgba(0,0,0,0.04)",
  },
  previewTitle: {
    fontSize: 11, fontWeight: 600, color: "#9ca3af",
    textTransform: "uppercase", letterSpacing: "0.07em",
    marginBottom: 10,
  },
  previewRow: { display: "flex", flexWrap: "wrap", gap: 8 },
  previewChip: {
    display: "flex", alignItems: "center", gap: 5,
    background: "#f3f4f6", borderRadius: 100,
    padding: "4px 10px",
  },
  previewChipLabel: { fontSize: 10, color: "#9ca3af", fontWeight: 500, textTransform: "uppercase" },
  previewChipValue: { fontSize: 13, color: "#1f2937", fontWeight: 500 },
  errorBox: {
    background: "#fef2f2",
    border: "1px solid #fecaca",
    borderRadius: 10,
    padding: "10px 14px",
    fontSize: 13, color: "#dc2626",
    marginBottom: 12,
    display: "flex", alignItems: "center", gap: 8,
  },
  successBox: {
    background: "#f0fdf4",
    border: "1px solid #bbf7d0",
    borderRadius: 10,
    padding: "10px 14px",
    fontSize: 13, color: "#16a34a",
    marginBottom: 12,
    display: "flex", alignItems: "center", gap: 8,
  },
  alertIcon: { fontSize: 15 },
  genBtn: {
    width: "100%",
    padding: "14px",
    fontSize: 15,
    fontWeight: 600,
    cursor: "pointer",
    background: "linear-gradient(135deg, #6366f1, #8b5cf6)",
    color: "#fff",
    border: "none",
    borderRadius: 12,
    fontFamily: "inherit",
    letterSpacing: "-0.01em",
    transition: "opacity 0.15s, transform 0.1s",
    boxShadow: "0 4px 14px rgba(99,102,241,0.4)",
  },
  genBtnLoading: { opacity: 0.6 },
  genBtnInner: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8 },
  spinner: {
    display: "inline-block",
    width: 14, height: 14,
    border: "2px solid rgba(255,255,255,0.4)",
    borderTopColor: "#fff",
    borderRadius: "50%",
    animation: "spin 0.7s linear infinite",
  },

  // Contact button
  contactBtn: {
    position: "fixed",
    bottom: 24, right: 24,
    width: 44, height: 44,
    borderRadius: "50%",
    background: "linear-gradient(135deg, #0f172a, #1e293b)",
    color: "#fff",
    border: "none",
    fontSize: 13,
    fontWeight: 700,
    cursor: "pointer",
    boxShadow: "0 4px 14px rgba(0,0,0,0.25)",
    zIndex: 100,
    fontFamily: "inherit",
    letterSpacing: "0.03em",
    transition: "transform 0.15s, box-shadow 0.15s",
  },
  contactCard: {
    position: "fixed",
    bottom: 78, right: 24,
    width: 220,
    background: "#fff",
    borderRadius: 14,
    boxShadow: "0 8px 32px rgba(0,0,0,0.14), 0 1px 3px rgba(0,0,0,0.08)",
    padding: "18px 16px 14px",
    zIndex: 99,
    border: "1px solid #e5e7eb",
  },
  contactClose: {
    position: "absolute",
    top: 10, right: 12,
    background: "none", border: "none",
    fontSize: 12, color: "#9ca3af",
    cursor: "pointer", padding: 2,
  },
  contactName: {
    fontSize: 15, fontWeight: 700, color: "#0f172a",
    letterSpacing: "-0.02em",
  },
  contactCompany: {
    fontSize: 12, color: "#6366f1", fontWeight: 500,
    marginTop: 2, marginBottom: 10,
  },
  contactDivider: {
    height: 1, background: "#f3f4f6", marginBottom: 10,
  },
  contactLink: {
    display: "block",
    fontSize: 13, color: "#374151",
    textDecoration: "none",
    padding: "5px 0",
    transition: "color 0.1s",
  },
};
