import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const DISC_LABELS = {
  D: "D – Dominant",
  I: "i – Influential",
  S: "S – Steady",
  C: "C – Conscientious",
};
const DISC_COLORS = {
  D: { bg: "#ffe8e8", color: "#c0392b", border: "#e05c5c" },
  I: { bg: "#fff3d4", color: "#b07a00", border: "#e0a020" },
  S: { bg: "#e8f8e8", color: "#2e7d32", border: "#5cb85c" },
  C: { bg: "#e8f0ff", color: "#1a56b0", border: "#4a80e0" },
};

const FIELDS = [
  {
    id: "strengths",
    label: "My top strengths",
    icon: "💪",
    placeholder: "What do you bring at your best?",
  },
  {
    id: "energises",
    label: "What energises me",
    icon: "⚡",
    placeholder: "What lights you up at work?",
  },
  {
    id: "drains",
    label: "What drains me",
    icon: "🪫",
    placeholder: "What saps your energy or focus?",
  },
  {
    id: "pressure",
    label: "My pressure signs",
    icon: "🌡️",
    placeholder: "How do others know when you're under stress?",
  },
  {
    id: "approach",
    label: "Best way to approach me",
    icon: "🤝",
    placeholder: "What helps you feel heard and respected?",
  },
  {
    id: "challenge",
    label: "Best way to challenge me",
    icon: "💬",
    placeholder: "How can others push your thinking productively?",
  },
  {
    id: "onemore",
    label: "One thing I want this team to know",
    icon: "✨",
    placeholder: "Anything else that helps people work with you well…",
    full: true,
  },
];

const EMPTY_FORM = {
  name: "",
  disc: "",
  strengths: "",
  energises: "",
  drains: "",
  pressure: "",
  approach: "",
  challenge: "",
  onemore: "",
};

// ── Styles ──
const s = {
  app: {
    minHeight: "100vh",
    background: "linear-gradient(135deg,#e8e4f8 0%,#f0eeff 50%,#e4eaf8 100%)",
    fontFamily: "'DM Sans', sans-serif",
    color: "#1a1a2e",
  },
  center: { maxWidth: 620, margin: "0 auto", padding: "36px 20px 60px" },
  header: { textAlign: "center", marginBottom: 32 },
  eyebrow: {
    fontSize: 11,
    fontWeight: 700,
    letterSpacing: "0.2em",
    textTransform: "uppercase",
    color: "#5c4fcf",
    marginBottom: 8,
  },
  h1: {
    fontFamily: "'DM Serif Display', serif",
    fontSize: "clamp(2rem,5vw,3rem)",
    lineHeight: 1.1,
    color: "#1a1a2e",
    margin: "0 0 8px",
  },
  h1em: { fontStyle: "italic", color: "#5c4fcf" },
  subtitle: {
    fontSize: 14,
    color: "#8b85b0",
    lineHeight: 1.6,
    maxWidth: 380,
    margin: "0 auto",
  },

  // tabs
  tabs: {
    display: "flex",
    gap: 6,
    marginBottom: 24,
    background: "#fff",
    borderRadius: 100,
    padding: 4,
    boxShadow: "0 2px 12px rgba(92,79,207,.1)",
  },
  tab: (active) => ({
    flex: 1,
    padding: "9px 0",
    borderRadius: 100,
    border: "none",
    fontFamily: "'DM Sans',sans-serif",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
    transition: "all .2s",
    background: active ? "#5c4fcf" : "transparent",
    color: active ? "#fff" : "#8b85b0",
  }),

  // card
  card: {
    background: "#fff",
    borderRadius: 20,
    overflow: "hidden",
    boxShadow: "0 4px 6px rgba(92,79,207,.06),0 20px 60px rgba(92,79,207,.12)",
  },
  stripe: {
    height: 6,
    background: "linear-gradient(90deg,#5c4fcf 0%,#9b8ff0 50%,#c4b8ff 100%)",
  },
  cardBody: { padding: "28px 28px 24px" },

  // form elements
  nameRow: {
    display: "flex",
    alignItems: "flex-end",
    gap: 14,
    marginBottom: 24,
    paddingBottom: 22,
    borderBottom: "1.5px solid #d4cff0",
  },
  avatar: {
    width: 50,
    height: 50,
    borderRadius: "50%",
    background: "#ede9ff",
    border: "2px dashed #5c4fcf",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 20,
    flexShrink: 0,
  },
  nameLabel: {
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.15em",
    textTransform: "uppercase",
    color: "#8b85b0",
    marginBottom: 4,
    display: "block",
  },
  nameInput: {
    width: "100%",
    border: "none",
    borderBottom: "2px solid #d4cff0",
    background: "transparent",
    fontFamily: "'DM Serif Display',serif",
    fontSize: 22,
    color: "#1a1a2e",
    padding: "4px 0",
    outline: "none",
  },
  fieldLabel: {
    display: "flex",
    alignItems: "center",
    gap: 7,
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.13em",
    textTransform: "uppercase",
    color: "#5c4fcf",
    marginBottom: 6,
  },
  textarea: {
    width: "100%",
    background: "#f7f6fb",
    border: "1.5px solid transparent",
    borderRadius: 10,
    padding: "10px 14px",
    fontFamily: "'DM Sans',sans-serif",
    fontSize: 14,
    color: "#1a1a2e",
    lineHeight: 1.6,
    resize: "none",
    minHeight: 54,
    outline: "none",
    transition: "border-color .2s,background .2s",
    boxSizing: "border-box",
  },
  divider: {
    height: 1.5,
    background: "linear-gradient(90deg,transparent,#d4cff0,transparent)",
    margin: "6px 0",
  },

  // disc pills
  discRow: { display: "flex", gap: 8, flexWrap: "wrap" },
  discPill: (active, style) => ({
    padding: "7px 14px",
    borderRadius: 100,
    border: `1.5px solid ${active && style ? style.border : "#d4cff0"}`,
    fontSize: 13,
    fontWeight: active ? 700 : 500,
    cursor: "pointer",
    transition: "all .2s",
    background: active && style ? style.bg : "#fff",
    color: active && style ? style.color : "#8b85b0",
    fontFamily: "'DM Sans',sans-serif",
  }),

  // buttons
  btnRow: {
    display: "flex",
    gap: 10,
    justifyContent: "flex-end",
    padding: "18px 28px 26px",
    flexWrap: "wrap",
  },
  btn: (variant) => {
    const base = {
      padding: "10px 22px",
      borderRadius: 100,
      fontFamily: "'DM Sans',sans-serif",
      fontSize: 13,
      fontWeight: 700,
      cursor: "pointer",
      border: "none",
      transition: "all .2s",
      letterSpacing: "0.02em",
    };
    if (variant === "primary")
      return {
        ...base,
        background: "#5c4fcf",
        color: "#fff",
        boxShadow: "0 4px 14px rgba(92,79,207,.3)",
      };
    if (variant === "green")
      return {
        ...base,
        background: "#2e7d32",
        color: "#fff",
        boxShadow: "0 4px 14px rgba(46,125,50,.25)",
      };
    return {
      ...base,
      background: "transparent",
      border: "1.5px solid #d4cff0",
      color: "#8b85b0",
    };
  },

  // completed preview
  previewHeader: {
    display: "flex",
    alignItems: "center",
    gap: 14,
    marginBottom: 22,
    paddingBottom: 18,
    borderBottom: "1.5px solid #d4cff0",
  },
  previewAvatar: {
    width: 48,
    height: 48,
    borderRadius: "50%",
    background: "linear-gradient(135deg,#5c4fcf,#9b8ff0)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#fff",
    fontSize: 20,
    flexShrink: 0,
  },
  previewName: {
    fontFamily: "'DM Serif Display',serif",
    fontSize: 22,
    color: "#1a1a2e",
  },
  grid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  previewField: (full) => ({
    background: "#f7f6fb",
    borderRadius: 12,
    padding: "11px 14px",
    gridColumn: full ? "1/-1" : "auto",
  }),
  previewFieldLabel: {
    fontSize: 9,
    fontWeight: 700,
    letterSpacing: "0.14em",
    textTransform: "uppercase",
    color: "#5c4fcf",
    marginBottom: 4,
    display: "flex",
    alignItems: "center",
    gap: 4,
  },
  previewFieldValue: (empty) => ({
    fontSize: 13,
    color: empty ? "#ccc" : "#1a1a2e",
    fontStyle: empty ? "italic" : "normal",
    lineHeight: 1.5,
  }),
  branding: {
    textAlign: "center",
    paddingTop: 14,
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.15em",
    textTransform: "uppercase",
    color: "#d4cff0",
  },

  // admin
  adminHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 20,
    flexWrap: "wrap",
    gap: 10,
  },
  adminTitle: {
    fontFamily: "'DM Serif Display',serif",
    fontSize: 22,
    color: "#1a1a2e",
  },
  countBadge: {
    background: "#ede9ff",
    color: "#5c4fcf",
    fontSize: 12,
    fontWeight: 700,
    padding: "4px 12px",
    borderRadius: 100,
  },
  responseCard: {
    background: "#fff",
    borderRadius: 16,
    marginBottom: 14,
    overflow: "hidden",
    boxShadow: "0 2px 12px rgba(92,79,207,.07)",
    border: "1.5px solid #ede9ff",
  },
  responseCardHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "14px 18px",
    cursor: "pointer",
    gap: 12,
  },
  responseCardLeft: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    flex: 1,
    minWidth: 0,
  },
  responseAvatar: (disc) => {
    const c =
      disc && DISC_COLORS[disc]
        ? DISC_COLORS[disc]
        : { bg: "#ede9ff", color: "#5c4fcf" };
    return {
      width: 38,
      height: 38,
      borderRadius: "50%",
      background: c.bg,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      color: c.color,
      fontSize: 15,
      fontWeight: 700,
      flexShrink: 0,
    };
  },
  responseName: {
    fontFamily: "'DM Serif Display',serif",
    fontSize: 17,
    color: "#1a1a2e",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  responseTime: { fontSize: 11, color: "#8b85b0", marginTop: 1 },
  discTag: (disc) => {
    const c =
      disc && DISC_COLORS[disc]
        ? DISC_COLORS[disc]
        : { bg: "#ede9ff", color: "#5c4fcf", border: "#5c4fcf" };
    return {
      padding: "4px 12px",
      borderRadius: 100,
      fontSize: 11,
      fontWeight: 700,
      background: c.bg,
      color: c.color,
      border: `1.5px solid ${c.border}`,
      whiteSpace: "nowrap",
    };
  },
  chevron: (open) => ({
    fontSize: 14,
    color: "#8b85b0",
    transform: open ? "rotate(180deg)" : "rotate(0deg)",
    transition: "transform .2s",
    flexShrink: 0,
  }),
  responseBody: {
    padding: "0 18px 16px",
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 10,
  },
  responseField: (full) => ({
    background: "#f7f6fb",
    borderRadius: 10,
    padding: "10px 13px",
    gridColumn: full ? "1/-1" : "auto",
  }),
  responseFieldLabel: {
    fontSize: 9,
    fontWeight: 700,
    letterSpacing: "0.13em",
    textTransform: "uppercase",
    color: "#5c4fcf",
    marginBottom: 4,
  },
  responseFieldValue: { fontSize: 13, color: "#1a1a2e", lineHeight: 1.5 },
  deleteBtn: {
    background: "none",
    border: "none",
    color: "#d4cff0",
    fontSize: 16,
    cursor: "pointer",
    padding: "2px 6px",
    borderRadius: 6,
    flexShrink: 0,
  },
  emptyState: { textAlign: "center", padding: "48px 20px", color: "#8b85b0" },
  emptyIcon: { fontSize: 40, marginBottom: 12 },
  emptyTitle: {
    fontFamily: "'DM Serif Display',serif",
    fontSize: 20,
    color: "#1a1a2e",
    marginBottom: 6,
  },
  emptyText: { fontSize: 14, lineHeight: 1.6 },
};

// ─────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("form"); // "form" | "admin"
  const [form, setForm] = useState(EMPTY_FORM);
  const [stage, setStage] = useState("fill"); // "fill" | "preview"
  const [responses, setResponses] = useState([]);
  const [expanded, setExpanded] = useState(null);
  const [nameError, setNameError] = useState(false);
  const previewRef = useRef(null);

  // Load from storage on mount
  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("buom-responses");
        if (r && r.value) setResponses(JSON.parse(r.value));
      } catch (_) {}
    })();
  }, []);

  // Save to storage whenever responses change
  useEffect(() => {
    if (responses.length === 0) return;
    (async () => {
      try {
        await window.storage.set("buom-responses", JSON.stringify(responses));
      } catch (_) {}
    })();
  }, [responses]);

  const update = (id, v) => setForm((f) => ({ ...f, [id]: v }));

  function submitForm() {
    if (!form.name.trim()) {
      setNameError(true);
      setTimeout(() => setNameError(false), 1500);
      return;
    }
    setStage("preview");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function saveResponse() {
    const entry = {
      ...form,
      id: Date.now(),
      submittedAt: new Date().toLocaleString(),
    };
    const updated = [entry, ...responses];
    setResponses(updated);
    // Also persist immediately
    (async () => {
      try {
        await window.storage.set("buom-responses", JSON.stringify(updated));
      } catch (_) {}
    })();
    setForm(EMPTY_FORM);
    setStage("fill");
    setTab("admin");
  }

  function deleteResponse(id) {
    const updated = responses.filter((r) => r.id !== id);
    setResponses(updated);
    (async () => {
      try {
        if (updated.length === 0) await window.storage.delete("buom-responses");
        else
          await window.storage.set("buom-responses", JSON.stringify(updated));
      } catch (_) {}
    })();
  }

  function exportExcel() {
    const headers = [
      "Name",
      "DiSC Style",
      "Top Strengths",
      "What Energises Me",
      "What Drains Me",
      "Pressure Signs",
      "Best Way to Approach",
      "Best Way to Challenge",
      "One Thing to Know",
      "Submitted At",
    ];
    const rows = responses.map((r) => [
      r.name,
      r.disc ? DISC_LABELS[r.disc] : "",
      r.strengths,
      r.energises,
      r.drains,
      r.pressure,
      r.approach,
      r.challenge,
      r.onemore,
      r.submittedAt,
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws["!cols"] = [
      { wch: 20 },
      { wch: 18 },
      { wch: 35 },
      { wch: 35 },
      { wch: 35 },
      { wch: 35 },
      { wch: 35 },
      { wch: 35 },
      { wch: 40 },
      { wch: 22 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Best Use of Me");
    XLSX.writeFile(
      wb,
      `BestUseOfMe_Responses_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
  }

  // ── Render form ──
  function FormView() {
    return (
      <>
        <div style={s.card}>
          <div style={s.stripe} />
          <div style={s.cardBody}>
            {/* Name */}
            <div style={s.nameRow}>
              <div style={s.avatar}>👤</div>
              <div style={{ flex: 1 }}>
                <label style={s.nameLabel}>Your name</label>
                <input
                  style={{
                    ...s.nameInput,
                    borderBottomColor: nameError ? "#e05c5c" : "#d4cff0",
                  }}
                  value={form.name}
                  onChange={(e) => update("name", e.target.value)}
                  placeholder="e.g. Alex Johnson"
                />
              </div>
            </div>

            {/* DiSC */}
            <div style={{ marginBottom: 18 }}>
              <div style={s.fieldLabel}>
                <span>🎯</span> My DiSC style
              </div>
              <div style={s.discRow}>
                {Object.entries(DISC_LABELS).map(([k, v]) => (
                  <button
                    key={k}
                    style={s.discPill(form.disc === k, DISC_COLORS[k])}
                    onClick={() => update("disc", form.disc === k ? "" : k)}
                  >
                    {v}
                  </button>
                ))}
              </div>
            </div>

            <div style={s.divider} />

            {/* Fields */}
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                gap: 16,
                marginTop: 16,
              }}
            >
              {FIELDS.map((f, i) => (
                <div key={f.id}>
                  {(i === 3 || i === 6) && (
                    <div style={{ ...s.divider, marginBottom: 16 }} />
                  )}
                  <div style={s.fieldLabel}>
                    <span>{f.icon}</span> {f.label}
                  </div>
                  <textarea
                    style={s.textarea}
                    value={form[f.id]}
                    onChange={(e) => update(f.id, e.target.value)}
                    placeholder={f.placeholder}
                    rows={2}
                    onInput={(e) => {
                      e.target.style.height = "auto";
                      e.target.style.height = e.target.scrollHeight + "px";
                    }}
                  />
                </div>
              ))}
            </div>
          </div>
          <div style={s.btnRow}>
            <button style={s.btn("ghost")} onClick={() => setForm(EMPTY_FORM)}>
              Clear
            </button>
            <button style={s.btn("primary")} onClick={submitForm}>
              Preview card →
            </button>
          </div>
        </div>
      </>
    );
  }

  // ── Render preview ──
  function PreviewView() {
    const disc = form.disc;
    const dc = disc ? DISC_COLORS[disc] : null;
    return (
      <>
        <div
          style={{
            display: "flex",
            gap: 10,
            justifyContent: "center",
            flexWrap: "wrap",
            marginBottom: 20,
          }}
        >
          <button style={s.btn("ghost")} onClick={() => setStage("fill")}>
            ← Edit
          </button>
          <button style={{ ...s.btn("primary") }} onClick={saveResponse}>
            ✓ Submit & save response
          </button>
        </div>

        <div ref={previewRef} style={s.card} id="previewCard">
          <div style={s.stripe} />
          <div style={s.cardBody}>
            <div style={s.previewHeader}>
              <div style={s.previewAvatar}>✦</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    letterSpacing: "0.15em",
                    textTransform: "uppercase",
                    color: "#8b85b0",
                    marginBottom: 3,
                  }}
                >
                  Best Use of Me
                </div>
                <div style={s.previewName}>{form.name}</div>
              </div>
              {disc && (
                <div
                  style={{
                    padding: "5px 14px",
                    borderRadius: 100,
                    fontSize: 12,
                    fontWeight: 700,
                    background: dc.bg,
                    color: dc.color,
                    border: `1.5px solid ${dc.border}`,
                    whiteSpace: "nowrap",
                  }}
                >
                  {DISC_LABELS[disc]}
                </div>
              )}
            </div>

            <div style={s.grid}>
              {FIELDS.map((f) => (
                <div key={f.id} style={s.previewField(f.full)}>
                  <div style={s.previewFieldLabel}>
                    <span>{f.icon}</span> {f.label}
                  </div>
                  <div style={s.previewFieldValue(!form[f.id])}>
                    {form[f.id] || "Not filled in"}
                  </div>
                </div>
              ))}
            </div>
            <div style={s.branding}>Best Use of Me · Team Workshop</div>
          </div>
        </div>
      </>
    );
  }

  // ── Render admin ──
  function AdminView() {
    return (
      <div style={s.card}>
        <div style={s.stripe} />
        <div style={{ padding: "24px 24px 6px" }}>
          <div style={s.adminHeader}>
            <div>
              <div style={s.eyebrow}>Facilitator View</div>
              <div style={s.adminTitle}>All responses</div>
            </div>
            <div
              style={{
                display: "flex",
                gap: 10,
                alignItems: "center",
                flexWrap: "wrap",
              }}
            >
              <div style={s.countBadge}>
                {responses.length}{" "}
                {responses.length === 1 ? "response" : "responses"}
              </div>
              {responses.length > 0 && (
                <button style={s.btn("green")} onClick={exportExcel}>
                  📊 Export all to Excel
                </button>
              )}
            </div>
          </div>
        </div>

        <div style={{ padding: "0 16px 24px" }}>
          {responses.length === 0 ? (
            <div style={s.emptyState}>
              <div style={s.emptyIcon}>📋</div>
              <div style={s.emptyTitle}>No responses yet</div>
              <div style={s.emptyText}>
                Switch to the Form tab and fill in the first card.
                <br />
                All submissions will appear here.
              </div>
            </div>
          ) : (
            responses.map((r) => {
              const open = expanded === r.id;
              const disc = r.disc;
              const dc = disc ? DISC_COLORS[disc] : null;
              return (
                <div key={r.id} style={s.responseCard}>
                  <div
                    style={s.responseCardHeader}
                    onClick={() => setExpanded(open ? null : r.id)}
                  >
                    <div style={s.responseCardLeft}>
                      <div style={s.responseAvatar(disc)}>
                        {r.name?.[0]?.toUpperCase() || "?"}
                      </div>
                      <div style={{ minWidth: 0 }}>
                        <div style={s.responseName}>{r.name}</div>
                        <div style={s.responseTime}>{r.submittedAt}</div>
                      </div>
                    </div>
                    {disc && (
                      <div style={s.discTag(disc)}>{DISC_LABELS[disc]}</div>
                    )}
                    <div style={s.chevron(open)}>▼</div>
                    <button
                      style={s.deleteBtn}
                      onClick={(e) => {
                        e.stopPropagation();
                        deleteResponse(r.id);
                      }}
                      title="Delete"
                    >
                      ✕
                    </button>
                  </div>

                  {open && (
                    <div style={s.responseBody}>
                      {FIELDS.map((f) => (
                        <div key={f.id} style={s.responseField(f.full)}>
                          <div style={s.responseFieldLabel}>
                            {f.icon} {f.label}
                          </div>
                          <div
                            style={{
                              ...s.responseFieldValue,
                              color: r[f.id] ? "#1a1a2e" : "#ccc",
                              fontStyle: r[f.id] ? "normal" : "italic",
                            }}
                          >
                            {r[f.id] || "—"}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  }

  return (
    <>
      <link
        href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap"
        rel="stylesheet"
      />
      <div style={s.app}>
        <div style={s.center}>
          <div style={s.header}>
            <div style={s.eyebrow}>Team Workshop</div>
            <h1 style={s.h1}>
              Best Use of <em style={s.h1em}>Me</em>
            </h1>
            <p style={s.subtitle}>
              Fill in your card so your teammates know how to work with you at
              your best.
            </p>
          </div>

          {/* Tabs */}
          <div style={s.tabs}>
            <button
              style={s.tab(tab === "form")}
              onClick={() => {
                setTab("form");
              }}
            >
              📝 Fill in my card
            </button>
            <button
              style={s.tab(tab === "admin")}
              onClick={() => setTab("admin")}
            >
              📋 All responses{" "}
              {responses.length > 0 ? `(${responses.length})` : ""}
            </button>
          </div>

          {tab === "form" && stage === "fill" && <FormView />}
          {tab === "form" && stage === "preview" && <PreviewView />}
          {tab === "admin" && <AdminView />}
        </div>
      </div>
    </>
  );
}
