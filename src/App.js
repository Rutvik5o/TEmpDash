import { useState, useCallback, useMemo } from "react";
import {
  BarChart, Bar, PieChart, Pie, Cell,
  ScatterChart, Scatter, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";
import Papa from "papaparse";

/* ─── DESIGN TOKENS ─── */
const T = {
  bg:       "#0a0b0f",
  surface:  "#111318",
  card:     "#161820",
  border:   "#1e2130",
  accent:   "#00f5c4",
  accent2:  "#7c6fff",
  accent3:  "#ff6b6b",
  amber:    "#ffb347",
  text:     "#e8eaf0",
  muted:    "#5a607a",
  dim:      "#2a2f45",
};

const PALETTE = ["#00f5c4","#7c6fff","#ff6b6b","#ffb347","#38bdf8","#f472b6","#a3e635","#fb923c"];

/* ─── GLOBAL STYLES injected once ─── */
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${T.bg}; color: ${T.text}; font-family: 'Syne', sans-serif; }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: ${T.bg}; }
  ::-webkit-scrollbar-thumb { background: ${T.dim}; border-radius: 3px; }
  input[type=file] { display: none; }
  @keyframes fadeUp {
    from { opacity:0; transform:translateY(24px); }
    to   { opacity:1; transform:translateY(0); }
  }
  @keyframes pulse {
    0%,100% { box-shadow: 0 0 0 0 rgba(0,245,196,0.4); }
    50%      { box-shadow: 0 0 0 12px rgba(0,245,196,0); }
  }
  @keyframes scanline {
    0%   { transform: translateY(-100%); }
    100% { transform: translateY(100vh); }
  }
  @keyframes blink { 0%,100%{opacity:1} 50%{opacity:0} }
  @keyframes shimmer {
    0%   { background-position: -200% center; }
    100% { background-position:  200% center; }
  }
  .fadeUp { animation: fadeUp 0.6s ease both; }
  .card {
    background: ${T.card};
    border: 1px solid ${T.border};
    border-radius: 16px;
    padding: 24px;
    position: relative;
    overflow: hidden;
  }
  .card::before {
    content:'';
    position:absolute; top:0; left:0; right:0; height:1px;
    background: linear-gradient(90deg, transparent, ${T.accent}40, transparent);
  }
  .btn {
    background: transparent;
    border: 1px solid ${T.accent};
    color: ${T.accent};
    padding: 10px 22px;
    border-radius: 8px;
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    cursor: pointer;
    transition: all 0.2s;
    letter-spacing: 0.5px;
  }
  .btn:hover { background: ${T.accent}18; box-shadow: 0 0 16px ${T.accent}40; }
  .btn-solid {
    background: ${T.accent};
    color: #000;
    border: none;
    font-weight: 500;
  }
  .btn-solid:hover { background: #00ddb0; box-shadow: 0 0 24px ${T.accent}60; }
  .tag {
    display: inline-flex; align-items: center; gap: 6px;
    background: ${T.dim}; border-radius: 6px;
    padding: 4px 10px; font-size: 11px;
    font-family: 'DM Mono', monospace;
    color: ${T.muted}; border: 1px solid ${T.border};
  }
  .section-title {
    font-size: 11px; font-family: 'DM Mono', monospace;
    color: ${T.accent}; letter-spacing: 2px; text-transform: uppercase;
    margin-bottom: 16px; display: flex; align-items: center; gap: 8px;
  }
  .section-title::after {
    content:''; flex:1; height:1px;
    background: linear-gradient(90deg, ${T.accent}40, transparent);
  }
  .insight-chip {
    background: ${T.dim}; border: 1px solid ${T.border};
    border-radius: 10px; padding: 12px 16px;
    font-size: 13px; color: ${T.text}; line-height: 1.5;
    border-left: 3px solid ${T.accent};
  }
  .stat-num {
    font-size: 28px; font-weight: 800; color: ${T.accent};
    font-family: 'DM Mono', monospace;
  }
  .stat-label {
    font-size: 11px; color: ${T.muted}; text-transform: uppercase;
    letter-spacing: 1px; margin-top: 4px;
  }
`;

/* ─── UTILITIES ─── */
const num = (v) => parseFloat(v);
const isNum = (v) => !isNaN(num(v)) && v !== "" && v !== null;

function calcStats(data, col) {
  const vals = data.map(r => num(r[col])).filter(v => !isNaN(v));
  if (!vals.length) return null;
  const sorted = [...vals].sort((a, b) => a - b);
  const sum = vals.reduce((a, b) => a + b, 0);
  const mean = sum / vals.length;
  const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length;
  const median = sorted.length % 2 === 0
    ? (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2
    : sorted[Math.floor(sorted.length / 2)];
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  return {
    count: vals.length, sum, mean, median,
    min: sorted[0], max: sorted[sorted.length - 1],
    std: Math.sqrt(variance), q1, q3,
    iqr: q3 - q1,
    range: sorted[sorted.length - 1] - sorted[0],
  };
}

function detectOutliers(data, col) {
  const s = calcStats(data, col);
  if (!s) return [];
  const lo = s.q1 - 1.5 * s.iqr, hi = s.q3 + 1.5 * s.iqr;
  return data.filter(r => { const v = num(r[col]); return !isNaN(v) && (v < lo || v > hi); });
}

function correlation(data, colA, colB) {
  const pairs = data.map(r => [num(r[colA]), num(r[colB])]).filter(([a, b]) => !isNaN(a) && !isNaN(b));
  if (pairs.length < 2) return 0;
  const meanA = pairs.reduce((s, [a]) => s + a, 0) / pairs.length;
  const meanB = pairs.reduce((s, [, b]) => s + b, 0) / pairs.length;
  const num_ = pairs.reduce((s, [a, b]) => s + (a - meanA) * (b - meanB), 0);
  const den = Math.sqrt(
    pairs.reduce((s, [a]) => s + (a - meanA) ** 2, 0) *
    pairs.reduce((s, [, b]) => s + (b - meanB) ** 2, 0)
  );
  return den === 0 ? 0 : num_ / den;
}

function buildHistogram(data, col, bins = 10) {
  const vals = data.map(r => num(r[col])).filter(v => !isNaN(v));
  if (!vals.length) return [];
  const min = Math.min(...vals), max = Math.max(...vals);
  const size = (max - min) / bins || 1;
  const hist = Array.from({ length: bins }, (_, i) => ({
    range: `${(min + i * size).toFixed(1)}–${(min + (i + 1) * size).toFixed(1)}`,
    count: 0,
  }));
  vals.forEach(v => {
    const idx = Math.min(Math.floor((v - min) / size), bins - 1);
    hist[idx].count++;
  });
  return hist;
}

function categoryCounts(data, col, limit = 12) {
  const counts = {};
  data.forEach(r => { const v = r[col] ?? "N/A"; counts[v] = (counts[v] || 0) + 1; });
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([name, value]) => ({ name, value, pct: ((value / data.length) * 100).toFixed(1) }));
}

function autoInsights(data, numericCols, textCols) {
  const insights = [];
  numericCols.forEach(col => {
    const s = calcStats(data, col);
    if (!s) return;
    const outliers = detectOutliers(data, col);
    if (outliers.length > 0)
      insights.push(`⚠️ ${col}: ${outliers.length} outlier${outliers.length > 1 ? "s" : ""} detected (beyond IQR bounds)`);
    const cv = (s.std / Math.abs(s.mean)) * 100;
    if (cv > 50) insights.push(`📊 ${col} is highly variable — CV ${cv.toFixed(0)}%, values spread widely`);
    if (s.mean > s.median * 1.2) insights.push(`📈 ${col} is right-skewed — mean (${s.mean.toFixed(2)}) > median (${s.median.toFixed(2)})`);
    if (s.mean < s.median * 0.8) insights.push(`📉 ${col} is left-skewed — mean (${s.mean.toFixed(2)}) < median (${s.median.toFixed(2)})`);
  });
  for (let i = 0; i < numericCols.length; i++) {
    for (let j = i + 1; j < numericCols.length; j++) {
      const r = correlation(data, numericCols[i], numericCols[j]);
      if (Math.abs(r) > 0.7)
        insights.push(`🔗 Strong ${r > 0 ? "positive" : "negative"} correlation (${r.toFixed(2)}) between ${numericCols[i]} & ${numericCols[j]}`);
    }
  }
  textCols.forEach(col => {
    const cats = categoryCounts(data, col);
    if (cats.length === 1) insights.push(`ℹ️ ${col}: only one unique value — not useful for analysis`);
    if (cats[0] && parseFloat(cats[0].pct) > 70) insights.push(`🏆 ${col}: "${cats[0].name}" dominates at ${cats[0].pct}% of records`);
    const missing = data.filter(r => !r[col] || r[col].trim() === "").length;
    if (missing > 0) insights.push(`🚨 ${col}: ${missing} missing / empty values (${((missing / data.length) * 100).toFixed(1)}%)`);
  });
  if (!insights.length) insights.push("✅ Data looks clean — no major anomalies detected");
  return insights;
}

/* ─── TOOLTIP ─── */
const DarkTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px", fontFamily: "'DM Mono', monospace", fontSize: 12 }}>
      {label && <div style={{ color: T.muted, marginBottom: 6 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || T.accent }}>
          {p.name}: <strong>{typeof p.value === "number" ? p.value.toLocaleString(undefined, { maximumFractionDigits: 2 }) : p.value}</strong>
        </div>
      ))}
    </div>
  );
};

/* ─── PASSWORD SCREEN ─── */
function PasswordScreen({ onUnlock }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const [shake, setShake] = useState(false);

  const attempt = () => {
    if (pw === "Foundation") { onUnlock(); }
    else {
      setErr(true); setShake(true);
      setTimeout(() => setShake(false), 500);
    }
  };

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: T.bg, position: "relative", overflow: "hidden",
    }}>
      {/* animated background grid */}
      <div style={{
        position: "absolute", inset: 0,
        backgroundImage: `linear-gradient(${T.dim}22 1px, transparent 1px), linear-gradient(90deg, ${T.dim}22 1px, transparent 1px)`,
        backgroundSize: "48px 48px",
      }} />
      {/* glow */}
      <div style={{
        position: "absolute", width: 500, height: 500, borderRadius: "50%",
        background: `radial-gradient(circle, ${T.accent}12 0%, transparent 70%)`,
        top: "50%", left: "50%", transform: "translate(-50%,-50%)",
      }} />

      <div className="fadeUp" style={{
        position: "relative", zIndex: 1, textAlign: "center",
        animation: shake ? "none" : undefined,
        transform: shake ? "translateX(-8px)" : "none",
        transition: "transform 0.1s",
      }}>
        {/* logo */}
        <div style={{
          width: 72, height: 72, borderRadius: 20, margin: "0 auto 28px",
          background: `linear-gradient(135deg, ${T.accent}22, ${T.accent2}22)`,
          border: `1px solid ${T.accent}40`,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 32,
        }}>📊</div>

        <h1 style={{ fontSize: 36, fontWeight: 800, letterSpacing: -1, marginBottom: 6 }}>
          <span style={{ color: T.accent }}>Data</span>Scope
        </h1>
        <p style={{ color: T.muted, fontSize: 13, fontFamily: "'DM Mono', monospace", marginBottom: 40 }}>
          ANALYTICS PLATFORM · RESTRICTED ACCESS
        </p>

        <div style={{
          background: T.card, border: `1px solid ${err ? T.accent3 : T.border}`,
          borderRadius: 20, padding: "36px 40px", width: 380,
          boxShadow: err ? `0 0 24px ${T.accent3}30` : `0 0 40px #00000080`,
          transition: "border-color 0.3s, box-shadow 0.3s",
        }}>
          <div style={{ fontSize: 12, color: T.muted, fontFamily: "'DM Mono', monospace", marginBottom: 20, textAlign: "left" }}>
            ACCESS CODE
          </div>
          <input
            type="password"
            value={pw}
            onChange={e => { setPw(e.target.value); setErr(false); }}
            onKeyDown={e => e.key === "Enter" && attempt()}
            placeholder="Enter password"
            autoFocus
            style={{
              width: "100%", padding: "14px 18px",
              background: T.surface, border: `1px solid ${err ? T.accent3 : T.dim}`,
              borderRadius: 10, color: T.text,
              fontFamily: "'DM Mono', monospace", fontSize: 16,
              outline: "none", marginBottom: 8,
              letterSpacing: pw ? 4 : 0,
              transition: "border-color 0.2s",
            }}
          />
          {err && (
            <div style={{ color: T.accent3, fontSize: 12, fontFamily: "'DM Mono', monospace", marginBottom: 12, textAlign: "left" }}>
              ✗ Incorrect password. Try again.
            </div>
          )}
          <button
            onClick={attempt}
            className="btn btn-solid"
            style={{ width: "100%", padding: "13px", fontSize: 14, borderRadius: 10, marginTop: 8, fontFamily: "'Syne', sans-serif", fontWeight: 700 }}
          >
            UNLOCK DASHBOARD →
          </button>
        </div>

        <p style={{ color: T.dim, fontSize: 11, fontFamily: "'DM Mono', monospace", marginTop: 24 }}>
          Unauthorized access is strictly prohibited
        </p>
      </div>
    </div>
  );
}

/* ─── STAT CARD ─── */
function StatCard({ label, value, sub, color, delay = 0 }) {
  return (
    <div className="card fadeUp" style={{ animationDelay: `${delay}ms` }}>
      <div className="stat-num" style={{ color: color || T.accent }}>{value}</div>
      <div className="stat-label">{label}</div>
      {sub && <div style={{ fontSize: 11, color: T.muted, marginTop: 6, fontFamily: "'DM Mono', monospace" }}>{sub}</div>}
    </div>
  );
}

/* ─── CHART SECTION ─── */
function ChartSection({ title, accent, children }) {
  return (
    <div className="card fadeUp" style={{ marginBottom: 24 }}>
      <div className="section-title" style={{ color: accent || T.accent }}>{title}</div>
      {children}
    </div>
  );
}

/* ─── MAIN DASHBOARD ─── */
function Dashboard() {
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [fileName, setFileName] = useState("");
  const [dragging, setDragging] = useState(false);
  const [activeTab, setActiveTab] = useState("overview");

  const processFile = (file) => {
    setFileName(file.name);
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (result) => { setData(result.data); setColumns(result.meta.fields || []); }
    });
  };

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, []);

  const numericCols = useMemo(() => columns.filter(c => data.length > 0 && isNum(data[0][c])), [columns, data]);
  const textCols = useMemo(() => columns.filter(c => !numericCols.includes(c)), [columns, numericCols]);
  const insights = useMemo(() => data.length ? autoInsights(data, numericCols, textCols) : [], [data, numericCols, textCols]);

  const TABS = ["overview", "distributions", "correlations", "insights"];

  if (!data.length) {
    return (
      <div style={{ minHeight: "100vh", background: T.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24 }}>
        <div style={{ position: "absolute", inset: 0, backgroundImage: `linear-gradient(${T.dim}22 1px, transparent 1px), linear-gradient(90deg, ${T.dim}22 1px, transparent 1px)`, backgroundSize: "48px 48px" }} />
        <div className="fadeUp" style={{ position: "relative", zIndex: 1, textAlign: "center" }}>
          <h1 style={{ fontSize: 42, fontWeight: 800, marginBottom: 8, letterSpacing: -1 }}>
            <span style={{ color: T.accent }}>Data</span>Scope
          </h1>
          <p style={{ color: T.muted, fontFamily: "'DM Mono', monospace", fontSize: 13, marginBottom: 48 }}>
            ADVANCED CSV ANALYTICS DASHBOARD
          </p>
          <div
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            style={{
              border: `2px dashed ${dragging ? T.accent : T.dim}`,
              borderRadius: 24, background: dragging ? `${T.accent}08` : T.card,
              padding: "64px 80px", cursor: "pointer", transition: "all 0.3s",
              boxShadow: dragging ? `0 0 40px ${T.accent}30` : "none",
              animation: "pulse 3s infinite",
            }}
          >
            <div style={{ fontSize: 56, marginBottom: 20 }}>⬆</div>
            <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 8 }}>Drop CSV file here</h2>
            <p style={{ color: T.muted, fontSize: 13, marginBottom: 28 }}>or click to browse your files</p>
            <label className="btn btn-solid" style={{ padding: "13px 32px", borderRadius: 10, cursor: "pointer", fontSize: 14 }}>
              CHOOSE FILE
              <input type="file" accept=".csv" onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
            </label>
          </div>
          <div style={{ display: "flex", gap: 16, justifyContent: "center", marginTop: 32 }}>
            {["Auto-detect types", "Outlier detection", "Correlation matrix", "Smart insights"].map(f => (
              <span className="tag" key={f}>{f}</span>
            ))}
          </div>
        </div>
      </div>
    );
  }

  const chartData = data.slice(0, 40);

  return (
    <div style={{ minHeight: "100vh", background: T.bg, color: T.text }}>
      {/* header */}
      <div style={{
        background: `${T.surface}ee`, backdropFilter: "blur(12px)",
        borderBottom: `1px solid ${T.border}`, padding: "0 32px",
        position: "sticky", top: 0, zIndex: 100,
        display: "flex", alignItems: "center", gap: 24,
      }}>
        <div style={{ fontWeight: 800, fontSize: 20, paddingBlock: 16 }}>
          <span style={{ color: T.accent }}>Data</span>Scope
        </div>
        <div style={{ flex: 1, display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 12, color: T.muted, fontFamily: "'DM Mono', monospace" }}>▸</span>
          <span style={{ fontSize: 13, color: T.text, fontFamily: "'DM Mono', monospace" }}>{fileName}</span>
          <span className="tag">{data.length} rows</span>
          <span className="tag">{columns.length} cols</span>
        </div>
        <button className="btn" onClick={() => { setData([]); setColumns([]); setFileName(""); }} style={{ fontSize: 12 }}>
          + New File
        </button>
      </div>

      {/* tabs */}
      <div style={{ borderBottom: `1px solid ${T.border}`, padding: "0 32px", display: "flex", gap: 0 }}>
        {TABS.map(t => (
          <button key={t} onClick={() => setActiveTab(t)} style={{
            background: "none", border: "none", cursor: "pointer",
            padding: "14px 20px", fontSize: 13, fontFamily: "'DM Mono', monospace",
            color: activeTab === t ? T.accent : T.muted,
            borderBottom: `2px solid ${activeTab === t ? T.accent : "transparent"}`,
            textTransform: "uppercase", letterSpacing: 1, transition: "all 0.2s",
          }}>
            {t}
          </button>
        ))}
      </div>

      <div style={{ padding: "32px", maxWidth: 1300, margin: "0 auto" }}>

        {/* ── OVERVIEW TAB ── */}
        {activeTab === "overview" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16, marginBottom: 32 }}>
              <StatCard label="Total Records" value={data.length.toLocaleString()} delay={0} />
              <StatCard label="Total Columns" value={columns.length} delay={60} />
              <StatCard label="Numeric Columns" value={numericCols.length} color={T.accent2} delay={120} />
              <StatCard label="Category Columns" value={textCols.length} color={T.amber} delay={180} />
              <StatCard label="Auto Insights" value={insights.length} color={T.accent3} sub="anomalies + patterns" delay={240} />
            </div>

            {numericCols.length > 0 && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16, marginBottom: 32 }}>
                {numericCols.slice(0, 4).map((col, ci) => {
                  const s = calcStats(data, col);
                  if (!s) return null;
                  return (
                    <div key={col} className="card fadeUp" style={{ animationDelay: `${ci * 80}ms` }}>
                      <div style={{ fontSize: 12, color: PALETTE[ci], fontFamily: "'DM Mono', monospace", marginBottom: 12, textTransform: "uppercase", letterSpacing: 1 }}>
                        ◈ {col}
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                        {[
                          ["Mean", s.mean.toFixed(2)],
                          ["Median", s.median.toFixed(2)],
                          ["Std Dev", s.std.toFixed(2)],
                          ["Range", s.range.toFixed(2)],
                          ["Min", s.min.toFixed(2)],
                          ["Max", s.max.toFixed(2)],
                        ].map(([k, v]) => (
                          <div key={k} style={{ background: T.surface, borderRadius: 8, padding: "8px 12px" }}>
                            <div style={{ fontSize: 10, color: T.muted, marginBottom: 2 }}>{k}</div>
                            <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "'DM Mono', monospace", color: PALETTE[ci] }}>{v}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {numericCols[0] && (
              <ChartSection title={`${numericCols[0]} — Trend View`} accent={T.accent}>
                <ResponsiveContainer width="100%" height={260}>
                  <AreaChart data={chartData}>
                    <defs>
                      <linearGradient id="g1" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={T.accent} stopOpacity={0.3} />
                        <stop offset="95%" stopColor={T.accent} stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                    <XAxis dataKey={textCols[0] || columns[0]} tick={{ fontSize: 10, fill: T.muted }} hide={chartData.length > 20} />
                    <YAxis tick={{ fontSize: 10, fill: T.muted }} />
                    <Tooltip content={<DarkTooltip />} />
                    <Area type="monotone" dataKey={numericCols[0]} stroke={T.accent} fill="url(#g1)" strokeWidth={2} dot={false} />
                  </AreaChart>
                </ResponsiveContainer>
              </ChartSection>
            )}

            {numericCols[1] && (
              <ChartSection title={`${numericCols[1]} — Bar Chart`} accent={T.accent2}>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                    <XAxis dataKey={textCols[0] || columns[0]} tick={{ fontSize: 10, fill: T.muted }} hide={chartData.length > 20} />
                    <YAxis tick={{ fontSize: 10, fill: T.muted }} />
                    <Tooltip content={<DarkTooltip />} />
                    <Bar dataKey={numericCols[1]} fill={T.accent2} radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartSection>
            )}
          </>
        )}

        {/* ── DISTRIBUTIONS TAB ── */}
        {activeTab === "distributions" && (
          <>
            {numericCols.map((col, ci) => {
              const hist = buildHistogram(data, col);
              const s = calcStats(data, col);
              return (
                <ChartSection key={col} title={`${col} — Histogram`} accent={PALETTE[ci % PALETTE.length]}>
                  {s && (
                    <div style={{ display: "flex", gap: 24, marginBottom: 20, flexWrap: "wrap" }}>
                      {[["Mean", s.mean.toFixed(2)], ["Median", s.median.toFixed(2)], ["Std Dev", s.std.toFixed(2)], ["Skew", s.mean > s.median ? "Right ↗" : "Left ↙"]].map(([k, v]) => (
                        <div key={k}>
                          <div style={{ fontSize: 10, color: T.muted }}>{k}</div>
                          <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'DM Mono', monospace", color: PALETTE[ci % PALETTE.length] }}>{v}</div>
                        </div>
                      ))}
                    </div>
                  )}
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={hist}>
                      <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                      <XAxis dataKey="range" tick={{ fontSize: 9, fill: T.muted }} />
                      <YAxis tick={{ fontSize: 10, fill: T.muted }} />
                      <Tooltip content={<DarkTooltip />} />
                      <Bar dataKey="count" fill={PALETTE[ci % PALETTE.length]} radius={[3, 3, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </ChartSection>
              );
            })}

            {textCols.map((col, ci) => {
              const cats = categoryCounts(data, col);
              return (
                <ChartSection key={col} title={`${col} — Category Breakdown`} accent={PALETTE[(ci + 3) % PALETTE.length]}>
                  <div style={{ display: "flex", gap: 32, alignItems: "center", flexWrap: "wrap" }}>
                    <ResponsiveContainer width={220} height={220}>
                      <PieChart>
                        <Pie data={cats} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} innerRadius={40}>
                          {cats.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
                        </Pie>
                        <Tooltip content={<DarkTooltip />} />
                      </PieChart>
                    </ResponsiveContainer>
                    <div style={{ flex: 1, minWidth: 200 }}>
                      {cats.map((d, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                          <div style={{ width: 10, height: 10, borderRadius: "50%", background: PALETTE[i % PALETTE.length], flexShrink: 0 }} />
                          <div style={{ flex: 1, fontSize: 13, color: T.text }}>{d.name}</div>
                          <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 13, color: PALETTE[i % PALETTE.length], fontWeight: 700 }}>{d.pct}%</div>
                          <div style={{ background: T.surface, borderRadius: 4, padding: "2px 8px", fontSize: 11, color: T.muted, fontFamily: "'DM Mono', monospace" }}>{d.value}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </ChartSection>
              );
            })}
          </>
        )}

        {/* ── CORRELATIONS TAB ── */}
        {activeTab === "correlations" && (
          <>
            {numericCols.length >= 2 ? (
              <>
                <div className="card fadeUp" style={{ marginBottom: 24 }}>
                  <div className="section-title">Correlation Matrix</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", fontFamily: "'DM Mono', monospace", fontSize: 12 }}>
                      <thead>
                        <tr>
                          <th style={{ padding: "8px 14px", color: T.muted, textAlign: "left" }}></th>
                          {numericCols.map(c => (
                            <th key={c} style={{ padding: "8px 14px", color: T.muted, fontWeight: 500, whiteSpace: "nowrap" }}>{c.length > 10 ? c.slice(0, 10) + "…" : c}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {numericCols.map(rowCol => (
                          <tr key={rowCol}>
                            <td style={{ padding: "8px 14px", color: T.muted, whiteSpace: "nowrap" }}>{rowCol.length > 10 ? rowCol.slice(0, 10) + "…" : rowCol}</td>
                            {numericCols.map(colCol => {
                              const r = rowCol === colCol ? 1 : correlation(data, rowCol, colCol);
                              const abs = Math.abs(r);
                              const bg = r === 1 ? T.dim
                                : r > 0.7 ? `${T.accent}40`
                                : r > 0.4 ? `${T.accent}20`
                                : r < -0.7 ? `${T.accent3}40`
                                : r < -0.4 ? `${T.accent3}20`
                                : T.surface;
                              return (
                                <td key={colCol} style={{
                                  padding: "8px 14px", textAlign: "center",
                                  background: bg, borderRadius: 4,
                                  color: abs > 0.7 ? T.text : T.muted,
                                  fontWeight: abs > 0.7 ? 700 : 400,
                                }}>
                                  {r.toFixed(2)}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{ marginTop: 16, display: "flex", gap: 16, flexWrap: "wrap" }}>
                    {[["Strong +", T.accent, "> 0.7"], ["Moderate +", `${T.accent}60`, "0.4–0.7"], ["Strong −", T.accent3, "< −0.7"]].map(([l, c, r]) => (
                      <div key={l} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, color: T.muted }}>
                        <div style={{ width: 12, height: 12, borderRadius: 3, background: c }} />
                        {l} ({r})
                      </div>
                    ))}
                  </div>
                </div>

                {numericCols.length >= 2 && (
                  <ChartSection title={`Scatter: ${numericCols[0]} vs ${numericCols[1]}`} accent={T.amber}>
                    <ResponsiveContainer width="100%" height={280}>
                      <ScatterChart>
                        <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                        <XAxis dataKey={numericCols[0]} type="number" name={numericCols[0]} tick={{ fontSize: 10, fill: T.muted }} label={{ value: numericCols[0], position: "insideBottom", offset: -4, fill: T.muted, fontSize: 11 }} />
                        <YAxis dataKey={numericCols[1]} type="number" name={numericCols[1]} tick={{ fontSize: 10, fill: T.muted }} />
                        <Tooltip content={<DarkTooltip />} cursor={{ fill: `${T.accent}10` }} />
                        <Scatter
                          data={data.map(r => ({ [numericCols[0]]: num(r[numericCols[0]]), [numericCols[1]]: num(r[numericCols[1]]) })).filter(r => !isNaN(r[numericCols[0]]) && !isNaN(r[numericCols[1]]))}
                          fill={T.amber}
                          fillOpacity={0.7}
                        />
                      </ScatterChart>
                    </ResponsiveContainer>
                    <div style={{ marginTop: 12, fontFamily: "'DM Mono', monospace", fontSize: 12, color: T.muted }}>
                      Correlation coefficient: <span style={{ color: T.amber, fontWeight: 700 }}>{correlation(data, numericCols[0], numericCols[1]).toFixed(3)}</span>
                    </div>
                  </ChartSection>
                )}
              </>
            ) : (
              <div style={{ textAlign: "center", color: T.muted, padding: 80, fontFamily: "'DM Mono', monospace", fontSize: 13 }}>
                Need at least 2 numeric columns for correlation analysis.
              </div>
            )}
          </>
        )}

        {/* ── INSIGHTS TAB ── */}
        {activeTab === "insights" && (
          <>
            <div className="card fadeUp" style={{ marginBottom: 24 }}>
              <div className="section-title">Auto-Generated Insights</div>
              <div style={{ display: "grid", gap: 12 }}>
                {insights.map((ins, i) => (
                  <div key={i} className="insight-chip fadeUp" style={{ animationDelay: `${i * 60}ms` }}>{ins}</div>
                ))}
              </div>
            </div>

            {numericCols.map((col, ci) => {
              const outliers = detectOutliers(data, col);
              if (!outliers.length) return null;
              return (
                <div key={col} className="card fadeUp" style={{ marginBottom: 20, borderColor: `${T.accent3}40` }}>
                  <div className="section-title" style={{ color: T.accent3 }}>⚠ Outliers in {col} ({outliers.length} found)</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'DM Mono', monospace", fontSize: 12 }}>
                      <thead>
                        <tr>
                          {columns.slice(0, 6).map(c => (
                            <th key={c} style={{ padding: "8px 12px", textAlign: "left", color: T.muted, borderBottom: `1px solid ${T.border}` }}>{c}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {outliers.slice(0, 10).map((row, i) => (
                          <tr key={i} style={{ borderBottom: `1px solid ${T.border}` }}>
                            {columns.slice(0, 6).map(c => (
                              <td key={c} style={{ padding: "8px 12px", color: c === col ? T.accent3 : T.text }}>{row[c]}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}

            <div className="card fadeUp">
              <div className="section-title">Column Summary Table</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'DM Mono', monospace", fontSize: 12 }}>
                  <thead>
                    <tr>
                      {["Column", "Type", "Count", "Unique", "Missing", "Min", "Max", "Mean"].map(h => (
                        <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: T.muted, borderBottom: `1px solid ${T.border}` }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {columns.map((col, i) => {
                      const isN = numericCols.includes(col);
                      const s = isN ? calcStats(data, col) : null;
                      const uniq = new Set(data.map(r => r[col])).size;
                      const miss = data.filter(r => !r[col] || r[col] === "").length;
                      return (
                        <tr key={col} style={{ borderBottom: `1px solid ${T.border}` }}>
                          <td style={{ padding: "8px 12px", color: T.text, fontWeight: 500 }}>{col}</td>
                          <td style={{ padding: "8px 12px" }}>
                            <span style={{ color: isN ? T.accent : T.amber, background: `${isN ? T.accent : T.amber}15`, borderRadius: 4, padding: "2px 8px", fontSize: 11 }}>
                              {isN ? "numeric" : "text"}
                            </span>
                          </td>
                          <td style={{ padding: "8px 12px", color: T.muted }}>{data.length}</td>
                          <td style={{ padding: "8px 12px", color: T.muted }}>{uniq}</td>
                          <td style={{ padding: "8px 12px", color: miss > 0 ? T.accent3 : T.muted }}>{miss}</td>
                          <td style={{ padding: "8px 12px", color: T.muted }}>{s ? s.min.toFixed(2) : "—"}</td>
                          <td style={{ padding: "8px 12px", color: T.muted }}>{s ? s.max.toFixed(2) : "—"}</td>
                          <td style={{ padding: "8px 12px", color: T.muted }}>{s ? s.mean.toFixed(2) : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ─── ROOT ─── */
export default function App() {
  const [unlocked, setUnlocked] = useState(false);
  return (
    <>
      <style>{GLOBAL_CSS}</style>
      {unlocked ? <Dashboard /> : <PasswordScreen onUnlock={() => setUnlocked(true)} />}
    </>
  );
}