"use client";

import React, { useState, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, Search, Download, Stamp, AlertCircle, CheckCircle2,
  ArrowUpRight, MinusCircle, Users, Trophy, LineChart as LineChartIcon,
  Trash2, ClipboardPaste, RefreshCw
} from "lucide-react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell
} from "recharts";
import { storage } from "@/lib/storage";

const INK = "#1B2A4A";
const PAPER = "#FAF9F5";
const PAPER_LINE = "#E4E0D6";
const CHURN_RED = "#B23A2E";
const RETAIN_GREEN = "#3F6C51";
const NEW_BLUE = "#2A5C7A";
const DORMANT_GRAY = "#8A8578";
const GOLD = "#B8863B";

const PRODUCTS = [
  { key: "duchat", label: "Duchat" },
  { key: "premical", label: "Premical" },
  { key: "smartkid", label: "Smartkid" },
  { key: "tigin", label: "Tigin" },
  { key: "baciforte", label: "Baciforte" },
  { key: "nebusal", label: "Nebusal" },
];

const CREATOR_KEYS = ["tên người dùng của nhà sáng tạo", "creator", "username", "nhà sáng tạo"];
const DATE_KEYS = ["ngày đăng video", "ngày đăng", "date"];
const GMV_KEYS = ["gmv"];
const VIDEO_KEYS = ["tên video", "video"];
const CHART_COLORS = [NEW_BLUE, RETAIN_GREEN, GOLD, CHURN_RED, "#6B5B95", DORMANT_GRAY];

function findKey(row: any, candidates: string[]) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const hit = keys.find((k) => k.trim().toLowerCase() === cand);
    if (hit) return hit;
  }
  for (const cand of candidates) {
    const hit = keys.find((k) => k.trim().toLowerCase().includes(cand));
    if (hit) return hit;
  }
  return null;
}

function toDate(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const parsed = new Date(val);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function monthKeyOf(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(key: string) {
  const [y, m] = key.split("-");
  return `Tháng ${parseInt(m, 10)}/${y}`;
}

function detectProduct(videoName: any) {
  if (!videoName) return null;
  const lower = videoName.toString().toLowerCase();
  const hit = PRODUCTS.find((p) => lower.includes(p.key));
  return hit ? hit.key : null;
}

function parseVideoListFile(arrayBuffer: ArrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: null });
  if (!rows.length) return { ok: false as const, reason: "File rỗng." };
  const creatorKey = findKey(rows[0], CREATOR_KEYS);
  const dateKey = findKey(rows[0], DATE_KEYS);
  const gmvKey = findKey(rows[0], GMV_KEYS);
  const videoKey = findKey(rows[0], VIDEO_KEYS);
  if (!creatorKey || !dateKey || !gmvKey) {
    return { ok: false as const, reason: "Thiếu cột Creator / Ngày đăng video / GMV." };
  }
  const byMonth: Record<string, Map<string, any>> = {};
  rows.forEach((r) => {
    const d = toDate(r[dateKey]);
    if (!d) return;
    const mk = monthKeyOf(d);
    const name = (r[creatorKey] || "").toString().trim();
    if (!name) return;
    const gmv = Number(r[gmvKey]) || 0;
    const product = videoKey ? detectProduct(r[videoKey]) : null;
    if (!byMonth[mk]) byMonth[mk] = new Map();
    if (!byMonth[mk].has(name)) byMonth[mk].set(name, { name, videoCount: 0, gmv: 0, products: {} });
    const entry = byMonth[mk].get(name);
    entry.videoCount += 1;
    entry.gmv += gmv;
    if (product) entry.products[product] = (entry.products[product] || 0) + 1;
  });
  const months: Record<string, any[]> = {};
  Object.keys(byMonth).forEach((mk) => {
    months[mk] = Array.from(byMonth[mk].values()).map((c: any) => ({ ...c, products: Object.keys(c.products) }));
  });
  return { ok: true as const, months };
}

function parseAssignmentPaste(text: string) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const map: Record<string, string> = {};
  let count = 0;
  lines.forEach((line) => {
    const cols = line.split(/\t|,|;/).map((c) => c.trim());
    if (cols.length < 2) return;
    const name = cols[0];
    const person = cols[1];
    if (!name || !person) return;
    const lname = name.toLowerCase();
    if (lname.includes("creator") || lname.includes("nhà sáng tạo") || lname.includes("username") || lname.includes("tên ")) return;
    map[name] = person;
    count++;
  });
  return { map, count };
}

async function loadAssignments() {
  try {
    const a = await storage.get("koc_assignments");
    if (a && a.value) return JSON.parse(a.value);
  } catch (e) {}
  return {};
}

async function loadPeriods() {
  const periods: Record<string, any[]> = {};
  try {
    const list = await storage.list("koc_period_");
    const keys = list && list.keys ? list.keys : [];
    for (const k of keys) {
      try {
        const res = await storage.get(k);
        if (res && res.value) periods[k.replace("koc_period_", "")] = JSON.parse(res.value);
      } catch (e) {}
    }
  } catch (e) {}
  return periods;
}

function TabButton({ active, onClick, icon: Icon, label }: any) {
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-2 px-4 py-2.5 rounded-sm text-sm"
      style={{
        fontFamily: "Inter", fontWeight: 500,
        background: active ? INK : "#fff", color: active ? "#fff" : INK,
        border: `1px solid ${active ? INK : PAPER_LINE}`,
      }}
    >
      <Icon size={15} />
      {label}
    </button>
  );
}

function StatCard({ label, value, color, sub }: any) {
  return (
    <div className="flex-1 rounded-sm px-4 py-3" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}`, borderTop: `3px solid ${color}` }}>
      <div style={{ fontFamily: "Inter", fontSize: 11, color: INK, opacity: 0.55, textTransform: "uppercase", letterSpacing: "0.08em" }}>{label}</div>
      <div style={{ fontFamily: "JetBrains Mono", fontSize: 24, fontWeight: 600, color: INK, marginTop: 4 }}>{value}</div>
      {sub && <div style={{ fontFamily: "Inter", fontSize: 11, color: INK, opacity: 0.45, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function EmptyState({ text }: { text: string }) {
  return (
    <div className="rounded-sm px-6 py-10 text-center" style={{ border: `1px dashed ${PAPER_LINE}` }}>
      <span style={{ fontFamily: "Inter", fontSize: 13, color: INK, opacity: 0.45 }}>{text}</span>
    </div>
  );
}

function ChartCard({ title, children }: any) {
  return (
    <div className="rounded-sm p-5" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}` }}>
      <div style={{ fontFamily: "Inter", fontSize: 12, fontWeight: 600, color: INK, marginBottom: 12 }}>{title}</div>
      {children}
    </div>
  );
}

const thStyle: React.CSSProperties = {
  textAlign: "left", padding: "10px 14px", fontFamily: "Inter", fontSize: 11,
  fontWeight: 600, color: INK, opacity: 0.6, textTransform: "uppercase", letterSpacing: "0.04em",
};
const tdStyle: React.CSSProperties = { padding: "9px 14px", fontSize: 13, color: INK };
const tdMono: React.CSSProperties = { ...tdStyle, fontFamily: "JetBrains Mono", fontSize: 12.5 };

const CLASS_META: Record<string, any> = {
  churn: { label: "Rời bỏ", color: CHURN_RED, icon: Stamp },
  retained: { label: "Duy trì", color: RETAIN_GREEN, icon: CheckCircle2 },
  new: { label: "Mới / Quay lại", color: NEW_BLUE, icon: ArrowUpRight },
  dormant: { label: "Không hoạt động", color: DORMANT_GRAY, icon: MinusCircle },
};

function ChurnUploadSlot({ label, meta, onFile, selectedMonth, onSelectMonth, error }: any) {
  const inputId = `churn-file-${label}`;
  return (
    <div className="flex-1 rounded-sm p-5" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}` }}>
      <div className="flex items-center justify-between mb-3">
        <span className="text-xs tracking-widest uppercase" style={{ fontFamily: "Inter", color: INK, opacity: 0.55, letterSpacing: "0.12em" }}>{label}</span>
        {meta && <FileSpreadsheet size={16} color={RETAIN_GREEN} />}
      </div>
      <label htmlFor={inputId} className="flex flex-col items-center justify-center gap-2 rounded-sm cursor-pointer" style={{ border: `1.5px dashed ${meta ? RETAIN_GREEN : PAPER_LINE}`, padding: "26px 12px", background: meta ? "#F4F8F5" : PAPER }}>
        <Upload size={20} color={meta ? RETAIN_GREEN : INK} style={{ opacity: meta ? 1 : 0.4 }} />
        <span style={{ fontFamily: "Inter", fontSize: 13, color: INK, opacity: 0.7, textAlign: "center" }}>{meta ? meta.fileName : "Kéo file .xlsx video list vào đây"}</span>
        <input id={inputId} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => onFile(e.target.files?.[0])} />
      </label>
      {error && (
        <div className="mt-2 flex items-center gap-1.5" style={{ color: CHURN_RED }}>
          <AlertCircle size={13} /><span style={{ fontFamily: "Inter", fontSize: 12 }}>{error}</span>
        </div>
      )}
      {meta && meta.monthKeys.length > 0 && (
        <div className="mt-3">
          <select value={selectedMonth || ""} onChange={(e) => onSelectMonth(e.target.value)} className="w-full mt-1 rounded-sm px-2 py-2" style={{ fontFamily: "JetBrains Mono", fontSize: 13, border: `1px solid ${PAPER_LINE}`, background: "#fff", color: INK }}>
            {meta.monthKeys.map((m: string) => <option key={m} value={m}>{monthLabel(m)} — {meta.months[m].length} creator</option>)}
          </select>
        </div>
      )}
    </div>
  );
}

export default function KocSuite() {
  const [tab, setTab] = useState("master");
  const [loading, setLoading] = useState(true);
  const [assignments, setAssignments] = useState<Record<string, string>>({});
  const [periods, setPeriods] = useState<Record<string, any[]>>({});

  useEffect(() => {
    (async () => {
      const [a, p] = await Promise.all([loadAssignments(), loadPeriods()]);
      setAssignments(a);
      setPeriods(p);
      setLoading(false);
    })();
  }, []);

  const monthList = useMemo(() => Object.keys(periods).sort(), [periods]);

  const master = useMemo(() => {
    const map = new Map<string, any>();
    monthList.forEach((mk) => {
      (periods[mk] || []).forEach((c: any) => {
        if (!map.has(c.name)) map.set(c.name, { name: c.name, totalVideo: 0, totalGmv: 0, products: new Set(), months: new Set() });
        const e = map.get(c.name);
        e.totalVideo += c.videoCount;
        e.totalGmv += c.gmv;
        (c.products || []).forEach((p: string) => e.products.add(p));
        e.months.add(mk);
      });
    });
    return Array.from(map.values()).map((e: any) => ({
      ...e, products: Array.from(e.products), monthsActive: e.months.size,
      person: assignments[e.name] || "—",
    }));
  }, [periods, monthList, assignments]);

  const ranked = useMemo(() => {
    if (master.length === 0) return [];
    const withAvg = master.map((c: any) => ({ ...c, avgGmv: c.totalVideo > 0 ? c.totalGmv / c.totalVideo : 0 }));
    const sorted = [...withAvg].sort((a, b) => a.avgGmv - b.avgGmv);
    const n = sorted.length;
    sorted.forEach((c, i) => { c.avgPercentile = n > 1 ? i / (n - 1) : 1; });
    const totalPeriods = Math.max(monthList.length, 1);
    sorted.forEach((c) => {
      c.consistency = c.monthsActive / totalPeriods;
      c.diversity = c.products.length / PRODUCTS.length;
      c.score = c.avgPercentile * 0.5 + c.consistency * 0.3 + c.diversity * 0.2;
    });
    sorted.sort((a, b) => b.score - a.score);
    const cut1 = Math.ceil(sorted.length * 0.2);
    const cut2 = Math.ceil(sorted.length * 0.6);
    return sorted.map((c, i) => ({ ...c, tier: i < cut1 ? "Top" : i < cut2 ? "Mid" : "Low" }));
  }, [master, monthList]);

const [pasteText, setPasteText] = useState("");
  const [pasteMsg, setPasteMsg] = useState<any>(null);
  const [sheetSyncing, setSheetSyncing] = useState(false);

  const syncAssignments = useCallback(async () => {
    const { map, count } = parseAssignmentPaste(pasteText);
    if (count === 0) {
      setPasteMsg({ type: "err", text: "Không nhận diện được dòng nào. Đảm bảo đã dán 2 cột: Creator và Người phụ trách." });
      return;
    }
    const merged = { ...assignments, ...map };
    setAssignments(merged);
    await storage.set("koc_assignments", JSON.stringify(merged));
    setPasteMsg({ type: "ok", text: `Đã đồng bộ ${count} creator từ dữ liệu dán.` });
  }, [pasteText, assignments]);

  const syncFromGoogleSheets = useCallback(async () => {
    setSheetSyncing(true);
    setPasteMsg(null);
    try {
      const res = await fetch("/api/assignments");
      const data = await res.json();
      if (!data.ok) {
        setPasteMsg({ type: "err", text: `Lỗi kết nối Google Sheets: ${data.error}` });
      } else {
        const merged = { ...assignments, ...data.map };
        setAssignments(merged);
        await storage.set("koc_assignments", JSON.stringify(merged));
        setPasteMsg({ type: "ok", text: `Đã đồng bộ ${data.count} creator trực tiếp từ Google Sheets.` });
      }
    } catch (e) {
      setPasteMsg({ type: "err", text: "Không kết nối được API." });
    }
    setSheetSyncing(false);
  }, [assignments]);

  const exportMasterCsv = () => {
    const header = ["Creator", "Nguoi phu trach", "Tong video", "Tong GMV", "So san pham", "So tuan hoat dong"];
    const lines = [header.join(",")];
    master.forEach((c: any) => lines.push([c.name, c.person, c.totalVideo, c.totalGmv, c.products.length, c.monthsActive].join(",")));
    const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "tong_hop_koc.csv"; a.click();
    URL.revokeObjectURL(url);
  };

  const [rankBusy, setRankBusy] = useState(false);
  const [rankMsg, setRankMsg] = useState<any>(null);
  const handleWeeklyUpload = useCallback(async (file: File) => {
    if (!file) return;
    setRankBusy(true);
    setRankMsg(null);
    try {
      const buf = await file.arrayBuffer();
      const res = parseVideoListFile(buf);
      if (!res.ok) {
        setRankMsg({ type: "err", text: res.reason });
      } else {
        const monthsFound = Object.keys(res.months);
        const updated = { ...periods };
        for (const mk of monthsFound) {
          updated[mk] = res.months[mk];
          await storage.set(`koc_period_${mk}`, JSON.stringify(res.months[mk]));
        }
        setPeriods(updated);
        setRankMsg({ type: "ok", text: `Đã lưu doanh thu tuần: ${monthsFound.map(monthLabel).join(", ")}.` });
      }
    } catch (e) {
      setRankMsg({ type: "err", text: "Không đọc được file." });
    }
    setRankBusy(false);
  }, [periods]);

  const deletePeriod = useCallback(async (mk: string) => {
    try {
      await storage.delete(`koc_period_${mk}`);
      const updated = { ...periods };
      delete updated[mk];
      setPeriods(updated);
    } catch (e) {}
  }, [periods]);

  const [fileAMeta, setFileAMeta] = useState<any>(null);
  const [fileBMeta, setFileBMeta] = useState<any>(null);
  const [monthA, setMonthA] = useState<string | null>(null);
  const [monthB, setMonthB] = useState<string | null>(null);
  const [errorA, setErrorA] = useState<string | null>(null);
  const [errorB, setErrorB] = useState<string | null>(null);
  const [threshold, setThreshold] = useState(0);
  const [churnTab, setChurnTab] = useState("churn");
  const [churnSearch, setChurnSearch] = useState("");

  const handleChurnFile = useCallback(async (file: File, which: "A" | "B") => {
    if (!file) return;
    const setMeta = which === "A" ? setFileAMeta : setFileBMeta;
    const setMonth = which === "A" ? setMonthA : setMonthB;
    const setErr = which === "A" ? setErrorA : setErrorB;
    setErr(null);
    try {
      const buf = await file.arrayBuffer();
      const res = parseVideoListFile(buf);
      if (!res.ok) {
        setErr(res.reason);
        setMeta(null);
        return;
      }
      const monthKeys = Object.keys(res.months).sort().reverse();
      setMeta({ fileName: file.name, months: res.months, monthKeys });
      setMonth(monthKeys[0] || null);
    } catch (e) {
      setErr("Không đọc được file.");
      setMeta(null);
    }
  }, []);

  const churnResults = useMemo(() => {
    if (!fileAMeta || !fileBMeta || !monthA || !monthB) return null;
    const prevMap = new Map((fileAMeta.months[monthA] || []).map((c: any) => [c.name, c]));
    const currMap = new Map((fileBMeta.months[monthB] || []).map((c: any) => [c.name, c]));
    const allNames = new Set([...prevMap.keys(), ...currMap.keys()]);
    const rows: any[] = [];
    allNames.forEach((name: any) => {
      const prev: any = prevMap.get(name);
      const curr: any = currMap.get(name);
      const hasPrevRevenue = prev && prev.gmv > threshold;
      const hasCurrVideo = curr && curr.videoCount > 0;
      let cls;
      if (hasPrevRevenue && !hasCurrVideo) cls = "churn";
      else if (hasPrevRevenue && hasCurrVideo) cls = "retained";
      else if (!hasPrevRevenue && hasCurrVideo) cls = "new";
      else cls = "dormant";
      rows.push({
        name, person: assignments[name] || "—",
        prevVideo: prev ? prev.videoCount : 0, prevGmv: prev ? prev.gmv : 0,
        currVideo: curr ? curr.videoCount : 0, currGmv: curr ? curr.gmv : 0, cls,
      });
    });
    const counts: any = { churn: 0, retained: 0, new: 0, dormant: 0 };
    rows.forEach((r) => counts[r.cls]++);
    const churnGmvAtRisk = rows.filter((r) => r.cls === "churn").reduce((s, r) => s + r.prevGmv, 0);
    return { rows, counts, churnGmvAtRisk };
  }, [fileAMeta, fileBMeta, monthA, monthB, threshold, assignments]);

  const filteredChurnRows = useMemo(() => {
    if (!churnResults) return [];
    let rows = churnResults.rows;
    if (churnTab !== "all") rows = rows.filter((r) => r.cls === churnTab);
    if (churnSearch.trim()) {
      const q = churnSearch.trim().toLowerCase();
      rows = rows.filter((r) => r.name.toLowerCase().includes(q));
    }
    return [...rows].sort((a, b) => b.prevGmv - a.prevGmv);
  }, [churnResults, churnTab, churnSearch]);

  const monthlyTrend = useMemo(() => monthList.map((mk) => ({
    month: monthLabel(mk).replace("Tháng ", "T"),
    gmv: Math.round((periods[mk] || []).reduce((s: number, c: any) => s + c.gmv, 0)),
  })), [periods, monthList]);

  const productBreakdown = useMemo(() => {
    const sums: Record<string, number> = {};
    PRODUCTS.forEach((p) => { sums[p.key] = 0; });
    monthList.forEach((mk) => {
      (periods[mk] || []).forEach((c: any) => {
        const nProd = (c.products || []).length || 1;
        const share = c.gmv / nProd;
        (c.products || []).forEach((p: string) => { sums[p] += share; });
      });
    });
    return PRODUCTS.map((p) => ({ name: p.label, gmv: Math.round(sums[p.key]) })).filter((d) => d.gmv > 0);
  }, [periods, monthList]);

  const memberKpi = useMemo(() => {
    const map: Record<string, any> = {};
    master.forEach((c: any) => {
      const person = c.person === "—" ? "Chưa gán" : c.person;
      if (!map[person]) map[person] = { person, creators: 0, totalGmv: 0, totalVideo: 0, top: 0, mid: 0, low: 0 };
      map[person].creators += 1;
      map[person].totalGmv += c.totalGmv;
      map[person].totalVideo += c.totalVideo;
    });
    ranked.forEach((r: any) => {
      const person = r.person === "—" ? "Chưa gán" : r.person;
      if (!map[person]) map[person] = { person, creators: 0, totalGmv: 0, totalVideo: 0, top: 0, mid: 0, low: 0 };
      if (r.tier === "Top") map[person].top++;
      if (r.tier === "Mid") map[person].mid++;
      if (r.tier === "Low") map[person].low++;
    });
    return Object.values(map).sort((a: any, b: any) => b.totalGmv - a.totalGmv);
  }, [master, ranked]);

  const fmt = (n: number) => Math.round(n).toLocaleString("vi-VN");

  if (loading) {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <span style={{ fontFamily: "Inter", color: INK, opacity: 0.6 }}>Đang tải dữ liệu đã lưu...</span>
      </div>
    );
  }

  return (
    <div style={{ background: PAPER, minHeight: "100vh", fontFamily: "Inter" }}>
      <div className="max-w-6xl mx-auto px-6 py-8">
        <div className="flex items-start justify-between mb-6 pb-5" style={{ borderBottom: `2px solid ${INK}` }}>
          <div>
            <div style={{ fontFamily: "Inter", fontSize: 11, letterSpacing: "0.15em", color: INK, opacity: 0.5, textTransform: "uppercase" }}>
              Hệ thống quản lý Creator · TikTok Shop
            </div>
            <h1 style={{ fontFamily: "Roboto Slab", fontWeight: 700, fontSize: 28, color: INK, marginTop: 4 }}>KOC Management Suite</h1>
          </div>
          <div className="flex items-center gap-1.5 rounded-full px-3 py-1.5" style={{ border: `1px solid ${PAPER_LINE}`, transform: "rotate(-2deg)" }}>
            <Stamp size={13} color={CHURN_RED} />
            <span style={{ fontFamily: "JetBrains Mono", fontSize: 11, color: CHURN_RED }}>SO4 · CPC1HN</span>
          </div>
        </div>

        <div className="flex gap-2 mb-6 flex-wrap">
          <TabButton active={tab === "master"} onClick={() => setTab("master")} icon={Users} label="Tổng hợp KOC" />
          <TabButton active={tab === "ranking"} onClick={() => setTab("ranking")} icon={Trophy} label="Xếp hạng KOC" />
          <TabButton active={tab === "churn"} onClick={() => setTab("churn")} icon={Stamp} label="KOC Rời bỏ" />
          <TabButton active={tab === "analytics"} onClick={() => setTab("analytics")} icon={LineChartIcon} label="Phân tích KOC" />
        </div>

        {tab === "master" && (
          <div className="space-y-5">
<div className="rounded-sm p-5" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}` }}>
              <div className="flex items-center gap-2 mb-2">
                <ClipboardPaste size={15} color={INK} />
                <span style={{ fontFamily: "Inter", fontSize: 13, fontWeight: 600, color: INK }}>Người phụ trách — Google Sheets</span>
              </div>

              <button
                onClick={syncFromGoogleSheets}
                disabled={sheetSyncing}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-sm text-xs mb-3"
                style={{ fontFamily: "Inter", fontWeight: 500, background: INK, color: "#fff", opacity: sheetSyncing ? 0.6 : 1 }}
              >
                <RefreshCw size={13} className={sheetSyncing ? "animate-spin" : ""} />
                {sheetSyncing ? "Đang đồng bộ..." : "Đồng bộ trực tiếp từ Google Sheets"}
              </button>

              <details className="mt-1">
                <summary style={{ fontFamily: "Inter", fontSize: 11, color: INK, opacity: 0.5, cursor: "pointer" }}>Hoặc dán thủ công (nếu API lỗi)</summary>
                <div className="mt-2">
                  <textarea
                    value={pasteText}
                    onChange={(e) => setPasteText(e.target.value)}
                    placeholder={"creator_username_1\tThùy Trang\ncreator_username_2\tThái Bình"}
                    rows={4}
                    className="w-full rounded-sm px-3 py-2"
                    style={{ fontFamily: "JetBrains Mono", fontSize: 12, border: `1px solid ${PAPER_LINE}`, background: PAPER, color: INK }}
                  />
                  <button onClick={syncAssignments} className="flex items-center gap-1.5 px-3 py-1.5 rounded-sm text-xs mt-2" style={{ fontFamily: "Inter", fontWeight: 500, background: "#fff", border: `1px solid ${PAPER_LINE}`, color: INK }}>
                    Đồng bộ dữ liệu dán
                  </button>
                </div>
              </details>

              <div className="flex items-center justify-between mt-3">
                <span style={{ fontFamily: "JetBrains Mono", fontSize: 11, color: INK, opacity: 0.5 }}>{Object.keys(assignments).length} creator đã gán</span>
              </div>
              {pasteMsg && (
                <div className="flex items-center gap-1.5 mt-2" style={{ color: pasteMsg.type === "err" ? CHURN_RED : RETAIN_GREEN }}>
                  {pasteMsg.type === "err" ? <AlertCircle size={13} /> : <CheckCircle2 size={13} />}
                  <span style={{ fontFamily: "Inter", fontSize: 12 }}>{pasteMsg.text}</span>
                </div>
              )}
            </div>
            <div className="flex gap-3">
              <StatCard label="Tổng creator" value={master.length} color={NEW_BLUE} />
              <StatCard label="Tổng GMV lũy kế" value={`${fmt(master.reduce((s: number, c: any) => s + c.totalGmv, 0))} đ`} color={GOLD} />
              <StatCard label="Số tuần dữ liệu" value={monthList.length} color={RETAIN_GREEN} />
            </div>

            {master.length === 0 ? (
              <EmptyState text="Chưa có dữ liệu doanh thu — vào tab Xếp hạng KOC để upload file doanh thu hàng tuần trước." />
            ) : (
              <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${PAPER_LINE}`, background: "#fff" }}>
                <div className="flex justify-end p-2" style={{ borderBottom: `1px solid ${PAPER_LINE}` }}>
                  <button onClick={exportMasterCsv} className="flex items-center gap-1.5 px-3 py-1.5 rounded-sm text-xs" style={{ fontFamily: "Inter", border: `1px solid ${PAPER_LINE}`, color: INK }}>
                    <Download size={13} /> Tải CSV
                  </button>
                </div>
                <table className="w-full" style={{ borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: PAPER }}>
                      <th style={thStyle}>Creator</th>
                      <th style={thStyle}>Người phụ trách</th>
                      <th style={thStyle}>Tổng video</th>
                      <th style={thStyle}>Tổng GMV</th>
                      <th style={thStyle}>Sản phẩm</th>
                      <th style={thStyle}>Số tuần hoạt động</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...master].sort((a: any, b: any) => b.totalGmv - a.totalGmv).slice(0, 300).map((c: any, i: number) => (
                      <tr key={c.name} style={{ borderTop: `1px solid ${PAPER_LINE}`, background: i % 2 ? "#FCFBF8" : "#fff" }}>
                        <td style={{ ...tdStyle, fontWeight: 500 }}>{c.name}</td>
                        <td style={tdStyle}>{c.person}</td>
                        <td style={tdMono}>{c.totalVideo}</td>
                        <td style={tdMono}>{fmt(c.totalGmv)}</td>
                        <td style={tdStyle}>{c.products.length ? c.products.join(", ") : "—"}</td>
                        <td style={tdMono}>{c.monthsActive}/{monthList.length}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {tab === "ranking" && (
          <div className="space-y-5">
            <div className="rounded-sm p-5" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}` }}>
              <div style={{ fontFamily: "Inter", fontSize: 13, fontWeight: 600, color: INK, marginBottom: 8 }}>Upload doanh thu KOC hàng tuần</div>
              <label className="flex items-center justify-center gap-2 rounded-sm cursor-pointer" style={{ border: `1.5px dashed ${PAPER_LINE}`, padding: "18px", background: PAPER }}>
                <Upload size={17} color={INK} style={{ opacity: 0.5 }} />
                <span style={{ fontFamily: "Inter", fontSize: 13, color: INK, opacity: 0.7 }}>Chọn file doanh thu tuần này (.xlsx/.csv)</span>
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => e.target.files && handleWeeklyUpload(e.target.files[0])} />
              </label>
              {rankBusy && <div style={{ fontFamily: "Inter", fontSize: 12, color: INK, opacity: 0.5, marginTop: 8 }}>Đang xử lý...</div>}
              {rankMsg && (
                <div className="flex items-center gap-1.5 mt-2" style={{ color: rankMsg.type === "err" ? CHURN_RED : RETAIN_GREEN }}>
                  {rankMsg.type === "err" ? <AlertCircle size={13} /> : <CheckCircle2 size={13} />}
                  <span style={{ fontFamily: "Inter", fontSize: 12 }}>{rankMsg.text}</span>
                </div>
              )}
              {monthList.length > 0 && (
                <div className="mt-4">
                  <div style={{ fontFamily: "Inter", fontSize: 11, color: INK, opacity: 0.5, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>
                    Các kỳ đã lưu ({monthList.length})
                  </div>
                  <div className="space-y-1.5">
                    {monthList.map((mk) => (
                      <div key={mk} className="flex items-center justify-between rounded-sm px-3 py-2" style={{ border: `1px solid ${PAPER_LINE}`, background: PAPER }}>
                        <span style={{ fontFamily: "JetBrains Mono", fontSize: 12.5, color: INK }}>{monthLabel(mk)} — {(periods[mk] || []).length} creator</span>
                        <button onClick={() => deletePeriod(mk)} style={{ color: CHURN_RED, opacity: 0.6 }}><Trash2 size={14} /></button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {ranked.length === 0 ? (
              <EmptyState text="Chưa có dữ liệu — upload file doanh thu tuần này ở trên." />
            ) : (
              <>
                <div className="flex gap-3">
                  <StatCard label="Top" value={ranked.filter((r: any) => r.tier === "Top").length} color={GOLD} sub="Ưu tiên booking lại" />
                  <StatCard label="Mid" value={ranked.filter((r: any) => r.tier === "Mid").length} color={NEW_BLUE} sub="Theo dõi thêm" />
                  <StatCard label="Low" value={ranked.filter((r: any) => r.tier === "Low").length} color={DORMANT_GRAY} sub="Cân nhắc dừng đầu tư" />
                </div>
                <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${PAPER_LINE}`, background: "#fff" }}>
                  <table className="w-full" style={{ borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: PAPER }}>
                        <th style={thStyle}>#</th>
                        <th style={thStyle}>Creator</th>
                        <th style={thStyle}>Nhóm</th>
                        <th style={thStyle}>GMV TB/video</th>
                        <th style={thStyle}>Độ ổn định</th>
                        <th style={thStyle}>Đa dạng SP</th>
                        <th style={thStyle}>Điểm</th>
                      </tr>
                    </thead>
                    <tbody>
                      {ranked.slice(0, 300).map((c: any, i: number) => (
                        <tr key={c.name} style={{ borderTop: `1px solid ${PAPER_LINE}`, background: i % 2 ? "#FCFBF8" : "#fff" }}>
                          <td style={tdMono}>{i + 1}</td>
                          <td style={{ ...tdStyle, fontWeight: 500 }}>{c.name}</td>
                          <td style={tdStyle}>
                            <span style={{ fontFamily: "JetBrains Mono", fontSize: 11, padding: "2px 8px", borderRadius: 999, border: `1px solid ${c.tier === "Top" ? GOLD : c.tier === "Mid" ? NEW_BLUE : DORMANT_GRAY}`, color: c.tier === "Top" ? GOLD : c.tier === "Mid" ? NEW_BLUE : DORMANT_GRAY }}>{c.tier}</span>
                          </td>
                          <td style={tdMono}>{fmt(c.avgGmv)} đ</td>
                          <td style={tdMono}>{Math.round(c.consistency * 100)}%</td>
                          <td style={tdMono}>{c.products.length}/6</td>
                          <td style={tdMono}>{c.score.toFixed(2)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}

        {tab === "churn" && (
          <div className="space-y-5">
            <div className="flex gap-4">
              <ChurnUploadSlot label="Kỳ trước" meta={fileAMeta} onFile={(f: File) => handleChurnFile(f, "A")} selectedMonth={monthA} onSelectMonth={setMonthA} error={errorA} />
              <ChurnUploadSlot label="Kỳ này" meta={fileBMeta} onFile={(f: File) => handleChurnFile(f, "B")} selectedMonth={monthB} onSelectMonth={setMonthB} error={errorB} />
            </div>

            <div className="flex items-center gap-2 px-1">
              <span style={{ fontFamily: "Inter", fontSize: 12, color: INK, opacity: 0.6 }}>Ngưỡng GMV kỳ trước để tính là "có doanh số":</span>
              <input type="number" min={0} value={threshold} onChange={(e) => setThreshold(Number(e.target.value) || 0)} className="rounded-sm px-2 py-1 w-32" style={{ fontFamily: "JetBrains Mono", fontSize: 12, border: `1px solid ${PAPER_LINE}`, background: "#fff", color: INK }} />
            </div>

            {!churnResults && <EmptyState text="Tải lên đủ 2 file (kỳ trước & kỳ này) để bắt đầu phân tích." />}

            {churnResults && (
              <>
                <div className="flex gap-3">
                  <StatCard label="Rời bỏ" value={churnResults.counts.churn} color={CHURN_RED} sub={`${fmt(churnResults.churnGmvAtRisk)} đ rủi ro`} />
                  <StatCard label="Duy trì" value={churnResults.counts.retained} color={RETAIN_GREEN} />
                  <StatCard label="Mới/Quay lại" value={churnResults.counts.new} color={NEW_BLUE} />
                  <StatCard label="Không hoạt động" value={churnResults.counts.dormant} color={DORMANT_GRAY} />
                </div>

                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div className="flex gap-1">
                    {["churn", "retained", "new", "dormant", "all"].map((t) => (
                      <button key={t} onClick={() => setChurnTab(t)} className="px-3 py-1.5 rounded-sm text-xs" style={{ fontFamily: "Inter", fontWeight: 500, background: churnTab === t ? INK : "#fff", color: churnTab === t ? "#fff" : INK, border: `1px solid ${churnTab === t ? INK : PAPER_LINE}` }}>
                        {t === "all" ? "Tất cả" : CLASS_META[t].label}
                      </button>
                    ))}
                  </div>
                  <div className="flex items-center gap-1.5 rounded-sm px-2" style={{ border: `1px solid ${PAPER_LINE}`, background: "#fff" }}>
                    <Search size={13} color={INK} style={{ opacity: 0.4 }} />
                    <input value={churnSearch} onChange={(e) => setChurnSearch(e.target.value)} placeholder="Tìm creator..." className="py-1.5 outline-none text-xs" style={{ fontFamily: "Inter", border: "none", background: "transparent", color: INK }} />
                  </div>
                </div>

                <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${PAPER_LINE}`, background: "#fff" }}>
                  <table className="w-full" style={{ borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: PAPER }}>
                        <th style={thStyle}>Creator</th>
                        <th style={thStyle}>Người phụ trách</th>
                        <th style={thStyle}>GMV kỳ trước</th>
                        <th style={thStyle}>GMV kỳ này</th>
                        <th style={thStyle}>Phân loại</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredChurnRows.slice(0, 300).map((r: any, i: number) => {
                        const meta = CLASS_META[r.cls];
                        const Icon = meta.icon;
                        return (
                          <tr key={r.name} style={{ borderTop: `1px solid ${PAPER_LINE}`, background: i % 2 ? "#FCFBF8" : "#fff" }}>
                            <td style={{ ...tdStyle, fontWeight: 500 }}>{r.name}</td>
                            <td style={tdStyle}>{r.person}</td>
                            <td style={tdMono}>{fmt(r.prevGmv)}</td>
                            <td style={tdMono}>{fmt(r.currGmv)}</td>
                            <td style={tdStyle}>
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full" style={{ border: `1px solid ${meta.color}`, color: meta.color, fontFamily: "JetBrains Mono", fontSize: 11 }}>
                                <Icon size={11} /> {meta.label}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {filteredChurnRows.length === 0 && (
                    <div className="py-8 text-center" style={{ fontFamily: "Inter", fontSize: 13, color: INK, opacity: 0.4 }}>Không có creator nào khớp bộ lọc</div>
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {tab === "analytics" && (
          <div className="space-y-5">
            {monthList.length === 0 ? (
              <EmptyState text="Chưa có dữ liệu — vào tab Xếp hạng KOC để upload file doanh thu hàng tuần." />
            ) : (
              <>
                <ChartCard title="Xu hướng GMV theo tuần/kỳ">
                  <ResponsiveContainer width="100%" height={240}>
                    <LineChart data={monthlyTrend}>
                      <CartesianGrid stroke={PAPER_LINE} strokeDasharray="3 3" />
                      <XAxis dataKey="month" tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                      <YAxis tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                      <Tooltip formatter={(v: any) => fmt(v) + " đ"} contentStyle={{ fontFamily: "Inter", fontSize: 12, borderRadius: 4 }} />
                      <Line type="monotone" dataKey="gmv" stroke={NEW_BLUE} strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </ChartCard>

                <ChartCard title="GMV theo sản phẩm">
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart data={productBreakdown}>
                      <CartesianGrid stroke={PAPER_LINE} strokeDasharray="3 3" />
                      <XAxis dataKey="name" tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                      <YAxis tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                      <Tooltip formatter={(v: any) => fmt(v) + " đ"} contentStyle={{ fontFamily: "Inter", fontSize: 12, borderRadius: 4 }} />
                      <Bar dataKey="gmv" radius={[3, 3, 0, 0]}>
                        {productBreakdown.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </ChartCard>

                <div className="rounded-sm p-5" style={{ background: "#fff", border: `1px solid ${PAPER_LINE}` }}>
                  <div style={{ fontFamily: "Inter", fontSize: 12, fontWeight: 600, color: INK, marginBottom: 12 }}>KPI theo thành viên</div>
                  {memberKpi.length === 0 ? (
                    <EmptyState text="Chưa có người phụ trách — dán dữ liệu Google Sheets ở tab Tổng hợp KOC." />
                  ) : (
                    <>
                      <ResponsiveContainer width="100%" height={200}>
                        <BarChart data={memberKpi} layout="vertical" margin={{ left: 20 }}>
                          <CartesianGrid stroke={PAPER_LINE} strokeDasharray="3 3" />
                          <XAxis type="number" tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                          <YAxis type="category" dataKey="person" width={100} tick={{ fontFamily: "Inter", fontSize: 11, fill: INK }} />
                          <Tooltip formatter={(v: any) => fmt(v) + " đ"} contentStyle={{ fontFamily: "Inter", fontSize: 12, borderRadius: 4 }} />
                          <Bar dataKey="totalGmv" fill={RETAIN_GREEN} radius={[0, 3, 3, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                      <table className="w-full mt-3" style={{ borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: PAPER }}>
                            <th style={thStyle}>Thành viên</th>
                            <th style={thStyle}>Số creator quản lý</th>
                            <th style={thStyle}>Tổng GMV</th>
                            <th style={thStyle}>Top</th>
                            <th style={thStyle}>Mid</th>
                            <th style={thStyle}>Low</th>
                          </tr>
                        </thead>
                        <tbody>
                          {memberKpi.map((m: any, i: number) => (
                            <tr key={m.person} style={{ borderTop: `1px solid ${PAPER_LINE}`, background: i % 2 ? "#FCFBF8" : "#fff" }}>
                              <td style={{ ...tdStyle, fontWeight: 500 }}>{m.person}</td>
                              <td style={tdMono}>{m.creators}</td>
                              <td style={tdMono}>{fmt(m.totalGmv)} đ</td>
                              <td style={tdMono}>{m.top}</td>
                              <td style={tdMono}>{m.mid}</td>
                              <td style={tdMono}>{m.low}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </>
                  )}
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
