import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const GlobalStyles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: #07090f; font-family: 'IBM Plex Sans', sans-serif; color: #e2e8f0; -webkit-font-smoothing: antialiased; min-height: 100vh; }
    input, select, textarea { font-family: 'IBM Plex Sans', sans-serif; font-size: 16px !important; -webkit-appearance: none; }
    button { touch-action: manipulation; -webkit-tap-highlight-color: transparent; cursor: pointer; }
    ::-webkit-scrollbar { width: 4px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #1e2d3d; border-radius: 2px; }
    @keyframes fadeSlide { from { opacity:0; transform:translateY(6px); } to { opacity:1; transform:translateY(0); } }
    @keyframes blink { 0%,100% { opacity:1; } 50% { opacity:0.3; } }
    .card { animation: fadeSlide 0.22s ease forwards; }
    .blink { animation: blink 1.6s ease-in-out infinite; }
  `}</style>
);

// ── STAŁE ────────────────────────────────────────────────────────
const TASK_TYPES = [
  { id: "regulatory", label: "Regulacyjne", color: "#a78bfa", icon: "📋" },
  { id: "internal",   label: "Wewnętrzne",  color: "#38bdf8", icon: "🏦" },
  { id: "adhoc",      label: "Ad hoc",      color: "#fb923c", icon: "⚡" },
];

const REGULATORS = ["UKNF", "EBA", "GIIF", "UE/AMLA", "Wewnętrzne", "Inne"];
const RECURRENCE = [
  { id: "none",      label: "Jednorazowe" },
  { id: "monthly",   label: "Miesięczne"  },
  { id: "quarterly", label: "Kwartalne"   },
  { id: "annual",    label: "Roczne"      },
];

// ── DANE STARTOWE ────────────────────────────────────────────────
const addDays = (n) => { const d = new Date(); d.setDate(d.getDate()+n); return d.toISOString().split("T")[0]; };
const thisYear = new Date().getFullYear();

const INITIAL = [
  // ── W TRAKCIE ────────────────────────────────────────────────────
  { id:1,  type:"internal",   title:"Połączenia komitetów CIRC KYC", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-04-14`, recurrence:"none", status:"pending", starred:false, completedDate:null, notes:"Przejrzeć regulamin CIRC." },
  { id:2,  type:"regulatory", title:"EA przegląd wykonania ISS-5001189", regulator:"Wewnętrzne", source:"ISS-5001189", deadline:`${thisYear}-04-15`, recurrence:"none", status:"inprogress", starred:false, completedDate:null, notes:"Do zrobienia raport." },
  { id:3,  type:"internal",   title:"Przegląd CRRM", regulator:"Wewnętrzne", source:"Art. 27 Ustawa AML / EBA/GL/2022/05", deadline:`${thisYear}-04-15`, recurrence:"annual", status:"inprogress", starred:true, completedDate:null, notes:"Poszedł mail do grupy." },
  { id:4,  type:"internal",   title:"Scenariusze TM", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-04-15`, recurrence:"none", status:"pending", starred:false, completedDate:null, notes:"Do przejrzenia dokumenty." },
  { id:5,  type:"internal",   title:"CTRLv2", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-04-16`, recurrence:"none", status:"inprogress", starred:true, completedDate:null, notes:"Czekam KYCO." },
  { id:6,  type:"regulatory", title:"Weryfikacja kompletności i adekwatności kontroli AML/CFT w biurze maklerskim", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-04-30`, recurrence:"none", status:"pending", starred:false, completedDate:null, notes:"" },
  { id:7,  type:"internal",   title:"Mapa wdrożenia w tematykę TM", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-04-30`, recurrence:"none", status:"pending", starred:false, completedDate:null, notes:"" },

  // ── ZAKOŃCZONE ───────────────────────────────────────────────────
  { id:8,  type:"internal",   title:"Uchylić procedury", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:9,  type:"internal",   title:"Weryfikacja PCS 2.4 Grupowe", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-16`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-16`, notes:"" },
  { id:10, type:"internal",   title:"Data Governance: Data Usage Agreement #DUA-1005 — Chief Data Officer approval / MBB RAS", regulator:"Wewnętrzne", source:"DUA-1005", deadline:`${thisYear}-01-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:11, type:"internal",   title:"Operational Effectivess (OE) Report", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-31`, recurrence:"quarterly", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:12, type:"internal",   title:"Porządek w kontrolach Lease/Com-", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:13, type:"internal",   title:"Plan KCM 2026 — prośba o weryfikację", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-02`, recurrence:"annual", status:"done", starred:false, completedDate:`${thisYear}-02-02`, notes:"" },
  { id:14, type:"internal",   title:"CTRL 29.01 Komunikacja", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-29`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-29`, notes:"" },
  { id:15, type:"internal",   title:"OE Assessment", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-06`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-06`, notes:"" },
  { id:16, type:"internal",   title:"Kontrole PCS oraz kwartalne", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-02`, recurrence:"quarterly", status:"done", starred:false, completedDate:`${thisYear}-02-02`, notes:"" },
  { id:17, type:"internal",   title:"Zapisać się dot. szkoleń Prince2 Foundation i Practitioner", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-01-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:18, type:"internal",   title:"ISS-0003862 — rekomendacja oczekująca na zamknięcie", regulator:"Wewnętrzne", source:"ISS-0003862", deadline:`${thisYear}-01-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-01-31`, notes:"" },
  { id:19, type:"internal",   title:"Napisać wnioski CIRC", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-13`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-13`, notes:"" },
  { id:20, type:"internal",   title:"Fundacja rodzinne BM", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-13`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-13`, notes:"" },
  { id:21, type:"internal",   title:"Kontrole w UdB", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-10`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-10`, notes:"" },
  { id:22, type:"internal",   title:"Porządek w kontrolach Lease/Com- (luty)", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-28`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-28`, notes:"" },
  { id:23, type:"internal",   title:"IRN-20059145", regulator:"Wewnętrzne", source:"IRN-20059145", deadline:`${thisYear}-02-28`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-28`, notes:"" },
  { id:24, type:"internal",   title:"Zrobić mapowanie Ryzyka, Kontrole GS", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-02-20`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-02-20`, notes:"" },
  { id:25, type:"internal",   title:"Nowy PCS RIA", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-31`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-03-31`, notes:"" },
  { id:26, type:"internal",   title:"Kwartalny raport CIRC", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-13`, recurrence:"quarterly", status:"done", starred:true, completedDate:`${thisYear}-03-13`, notes:"" },
  { id:27, type:"internal",   title:"Audyt WB Lending — obserwacja audytora", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-11`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-03-11`, notes:"" },
  { id:28, type:"regulatory", title:"ISS-5000606 MIA Walidacja", regulator:"Wewnętrzne", source:"ISS-5000606", deadline:`${thisYear}-03-13`, recurrence:"none", status:"done", starred:true, completedDate:`${thisYear}-03-13`, notes:"" },
  { id:29, type:"regulatory", title:"Do walidacji MIA Globalne", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-31`, recurrence:"none", status:"done", starred:true, completedDate:`${thisYear}-03-31`, notes:"" },
  { id:30, type:"internal",   title:"Analiza celowości — uzupełnienie obszar backoffice HOC 1k", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-18`, recurrence:"none", status:"done", starred:false, completedDate:`${thisYear}-03-18`, notes:"" },
  { id:31, type:"regulatory", title:"Opinia prawna dotycząca weryfikacji tożsamości", regulator:"Wewnętrzne", source:"", deadline:`${thisYear}-03-31`, recurrence:"none", status:"done", starred:true, completedDate:`${thisYear}-03-31`, notes:"" },
  { id:32, type:"regulatory", title:"Szkolenie AML", regulator:"Wewnętrzne", source:"Art. 52 Ustawa AML", deadline:`${thisYear}-03-31`, recurrence:"annual", status:"done", starred:true, completedDate:`${thisYear}-03-31`, notes:"" },
  { id:33, type:"regulatory", title:"BION", regulator:"UKNF", source:"", deadline:`${thisYear}-03-27`, recurrence:"annual", status:"done", starred:true, completedDate:`${thisYear}-03-27`, notes:"" },
  { id:34, type:"regulatory", title:"Stanowisko KNF z dnia 19 marca 2026 dot. stosowania AML — analiza", regulator:"UKNF", source:"Stanowisko KNF 19.03.2026", deadline:`${thisYear}-03-31`, recurrence:"none", status:"done", starred:true, completedDate:`${thisYear}-03-31`, notes:"" },
];

// ── HELPERS ──────────────────────────────────────────────────────
const getDays   = (d) => { const t=new Date(); t.setHours(0,0,0,0); return Math.ceil((new Date(d)-t)/86400000); };
const fmtDate   = (iso) => iso ? new Date(iso).toLocaleDateString("pl-PL",{day:"2-digit",month:"short",year:"numeric"}) : "—";
const fmtDateLong = (iso) => iso ? new Date(iso).toLocaleDateString("pl-PL",{day:"2-digit",month:"long",year:"numeric"}) : "—";
const getType   = (id)  => TASK_TYPES.find(t=>t.id===id)||TASK_TYPES[0];
const getRec    = (id)  => RECURRENCE.find(r=>r.id===id)?.label||"—";

const getLight = (days, status) => {
  if (status==="done") return { color:"#10b981", bg:"#061410", border:"#10b98133", ring:"#10b98155" };
  if (days<0)          return { color:"#ef4444", bg:"#130808", border:"#ef444455", ring:"#ef444488" };
  if (days<=7)         return { color:"#ef4444", bg:"#130808", border:"#ef444444", ring:"#ef444466" };
  if (days<=21)        return { color:"#f59e0b", bg:"#111008", border:"#f59e0b44", ring:"#f59e0b55" };
  return                      { color:"#10b981", bg:"#061410", border:"#10b98133", ring:"#10b98144" };
};

const inputSt = {
  width:"100%", background:"#0d1117", border:"1px solid #1e2d3d",
  borderRadius:10, padding:"11px 14px", color:"#e2e8f0",
  fontSize:16, outline:"none", fontFamily:"'IBM Plex Sans', sans-serif",
};

// ── EKSPORT EXCEL ────────────────────────────────────────────────
const exportExcel = (items) => {
  const wb = XLSX.utils.book_new();
  const rows = [...items].sort((a,b)=>{
    const o = {"done":99}; const da=getDays(a.deadline); const db=getDays(b.deadline);
    if(a.status==="done"&&b.status!=="done") return 1;
    if(b.status==="done"&&a.status!=="done") return -1;
    return da-db;
  }).map(t=>({
    "⭐": t.starred ? "★" : "",
    "Typ": getType(t.type).label,
    "Tytuł": t.title,
    "Regulator": t.regulator,
    "Podstawa prawna": t.source,
    "Termin": t.deadline,
    "Dni": t.status==="done"?"✓":getDays(t.deadline)<0?`-${Math.abs(getDays(t.deadline))}d`:`${getDays(t.deadline)}d`,
    "Powtarzalność": getRec(t.recurrence),
    "Status": t.status==="done"?"Zakończone":t.status==="inprogress"?"W toku":"Do zrobienia",
    "Data zakończenia": t.completedDate||"",
    "Notatki": t.notes,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{wch:4},{wch:14},{wch:50},{wch:12},{wch:22},{wch:13},{wch:8},{wch:14},{wch:14},{wch:16},{wch:55}];
  XLSX.utils.book_append_sheet(wb, ws, "Zadania");

  const today = new Date().toISOString().split("T")[0];
  const sum = [
    {"Statystyka":"Data raportu","Wartość":today},
    {"Statystyka":"Wszystkie","Wartość":items.length},
    {"Statystyka":"Regulacyjne","Wartość":items.filter(x=>x.type==="regulatory").length},
    {"Statystyka":"Wewnętrzne","Wartość":items.filter(x=>x.type==="internal").length},
    {"Statystyka":"Ad hoc","Wartość":items.filter(x=>x.type==="adhoc").length},
    {"Statystyka":"Zakończone","Wartość":items.filter(x=>x.status==="done").length},
    {"Statystyka":"Przeterminowane","Wartość":items.filter(x=>x.status!=="done"&&getDays(x.deadline)<0).length},
    {"Statystyka":"Pilne ≤7d","Wartość":items.filter(x=>x.status!=="done"&&getDays(x.deadline)>=0&&getDays(x.deadline)<=7).length},
    {"Statystyka":"Oznaczone ★","Wartość":items.filter(x=>x.starred).length},
  ];
  const wss = XLSX.utils.json_to_sheet(sum);
  wss["!cols"]=[{wch:22},{wch:12}];
  XLSX.utils.book_append_sheet(wb, wss, "Podsumowanie");
  XLSX.writeFile(wb, `Compliance_${today}.xlsx`);
};

// ── OKRESY RAPORTOWANIA ───────────────────────────────────────────
const REPORT_PERIODS = [
  { id: `${thisYear}-h1`, label: `H1 ${thisYear}`, short: "Półrocze I",   from: `${thisYear}-01-01`, to: `${thisYear}-06-30` },
  { id: `${thisYear}-h2`, label: `H2 ${thisYear}`, short: "Półrocze II",  from: `${thisYear}-07-01`, to: `${thisYear}-12-31` },
  { id: `${thisYear}`,    label: `Rok ${thisYear}`, short: "Rok pełny",   from: `${thisYear}-01-01`, to: `${thisYear}-12-31` },
  { id: `${thisYear-1}`,  label: `Rok ${thisYear-1}`, short: "Rok poprzedni", from: `${thisYear-1}-01-01`, to: `${thisYear-1}-12-31` },
];

const getPeriodItems = (items, period) => {
  const p = REPORT_PERIODS.find(x => x.id === period);
  if (!p) return items;
  return items.filter(x => {
    const d = x.completedDate || x.deadline;
    return d && d >= p.from && d <= p.to;
  });
};

// ── BUDUJ DANE RAPORTU ────────────────────────────────────────────
const buildReportData = (items, periodId) => {
  const period   = REPORT_PERIODS.find(x => x.id === periodId);
  const filtered = getPeriodItems(items, periodId);
  const done      = filtered.filter(x => x.status === "done");
  const starred   = done.filter(x => x.starred);
  const regulatory= done.filter(x => x.type === "regulatory");
  const internal  = done.filter(x => x.type === "internal");
  const adhoc     = done.filter(x => x.type === "adhoc");
  const implemented = done.filter(x => x.type === "regulatory" && x.source);
  return { period, done, starred, regulatory, internal, adhoc, implemented };
};

// ── PODGLĄD RAPORTU (w apce) ──────────────────────────────────────
const ReportPreview = ({ items, periodId, onClose, onPrint }) => {
  const { period, done, starred, regulatory, internal, adhoc, implemented } = buildReportData(items, periodId);

  const SectionHeader = ({ color, title, count }) => (
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
      borderBottom:`2px solid ${color}33`, paddingBottom:8, marginBottom:10, marginTop:20 }}>
      <span style={{ fontSize:11, fontWeight:700, color, textTransform:"uppercase", letterSpacing:"0.1em" }}>{title}</span>
      <span style={{ fontFamily:"'IBM Plex Mono', monospace", fontSize:11, color:"#475569" }}>{count} zadań</span>
    </div>
  );

  const TaskRow = ({ item, highlight=false }) => {
    const type = getType(item.type);
    return (
      <div style={{ padding:"10px 12px", borderRadius:10, marginBottom:8,
        background: highlight ? "#111208" : "#0d1117",
        border:`1px solid ${highlight ? "#f59e0b33" : "#1e2d3d"}`,
        borderLeft:`3px solid ${highlight ? "#f59e0b" : type.color}`,
      }}>
        <div style={{ fontSize:13, fontWeight:600, color:"#e2e8f0", marginBottom:5 }}>
          {item.starred && <span style={{ color:"#f59e0b", marginRight:5 }}>★</span>}
          {item.title}
        </div>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          <span style={{ fontSize:10, color:type.color, background:type.color+"18",
            border:`1px solid ${type.color}44`, borderRadius:4, padding:"1px 7px", fontWeight:700 }}>
            {type.icon} {type.label}
          </span>
          {item.regulator && item.regulator !== "Wewnętrzne" && (
            <span style={{ fontSize:10, color:"#64748b" }}>🏛 {item.regulator}</span>
          )}
          {item.source && (
            <span style={{ fontFamily:"'IBM Plex Mono', monospace", fontSize:10, color:"#475569" }}>§ {item.source}</span>
          )}
          <span style={{ fontFamily:"'IBM Plex Mono', monospace", fontSize:10, color:"#10b981" }}>
            📅 {fmtDate(item.completedDate || item.deadline)}
          </span>
        </div>
        {item.notes && (
          <div style={{ fontSize:11, color:"#64748b", marginTop:6,
            borderLeft:"2px solid #1e2d3d", paddingLeft:8, lineHeight:1.5 }}>{item.notes}</div>
        )}
      </div>
    );
  };

  return (
    <div style={{ position:"fixed", inset:0, zIndex:150, background:"rgba(0,0,0,0.92)",
      backdropFilter:"blur(10px)", display:"flex", flexDirection:"column",
      alignItems:"center", overflowY:"auto", padding:"20px 0 40px" }}>

      <div style={{ width:"100%", maxWidth:520, padding:"0 16px" }}>

        {/* Nagłówek podglądu */}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
          marginBottom:20, position:"sticky", top:0, zIndex:10,
          background:"rgba(7,9,15,0.95)", backdropFilter:"blur(8px)", padding:"16px 0 12px" }}>
          <div>
            <div style={{ fontWeight:700, fontSize:17 }}>Podgląd raportu</div>
            <div style={{ fontSize:11, color:"#475569", fontFamily:"'IBM Plex Mono', monospace", marginTop:2 }}>
              {period?.label} · {period?.short}
            </div>
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <button onClick={onPrint} style={{ background:"linear-gradient(135deg,#1e3a5f,#312e81)",
              border:"1px solid #2563eb55", borderRadius:10, padding:"8px 14px",
              color:"#93c5fd", fontWeight:700, fontSize:12 }}>🖨 Drukuj / PDF</button>
            <button onClick={onClose} style={{ background:"#1e2d3d", border:"none",
              borderRadius:10, padding:"8px 12px", color:"#94a3b8", fontSize:16 }}>×</button>
          </div>
        </div>

        {/* Baner edycji */}
        <div style={{ background:"#111208", border:"1px solid #f59e0b33",
          borderRadius:10, padding:"10px 14px", marginBottom:16,
          fontSize:12, color:"#92400e" }}>
          📝 <strong style={{ color:"#f59e0b" }}>Dokument do edycji</strong>
          <span style={{ color:"#78716c" }}> — po wydruku uzupełnij o kontekst przed rozmową</span>
        </div>

        {/* Statystyki */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr 1fr", gap:6, marginBottom:20 }}>
          {[
            { val:done.length,       label:"Łącznie",   color:"#10b981" },
            { val:regulatory.length, label:"Regulac.",  color:"#a78bfa" },
            { val:internal.length,   label:"Wewn.",     color:"#38bdf8" },
            { val:adhoc.length,      label:"Ad hoc",    color:"#fb923c" },
            { val:starred.length,    label:"★ Kluczowe",color:"#f59e0b" },
          ].map(({ val, label, color }) => (
            <div key={label} style={{ background:"#0d1117", border:`1px solid ${color}33`,
              borderRadius:10, padding:"8px 4px", textAlign:"center" }}>
              <div style={{ fontFamily:"'IBM Plex Mono', monospace", fontSize:20,
                fontWeight:700, color, lineHeight:1 }}>{val}</div>
              <div style={{ fontSize:8, color:"#475569", fontWeight:700,
                textTransform:"uppercase", letterSpacing:"0.04em", marginTop:3 }}>{label}</div>
            </div>
          ))}
        </div>

        {/* Kluczowe projekty ★ */}
        <SectionHeader color="#f59e0b" title="Kluczowe projekty i inicjatywy ★" count={starred.length}/>
        {starred.length > 0
          ? starred.map(t => <TaskRow key={t.id} item={t} highlight/>)
          : <div style={{ fontSize:12, color:"#334155", padding:"12px 0", textAlign:"center" }}>Brak zadań oznaczonych ★</div>
        }

        {/* Wdrożone regulacje */}
        {implemented.length > 0 && <>
          <SectionHeader color="#a78bfa" title="Wdrożone regulacje / przepisy" count={implemented.length}/>
          {implemented.map(t => <TaskRow key={t.id} item={t}/>)}
        </>}

        {/* Regulacyjne */}
        {regulatory.length > 0 && <>
          <SectionHeader color="#a78bfa" title="📋 Regulacyjne" count={regulatory.length}/>
          {regulatory.map(t => <TaskRow key={t.id} item={t}/>)}
        </>}

        {/* Wewnętrzne */}
        {internal.length > 0 && <>
          <SectionHeader color="#38bdf8" title="🏦 Wewnętrzne" count={internal.length}/>
          {internal.map(t => <TaskRow key={t.id} item={t}/>)}
        </>}

        {/* Ad hoc */}
        {adhoc.length > 0 && <>
          <SectionHeader color="#fb923c" title="⚡ Ad hoc" count={adhoc.length}/>
          {adhoc.map(t => <TaskRow key={t.id} item={t}/>)}
        </>}

        {done.length === 0 && (
          <div style={{ textAlign:"center", padding:"40px 0", color:"#334155" }}>
            <div style={{ fontSize:32, marginBottom:12 }}>📭</div>
            <div style={{ fontSize:13 }}>Brak zakończonych zadań w tym okresie</div>
          </div>
        )}

        {/* Stopka */}
        <div style={{ marginTop:24, paddingTop:14, borderTop:"1px solid #1e2d3d",
          fontSize:10, color:"#334155", textAlign:"center",
          fontFamily:"'IBM Plex Mono', monospace" }}>
          Compliance Tracker PRO · {fmtDate(new Date().toISOString().split("T")[0])}
        </div>
      </div>
    </div>
  );
};

// ── GENERUJ HTML DO DRUKU ─────────────────────────────────────────
const generatePrintReport = (items, periodId) => {
  const { period, done, starred, regulatory, internal, adhoc, implemented } = buildReportData(items, periodId);

  const section = (title, color, list, showSource=false) => {
    if (!list.length) return "";
    return `
      <h3 style="color:${color};font-size:12px;text-transform:uppercase;letter-spacing:0.1em;margin:20px 0 8px;padding-bottom:6px;border-bottom:2px solid ${color}33">${title} (${list.length})</h3>
      ${list.map(t => `
        <div style="margin-bottom:10px;padding:10px 14px;border:1px solid #e2e8f0;border-left:3px solid ${color};border-radius:6px;break-inside:avoid">
          <div style="font-weight:700;font-size:13px;margin-bottom:4px">${t.starred?"★ ":""}${t.title}</div>
          <div style="display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:#64748b">
            ${showSource&&t.source?`<span>§ ${t.source}</span>`:""}
            ${t.regulator&&t.regulator!=="Wewnętrzne"?`<span>🏛 ${t.regulator}</span>`:""}
            <span>📅 ${fmtDateLong(t.completedDate||t.deadline)}</span>
          </div>
          ${t.notes?`<div style="font-size:11px;color:#475569;margin-top:5px;padding-left:8px;border-left:2px solid #e2e8f0">${t.notes}</div>`:""}
        </div>`).join("")}`;
  };

  const html = `<!DOCTYPE html><html lang="pl"><head><meta charset="UTF-8">
    <title>Podsumowanie Compliance — ${period?.label}</title>
    <style>
      body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;max-width:800px;margin:0 auto;padding:32px;font-size:14px;line-height:1.5}
      h1{font-size:24px;font-weight:700;margin-bottom:4px}
      h2{font-size:16px;font-weight:700;margin:28px 0 12px;text-transform:uppercase;letter-spacing:0.08em}
      .subtitle{color:#64748b;font-size:13px;margin-bottom:24px;padding-bottom:14px;border-bottom:2px solid #e2e8f0}
      .stats{display:flex;gap:12px;margin-bottom:28px;flex-wrap:wrap}
      .stat{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;text-align:center;flex:1;min-width:90px}
      .stat-val{font-size:26px;font-weight:700;margin-bottom:2px}
      .stat-lbl{font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.08em}
      .note{background:#fffbeb;border:1px solid #fcd34d;border-radius:6px;padding:10px 14px;font-size:12px;color:#92400e;margin-bottom:20px}
      @media print{body{padding:16px}.note{display:none}}
    </style></head><body>
    <h1>Podsumowanie Compliance</h1>
    <div class="subtitle">
      <strong>${period?.label}</strong> · ${period?.from} – ${period?.to}
      &nbsp;·&nbsp; Wygenerowano: ${fmtDateLong(new Date().toISOString().split("T")[0])}
    </div>
    <div class="note">📝 <strong>Dokument do edycji</strong> — uzupełnij o kontekst i wagi przed rozmową oceniającą.</div>
    <h2>Statystyki</h2>
    <div class="stats">
      <div class="stat"><div class="stat-val" style="color:#10b981">${done.length}</div><div class="stat-lbl">Zakończonych</div></div>
      <div class="stat"><div class="stat-val" style="color:#a78bfa">${regulatory.length}</div><div class="stat-lbl">Regulacyjnych</div></div>
      <div class="stat"><div class="stat-val" style="color:#38bdf8">${internal.length}</div><div class="stat-lbl">Wewnętrznych</div></div>
      <div class="stat"><div class="stat-val" style="color:#fb923c">${adhoc.length}</div><div class="stat-lbl">Ad hoc</div></div>
      <div class="stat"><div class="stat-val" style="color:#f59e0b">${starred.length}</div><div class="stat-lbl">Kluczowych ★</div></div>
    </div>
    <h2>Kluczowe projekty i inicjatywy ★</h2>
    ${starred.length ? starred.map(t=>`
      <div style="margin-bottom:12px;padding:12px 16px;background:#fffbeb;border:1px solid #fcd34d44;border-left:4px solid #f59e0b;border-radius:8px;break-inside:avoid">
        <div style="font-weight:700;font-size:14px;margin-bottom:5px">★ ${t.title}</div>
        <div style="display:flex;gap:14px;font-size:11px;color:#64748b;flex-wrap:wrap;margin-bottom:4px">
          <span style="background:${getType(t.type).color}18;color:${getType(t.type).color};border:1px solid ${getType(t.type).color}44;border-radius:4px;padding:1px 7px;font-weight:700">${getType(t.type).label}</span>
          ${t.regulator&&t.regulator!=="Wewnętrzne"?`<span>🏛 ${t.regulator}</span>`:""}
          ${t.source?`<span>§ ${t.source}</span>`:""}
          <span>📅 ${fmtDateLong(t.completedDate||t.deadline)}</span>
        </div>
        ${t.notes?`<div style="font-size:12px;color:#78716c;margin-top:4px">${t.notes}</div>`:""}
      </div>`).join("") : "<p style='color:#94a3b8;font-size:13px'>Brak zadań oznaczonych gwiazdką</p>"}
    <h2>Wdrożone regulacje / przepisy</h2>
    ${section("Regulacyjne", "#6d28d9", implemented, true)}
    ${!implemented.length?"<p style='color:#94a3b8;font-size:13px'>Brak wdrożeń w tym okresie</p>":""}
    <h2>Wszystkie zakończone zadania</h2>
    ${section("📋 Regulacyjne", "#7c3aed", regulatory, true)}
    ${section("🏦 Wewnętrzne", "#0284c7", internal)}
    ${section("⚡ Ad hoc", "#ea580c", adhoc)}
    ${!done.length?"<p style='color:#94a3b8;font-size:13px'>Brak zakończonych zadań w tym okresie</p>":""}
    <div style="margin-top:40px;padding-top:14px;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;text-align:center">
      Compliance Tracker PRO · ${fmtDateLong(new Date().toISOString().split("T")[0])}
    </div>
  </body></html>`;

  const w = window.open("", "_blank");
  w.document.write(html);
  w.document.close();
  setTimeout(() => w.print(), 600);
};

// ── DANE STARTOWE — CELE ROCZNE ──────────────────────────────────
const INITIAL_GOALS = [
  {
    id: 1,
    title: "Zapewnię sprawną i zgodną ze strategią ESG realizację zadań BAU oraz będę skutecznie zarządzać wąskimi gardłami i ryzykami",
    color: "#a78bfa",
    subtasks: [
      { id: 11, title: "Zadania BAU realizowane zgodnie z ustaloną normą jakości/czasu a ich wpływ na cele ESG świadomie uwzględniany", done: false, completedDate: null },
      { id: 12, title: "Wąskie gardła i ryzyka identyfikowane i adresowane na bieżąco. Zagrożenia zgłaszane i eskalowane zgodnie ze standardem", done: false, completedDate: null },
    ]
  },
  {
    id: 2,
    title: "Usprawnienia identyfikowane i wdrażane, z widoczną poprawą efektywności, jakości lub redukcją ryzyka — min. 2/rok",
    color: "#38bdf8",
    subtasks: [
      { id: 21, title: "Usunięcie ODD PL_002", done: false, completedDate: null },
      { id: 22, title: "Zmiana kontroli w spółkach zależnych co ma pozytywny wpływ na wynik Compliance", done: true, completedDate: `${thisYear}-02-01` },
      { id: 23, title: "Zmniejszenie liczby wpływających wniosków na Komitet CIRC", done: true, completedDate: `${thisYear}-02-01` },
      { id: 24, title: "Zmniejszenie liczby wpływających maili z prośbą o rekomendacje do zespołu FCC Operations", done: true, completedDate: `${thisYear}-03-26` },
      { id: 25, title: "Backlog zarządzany zgodnie ze standardem (Azure). Zadania bez uzasadnienia (niezgodne z rolą/bez wartości dodanej/nieistotne z perspektywy ryzyka) eliminowane", done: false, completedDate: null },
      { id: 26, title: "Przygotowanie i realizacja planu wdrożenia w obszar TM, zakończone samodzielnym i poprawnym wykonywaniem zadań w tym zakresie", done: false, completedDate: null },
      { id: 27, title: "Spotkania FCC TM", done: false, completedDate: null },
      { id: 28, title: "Analiza Global Set Function", done: false, completedDate: null },
    ]
  },
  {
    id: 3,
    title: "W ramach projektu Golden, dostarczam na czas i w oczekiwanej jakości zadania, które zapewniające integrację Banku z Goldman Sachs TFI w obszarach dla których jestem SME",
    color: "#fb923c",
    subtasks: [
      { id: 31, title: "Przygotowanie ryzyk i kontroli z obszarów FCC", done: true, completedDate: `${thisYear}-02-25` },
    ]
  },
];

// ── WIDOK CELÓW ROCZNYCH ─────────────────────────────────────────
const GoalsView = ({ goals, setGoals }) => {
  const [expandedGoal, setExpandedGoal] = useState(null);
  const [addingTo, setAddingTo]         = useState(null);
  const [newSubtask, setNewSubtask]     = useState("");
  const [completingId, setCompletingId] = useState(null); // { goalId, subtaskId }
  const [editingGoal, setEditingGoal]   = useState(null);
  const [newGoalTitle, setNewGoalTitle] = useState("");
  const [addingGoal, setAddingGoal]     = useState(false);

  const saveGoals = (next) => {
    setGoals(next);
    try { localStorage.setItem("ct_goals", JSON.stringify(next)); } catch {}
  };

  const toggleSubtask = (goalId, subtaskId) => {
    const goal = goals.find(g => g.id === goalId);
    const sub  = goal?.subtasks.find(s => s.id === subtaskId);
    if (!sub) return;
    if (sub.done) {
      // odznacz
      saveGoals(goals.map(g => g.id !== goalId ? g : {
        ...g, subtasks: g.subtasks.map(s => s.id !== subtaskId ? s : { ...s, done: false, completedDate: null })
      }));
    } else {
      setCompletingId({ goalId, subtaskId });
    }
  };

  const confirmSubDone = (date) => {
    const { goalId, subtaskId } = completingId;
    saveGoals(goals.map(g => g.id !== goalId ? g : {
      ...g, subtasks: g.subtasks.map(s => s.id !== subtaskId ? s : { ...s, done: true, completedDate: date })
    }));
    setCompletingId(null);
  };

  const deleteSubtask = (goalId, subtaskId) => {
    saveGoals(goals.map(g => g.id !== goalId ? g : {
      ...g, subtasks: g.subtasks.filter(s => s.id !== subtaskId)
    }));
  };

  const addSubtask = (goalId) => {
    if (!newSubtask.trim()) return;
    saveGoals(goals.map(g => g.id !== goalId ? g : {
      ...g, subtasks: [...g.subtasks, { id: Date.now(), title: newSubtask.trim(), done: false, completedDate: null }]
    }));
    setNewSubtask("");
    setAddingTo(null);
  };

  const deleteGoal = (goalId) => {
    saveGoals(goals.filter(g => g.id !== goalId));
  };

  const addGoal = () => {
    if (!newGoalTitle.trim()) return;
    const colors = ["#a78bfa","#38bdf8","#fb923c","#10b981","#f59e0b","#ec4899"];
    saveGoals([...goals, {
      id: Date.now(),
      title: newGoalTitle.trim(),
      color: colors[goals.length % colors.length],
      subtasks: []
    }]);
    setNewGoalTitle("");
    setAddingGoal(false);
  };

  const exportGoalsExcel = () => {
    const wb = XLSX.utils.book_new();
    const rows = [];
    goals.forEach(g => {
      const done = g.subtasks.filter(s => s.done).length;
      rows.push({ "Cel": g.title, "Podzadanie": "", "Status": `${done}/${g.subtasks.length}`, "Data realizacji": "" });
      g.subtasks.forEach(s => {
        rows.push({ "Cel": "", "Podzadanie": s.title, "Status": s.done ? "✓ Zakończone" : "W toku", "Data realizacji": s.completedDate || "" });
      });
      rows.push({ "Cel": "", "Podzadanie": "", "Status": "", "Data realizacji": "" });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{wch:60},{wch:60},{wch:16},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws, "Cele roczne");
    XLSX.writeFile(wb, `Cele_${thisYear}_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  return (
    <div style={{padding:"8px 16px 140px",display:"flex",flexDirection:"column",gap:12}}>

      {/* Pasek podsumowania */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:4}}>
        {[
          { val: goals.length, label: "Cele", color: "#a78bfa" },
          { val: goals.reduce((a,g)=>a+g.subtasks.length,0), label: "Podzadania", color: "#38bdf8" },
          { val: goals.reduce((a,g)=>a+g.subtasks.filter(s=>s.done).length,0), label: "Zrobione ✓", color: "#10b981" },
        ].map(({val,label,color})=>(
          <div key={label} style={{background:"#0d1117",border:`1px solid ${color}33`,borderRadius:10,padding:"8px 6px",textAlign:"center"}}>
            <div style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:22,fontWeight:600,color,lineHeight:1}}>{val}</div>
            <div style={{fontSize:9,color:"#475569",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginTop:3}}>{label}</div>
          </div>
        ))}
      </div>

      {/* Lista celów */}
      {goals.map(goal => {
        const doneCount = goal.subtasks.filter(s => s.done).length;
        const total     = goal.subtasks.length;
        const pct       = total ? Math.round(doneCount / total * 100) : 0;
        const expanded  = expandedGoal === goal.id;

        return (
          <div key={goal.id} className="card" style={{background:"#0d1117",border:`1px solid ${goal.color}33`,borderLeft:`3px solid ${goal.color}`,borderRadius:14,overflow:"hidden"}}>

            {/* Nagłówek celu */}
            <div style={{padding:"14px 14px 10px",cursor:"pointer"}} onClick={()=>setExpandedGoal(expanded ? null : goal.id)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
                <div style={{fontSize:13,fontWeight:600,color:"#e2e8f0",lineHeight:1.45,flex:1}}>{goal.title}</div>
                <div style={{display:"flex",gap:4,flexShrink:0}}>
                  <button onClick={e=>{e.stopPropagation();deleteGoal(goal.id);}} style={{background:"none",border:"none",color:"#334155",fontSize:16,padding:"2px 4px",lineHeight:1}}>×</button>
                  <span style={{fontSize:16,color:goal.color,transition:"transform 0.2s",display:"inline-block",transform:expanded?"rotate(180deg)":"rotate(0deg)"}}>▾</span>
                </div>
              </div>

              {/* Progress bar */}
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{flex:1,height:5,background:"#1e2d3d",borderRadius:3,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${pct}%`,background:goal.color,borderRadius:3,transition:"width 0.4s"}}/>
                </div>
                <span style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:10,color:goal.color,fontWeight:700,whiteSpace:"nowrap"}}>
                  {doneCount}/{total} · {pct}%
                </span>
              </div>
            </div>

            {/* Podzadania */}
            {expanded && (
              <div style={{borderTop:`1px solid ${goal.color}22`,padding:"8px 14px 12px",display:"flex",flexDirection:"column",gap:6}}>
                {goal.subtasks.map(sub => (
                  <div key={sub.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"8px 10px",borderRadius:9,background:sub.done?"#061410":"#060b14",border:`1px solid ${sub.done?"#10b98122":"#1e2d3d"}`}}>
                    <button onClick={()=>toggleSubtask(goal.id, sub.id)} style={{width:20,height:20,borderRadius:5,border:`2px solid ${sub.done?"#10b981":goal.color+"88"}`,background:sub.done?"#10b981":"transparent",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",marginTop:1}}>
                      {sub.done && <span style={{color:"white",fontSize:11,fontWeight:700}}>✓</span>}
                    </button>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,color:sub.done?"#475569":"#cbd5e1",textDecoration:sub.done?"line-through":"none",lineHeight:1.4}}>{sub.title}</div>
                      {sub.done && sub.completedDate && (
                        <div style={{fontSize:10,color:"#10b981",fontFamily:"'IBM Plex Mono', monospace",marginTop:3}}>✓ {fmtDate(sub.completedDate)}</div>
                      )}
                    </div>
                    <button onClick={()=>deleteSubtask(goal.id, sub.id)} style={{background:"none",border:"none",color:"#1e2d3d",fontSize:14,flexShrink:0,padding:"0 2px"}}>×</button>
                  </div>
                ))}

                {/* Dodaj podzadanie */}
                {addingTo === goal.id ? (
                  <div style={{display:"flex",gap:6,marginTop:4}}>
                    <input
                      autoFocus
                      value={newSubtask}
                      onChange={e=>setNewSubtask(e.target.value)}
                      onKeyDown={e=>{ if(e.key==="Enter") addSubtask(goal.id); if(e.key==="Escape") { setAddingTo(null); setNewSubtask(""); } }}
                      placeholder="Nowe podzadanie..."
                      style={{...inputSt,flex:1,padding:"8px 12px",fontSize:13}}
                    />
                    <button onClick={()=>addSubtask(goal.id)} style={{background:`${goal.color}22`,border:`1px solid ${goal.color}55`,borderRadius:9,padding:"8px 12px",color:goal.color,fontWeight:700,fontSize:12}}>Dodaj</button>
                    <button onClick={()=>{setAddingTo(null);setNewSubtask("");}} style={{background:"transparent",border:"1px solid #1e2d3d",borderRadius:9,padding:"8px 10px",color:"#475569",fontSize:12}}>✕</button>
                  </div>
                ) : (
                  <button onClick={()=>setAddingTo(goal.id)} style={{marginTop:4,background:"transparent",border:`1px dashed ${goal.color}44`,borderRadius:9,padding:"7px 0",color:goal.color+"99",fontSize:12,fontWeight:600,width:"100%"}}>+ Dodaj podzadanie</button>
                )}
              </div>
            )}
          </div>
        );
      })}

      {/* Dodaj nowy cel */}
      {addingGoal ? (
        <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:14,padding:"14px"}}>
          <div style={{fontSize:11,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:8}}>Nowy cel roczny</div>
          <textarea
            autoFocus
            value={newGoalTitle}
            onChange={e=>setNewGoalTitle(e.target.value)}
            placeholder="Treść celu..."
            rows={3}
            style={{...inputSt,resize:"vertical",marginBottom:10}}
          />
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>{setAddingGoal(false);setNewGoalTitle("");}} style={{flex:1,background:"transparent",border:"1px solid #1e2d3d",borderRadius:10,padding:"10px 0",color:"#475569",fontWeight:700,fontSize:13}}>Anuluj</button>
            <button onClick={addGoal} style={{flex:2,background:"linear-gradient(135deg,#1e3a5f,#312e81)",border:"1px solid #2563eb55",borderRadius:10,padding:"10px 0",color:"#93c5fd",fontWeight:700,fontSize:13}}>Dodaj cel</button>
          </div>
        </div>
      ) : (
        <button onClick={()=>setAddingGoal(true)} style={{background:"transparent",border:"1px dashed #1e2d3d",borderRadius:14,padding:"14px 0",color:"#334155",fontSize:13,fontWeight:600}}>+ Nowy cel roczny</button>
      )}

      {/* Eksport */}
      <button onClick={exportGoalsExcel} style={{background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:12,padding:"11px 0",color:"#34d399",fontWeight:700,fontSize:13}}>⬇ Eksport celów do Excel</button>

      {/* Modal — data zakończenia podzadania */}
      {completingId && (
        <div style={{position:"fixed",inset:0,zIndex:200,background:"rgba(0,0,0,0.9)",backdropFilter:"blur(8px)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:16,padding:"24px 20px",width:"100%",maxWidth:360}}>
            <div style={{fontWeight:700,fontSize:16,marginBottom:6}}>Data realizacji</div>
            <div style={{fontSize:12,color:"#475569",marginBottom:16}}>Kiedy zostało zrealizowane podzadanie?</div>
            <input type="date" defaultValue={new Date().toISOString().split("T")[0]} id="subCompleteDate" style={{...inputSt,marginBottom:14}}/>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setCompletingId(null)} style={{flex:1,background:"transparent",border:"1px solid #1e2d3d",borderRadius:10,padding:"11px 0",color:"#475569",fontWeight:700,fontSize:13}}>Anuluj</button>
              <button onClick={()=>confirmSubDone(document.getElementById("subCompleteDate").value)} style={{flex:2,background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:10,padding:"11px 0",color:"#34d399",fontWeight:700,fontSize:13}}>✓ Zakończ</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ── KOMPONENTY ────────────────────────────────────────────────────
const TypeBadge = ({typeId}) => {
  const t = getType(typeId);
  return <span style={{background:t.color+"18",color:t.color,border:`1px solid ${t.color}44`,borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700,fontFamily:"'IBM Plex Mono', monospace"}}>{t.icon} {t.label}</span>;
};

const RegBadge = ({name}) => {
  const c = {UKNF:"#38bdf8",EBA:"#a78bfa",GIIF:"#fb923c","UE/AMLA":"#facc15",Wewnętrzne:"#64748b",Inne:"#94a3b8"}[name]||"#94a3b8";
  return <span style={{background:c+"18",color:c,border:`1px solid ${c}44`,borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700,fontFamily:"'IBM Plex Mono', monospace"}}>{name}</span>;
};

const Card = ({item, onEdit, onDelete, onStatus, onStar, onComplete}) => {
  const days = getDays(item.deadline);
  const tl   = getLight(days, item.status);
  const urgent = item.status!=="done" && days<=7;

  return (
    <div className="card" style={{background:tl.bg,border:`1px solid ${tl.border}`,borderLeft:`3px solid ${tl.color}`,borderRadius:14,padding:"14px 16px",opacity:item.status==="done"?0.65:1}}>

      {/* Górna sekcja */}
      <div style={{display:"flex",gap:10,alignItems:"flex-start",marginBottom:10}}>
        <div style={{width:11,height:11,borderRadius:"50%",background:tl.color,boxShadow:`0 0 7px ${tl.ring}`,flexShrink:0,marginTop:3}} className={urgent?"blink":""}/>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:14,fontWeight:600,lineHeight:1.4,color:item.status==="done"?"#64748b":"#e2e8f0",textDecoration:item.status==="done"?"line-through":"none",marginBottom:7}}>
            {item.starred && <span style={{color:"#f59e0b",marginRight:5}}>★</span>}
            {item.title}
          </div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            <TypeBadge typeId={item.type}/>
            {item.regulator && <RegBadge name={item.regulator}/>}
            {item.source && <span style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:10,color:"#475569",background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:5,padding:"2px 8px"}}>{item.source}</span>}
            {item.recurrence!=="none" && <span style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:10,color:"#334155",background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:5,padding:"2px 8px"}}>🔄 {getRec(item.recurrence)}</span>}
          </div>
        </div>

        {/* Akcje */}
        <div style={{display:"flex",gap:4,flexShrink:0,alignItems:"center"}}>
          <button onClick={()=>onStar(item.id)} title={item.starred?"Usuń gwiazdkę":"Oznacz jako ważne"} style={{background:item.starred?"#1a1208":"transparent",border:`1px solid ${item.starred?"#f59e0b44":"#1e2d3d"}`,borderRadius:7,padding:"4px 7px",color:item.starred?"#f59e0b":"#334155",fontSize:14}}>★</button>
          <button onClick={()=>onEdit(item)} style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:7,padding:"4px 9px",color:"#60a5fa",fontSize:11,fontWeight:600}}>Edytuj</button>
          <button onClick={()=>onDelete(item.id)} style={{background:"none",border:"none",color:"#334155",fontSize:18,lineHeight:1,padding:"2px 4px"}}>×</button>
        </div>
      </div>

      {/* Termin */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"#0d1117",borderRadius:8,padding:"8px 12px",marginBottom:item.notes?10:10}}>
        <div>
          <div style={{fontSize:9,color:"#334155",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>
            {item.status==="done"?"Zakończono":"Termin"}
          </div>
          <div style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:12,color:"#94a3b8"}}>
            {item.status==="done" ? fmtDate(item.completedDate||item.deadline) : fmtDate(item.deadline)}
          </div>
        </div>
        <div style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:16,fontWeight:700,color:tl.color}}>
          {item.status==="done" ? "✓" : days<0 ? `⚠ +${Math.abs(days)}d` : days===0 ? "Dziś!" : `${days}d`}
        </div>
      </div>

      {/* Notatki */}
      {item.notes && <div style={{fontSize:12,color:"#64748b",lineHeight:1.55,borderLeft:"2px solid #1e2d3d",paddingLeft:10,marginBottom:10}}>{item.notes}</div>}

      {/* Status */}
      <div style={{display:"flex",gap:5}}>
        {[["pending","Do zrobienia"],["inprogress","W toku"],["done","Zakończone"]].map(([v,l])=>(
          <button key={v} onClick={()=>{ if(v==="done") onComplete(item.id); else onStatus(item.id,v); }} style={{flex:1,padding:"5px 0",borderRadius:7,fontSize:10,fontWeight:700,fontFamily:"'IBM Plex Mono', monospace",background:item.status===v?"#1e3a5f":"transparent",border:`1px solid ${item.status===v?"#2563eb":"#1e2d3d"}`,color:item.status===v?"#60a5fa":"#334155"}}>{l}</button>
        ))}
      </div>
    </div>
  );
};

const Modal = ({open,onClose,title,children}) => {
  if(!open) return null;
  return (
    <div style={{position:"fixed",inset:0,zIndex:100,background:"rgba(0,0,0,0.88)",backdropFilter:"blur(8px)",display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}>
      <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:"20px 20px 0 0",width:"100%",maxWidth:520,padding:"24px 20px 48px",maxHeight:"92vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:22}}>
          <span style={{fontWeight:700,fontSize:17}}>{title}</span>
          <button onClick={onClose} style={{background:"#1e2d3d",border:"none",borderRadius:8,padding:"6px 11px",color:"#94a3b8",fontSize:17}}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
};

const Field = ({label,children}) => (
  <div style={{marginBottom:14}}>
    <div style={{fontSize:10,fontWeight:700,color:"#475569",marginBottom:6,textTransform:"uppercase",letterSpacing:"0.1em"}}>{label}</div>
    {children}
  </div>
);

// ── GŁÓWNA APLIKACJA ─────────────────────────────────────────────
const EMPTY = { type:"adhoc", title:"", regulator:"Wewnętrzne", source:"", deadline:addDays(7), recurrence:"none", status:"pending", starred:false, completedDate:null, notes:"" };

export default function ComplianceTracker() {
  const [tab, setTab]           = useState("tasks"); // "tasks" | "goals"
  const [items, setItems]       = useState(()=>{ try{const s=localStorage.getItem("ct_v2");return s?JSON.parse(s):INITIAL;}catch{return INITIAL;}});
  const [goals, setGoals]       = useState(()=>{ try{const s=localStorage.getItem("ct_goals");return s?JSON.parse(s):INITIAL_GOALS;}catch{return INITIAL_GOALS;}});
  const [modal, setModal]       = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm]         = useState(EMPTY);
  const [filterType, setFT]     = useState("all");
  const [filterSt, setFS]       = useState("active");
  const [showExport, setShowEx]   = useState(false);
  const [reportPeriod, setRP]     = useState(`${thisYear}-h1`);
  const [showPreview, setPreview] = useState(false);
  const [completedDateModal, setCDM] = useState(null);

  const save = (next) => { setItems(next); try{localStorage.setItem("ct_v2",JSON.stringify(next));}catch{} };

  const openAdd  = () => { setEditItem(null); setForm(EMPTY); setModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({...item}); setModal(true); };

  const saveForm = () => {
    if(!form.title||!form.deadline) return;
    save(editItem ? items.map(x=>x.id===editItem.id?{...x,...form}:x) : [...items,{id:Date.now(),...form}]);
    setModal(false);
  };

  const del      = (id) => save(items.filter(x=>x.id!==id));
  const onStatus = (id, s) => save(items.map(x=>x.id===id?{...x,status:s}:x));
  const onStar   = (id) => save(items.map(x=>x.id===id?{...x,starred:!x.starred}:x));

  // Przy oznaczaniu jako "done" — pyta o datę zakończenia
  const onComplete = (id) => {
    const item = items.find(x=>x.id===id);
    if(!item) return;
    if(item.status==="done") { onStatus(id,"pending"); return; }
    setCDM(id);
  };

  const confirmComplete = (date) => {
    save(items.map(x=>x.id===completedDateModal?{...x,status:"done",completedDate:date}:x));
    setCDM(null);
  };

  const filtered = useMemo(()=>{
    let list=[...items];
    if(filterType!=="all") list=list.filter(x=>x.type===filterType);
    if(filterSt==="active") list=list.filter(x=>x.status!=="done");
    if(filterSt==="done")   list=list.filter(x=>x.status==="done");
    if(filterSt==="starred") list=list.filter(x=>x.starred);
    list.sort((a,b)=>{
      if(a.status==="done"&&b.status!=="done") return 1;
      if(b.status==="done"&&a.status!=="done") return -1;
      return new Date(a.deadline)-new Date(b.deadline);
    });
    return list;
  },[items,filterType,filterSt]);

  const overdue  = items.filter(x=>x.status!=="done"&&getDays(x.deadline)<0).length;
  const urgent   = items.filter(x=>x.status!=="done"&&getDays(x.deadline)>=0&&getDays(x.deadline)<=7).length;
  const upcoming = items.filter(x=>x.status!=="done"&&getDays(x.deadline)>7&&getDays(x.deadline)<=21).length;
  const starred  = items.filter(x=>x.starred).length;

  const overdueList = filtered.filter(x=>x.status!=="done"&&getDays(x.deadline)<0);
  const activeList  = filtered.filter(x=>x.status!=="done"&&getDays(x.deadline)>=0);
  const doneList    = filtered.filter(x=>x.status==="done");

  return (
    <div style={{fontFamily:"'IBM Plex Sans', sans-serif",background:"#07090f",color:"#e2e8f0",minHeight:"100vh",maxWidth:520,margin:"0 auto"}}>
      <GlobalStyles/>

      {/* NAGŁÓWEK */}
      <div style={{position:"sticky",top:0,zIndex:50,background:"linear-gradient(180deg,#07090f 78%,transparent)",padding:"18px 16px 12px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:34,height:34,borderRadius:10,background:"linear-gradient(135deg,#1e3a5f,#312e81)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,border:"1px solid #2563eb44"}}>⚖️</div>
            <div>
              <div style={{fontWeight:700,fontSize:16}}>Compliance Tracker</div>
              <div style={{fontSize:10,color:"#334155",fontFamily:"'IBM Plex Mono', monospace",letterSpacing:"0.05em"}}>AML · PERSONAL</div>
            </div>
          </div>
          {tab === "tasks" && <button onClick={openAdd} style={{background:"linear-gradient(135deg,#1e3a5f,#312e81)",border:"1px solid #2563eb55",borderRadius:10,padding:"9px 14px",color:"#93c5fd",fontWeight:700,fontSize:13}}>+ Dodaj</button>}
        </div>

        {/* Zakładki */}
        <div style={{display:"flex",gap:5,marginBottom:14}}>
          {[["tasks","📋 Zadania"],["goals","🎯 Cele roczne"]].map(([v,l])=>(
            <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"8px 0",borderRadius:10,fontSize:12,fontWeight:700,background:tab===v?"linear-gradient(135deg,#1e3a5f,#312e81)":"#0d1117",border:`1px solid ${tab===v?"#2563eb55":"#1e2d3d"}`,color:tab===v?"#93c5fd":"#475569"}}>{l}</button>
          ))}
        </div>

        {/* Statystyki i filtry — tylko w zakładce zadań */}
        {tab === "tasks" && (
          <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:14}}>
          {[
            {val:overdue, label:"Po term.", color:"#ef4444"},
            {val:urgent,  label:"Pilne",    color:"#ef4444"},
            {val:upcoming,label:"Wkrótce",  color:"#f59e0b"},
            {val:starred, label:"★ Ważne",  color:"#f59e0b"},
          ].map(({val,label,color})=>(
            <div key={label} style={{background:"#0d1117",border:`1px solid ${color}33`,borderRadius:10,padding:"8px 6px",textAlign:"center"}}>
              <div style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:22,fontWeight:600,color,lineHeight:1}}>{val}</div>
              <div style={{fontSize:9,color:"#475569",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginTop:3}}>{label}</div>
            </div>
          ))}
        </div>

        {/* Filtry typ */}
        <div style={{display:"flex",gap:5,marginBottom:8,overflowX:"auto",scrollbarWidth:"none"}}>
          {[["all","Wszystkie"],["regulatory","📋 Regulacyjne"],["internal","🏦 Wewnętrzne"],["adhoc","⚡ Ad hoc"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFT(v)} style={{padding:"4px 10px",borderRadius:7,fontSize:11,fontWeight:700,whiteSpace:"nowrap",flexShrink:0,background:filterType===v?"#1e3a5f":"transparent",border:`1px solid ${filterType===v?"#2563eb":"#1e2d3d"}`,color:filterType===v?"#60a5fa":"#334155"}}>{l}</button>
          ))}
        </div>

        {/* Filtry status */}
        <div style={{display:"flex",gap:5,alignItems:"center"}}>
          {[["active","Aktywne"],["starred","★ Ważne"],["all","Wszystkie"],["done","Zakończone"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFS(v)} style={{padding:"4px 10px",borderRadius:7,fontSize:10,fontWeight:700,fontFamily:"'IBM Plex Mono', monospace",whiteSpace:"nowrap",background:filterSt===v?"#1e3a5f":"transparent",border:`1px solid ${filterSt===v?"#2563eb":"#1e2d3d"}`,color:filterSt===v?"#60a5fa":"#334155"}}>{l}</button>
          ))}
          <span style={{marginLeft:"auto",fontSize:10,color:"#334155",fontFamily:"'IBM Plex Mono', monospace"}}>{filtered.length}</span>
        </div>
          </div>
        )}
      </div>

      {/* ZAWARTOŚĆ ZAKŁADEK */}
      {tab === "goals"
        ? <GoalsView goals={goals} setGoals={setGoals} />
        : (
      <div style={{padding:"8px 16px 160px",display:"flex",flexDirection:"column",gap:10}}>
        {overdueList.length>0&&<><div style={{fontSize:10,color:"#ef4444",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.12em",paddingTop:4}}>⚠ Przeterminowane</div>{overdueList.map(item=><Card key={item.id} item={item} onEdit={openEdit} onDelete={del} onStatus={onStatus} onStar={onStar} onComplete={onComplete}/>)}</>}
        {activeList.length>0&&<><div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.12em",paddingTop:4}}>Nadchodzące</div>{activeList.map(item=><Card key={item.id} item={item} onEdit={openEdit} onDelete={del} onStatus={onStatus} onStar={onStar} onComplete={onComplete}/>)}</>}
        {filterSt!=="active"&&doneList.length>0&&<><div style={{fontSize:10,color:"#10b981",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.12em",paddingTop:4}}>✓ Zakończone</div>{doneList.map(item=><Card key={item.id} item={item} onEdit={openEdit} onDelete={del} onStatus={onStatus} onStar={onStar} onComplete={onComplete}/>)}</>}
        {filtered.length===0&&<div style={{textAlign:"center",padding:"56px 0",color:"#1e2d3d"}}><div style={{fontSize:36,marginBottom:12}}>✓</div><div style={{fontSize:13}}>Brak zadań</div></div>}
      </div>
        )
      }

      {/* DOLNY PASEK */}
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:520,background:"linear-gradient(180deg,transparent,#07090f 45%)",padding:"16px 16px 14px"}}>
        {showExport&&(
          <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:14,padding:"14px 16px",marginBottom:10,display:"flex",flexDirection:"column",gap:10}}>
            {/* Eksport Excel */}
            <button onClick={()=>{exportExcel(items);setShowEx(false);}} style={{background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:10,padding:"11px 0",color:"#34d399",fontWeight:700,fontSize:13}}>⬇ Eksport Excel (.xlsx)</button>
            {/* Wybór okresu */}
            <div>
              <div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8}}>Raport — wybierz okres</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:10}}>
                {REPORT_PERIODS.map(p=>(
                  <button key={p.id} onClick={()=>setRP(p.id)} style={{padding:"8px 0",borderRadius:8,fontSize:11,fontWeight:700,fontFamily:"'IBM Plex Mono', monospace",background:reportPeriod===p.id?"#1e3a5f":"transparent",border:`1px solid ${reportPeriod===p.id?"#2563eb":"#1e2d3d"}`,color:reportPeriod===p.id?"#60a5fa":"#475569"}}>
                    {p.label}
                    <div style={{fontSize:9,color:reportPeriod===p.id?"#3b82f6":"#334155",marginTop:1}}>{p.short}</div>
                  </button>
                ))}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>{setShowEx(false);setPreview(true);}} style={{flex:1,background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:10,padding:"10px 0",color:"#94a3b8",fontWeight:700,fontSize:12}}>👁 Podgląd</button>
                <button onClick={()=>{generatePrintReport(items,reportPeriod);setShowEx(false);}} style={{flex:2,background:"linear-gradient(135deg,#1e3a5f,#312e81)",border:"1px solid #2563eb55",borderRadius:10,padding:"10px 0",color:"#93c5fd",fontWeight:700,fontSize:12}}>📄 Drukuj / PDF</button>
              </div>
            </div>
          </div>
        )}

        <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:12,padding:"10px 14px",display:"flex",gap:8,alignItems:"center"}}>
          <div style={{display:"flex",gap:10,flex:1}}>
            {[["#ef4444","≤7d"],["#f59e0b","8–21d"],["#10b981","OK"]].map(([c,l])=>(
              <div key={l} style={{display:"flex",alignItems:"center",gap:5}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:c,boxShadow:`0 0 5px ${c}`}}/>
                <span style={{fontSize:10,color:"#475569",fontWeight:600}}>{l}</span>
              </div>
            ))}
          </div>
          <button onClick={()=>setShowEx(s=>!s)} style={{background:showExport?"#1e3a5f":"#060b14",border:`1px solid ${showExport?"#2563eb":"#1e2d3d"}`,borderRadius:9,padding:"7px 12px",color:showExport?"#60a5fa":"#475569",fontWeight:700,fontSize:12}}>📤 Eksport</button>
        </div>
      </div>

      {/* PODGLĄD RAPORTU */}
      {showPreview && (
        <ReportPreview
          items={items}
          periodId={reportPeriod}
          onClose={()=>setPreview(false)}
          onPrint={()=>{ generatePrintReport(items,reportPeriod); }}
        />
      )}

      {/* MODAL — Data zakończenia */}
      {completedDateModal&&(
        <div style={{position:"fixed",inset:0,zIndex:200,background:"rgba(0,0,0,0.9)",backdropFilter:"blur(8px)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:"#0d1117",border:"1px solid #1e2d3d",borderRadius:16,padding:"24px 20px",width:"100%",maxWidth:360}}>
            <div style={{fontWeight:700,fontSize:16,marginBottom:6}}>Data zakończenia</div>
            <div style={{fontSize:12,color:"#475569",marginBottom:16}}>Kiedy zostało zakończone to zadanie?</div>
            <input type="date" defaultValue={new Date().toISOString().split("T")[0]} id="completeDate" style={{...inputSt,marginBottom:14}}/>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setCDM(null)} style={{flex:1,background:"transparent",border:"1px solid #1e2d3d",borderRadius:10,padding:"11px 0",color:"#475569",fontWeight:700,fontSize:13}}>Anuluj</button>
              <button onClick={()=>confirmComplete(document.getElementById("completeDate").value)} style={{flex:2,background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:10,padding:"11px 0",color:"#34d399",fontWeight:700,fontSize:13}}>✓ Zakończ</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL — Dodaj / Edytuj */}
      <Modal open={modal} onClose={()=>setModal(false)} title={editItem?"Edytuj zadanie":"Nowe zadanie"}>
        <Field label="Typ">
          <div style={{display:"flex",gap:7}}>
            {TASK_TYPES.map(t=>(
              <button key={t.id} onClick={()=>setForm(f=>({...f,type:t.id}))} style={{flex:1,padding:"9px 0",borderRadius:9,background:form.type===t.id?t.color+"22":"transparent",border:`1px solid ${form.type===t.id?t.color:"#1e2d3d"}`,color:form.type===t.id?t.color:"#475569",fontWeight:700,fontSize:12}}>{t.icon} {t.label}</button>
            ))}
          </div>
        </Field>

        <Field label="Tytuł / opis">
          <input value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))} placeholder="np. Raport kwartalny AML do UKNF" style={inputSt}/>
        </Field>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <Field label="Regulator">
            <select value={form.regulator} onChange={e=>setForm(f=>({...f,regulator:e.target.value}))} style={inputSt}>
              {REGULATORS.map(r=><option key={r} value={r}>{r}</option>)}
            </select>
          </Field>
          <Field label="Termin">
            <input type="date" value={form.deadline} onChange={e=>setForm(f=>({...f,deadline:e.target.value}))} style={inputSt}/>
          </Field>
        </div>

        <Field label="Podstawa prawna">
          <input value={form.source} onChange={e=>setForm(f=>({...f,source:e.target.value}))} placeholder="np. Art. 72 Ustawa AML / EBA/GL/2022/05" style={inputSt}/>
        </Field>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <Field label="Powtarzalność">
            <select value={form.recurrence} onChange={e=>setForm(f=>({...f,recurrence:e.target.value}))} style={inputSt}>
              {RECURRENCE.map(r=><option key={r.id} value={r.id}>{r.label}</option>)}
            </select>
          </Field>
          <Field label="Status">
            <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={inputSt}>
              <option value="pending">Do zrobienia</option>
              <option value="inprogress">W toku</option>
              <option value="done">Zakończone</option>
            </select>
          </Field>
        </div>

        {/* Oznacz jako ważne */}
        <div style={{marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between",background:"#060b14",border:`1px solid ${form.starred?"#f59e0b44":"#1e2d3d"}`,borderRadius:10,padding:"12px 14px"}}>
          <div>
            <div style={{fontSize:13,fontWeight:600,color:form.starred?"#f59e0b":"#e2e8f0"}}>★ Ważne — uwzględnij w raporcie rocznym</div>
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>Pojawi się w sekcji "Kluczowe projekty"</div>
          </div>
          <button onClick={()=>setForm(f=>({...f,starred:!f.starred}))} style={{width:44,height:24,borderRadius:12,border:"none",background:form.starred?"#f59e0b":"#1a2744",position:"relative",transition:"background 0.2s"}}>
            <div style={{width:18,height:18,borderRadius:9,background:"white",position:"absolute",top:3,left:form.starred?23:3,transition:"left 0.2s"}}/>
          </button>
        </div>

        <Field label="Notatki / zakres działania">
          <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder="Zakres, wymagane dokumenty, kontekst..." rows={3} style={{...inputSt,resize:"vertical",minHeight:80}}/>
        </Field>

        <button onClick={saveForm} style={{width:"100%",background:"linear-gradient(135deg,#1e3a5f,#312e81)",border:"1px solid #2563eb55",borderRadius:12,padding:"14px 0",color:"#93c5fd",fontWeight:700,fontSize:15,marginTop:4}}>{editItem?"Zapisz zmiany":"Dodaj zadanie"}</button>
      </Modal>
    </div>
  );
}
