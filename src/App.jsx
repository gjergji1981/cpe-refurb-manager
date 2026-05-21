import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ─────────────────────────────────────────────
// CONSTANTS & HELPERS
// ─────────────────────────────────────────────
const DEVICE_TYPES = ["Router/Modem", "Set-top Box", "ONT/OLT"];
const STAGES = ["Intake","Triage","Refurbishment","QC Check","Stock","Scrap","Escalated"];
const TYPE_ICON = { "Router/Modem":"⬡", "Set-top Box":"▦", "ONT/OLT":"◈" };
const STAGE_COLOR = {
  Intake:        { bg:"#EEF2FF", text:"#3730A3", dot:"#6366F1" },
  Triage:        { bg:"#FFF7ED", text:"#9A3412", dot:"#F97316" },
  Refurbishment: { bg:"#FFFBEB", text:"#92400E", dot:"#F59E0B" },
  "QC Check":    { bg:"#EFF6FF", text:"#1E40AF", dot:"#3B82F6" },
  Stock:         { bg:"#F0FDF4", text:"#166534", dot:"#22C55E" },
  Scrap:         { bg:"#FEF2F2", text:"#991B1B", dot:"#EF4444" },
  Escalated:     { bg:"#FDF4FF", text:"#7E22CE", dot:"#A855F7" },
};

const OUI = {
  "E8:65:D4":"Huawei","54:89:98":"Huawei","70:72:CF":"Huawei","00:9A:CD":"Huawei",
  "04:C0:6F":"Huawei","28:31:52":"Huawei","48:46:FB":"Huawei","AC:44:F2":"Huawei",
  "20:F3:A3":"Huawei","D4:6A:A8":"Huawei","F4:9F:F3":"Huawei","B4:15:13":"Huawei",
  "00:1A:2B":"Cisco","00:1B:2F":"Cisco","00:21:A0":"Cisco","00:26:99":"Cisco",
  "00:19:CB":"ZTE","00:26:ED":"ZTE","34:4B:50":"ZTE","8C:A6:DF":"ZTE",
  "BC:F6:12":"ZTE","C8:6C:87":"ZTE","E4:63:DA":"ZTE","FC:AF:6A":"ZTE",
  "00:24:E8":"Sagemcom","7C:4C:A5":"Sagemcom","C8:D7:19":"Sagemcom","E0:AB:39":"Sagemcom",
  "44:E9:DD":"Sagemcom","78:32:1B":"Sagemcom","48:EE:0C":"Sagemcom",
  "00:1A:C1":"Technicolor","00:24:D4":"Technicolor","28:C6:8E":"Technicolor",
  "50:E5:49":"Technicolor","78:E4:00":"Technicolor","A0:21:95":"Technicolor",
  "C0:25:2F":"Technicolor","DC:A6:32":"Technicolor",
  "00:17:10":"Netgear","00:1E:2A":"Netgear","20:E5:2A":"Netgear","2C:B0:5D":"Netgear",
  "84:1B:5E":"Netgear","A0:04:60":"Netgear","C4:04:15":"Netgear",
  "00:1D:7E":"Arris","00:21:5C":"Arris","00:26:B8":"Arris","18:1B:EB":"Arris",
  "3C:DF:A9":"Arris","70:2A:D5":"Arris","AC:20:2D":"Arris","E0:18:54":"Arris",
  "54:A7:03":"Nokia","9C:97:26":"Nokia","00:25:9C":"Nokia","D4:CA:6D":"Nokia",
  "30:46:9A":"TP-Link","50:C7:BF":"TP-Link","98:DE:D0":"TP-Link","E8:48:B8":"TP-Link",
  "00:27:22":"TP-Link","14:CC:20":"TP-Link","B0:BE:76":"TP-Link","EC:08:6B":"TP-Link",
};

function lookupOUI(mac) {
  if (!mac) return null;
  const c = mac.toUpperCase().replace(/[^0-9A-F]/g,"");
  if (c.length < 6) return null;
  return OUI[`${c.slice(0,2)}:${c.slice(2,4)}:${c.slice(4,6)}`] || null;
}
function formatMac(raw) {
  const c = raw.replace(/[^0-9a-fA-F]/g,"").slice(0,12);
  return c.match(/.{1,2}/g)?.join(":").toUpperCase() || c;
}
let _id = 13;
function genId() { return `D-${String(_id++).padStart(4,"0")}`; }
function today() { return new Date().toISOString().slice(0,10); }
function useIsMobile() {
  const [mobile, setMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const h = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return mobile;
}

// ─────────────────────────────────────────────
// SEED DATA
// ─────────────────────────────────────────────
const SEED = [
  { id:"D-0001", serial:"SN-88421", mac:"E8:65:D4:11:22:33", model:"Huawei HG8245H",        type:"Router/Modem", stage:"Stock",         outcome:"Working",     received:"2025-05-01", notes:"Firmware reset",      sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0002", serial:"SN-33190", mac:"28:C6:8E:44:55:66", model:"Technicolor TC7200",     type:"Set-top Box",  stage:"Stock",         outcome:"Working",     received:"2025-05-02", notes:"HDMI port replaced",  sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0003", serial:"SN-77042", mac:"9C:97:26:77:88:99", model:"Nokia G-010G-P",         type:"ONT/OLT",      stage:"Stock",         outcome:"Working",     received:"2025-05-03", notes:"Clean & reconfigure", sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0004", serial:"SN-55610", mac:"34:4B:50:AA:BB:CC", model:"ZTE ZXHN H108N",         type:"Router/Modem", stage:"Scrap",         outcome:"Scrap",       received:"2025-05-04", notes:"PCB damage",          sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0005", serial:"SN-12983", mac:"",                  model:"",                       type:"Set-top Box",  stage:"Scrap",         outcome:"Scrap",       received:"2025-05-05", notes:"Burned PSU",          sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0006", serial:"SN-66101", mac:"54:89:98:DD:EE:FF", model:"Huawei EchoLife EG8145", type:"ONT/OLT",      stage:"Refurbishment", outcome:null,          received:"2025-05-06", notes:"In repair",           sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0007", serial:"SN-29344", mac:"C8:D7:19:12:34:56", model:"Sagemcom F@ST 5366",     type:"Router/Modem", stage:"Refurbishment", outcome:null,          received:"2025-05-07", notes:"Awaiting partner",    sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0008", serial:"SN-84720", mac:"",                  model:"",                       type:"Set-top Box",  stage:"Triage",        outcome:null,          received:"2025-05-08", notes:"",                    sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0009", serial:"SN-39011", mac:"BC:F6:12:AB:CD:EF", model:"ZTE F660",               type:"ONT/OLT",      stage:"Escalated",     outcome:"Not Working", received:"2025-05-09", notes:"Failed QC twice",     sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0010", serial:"SN-47832", mac:"00:1A:2B:33:44:55", model:"Cisco DPC3825",          type:"Router/Modem", stage:"Stock",         outcome:"Working",     received:"2025-05-10", notes:"",                    sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0011", serial:"SN-90012", mac:"7C:4C:A5:66:77:88", model:"Sagemcom FAST 3686",     type:"Set-top Box",  stage:"Refurbishment", outcome:null,          received:"2025-05-11", notes:"Screen replaced",     sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0012", serial:"SN-55123", mac:"",                  model:"",                       type:"ONT/OLT",      stage:"Triage",        outcome:null,          received:"2025-05-12", notes:"",                    sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
];

// ─────────────────────────────────────────────
// DESIGN TOKENS
// ─────────────────────────────────────────────
const C = {
  indigo:"#6366F1", indigoDark:"#4338CA", indigoLight:"#EEF2FF",
  purple:"#7C3AED", purpleLight:"#FDF4FF",
  amber:"#F59E0B",  amberLight:"#FFFBEB",
  green:"#22C55E",  greenLight:"#F0FDF4",  greenDark:"#166534",
  red:"#EF4444",    redLight:"#FEF2F2",    redDark:"#991B1B",
  slate:"#0F172A",  slate2:"#1E293B",      slate3:"#475569",
  slate4:"#94A3B8", slate5:"#CBD5E1",      slate6:"#E2E8F0",
  slate7:"#F1F5F9", slate8:"#F8FAFC",
  white:"#fff",
};

// ─────────────────────────────────────────────
// SHARED COMPONENTS
// ─────────────────────────────────────────────
function Badge({ stage }) {
  const c = STAGE_COLOR[stage] || { bg:C.slate7, text:C.slate3, dot:C.slate4 };
  return (
    <span style={{ background:c.bg, color:c.text, fontSize:11, fontWeight:700,
      padding:"3px 9px", borderRadius:20, display:"inline-flex", alignItems:"center", gap:4, whiteSpace:"nowrap" }}>
      <span style={{ width:6, height:6, borderRadius:"50%", background:c.dot, flexShrink:0 }}/>
      {stage}
    </span>
  );
}

function Btn({ children, onClick, variant="default", size="md", full=false, disabled=false, style:overrideStyle={} }) {
  const v = {
    default: { bg:C.slate8,     color:C.slate3,   border:`1px solid ${C.slate6}` },
    primary: { bg:C.indigo,     color:C.white,     border:"none" },
    success: { bg:C.greenLight, color:C.greenDark, border:`1px solid #BBF7D0` },
    danger:  { bg:C.redLight,   color:C.redDark,   border:`1px solid #FECACA` },
    purple:  { bg:C.purpleLight,color:C.purple,    border:`1px solid #E9D5FF` },
    amber:   { bg:C.amberLight, color:"#92400E",   border:`1px solid #FCD34D` },
    ghost:   { bg:"transparent",color:C.slate3,    border:`1px solid ${C.slate6}` },
    dark:    { bg:C.slate,      color:C.white,     border:"none" },
  }[variant] || {};
  const pad = size==="sm" ? "5px 12px" : size==="lg" ? "12px 24px" : "8px 16px";
  const fs  = size==="sm" ? 12 : size==="lg" ? 15 : 13;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ ...v, padding:pad, borderRadius:9, fontSize:fs, fontWeight:700,
        cursor:disabled?"not-allowed":"pointer", opacity:disabled?.5:1,
        whiteSpace:"nowrap", width:full?"100%":"auto",
        display:"inline-flex", alignItems:"center", justifyContent:"center", gap:6, transition:"opacity .15s", ...overrideStyle }}>
      {children}
    </button>
  );
}

function Card({ children, style={}, accent }) {
  return (
    <div style={{ background:C.white, border:`1.5px solid ${C.slate6}`, borderRadius:14,
      borderLeft: accent ? `4px solid ${accent}` : undefined,
      padding:16, ...style }}>
      {children}
    </div>
  );
}

function SectionTitle({ children }) {
  return <h3 style={{ margin:"0 0 14px", fontSize:14, fontWeight:700, color:C.slate }}>{children}</h3>;
}

function Label({ children }) {
  return <label style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em", display:"block", marginBottom:4 }}>{children}</label>;
}

const iStyle = (extra={}) => ({
  width:"100%", padding:"9px 12px", borderRadius:9, border:`1.5px solid ${C.slate6}`,
  fontSize:13, outline:"none", boxSizing:"border-box", background:C.white,
  WebkitAppearance:"none", ...extra
});

function StatCard({ label, value, sub, accent, badge }) {
  return (
    <Card accent={accent} style={{ display:"flex", flexDirection:"column", gap:4 }}>
      <span style={{ fontSize:11, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".07em" }}>{label}</span>
      <div style={{ display:"flex", alignItems:"baseline", gap:8 }}>
        <span style={{ fontSize:28, fontWeight:800, color:C.slate, lineHeight:1 }}>{value}</span>
        {badge && <span style={{ fontSize:11, fontWeight:700, padding:"2px 7px", borderRadius:20,
          background:badge.bg, color:badge.color }}>{badge.text}</span>}
      </div>
      {sub && <span style={{ fontSize:12, color:C.slate4 }}>{sub}</span>}
    </Card>
  );
}

// Mobile-friendly device card (replaces table rows on mobile)
function DeviceCard({ d, actions, fields }) {
  return (
    <div style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, background:C.white, marginBottom:10 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8 }}>
        <div>
          <span style={{ fontWeight:800, color:C.indigo, fontSize:13 }}>{d.id}</span>
          <span style={{ fontFamily:"monospace", fontSize:12, color:C.slate2, marginLeft:8 }}>{d.serial}</span>
        </div>
        <Badge stage={d.stage}/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:10 }}>
        {fields.map(([k,v])=>(
          <div key={k}>
            <span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".05em" }}>{k}</span>
            <div style={{ fontSize:12, color:C.slate2, marginTop:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{v||"—"}</div>
          </div>
        ))}
      </div>
      {actions && <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>{actions}</div>}
    </div>
  );
}

// Alert banner
function Alert({ children, type="info" }) {
  const colors = {
    info:    { bg:"#EEF2FF", border:`1.5px solid #C7D2FE`, color:"#3730A3" },
    warning: { bg:"#FFFBEB", border:`1.5px solid #FCD34D`, color:"#92400E" },
    success: { bg:C.greenLight, border:`1.5px solid #BBF7D0`, color:C.greenDark },
    danger:  { bg:C.redLight,   border:`1.5px solid #FECACA`,  color:C.redDark },
  }[type];
  return (
    <div style={{ ...colors, borderRadius:10, padding:"10px 14px", fontSize:13, display:"flex", alignItems:"center", gap:8 }}>
      {children}
    </div>
  );
}

// Tab switcher
function Tabs({ tabs, active, onChange }) {
  return (
    <div style={{ display:"flex", background:C.slate7, borderRadius:10, padding:3, gap:2 }}>
      {tabs.map(t=>(
        <button key={t.id} onClick={()=>onChange(t.id)}
          style={{ flex:1, padding:"7px 12px", border:"none", borderRadius:8, fontSize:12, fontWeight:700,
            cursor:"pointer", background:active===t.id?C.white:"transparent",
            color:active===t.id?C.slate:C.slate3, transition:"all .15s",
            boxShadow:active===t.id?"0 1px 3px rgba(0,0,0,.08)":undefined, whiteSpace:"nowrap" }}>
          {t.label}
        </button>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────
function Dashboard({ devices, isMobile }) {
  const total       = devices.length;
  const inFlow      = devices.filter(d=>!["Stock","Scrap"].includes(d.stage)).length;
  const stock       = devices.filter(d=>d.stage==="Stock").length;
  const scrap       = devices.filter(d=>d.stage==="Scrap").length;
  const atPartner   = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome).length;
  const pendingConf = devices.filter(d=>d.partnerOutcome && d.stage==="Refurbishment").length;
  const byStage     = STAGES.map(s=>({ stage:s, count:devices.filter(d=>d.stage===s).length }));
  const byType      = DEVICE_TYPES.map(t=>({ type:t, total:devices.filter(d=>d.type===t).length, stock:devices.filter(d=>d.type===t&&d.stage==="Stock").length, scrap:devices.filter(d=>d.type===t&&d.stage==="Scrap").length }));
  const maxStage    = Math.max(...byStage.map(s=>s.count), 1);
  const done        = devices.filter(d=>d.outcome==="Working"||d.outcome==="Not Working");
  const working     = done.filter(d=>d.outcome==="Working").length;
  const rate        = done.length ? Math.round(working/done.length*100) : 0;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Dashboard</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Live pipeline overview</p>
      </div>

      {pendingConf>0 && <Alert type="warning">⚠ <strong>{pendingConf}</strong> partner outcome{pendingConf>1?"s":""} awaiting your confirmation in Refurbishment</Alert>}

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(5,1fr)", gap:10 }}>
        <StatCard label="Total" value={total} sub="All devices" accent={C.indigo}/>
        <StatCard label="Pipeline" value={inFlow} sub="In progress" accent={C.amber}/>
        <StatCard label="Stock" value={stock} sub="Ready" accent={C.green}/>
        <StatCard label="Scrapped" value={scrap} sub="Disposed" accent={C.red}/>
        <StatCard label="At Partner" value={atPartner}
          badge={pendingConf?{text:`${pendingConf} pending`,bg:"#FFFBEB",color:"#92400E"}:undefined}
          sub="Awaiting results" accent={C.purple}/>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:16 }}>
        <Card>
          <SectionTitle>Devices by stage</SectionTitle>
          {byStage.map(({ stage, count })=>{
            const sc = STAGE_COLOR[stage] || { dot:C.slate4 };
            return (
              <div key={stage} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
                <span style={{ width:isMobile?80:96, fontSize:12, fontWeight:600, color:C.slate3, textAlign:"right", flexShrink:0 }}>{stage}</span>
                <div style={{ flex:1, background:C.slate7, borderRadius:6, height:20, overflow:"hidden" }}>
                  <div style={{ width:`${(count/maxStage)*100}%`, minWidth:count?28:0, height:"100%",
                    background:sc.dot, borderRadius:6, display:"flex", alignItems:"center", paddingLeft:6, transition:"width .4s" }}>
                    {count>0 && <span style={{ fontSize:11, fontWeight:800, color:C.white }}>{count}</span>}
                  </div>
                </div>
                {count===0 && <span style={{ fontSize:11, color:C.slate5, width:14 }}>0</span>}
              </div>
            );
          })}
        </Card>

        <Card>
          <SectionTitle>By device type</SectionTitle>
          {byType.map(r=>(
            <div key={r.type} style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
              padding:"10px 0", borderBottom:`1px solid ${C.slate7}` }}>
              <span style={{ fontSize:13, fontWeight:600, color:C.slate }}>{TYPE_ICON[r.type]} {r.type}</span>
              <div style={{ display:"flex", gap:16 }}>
                <span style={{ fontSize:13, fontWeight:700, color:C.slate }}>{r.total}</span>
                <span style={{ fontSize:13, fontWeight:700, color:C.greenDark }}>↑{r.stock}</span>
                <span style={{ fontSize:13, fontWeight:700, color:C.redDark }}>↓{r.scrap}</span>
              </div>
            </div>
          ))}
          <div style={{ marginTop:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
              <span style={{ fontSize:12, color:C.slate3 }}>Refurb success rate</span>
              <span style={{ fontSize:12, fontWeight:800, color:C.greenDark }}>{rate}%</span>
            </div>
            <div style={{ background:C.slate7, borderRadius:8, height:10, overflow:"hidden" }}>
              <div style={{ width:`${rate}%`, height:"100%", background:C.green, borderRadius:8, transition:"width .4s" }}/>
            </div>
            <div style={{ fontSize:11, color:C.slate4, marginTop:4 }}>{working} of {done.length} tested units passed</div>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// INTAKE & TRIAGE
// ─────────────────────────────────────────────
function IntakeTriage({ devices, setDevices, isMobile }) {
  const ACTIONS = ["Refurbishment", "Scrap"];

  const [tab, setTab]               = useState("manual");
  const [form, setForm]             = useState({ serial:"", type:DEVICE_TYPES[0], mac:"", model:"", notes:"", action:"Refurbishment" });
  const [ok, setOk]                 = useState(null);
  const [formError, setFormError]   = useState("");
  const [dragOver, setDragOver]     = useState(false);
  const [processing, setProcessing] = useState(false);
  const [uploadResult, setUploadResult] = useState(null);
  const [uploadLogs, setUploadLogs]     = useState([]);
  const [editingId, setEditingId]   = useState(null);
  const [editBuf, setEditBuf]       = useState({});
  const [deleteId, setDeleteId]     = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [showConfirm, setShowConfirm] = useState(false);

  // Queue = devices in Intake/Triage that have a pendingAction set
  const queue = devices.filter(d => ["Intake","Triage"].includes(d.stage));

  // Counts for confirmation popup
  const toRefurb = queue.filter(d => d.pendingAction === "Refurbishment").length;
  const toScrap  = queue.filter(d => d.pendingAction === "Scrap").length;
  const unassigned = queue.filter(d => !d.pendingAction).length;
  const readyCount = toRefurb + toScrap;

  // ── MAC / OUI helper ──
  function handleMacChange(raw, setBuf) {
    const mac = formatMac(raw);
    const mfr = lookupOUI(mac);
    setBuf(f => ({ ...f, mac, model: mfr && !f.model ? `${mfr} - ` : f.model }));
  }

  // ── Save single device to queue ──
  function saveToQueue() {
    if (!form.serial.trim()) { setFormError("Serial number is required."); return; }
    const duplicate = devices.find(d => d.serial.toLowerCase() === form.serial.trim().toLowerCase());
    if (duplicate) { setFormError("This serial number already exists."); return; }
    setFormError("");
    const d = {
      id: genId(), serial: form.serial.trim(), type: form.type,
      mac: form.mac, model: form.model, stage: "Triage",
      outcome: null, received: today(), notes: form.notes,
      sentToPartner: false, partnerOutcome: null, partnerNotes: "",
      pendingAction: form.action,
    };
    setDevices(p => [d, ...p]);
    setOk(d.id);
    setForm({ serial:"", type:DEVICE_TYPES[0], mac:"", model:"", notes:"", action:"Refurbishment" });
    setTimeout(() => setOk(null), 2500);
  }

  // ── Set pending action on a queued device ──
  function setPendingAction(id, action) {
    setDevices(p => p.map(d => d.id === id ? { ...d, pendingAction: action } : d));
  }

  // ── Edit / Delete ──
  function startEdit(d) {
    setEditingId(d.id);
    setEditBuf({ serial:d.serial, type:d.type, mac:d.mac||"", model:d.model||"", notes:d.notes||"", pendingAction:d.pendingAction||"Refurbishment" });
  }
  function saveEdit(id) {
    setDevices(p => p.map(d => d.id === id ? { ...d, ...editBuf } : d));
    setEditingId(null);
  }
  function deleteDevice(id) {
    setDevices(p => p.filter(d => d.id !== id));
    setDeleteId(null);
  }

  // ── Execute queue ──
  function executeQueue() {
    setDevices(p => p.map(d => {
      if (!["Intake","Triage"].includes(d.stage)) return d;
      if (!d.pendingAction) return d; // leave unassigned in queue
      const isScrap = d.pendingAction === "Scrap";
      return {
        ...d,
        stage: isScrap ? "Scrap" : "Refurbishment",
        outcome: isScrap ? "Scrap" : null,
        sentToPartner: !isScrap,
        pendingAction: null,
      };
    }));
    setShowConfirm(false);
  }

  // ── Bulk upload helpers ──
  function resolveType(raw="") {
    const v = raw.toLowerCase();
    if (v.includes("router")||v.includes("modem")) return "Router/Modem";
    if (v.includes("set-top")||v.includes("stb")||v.includes("settop")) return "Set-top Box";
    if (v.includes("ont")||v.includes("olt")) return "ONT/OLT";
    return null;
  }
  function resolveAction(raw="") {
    const v = raw.toLowerCase();
    if (v.includes("scrap")) return "Scrap";
    if (v.includes("refurb")||v.includes("repair")) return "Refurbishment";
    return null;
  }
  function normalizeRow(row) {
    const o={};
    Object.keys(row).forEach(k => { o[k.trim().toLowerCase().replace(/\s+/g,"_")] = String(row[k]||"").trim(); });
    return o;
  }
  function applyRows(rows) {
    const norm = rows.map(normalizeRow).filter(r => r.serial_number||r.serial);
    const added=[], skipped=[], warnings=[];
    const existing = new Set(devices.map(d => d.serial.toLowerCase()));
    const newDevs = [];
    norm.forEach(row => {
      const serial = (row.serial_number||row.serial||"").trim();
      if (!serial) { skipped.push("(empty)"); return; }
      if (existing.has(serial.toLowerCase())) { skipped.push(serial); return; }
      const rawType = row.device_type||row.type||"";
      const type = resolveType(rawType) || DEVICE_TYPES[0];
      if (rawType && !resolveType(rawType)) warnings.push(`${serial}: unknown type → Router/Modem`);
      const mac    = row.mac_address||row.mac||"";
      const mfr    = lookupOUI(mac);
      const model  = row.model||row.device_model||(mfr?`${mfr} - `:"");
      const rawAction = row.action||row.routing||"";
      const pendingAction = resolveAction(rawAction) || null; // null = unassigned in queue
      if (rawAction && !resolveAction(rawAction)) warnings.push(`${serial}: unknown action "${rawAction}" — unassigned in queue`);
      newDevs.push({
        id: genId(), serial, type, mac, model, stage:"Triage", outcome:null,
        received: row.received_date||row.date||today(), notes:row.notes||"",
        sentToPartner:false, partnerOutcome:null, partnerNotes:"", pendingAction,
      });
      existing.add(serial.toLowerCase());
      added.push(serial);
    });
    setDevices(p => [...newDevs, ...p]);
    const log = {
      id:Date.now(), timestamp:new Date().toLocaleString(), total:norm.length,
      added:added.length, skipped:skipped.length, warnings:warnings.length,
      skippedSerials:skipped, warningMessages:warnings,
    };
    setUploadLogs(p => [log, ...p]);
    setUploadResult(log);
  }
  function parseFile(file) {
    setProcessing(true); setUploadResult(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext==="csv") {
      Papa.parse(file,{ header:true, skipEmptyLines:true,
        complete:(r)=>{ applyRows(r.data); setProcessing(false); }, error:()=>setProcessing(false) });
    } else if (ext==="xlsx"||ext==="xls") {
      const rd = new FileReader();
      rd.onload = e => { const wb=XLSX.read(e.target.result,{type:"array"}); applyRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""})); setProcessing(false); };
      rd.readAsArrayBuffer(file);
    } else setProcessing(false);
  }
  function downloadTemplate() {
    const rows=[
      ["serial_number","device_type","mac_address","model","received_date","action","notes"],
      ["SN-10001","Router/Modem","E8:65:D4:11:22:33","Huawei HG8245H","2025-05-20","Refurbishment","Customer return"],
      ["SN-10002","Set-top Box","","","2025-05-20","Scrap","Physical damage"],
      ["SN-10003","ONT/OLT","","","2025-05-20","Refurbishment",""],
    ];
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([rows.map(r=>r.join(",")).join("\n")],{type:"text/csv"}));
    a.download = "intake_template.csv"; a.click();
  }

  const tabStyle = active => ({
    padding:"8px 20px", border:"none", borderRadius:8, fontSize:13, fontWeight:700,
    cursor:"pointer", background:active?C.indigo:"transparent",
    color:active?C.white:C.slate3, transition:"all .15s"
  });

  // ── Action badge style ──
  function actionBadge(action) {
    if (!action) return { bg:"#F1F5F9", color:C.slate3, label:"— Unassigned" };
    if (action==="Scrap") return { bg:C.redLight, color:C.redDark, label:"🗑 Scrap" };
    return { bg:"#EFF6FF", color:"#1E40AF", label:"🔧 Refurbishment" };
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Intake & Triage</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Register returns, assign actions, then execute the queue</p>
      </div>

      {/* ── Register card ── */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <SectionTitle>Register returned devices</SectionTitle>
          <div style={{ display:"flex", background:C.slate7, borderRadius:10, padding:3, gap:2 }}>
            <button style={tabStyle(tab==="manual")} onClick={()=>setTab("manual")}>✏️ Manual</button>
            <button style={tabStyle(tab==="bulk")}   onClick={()=>setTab("bulk")}>📂 Bulk</button>
          </div>
        </div>

        {/* ── MANUAL ENTRY ── */}
        {tab==="manual" && (
          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
            <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(5,1fr)", gap:10 }}>
              <div>
                <Label>SERIAL NUMBER *</Label>
                <input value={form.serial} onChange={e=>{ setFormError(""); setForm(f=>({...f,serial:e.target.value})); }}
                  onKeyDown={e=>e.key==="Enter"&&saveToQueue()} placeholder="SN-12345"
                  style={iStyle({ borderColor: formError && !form.serial ? C.red : C.slate6 })}/>
              </div>
              <div>
                <Label>DEVICE TYPE</Label>
                <select value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))} style={iStyle()}>
                  {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
                </select>
              </div>
              <div>
                <Label>MAC ADDRESS</Label>
                <input value={form.mac} onChange={e=>handleMacChange(e.target.value, setForm)}
                  placeholder="E8:65:D4:11:22:33" style={iStyle({fontFamily:"monospace"})} maxLength={17}/>
              </div>
              <div>
                <Label>DEVICE MODEL</Label>
                <input value={form.model} onChange={e=>setForm(f=>({...f,model:e.target.value}))}
                  placeholder={lookupOUI(form.mac)?`${lookupOUI(form.mac)} - model`:"Auto from MAC"}
                  style={iStyle({ background: lookupOUI(form.mac)&&!form.model?"#FFFBEB":C.white })}/>
              </div>
              <div>
                <Label>ACTION *</Label>
                <select value={form.action} onChange={e=>setForm(f=>({...f,action:e.target.value}))}
                  style={iStyle({ fontWeight:700,
                    background: form.action==="Scrap" ? "#FFF0F0" : "#EFF6FF",
                    color: form.action==="Scrap" ? C.redDark : "#1E40AF",
                    borderColor: form.action==="Scrap" ? "#FECACA" : "#BFDBFE" })}>
                  <option value="Refurbishment">🔧 Refurbishment</option>
                  <option value="Scrap">🗑 Scrap</option>
                </select>
              </div>
            </div>
            {lookupOUI(form.mac) && (
              <Alert type="success">🔍 OUI match: <strong>{lookupOUI(form.mac)}</strong> — complete the model number above</Alert>
            )}
            <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr auto", gap:10, alignItems:"end" }}>
              <div>
                <Label>NOTES</Label>
                <input value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder="Optional notes" style={iStyle()}/>
              </div>
              <Btn onClick={saveToQueue} variant="primary" size={isMobile?"lg":"md"} full={isMobile}>
                💾 Save to Queue
              </Btn>
            </div>
            {formError && <Alert type="danger">⚠ {formError}</Alert>}
            {ok && <Alert type="success">✓ Device {ok} saved to triage queue with action: {form.action||"Refurbishment"}</Alert>}
          </div>
        )}

        {/* ── BULK UPLOAD ── */}
        {tab==="bulk" && (
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
              <span style={{ fontSize:13, color:C.slate3 }}>
                Upload CSV or Excel. Include an <code style={{ background:C.slate7, padding:"1px 6px", borderRadius:4, fontSize:12 }}>action</code> column
                (Refurbishment / Scrap) — or leave it blank to assign in the queue. Duplicates skipped automatically.
              </span>
              <Btn onClick={downloadTemplate} variant="purple" size="sm">↓ Template</Btn>
            </div>
            <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)parseFile(f);}}
              onClick={()=>document.getElementById("intake-file").click()}
              style={{ border:`2px dashed ${dragOver?C.indigo:C.slate5}`, borderRadius:12, padding:"28px 16px",
                textAlign:"center", background:dragOver?C.indigoLight:C.slate8, transition:"all .2s", cursor:"pointer" }}>
              <input id="intake-file" type="file" accept=".csv,.xlsx,.xls"
                onChange={e=>{const f=e.target.files[0];if(f)parseFile(f);e.target.value="";}} style={{ display:"none" }}/>
              <div style={{ fontSize:28, marginBottom:6 }}>{processing?"⏳":"📋"}</div>
              <div style={{ fontSize:13, fontWeight:700, color:C.slate2, marginBottom:4 }}>
                {processing?"Processing…":"Drop file here or tap to browse"}
              </div>
              <div style={{ fontSize:11, color:C.slate4 }}>
                serial_number · device_type · mac_address · model · action · received_date · notes
              </div>
            </div>

            {/* Column reference */}
            <div style={{ background:C.slate8, borderRadius:8, padding:"10px 14px" }}>
              <div style={{ fontSize:10, fontWeight:700, color:C.slate4, letterSpacing:".06em", marginBottom:8 }}>ACCEPTED COLUMNS</div>
              <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:6 }}>
                {[
                  ["serial_number","Required · must be unique"],
                  ["device_type","Router/Modem · Set-top Box · ONT/OLT"],
                  ["action","Refurbishment or Scrap (optional — can assign in queue)"],
                  ["mac_address","Optional · triggers OUI manufacturer lookup"],
                  ["model","Optional · device model name"],
                  ["received_date","Optional · defaults to today"],
                  ["notes","Optional · free text"],
                ].map(([col,desc])=>(
                  <div key={col} style={{ fontSize:12 }}>
                    <code style={{ background:C.slate7, padding:"1px 6px", borderRadius:4, fontSize:11 }}>{col}</code>
                    <span style={{ color:C.slate4, marginLeft:6 }}>{desc}</span>
                  </div>
                ))}
              </div>
            </div>

            {uploadResult && (
              <div style={{ border:`1.5px solid ${C.slate6}`, borderRadius:12, padding:14 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.slate, marginBottom:10 }}>
                  Upload complete · {uploadResult.timestamp}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10 }}>
                  {[[uploadResult.added,"Added",C.greenDark,C.greenLight],[uploadResult.skipped,"Skipped","#92400E",C.amberLight],[uploadResult.warnings,"Warnings",C.purple,C.purpleLight]].map(([n,l,c,bg])=>(
                    <div key={l} style={{ background:bg, borderRadius:8, padding:10, textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:800, color:c }}>{n}</div>
                      <div style={{ fontSize:11, color:c, fontWeight:700 }}>{l}</div>
                    </div>
                  ))}
                </div>
                {uploadResult.skippedSerials.length>0 && <Alert type="warning">Skipped: {uploadResult.skippedSerials.join(", ")}</Alert>}
                {uploadResult.warningMessages.length>0 && (
                  <div style={{ marginTop:8 }}>
                    {uploadResult.warningMessages.map((w,i)=><Alert key={i} type="warning">⚠ {w}</Alert>)}
                  </div>
                )}
                {unassigned>0 && <Alert type="info" style={{ marginTop:8 }}>ℹ {unassigned} device{unassigned>1?"s":""} have no action assigned — set them in the queue below before executing.</Alert>}
              </div>
            )}
          </div>
        )}
      </Card>

      {/* ── Triage Queue ── */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <SectionTitle>Triage queue ({queue.length})</SectionTitle>
            {unassigned>0 && (
              <span style={{ background:C.amberLight, color:"#92400E", fontSize:11, fontWeight:700,
                padding:"2px 9px", borderRadius:20 }}>{unassigned} unassigned</span>
            )}
          </div>
          <Btn onClick={()=>setShowConfirm(true)} size={isMobile?"lg":"md"} disabled={readyCount===0}
            style={{ background:readyCount>0?"#16A34A":"#D1FAE5", color:readyCount>0?"#fff":"#6EE7B7",
              border:"none", fontWeight:800, letterSpacing:".02em",
              boxShadow:readyCount>0?"0 2px 8px rgba(22,163,74,.35)":"none",
              transition:"all .2s" }}>
            ▶ Execute Queue {readyCount>0?`(${readyCount})`:""}
          </Btn>
        </div>

        {queue.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>No devices awaiting triage. Register devices above to get started.</p>
          : isMobile
            ? queue.map(d => {
                const isEditing  = editingId===d.id;
                const isDeleting = deleteId===d.id;
                const expanded   = expandedId===d.id;
                const ab = actionBadge(d.pendingAction);

                if (isEditing) return (
                  <div key={d.id} style={{ border:`2px solid ${C.indigo}`, borderRadius:12, padding:14, marginBottom:10, background:C.indigoLight }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.indigo, marginBottom:10 }}>Editing {d.id}</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:8 }}>
                      <div><Label>SERIAL</Label><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>TYPE</Label>
                        <select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={iStyle()}>
                          {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
                        </select>
                      </div>
                      <div><Label>MAC</Label><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={iStyle({fontFamily:"monospace"})} maxLength={17}/></div>
                      <div><Label>MODEL</Label><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>ACTION</Label>
                        <select value={editBuf.pendingAction||"Refurbishment"} onChange={e=>setEditBuf(b=>({...b,pendingAction:e.target.value}))} style={iStyle()}>
                          <option value="Refurbishment">🔧 Refurbishment</option>
                          <option value="Scrap">🗑 Scrap</option>
                        </select>
                      </div>
                      <div><Label>NOTES</Label><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={iStyle()}/></div>
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>saveEdit(d.id)} variant="success" full>✓ Save</Btn>
                      <Btn onClick={()=>setEditingId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isDeleting) return (
                  <div key={d.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>Delete {d.serial}? This cannot be undone.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>deleteDevice(d.id)} variant="danger" full>Delete</Btn>
                      <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                return (
                  <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, marginBottom:10,
                    borderLeft:`3px solid ${d.pendingAction==="Scrap"?C.red:d.pendingAction?C.indigo:C.slate5}` }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                      <div>
                        <span style={{ fontWeight:800, color:C.indigo, fontSize:12 }}>{d.id}</span>
                        <span style={{ fontFamily:"monospace", fontSize:12, color:C.slate2, marginLeft:8 }}>{d.serial}</span>
                      </div>
                      <button onClick={()=>setExpandedId(expanded?null:d.id)}
                        style={{ background:"none", border:"none", cursor:"pointer", fontSize:14, color:C.slate4 }}>
                        {expanded?"▲":"▼"}
                      </button>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:10 }}>
                      <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Type</span><div style={{ fontSize:12 }}>{TYPE_ICON[d.type]} {d.type}</div></div>
                      <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Received</span><div style={{ fontSize:12, color:C.slate3 }}>{d.received}</div></div>
                      {expanded && <>
                        <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>MAC</span><div style={{ fontSize:11, fontFamily:"monospace" }}>{d.mac||"—"}</div></div>
                        <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Model</span><div style={{ fontSize:12 }}>{d.model||"—"}</div></div>
                        {d.notes && <div style={{ gridColumn:"1/-1" }}><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Notes</span><div style={{ fontSize:12, color:C.slate3 }}>{d.notes}</div></div>}
                      </>}
                    </div>
                    {/* Action status badge — read-only, editable via ✏️ Edit */}
                    <div style={{ marginBottom:10 }}>
                      <span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".05em" }}>ACTION</span>
                      <div style={{ marginTop:4 }}>
                        {d.pendingAction
                          ? <span style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:700,
                              background: d.pendingAction==="Scrap" ? C.redLight : C.indigoLight,
                              color: d.pendingAction==="Scrap" ? C.redDark : C.indigoDark,
                              border: `1px solid ${d.pendingAction==="Scrap" ? "#FECACA" : "#C7D2FE"}` }}>
                              {d.pendingAction==="Scrap" ? "🗑 Scrap" : "🔧 Refurbishment"}
                            </span>
                          : <span style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:700,
                              background:C.slate7, color:C.slate4, border:`1px solid ${C.slate6}` }}>
                              — Unassigned
                            </span>
                        }
                      </div>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
                      <Btn onClick={()=>startEdit(d)} variant="ghost" size="sm" full>✏️ Edit</Btn>
                      <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm" full>🗑️ Delete</Btn>
                    </div>
                  </div>
                );
              })
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["ID","Serial","Type","MAC","Model","Notes","Action","Controls"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {queue.map(d => {
                      const isEditing  = editingId===d.id;
                      const isDeleting = deleteId===d.id;
                      const ab = actionBadge(d.pendingAction);
                      return (
                        <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                          background: isEditing?C.indigoLight : isDeleting?C.redLight : "transparent",
                          borderLeft:`3px solid ${d.pendingAction==="Scrap"?C.red:d.pendingAction?C.indigo:C.slate5}` }}>
                          <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>

                          {isEditing ? (
                            <>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={{ ...iStyle(), width:120 }}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={{ ...iStyle({fontFamily:"monospace"}), width:130 }} maxLength={17}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={{ ...iStyle(), width:130 }}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}>
                                <select value={editBuf.pendingAction||"Refurbishment"} onChange={e=>setEditBuf(b=>({...b,pendingAction:e.target.value}))} style={{ ...iStyle(), width:130 }}>
                                  <option value="Refurbishment">🔧 Refurbishment</option>
                                  <option value="Scrap">🗑 Scrap</option>
                                </select>
                              </td>
                              <td style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>saveEdit(d.id)} variant="success" size="sm">✓ Save</Btn>
                                  <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isDeleting ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:C.redDark, fontWeight:600 }}>
                                Delete <span style={{ fontFamily:"monospace" }}>{d.serial}</span>? This cannot be undone.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>deleteDevice(d.id)} variant="danger" size="sm">Delete</Btn>
                                  <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : (
                            <>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                              <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                              <td style={{ padding:"9px 10px", fontSize:12 }}>
                                {d.model ? d.model : d.mac&&lookupOUI(d.mac)
                                  ? <span style={{ color:C.amber, fontStyle:"italic" }}>{lookupOUI(d.mac)} - ?</span>
                                  : <span style={{ color:C.slate5 }}>—</span>}
                              </td>
                              <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:120, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                              {/* Action status — read-only badge */}
                              <td style={{ padding:"9px 10px" }}>
                                {d.pendingAction
                                  ? <span style={{ display:"inline-flex", alignItems:"center", gap:4, padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:700, whiteSpace:"nowrap",
                                      background: d.pendingAction==="Scrap" ? C.redLight : C.indigoLight,
                                      color: d.pendingAction==="Scrap" ? C.redDark : C.indigoDark,
                                      border:`1px solid ${d.pendingAction==="Scrap"?"#FECACA":"#C7D2FE"}` }}>
                                      {d.pendingAction==="Scrap" ? "🗑 Scrap" : "🔧 Refurbishment"}
                                    </span>
                                  : <span style={{ display:"inline-flex", alignItems:"center", padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:600,
                                      background:C.slate7, color:C.slate4, border:`1px solid ${C.slate6}` }}>
                                      — Unassigned
                                    </span>
                                }
                              </td>
                              <td style={{ padding:"5px 8px" }}>
                                <div style={{ display:"flex", gap:4 }}>
                                  <Btn onClick={()=>startEdit(d)} variant="ghost" size="sm">✏️</Btn>
                                  <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm">🗑️</Btn>
                                </div>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )
        }

        {/* Execute Queue button at bottom */}
        {queue.length > 0 && (
          <div style={{ marginTop:16, paddingTop:14, borderTop:`1px solid ${C.slate7}`,
            display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
            <div style={{ fontSize:13, color:C.slate3 }}>
              {readyCount>0
                ? <span><strong style={{ color:C.indigoDark }}>{toRefurb}</strong> → Refurb &nbsp;·&nbsp; <strong style={{ color:C.redDark }}>{toScrap}</strong> → Scrap {unassigned>0 && <span style={{ color:"#92400E" }}>· <strong>{unassigned}</strong> unassigned</span>}</span>
                : <span style={{ color:C.slate4 }}>Assign actions to devices above, then execute.</span>
              }
            </div>
            <Btn onClick={()=>setShowConfirm(true)} disabled={readyCount===0}
              style={{ background:readyCount>0?"#16A34A":"#D1FAE5", color:readyCount>0?"#fff":"#6EE7B7",
                border:"none", fontWeight:800, letterSpacing:".02em",
                boxShadow:readyCount>0?"0 2px 8px rgba(22,163,74,.35)":"none",
                transition:"all .2s" }}>
              ▶ Execute Queue {readyCount>0?`(${readyCount})`:""}
            </Btn>
          </div>
        )}
      </Card>

      {/* ── CONFIRMATION POPUP ── */}
      {showConfirm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.55)", zIndex:100,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:440, width:"100%",
            boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>
            <div style={{ fontSize:24, marginBottom:8 }}>⚙️</div>
            <h3 style={{ margin:"0 0 6px", fontSize:18, fontWeight:800, color:C.slate }}>Execute Triage Queue</h3>
            <p style={{ margin:"0 0 20px", fontSize:13, color:C.slate3 }}>
              This will dispatch all assigned devices to their target stages. This action cannot be undone.
            </p>

            {/* Summary */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:20 }}>
              <div style={{ background:C.indigoLight, borderRadius:10, padding:"14px 16px",
                border:`1.5px solid #C7D2FE`, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:800, color:C.indigo }}>{toRefurb}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.indigoDark }}>🔧 → Refurbishment</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Sent to partner</div>
              </div>
              <div style={{ background:C.redLight, borderRadius:10, padding:"14px 16px",
                border:`1.5px solid #FECACA`, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:800, color:C.red }}>{toScrap}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.redDark }}>🗑 → Scrap</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Removed from circulation</div>
              </div>
            </div>

            {unassigned>0 && (
              <Alert type="warning">
                ⚠ <strong>{unassigned}</strong> device{unassigned>1?"s":""} have no action assigned and will stay in the queue.
              </Alert>
            )}

            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:16 }}>
              <Btn onClick={()=>setShowConfirm(false)} variant="ghost" full size="lg">Cancel</Btn>
              <Btn onClick={executeQueue} full size="lg" disabled={readyCount===0}
                style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                  boxShadow:"0 2px 10px rgba(22,163,74,.4)", transition:"all .2s" }}>
                ✓ Confirm & Execute
              </Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
// ─────────────────────────────────────────────
// REFURB TRACKING
// ─────────────────────────────────────────────
function RefurbTracking({ devices, setDevices, isMobile }) {
  const inRefurb    = devices.filter(d=>d.stage==="Refurbishment" && !d.partnerOutcome);
  const pendingConf = devices.filter(d=>d.stage==="Refurbishment" && d.partnerOutcome);
  const inQC        = devices.filter(d=>d.stage==="QC Check");
  const escalated   = devices.filter(d=>d.stage==="Escalated");

  // In Refurbishment row state
  const [editingId, setEditingId] = useState(null);
  const [editBuf,   setEditBuf]   = useState({});
  const [deleteId,  setDeleteId]  = useState(null);
  const [returnId,  setReturnId]  = useState(null);

  function setOutcome(id, outcome) { setDevices(p=>p.map(d=>d.id===id?{...d,stage:outcome==="Working"?"Stock":"Escalated",outcome}:d)); }
  function escalateToScrap(id)     { setDevices(p=>p.map(d=>d.id===id?{...d,stage:"Scrap",outcome:"Scrap"}:d)); }
  function requeue(id)             { setDevices(p=>p.map(d=>d.id===id?{...d,stage:"Refurbishment",outcome:null,partnerOutcome:null,partnerNotes:""}:d)); }

  // Return device to Triage queue (unassigned, ready to re-triage)
  function returnToTriage(id) {
    setDevices(p=>p.map(d=>d.id===id
      ? { ...d, stage:"Triage", outcome:null, sentToPartner:false,
          partnerOutcome:null, partnerNotes:"", pendingAction:null }
      : d));
    setReturnId(null);
  }

  function saveEdit(id) {
    setDevices(p=>p.map(d=>d.id===id ? { ...d,...editBuf } : d));
    setEditingId(null);
  }

  function deleteDevice(id) {
    setDevices(p=>p.filter(d=>d.id!==id));
    setDeleteId(null);
  }

  function DevSection({ title, items, accent, renderActions, extra }) {
    return (
      <div style={{ border:`1.5px solid ${accent}22`, borderTop:`3px solid ${accent}`, borderRadius:14, padding:16, background:C.white }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
          <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:C.slate }}>{title}</h3>
          <span style={{ fontSize:12, color:C.slate4 }}>({items.length})</span>
        </div>
        {extra}
        {items.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>None at this stage.</p>
          : isMobile
            ? items.map(d=>(
                <DeviceCard key={d.id} d={d}
                  fields={[["Model",d.model],["Type",`${TYPE_ICON[d.type]} ${d.type}`],["Received",d.received],["Notes",d.notes]]}
                  actions={renderActions(d)}/>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead><tr style={{ borderBottom:`1.5px solid ${C.slate6}` }}>
                    {["ID","Serial","Model","Type","Received","Notes","Action"].map(h=>(
                      <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {items.map(d=>(
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                        <td style={{ padding:"9px 10px" }}><div style={{ display:"flex", gap:6 }}>{renderActions(d)}</div></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </div>
    );
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Refurbishment</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Track devices through repair, QC, and confirmation</p>
      </div>

      {/* Pending partner confirmation */}
      {pendingConf.length>0 && (
        <div style={{ border:`2px solid #FCD34D`, borderTop:`3px solid ${C.amber}`, borderRadius:14, padding:16, background:C.white }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
            <span style={{ fontSize:16 }}>⚠️</span>
            <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:"#92400E" }}>Partner Results — Pending Confirmation ({pendingConf.length})</h3>
          </div>
          {isMobile
            ? pendingConf.map(d=>(
                <div key={d.id} style={{ border:`1px solid #FCD34D`, borderRadius:12, padding:12, marginBottom:10, background:"#FFFBEB" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
                    <span style={{ fontWeight:800, color:C.indigo, fontSize:12 }}>{d.id} · {d.serial}</span>
                    <span style={{ fontSize:12, fontWeight:700, color:d.partnerOutcome==="Working"?C.greenDark:C.redDark,
                      background:d.partnerOutcome==="Working"?C.greenLight:C.redLight, padding:"2px 8px", borderRadius:20 }}>
                      {d.partnerOutcome==="Working"?"✓ Working":"✗ Not Working"}
                    </span>
                  </div>
                  {d.partnerNotes && <p style={{ margin:"0 0 8px", fontSize:12, color:C.slate3 }}>Notes: {d.partnerNotes}</p>}
                  <div style={{ display:"flex", gap:6 }}>
                    <Btn onClick={()=>confirmPartner(d.id)} variant="success" size="sm" full>✓ Confirm</Btn>
                    <Btn onClick={()=>rejectPartner(d.id)} variant="danger" size="sm" full>✕ Reject</Btn>
                  </div>
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead><tr style={{ borderBottom:`1.5px solid #FCD34D` }}>
                    {["ID","Serial","Model","Type","Partner Outcome","Partner Notes","Action"].map(h=>(
                      <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:"#92400E", letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {pendingConf.map(d=>(
                      <tr key={d.id} style={{ borderBottom:`1px solid #FEF3C7` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <span style={{ background:d.partnerOutcome==="Working"?C.greenLight:C.redLight,
                            color:d.partnerOutcome==="Working"?C.greenDark:C.redDark,
                            fontSize:12, fontWeight:700, padding:"3px 10px", borderRadius:20 }}>
                            {d.partnerOutcome==="Working"?"✓ Working":"✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.partnerNotes||"—"}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <div style={{ display:"flex", gap:6 }}>
                            <Btn onClick={()=>confirmPartner(d.id)} variant="success" size="sm">✓ Confirm</Btn>
                            <Btn onClick={()=>rejectPartner(d.id)} variant="danger" size="sm">✕ Reject</Btn>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          }
        </div>
      )}

      {/* ── IN REFURBISHMENT — custom section with return/edit/delete ── */}
      <div style={{ border:`1.5px solid ${C.amber}22`, borderTop:`3px solid ${C.amber}`, borderRadius:14, padding:16, background:C.white }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
          <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:C.slate }}>In Refurbishment</h3>
          <span style={{ fontSize:12, color:C.slate4 }}>({inRefurb.length})</span>
        </div>
        {inRefurb.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>None at this stage.</p>
          : isMobile
            ? inRefurb.map(d=>{
                const isEditing  = editingId===d.id;
                const isDeleting = deleteId===d.id;
                const isReturning= returnId===d.id;
                if (isEditing) return (
                  <div key={d.id} style={{ border:`2px solid ${C.indigo}`, borderRadius:12, padding:14, marginBottom:10, background:C.indigoLight }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.indigo, marginBottom:10 }}>Editing {d.id}</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:8 }}>
                      <div><Label>SERIAL</Label><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>TYPE</Label><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={iStyle()}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></div>
                      <div><Label>MAC</Label><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={iStyle({fontFamily:"monospace"})} maxLength={17}/></div>
                      <div><Label>MODEL</Label><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={iStyle()}/></div>
                    </div>
                    <div style={{ marginBottom:8 }}><Label>NOTES</Label><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={iStyle()}/></div>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>saveEdit(d.id)} variant="success" full>✓ Save</Btn>
                      <Btn onClick={()=>setEditingId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isDeleting) return (
                  <div key={d.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>Delete {d.serial}? This cannot be undone.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>deleteDevice(d.id)} variant="danger" full>Delete</Btn>
                      <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isReturning) return (
                  <div key={d.id} style={{ border:`2px solid ${C.amber}`, borderRadius:12, padding:14, marginBottom:10, background:C.amberLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:"#92400E" }}>Return {d.serial} to Triage queue? It will be unassigned and need re-triaging.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>returnToTriage(d.id)} variant="amber" full>↩ Confirm Return</Btn>
                      <Btn onClick={()=>setReturnId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                return (
                  <DeviceCard key={d.id} d={d}
                    fields={[["Model",d.model],["Type",`${TYPE_ICON[d.type]} ${d.type}`],["Received",d.received],["Notes",d.notes]]}
                    actions={[
                      <Btn key="ret" onClick={()=>setReturnId(d.id)}  variant="amber"  size="sm">↩ Return to Triage</Btn>,
                      <Btn key="ed"  onClick={()=>{ setEditingId(d.id); setEditBuf({serial:d.serial,type:d.type,mac:d.mac||"",model:d.model||"",notes:d.notes||""}); }} variant="ghost" size="sm">✏️ Edit</Btn>,
                      <Btn key="del" onClick={()=>setDeleteId(d.id)}  variant="ghost"  size="sm">🗑️ Delete</Btn>,
                    ]}/>
                );
              })
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["ID","Serial","Model","Type","MAC","Received","Notes","Actions"].map(h=>(
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {inRefurb.map(d=>{
                      const isEditing  = editingId===d.id;
                      const isDeleting = deleteId===d.id;
                      const isReturning= returnId===d.id;
                      return (
                        <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                          background: isEditing?C.indigoLight : isDeleting?C.redLight : isReturning?"#FFFBEB" : "transparent" }}>
                          <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                          {isEditing ? (
                            <>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={{ ...iStyle(), width:100 }}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.model}  onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))}  style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={{ ...iStyle(), width:120 }}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={{ ...iStyle({fontFamily:"monospace"}), width:130 }} maxLength={17}/></td>
                              <td style={{ padding:"5px 6px" }}>{d.received}</td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={{ ...iStyle(), width:120 }}/></td>
                              <td style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>saveEdit(d.id)} variant="success" size="sm">✓ Save</Btn>
                                  <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isDeleting ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:C.redDark, fontWeight:600 }}>
                                Delete <span style={{ fontFamily:"monospace" }}>{d.serial}</span>? This cannot be undone.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>deleteDevice(d.id)} variant="danger" size="sm">Delete</Btn>
                                  <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isReturning ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:"#92400E", fontWeight:600 }}>
                                Return <span style={{ fontFamily:"monospace" }}>{d.serial}</span> to Triage queue? It will be unassigned.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>returnToTriage(d.id)} variant="amber" size="sm">↩ Confirm</Btn>
                                  <Btn onClick={()=>setReturnId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : (
                            <>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                              <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                              <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                              <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                              <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:130, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                              <td style={{ padding:"9px 10px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>setReturnId(d.id)} variant="amber" size="sm">↩ Return to Triage</Btn>
                                  <Btn onClick={()=>{ setEditingId(d.id); setEditBuf({serial:d.serial,type:d.type,mac:d.mac||"",model:d.model||"",notes:d.notes||""}); }} variant="ghost" size="sm">✏️</Btn>
                                  <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm">🗑️</Btn>
                                </div>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )
        }
      </div>
      <DevSection title="QC Check" items={inQC} accent="#3B82F6"
        renderActions={d=>[
          <Btn key="w" onClick={()=>setOutcome(d.id,"Working")}     variant="success" size="sm">✓ Working</Btn>,
          <Btn key="n" onClick={()=>setOutcome(d.id,"Not Working")} variant="danger"  size="sm">✗ Not Working</Btn>,
        ]}/>
      <DevSection title="Escalated" items={escalated} accent={C.purple}
        renderActions={d=>[
          <Btn key="r" onClick={()=>requeue(d.id)}        variant="purple" size="sm">↻ Re-queue</Btn>,
          <Btn key="s" onClick={()=>escalateToScrap(d.id)} variant="danger" size="sm">→ Scrap</Btn>,
        ]}/>
    </div>
  );
}

// ─────────────────────────────────────────────
// STOCK & SCRAP (tabbed)
// ─────────────────────────────────────────────
function StockAndScrap({ devices, isMobile }) {
  const [activeTab, setActiveTab] = useState("stock");
  const [filter, setFilter]       = useState("All");

  const isStock = activeTab === "stock";
  const pool    = devices.filter(d => d.stage === (isStock ? "Stock" : "Scrap"));
  const shown   = filter === "All" ? pool : pool.filter(d => d.type === filter);

  // Per-type counts for summary cards
  const typeCounts = DEVICE_TYPES.map(t => ({
    type: t,
    count: pool.filter(d => d.type === t).length,
  }));

  // Download this view as Excel
  function downloadExcel() {
    const rows = shown.map(d => ({
      "Device ID":    d.id,
      "Serial":       d.serial,
      "MAC Address":  d.mac || "",
      "Model":        d.model || "",
      "Type":         d.type,
      "Stage":        d.stage,
      "Outcome":      d.outcome || "",
      "Received":     d.received,
      "Notes":        d.notes || "",
      "Partner":      d.sentToPartner ? "Yes" : "No",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, isStock ? "Stock" : "Scrap");
    XLSX.writeFile(wb, `cpe_${isStock?"stock":"scrap"}_${today()}.xlsx`);
  }

  const accentColor = isStock ? C.green    : C.red;
  const badgeBg     = isStock ? C.greenLight : C.redLight;
  const badgeColor  = isStock ? C.greenDark  : C.redDark;
  const badgeText   = isStock ? "✓ Ready"    : "🗑 Scrapped";

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {/* Header + tab switcher */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>
            {isStock ? "Warehouse Stock" : "Scrap"}
          </h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>
            {isStock ? "Refurbished devices ready for redeployment" : "Devices removed from circulation"}
          </p>
        </div>
        <Tabs
          tabs={[
            { id:"stock", label:`📦 Stock (${devices.filter(d=>d.stage==="Stock").length})` },
            { id:"scrap", label:`🗑 Scrap (${devices.filter(d=>d.stage==="Scrap").length})` },
          ]}
          active={activeTab}
          onChange={t => { setActiveTab(t); setFilter("All"); }}
        />
      </div>

      {/* Type summary cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10 }}>
        {typeCounts.map(({ type, count }) => (
          <Card key={type} accent={accentColor} style={{ display:"flex", alignItems:"center", gap:12 }}>
            <span style={{ fontSize:isMobile?22:28 }}>{TYPE_ICON[type]}</span>
            <div>
              <div style={{ fontSize:isMobile?20:26, fontWeight:800, color:C.slate }}>{count}</div>
              <div style={{ fontSize:11, color:C.slate4 }}>{isMobile ? type.split("/")[0] : type}</div>
            </div>
          </Card>
        ))}
      </div>

      {/* Table / card list */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <SectionTitle>{isStock ? "Stock" : "Scrap"} list ({shown.length})</SectionTitle>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
            {["All", ...DEVICE_TYPES].map(f => (
              <button key={f} onClick={() => setFilter(f)}
                style={{ padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:600, cursor:"pointer", border:"none",
                  background:filter===f ? C.indigo : C.slate7, color:filter===f ? C.white : C.slate3, transition:"all .15s" }}>
                {f === "All" ? "All" : `${TYPE_ICON[f]} ${f.split("/")[0]}`}
              </button>
            ))}
            <Btn onClick={downloadExcel} variant="success" size="sm">↓ Excel</Btn>
          </div>
        </div>

        {shown.length === 0
          ? <p style={{ color:C.slate4, fontSize:13 }}>No {isStock?"stock":"scrap"} for this filter.</p>
          : isMobile
            ? shown.map(d => (
                <DeviceCard key={d.id} d={d}
                  fields={[
                    ["MAC",      d.mac || "—"],
                    ["Model",    d.model || "—"],
                    ["Type",     `${TYPE_ICON[d.type]} ${d.type}`],
                    ["Received", d.received],
                    ["Notes",    d.notes || "—"],
                  ]}
                  actions={[
                    <span key="status" style={{ background:badgeBg, color:badgeColor, fontSize:11,
                      fontWeight:700, padding:"4px 12px", borderRadius:20 }}>{badgeText}</span>
                  ]}/>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}` }}>
                      {["ID","Serial","MAC","Model","Type","Received","Notes","Status"].map(h => (
                        <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map(d => (
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac || "—"}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model || "—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3 }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes || "—"}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <span style={{ background:badgeBg, color:badgeColor, fontSize:11,
                            fontWeight:700, padding:"3px 10px", borderRadius:20 }}>{badgeText}</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// ALL DEVICES + REPORT EXPORT
// ─────────────────────────────────────────────
function AllDevices({ devices, isMobile }) {
  const [search, setSearch]         = useState("");
  const [stageFilter, setStage]     = useState("All");
  const [typeFilter, setTypeFilter] = useState("All");
  const [showReport, setShowReport] = useState(false);
  const [reportStage, setReportStage] = useState("All");
  const [reportType, setReportType]   = useState("All");
  const [reportOutcome, setReportOutcome] = useState("All");
  const [exportSuccess, setExportSuccess] = useState(false);

  const filtered = useMemo(() => devices.filter(d => {
    const ms = stageFilter === "All" || d.stage === stageFilter;
    const mt = typeFilter  === "All" || d.type  === typeFilter;
    const q  = search.toLowerCase();
    return ms && mt && (!q || [d.id, d.serial, d.type, d.mac, d.model, d.notes]
      .some(v => v && v.toLowerCase().includes(q)));
  }), [devices, search, stageFilter, typeFilter]);

  // Report preview — applies its own independent filters
  const reportData = useMemo(() => devices.filter(d => {
    const ms = reportStage   === "All" || d.stage   === reportStage;
    const mt = reportType    === "All" || d.type    === reportType;
    const mo = reportOutcome === "All" || (d.outcome || "—") === reportOutcome;
    return ms && mt && mo;
  }), [devices, reportStage, reportType, reportOutcome]);

  // Stage breakdown for report summary
  const stageSummary = useMemo(() => {
    const base = reportStage === "All" ? STAGES : [reportStage];
    return base.map(s => ({ stage: s, count: reportData.filter(d => d.stage === s).length }))
               .filter(r => r.count > 0);
  }, [reportData, reportStage]);

  function exportReport() {
    // Sheet 1 – device list
    const rows = reportData.map(d => ({
      "Device ID":      d.id,
      "Serial Number":  d.serial,
      "MAC Address":    d.mac    || "",
      "Model":          d.model  || "",
      "Device Type":    d.type,
      "Stage":          d.stage,
      "Outcome":        d.outcome || "",
      "Partner Managed":d.sentToPartner ? "Yes" : "No",
      "Received Date":  d.received,
      "Notes":          d.notes || "",
    }));

    // Sheet 2 – summary by stage
    const summaryRows = STAGES.map(s => ({
      "Stage":       s,
      "Total":       reportData.filter(d => d.stage === s).length,
      "Router/Modem":reportData.filter(d => d.stage === s && d.type === "Router/Modem").length,
      "Set-top Box": reportData.filter(d => d.stage === s && d.type === "Set-top Box").length,
      "ONT/OLT":     reportData.filter(d => d.stage === s && d.type === "ONT/OLT").length,
    }));

    // Sheet 3 – outcome summary
    const outcomes = ["Working","Not Working","Scrap"];
    const outcomeRows = outcomes.map(o => ({
      "Outcome": o,
      "Total":   reportData.filter(d => d.outcome === o).length,
      "Router/Modem": reportData.filter(d => d.outcome === o && d.type === "Router/Modem").length,
      "Set-top Box":  reportData.filter(d => d.outcome === o && d.type === "Set-top Box").length,
      "ONT/OLT":      reportData.filter(d => d.outcome === o && d.type === "ONT/OLT").length,
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),        "Devices");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Stage Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outcomeRows), "Outcome Summary");

    const label = [
      reportStage   !== "All" ? reportStage.replace(/\s+/g,"_")   : "AllStages",
      reportType    !== "All" ? reportType.replace(/[/\s]+/g,"_") : "",
      reportOutcome !== "All" ? reportOutcome.replace(/\s+/g,"_") : "",
    ].filter(Boolean).join("_");

    XLSX.writeFile(wb, `cpe_report_${label}_${today()}.xlsx`);
    setExportSuccess(true);
    setTimeout(() => setExportSuccess(false), 2500);
  }

  const OUTCOMES = ["All", "Working", "Not Working", "Scrap"];

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>All Devices</h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Full device registry · filter, search and export</p>
        </div>
        <Btn onClick={() => setShowReport(r => !r)} variant={showReport ? "primary" : "default"}>
          📊 {showReport ? "Hide" : "Generate"} Report
        </Btn>
      </div>

      {/* ── REPORT PANEL ── */}
      {showReport && (
        <Card style={{ border:`1.5px solid ${C.indigo}44`, background:C.indigoLight }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
            <SectionTitle>📊 Report Builder</SectionTitle>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              {exportSuccess && (
                <span style={{ fontSize:12, color:C.greenDark, fontWeight:700 }}>✓ Downloaded!</span>
              )}
              <Btn onClick={exportReport} variant="success">↓ Export Excel ({reportData.length} devices)</Btn>
            </div>
          </div>

          {/* Report filters */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10, marginBottom:16 }}>
            <div>
              <Label>FILTER BY STAGE</Label>
              <select value={reportStage} onChange={e => setReportStage(e.target.value)} style={iStyle()}>
                <option value="All">All stages</option>
                {STAGES.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <Label>FILTER BY TYPE</Label>
              <select value={reportType} onChange={e => setReportType(e.target.value)} style={iStyle()}>
                <option value="All">All types</option>
                {DEVICE_TYPES.map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
            <div>
              <Label>FILTER BY OUTCOME</Label>
              <select value={reportOutcome} onChange={e => setReportOutcome(e.target.value)} style={iStyle()}>
                {OUTCOMES.map(o => <option key={o}>{o}</option>)}
              </select>
            </div>
          </div>

          {/* Report preview stats */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)", gap:10, marginBottom:14 }}>
            {[
              ["Total",        reportData.length,                                       C.indigo],
              ["Working",      reportData.filter(d=>d.outcome==="Working").length,      C.green],
              ["Not Working",  reportData.filter(d=>d.outcome==="Not Working").length,  C.red],
              ["Scrap",        reportData.filter(d=>d.outcome==="Scrap").length,        "#92400E"],
            ].map(([label, count, color]) => (
              <div key={label} style={{ background:C.white, borderRadius:10, padding:"10px 14px",
                borderLeft:`3px solid ${color}` }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.slate4, textTransform:"uppercase", marginBottom:2 }}>{label}</div>
                <div style={{ fontSize:22, fontWeight:800, color }}>{count}</div>
              </div>
            ))}
          </div>

          {/* Stage breakdown table */}
          {stageSummary.length > 0 && (
            <div style={{ background:C.white, borderRadius:10, overflow:"hidden" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:C.slate7 }}>
                    {["Stage","Total","Router/Modem","Set-top Box","ONT/OLT"].map(h => (
                      <th key={h} style={{ padding:"7px 12px", textAlign:"left", fontSize:11, fontWeight:700,
                        color:C.slate4, letterSpacing:".05em" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {stageSummary.map(({ stage }) => {
                    const sc = STAGE_COLOR[stage] || { dot:C.slate4 };
                    return (
                      <tr key={stage} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"7px 12px" }}>
                          <span style={{ display:"inline-flex", alignItems:"center", gap:6 }}>
                            <span style={{ width:7, height:7, borderRadius:"50%", background:sc.dot }}/>
                            <strong>{stage}</strong>
                          </span>
                        </td>
                        <td style={{ padding:"7px 12px", fontWeight:700 }}>{reportData.filter(d=>d.stage===stage).length}</td>
                        {DEVICE_TYPES.map(t => (
                          <td key={t} style={{ padding:"7px 12px", color:C.slate3 }}>
                            {reportData.filter(d=>d.stage===stage&&d.type===t).length}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          <p style={{ margin:"10px 0 0", fontSize:11, color:C.slate4 }}>
            Export generates a 3-sheet Excel: Device list · Stage summary · Outcome summary
          </p>
        </Card>
      )}

      {/* ── SEARCH & FILTER ROW ── */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"3fr 1fr 1fr", gap:10 }}>
        <input value={search} onChange={e => setSearch(e.target.value)}
          placeholder="Search ID, serial, MAC, model, notes…"
          style={iStyle()}/>
        <select value={stageFilter} onChange={e => setStage(e.target.value)} style={iStyle()}>
          <option value="All">All stages</option>
          {STAGES.map(s => <option key={s}>{s}</option>)}
        </select>
        <select value={typeFilter} onChange={e => setTypeFilter(e.target.value)} style={iStyle()}>
          <option value="All">All types</option>
          {DEVICE_TYPES.map(t => <option key={t}>{t}</option>)}
        </select>
      </div>

      {/* ── DEVICE TABLE / CARDS ── */}
      <Card>
        <div style={{ marginBottom:12, fontSize:13, color:C.slate4, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span>{filtered.length} device{filtered.length !== 1 ? "s" : ""} shown</span>
          {filtered.length > 0 && (
            <button onClick={() => {
              setReportStage(stageFilter);
              setReportType(typeFilter);
              setReportOutcome("All");
              setShowReport(true);
            }} style={{ background:"none", border:"none", color:C.indigo, fontSize:12, fontWeight:600, cursor:"pointer" }}>
              Export this view →
            </button>
          )}
        </div>

        {filtered.length === 0
          ? <p style={{ color:C.slate4, fontSize:13 }}>No devices match the current filters.</p>
          : isMobile
            ? filtered.map(d => (
                <DeviceCard key={d.id} d={d}
                  fields={[
                    ["Type",    `${TYPE_ICON[d.type]} ${d.type}`],
                    ["Stage",   d.stage],
                    ["Outcome", d.outcome || "—"],
                    ["Model",   d.model   || "—"],
                    ["MAC",     d.mac     || "—"],
                    ["Received",d.received],
                  ]}/>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}` }}>
                      {["ID","Serial","MAC","Model","Type","Stage","Outcome","Partner","Received","Notes"].map(h => (
                        <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map(d => (
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac || "—"}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model || "—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}><Badge stage={d.stage}/></td>
                        <td style={{ padding:"9px 10px", fontWeight:600, fontSize:12,
                          color: d.outcome==="Working" ? C.greenDark : d.outcome && d.outcome !== "null" ? C.redDark : C.slate4 }}>
                          {d.outcome || "—"}
                        </td>
                        <td style={{ padding:"9px 10px" }}>
                          {d.sentToPartner
                            ? <span style={{ background:C.purpleLight, color:C.purple, fontSize:11, fontWeight:600, padding:"2px 8px", borderRadius:20 }}>Partner</span>
                            : <span style={{ color:C.slate5 }}>—</span>}
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes || "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// PARTNER PORTAL
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// PARTNER REPORTS  (proper component — hooks at top level)
// ─────────────────────────────────────────────
function PartnerReports({ devices, isMobile }) {
  const [dateFrom,    setDateFrom]    = useState("");
  const [dateTo,      setDateTo]      = useState("");
  const [repOutcome,  setRepOutcome]  = useState("All");
  const [repType,     setRepType]     = useState("All");
  const [repExported, setRepExported] = useState(false);

  const allProcessed = devices.filter(d =>
    d.sentToPartner && d.outcome && ["Stock","Escalated","Scrap"].includes(d.stage)
  );

  const reportData = allProcessed.filter(d => {
    const afterFrom = !dateFrom || d.received >= dateFrom;
    const beforeTo  = !dateTo   || d.received <= dateTo;
    const matchOut  = repOutcome === "All" || d.outcome === repOutcome;
    const matchType = repType    === "All" || d.type    === repType;
    return afterFrom && beforeTo && matchOut && matchType;
  });

  const working    = reportData.filter(d => d.outcome === "Working").length;
  const notWorking = reportData.filter(d => d.outcome === "Not Working").length;
  const rate       = reportData.length ? Math.round(working / reportData.length * 100) : 0;

  function clearFilters() {
    setDateFrom(""); setDateTo(""); setRepOutcome("All"); setRepType("All");
  }

  function exportReport() {
    const rows = reportData.map(d => ({
      "Device ID":     d.id,
      "Serial Number": d.serial,
      "MAC Address":   d.mac    || "",
      "Model":         d.model  || "",
      "Device Type":   d.type,
      "Outcome":       d.outcome || "",
      "Received Date": d.received,
      "Notes":         d.notes  || "",
      "Partner Notes": d.partnerNotes || "",
    }));
    const summaryRows = [
      { "Metric": "Period From",    "Value": dateFrom    || "All time" },
      { "Metric": "Period To",      "Value": dateTo      || "All time" },
      { "Metric": "Device Type",    "Value": repType                   },
      { "Metric": "Outcome Filter", "Value": repOutcome                },
      { "Metric": "Total Devices",  "Value": reportData.length         },
      { "Metric": "Working",        "Value": working                   },
      { "Metric": "Not Working",    "Value": notWorking                },
      { "Metric": "Success Rate %", "Value": rate                      },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),        "Devices");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Summary");
    const safeType    = repType    !== "All" ? repType.replace(/[\/\s]+/g, "_") : "";
    const safeOutcome = repOutcome !== "All" ? repOutcome.replace(/\s+/g, "_")  : "";
    const label = [dateFrom || "start", dateTo || "end", safeType, safeOutcome]
                    .filter(Boolean).join("_");
    XLSX.writeFile(wb, `partner_report_${label}.xlsx`);
    setRepExported(true);
    setTimeout(() => setRepExported(false), 2500);
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>

      {/* ── FILTER CARD ── */}
      <Card>
        <SectionTitle>Report Builder</SectionTitle>
        <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)", gap:12, marginBottom:16 }}>
          <div>
            <Label>DATE FROM</Label>
            <input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={iStyle()}/>
          </div>
          <div>
            <Label>DATE TO</Label>
            <input type="date" value={dateTo}   onChange={e=>setDateTo(e.target.value)}   style={iStyle()}/>
          </div>
          <div>
            <Label>OUTCOME</Label>
            <select value={repOutcome} onChange={e=>setRepOutcome(e.target.value)} style={iStyle()}>
              <option value="All">All outcomes</option>
              <option value="Working">Working</option>
              <option value="Not Working">Not Working</option>
            </select>
          </div>
          <div>
            <Label>DEVICE TYPE</Label>
            <select value={repType} onChange={e=>setRepType(e.target.value)} style={iStyle()}>
              <option value="All">All types</option>
              {DEVICE_TYPES.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
          <span style={{ fontSize:13, color: (dateFrom||dateTo) ? C.slate3 : C.slate4 }}>
            {dateFrom || dateTo
              ? <>Period: <strong>{dateFrom||"start"}</strong> → <strong>{dateTo||"today"}</strong></>
              : "No date filter — showing all time"}
          </span>
          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            {repExported && <span style={{ fontSize:12, color:C.greenDark, fontWeight:700 }}>✓ Downloaded!</span>}
            <Btn onClick={clearFilters} variant="ghost" size="sm">Clear filters</Btn>
            <Btn onClick={exportReport} variant="success">
              ↓ Export Excel ({reportData.length})
            </Btn>
          </div>
        </div>
      </Card>

      {/* ── STAT CARDS ── */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:12 }}>
        <StatCard label="Total in Period" value={reportData.length} sub="Matching filters"   accent={C.purple}/>
        <StatCard label="Working"         value={working}           sub={`${rate}% success`} accent={C.green}/>
        <StatCard label="Not Working"     value={notWorking}        sub="Failed QC"          accent={C.red}/>
      </div>

      {/* ── SUCCESS RATE ── */}
      {reportData.length > 0 && (
        <Card>
          <SectionTitle>Success Rate — Filtered Period</SectionTitle>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
            <span style={{ fontSize:13, color:C.slate3 }}>Working after refurbishment</span>
            <span style={{ fontSize:14, fontWeight:800, color:C.greenDark }}>{rate}%</span>
          </div>
          <div style={{ background:C.slate7, borderRadius:8, height:14, overflow:"hidden", marginBottom:16 }}>
            <div style={{ width:`${rate}%`, height:"100%",
              background:"linear-gradient(90deg,#16A34A,#22C55E)", borderRadius:8, transition:"width .5s" }}/>
          </div>

          {/* Per-type breakdown */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"repeat(3,1fr)", gap:10 }}>
            {DEVICE_TYPES.map(t => {
              const tData    = reportData.filter(d => d.type === t);
              const tWorking = tData.filter(d => d.outcome === "Working").length;
              const tRate    = tData.length ? Math.round(tWorking / tData.length * 100) : 0;
              return (
                <div key={t} style={{ background:C.slate8, borderRadius:10, padding:"10px 14px" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                    <span style={{ fontSize:12, fontWeight:700, color:C.slate }}>{TYPE_ICON[t]} {t}</span>
                    <span style={{ fontSize:12, fontWeight:800, color:tData.length ? C.greenDark : C.slate4 }}>
                      {tData.length ? `${tRate}%` : "—"}
                    </span>
                  </div>
                  <div style={{ background:C.slate6, borderRadius:6, height:8, overflow:"hidden", marginBottom:4 }}>
                    <div style={{ width:`${tRate}%`, height:"100%", background:C.green, borderRadius:6 }}/>
                  </div>
                  <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.slate4 }}>
                    <span>✓ {tWorking} working</span>
                    <span>✗ {tData.length - tWorking} not working</span>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* ── DEVICE LIST ── */}
      {reportData.length > 0 ? (
        <Card>
          <SectionTitle>Devices in report ({reportData.length})</SectionTitle>
          {isMobile
            ? reportData.map(d => (
                <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:10, padding:12, marginBottom:8,
                  borderLeft:`3px solid ${d.outcome==="Working" ? C.green : C.red}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                    <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</span>
                    <span style={{ fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:20,
                      background: d.outcome==="Working" ? C.greenLight : C.redLight,
                      color:      d.outcome==="Working" ? C.greenDark  : C.redDark }}>
                      {d.outcome==="Working" ? "✓ Working" : "✗ Not Working"}
                    </span>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4, fontSize:11, color:C.slate3 }}>
                    <span>{TYPE_ICON[d.type]} {d.type}</span>
                    <span>{d.model || "—"}</span>
                    <span>Received: {d.received}</span>
                    {(d.partnerNotes||d.notes) && <span>Note: {d.partnerNotes||d.notes}</span>}
                  </div>
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["Serial","Model","Type","Outcome","Received","Notes"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {reportData.map(d => (
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12, fontWeight:600 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model || "—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <span style={{ fontSize:11, fontWeight:700, padding:"3px 10px", borderRadius:20,
                            background: d.outcome==="Working" ? C.greenLight : C.redLight,
                            color:      d.outcome==="Working" ? C.greenDark  : C.redDark }}>
                            {d.outcome==="Working" ? "✓ Working" : "✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:160,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {d.partnerNotes || d.notes || "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          }
        </Card>
      ) : (
        <Card>
          <p style={{ color:C.slate4, fontSize:13, textAlign:"center", padding:"20px 0", margin:0 }}>
            No processed devices match the selected filters.
          </p>
        </Card>
      )}
    </div>
  );
}

function PartnerPortal({ devices, setDevices, uploadLogs, setUploadLogs, isMobile }) {
  const [portalTab, setPortalTab]   = useState("live");
  const [edits, setEdits]           = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted]   = useState(false);
  const [showExecConfirm, setShowExecConfirm] = useState(false);
  const [dragOver, setDragOver]     = useState(false);
  const [processing, setProcessing] = useState(false);
  const [lastResult, setLastResult] = useState(null);

  const assigned     = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome);
  const submitted2   = devices.filter(d=>d.sentToPartner && d.partnerOutcome).length;
  const done         = devices.filter(d=>d.sentToPartner && ["Stock","Escalated","Scrap"].includes(d.stage) && d.outcome).length;
  const pendingEdits = Object.values(edits).filter(v=>v.outcome).length;
  const pendingWorking    = Object.values(edits).filter(v=>v.outcome==="Working").length;
  const pendingNotWorking = Object.values(edits).filter(v=>v.outcome==="Not Working").length;

  function setEdit(id, field, value) { setEdits(p=>({ ...p, [id]:{ ...p[id], [field]:value } })); }
  function getEdit(id, field, fb="") { return edits[id]?.[field] ?? fb; }

  // Execute the queue — submit all marked outcomes
  function executeQueue() {
    setSubmitting(true);
    setTimeout(()=>{
      setDevices(p=>p.map(d=>{ const e=edits[d.id]; return e?.outcome ? { ...d,partnerOutcome:e.outcome,partnerNotes:e.notes||"" } : d; }));
      setEdits({}); setSubmitting(false); setSubmitted(true);
      setShowExecConfirm(false);
      setTimeout(()=>setSubmitted(false), 3000);
    }, 600);
  }

  function normalizeRow(row) { const o={}; Object.keys(row).forEach(k=>{ o[k.trim().toLowerCase().replace(/\s+/g,"_")]=String(row[k]||"").trim(); }); return o; }
  function applyRows(rows) {
    const norm=rows.map(normalizeRow).filter(r=>r.serial_number||r.serial);
    const matched=[], unmatched=[], invalid=[], updates={};
    norm.forEach(row=>{
      const serial=(row.serial_number||row.serial||"").trim();
      const rawO=row.outcome||"";
      const outcome=rawO.toLowerCase().includes("not")?"Not Working":rawO.toLowerCase().includes("work")?"Working":null;
      const dev=devices.find(d=>d.serial.toLowerCase()===serial.toLowerCase()&&d.sentToPartner&&d.stage==="Refurbishment");
      if (!dev)    { unmatched.push(serial); return; }
      if (!outcome){ invalid.push(serial);   return; }
      matched.push(serial); updates[dev.id]={ partnerOutcome:outcome, partnerNotes:row.notes||"" };
    });
    setDevices(p=>p.map(d=>updates[d.id]?{...d,...updates[d.id]}:d));
    const log={ id:Date.now(), timestamp:new Date().toLocaleString(), total:norm.length,
      matched:matched.length, unmatched:unmatched.length, invalid:invalid.length,
      unmatchedSerials:unmatched, invalidSerials:invalid };
    setUploadLogs(p=>[log,...p]); setLastResult(log);
  }
  function parseFile(file) {
    setProcessing(true); setLastResult(null);
    const ext=file.name.split(".").pop().toLowerCase();
    if (ext==="csv") { Papa.parse(file,{ header:true, skipEmptyLines:true, complete:(r)=>{ applyRows(r.data); setProcessing(false); }, error:()=>setProcessing(false) }); }
    else if (ext==="xlsx"||ext==="xls") { const rd=new FileReader(); rd.onload=(e)=>{ const wb=XLSX.read(e.target.result,{type:"array"}); applyRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""})); setProcessing(false); }; rd.readAsArrayBuffer(file); }
    else setProcessing(false);
  }
  function downloadTemplate() {
    const rows=[["serial_number","outcome","notes","return_date"],["SN-EXAMPLE1","Working","Firmware updated","2025-05-20"],["SN-EXAMPLE2","Not Working","Power failure","2025-05-20"]];
    const a=document.createElement("a"); a.href=URL.createObjectURL(new Blob([rows.map(r=>r.join(",")).join("\n")],{type:"text/csv"})); a.download="partner_results_template.csv"; a.click();
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ background:"linear-gradient(135deg,#4F46E5,#7C3AED)", borderRadius:14, padding:"20px 20px 16px" }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
          <span style={{ fontSize:18 }}>🔒</span>
          <h2 style={{ margin:0, fontSize:isMobile?16:20, fontWeight:800, color:C.white }}>Partner Portal</h2>
          <span style={{ background:"rgba(255,255,255,.2)", fontSize:10, fontWeight:700, padding:"2px 8px", borderRadius:20, color:C.white }}>RESTRICTED</span>
        </div>
        <p style={{ margin:"0 0 14px", opacity:.8, fontSize:12, color:C.white }}>Submit QC outcomes — internal team will confirm before stock is updated</p>
        <div style={{ display:"flex", background:"rgba(0,0,0,.2)", borderRadius:10, padding:3, gap:2 }}>
          {[{id:"live",label:"📋 Refurb View"},{id:"upload",label:"📂 Bulk Upload"},{id:"history",label:"📊 History"},{id:"reports",label:"📑 Reports"}].map(t=>(
            <button key={t.id} onClick={()=>setPortalTab(t.id)}
              style={{ flex:1, padding:"6px 10px", border:"none", borderRadius:8, fontSize:12, fontWeight:700, cursor:"pointer",
                background:portalTab===t.id?"rgba(255,255,255,.2)":"transparent",
                color:C.white, transition:"all .15s" }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10 }}>
        <StatCard label="Assigned" value={assigned.length} sub="Awaiting your results" accent={C.purple}/>
        <StatCard label="Submitted" value={submitted2} sub="Pending confirmation" accent={C.amber}/>
        <StatCard label="Completed" value={done} sub="Confirmed & closed" accent={C.green}/>
      </div>

      {/* LIVE INPUT — with queue + execute */}
      {portalTab==="live" && (
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <Card>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                <SectionTitle>Assigned devices ({assigned.length})</SectionTitle>
                {pendingEdits>0 && (
                  <span style={{ background:C.indigoLight, color:C.indigoDark, fontSize:11, fontWeight:700,
                    padding:"2px 9px", borderRadius:20 }}>{pendingEdits} in queue</span>
                )}
              </div>
              <Btn onClick={()=>setShowExecConfirm(true)} disabled={pendingEdits===0}
                style={{ background:pendingEdits>0?"#16A34A":"#D1FAE5", color:pendingEdits>0?"#fff":"#6EE7B7",
                  border:"none", fontWeight:800, boxShadow:pendingEdits>0?"0 2px 8px rgba(22,163,74,.35)":"none", transition:"all .2s" }}>
                ▶ Execute Queue {pendingEdits>0?`(${pendingEdits})`:""}
              </Btn>
            </div>

            {submitted && <Alert type="success">✓ Batch submitted — awaiting internal confirmation</Alert>}

            {assigned.length===0
              ? <p style={{ color:C.slate4, fontSize:13 }}>No devices currently assigned to your facility.</p>
              : (
                <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                  {assigned.map(d=>{
                    const outcome = getEdit(d.id,"outcome","");
                    const notes   = getEdit(d.id,"notes","");
                    return (
                      <div key={d.id} style={{
                        border:`1.5px solid ${outcome==="Working"?"#BBF7D0":outcome==="Not Working"?"#FECACA":C.slate6}`,
                        borderRadius:12, padding:14,
                        background:outcome==="Working"?"#F0FFF4":outcome==="Not Working"?"#FFF5F5":C.white,
                        borderLeft:`3px solid ${outcome==="Working"?C.green:outcome==="Not Working"?C.red:C.slate5}`,
                        transition:"all .2s" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:10, flexWrap:"wrap", gap:6 }}>
                          <div>
                            <span style={{ fontFamily:"monospace", fontSize:13, fontWeight:700, color:C.slate2 }}>{d.serial}</span>
                            {d.model && <span style={{ fontSize:12, color:C.slate3, marginLeft:8 }}>{d.model}</span>}
                          </div>
                          <span style={{ fontSize:12, color:C.slate4 }}>{TYPE_ICON[d.type]} {d.type} · {d.received}</span>
                        </div>
                        {d.notes && <p style={{ margin:"0 0 10px", fontSize:12, color:C.slate3, fontStyle:"italic" }}>Internal note: {d.notes}</p>}
                        <div style={{ display:"flex", flexDirection:isMobile?"column":"row", gap:10, alignItems:isMobile?"stretch":"center" }}>
                          <div style={{ display:"flex", gap:8 }}>
                            <button onClick={()=>setEdit(d.id,"outcome",outcome==="Working"?"":"Working")}
                              style={{ flex:1, padding:"8px 14px", borderRadius:9, fontSize:13, fontWeight:700, cursor:"pointer",
                                border:`2px solid ${outcome==="Working"?"#16A34A":C.slate6}`,
                                background:outcome==="Working"?C.greenLight:C.white,
                                color:outcome==="Working"?C.greenDark:C.slate3, transition:"all .15s" }}>
                              ✓ Working
                            </button>
                            <button onClick={()=>setEdit(d.id,"outcome",outcome==="Not Working"?"":"Not Working")}
                              style={{ flex:1, padding:"8px 14px", borderRadius:9, fontSize:13, fontWeight:700, cursor:"pointer",
                                border:`2px solid ${outcome==="Not Working"?C.redDark:C.slate6}`,
                                background:outcome==="Not Working"?C.redLight:C.white,
                                color:outcome==="Not Working"?C.redDark:C.slate3, transition:"all .15s" }}>
                              ✗ Not Working
                            </button>
                          </div>
                          <input value={notes} onChange={e=>setEdit(d.id,"notes",e.target.value)}
                            placeholder="Repair notes…" style={{ ...iStyle(), flex:1, minWidth:0 }}/>
                        </div>
                        {/* Queue status tag */}
                        {outcome && (
                          <div style={{ marginTop:8, display:"flex", alignItems:"center", gap:6 }}>
                            <span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".05em" }}>Queued as:</span>
                            <span style={{ fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20,
                              background:outcome==="Working"?C.greenLight:C.redLight,
                              color:outcome==="Working"?C.greenDark:C.redDark }}>
                              {outcome==="Working"?"✓ Working":"✗ Not Working"}
                            </span>
                            <button onClick={()=>setEdit(d.id,"outcome","")}
                              style={{ background:"none", border:"none", cursor:"pointer", color:C.slate4, fontSize:11, padding:"0 2px" }}>✕ clear</button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )
            }

            {/* Queue summary bar at bottom */}
            {pendingEdits>0 && (
              <div style={{ marginTop:14, paddingTop:12, borderTop:`1px solid ${C.slate7}`,
                display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
                <div style={{ fontSize:13, color:C.slate3 }}>
                  <strong style={{ color:C.greenDark }}>{pendingWorking}</strong> Working ·{" "}
                  <strong style={{ color:C.redDark }}>{pendingNotWorking}</strong> Not Working queued
                </div>
                <Btn onClick={()=>setShowExecConfirm(true)} disabled={pendingEdits===0}
                  style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                    boxShadow:"0 2px 8px rgba(22,163,74,.35)", transition:"all .2s" }}>
                  ▶ Execute Queue ({pendingEdits})
                </Btn>
              </div>
            )}
          </Card>

          {/* ── EXECUTE CONFIRMATION POPUP ── */}
          {showExecConfirm && (
            <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.55)", zIndex:100,
              display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
              <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:440, width:"100%",
                boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>
                <div style={{ fontSize:24, marginBottom:8 }}>📤</div>
                <h3 style={{ margin:"0 0 6px", fontSize:18, fontWeight:800, color:C.slate }}>Execute Outcome Queue</h3>
                <p style={{ margin:"0 0 20px", fontSize:13, color:C.slate3 }}>
                  This will submit all marked outcomes to the internal team for confirmation. You cannot change them after submission.
                </p>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:20 }}>
                  <div style={{ background:C.greenLight, borderRadius:10, padding:"14px 16px",
                    border:`1.5px solid #BBF7D0`, textAlign:"center" }}>
                    <div style={{ fontSize:28, fontWeight:800, color:C.greenDark }}>{pendingWorking}</div>
                    <div style={{ fontSize:12, fontWeight:700, color:C.greenDark }}>✓ Working</div>
                    <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Sent for stock confirmation</div>
                  </div>
                  <div style={{ background:C.redLight, borderRadius:10, padding:"14px 16px",
                    border:`1.5px solid #FECACA`, textAlign:"center" }}>
                    <div style={{ fontSize:28, fontWeight:800, color:C.redDark }}>{pendingNotWorking}</div>
                    <div style={{ fontSize:12, fontWeight:700, color:C.redDark }}>✗ Not Working</div>
                    <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Sent for escalation</div>
                  </div>
                </div>
                {assigned.length - pendingEdits > 0 && (
                  <Alert type="warning">
                    ⚠ <strong>{assigned.length - pendingEdits}</strong> device{assigned.length-pendingEdits>1?"s":""} have no outcome set and will stay in the queue.
                  </Alert>
                )}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:16 }}>
                  <Btn onClick={()=>setShowExecConfirm(false)} variant="ghost" full size="lg">Cancel</Btn>
                  <Btn onClick={executeQueue} full size="lg" disabled={submitting}
                    style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                      boxShadow:"0 2px 10px rgba(22,163,74,.4)" }}>
                    {submitting?"Submitting…":"✓ Confirm & Execute"}
                  </Btn>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* BULK UPLOAD */}
      {portalTab==="upload" && (
        <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:16 }}>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <Card>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                <SectionTitle>Upload QC Results</SectionTitle>
                <Btn onClick={downloadTemplate} variant="purple" size="sm">↓ Template</Btn>
              </div>
              <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
                onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)parseFile(f);}}
                onClick={()=>document.getElementById("partner-file").click()}
                style={{ border:`2px dashed ${dragOver?C.purple:C.slate5}`, borderRadius:12, padding:"28px 16px",
                  textAlign:"center", background:dragOver?C.purpleLight:C.slate8, transition:"all .2s", cursor:"pointer" }}>
                <input id="partner-file" type="file" accept=".csv,.xlsx,.xls"
                  onChange={e=>{const f=e.target.files[0];if(f)parseFile(f);e.target.value="";}} style={{ display:"none" }}/>
                <div style={{ fontSize:28, marginBottom:6 }}>{processing?"⏳":"📂"}</div>
                <div style={{ fontSize:13, fontWeight:700, color:C.slate2, marginBottom:4 }}>{processing?"Processing…":"Drop or tap to browse"}</div>
                <div style={{ fontSize:11, color:C.slate4 }}>CSV or Excel · serial_number · outcome · notes · return_date</div>
              </div>
              <div style={{ marginTop:12, background:C.slate8, borderRadius:8, padding:"10px 14px", fontSize:12, color:C.slate3 }}>
                ℹ Bulk uploads also go through internal confirmation before stock is updated.
              </div>
            </Card>
            {lastResult && (
              <Card>
                <SectionTitle>Upload Result</SectionTitle>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8 }}>
                  {[[lastResult.matched,"Matched",C.greenDark,C.greenLight],[lastResult.unmatched,"Unmatched",C.redDark,C.redLight],[lastResult.invalid,"Invalid","#92400E",C.amberLight]].map(([n,l,c,bg])=>(
                    <div key={l} style={{ background:bg, borderRadius:8, padding:10, textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:800, color:c }}>{n}</div>
                      <div style={{ fontSize:11, color:c, fontWeight:700 }}>{l}</div>
                    </div>
                  ))}
                </div>
                {lastResult.unmatchedSerials.length>0 && <div style={{ marginTop:10 }}><Alert type="danger">Unmatched: {lastResult.unmatchedSerials.join(", ")}</Alert></div>}
              </Card>
            )}
          </div>
          <Card>
            <SectionTitle>Awaiting results ({assigned.length})</SectionTitle>
            {assigned.length===0
              ? <p style={{ color:C.slate4, fontSize:13 }}>None pending.</p>
              : assigned.map(d=>(
                  <div key={d.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
                    padding:"10px 0", borderBottom:`1px solid ${C.slate7}`, gap:8 }}>
                    <div>
                      <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</span>
                      <span style={{ fontSize:11, color:C.slate4, marginLeft:6 }}>{d.model||d.type}</span>
                    </div>
                    <span style={{ fontSize:11, color:C.slate4 }}>{d.received}</span>
                  </div>
                ))
            }
            {uploadLogs.length>0 && (
              <div style={{ marginTop:16, borderTop:`1px solid ${C.slate7}`, paddingTop:14 }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.slate4, marginBottom:8, letterSpacing:".06em" }}>UPLOAD HISTORY</div>
                {uploadLogs.slice(0,5).map(log=>(
                  <div key={log.id} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:6, color:C.slate3 }}>
                    <span>{log.timestamp}</span>
                    <span style={{ color:C.greenDark, fontWeight:700 }}>{log.matched} matched</span>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      )}
      {/* ── HISTORY TAB ── */}
      {portalTab==="history" && (
        <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
          {/* Summary stat cards */}
          {(() => {
            const allProcessed = devices.filter(d => d.sentToPartner && d.outcome && ["Stock","Escalated","Scrap"].includes(d.stage));
            const working    = allProcessed.filter(d => d.outcome==="Working");
            const notWorking = allProcessed.filter(d => d.outcome==="Not Working");
            const byType = DEVICE_TYPES.map(t => ({
              type: t,
              total:      allProcessed.filter(d=>d.type===t).length,
              working:    allProcessed.filter(d=>d.type===t && d.outcome==="Working").length,
              notWorking: allProcessed.filter(d=>d.type===t && d.outcome==="Not Working").length,
            }));
            const rate = allProcessed.length ? Math.round(working.length/allProcessed.length*100) : 0;

            return (
              <>
                {/* Top summary */}
                <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:12 }}>
                  <StatCard label="Total Processed" value={allProcessed.length} sub="All confirmed outcomes" accent={C.purple}/>
                  <StatCard label="Working"     value={working.length}    sub={`${rate}% success rate`} accent={C.green}/>
                  <StatCard label="Not Working" value={notWorking.length} sub="Failed QC" accent={C.red}/>
                </div>

                {/* Success rate bar */}
                <Card>
                  <SectionTitle>Overall Success Rate</SectionTitle>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                    <span style={{ fontSize:13, color:C.slate3 }}>Working after refurbishment</span>
                    <span style={{ fontSize:14, fontWeight:800, color:C.greenDark }}>{rate}%</span>
                  </div>
                  <div style={{ background:C.slate7, borderRadius:8, height:14, overflow:"hidden", marginBottom:8 }}>
                    <div style={{ width:`${rate}%`, height:"100%", background:`linear-gradient(90deg,#16A34A,#22C55E)`, borderRadius:8, transition:"width .5s" }}/>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:14 }}>
                    {byType.map(r => {
                      const typeRate = r.total ? Math.round(r.working/r.total*100) : 0;
                      return (
                        <div key={r.type} style={{ background:C.slate8, borderRadius:10, padding:"10px 14px" }}>
                          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                            <span style={{ fontSize:12, fontWeight:700, color:C.slate }}>{TYPE_ICON[r.type]} {r.type}</span>
                            <span style={{ fontSize:12, fontWeight:800, color:C.greenDark }}>{typeRate}%</span>
                          </div>
                          <div style={{ background:C.slate6, borderRadius:6, height:8, overflow:"hidden", marginBottom:4 }}>
                            <div style={{ width:`${typeRate}%`, height:"100%", background:C.green, borderRadius:6 }}/>
                          </div>
                          <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.slate4 }}>
                            <span>✓ {r.working} working</span>
                            <span>✗ {r.notWorking} not working</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </Card>

                {/* Processed device list split by outcome */}
                <Card>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
                    <SectionTitle>Processed Devices ({allProcessed.length})</SectionTitle>
                    <div style={{ display:"flex", gap:6 }}>
                      {[["all","All",C.slate],["Working","✓ Working",C.greenDark],["Not Working","✗ Not Working",C.redDark]].map(([val,label,color])=>{
                        const [hFilter, setHFilter] = [null, null]; // handled in parent scope
                        return null; // placeholder — handled inline below
                      })}
                    </div>
                  </div>
                  {allProcessed.length === 0
                    ? <p style={{ color:C.slate4, fontSize:13 }}>No devices have been fully processed yet.</p>
                    : (() => {
                        const working   = allProcessed.filter(d=>d.outcome==="Working");
                        const notWorking= allProcessed.filter(d=>d.outcome==="Not Working");

                        function OutcomeSection({ title, items, outcomeColor, outcomeBg, icon }) {
                          if (items.length === 0) return null;
                          return (
                            <div style={{ marginBottom:16 }}>
                              <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10,
                                paddingBottom:8, borderBottom:`2px solid ${outcomeColor}` }}>
                                <span style={{ fontSize:16 }}>{icon}</span>
                                <h4 style={{ margin:0, fontSize:13, fontWeight:800, color:outcomeColor }}>{title}</h4>
                                <span style={{ background:outcomeBg, color:outcomeColor, fontSize:11, fontWeight:700,
                                  padding:"2px 9px", borderRadius:20 }}>{items.length} device{items.length>1?"s":""}</span>
                              </div>
                              {isMobile
                                ? items.map(d=>(
                                    <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:10, padding:12,
                                      marginBottom:8, borderLeft:`3px solid ${outcomeColor}` }}>
                                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                                        <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</span>
                                        <span style={{ fontSize:11, color:C.slate4 }}>{d.received}</span>
                                      </div>
                                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4, fontSize:12 }}>
                                        <span style={{ color:C.slate3 }}>{TYPE_ICON[d.type]} {d.type}</span>
                                        <span style={{ color:C.slate3 }}>{d.model||"—"}</span>
                                      </div>
                                      {d.notes && <div style={{ fontSize:11, color:C.slate4, marginTop:4, fontStyle:"italic" }}>{d.notes}</div>}
                                    </div>
                                  ))
                                : (
                                  <div style={{ overflowX:"auto" }}>
                                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                                      <thead>
                                        <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:outcomeBg }}>
                                          {["Serial","Model","Type","MAC","Received","Notes"].map(h=>(
                                            <th key={h} style={{ padding:"7px 10px", textAlign:"left", fontSize:11,
                                              fontWeight:700, color:outcomeColor, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                                              {h.toUpperCase()}
                                            </th>
                                          ))}
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {items.map(d=>(
                                          <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                                            <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12, fontWeight:600 }}>{d.serial}</td>
                                            <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                                            <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                                            <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                                            <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                                            <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:150,
                                              overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                )
                              }
                            </div>
                          );
                        }

                        return (
                          <>
                            <OutcomeSection title="Working" items={working}
                              outcomeColor={C.greenDark} outcomeBg={C.greenLight} icon="✅"/>
                            <OutcomeSection title="Not Working" items={notWorking}
                              outcomeColor={C.redDark} outcomeBg={C.redLight} icon="❌"/>
                          </>
                        );
                      })()
                  }
                </Card>
              </>
            );
          })()}
        </div>
      )}
      {/* ── REPORTS TAB ── */}
      {portalTab==="reports" && (
        <PartnerReports devices={devices} isMobile={isMobile}/>
      )}

    </div>
  );
}

// ─────────────────────────────────────────────
// APP SHELL
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// ROLE CONFIGURATION
// ─────────────────────────────────────────────
const ROLES = {
  admin: {
    label:   "Administrator",
    icon:    "🛡️",
    color:   "#6366F1",
    bg:      "#EEF2FF",
    views:   ["dashboard","intake","refurb","stock","all","users","partnerportal"],
    canSwitchToPartner: true,
    canManageUsers:     true,
  },
  stock: {
    label:   "Stock Management",
    icon:    "📦",
    color:   "#16A34A",
    bg:      "#F0FDF4",
    views:   ["dashboard","intake","stock","all"],
    canSwitchToPartner: false,
    canManageUsers:     false,
  },
  partner: {
    label:   "Refurbishment Partner",
    icon:    "🔧",
    color:   "#7C3AED",
    bg:      "#FDF4FF",
    views:   [],                // partner has its own portal, no internal views
    canSwitchToPartner: true,
    canManageUsers:     false,
    defaultMode:        "partner",
  },
};

// ─────────────────────────────────────────────
// DEMO USERS  (prototype — no real auth)
// ─────────────────────────────────────────────
const DEMO_USERS_INIT = [
  { id:1, name:"Alice Admin",   username:"admin",   password:"admin123",   role:"admin"   },
  { id:2, name:"Sam Stock",     username:"stock",   password:"stock123",   role:"stock"   },
  { id:3, name:"Paula Partner", username:"partner", password:"partner123", role:"partner" },
];

// ─────────────────────────────────────────────
// LOGIN SCREEN
// ─────────────────────────────────────────────
function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [showPw, setShowPw]     = useState(false);
  const [loading, setLoading]   = useState(false);

  function handleLogin() {
    if (!username || !password) { setError("Please enter both username and password."); return; }
    setLoading(true);
    setTimeout(() => {
      const user = users.find(u => u.username === username && u.password === password);
      if (user) {
        onLogin(user);
      } else {
        setError("Invalid username or password.");
        setLoading(false);
      }
    }, 600);
  }

  return (
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg, ${C.slate} 0%, #1E293B 100%)`,
      display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ width:"100%", maxWidth:400 }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ fontSize:40, marginBottom:10 }}>⚙️</div>
          <h1 style={{ margin:"0 0 4px", fontSize:24, fontWeight:800, color:C.white }}>CPE Refurb Manager</h1>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Sign in to continue</p>
        </div>

        {/* Card */}
        <div style={{ background:C.white, borderRadius:16, padding:28,
          boxShadow:"0 20px 60px rgba(0,0,0,.4)" }}>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <div>
              <Label>USERNAME</Label>
              <input value={username} onChange={e=>{ setError(""); setUsername(e.target.value); }}
                onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                placeholder="Enter your username" autoFocus
                style={iStyle({ fontSize:14 })}/>
            </div>
            <div>
              <Label>PASSWORD</Label>
              <div style={{ position:"relative" }}>
                <input value={password} onChange={e=>{ setError(""); setPassword(e.target.value); }}
                  onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                  type={showPw?"text":"password"} placeholder="Enter your password"
                  style={iStyle({ paddingRight:40, fontSize:14 })}/>
                <button onClick={()=>setShowPw(p=>!p)}
                  style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)",
                    background:"none", border:"none", cursor:"pointer", color:C.slate4, fontSize:16, padding:0 }}>
                  {showPw?"🙈":"👁"}
                </button>
              </div>
            </div>

            {error && <Alert type="danger">⚠ {error}</Alert>}

            <button onClick={handleLogin} disabled={loading}
              style={{ marginTop:4, padding:"12px 0", background:loading?"#4F46E5":C.indigo, color:C.white,
                border:"none", borderRadius:10, fontSize:14, fontWeight:800, cursor:loading?"not-allowed":"pointer",
                boxShadow:"0 2px 10px rgba(99,102,241,.4)", transition:"all .2s", opacity:loading?.8:1 }}>
              {loading ? "Signing in…" : "Sign In →"}
            </button>
          </div>

          {/* Demo credentials hint */}
          <div style={{ marginTop:20, padding:"12px 14px", background:C.slate8,
            borderRadius:10, border:`1px solid ${C.slate6}` }}>
            <div style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em", marginBottom:8 }}>
              DEMO CREDENTIALS
            </div>
            {[
              ["admin","admin123","🛡️ Administrator"],
              ["stock","stock123","📦 Stock Management"],
              ["partner","partner123","🔧 Refurbishment Partner"],
            ].map(([u,p,label])=>(
              <div key={u} onClick={()=>{ setUsername(u); setPassword(p); setError(""); }}
                style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
                  padding:"5px 0", borderBottom:`1px solid ${C.slate7}`, cursor:"pointer" }}>
                <span style={{ fontSize:12, color:C.slate3 }}>{label}</span>
                <span style={{ fontFamily:"monospace", fontSize:11, color:C.indigo, fontWeight:600 }}>
                  {u} / {p}
                </span>
              </div>
            ))}
            <p style={{ margin:"8px 0 0", fontSize:11, color:C.slate4 }}>Click a row to auto-fill</p>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// USER MANAGEMENT  (Admin only)
// ─────────────────────────────────────────────
function UserManagement({ users, setUsers, currentUser, isMobile }) {
  const [showForm, setShowForm]   = useState(false);
  const [editUser, setEditUser]   = useState(null);
  const [deleteId, setDeleteId]   = useState(null);
  const [form, setForm]           = useState({ name:"", username:"", password:"", role:"stock" });
  const [formError, setFormError] = useState("");

  function openAdd()  { setForm({ name:"", username:"", password:"", role:"stock" }); setEditUser(null); setFormError(""); setShowForm(true); }
  function openEdit(u){ setForm({ name:u.name, username:u.username, password:u.password, role:u.role }); setEditUser(u); setFormError(""); setShowForm(true); }

  function saveUser() {
    if (!form.name.trim() || !form.username.trim() || !form.password.trim()) {
      setFormError("All fields are required."); return;
    }
    const duplicate = users.find(u => u.username === form.username && (!editUser || u.id !== editUser.id));
    if (duplicate) { setFormError("Username already exists."); return; }

    if (editUser) {
      setUsers(p => p.map(u => u.id === editUser.id ? { ...u, ...form } : u));
    } else {
      setUsers(p => [...p, { id: Date.now(), ...form }]);
    }
    setShowForm(false);
  }

  function deleteUser(id) {
    if (id === currentUser.id) return; // can't delete yourself
    setUsers(p => p.filter(u => u.id !== id));
    setDeleteId(null);
  }

  const roleConfig = (role) => ROLES[role] || { label:role, icon:"👤", color:C.slate3, bg:C.slate7 };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>User Management</h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Manage user accounts and role assignments</p>
        </div>
        <Btn onClick={openAdd} variant="primary">+ Add User</Btn>
      </div>

      {/* Role summary cards */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10 }}>
        {Object.entries(ROLES).map(([key, r]) => (
          <Card key={key} accent={r.color} style={{ display:"flex", alignItems:"center", gap:12 }}>
            <span style={{ fontSize:24 }}>{r.icon}</span>
            <div>
              <div style={{ fontSize:22, fontWeight:800, color:C.slate }}>
                {users.filter(u => u.role === key).length}
              </div>
              <div style={{ fontSize:12, color:C.slate4 }}>{r.label}</div>
            </div>
          </Card>
        ))}
      </div>

      {/* User list */}
      <Card>
        <SectionTitle>All Users ({users.length})</SectionTitle>
        {isMobile
          ? users.map(u => {
              const rc = roleConfig(u.role);
              const isDeleting = deleteId === u.id;
              if (isDeleting) return (
                <div key={u.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                  <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>
                    Delete user <strong>{u.name}</strong>? This cannot be undone.
                  </p>
                  <div style={{ display:"flex", gap:8 }}>
                    <Btn onClick={()=>deleteUser(u.id)} variant="danger" full>Delete</Btn>
                    <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                  </div>
                </div>
              );
              return (
                <div key={u.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, marginBottom:10,
                  borderLeft:`3px solid ${rc.color}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                    <div>
                      <span style={{ fontWeight:700, fontSize:13, color:C.slate }}>{u.name}</span>
                      {u.id === currentUser.id && <span style={{ marginLeft:6, fontSize:10, fontWeight:700, background:C.indigoLight, color:C.indigo, padding:"1px 7px", borderRadius:20 }}>YOU</span>}
                    </div>
                    <span style={{ background:rc.bg, color:rc.color, fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20 }}>
                      {rc.icon} {rc.label}
                    </span>
                  </div>
                  <div style={{ fontSize:12, color:C.slate4, fontFamily:"monospace", marginBottom:10 }}>@{u.username}</div>
                  <div style={{ display:"flex", gap:6 }}>
                    <Btn onClick={()=>openEdit(u)} variant="ghost" size="sm" full>✏️ Edit</Btn>
                    <Btn onClick={()=>setDeleteId(u.id)} variant="ghost" size="sm" full disabled={u.id===currentUser.id}>🗑️ Delete</Btn>
                  </div>
                </div>
              );
            })
          : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                <thead>
                  <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                    {["Name","Username","Role","Actions"].map(h=>(
                      <th key={h} style={{ padding:"8px 12px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {users.map(u => {
                    const rc = roleConfig(u.role);
                    const isDeleting = deleteId === u.id;
                    return (
                      <tr key={u.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                        background: isDeleting ? C.redLight : "transparent" }}>
                        <td style={{ padding:"10px 12px" }}>
                          <span style={{ fontWeight:600, color:C.slate }}>{u.name}</span>
                          {u.id === currentUser.id && <span style={{ marginLeft:6, fontSize:10, fontWeight:700, background:C.indigoLight, color:C.indigo, padding:"1px 7px", borderRadius:20 }}>YOU</span>}
                        </td>
                        <td style={{ padding:"10px 12px", fontFamily:"monospace", fontSize:12, color:C.slate3 }}>@{u.username}</td>
                        <td style={{ padding:"10px 12px" }}>
                          <span style={{ background:rc.bg, color:rc.color, fontSize:12, fontWeight:700, padding:"3px 10px", borderRadius:20 }}>
                            {rc.icon} {rc.label}
                          </span>
                        </td>
                        <td style={{ padding:"10px 12px" }}>
                          {isDeleting
                            ? <div style={{ display:"flex", gap:6 }}>
                                <span style={{ fontSize:12, color:C.redDark, fontWeight:600, marginRight:4 }}>Confirm delete?</span>
                                <Btn onClick={()=>deleteUser(u.id)} variant="danger" size="sm">Delete</Btn>
                                <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                              </div>
                            : <div style={{ display:"flex", gap:6 }}>
                                <Btn onClick={()=>openEdit(u)} variant="ghost" size="sm">✏️ Edit</Btn>
                                <Btn onClick={()=>setDeleteId(u.id)} variant="ghost" size="sm" disabled={u.id===currentUser.id}>🗑️ Delete</Btn>
                              </div>
                          }
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        }
      </Card>

      {/* Add / Edit modal */}
      {showForm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.5)", zIndex:100,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:420, width:"100%",
            boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>
            <h3 style={{ margin:"0 0 18px", fontSize:17, fontWeight:800, color:C.slate }}>
              {editUser ? "Edit User" : "Add New User"}
            </h3>
            <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
              <div><Label>FULL NAME</Label>
                <input value={form.name} onChange={e=>{ setFormError(""); setForm(f=>({...f,name:e.target.value})); }}
                  placeholder="e.g. Alice Admin" style={iStyle()}/>
              </div>
              <div><Label>USERNAME</Label>
                <input value={form.username} onChange={e=>{ setFormError(""); setForm(f=>({...f,username:e.target.value})); }}
                  placeholder="e.g. alice" style={iStyle({ fontFamily:"monospace" })}/>
              </div>
              <div><Label>PASSWORD</Label>
                <input value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))}
                  type="text" placeholder="Set a password" style={iStyle()}/>
              </div>
              <div><Label>ROLE</Label>
                <select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={iStyle()}>
                  {Object.entries(ROLES).map(([key, r])=>(
                    <option key={key} value={key}>{r.icon} {r.label}</option>
                  ))}
                </select>
              </div>
              {/* Role description */}
              <div style={{ background:ROLES[form.role]?.bg||C.slate8, borderRadius:8, padding:"10px 14px",
                border:`1px solid ${ROLES[form.role]?.color||C.slate6}22` }}>
                <div style={{ fontSize:11, fontWeight:700, color:ROLES[form.role]?.color||C.slate4, marginBottom:4 }}>
                  {ROLES[form.role]?.icon} {ROLES[form.role]?.label} — Access
                </div>
                <div style={{ fontSize:12, color:C.slate3 }}>
                  {form.role==="admin"   && "Full access to all views, user management, and partner portal."}
                  {form.role==="stock"   && "Dashboard, Intake & Triage, Stock & Scrap, and All Devices."}
                  {form.role==="partner" && "Partner portal only — submit QC outcomes and view assigned devices."}
                </div>
              </div>
              {formError && <Alert type="danger">⚠ {formError}</Alert>}
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:20 }}>
              <Btn onClick={()=>setShowForm(false)} variant="ghost" full size="lg">Cancel</Btn>
              <Btn onClick={saveUser} variant="primary" full size="lg">{editUser?"Save Changes":"Add User"}</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// NAV VIEWS CONFIG (all possible views)
// ─────────────────────────────────────────────
const ALL_VIEWS = [
  { id:"dashboard",     label:"Dashboard",        icon:"📊" },
  { id:"intake",        label:"Intake & Triage",  icon:"📥" },
  { id:"refurb",        label:"Refurbishment",    icon:"🔧" },
  { id:"stock",         label:"Stock & Scrap",    icon:"📦" },
  { id:"all",           label:"All Devices",      icon:"🗂️" },
  { id:"users",         label:"User Management",  icon:"👥" },
  { id:"partnerportal", label:"Partner Portal",   icon:"🔒" },
];

// ─────────────────────────────────────────────
// APP
// ─────────────────────────────────────────────
export default function App() {
  const isMobile = useIsMobile();

  // Auth state
  const [currentUser, setCurrentUser] = useState(null);
  const [users, setUsers]             = useState(DEMO_USERS_INIT);

  // App state
  const [view, setView]               = useState("dashboard");
  const [devices, setDevices]         = useState(SEED);
  const [uploadLogs, setUploadLogs]   = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // Derive role config and mode from current user
  const roleConfig    = currentUser ? (ROLES[currentUser.role] || ROLES.admin) : null;
  const isPartnerMode = currentUser?.role === "partner";
  const allowedViews  = roleConfig ? ALL_VIEWS.filter(v => roleConfig.views.includes(v.id)) : [];

  const pipelineCount  = devices.filter(d=>!["Stock","Scrap"].includes(d.stage)).length;
  const partnerPending = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome).length;
  const pendingConf    = devices.filter(d=>d.partnerOutcome && d.stage==="Refurbishment").length;

  function login(user) {
    setCurrentUser(user);
    // Set default view based on role
    if (user.role === "partner") {
      // partner goes straight to partner portal — no internal view needed
    } else if (user.role === "stock") {
      setView("dashboard");
    } else {
      setView("dashboard");
    }
  }

  function logout() {
    setCurrentUser(null);
    setView("dashboard");
    setSidebarOpen(false);
  }

  // ── Not logged in ──
  if (!currentUser) {
    return <LoginScreen users={users} onLogin={login}/>;
  }

  // ── Partner role → go straight to partner portal ──
  if (isPartnerMode) {
    return (
      <div style={{ minHeight:"100vh", background:C.slate8, fontFamily:"system-ui,-apple-system,sans-serif", display:"flex", flexDirection:"column" }}>
        <header style={{ background:C.slate, padding:`0 ${isMobile?14:24}px`, height:52,
          display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, position:"sticky", top:0, zIndex:50 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ color:C.white, fontWeight:800, fontSize:isMobile?13:15 }}>⚙️ CPE Refurb</span>
            <span style={{ background:C.purple, color:C.white, fontSize:9, fontWeight:700, padding:"2px 7px", borderRadius:20 }}>PARTNER</span>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ color:"rgba(255,255,255,.6)", fontSize:12 }}>
              {roleConfig.icon} {currentUser.name}
            </span>
            <button onClick={logout}
              style={{ padding:"4px 12px", background:"rgba(255,255,255,.1)", color:"rgba(255,255,255,.8)",
                border:"1px solid rgba(255,255,255,.2)", borderRadius:8, fontSize:12, fontWeight:600, cursor:"pointer" }}>
              Sign Out
            </button>
          </div>
        </header>
        <main style={{ flex:1, padding:isMobile?"14px":"28px", overflowY:"auto" }}>
          <PartnerPortal devices={devices} setDevices={setDevices}
            uploadLogs={uploadLogs} setUploadLogs={setUploadLogs} isMobile={isMobile}/>
        </main>
      </div>
    );
  }

  const props = { devices, setDevices, isMobile };

  // ── Internal roles (admin / stock) ──
  return (
    <div style={{ minHeight:"100vh", background:C.slate8, fontFamily:"system-ui,-apple-system,sans-serif", display:"flex", flexDirection:"column" }}>

      {/* ── TOP BAR ── */}
      <header style={{ background:C.slate, padding:`0 ${isMobile?14:24}px`, height:52,
        display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, position:"sticky", top:0, zIndex:50 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {isMobile && (
            <button onClick={()=>setSidebarOpen(o=>!o)}
              style={{ background:"none", border:"none", color:C.white, fontSize:20, cursor:"pointer", padding:"0 4px", lineHeight:1 }}>☰</button>
          )}
          <span style={{ color:C.white, fontWeight:800, fontSize:isMobile?13:15 }}>⚙️ CPE Refurb</span>
          <span style={{ background:C.indigo, color:C.white, fontSize:9, fontWeight:700, padding:"2px 7px", borderRadius:20 }}>PROTO</span>
        </div>

        {/* Role badge + user info */}
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ background:roleConfig.bg, color:roleConfig.color, fontSize:11, fontWeight:700,
            padding:"3px 10px", borderRadius:20, display:isMobile?"none":"inline-flex", alignItems:"center", gap:4 }}>
            {roleConfig.icon} {roleConfig.label}
          </span>
          {!isMobile && <span style={{ color:"rgba(255,255,255,.6)", fontSize:12 }}>{currentUser.name}</span>}
          <button onClick={logout}
            style={{ padding:"4px 12px", background:"rgba(255,255,255,.1)", color:"rgba(255,255,255,.8)",
              border:"1px solid rgba(255,255,255,.2)", borderRadius:8, fontSize:12, fontWeight:600, cursor:"pointer" }}>
            {isMobile?"↩":"Sign Out"}
          </button>
        </div>
      </header>

      <div style={{ display:"flex", flex:1, overflow:"hidden", position:"relative" }}>

        {/* ── SIDEBAR / DRAWER ── */}
        <>
          {isMobile && sidebarOpen && (
            <div onClick={()=>setSidebarOpen(false)}
              style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.4)", zIndex:40 }}/>
          )}
          <nav style={{
            width:210, background:C.white, borderRight:`1.5px solid ${C.slate6}`,
            padding:"16px 0", flexShrink:0, display:"flex", flexDirection:"column", gap:2,
            ...(isMobile ? {
              position:"fixed", top:52, left:sidebarOpen?0:-230, bottom:0,
              zIndex:45, width:230, transition:"left .25s", boxShadow:sidebarOpen?"4px 0 20px rgba(0,0,0,.15)":"none"
            } : {})
          }}>
            {/* User info in sidebar */}
            <div style={{ padding:"0 16px 14px", borderBottom:`1px solid ${C.slate7}`, marginBottom:6 }}>
              <div style={{ fontSize:12, fontWeight:700, color:C.slate }}>{currentUser.name}</div>
              <span style={{ background:roleConfig.bg, color:roleConfig.color, fontSize:10, fontWeight:700,
                padding:"2px 8px", borderRadius:20, display:"inline-block", marginTop:3 }}>
                {roleConfig.icon} {roleConfig.label}
              </span>
            </div>

            {allowedViews.map(v=>(
              <button key={v.id} onClick={()=>{ setView(v.id); setSidebarOpen(false); }}
                style={{ width:"100%", padding:"11px 18px", border:"none", textAlign:"left",
                  background:view===v.id?C.indigoLight:"transparent",
                  color:view===v.id?C.indigoDark:C.slate3,
                  fontWeight:view===v.id?700:500, fontSize:13, cursor:"pointer",
                  borderLeft:view===v.id?`3px solid ${C.indigo}`:"3px solid transparent",
                  display:"flex", alignItems:"center", gap:8, transition:"all .15s" }}>
                <span>{v.icon}</span>
                {v.label}
                {v.id==="refurb" && pendingConf>0 && (
                  <span style={{ marginLeft:"auto", background:C.amber, color:C.white, fontSize:9, fontWeight:800,
                    minWidth:16, height:16, borderRadius:20, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 3px" }}>{pendingConf}</span>
                )}
              </button>
            ))}

            {/* Pipeline summary — admin only */}
            {!isMobile && currentUser.role === "admin" && (
              <div style={{ marginTop:"auto", padding:"14px 18px", borderTop:`1px solid ${C.slate7}` }}>
                <div style={{ fontSize:10, color:C.slate4, fontWeight:700, letterSpacing:".07em", marginBottom:8 }}>PIPELINE</div>
                {[
                  ["Triage",       d=>["Intake","Triage"].includes(d.stage),              "#F97316"],
                  ["Refurb/QC",    d=>["Refurbishment","QC Check"].includes(d.stage),     C.amber],
                  ["Escalated",    d=>d.stage==="Escalated",                              C.purple],
                  ["At Partner",   d=>d.sentToPartner&&d.stage==="Refurbishment",         C.purple],
                  ["Pend.Confirm", d=>d.partnerOutcome&&d.stage==="Refurbishment",        C.red],
                ].map(([label,fn,color])=>(
                  <div key={label} style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:C.slate3, marginBottom:4 }}>
                    <span style={{ display:"flex", alignItems:"center", gap:5 }}>
                      <span style={{ width:6, height:6, borderRadius:"50%", background:color, display:"inline-block" }}/>
                      {label}
                    </span>
                    <span style={{ fontWeight:700 }}>{devices.filter(fn).length}</span>
                  </div>
                ))}
              </div>
            )}
          </nav>
        </>

        {/* ── MAIN CONTENT ── */}
        <main style={{ flex:1, padding:isMobile?"14px":"28px", overflowY:"auto", paddingBottom:isMobile?"80px":"28px" }}>
          {view==="dashboard" && <Dashboard {...props}/>}
          {view==="intake"    && allowedViews.find(v=>v.id==="intake")   && <IntakeTriage {...props}/>}
          {view==="refurb"    && allowedViews.find(v=>v.id==="refurb")   && <RefurbTracking {...props}/>}
          {view==="stock"     && allowedViews.find(v=>v.id==="stock")    && <StockAndScrap {...props}/>}
          {view==="all"       && allowedViews.find(v=>v.id==="all")      && <AllDevices {...props}/>}
          {view==="users"     && allowedViews.find(v=>v.id==="users")    && (
            <UserManagement users={users} setUsers={setUsers} currentUser={currentUser} isMobile={isMobile}/>
          )}
          {view==="partnerportal" && allowedViews.find(v=>v.id==="partnerportal") && (
            <PartnerPortal devices={devices} setDevices={setDevices}
              uploadLogs={uploadLogs} setUploadLogs={setUploadLogs} isMobile={isMobile}/>
          )}
        </main>
      </div>

      {/* ── MOBILE BOTTOM NAV ── */}
      {isMobile && (
        <nav style={{ position:"fixed", bottom:0, left:0, right:0, background:C.white,
          borderTop:`1px solid ${C.slate6}`, display:"flex", zIndex:40, paddingBottom:"env(safe-area-inset-bottom,0)" }}>
          {allowedViews.slice(0, 5).map(v=>(
            <button key={v.id} onClick={()=>{ setView(v.id); setSidebarOpen(false); }}
              style={{ flex:1, padding:"10px 4px 8px", border:"none", background:"transparent",
                display:"flex", flexDirection:"column", alignItems:"center", gap:2,
                color:view===v.id?C.indigo:C.slate4, cursor:"pointer" }}>
              <span style={{ fontSize:16, position:"relative" }}>
                {v.icon}
                {v.id==="refurb" && pendingConf>0 && (
                  <span style={{ position:"absolute", top:-3, right:-6, background:C.amber, color:C.white, fontSize:8, fontWeight:800,
                    minWidth:14, height:14, borderRadius:20, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 2px" }}>{pendingConf}</span>
                )}
              </span>
              <span style={{ fontSize:9, fontWeight:view===v.id?700:500 }}>{v.label.split(" ")[0]}</span>
            </button>
          ))}
        </nav>
      )}
    </div>
  );
}
