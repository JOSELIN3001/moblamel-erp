// MoblaMel ERP v52 - Logo PWA original, ícono etiquetas en sidebar
// PWA: Para activar el ícono en iPhone, agregar en public/index.html:
// <link rel="apple-touch-icon" href="/logo192.png">
// <meta name="apple-mobile-web-app-capable" content="yes">
// <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
// <meta name="apple-mobile-web-app-title" content="MoblaMel">
// <meta name="theme-color" content="#a0714f"> - Contadores desde 00001, PWA, KPIs compactos, scroll móvil, etiquetas, sin datos ficticios
import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX_LIB from "xlsx";
if (typeof window !== "undefined") { (window as any).XLSX = XLSX_LIB; }
import { AreaChart, Area, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line } from "recharts";
import { createClient } from "@supabase/supabase-js";

// ─── SUPABASE ────────────────────────────────────────────────────────────────
const SUPA_URL = "https://irhharltplzjfxbqxbqq.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlyaGhhcmx0cGx6amZ4YnF4YnFxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4NDAxNjYsImV4cCI6MjA5NTQxNjE2Nn0.5Ji9ORC5Myk-iDOWjGrmt50xzVa1qJwWjhl4h9Gs3U8";
const sb = createClient(SUPA_URL, SUPA_KEY);

// ─── HELPERS SUPABASE ────────────────────────────────────────────────────────
// Carga todos los registros de una tabla (cada fila tiene {id, data})
async function sbLoad(tabla: string): Promise<any[]> {
  const { data, error } = await sb.from(tabla).select("*");
  if (error) { console.error(`sbLoad ${tabla}:`, error); return []; }
  return (data || []).map((r: any) => ({ ...r.data, id: r.id }));
}

// Guarda (upsert) un registro en una tabla
async function sbSave(tabla: string, item: any): Promise<void> {
  const { error } = await sb.from(tabla).upsert({ id: item.id, data: item }, { onConflict: "id" });
  if (error) console.error(`sbSave ${tabla}:`, error);
}

// Guarda múltiples registros de una vez
async function sbSaveAll(tabla: string, items: any[]): Promise<void> {
  if (!items.length) return;
  const rows = items.map(i => ({ id: i.id, data: i }));
  const { error } = await sb.from(tabla).upsert(rows, { onConflict: "id" });
  if (error) console.error(`sbSaveAll ${tabla}:`, error);
}

// Elimina un registro
async function sbDelete(tabla: string, id: string): Promise<void> {
  const { error } = await sb.from(tabla).delete().eq("id", id);
  if (error) console.error(`sbDelete ${tabla}:`, error);
}

// Guarda un valor de configuración (ej: numB, numF)
async function sbSetConfig(key: string, value: any): Promise<void> {
  const { error } = await sb.from("config").upsert({ key, value }, { onConflict: "key" });
  if (error) console.error(`sbSetConfig ${key}:`, error);
}

// Carga un valor de configuración
async function sbGetConfig(key: string): Promise<any> {
  const { data, error } = await sb.from("config").select("value").eq("key", key).single();
  if (error) return null;
  return data?.value ?? null;
}

// ─── LOGOS MOBLAMEL ──────────────────────────────────────────────────────────
const LOGO_SMALL = "data:image/png;base64,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";
const LOGO_MED   = "data:image/png;base64,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";



// ─── PALETA ──────────────────────────────────────────────────────────────────
const C = {
  // Fondos — crema cálido, no blanco puro
  bg:"#f7f4f0", bg2:"#ede8e1", bg3:"#e0d9d0",
  white:"#fdfbf8",
  border:"#e8e0d5", borderD:"#d4c9bb",
  // Acento principal — café dorado del logo
  ac:"#a0714f", acD:"#7d5538", acL:"#c4956a", acBg:"#fdf3eb",
  // Semáforos
  gr:"#2d6a4f", grBg:"#e9f5ee", grL:"#40916c",
  rd:"#b5341c", rdBg:"#fdf0ed", rdL:"#d64123",
  or:"#9c5a1d", orBg:"#fdf4e7",
  bl:"#2c5f8a", blBg:"#eaf2fb",
  pu:"#5c3d8f", puBg:"#f2edfb",
  // Texto
  t1:"#2c2016", t2:"#5c4a38", t3:"#8a7560", t4:"#b5a090",
  // Sombras suaves y cálidas
  sh:"0 1px 4px rgba(100,70,40,0.10)",
  sh2:"0 4px 16px rgba(100,70,40,0.13)",
};

// ─── EXPORT EXCEL (CSV con BOM para Excel en español) ────────────────────────
function exportXLSX(nombre, headers, rows) {
  // Usar SheetJS si está disponible, si no fallback a CSV mejorado
  try {
    const XLSX = window.XLSX;
    if (XLSX) {
      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      // Ancho automático de columnas
      const colWidths = headers.map((h, i) => ({
        wch: Math.max(h.length, ...rows.map(r => String(r[i]||"").length), 10)
      }));
      ws['!cols'] = colWidths;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, nombre);
      XLSX.writeFile(wb, `${nombre}_${new Date().toISOString().slice(0,10)}.xlsx`);
      return;
    }
  } catch(_) {}
  // Fallback: CSV con separadores correctos y BOM
  const esc = v => {
    const s = v === null || v === undefined ? "" : String(v);
    if (s.startsWith("data:image")) return '"[imagen]"';
    return s.includes(",") || s.includes('"') || s.includes("\n") ? `"${s.replace(/"/g,'""')}"` : s;
  };
  const bom  = "\uFEFF";
  const csv  = bom + [headers, ...rows].map(r => r.map(esc).join(",")).join("\r\n");
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = `${nombre}_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── COMISIÓN ────────────────────────────────────────────────────────────────
const COMISION_PCT = 0.02; // 2%

const UBICACIONES = ["Tienda Principal","Tienda Pasaje","Taller/Almacén"];

const TODOS_MODULOS = [
  { id:"dashboard",  label:"🏠 Inicio",          roles:["admin"] },
  { id:"pos",        label:"🛒 Punto de Venta",   roles:["admin","vendedor"] },
  { id:"ventas",     label:"💰 Ventas",           roles:["admin","vendedor"] },
  { id:"caja",       label:"🏦 Cierre de Caja",   roles:["admin","vendedor"] },
  { id:"pedidos",    label:"📬 Pedidos",          roles:["admin","vendedor"] },
  { id:"traslados",  label:"🚚 Traslados",         roles:["admin"] },
  { id:"vendedores", label:"🏆 Vendedores",       roles:["admin"] },
  { id:"clientes",   label:"👥 Clientes",         roles:["admin","vendedor"] },
  { id:"cotiz",      label:"📋 Cotizaciones",     roles:["admin","vendedor"] },
  { id:"inventario", label:"📦 Inventario",       roles:["admin","vendedor","taller"] },
  { id:"kardex",     label:"📒 Kardex",           roles:["admin"] },
  { id:"produccion", label:"🔨 Producción",       roles:["admin","taller"] },
  { id:"compras",    label:"🏭 Compras",          roles:["admin"] },
  { id:"gastos",     label:"💸 Gastos",           roles:["admin"] },
  { id:"finanzas",   label:"📊 Finanzas",         roles:["admin"] },
  { id:"abc",        label:"🏆 Análisis ABC",     roles:["admin"] },
  { id:"prediccion", label:"🔮 Predicción",       roles:["admin"] },
  { id:"usuarios",   label:"👤 Usuarios",         roles:["admin"] },
  { id:"ia",         label:"🤖 Asistente IA",     roles:["admin"] },
  { id:"etiquetas",  label:"🏷️ Etiquetas",       roles:["admin"] },
];

// Módulos por defecto según rol
const modulosPorRol = (rol) => TODOS_MODULOS.filter(m => m.roles.includes(rol)).map(m => m.id);

// ─── USUARIOS (mutable via estado) ───────────────────────────────────────────
const USUARIOS_INIT = [
  { id:"U000", nombre:"Admin",    rol:"admin",    cel:"",          ico:"👑", activo:true, pass:"011207", sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia","etiquetas"] },
  { id:"U001", nombre:"Adolfo",   rol:"admin",    cel:"998649169", ico:"👨‍💼", activo:true, pass:"1234",   sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia","etiquetas"] },
  { id:"U002", nombre:"Joselin",  rol:"admin",    cel:"907569211", ico:"👩‍💼", activo:true, pass:"1234",   sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia","etiquetas"] },
  { id:"U003", nombre:"Rosa",     rol:"vendedor", cel:"985568624", ico:"🌸", activo:true,  pass:"1234",   sueldoSemanal:300, modulos:["pos","ventas","caja","pedidos","clientes","cotiz","inventario"] },
  { id:"U004", nombre:"Maricielo",rol:"vendedor", cel:"914727793", ico:"💐", activo:true,  pass:"1234",   sueldoSemanal:300, modulos:["pos","ventas","caja","pedidos","clientes","cotiz","inventario"] },
  { id:"U005", nombre:"Mia",      rol:"vendedor", cel:"",          ico:"🌺", activo:true,  pass:"1234",   sueldoSemanal:300, modulos:["pos","ventas","caja","pedidos","clientes","cotiz","inventario"] },
];

// ─── HELPERS DE STOCK POR UBICACIÓN + COLOR ──────────────────────────────────
// locs = { "Tienda Principal|Blanco": 2, "Tienda Principal|Nogal": 2, "Taller/Almacén|Blanco": 1, ... }
const locKey    = (ubi, col) => `${ubi}|${col}`;
const getTotalStk  = (p) => p.locs ? Object.values(p.locs).reduce((a,b)=>a+b,0) : (p.stk||0);
const getStkEnUbiCol = (p, ubi, col) => p.locs ? (p.locs[locKey(ubi,col)]||0) : 0;
const getStkEnUbi  = (p, ubi) => p.locs
  ? Object.entries(p.locs).filter(([k])=>k.startsWith(ubi+"|")).reduce((a,[,v])=>a+v,0)
  : (p.ubi===ubi?(p.stk||0):0);
const getLocs = (p) => {
  if (p.locs) return p.locs;
  const res = {};
  const perCol = Math.ceil((p.stk||0)/Math.max(1,(p.cols||[]).length));
  (p.cols||[]).forEach((c,i) => { res[locKey(p.ubi||"Taller/Almacén",c)] = i===0?(p.stk||0)-perCol*((p.cols.length)-1):perCol; });
  return res;
};
// Orígenes con stock (ubicaciones que tienen al menos 1 unidad del producto)
const ubisConStock = (p, col=null) => UBICACIONES.filter(ubi => col
  ? (getLocs(p)[locKey(ubi,col)]||0) > 0
  : Object.entries(getLocs(p)).some(([k,v])=>k.startsWith(ubi+"|")&&v>0)
);
// Colores con stock en una ubicación dada
const colsConStockEnUbi = (p, ubi) => (p.cols||[]).filter(col => (getLocs(p)[locKey(ubi,col)]||0) > 0);
const descontarStk = (p, qty, ubi, col) => {
  const locs = { ...getLocs(p) };
  const k = locKey(ubi,col);
  locs[k] = Math.max(0,(locs[k]||0)-qty);
  return { ...p, locs, stk:Object.values(locs).reduce((a,b)=>a+b,0) };
};
const agregarStk = (p, qty, ubi, col) => {
  const locs = { ...getLocs(p) };
  const k = locKey(ubi,col);
  locs[k] = (locs[k]||0)+qty;
  return { ...p, locs, stk:Object.values(locs).reduce((a,b)=>a+b,0) };
};
const trasladarStk = (p, qty, de, a, col) => {
  const locs = { ...getLocs(p) };
  const kDe = locKey(de,col); const kA = locKey(a,col);
  locs[kDe] = Math.max(0,(locs[kDe]||0)-qty);
  locs[kA]  = (locs[kA]||0)+qty;
  return { ...p, locs, stk:Object.values(locs).reduce((a,b)=>a+b,0) };
};

// ─── DATOS INICIALES ─────────────────────────────────────────────────────────
const PRODS_INIT = [];

const VENTAS_INIT = [];

const GASTOS_INIT = [];

const KARDEX_INIT = [];

const LOTES_INIT = [];

const CLIENTES_INIT = [];

const COTIZ_INIT = [];

const COMPRAS_INIT = [];

const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

const PEDIDOS_INIT = [];
const CAT_C = {Ropero:"#e07b39",Zapatero:"#2b6cb0","Cómoda":"#6b46c1","Mesa de Noche":"#276749",Cabecera:"#c53030"};
const HEX_COLOR = {Blanco:"#f5f0e8",Nogal:"#7c5230","Wengué":"#2e1a0e",Cerezo:"#8b2020",Gris:"#7a7a7a",Beige:"#c8a87a"};
const METODOS_PAGO = ["Efectivo","Yape","Plin","Trans. BCP","Trans. Interbank","Trans. BBVA","POS Tarjeta"];
const HOY = new Date().toISOString().split("T")[0];
const uid = () => Math.random().toString(36).slice(2,8).toUpperCase();
const fmt = n => `S/ ${Number(n).toFixed(2)}`;
const fmtK = n => n >= 1000 ? `S/ ${(n/1000).toFixed(1)}k` : `S/ ${Math.round(n)}`;

// ─── COMPONENTES UI ───────────────────────────────────────────────────────────
function Card({ children, style = {} }) {
  return (
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 10, boxShadow: C.sh, ...style }}>
      {children}
    </div>
  );
}

function Badge({ children, color = C.ac }) {
  return (
    <span style={{ display:"inline-flex", alignItems:"center", padding:"2px 8px", borderRadius:6, fontSize:11, fontWeight:700, color, background: color + "18", whiteSpace:"nowrap" }}>
      {children}
    </span>
  );
}

function Btn({ children, onClick, variant = "ghost", sm, full, disabled }) {
  const styles = {
    primary: { bg: C.ac,  col: "#fff", br: C.ac },
    green:   { bg: C.grL, col: "#fff", br: C.grL },
    red:     { bg: C.rdL, col: "#fff", br: C.rdL },
    ghost:   { bg: "transparent", col: C.t2, br: C.border },
  };
  const st = styles[variant] || styles.ghost;
  return (
    <button onClick={onClick} disabled={disabled} style={{
      padding: sm ? "4px 10px" : "8px 16px", borderRadius: 7,
      border: `1px solid ${st.br}`, background: st.bg, color: st.col,
      fontSize: sm ? 11 : 13, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer",
      fontFamily: "inherit", opacity: disabled ? 0.5 : 1,
      width: full ? "100%" : "auto", transition: "opacity 0.15s", whiteSpace: "nowrap",
    }}>
      {children}
    </button>
  );
}

function Inp({ label, ...props }) {
  return (
    <div>
      {label && <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>{label}</label>}
      <input {...props} style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:7, padding:"8px 11px", color:C.t1, fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", ...(props.style||{}) }} />
    </div>
  );
}

function Sel({ label, children, ...props }) {
  return (
    <div>
      {label && <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>{label}</label>}
      <select {...props} style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:7, padding:"8px 11px", color:C.t1, fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", ...(props.style||{}) }}>
        {children}
      </select>
    </div>
  );
}

function KPI({ icon, label, value, sub, color = C.ac, trend }) {
  return (
    <Card style={{ padding:"16px 18px" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
        <div style={{ width:38, height:38, borderRadius:10, background: color + "18", display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>{icon}</div>
        {trend !== undefined && (
          <span style={{ fontSize:11, fontWeight:700, padding:"2px 7px", borderRadius:99, color: trend >= 0 ? C.grL : C.rdL, background: trend >= 0 ? C.grBg : C.rdBg }}>
            {trend >= 0 ? "↑" : "↓"}{Math.abs(trend)}%
          </span>
        )}
      </div>
      <div style={{ fontSize:22, fontWeight:800, color:C.t1, letterSpacing:"-0.5px" }}>{value}</div>
      <div style={{ fontSize:11, color:C.t3, marginTop:3, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.4px" }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:C.t4, marginTop:4 }}>{sub}</div>}
    </Card>
  );
}

function PageTitle({ title, sub, action }) {
  return (
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
      <div>
        <h1 style={{ fontSize:20, fontWeight:700, color:C.t1, margin:0 }}>{title}</h1>
        {sub && <p style={{ fontSize:13, color:C.t3, margin:"3px 0 0" }}>{sub}</p>}
      </div>
      {action}
    </div>
  );
}

function TH({ children, right }) {
  return <th style={{ padding:"9px 12px", color:C.t3, fontWeight:600, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px", textAlign: right ? "right" : "left", background:C.bg, borderBottom:`1px solid ${C.border}`, whiteSpace:"nowrap" }}>{children}</th>;
}
function TD({ children, right, bold, color, sm }) {
  return <td style={{ padding:"10px 12px", color: color || C.t2, fontWeight: bold ? 700 : 400, fontSize: sm ? 11 : 13, textAlign: right ? "right" : "left", verticalAlign:"middle", borderBottom:`1px solid ${C.bg3}` }}>{children}</td>;
}

function ChartTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:12, boxShadow:C.sh2 }}>
      <div style={{ color:C.t3, marginBottom:4, fontWeight:600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, fontWeight:700 }}>{p.name}: {typeof p.value === "number" ? `S/ ${Math.round(p.value).toLocaleString()}` : p.value}</div>
      ))}
    </div>
  );
}

function Modal({ children, onClose, wide }) {
  return (
    <div onClick={e => { if (e.target === e.currentTarget) onClose(); }}
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.4)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:16 }}>
      <div style={{ background:C.white, borderRadius:14, padding:24, width: wide ? 620 : 520, maxWidth:"100%", maxHeight:"90dvh", overflowY:"auto", WebkitOverflowScrolling:"touch", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
        {children}
      </div>
    </div>
  );
}

function ModalTitle({ children }) {
  return <div style={{ fontSize:17, fontWeight:700, color:C.t1, marginBottom:20, paddingBottom:14, borderBottom:`1px solid ${C.border}` }}>{children}</div>;
}

// ─── COMPONENTE PRINCIPAL ────────────────────────────────────────────────────
export default function MaderERP() {
  const [user,   setUser]   = useState(null);
  const [lU,     setLU]     = useState("");
  const [lP,     setLP]     = useState("");
  const [lErr,   setLErr]   = useState("");
  const [mod,    setMod]    = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [prods,  setProds]  = useState(PRODS_INIT);
  const [ventas, setVentas] = useState(VENTAS_INIT);
  const [gastos, setGastos] = useState(GASTOS_INIT);
  const [cierres, setCierres] = useState([]); // { id, f, vendedor, ventas[], anticipos[], total, totalAnticipo, metodos:{}, hora }
  const [kardex, setKardex] = useState(KARDEX_INIT);
  const [lotes,  setLotes]  = useState(LOTES_INIT);
  const [clientes,setClientes] = useState(CLIENTES_INIT);
  const [cotizs, setCotizs] = useState(COTIZ_INIT);
  const [compras,setCompras]= useState(COMPRAS_INIT);
  const [usuarios,setUsuarios] = useState(USUARIOS_INIT);
  const [traslados, setTraslados] = useState([]);
  const [trasladoForm, setTrasladoForm] = useState(null);
  const [pedidos, setPedidos] = useState(PEDIDOS_INIT);
  const [pedidoForm, setPedidoForm] = useState(null);
  const [abonoModal, setAbonoModal] = useState(null);
  const [cotizPagoModal, setCotizPagoModal] = useState(null); // {pd, monto}
  const [modal,  setModal]  = useState(null);  const [toast,  setToast]  = useState(null);
  const [numB,   setNumB]   = useState(1);
  const [numF,   setNumF]   = useState(1);

  // POS
  const [posBusq,   setPosBusq]   = useState("");
  const [posCat,    setPosCat]    = useState("Todos");
  const [posCarrito,setPosCarrito]= useState([]);
  const [posCliente,setPosCliente]= useState("");
  const [posDni,    setPosDni]    = useState("");
  const [dniLoad,   setDniLoad]   = useState(false);
  const [posComp,   setPosComp]   = useState("Nota de Venta");
  const [posPagos,  setPosPagos]  = useState([{met:"Efectivo",monto:""}]);
  const [posDesc,   setPosDesc]   = useState("");
  const [scanMode,  setScanMode]  = useState(false);
  const [scanVal,   setScanVal]   = useState("");
  const [voucher,   setVoucher]   = useState(null);
  const [voucherFmt, setVoucherFmt] = useState("ticket");
  const [posTab,    setPosTab]    = useState("catalogo"); // "catalogo" | "cobro"
  const posFase = posTab;

  // Filtros existentes
  const [abcFiltro, setAbcFiltro] = useState("Todos");
  const [kFiltProd, setKFiltProd] = useState("");
  const [kFiltCol,  setKFiltCol]  = useState("");
  const [kFiltTipo, setKFiltTipo] = useState("Todos");
  const [kFiltBusq, setKFiltBusq] = useState("");
  const [kFiltDesde, setKFiltDesde] = useState("");
  const [kFiltHasta, setKFiltHasta] = useState("");
  const [kFiltUbi,   setKFiltUbi]   = useState("");

  // Filtros de fecha para módulos
  const [prodFiltDesde, setProdFiltDesde] = useState("");
  const [prodFiltHasta, setProdFiltHasta] = useState("");
  const [gastoFiltDesde,setGastoFiltDesde] = useState("");
  const [gastoFiltHasta,setGastoFiltHasta] = useState("");
  const [gastoFiltCat,  setGastoFiltCat]  = useState("Todas");
  const [compraFiltDesde,setCompraFiltDesde] = useState("");
  const [compraFiltHasta,setCompraFiltHasta] = useState("");

  // Vendedores
  const [vendFiltroDesde, setVendFiltroDesde] = useState("2025-05-01");
  const [vendFiltroHasta, setVendFiltroHasta] = useState("2025-05-31");
  const [vendBusq, setVendBusq] = useState("");
  const [comisionPct,     setComisionPct]     = useState(2); // porcentaje editable

  // Cotizaciones
  const [cotizForm, setCotizForm] = useState(null);
  // Etiquetas
  const [etProd, setEtProd] = useState("");
  const [etCol, setEtCol] = useState("");
  const [etQty, setEtQty] = useState(1);
  const [etPrecio, setEtPrecio] = useState(true);
  const [etLogo, setEtLogo] = useState(true);
  const [etNombre, setEtNombre] = useState(true);
  const [etTamaño, setEtTamaño] = useState("ticket58");

  // Compras
  const [compraForm, setCompraForm] = useState(null);

  // Inventario
  const [invBusq, setInvBusq] = useState("");
  const [invCat,  setInvCat]  = useState("Todos");
  const [invEst,  setInvEst]  = useState("Todos");
  const [invUbi,  setInvUbi]  = useState("Todas");
  const [prodEdit,setProdEdit] = useState(null);
  const [invScan,    setInvScan]    = useState(false);
  const [invScanVal, setInvScanVal] = useState("");
  const [camScan,    setCamScan]    = useState(false);
  const [invVista,   setInvVista]   = useState("tabla"); // "tabla" | "cards"
  const [catSel,     setCatSel]     = useState(null);    // dashboard: categoria seleccionada
  const invScanRef = useRef(null);

  // Finanzas
  const [finMes,  setFinMes]  = useState("2025-05");

  // Ventas filtros
  const [vFiltDesde, setVFiltDesde] = useState("");
  const [vFiltHasta, setVFiltHasta] = useState("");
  const [vFiltVend,  setVFiltVend]  = useState("Todos");
  const [vFiltComp,  setVFiltComp]  = useState("Todos");
  const [vFiltBusq,  setVFiltBusq]  = useState("");

  // IA — modo informe
  const [iaInput,   setIaInput]   = useState("");
  const [iaGrafico, setIaGrafico] = useState(null); // { tipo, datos, titulo }
  const [iaMsgs,    setIaMsgs]    = useState([{rol:"ai",txt:"¡Hola! 👋 Soy el asistente de MoblaMel. Analizo tus ventas, stock, gastos y pedidos en tiempo real.\n\nPuedes preguntarme cosas como:\n• ¿Qué vendedor vendió más este mes?\n• ¿Cuánto vendí este mes?\n• ¿Qué producto tiene mejor margen?\n• ¿Cuánto debo pagar de comisiones?\n• Dame un resumen del negocio\n\nO usa los informes rápidos del panel derecho 👉"}]);
  const [iaLoad,    setIaLoad]    = useState(false);

  const scanRef = useRef(null);
  useEffect(() => { if (scanMode && scanRef.current) scanRef.current.focus(); }, [scanMode]);

  // ─── PWA META TAGS (iPhone home screen) ─────────────────────────────────
  useEffect(() => {
    // apple-touch-icon — ícono cuadrado 180×180 con la M de MoblaMel bien encuadrada
    const PWA_ICON_URL = "data:image/png;base64,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";
    let link = document.querySelector("link[rel='apple-touch-icon']") as HTMLLinkElement;
    if (!link) { link = document.createElement("link") as HTMLLinkElement; link.rel = "apple-touch-icon"; document.head.appendChild(link); }
    link.href = PWA_ICON_URL;
    // theme-color
    let meta = document.querySelector("meta[name='theme-color']") as HTMLMetaElement;
    if (!meta) { meta = document.createElement("meta") as HTMLMetaElement; meta.name = "theme-color"; document.head.appendChild(meta); }
    meta.content = "#a0714f";
    // apple PWA
    const setMeta = (name:string, content:string) => {
      let m = document.querySelector(`meta[name='${name}']`) as HTMLMetaElement;
      if (!m) { m = document.createElement("meta") as HTMLMetaElement; m.name = name; document.head.appendChild(m); }
      m.content = content;
    };
    setMeta("apple-mobile-web-app-capable", "yes");
    setMeta("apple-mobile-web-app-status-bar-style", "black-translucent");
    setMeta("apple-mobile-web-app-title", "MoblaMel");
    // viewport-fit=cover — necesario para safe-area-inset en iPhone X+
    let vp = document.querySelector("meta[name='viewport']") as HTMLMetaElement;
    if (!vp) { vp = document.createElement("meta") as HTMLMetaElement; vp.name = "viewport"; document.head.appendChild(vp); }
    if (!vp.content.includes("viewport-fit")) {
      vp.content = (vp.content || "width=device-width,initial-scale=1") + ",viewport-fit=cover";
    }
    document.title = "MoblaMel ERP";
  }, []);

  // ─── SUPABASE: CARGA INICIAL ──────────────────────────────────────────────
  const [dbReady, setDbReady] = useState(false);
  const [dbError, setDbError] = useState("");

  useEffect(() => {
    let cancelled = false;
    async function cargarDatos() {
      try {
        const [
          dbProds, dbVentas, dbGastos, dbCierres, dbKardex, dbLotes,
          dbClientes, dbCotizs, dbCompras, dbUsuarios, dbTraslados, dbPedidos,
          cfgNumB, cfgNumF
        ] = await Promise.all([
          sbLoad("productos"), sbLoad("ventas"), sbLoad("gastos"),
          sbLoad("cierres"), sbLoad("kardex"), sbLoad("lotes"),
          sbLoad("clientes"), sbLoad("cotizaciones"), sbLoad("compras"),
          sbLoad("usuarios"), sbLoad("traslados"), sbLoad("pedidos"),
          sbGetConfig("numB"), sbGetConfig("numF")
        ]);
        if (cancelled) return;

        // Si hay datos en Supabase, usarlos; si no, usar datos iniciales y sembrarlos
        if (dbProds.length)    setProds(dbProds);       else sbSaveAll("productos",   PRODS_INIT);
        if (dbVentas.length)   setVentas(dbVentas);     else sbSaveAll("ventas",      VENTAS_INIT);
        if (dbGastos.length)   setGastos(dbGastos);     else sbSaveAll("gastos",      GASTOS_INIT);
        if (dbCierres.length)  setCierres(dbCierres);
        if (dbKardex.length)   setKardex(dbKardex);     else sbSaveAll("kardex",      KARDEX_INIT);
        if (dbLotes.length)    setLotes(dbLotes);       else sbSaveAll("lotes",       LOTES_INIT);
        if (dbClientes.length) setClientes(dbClientes); else sbSaveAll("clientes",    CLIENTES_INIT);
        if (dbCotizs.length)   setCotizs(dbCotizs);    else sbSaveAll("cotizaciones", COTIZ_INIT);
        if (dbCompras.length)  setCompras(dbCompras);   else sbSaveAll("compras",     COMPRAS_INIT);
        if (dbUsuarios.length) setUsuarios(dbUsuarios); else sbSaveAll("usuarios",    USUARIOS_INIT);
        if (dbTraslados.length) setTraslados(dbTraslados);
        if (dbPedidos.length)  setPedidos(dbPedidos);   else sbSaveAll("pedidos",     PEDIDOS_INIT);
        if (cfgNumB !== null)  setNumB(cfgNumB);
        if (cfgNumF !== null)  setNumF(cfgNumF);

        setDbReady(true);
      } catch (e: any) {
        if (!cancelled) setDbError("No se pudo conectar a la base de datos. Usando datos locales.");
        setDbReady(true);
      }
    }
    cargarDatos();
    return () => { cancelled = true; };
  }, []);

  const showToast = (msg, t = "ok") => { setToast({ msg, t }); setTimeout(() => setToast(null), 2800); };

  // ─── LOGIN ──────────────────────────────────────────────────────────────────
  const login = () => {
    const u = usuarios.find(x => x.activo && x.nombre.toLowerCase() === lU.toLowerCase() && x.pass === lP);
    if (u) { setUser(u); setLErr(""); } else setLErr("Usuario o contraseña incorrectos");
  };

  // ─── KPIs ────────────────────────────────────────────────────────────────────
  const kpi = useMemo(() => {
    const mes = ventas.filter(v => v.f.startsWith("2025-05"));
    const ing = mes.reduce((a, v) => a + v.tot, 0);
    const cmv = mes.reduce((a, v) => a + v.items.reduce((b, i) => b + i.c * i.q, 0), 0);
    const gm  = gastos.filter(g => g.f.startsWith("2025-05")).reduce((a, g) => a + g.monto, 0);
    const ut  = ing - cmv - gm;
    const mg  = ing > 0 ? Math.round((ut / ing) * 100) : 0;
    const crit= prods.filter(p => getTotalStk(p) <= p.min).length;
    const vi  = prods.reduce((a, p) => a + p.stk * p.c, 0);
    return { ing, cmv, gm, ut, mg, crit, vi, tickets: mes.length, tp: mes.length ? ing / mes.length : 0 };
  }, [ventas, gastos, prods]);

  // ─── ABC (basado en ventas REALES registradas) ───────────────────────────────
  const abc = useMemo(() => {
    const gastoLocal   = gastos.filter(g => ["Alquiler","Servicios taller"].includes(g.cat)).reduce((a,g) => a + g.monto, 0);
    const costoPorProd = prods.length > 0 ? gastoLocal / prods.length : 0;

    const data = prods.map(p => {
      // Ventas reales del producto en el sistema
      const ventasProd = ventas.filter(v => v.items.some(i => i.id === p.id));
      const unids      = ventasProd.reduce((a,v) => a + v.items.filter(i => i.id === p.id).reduce((b,i) => b + i.q, 0), 0);
      const ing        = ventasProd.reduce((a,v) => a + v.items.filter(i => i.id === p.id).reduce((b,i) => b + i.q * i.p, 0), 0);
      const tieneVentas= ventasProd.length > 0;
      // Promedio últimos 3 meses (estimado del historial v[] si no hay suficiente real)
      const avgReal    = unids > 0 ? Math.round(unids / Math.max(1, new Set(ventasProd.map(v => v.f.slice(0,7))).size)) : 0;
      const avg        = tieneVentas ? avgReal : Math.round(p.v.slice(-3).reduce((a,b)=>a+b,0)/3);
      const margen     = p.p > 0 ? Math.round(((p.p - p.c) / p.p) * 100) : 0;
      const ingMes     = avg * p.p;
      const costMes    = avg * p.c + costoPorProd;
      const ganReal    = ingMes - costMes;
      const obsoleto   = avg < 1 && p.stk > 0;
      const capInmov   = p.stk * p.c;
      return { ...p, unids, ing, avg, margen, ingMes, costMes, ganReal, obsoleto, capInmov, costoPorProd, tieneVentas };
    }).sort((a,b) => b.ing - a.ing);

    // Solo clasificar productos CON ventas reales en el ABC A/B/C
    const conVentas = data.filter(d => d.tieneVentas);
    const sinVentas = data.filter(d => !d.tieneVentas);
    const totalIng  = conVentas.reduce((a,d) => a + d.ing, 0);
    let acc = 0;
    const clasificados = conVentas.map(d => {
      acc += d.ing;
      const pct = totalIng > 0 ? (acc / totalIng) * 100 : 100;
      const pctIng = totalIng > 0 ? Math.round((d.ing / totalIng) * 100) : 0;
      return { ...d, clase: pct <= 70 ? "A" : pct <= 90 ? "B" : "C", pctIng };
    });
    // Sin ventas reales → "Sin clasificación"
    const sinClasif = sinVentas.map(d => ({ ...d, clase:"—", pctIng:0 }));
    return [...clasificados, ...sinClasif];
  }, [prods, gastos, ventas]);

  const abcFiltrado = useMemo(() => {
    if (abcFiltro === "Obsoletos") return abc.filter(p => p.obsoleto);
    if (abcFiltro === "Todos") return abc;
    if (abcFiltro === "—") return abc.filter(p => p.clase === "—");
    return abc.filter(p => p.clase === abcFiltro);
  }, [abc, abcFiltro]);

  // ─── PREDICCIÓN ─────────────────────────────────────────────────────────────
  const pred = useMemo(() => {
    return prods.map(p => {
      const w = [0.1, 0.15, 0.2, 0.25, 0.3];
      const pred = Math.round(p.v.slice(-5).reduce((a, v, i) => a + v * w[i], 0) * 1.05);
      const dias = pred > 0 ? Math.round((p.stk / pred) * 30) : 999;
      const deficit = Math.max(0, pred - p.stk);
      return { ...p, pred, dias, deficit };
    }).sort((a,b) => a.dias - b.dias);
  }, [prods]);

  // ─── POS CÁLCULOS ───────────────────────────────────────────────────────────
  const posSub    = posCarrito.reduce((a,i) => a + i.q * i.pr, 0);
  const posBase   = posSub - (parseFloat(posDesc) || 0);
  // POS: la vendedora ingresa el precio del mueble (ej S/300)
  // El sistema calcula que el cliente paga S/315 en el datáfono (300 × 1.05)
  // La tienda recibe S/300, Izipay se queda con S/15 (5%)
  const montoPOS        = posPagos.filter(p => p.met === "POS Tarjeta").reduce((a,p) => a + (parseFloat(p.monto)||0), 0);
  const tienePOS        = posPagos.some(p => p.met === "POS Tarjeta");
  const recPOS          = tienePOS ? +(montoPOS * 0.05).toFixed(2) : 0;
  const montoEnDatafono = +(montoPOS * 1.05).toFixed(2); // cliente paga esto en el datáfono
  const montoNoPOS      = posPagos.filter(p=>p.met!=="POS Tarjeta").reduce((a,p)=>a+(parseFloat(p.monto)||0),0);
  // El total del carrito NO cambia — el recargo lo asume el cliente encima
  // Para cobrar exacto: montoNoPOS + montoPOS debe = posBase
  const posTotal  = posBase; // el total del carrito no cambia
  const posPagado = +(montoNoPOS + montoPOS).toFixed(2); // lo que ingresa la vendedora
  const posFalta  = +Math.max(0, posBase - posPagado).toFixed(2);
  const posVuelto = +Math.max(0, posPagado - posBase).toFixed(2);
  const basePOS   = montoPOS;

  const posFiltrados = useMemo(() => prods.filter(p => {
    const catOk = posCat === "Todos" || p.cat === posCat;
    const bOk   = !posBusq || p.n.toLowerCase().includes(posBusq.toLowerCase()) || p.cols.some(c => c.toLowerCase().includes(posBusq.toLowerCase()));
    return catOk && bOk;
  }), [prods, posCat, posBusq]);

  const posAgregar = (prod, col) => {
    setPosCarrito(prev => {
      const key = prod.id + "-" + col;
      const ex  = prev.find(i => i.key === key);
      if (ex) {
        if (ex.q >= prod.stk) { showToast("Stock máximo", "err"); return prev; }
        return prev.map(i => i.key === key ? { ...i, q: i.q + 1 } : i);
      }
      // maxStk = stock disponible de ese color específico en todas las ubicaciones
      const stkDelColor = Object.entries(getLocs(prod)).filter(([k])=>k.includes(`|${col}`)).reduce((a,[,v])=>a+v,0);
      return [...prev, { key, pid:prod.id, n:prod.n, ico:prod.ico, col, q:1, prLista:prod.p, pr:prod.p, costo:prod.c, maxStk:stkDelColor }];
    });
  };

  const updPago = (i, k, v) => {
    setPosPagos(prev => {
      const next = prev.map((x, j) => j === i ? { ...x, [k]: v } : x);
      if (k === "met" && v === "POS Tarjeta") {
        const otros  = next.filter((_, j) => j !== i).reduce((a, p) => a + (parseFloat(p.monto)||0), 0);
        const rest   = +Math.max(0, posBase - otros).toFixed(2);
        return next.map((x, j) => j === i ? { ...x, monto: String(+(rest * 1.05).toFixed(2)) } : x);
      }
      if (k === "monto") {
        const posIdx = next.findIndex((x, j) => j !== i && x.met === "POS Tarjeta");
        if (posIdx >= 0) {
          const otros = next.filter((_, j) => j !== posIdx).reduce((a, p) => a + (parseFloat(p.monto)||0), 0);
          const rest  = +Math.max(0, posBase - otros).toFixed(2);
          return next.map((x, j) => j === posIdx ? { ...x, monto: String(+(rest * 1.05).toFixed(2)) } : x);
        }
      }
      return next;
    });
  };

  const buscarDni = async (dni) => {
    if (dni.length !== 8) { showToast("DNI debe tener 8 dígitos", "err"); return; }
    setDniLoad(true);
    try {
      const res  = await fetch(`https://api.apis.net.pe/v2/reniec/dni?numero=${dni}`, { headers: { "Authorization": "Bearer apis-token-13752.dFVFbEFJKBuXM6QR6IA6cVb1STKBB3GF" } });
      const data = await res.json();
      if (data.nombreCompleto) { setPosCliente(data.nombreCompleto); showToast(`✓ ${data.nombreCompleto}`); }
      else showToast("DNI no encontrado", "err");
    } catch { showToast("No se pudo consultar RENIEC", "err"); }
    setDniLoad(false);
  };

  const posCobrar = () => {
    if (!posCarrito.length) { showToast("Carrito vacío", "err"); return; }
    if (posFalta > 0.01) { showToast(`Falta S/ ${posFalta.toFixed(2)}`, "err"); return; }
    const pfx = posComp === "Factura" ? "F001" : posComp === "Boleta" ? "B001" : "NV";
    const num = `${pfx}-${String(numB).padStart(5, "0")}`;
    const nv  = { id:"V"+uid(), f:HOY, cli: posCliente || "Consumidor final", vend: user?.nombre || "Joselin",
      items: posCarrito.map(i => ({ id:i.pid, n:i.n, col:i.col, q:i.q, p:i.pr, c:i.costo })),
      tot: posTotal, mp: posPagos.map(p => p.met).join("+"), comp: posComp, num };
    setVentas(v => [nv, ...v]);
    sbSave("ventas", nv);
    // Descontar stock por color+ubicación correctamente
    setProds(ps => {
      const updated = ps.map(p => {
        const itemsDelProd = posCarrito.filter(i => i.pid === p.id);
        if (!itemsDelProd.length) return p;
        let np = { ...p };
        itemsDelProd.forEach(item => {
          const ubiOrigen = UBICACIONES.find(u => (getLocs(np)[locKey(u, item.col)]||0) >= item.q)
                         || UBICACIONES.find(u => (getLocs(np)[locKey(u, item.col)]||0) > 0)
                         || UBICACIONES[0];
          np = descontarStk(np, item.q, ubiOrigen, item.col);
        });
        return np;
      });
      updated.filter(p => posCarrito.some(i => i.pid === p.id)).forEach(p => sbSave("productos", p));
      return updated;
    });
    const nk = posCarrito.map(i => ({ id:"K"+uid(), f:HOY, pid:i.pid, prod:i.n, col:i.col, tipo:"Venta", desc:`Venta ${num}`, ent:0, sal:i.q, saldo:(prods.find(p=>p.id===i.pid)?.stk||0)-i.q, costo:i.costo }));
    setKardex(k => { nk.forEach(k => sbSave("kardex", k)); return [...nk, ...k]; });
    if (posComp === "Factura") { setNumF(n => { const nv2 = n+1; sbSetConfig("numF", nv2); return nv2; }); }
    else { setNumB(n => { const nv2 = n+1; sbSetConfig("numB", nv2); return nv2; }); }
    const vData = { items:[...posCarrito], subtotal:posSub, desc:parseFloat(posDesc)||0, recPOS, total:posTotal, pagos:[...posPagos], vuelto:posVuelto, cliente:posCliente||"Consumidor final", vend:user?.nombre||"Joselin", comp:posComp, num, fecha:new Date().toLocaleString("es-PE") };
    setPosCarrito([]); setPosPagos([{met:"Efectivo",monto:""}]); setPosCliente(""); setPosDni(""); setPosDesc("");
    setPosTab("catalogo");
    setVoucher(vData);
    showToast(`✓ ${num} registrada`);
  };

  // ─── ASISTENTE IA LOCAL (sin API externa, analiza datos reales) ─────────────
  const enviarIA = () => {
    if (!iaInput.trim() || iaLoad) return;
    const preg = iaInput.trim().toLowerCase();
    setIaMsgs(m => [...m, { rol:"user", txt:iaInput.trim() }]);
    setIaInput(""); setIaLoad(true);

    setTimeout(() => {
      const ventasRealesIA = ventas.filter(v => !v.esAnticipo);
      const mesActual = HOY.slice(0,7);
      const ventasMes = ventasRealesIA.filter(v => v.f.startsWith(mesActual));
      const ingMes    = ventasMes.reduce((a,v)=>a+v.tot,0);
      const gastosMes = gastos.filter(g => g.f.startsWith(mesActual)).reduce((a,g)=>a+g.monto,0);
      const utilMes   = ingMes - gastosMes;

      let resp = "";

      // ── VENTAS ──────────────────────────────────────────────────────────────
      if (preg.match(/vend(edor|ió|io)|quién|quien|más vendió|top/)) {
        const rank = usuarios.filter(u=>u.rol!=="taller").map(u => {
          const tot = ventasMes.filter(v=>v.vend===u.nombre).reduce((a,v)=>a+v.tot,0);
          const qty = ventasMes.filter(v=>v.vend===u.nombre).length;
          return { nombre:u.nombre, tot, qty };
        }).sort((a,b)=>b.tot-a.tot).filter(u=>u.tot>0);
        if (rank.length === 0) {
          resp = `📊 Sin ventas registradas este mes aún.`;
        } else {
          resp = `🏆 **Ranking de ventas — ${mesActual}:**\n\n` +
            rank.map((v,i) => `${["🥇","🥈","🥉","4️⃣","5️⃣"][i]||"•"} **${v.nombre}**: ${fmt(v.tot)} (${v.qty} ventas)`).join("\n") +
            `\n\n💰 Total del equipo: **${fmt(ingMes)}**`;
        }
      }
      // ── INGRESOS / VENTAS DEL MES ────────────────────────────────────────────
      else if (preg.match(/ingres|cuánto (vendí|vendi|hice|gané|gane)|total.*mes|mes.*total/)) {
        resp = `💰 **Resumen de ingresos — ${mesActual}:**\n\n` +
          `• Ventas cerradas: **${fmt(ingMes)}** (${ventasMes.length} ventas)\n` +
          `• Gastos del mes: **${fmt(gastosMes)}**\n` +
          `• Utilidad estimada: **${fmt(utilMes)}** (${ingMes>0?Math.round((utilMes/ingMes)*100):0}% margen)\n\n` +
          `📅 Ticket promedio: ${fmt(ventasMes.length?ingMes/ventasMes.length:0)}`;
      }
      // ── STOCK / INVENTARIO ───────────────────────────────────────────────────
      else if (preg.match(/stock|inventario|queda|quedan|agot|falt/)) {
        const criticos = prods.filter(p => getTotalStk(p) <= p.min);
        const sinStock = prods.filter(p => getTotalStk(p) === 0);
        const ok       = prods.filter(p => getTotalStk(p) > p.min);
        resp = `📦 **Estado del inventario:**\n\n` +
          `✅ Con stock OK: **${ok.length}** productos\n` +
          `⚠️ Stock crítico: **${criticos.length}** productos\n` +
          `❌ Sin stock: **${sinStock.length}** productos\n\n`;
        if (criticos.length > 0) {
          resp += `**Productos a reabastecer:**\n` +
            criticos.map(p => `• ${p.ico} ${p.n}: ${getTotalStk(p)} uds (mín: ${p.min})`).join("\n");
        }
      }
      // ── PRODUCTO MÁS RENTABLE ────────────────────────────────────────────────
      else if (preg.match(/rentable|margen|ganancia|utilidad.*producto|producto.*utilidad/)) {
        const conMargen = prods.map(p => ({
          n:p.n, ico:p.ico,
          margen: p.p>0?Math.round(((p.p-p.c)/p.p)*100):0,
          ganUnit: p.p-p.c
        })).sort((a,b)=>b.margen-a.margen);
        resp = `💎 **Productos más rentables:**\n\n` +
          conMargen.slice(0,5).map((p,i) => `${["🥇","🥈","🥉","4️⃣","5️⃣"][i]} **${p.ico} ${p.n}**: ${p.margen}% margen (ganas S/${p.ganUnit} por unidad)`).join("\n");
      }
      // ── COMISIONES ───────────────────────────────────────────────────────────
      else if (preg.match(/comisi[oó]n|comision|le debo|pagar.*vend/)) {
        const pct = (comisionPct||2)/100;
        const coms = usuarios.filter(u=>u.rol==="vendedor").map(u => {
          const tot = ventasMes.filter(v=>v.vend===u.nombre).reduce((a,v)=>a+v.tot,0);
          return { nombre:u.nombre, comision: +(tot*pct).toFixed(2), ventas:tot };
        }).filter(u=>u.ventas>0).sort((a,b)=>b.comision-a.comision);
        if (coms.length === 0) {
          resp = `💳 Sin ventas registradas para calcular comisiones este mes.`;
        } else {
          resp = `💳 **Comisiones del mes (${comisionPct}%):**\n\n` +
            coms.map(c => `• **${c.nombre}**: ${fmt(c.comision)} (sobre ${fmt(c.ventas)} en ventas)`).join("\n") +
            `\n\n💰 Total a pagar: **${fmt(coms.reduce((a,c)=>a+c.comision,0))}**`;
        }
      }
      // ── PEDIDOS PENDIENTES ───────────────────────────────────────────────────
      else if (preg.match(/pedido|pendiente|entreg|saldo/)) {
        const pendientes = pedidos.filter(p => p.est !== "Entregado" && p.est !== "Cancelado");
        const saldoTotal = pendientes.reduce((a,p)=>a+(p.saldoPendiente||0),0);
        resp = `📬 **Pedidos en curso:**\n\n` +
          `• Total pedidos activos: **${pendientes.length}**\n` +
          `• Saldo pendiente total: **${fmt(saldoTotal)}** por cobrar\n\n`;
        if (pendientes.length > 0) {
          resp += `**Próximos a vencer:**\n` +
            pendientes.filter(p=>p.fEnt).sort((a,b)=>a.fEnt>b.fEnt?1:-1).slice(0,3)
              .map(p=>`• ${p.cli} — ${p.prod}: S/${p.saldoPendiente||0} pendiente (entrega: ${p.fEnt})`).join("\n");
        }
      }
      // ── GASTOS ───────────────────────────────────────────────────────────────
      else if (preg.match(/gasto|costo|alquiler|sueldo|cuánto.*gast/)) {
        const fijos    = gastos.filter(g=>g.esFijo&&g.f.startsWith(mesActual)).reduce((a,g)=>a+g.monto,0);
        const variables= gastos.filter(g=>!g.esFijo&&g.f.startsWith(mesActual)).reduce((a,g)=>a+g.monto,0);
        const porCat   = {};
        gastos.filter(g=>g.f.startsWith(mesActual)).forEach(g => { porCat[g.cat]=(porCat[g.cat]||0)+g.monto; });
        resp = `💸 **Gastos del mes ${mesActual}:**\n\n` +
          `• Gastos fijos: **${fmt(fijos)}** (alquileres, sueldos)\n` +
          `• Gastos variables: **${fmt(variables)}** (materiales, comisiones)\n` +
          `• **Total: ${fmt(fijos+variables)}**\n\n` +
          `**Por categoría:**\n` +
          Object.entries(porCat).sort((a,b)=>b[1]-a[1]).map(([cat,monto])=>`• ${cat}: ${fmt(monto)}`).join("\n");
      }
      // ── PUNTO DE EQUILIBRIO ──────────────────────────────────────────────────
      else if (preg.match(/equilibrio|cuánto.*vender|mínimo.*vender|punto de/)) {
        const gastosFijos = gastos.filter(g=>g.esFijo&&g.f.startsWith(mesActual)).reduce((a,g)=>a+g.monto,0);
        const margenProm  = ingMes>0?(ingMes-gastos.filter(g=>!g.esFijo&&g.f.startsWith(mesActual)).reduce((a,g)=>a+g.monto,0))/ingMes:0.35;
        const pe = margenProm>0?Math.round(gastosFijos/margenProm):0;
        const peDia = Math.round(pe/26);
        resp = `⚖️ **Punto de equilibrio:**\n\n` +
          `• Gastos fijos del mes: **${fmt(gastosFijos)}**\n` +
          `• Margen promedio: **${Math.round(margenProm*100)}%**\n` +
          `• Debes vender mínimo: **${fmt(pe)}/mes**\n` +
          `• Eso equivale a: **${fmt(peDia)}/día** (26 días hábiles)\n\n` +
          `📊 Este mes llevas **${fmt(ingMes)}** → ${ingMes>=pe?"✅ Ya superaste el punto de equilibrio":"⚠️ Aún no alcanzas el punto de equilibrio"}`;
      }
      // ── RESUMEN GENERAL ──────────────────────────────────────────────────────
      else if (preg.match(/resumen|cómo.*negocio|como.*negocio|estado.*negocio|negocio/)) {
        const mejorProd = [...prods].sort((a,b)=>{
          const vA=ventasMes.filter(v=>v.items.some(i=>i.id===a.id)).reduce((s,v)=>s+v.tot,0);
          const vB=ventasMes.filter(v=>v.items.some(i=>i.id===b.id)).reduce((s,v)=>s+v.tot,0);
          return vB-vA;
        })[0];
        const mejorVend = usuarios.filter(u=>u.rol!=="taller").map(u=>({
          n:u.nombre, t:ventasMes.filter(v=>v.vend===u.nombre).reduce((a,v)=>a+v.tot,0)
        })).sort((a,b)=>b.t-a.t)[0];
        resp = `📊 **Resumen del negocio — ${mesActual}:**\n\n` +
          `💰 Ingresos: **${fmt(ingMes)}** | Gastos: **${fmt(gastosMes)}** | Utilidad: **${fmt(utilMes)}**\n\n` +
          `📦 Inventario: ${prods.length} productos | ${prods.filter(p=>getTotalStk(p)<=p.min).length} en stock crítico\n\n` +
          `🏆 Mejor vendedor: **${mejorVend?.n||"—"}** (${fmt(mejorVend?.t||0)})\n\n` +
          `📬 Pedidos activos: **${pedidos.filter(p=>p.est!=="Entregado"&&p.est!=="Cancelado").length}**\n\n` +
          `${utilMes>0?"✅ El negocio está generando utilidad este mes.":"⚠️ Los gastos superan los ingresos este mes."}`;
      }
      // ── ABC ───────────────────────────────────────────────────────────────────
      else if (preg.match(/abc|clase|fabricar|producir|priorizar/)) {
        const claseA = abc.filter(p=>p.clase==="A");
        const claseB = abc.filter(p=>p.clase==="B");
        const claseC = abc.filter(p=>p.clase==="C"||p.clase==="—");
        resp = `🏆 **Análisis ABC de productos:**\n\n` +
          `⭐ **Clase A** (priorizar): ${claseA.map(p=>p.n).join(", ")||"—"}\n` +
          `✅ **Clase B** (mantener): ${claseB.map(p=>p.n).join(", ")||"—"}\n` +
          `📉 **Clase C / sin ventas**: ${claseC.map(p=>p.n).join(", ")||"—"}\n\n` +
          `💡 Enfócate en tener siempre stock de los productos Clase A — son los que más ingresos generan.`;
      }
      // ── CLIENTES ─────────────────────────────────────────────────────────────
      else if (preg.match(/cliente|comprador|quién.*compr/)) {
        const compras = {};
        ventasRealesIA.forEach(v => { if(v.cli) compras[v.cli]=(compras[v.cli]||0)+v.tot; });
        const top = Object.entries(compras).sort((a,b)=>b[1]-a[1]).slice(0,5);
        resp = `👥 **Top clientes por compras:**\n\n` +
          (top.length>0 ? top.map(([n,t],i)=>`${["🥇","🥈","🥉","4️⃣","5️⃣"][i]} **${n}**: ${fmt(t)}`).join("\n")
          : "Sin compras registradas con nombre de cliente.");
      }
      // ── AYUDA ────────────────────────────────────────────────────────────────
      else if (preg.match(/qué puedes|que puedes|ayuda|help|cómo funciona/)) {
        resp = `🤖 **Puedo responder preguntas sobre:**\n\n` +
          `💰 Ventas e ingresos del mes\n` +
          `🏆 Ranking de vendedores y comisiones\n` +
          `📦 Stock e inventario crítico\n` +
          `💎 Productos más rentables (márgenes)\n` +
          `📬 Pedidos pendientes y saldos\n` +
          `💸 Gastos del mes por categoría\n` +
          `⚖️ Punto de equilibrio\n` +
          `📊 Resumen general del negocio\n` +
          `🏆 Análisis ABC de productos\n` +
          `👥 Top clientes\n\n` +
          `💡 También puedes usar los **Informes rápidos** del panel derecho.`;
      }
      // ── RESPUESTA POR DEFECTO ────────────────────────────────────────────────
      else {
        // Intento genérico con datos
        resp = `🤔 No entendí exactamente tu pregunta, pero aquí un resumen rápido:\n\n` +
          `💰 Ingresos del mes: **${fmt(ingMes)}**\n` +
          `📦 Productos en stock crítico: **${prods.filter(p=>getTotalStk(p)<=p.min).length}**\n` +
          `📬 Pedidos pendientes: **${pedidos.filter(p=>p.est!=="Entregado"&&p.est!=="Cancelado").length}**\n\n` +
          `Prueba preguntas como:\n• "¿Qué vendedor vendió más?"\n• "¿Cuánto vendí este mes?"\n• "¿Qué producto tiene mejor margen?"\n• "Dame un resumen del negocio"`;
      }

      setIaMsgs(m => [...m, { rol:"ai", txt: resp }]);
      setIaLoad(false);
    }, 600); // pequeño delay para sensación de "pensando"
  };

  // ─── DATOS GRÁFICOS ──────────────────────────────────────────────────────────
  const grafMeses = MESES.map((mes, i) => ({
    mes, Ventas: Math.round(prods.reduce((a,p) => a + p.v[i]*p.p, 0)), Costo: Math.round(prods.reduce((a,p) => a + p.v[i]*p.c, 0))
  }));
  const grafCat = Object.keys(CAT_C).map(cat => {
    const ps = prods.filter(p => p.cat === cat);
    return { cat, Unidades: ps.reduce((a,p) => a + p.v.reduce((x,y) => x+y,0), 0), color: CAT_C[cat] };
  }).sort((a,b) => b.Unidades - a.Unidades);

  // ─── PANTALLA CARGA INICIAL ───────────────────────────────────────────────
  const isAdmin   = user?.rol === "admin";
  const userModulos = user?.modulos || modulosPorRol(user?.rol || "vendedor");
  const navItems = !user ? [] : TODOS_MODULOS.filter(m => userModulos.includes(m.id)).map(m => ({
    id: m.id, ico: { dashboard:"🏠",pos:"🛒",ventas:"💰",pedidos:"📬",traslados:"🚚",vendedores:"🏆",clientes:"👥",cotiz:"📋",inventario:"📦",kardex:"📒",produccion:"🔨",compras:"🏭",gastos:"💸",finanzas:"📊",abc:"🏆",prediccion:"🔮",usuarios:"👤",ia:"🤖",etiquetas:"🏷️" }[m.id], label: m.label.replace(/^.{2}/,"")
  })).map(m => ({ ...m, label: TODOS_MODULOS.find(x=>x.id===m.id)?.label.slice(2).trim() || m.label }));
  const modActual = navItems.find(n => n.id === mod) ? mod : navItems[0]?.id || "pos";

  // Sidebar NO se abre/cierra automáticamente — el usuario la controla manualmente con el botón triángulo
  // Solo se cierra automáticamente la primera vez que se entra al POS
  const posEntradaRef = useRef(false);
  useEffect(() => {
    if (user && modActual === "pos" && !posEntradaRef.current) {
      setSidebarOpen(false);
      posEntradaRef.current = true;
    }
  }, [modActual, user]);

  if (!dbReady) return (
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg,${C.bg} 0%,${C.bg2} 100%)`, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:"'Segoe UI',sans-serif", gap:16 }}>
      <img src={LOGO_MED} alt="MoblaMel" style={{ width:72, height:72, objectFit:"contain", borderRadius:16 }}/>
      <div style={{ fontSize:18, fontWeight:700, color:C.t1 }}>MoblaMel ERP</div>
      <div style={{ display:"flex", gap:6 }}>
        {[0,1,2].map(i => (
          <div key={i} style={{ width:8, height:8, borderRadius:"50%", background:C.ac, animation:"pulse 1.2s ease-in-out infinite", animationDelay:`${i*0.2}s` }}/>
        ))}
      </div>
      <div style={{ fontSize:12, color:C.t3 }}>Conectando con la base de datos...</div>
      <style>{`@keyframes pulse { 0%,100%{opacity:.2;transform:scale(.8)} 50%{opacity:1;transform:scale(1)} }`}</style>
    </div>
  );

  // ─── PANTALLA LOGIN ──────────────────────────────────────────────────────────
  if (!user) return (
    <div style={{ minHeight:"100vh", display:"flex", fontFamily:"'Segoe UI','Helvetica Neue',sans-serif" }}>
      <style>{`@media(max-width:640px){.login-panel-left{display:none!important}.login-panel-right{width:100%!important;max-width:100%!important}}`}</style>
      {/* Panel izquierdo — solo desktop */}
      <div className="login-panel-left" style={{ flex:1, background:`linear-gradient(160deg, #3d2b1a 0%, #6b4226 50%, #a0714f 100%)`, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:48, position:"relative", overflow:"hidden" }}>
        <div style={{ position:"absolute", width:400, height:400, borderRadius:"50%", border:"1px solid rgba(255,255,255,0.08)", top:-100, left:-100 }}/>
        <div style={{ position:"absolute", width:300, height:300, borderRadius:"50%", border:"1px solid rgba(255,255,255,0.06)", bottom:-80, right:-80 }}/>
        {/* Logo SIN filtro — muestra los colores originales */}
        <img src={LOGO_MED} alt="MoblaMel" style={{ width:110, height:110, objectFit:"cover", marginBottom:24, borderRadius:16 }}/>
        
        <div style={{ fontSize:15, color:"rgba(255,255,255,0.70)", textAlign:"center", lineHeight:1.6, maxWidth:260 }}>Sistema de gestión para mueblería de melamina</div>
        <div style={{ marginTop:36, display:"flex", gap:12, flexWrap:"wrap", justifyContent:"center" }}>
          {["📦 Inventario","💰 Ventas","📊 Reportes","🚚 Pedidos"].map(f => (
            <div key={f} style={{ background:"rgba(255,255,255,0.10)", border:"1px solid rgba(255,255,255,0.15)", borderRadius:8, padding:"6px 14px", fontSize:12, color:"rgba(255,255,255,0.85)" }}>{f}</div>
          ))}
        </div>
      </div>
      {/* Panel derecho — formulario */}
      <div className="login-panel-right" style={{ width:440, background:C.white, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"48px 40px", boxShadow:"-4px 0 24px rgba(60,35,15,0.10)" }}>
        <div style={{ width:"100%", maxWidth:340 }}>
          <div style={{ textAlign:"center", marginBottom:32 }}>
            {/* Logo en móvil */}
            <img src={LOGO_MED} alt="MoblaMel" style={{ width:72, height:72, objectFit:"cover", marginBottom:10, display:"block", margin:"0 auto 10px", borderRadius:12 }}/>
            <div style={{ fontSize:24, fontWeight:800, color:C.t1, marginBottom:4 }}>Bienvenido</div>
            <div style={{ fontSize:13, color:C.t3 }}>Ingresa tus credenciales para continuar</div>
          </div>
          <div style={{ marginBottom:16 }}>
            <label style={{ fontSize:11, fontWeight:700, color:C.t2, display:"block", marginBottom:6, letterSpacing:"0.5px" }}>USUARIO</label>
            <input value={lU} onChange={e => setLU(e.target.value)} onKeyDown={e => e.key === "Enter" && login()}
              placeholder="Tu nombre de usuario"
              style={{ width:"100%", background:C.bg, border:`1.5px solid ${C.border}`, borderRadius:10, padding:"11px 14px", fontSize:14, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}
              onFocus={e => e.target.style.borderColor=C.ac} onBlur={e => e.target.style.borderColor=C.border}/>
          </div>
          <div style={{ marginBottom:8 }}>
            <label style={{ fontSize:11, fontWeight:700, color:C.t2, display:"block", marginBottom:6, letterSpacing:"0.5px" }}>CONTRASEÑA</label>
            <input type="password" value={lP} onChange={e => setLP(e.target.value)} onKeyDown={e => e.key === "Enter" && login()}
              placeholder="••••••••"
              style={{ width:"100%", background:C.bg, border:`1.5px solid ${C.border}`, borderRadius:10, padding:"11px 14px", fontSize:14, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}
              onFocus={e => e.target.style.borderColor=C.ac} onBlur={e => e.target.style.borderColor=C.border}/>
          </div>
          {lErr && <div style={{ color:C.rd, fontSize:12, marginBottom:12, padding:"8px 12px", background:C.rdBg, borderRadius:8 }}>⚠️ {lErr}</div>}
          <button onClick={login} style={{ width:"100%", padding:"13px", borderRadius:10, border:"none", background:`linear-gradient(135deg, ${C.ac}, ${C.acD})`, color:"#fff", fontSize:15, fontWeight:700, cursor:"pointer", fontFamily:"inherit", marginTop:16, boxShadow:"0 4px 14px rgba(100,60,20,0.25)", letterSpacing:"0.3px" }}>
            Ingresar →
          </button>
          <div style={{ textAlign:"center", marginTop:20, fontSize:11, color:C.t4 }}>MoblaMel ERP · Villa El Salvador</div>
        </div>
      </div>
    </div>
  );

  // ─── RENDER PRINCIPAL ────────────────────────────────────────────────────────
  return (
    <div style={{ height:"100dvh", background:C.bg, color:C.t1, fontFamily:"'Segoe UI','Helvetica Neue',sans-serif", fontSize:14, display:"flex", flexDirection:"column", overflow:"hidden" }}>
      <style>{`
        *{-webkit-tap-highlight-color:transparent;box-sizing:border-box;}
        body{overscroll-behavior:none;}
        .scroll-touch{-webkit-overflow-scrolling:touch;overflow-y:auto;}
        @supports(padding:max(0px)){
          .safe-bottom{padding-bottom:max(0px,env(safe-area-inset-bottom));}
        }
      `}</style>

      {/* TOPBAR */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.border}`, height:54, display:"flex", alignItems:"center", padding:"0 20px", gap:16, flexShrink:0, boxShadow:"0 1px 8px rgba(100,60,20,0.08)", position:"sticky", top:0, zIndex:50 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <div style={{ width:32, height:32, borderRadius:8, overflow:"hidden", background:"#f9f6f0", display:"flex", alignItems:"center", justifyContent:"center" }}>
            <img src={LOGO_SMALL} alt="M" style={{ width:30, height:30, objectFit:"contain" }}/>
          </div>
          <span style={{ fontSize:15, fontWeight:800, color:C.ac }}>MoblaMel</span>
        </div>
        <div style={{ flex:1 }} />
        {/* Alerta de cumpleaños — solo admin */}
        {isAdmin && (() => {
          const hoy = HOY.slice(5);
          const cumpleaneros = usuarios.filter(u => u.cumple && u.cumple.slice(5) === hoy && u.activo);
          if (!cumpleaneros.length) return null;
          return (
            <div style={{ display:"flex", alignItems:"center", gap:6, background:"#fef9c3", border:"1px solid #fde047", borderRadius:8, padding:"5px 10px" }}>
              <span style={{ fontSize:14 }}>🎂</span>
              <span style={{ fontSize:12, fontWeight:700, color:"#854d0e" }}>¡Cumpleaños de {cumpleaneros.map(u=>u.nombre).join(" y ")} hoy!</span>
            </div>
          );
        })()}
        {kpi.crit > 0 && (() => {
          const critProds = prods.filter(p => getTotalStk(p) <= p.min);
          const msg = `🚨 *STOCK BAJO - ${new Date().toLocaleDateString("es-PE")}*\n\n` +
            critProds.map(p => {
              const stk = getTotalStk(p);
              return `• ${p.ico} *${p.n}* — ${stk} uds (mín: ${p.min})`;
            }).join("\n") +
            `\n\n_MaderERP - ${critProds.length} producto(s) requieren reabastecimiento_`;
          return (
            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
              <div onClick={() => setMod("inventario")} style={{ display:"flex", alignItems:"center", gap:6, background:C.rdBg, border:`1px solid ${C.rdL}44`, borderRadius:8, padding:"5px 10px", cursor:"pointer" }}>
                <span style={{ fontSize:12 }}>⚠️</span>
                <span style={{ fontSize:12, fontWeight:700, color:C.rd }}>{kpi.crit} stock crítico</span>
              </div>
              <a href={`https://wa.me/?text=${encodeURIComponent(msg)}`} target="_blank" rel="noopener noreferrer"
                title="Enviar alerta de stock por WhatsApp"
                style={{ display:"flex", alignItems:"center", gap:5, padding:"5px 10px", borderRadius:8, background:"#25D366", border:"none", color:"#fff", textDecoration:"none", fontSize:11, fontWeight:700, cursor:"pointer", whiteSpace:"nowrap" }}>
                📲 WA
              </a>
            </div>
          );
        })()}
        <div style={{ display:"flex", alignItems:"center", gap:8, background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"6px 12px" }}>
          <span style={{ fontSize:16 }}>{user.ico}</span>
          <div>
            <div style={{ fontSize:12, fontWeight:700, color:C.t1, lineHeight:1 }}>{user.nombre}</div>
            <div style={{ fontSize:10, color:C.t3, lineHeight:1, marginTop:1, textTransform:"capitalize" }}>{user.rol}</div>
          </div>
          <button onClick={() => setUser(null)} style={{ background:"none", border:"none", color:C.t4, cursor:"pointer", fontSize:11, marginLeft:4 }}>salir</button>
        </div>
      </div>

      <div style={{ display:"flex", flex:1, overflow:"hidden", minHeight:0 }}>

        {/* SIDEBAR */}
        <div className="scroll-touch" style={{ width: sidebarOpen ? 210 : (modActual === "pos" ? 0 : 56), background:`linear-gradient(180deg, #3d2b1a 0%, #5c3d22 100%)`, borderRight:"none", padding: sidebarOpen ? "16px 10px" : modActual === "pos" ? "0" : "16px 8px", flexShrink:0, overflowY:"auto", overflowX:"hidden", display:"flex", flexDirection:"column", transition:"width 0.2s ease", position:"relative", visibility: !sidebarOpen && modActual === "pos" ? "hidden" : "visible" }}>
          {/* Toggle collapse button */}
          <button onClick={() => setSidebarOpen(o => !o)}
            style={{ display:"flex", alignItems:"center", justifyContent:"center", width:"100%", padding:"7px 4px", marginBottom:12, borderRadius:8, border:"1px solid rgba(255,255,255,0.15)", background:"rgba(255,255,255,0.08)", cursor:"pointer", color:"rgba(255,255,255,0.7)", fontSize:14, flexShrink:0 }}
            title={sidebarOpen ? "Colapsar menú" : "Expandir menú"}>
            {sidebarOpen ? "◀" : "▶"}
          </button>
          {(isAdmin || user.rol === "vendedor") && (
            <button onClick={() => { setModal("venta"); }} style={{ width:"100%", padding:"9px", borderRadius:8, border:"none", background:`linear-gradient(135deg, ${C.ac}, ${C.acD})`, color:"#fff", fontSize:sidebarOpen?12:18, fontWeight:700, cursor:"pointer", fontFamily:"inherit", textAlign:"center", marginBottom:12, whiteSpace:"nowrap", overflow:"hidden", boxShadow:"0 2px 8px rgba(100,50,10,0.3)" }}
              title="Nueva Venta">
              {sidebarOpen ? "🧾 + Nueva Venta" : "🧾"}
            </button>
          )}
          {sidebarOpen && <div style={{ fontSize:10, color:"rgba(255,255,255,0.35)", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.8px", padding:"0 8px", marginBottom:6 }}>Módulos</div>}
          {navItems.map(n => (
            <button key={n.id} onClick={() => setMod(n.id)} title={n.label}
              style={{ width:"100%", padding: sidebarOpen?"8px 12px":"8px 4px", borderRadius:8, border:"none", background: modActual === n.id ? "rgba(255,255,255,0.15)" : "transparent", color: modActual === n.id ? "#fff" : "rgba(255,255,255,0.55)", fontSize:12, fontWeight: modActual === n.id ? 700 : 400, cursor:"pointer", textAlign:sidebarOpen?"left":"center", fontFamily:"inherit", marginBottom:1, display:"flex", alignItems:"center", gap:sidebarOpen?8:0, justifyContent:sidebarOpen?"flex-start":"center", transition:"all 0.12s", overflow:"hidden", borderLeft: modActual === n.id ? `3px solid ${C.acL}` : "3px solid transparent" }}>
              <span style={{ fontSize:sidebarOpen?15:18, flexShrink:0 }}>{n.ico}</span>
              {sidebarOpen && <span style={{ whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{n.label}</span>}
              {sidebarOpen && (n.id === "ia" || n.id === "prediccion" || n.id === "etiquetas") && <span style={{ marginLeft:"auto", fontSize:9, background: n.id === "ia" ? C.ac : n.id === "prediccion" ? C.bl : C.grL, color:"#fff", padding:"1px 5px", borderRadius:99, fontWeight:700 }}>{n.id === "etiquetas" ? "NEW" : "IA"}</span>}
            </button>
          ))}
          <div style={{ flex:1 }} />
          {sidebarOpen && (
            <div style={{ padding:"10px 8px", borderTop:`1px solid ${C.border}`, marginTop:10, fontSize:11, color:C.t3 }}>
              <div style={{ fontWeight:700, color:C.t2, marginBottom:2 }}>Moblamel</div>
              <div style={{ color:C.t4 }}>Moblamel · VES</div>
            </div>
          )}
        </div>

        {/* BOTÓN FLOTANTE para mostrar sidebar cuando está oculta en POS */}
        {!sidebarOpen && modActual === "pos" && (
          <button onClick={() => setSidebarOpen(true)}
            style={{ position:"absolute", left:8, top:"50%", transform:"translateY(-50%)", zIndex:100, background:"#3d2b1a", border:"1px solid rgba(255,255,255,0.2)", borderRadius:"0 8px 8px 0", padding:"12px 6px", cursor:"pointer", boxShadow:"2px 0 8px rgba(0,0,0,0.2)", color:"rgba(255,255,255,0.7)", fontSize:14 }}
            title="Mostrar menú">
            ▶
          </button>
        )}

        {/* CONTENIDO */}
        <div className="scroll-touch safe-bottom" style={{ flex:1, overflowY:"auto", overflowX:"hidden", padding:"16px", background:C.bg, minHeight:0 }}>

          {/* Banner de error de BD */}
          {dbError && (
            <div style={{ background:C.rdBg, border:`1px solid ${C.rd}44`, borderRadius:10, padding:"10px 14px", marginBottom:14, fontSize:12, color:C.rd, display:"flex", alignItems:"center", gap:8 }}>
              ⚠️ {dbError} — Los cambios se guardarán solo localmente.
            </div>
          )}

          {/* ═══════ DASHBOARD ═══════ */}
          {modActual === "dashboard" && (
            <>
              <PageTitle title="🏠 Panel de Control" sub={`Bienvenido ${user.nombre} · ${new Date().toLocaleDateString("es-PE",{weekday:"long",day:"2-digit",month:"long"})}`} />
              {/* KPIs — admin ve todo, vendedor solo sus ventas */}
              {isAdmin ? (
                <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:10, marginBottom:16 }}>
                  <KPI icon="💰" label="Ingresos del mes"  value={fmtK(kpi.ing)}  color={C.grL}  trend={12} sub={`${kpi.tickets} ventas`} />
                  <KPI icon="📉" label="Gastos del mes"    value={fmtK(kpi.gm)}   color={C.rdL}  sub="Mat+taller+local" />
                  <KPI icon="✨" label="Utilidad neta"     value={fmtK(kpi.ut)}   color={C.ac}   trend={8}  sub={`Margen ${kpi.mg}%`} />
                  <KPI icon="📦" label="Valor inventario"  value={fmtK(kpi.vi)}   color={C.bl}   sub="A precio de costo" />
                </div>
              ) : (
                // Vendedor: solo sus propias ventas del mes
                (() => {
                  const misVentas = ventas.filter(v => v.vend === user.nombre && !v.esAnticipo);
                  const mesActual = HOY.slice(0,7);
                  const misVentasMes = misVentas.filter(v => v.f.startsWith(mesActual));
                  const miTotal = misVentasMes.reduce((a,v)=>a+v.tot,0);
                  const miAnterior = misVentas.filter(v=>v.f.startsWith(HOY.slice(0,5)+(parseInt(HOY.slice(5,7))-1).toString().padStart(2,"0"))).reduce((a,v)=>a+v.tot,0);
                  const FRASES = [
                    "💪 ¡Cada cliente es una oportunidad! Sonríe, escucha y cierra esa venta — tu comisión te lo agradecerá.",
                    "🌟 Un espacio limpio y ordenado genera confianza. ¡El cliente que se siente bien, compra más!",
                    "🔥 ¡Hoy puede ser tu mejor día! Cada 'no' te acerca al próximo 'sí'. ¡Sigue adelante!",
                    "😊 Tu actitud es tu mejor herramienta. Un vendedor amable y atento siempre supera sus metas.",
                    "🚀 ¡Las comisiones no se esperan, se conquistan! Da el 100% en cada atención.",
                    "🏆 Los mejores vendedores no nacen — se forjan siendo amables, honestos y persistentes.",
                    "✨ ¡Mantén tu área impecable! Un ambiente agradable invita al cliente a quedarse y comprar.",
                    "🎯 Conoce cada producto, responde con seguridad y el cliente confiará en ti. ¡Tú puedes!",
                    "💡 Una buena atención hoy es un cliente fiel mañana. ¡Invierte en cada interacción!",
                    "🤝 Sé el vendedor que a ti te gustaría encontrar: honesto, amable y siempre dispuesto a ayudar.",
                    "⭐ ¡Tu energía se contagia! Si tú crees en el producto, el cliente también creerá en él.",
                    "🌺 Un 'gracias, fue un placer atenderle' puede traer al cliente de vuelta. ¡Cuida cada detalle!",
                  ];
                  const frase = FRASES[new Date().getDate() % FRASES.length];
                  return (
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:20 }}>
                      <KPI icon="💰" label="Mis ventas del mes" value={fmt(miTotal)} color={C.grL} sub={`${misVentasMes.length} ventas`}/>
                      <div style={{ background:"linear-gradient(135deg,#fff7ed,#fed7aa)", border:"1px solid #fdba74", borderRadius:14, padding:"14px 16px", display:"flex", flexDirection:"column", justifyContent:"center" }}>
                        <div style={{ fontSize:11, fontWeight:700, color:"#c05621", textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:6 }}>💬 Frase del día</div>
                        <div style={{ fontSize:13, color:"#7c2d12", fontWeight:600, lineHeight:1.4 }}>{frase}</div>
                      </div>
                      <KPI icon="📬" label="Pedidos activos" value={pedidos.filter(p=>p.vend===user.nombre&&p.est!=="Entregado"&&p.est!=="Cancelado").length} color={C.or} sub="En curso"/>
                    </div>
                  );
                })()
              )}
              <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:16, marginBottom:16 }}>
                <Card style={{ padding:18 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>📈 Ventas vs Costos · 12 meses</div>
                  <div style={{ fontSize:11, color:C.t4, marginBottom:12 }}>Línea naranja = lo que ingresó. Línea azul = lo que costó fabricar/comprar. La diferencia entre ambas es tu ganancia bruta.</div>
                  <ResponsiveContainer width="100%" height={160}>
                    <AreaChart data={grafMeses}>
                      <defs>
                        <linearGradient id="gv" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={C.ac} stopOpacity={0.15}/>
                          <stop offset="95%" stopColor={C.ac} stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <XAxis dataKey="mes" tick={{fill:C.t4,fontSize:10}} axisLine={false} tickLine={false}/>
                      <YAxis tick={{fill:C.t4,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v => `${Math.round(v/1000)}k`}/>
                      <Tooltip content={<ChartTip/>}/>
                      <Area type="monotone" dataKey="Ventas" stroke={C.ac} fill="url(#gv)" strokeWidth={2.5}/>
                      <Line type="monotone" dataKey="Costo" stroke={C.bl} strokeWidth={2} dot={false} strokeDasharray="4 2"/>
                    </AreaChart>
                  </ResponsiveContainer>
                  {/* Análisis automático */}
                  {(() => {
                    const best = grafMeses.reduce((a,b) => b.Ventas > a.Ventas ? b : a, grafMeses[0]);
                    const worst = grafMeses.reduce((a,b) => b.Ventas < a.Ventas ? b : a, grafMeses[0]);
                    const total = grafMeses.reduce((a,b) => a + b.Ventas, 0);
                    const totalCosto = grafMeses.reduce((a,b) => a + (b.Costo||0), 0);
                    const mgAnual = total > 0 ? Math.round(((total-totalCosto)/total)*100) : 0;
                    return (
                      <div style={{ marginTop:10, padding:"10px 12px", background:C.acBg, borderRadius:8, fontSize:11, color:C.t2, lineHeight:1.7 }}>
                        💡 <strong>Análisis:</strong> Tu mejor mes fue <strong>{best?.mes}</strong> con {fmt(best?.Ventas||0)} en ventas.
                        El mes más bajo fue <strong>{worst?.mes}</strong> con {fmt(worst?.Ventas||0)}.
                        Margen bruto anual: <strong style={{color:mgAnual>=30?C.grL:C.or}}>{mgAnual}%</strong>.
                        {mgAnual < 30 && " ⚠️ Margen bajo — revisar costos de producción."}
                        {mgAnual >= 40 && " ✅ Margen saludable."}
                      </div>
                    );
                  })()}
                </Card>
                <Card style={{ padding:18 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>🏆 Rotación por categoría</div>
                  <div style={{ fontSize:11, color:C.t4, marginBottom:10 }}>Unidades vendidas por tipo de mueble. Toca una barra para ver el detalle.</div>
                  {(() => {
                    const cats = Object.entries(
                      ventas.flatMap(v => v.items).reduce((acc, i) => {
                        const prod = prods.find(p => p.id === i.id);
                        const cat = prod?.cat || "Otro";
                        acc[cat] = (acc[cat]||0) + i.q;
                        return acc;
                      }, {})
                    ).map(([cat, qty]) => ({ cat, qty, color: CAT_C[cat]||C.t4 })).sort((a,b) => b.qty-a.qty);

                    const detalle = catSel ? ventas.flatMap(v => v.items.map(i => ({ ...i, vend:v.vend, f:v.f })))
                      .filter(i => prods.find(p=>p.id===i.id)?.cat === catSel)
                      .reduce((acc, i) => { acc[i.n] = (acc[i.n]||0)+i.q; return acc; }, {}) : null;

                    return (<>
                      {cats.map((c,i) => (
                        <div key={c.cat} onClick={() => setCatSel(catSel===c.cat?null:c.cat)}
                          style={{ marginBottom:8, cursor:"pointer", padding:"4px 6px", borderRadius:6, background:catSel===c.cat?c.color+"15":"transparent", border:`1px solid ${catSel===c.cat?c.color:C.border}` }}>
                          <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3 }}>
                            <span style={{ fontWeight:600, color:catSel===c.cat?c.color:C.t2 }}>{{Ropero:"🪞",Zapatero:"👟",Cómoda:"🗄️","Mesa de Noche":"🛏️",Cabecera:"🛋️"}[c.cat]||"📦"} {c.cat}</span>
                            <span style={{ fontWeight:700, color:c.color }}>{c.qty} uds</span>
                          </div>
                          <div style={{ height:5, borderRadius:99, background:C.bg3, overflow:"hidden" }}>
                            <div style={{ height:"100%", width:`${cats[0].qty>0?(c.qty/cats[0].qty)*100:0}%`, background:c.color, borderRadius:99 }}/>
                          </div>
                        </div>
                      ))}
                      {catSel && detalle && (
                        <div style={{ marginTop:8, padding:"8px 10px", background:C.bg, borderRadius:8, fontSize:11 }}>
                          <div style={{ fontWeight:700, color:C.t1, marginBottom:5 }}>Productos más vendidos en {catSel}:</div>
                          {Object.entries(detalle).sort((a,b)=>b[1]-a[1]).map(([n,q]) => (
                            <div key={n} style={{ display:"flex", justifyContent:"space-between", color:C.t2, padding:"2px 0" }}>
                              <span>{n}</span><span style={{ fontWeight:700, color:C.ac }}>{q} uds</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </>);
                  })()}
                </Card>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>
                <Card style={{ padding:16 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:10, textTransform:"uppercase", letterSpacing:"0.5px" }}>🏆 Vendedores del mes</div>
                  {usuarios.filter(u => u.rol !== "taller").map((u, i) => {
                    const tot = ventas.filter(v => v.vend === u.nombre && v.f.startsWith("2025-05")).reduce((a,v) => a+v.tot, 0);
                    const max = Math.max(...usuarios.filter(x => x.rol !== "taller").map(x => ventas.filter(v => v.vend === x.nombre && v.f.startsWith("2025-05")).reduce((a,v) => a+v.tot, 0)));
                    return (
                      <div key={u.id} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
                        <span style={{ fontSize:13 }}>{u.ico}</span>
                        <div style={{ flex:1 }}>
                          <div style={{ fontSize:12, fontWeight:600, color:C.t1, marginBottom:2 }}>{u.nombre}</div>
                          <div style={{ height:4, borderRadius:99, background:C.bg3, overflow:"hidden" }}>
                            <div style={{ height:"100%", width:`${max > 0 ? (tot/max)*100 : 0}%`, background: i === 0 ? C.ac : C.t4, borderRadius:99 }}/>
                          </div>
                        </div>
                        <span style={{ fontSize:12, fontWeight:700, color: i === 0 ? C.ac : C.t3 }}>{fmtK(tot)}</span>
                      </div>
                    );
                  })}
                </Card>
                <Card style={{ padding:16 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:10, textTransform:"uppercase", letterSpacing:"0.5px" }}>⚠️ Stock crítico</div>
                  {prods.filter(p => p.stk <= p.min).length === 0
                    ? <div style={{ color:C.grL, fontSize:13, padding:"10px 0" }}>✅ Todo el stock está OK</div>
                    : prods.filter(p => p.stk <= p.min).map(p => (
                      <div key={p.id} style={{ marginBottom:10, paddingBottom:10, borderBottom:`1px solid ${C.bg3}` }}>
                        <div style={{ fontSize:12, fontWeight:600, color:C.t1 }}>{p.ico} {p.n}</div>
                        <div style={{ fontSize:11, color:C.t3 }}>Stock: {p.stk} / Mínimo: {p.min}</div>
                      </div>
                    ))
                  }
                </Card>
                <Card style={{ padding:16 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:10, textTransform:"uppercase", letterSpacing:"0.5px" }}>🕐 Últimas ventas</div>
                  {ventas.slice(0,4).map(v => (
                    <div key={v.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:9, paddingBottom:9, borderBottom:`1px solid ${C.bg3}` }}>
                      <div>
                        <div style={{ fontSize:12, fontWeight:600, color:C.t1 }}>{v.cli}</div>
                        <div style={{ fontSize:11, color:C.t3 }}>{v.f} · {v.vend}</div>
                      </div>
                      <span style={{ fontSize:13, fontWeight:700, color:C.ac }}>{fmt(v.tot)}</span>
                    </div>
                  ))}
                </Card>
              </div>
            </>
          )}

          {/* ═══════ POS ═══════ */}
          {modActual === "pos" && (() => {
            // posFase: "catalogo" | "cobro"
            const CAT_TABS = [
              {key:"Todos",ico:"🔍",label:"Todos"},
              {key:"Ropero",ico:"🪞",label:"Roperos"},
              {key:"Zapatero",ico:"👟",label:"Zapateros"},
              {key:"Cómoda",ico:"🗄️",label:"Cómodas"},
              {key:"Mesa de Noche",ico:"🛏️",label:"Veladores"},
              {key:"Cabecera",ico:"🛋️",label:"Cabeceras"},
            ];
            const totalItems = posCarrito.reduce((a,i)=>a+i.q,0);
            return (
              <>
              <style>{`
                @media(max-width:767px){
                  .pos-cats{display:none!important;}
                  .pos-grid-inner{grid-template-columns:repeat(2,1fr)!important;}
                }
                .pos-card-hover:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,0.1);}
                .pos-card-hover{transition:transform 0.15s,box-shadow 0.15s;}
              `}</style>

              {/* ── FASE CATÁLOGO ── */}
              {posFase === "catalogo" && (
                <div style={{display:"flex",height:"calc(100dvh - 120px)",minHeight:400,borderRadius:14,overflow:"hidden",border:`1px solid ${C.border}`,boxShadow:"0 4px 24px rgba(0,0,0,0.08)"}}>

                  {/* Categorías verticales */}
                  <div className="pos-cats" style={{width:80,background:"#1e2a3a",display:"flex",flexDirection:"column",alignItems:"center",padding:"10px 0",gap:4,flexShrink:0,overflowY:"auto"}}>
                    {CAT_TABS.map(({key,ico,label})=>(
                      <button key={key} onClick={()=>setPosCat(key)}
                        style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",width:64,padding:"8px 4px",borderRadius:10,border:"none",background:posCat===key?"#e07b39":"transparent",cursor:"pointer",gap:3,flexShrink:0}}>
                        <span style={{fontSize:20}}>{ico}</span>
                        <span style={{fontSize:9,fontWeight:700,color:posCat===key?"#fff":"rgba(255,255,255,0.55)",textAlign:"center",lineHeight:1.2}}>{label}</span>
                      </button>
                    ))}
                  </div>

                  {/* Centro: búsqueda + grid de productos */}
                  <div style={{flex:1,display:"flex",flexDirection:"column",background:C.bg,minWidth:0}}>
                    {/* Barra superior */}
                    <div style={{padding:"10px 12px",background:C.white,borderBottom:`1px solid ${C.border}`,display:"flex",gap:8,flexShrink:0,alignItems:"center"}}>
                      <div style={{flex:1,position:"relative"}}>
                        <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:C.t4,fontSize:14}}>🔍</span>
                        <input style={{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 10px 8px 32px",fontSize:13,outline:"none",boxSizing:"border-box",color:C.t1,fontFamily:"inherit"}}
                          placeholder="Buscar producto..." value={posBusq} onChange={e=>{setPosBusq(e.target.value);setScanMode(false);}}/>
                      </div>
                      <button onClick={()=>{setScanMode(s=>!s);setPosBusq("");}}
                        style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${scanMode?C.ac:C.border}`,background:scanMode?C.acBg:C.white,color:scanMode?C.ac:C.t3,cursor:"pointer",fontSize:16,flexShrink:0}}>📷</button>
                      <div style={{display:"flex",gap:4,flexShrink:0}}>
                        {["B","F","NV"].map((c,i)=>{const full=["Boleta","Factura","Nota de Venta"][i];return(
                          <button key={c} onClick={()=>setPosComp(full)}
                            style={{padding:"6px 10px",borderRadius:7,border:`1px solid ${posComp===full?C.ac:C.border}`,background:posComp===full?C.acBg:"transparent",color:posComp===full?C.ac:C.t3,fontSize:11,fontWeight:700,cursor:"pointer"}}>
                            {c}
                          </button>
                        );})}
                      </div>
                    </div>
                    {scanMode&&(
                      <div style={{padding:"8px 12px",background:C.acBg,borderBottom:`1px solid ${C.border}`,display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
                        <span style={{fontSize:11,fontWeight:700,color:C.ac}}>📷 ESCÁNER</span>
                        <input ref={scanRef} autoFocus value={scanVal} onChange={e=>setScanVal(e.target.value)}
                          onKeyDown={e=>{if(e.key==="Enter"&&scanVal){const p=prods.find(x=>x.id===scanVal.toUpperCase()||x.barcode===scanVal);if(p){posAgregar(p,p.cols[0]);setScanVal("");setScanMode(false);showToast(`✓ ${p.n}`);}else showToast("Código no encontrado","err");setScanVal("");}}}
                          style={{flex:1,background:C.white,border:`2px solid ${C.ac}`,borderRadius:8,padding:"6px 12px",fontSize:13,outline:"none",fontFamily:"monospace",color:C.t1}}
                          placeholder="Escanea o escribe → Enter"/>
                        <button onClick={()=>setScanMode(false)} style={{padding:"6px 10px",borderRadius:7,border:`1px solid ${C.border}`,background:C.white,cursor:"pointer",fontSize:12}}>✕</button>
                      </div>
                    )}

                    {/* Grid de productos con scroll */}
                    <div style={{flex:1,overflowY:"auto",padding:12}}>
                      <div className="pos-grid-inner" style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10,alignContent:"start"}}>
                        {posFiltrados.map(prod=>{
                          const clr=CAT_C[prod.cat]||C.ac;
                          const enCarrito=posCarrito.filter(i=>i.pid===prod.id).reduce((a,i)=>a+i.q,0);
                          const stockDisp=Math.max(0,getTotalStk(prod)-(prod.stkRes||0));
                          const locsInfo=getLocs(prod);
                          return(
                            <div key={prod.id} className="pos-card-hover"
                              style={{background:C.white,border:`2px solid ${enCarrito?clr:C.border}`,borderRadius:12,overflow:"hidden",position:"relative",cursor:stockDisp>0?"pointer":"not-allowed",opacity:stockDisp>0?1:0.6,display:"flex",flexDirection:"column"}}>
                              {enCarrito>0&&<div style={{position:"absolute",top:7,right:7,width:22,height:22,borderRadius:"50%",background:clr,color:"#fff",fontSize:11,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",zIndex:2}}>{enCarrito}</div>}
                              {stockDisp<=0&&<div style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:3}}><span style={{color:"#fff",fontWeight:700,fontSize:11,background:"rgba(0,0,0,0.6)",padding:"3px 8px",borderRadius:5}}>SIN STOCK</span></div>}
                              {/* Franja + ícono */}
                              <div style={{position:"relative",flexShrink:0}}>
                                <div style={{height:4,background:clr}}/>
                                {prod.img
                                  ?<div style={{height:70,overflow:"hidden"}}><img src={prod.img} alt={prod.n} style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
                                  :<div style={{height:52,background:clr+"0d",display:"flex",alignItems:"center",justifyContent:"center"}}>
                                    <span style={{fontSize:28,opacity:0.55}}>{prod.ico}</span>
                                  </div>
                                }
                              </div>
                              {/* Info */}
                              <div style={{padding:"7px 9px 9px",flex:1}}>
                                <div style={{fontSize:12,fontWeight:700,color:C.t1,marginBottom:2,lineHeight:1.3,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{prod.n}</div>
                                <div style={{fontSize:14,fontWeight:800,color:clr,marginBottom:5}}>{fmt(prod.p)}</div>
                                {/* Stock mini por ubicación */}
                                <div style={{marginBottom:6,display:"flex",gap:3,flexWrap:"wrap"}}>
                                  {UBICACIONES.map(ubi=>{
                                    const tot=Object.entries(locsInfo).filter(([k])=>k.startsWith(ubi+"|")).reduce((a,[,v])=>a+v,0);
                                    if(!tot) return null;
                                    const uClr=ubi==="Tienda Principal"?C.bl:ubi==="Tienda Pasaje"?C.pu:C.or;
                                    return <span key={ubi} style={{fontSize:9,color:uClr,fontWeight:600,background:uClr+"15",padding:"1px 5px",borderRadius:99,whiteSpace:"nowrap"}}>
                                      {ubi==="Tienda Principal"?"T.Ppal":ubi==="Tienda Pasaje"?"T.Pasaje":"Almacén"}: {tot}
                                    </span>;
                                  })}
                                </div>
                                {/* Colores */}
                                <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                                  {prod.cols.map(col=>{
                                    const enC=posCarrito.find(i=>i.pid===prod.id&&i.col===col);
                                    const stkCol=Object.entries(locsInfo).filter(([k])=>k.includes(`|${col}`)).reduce((a,[,v])=>a+v,0);
                                    return(
                                      <button key={col} onClick={()=>{
                                        const yaEnCart=posCarrito.filter(i=>i.pid===prod.id&&i.col===col).reduce((a,i)=>a+i.q,0);
                                        if(stkCol>0&&yaEnCart<stkCol){posAgregar(prod,col);}
                                        else if(yaEnCart>=stkCol) showToast(`Stock máximo de ${col}: ${stkCol} uds`,"err");
                                      }}
                                        style={{display:"flex",alignItems:"center",gap:4,padding:"4px 8px",borderRadius:99,border:`2px solid ${enC?clr:stkCol>0?C.border:"#eee"}`,background:enC?clr+"18":stkCol>0?"transparent":"#f9f9f9",cursor:stkCol>0?"pointer":"not-allowed",opacity:stkCol>0?1:0.4,fontSize:11,fontWeight:700,color:enC?clr:stkCol>0?C.t2:C.t4}}>
                                        <div style={{width:8,height:8,borderRadius:"50%",background:HEX_COLOR[col]||"#888",flexShrink:0}}/>
                                        {col}{enC&&<span style={{background:clr,color:"#fff",fontSize:9,padding:"0 4px",borderRadius:99}}>{enC.q}</span>}
                                      </button>
                                    );
                                  })}
                                </div>
                              </div>
                            </div>
                          );
                        })}
                        {posFiltrados.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"50px 0",color:C.t4}}><div style={{fontSize:36,marginBottom:8}}>🔍</div>Sin resultados</div>}
                      </div>
                    </div>

                    {/* BARRA INFERIOR FLOTANTE — total acumulado + ir a cobrar */}
                    {posCarrito.length>0&&(
                      <div style={{background:"#1e2a3a",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,borderTop:"2px solid #e07b39"}}>
                        <div style={{display:"flex",gap:16,alignItems:"center"}}>
                          <div>
                            <div style={{fontSize:11,color:"rgba(255,255,255,0.6)"}}>TOTAL ({totalItems} items)</div>
                            <div style={{fontSize:22,fontWeight:800,color:"#e07b39"}}>{fmt(posBase)}</div>
                          </div>
                          {/* Mini lista de items */}
                          <div style={{display:"flex",gap:6,flexWrap:"wrap",maxWidth:300}}>
                            {posCarrito.map(item=>(
                              <div key={item.key} style={{display:"flex",alignItems:"center",gap:4,background:"rgba(255,255,255,0.1)",borderRadius:99,padding:"3px 8px"}}>
                                <div style={{width:7,height:7,borderRadius:"50%",background:HEX_COLOR[item.col]||"#888"}}/>
                                <span style={{fontSize:10,color:"#fff",fontWeight:600}}>{item.n.split(" ").slice(0,2).join(" ")} {item.col} ×{item.q}</span>
                                <button onClick={()=>setPosCarrito(c=>c.filter(x=>x.key!==item.key))} style={{background:"none",border:"none",color:"rgba(255,255,255,0.5)",cursor:"pointer",fontSize:12,padding:0,marginLeft:2}}>✕</button>
                              </div>
                            ))}
                          </div>
                        </div>
                        <div style={{display:"flex",gap:8,flexShrink:0}}>
                          <button onClick={()=>setPosCarrito([])} style={{padding:"10px 14px",borderRadius:8,border:"1px solid rgba(255,255,255,0.3)",background:"transparent",color:"rgba(255,255,255,0.7)",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑 Limpiar</button>
                          <button onClick={()=>setPosTab("cobro")}
                            style={{padding:"12px 24px",borderRadius:10,border:"none",background:"#e07b39",color:"#fff",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 12px rgba(224,123,57,0.4)"}}>
                            💳 Ir a cobrar →
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* ── FASE COBRO ── */}
              {posFase === "cobro" && (
                <div style={{maxWidth:520,margin:"0 auto"}}>
                  <button onClick={()=>setPosTab("catalogo")}
                    style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",color:C.ac,cursor:"pointer",fontWeight:700,fontSize:13,marginBottom:14,fontFamily:"inherit",padding:0}}>
                    ← Volver al catálogo
                  </button>
                  <Card style={{padding:0,overflow:"hidden"}}>
                    {/* Header */}
                    <div style={{background:"#1e2a3a",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div>
                        <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>🛒 Resumen del pedido</div>
                        <div style={{fontSize:11,color:"rgba(255,255,255,0.6)"}}>{totalItems} item(s) · {posComp}</div>
                      </div>
                      <div style={{display:"flex",gap:5}}>
                        {["B","F","NV"].map((c,i)=>{const full=["Boleta","Factura","Nota de Venta"][i];return(
                          <button key={c} onClick={()=>setPosComp(full)}
                            style={{padding:"4px 10px",borderRadius:6,border:`1px solid ${posComp===full?"#e07b39":"rgba(255,255,255,0.3)"}`,background:posComp===full?"#e07b39":"transparent",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                            {c}
                          </button>
                        );})}
                      </div>
                    </div>

                    {/* Items */}
                    <div style={{padding:"10px 14px",maxHeight:220,overflowY:"auto",borderBottom:`1px solid ${C.border}`}}>
                      {posCarrito.map(item=>{
                        const clr=CAT_C[prods.find(p=>p.id===item.pid)?.cat]||C.ac;
                        const hayDesc=item.pr<item.prLista;
                        return(
                          <div key={item.key} style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
                            <div style={{flex:1}}>
                              <div style={{fontSize:13,fontWeight:700,color:C.t1}}>{item.n}</div>
                              <div style={{display:"flex",alignItems:"center",gap:4}}>
                                <div style={{width:8,height:8,borderRadius:"50%",background:HEX_COLOR[item.col]||"#888"}}/>
                                <span style={{fontSize:11,color:C.t3}}>{item.col}</span>
                              </div>
                            </div>
                            <div style={{display:"flex",alignItems:"center",gap:5}}>
                              <button onClick={()=>setPosCarrito(c=>c.map(x=>x.key===item.key?{...x,q:Math.max(1,x.q-1)}:x))} style={{width:24,height:24,borderRadius:6,border:`1px solid ${C.border}`,background:C.bg,cursor:"pointer",fontWeight:700,fontSize:13}}>−</button>
                              <span style={{fontWeight:800,minWidth:18,textAlign:"center",color:clr}}>{item.q}</span>
                              <button onClick={()=>setPosCarrito(c=>c.map(x=>x.key===item.key&&x.q<x.maxStk?{...x,q:x.q+1}:x))} style={{width:24,height:24,borderRadius:6,border:`1px solid ${C.border}`,background:C.bg,cursor:"pointer",fontWeight:700,fontSize:13}}>+</button>
                            </div>
                            <div style={{position:"relative",width:90}}>
                              <span style={{position:"absolute",left:6,top:"50%",transform:"translateY(-50%)",fontSize:10,color:C.t3,pointerEvents:"none"}}>S/</span>
                              <input type="number" value={item.pr} onChange={e=>setPosCarrito(c=>c.map(x=>x.key===item.key?{...x,pr:parseFloat(e.target.value)||0}:x))}
                                style={{width:"100%",background:C.white,border:`1px solid ${hayDesc?C.or:C.border}`,borderRadius:6,padding:"4px 4px 4px 20px",fontSize:12,fontWeight:700,outline:"none",color:hayDesc?C.or:C.t1,boxSizing:"border-box"}}/>
                            </div>
                            <span style={{fontSize:12,fontWeight:700,color:C.t2,minWidth:55,textAlign:"right"}}>{fmt(item.q*item.pr)}</span>
                            <button onClick={()=>setPosCarrito(c=>c.filter(x=>x.key!==item.key))} style={{background:"none",border:"none",color:C.rdL,cursor:"pointer",fontSize:14}}>✕</button>
                          </div>
                        );
                      })}
                    </div>

                    {/* Cobro */}
                    <div style={{padding:"12px 14px"}}>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:C.t3,marginBottom:4}}><span>Subtotal</span><span>{fmt(posSub)}</span></div>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                        <span style={{fontSize:12,color:C.t3,flex:1}}>🏷️ Descuento S/</span>
                        <input type="number" min="0" value={posDesc} onChange={e=>setPosDesc(e.target.value)}
                          style={{width:80,background:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"4px 8px",fontSize:12,outline:"none",textAlign:"right"}}/>
                      </div>
                      {tienePOS&&montoPOS>0&&(
                        <div style={{background:C.orBg,borderRadius:7,padding:"6px 10px",marginBottom:8,fontSize:12,color:C.or,fontWeight:700}}>
                          💳 Izipay 5%: cliente paga {fmt(montoEnDatafono)} en datáfono · la tienda recibe {fmt(montoPOS)}
                        </div>
                      )}
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:24,fontWeight:800,color:C.ac,padding:"8px 0",borderTop:`1px solid ${C.border}`,borderBottom:`1px solid ${C.border}`,margin:"6px 0 10px"}}>
                        <span>TOTAL</span><span>S/ {posTotal.toFixed(2)}</span>
                      </div>
                      <input style={{width:"100%",background:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 10px",fontSize:12,outline:"none",boxSizing:"border-box",fontFamily:"inherit",color:C.t1,marginBottom:8}}
                        placeholder="👤 Cliente (opcional)" value={posCliente} onChange={e=>setPosCliente(e.target.value)}/>
                      {/* Métodos de pago */}
                      {posPagos.map((p,i)=>(
                        <div key={i} style={{marginBottom:6}}>
                          <div style={{display:"flex",gap:6,alignItems:"center"}}>
                            <select value={p.met} onChange={e=>{
                              const nuevoMet = e.target.value;
                              if(nuevoMet==="POS Tarjeta"){
                                // Auto-calcular: lo que falta cubrir con POS
                                const otrosPagos = posPagos.filter((_,j)=>j!==i).reduce((a,x)=>a+(parseFloat(x.monto)||0),0);
                                const montoPos = +(posBase - otrosPagos).toFixed(2);
                                setPosPagos(prev=>prev.map((x,j)=>j===i?{...x,met:nuevoMet,monto:String(Math.max(0,montoPos))}:x));
                              } else {
                                updPago(i,"met",nuevoMet);
                              }
                            }}
                              style={{flex:1,background:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 9px",fontSize:12,outline:"none",fontFamily:"inherit"}}>
                              {METODOS_PAGO.map(m=><option key={m}>{m}</option>)}
                            </select>
                            {p.met==="POS Tarjeta"
                              ? <div style={{width:100,background:"#fff7ed",border:`2px solid ${C.or}`,borderRadius:7,padding:"7px 9px",fontSize:13,fontWeight:800,textAlign:"right",color:C.or,flexShrink:0}}>
                                  {fmt(parseFloat(p.monto)||0)}
                                </div>
                              : <input type="number" value={p.monto} onChange={e=>updPago(i,"monto",e.target.value)}
                                  style={{width:100,background:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 9px",fontSize:13,fontWeight:700,outline:"none",textAlign:"right",flexShrink:0}}
                                  placeholder="Monto"/>
                            }
                            {i>0&&<button onClick={()=>setPosPagos(p=>p.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:C.rdL,cursor:"pointer",fontSize:14}}>✕</button>}
                          </div>
                          {p.met==="POS Tarjeta"&&(parseFloat(p.monto)||0)>0&&(
                            <div style={{background:"#fff7ed",borderRadius:6,padding:"5px 10px",marginTop:4}}>
                              <span style={{fontSize:12,color:C.or,fontWeight:800}}>
                                📱 Pasar por datáfono: S/{(parseFloat(p.monto)*1.05).toFixed(2)}
                              </span>
                              <span style={{fontSize:10,color:C.t3,marginLeft:8}}>
                                (tienda recibe S/{fmt(parseFloat(p.monto))} · Izipay retiene S/{(parseFloat(p.monto)*0.05).toFixed(2)})
                              </span>
                            </div>
                          )}
                        </div>
                      ))}
                      <button onClick={()=>setPosPagos(p=>[...p,{met:"Efectivo",monto:""}])}
                        style={{width:"100%",padding:"7px",borderRadius:7,border:`1px dashed ${C.ac}`,background:"transparent",color:C.ac,fontSize:12,fontWeight:700,cursor:"pointer",marginBottom:8}}>
                        + Agregar forma de pago
                      </button>
                      {posCarrito.length>0&&(<>
                        {tienePOS && montoPOS > 0 && (
                          <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:8,padding:"8px 12px",marginBottom:6,fontSize:12}}>
                            <div style={{fontWeight:700,color:C.or,marginBottom:4}}>💳 Resumen de pago:</div>
                            {posPagos.filter(p=>p.met!=="POS Tarjeta"&&parseFloat(p.monto)>0).map((p,i)=>(
                              <div key={i} style={{display:"flex",justifyContent:"space-between",color:C.t2}}>
                                <span>{p.met}</span><span>{fmt(parseFloat(p.monto))}</span>
                              </div>
                            ))}
                            <div style={{display:"flex",justifyContent:"space-between",color:C.or}}>
                              <span>POS datáfono (cliente paga)</span><span>{fmt(montoEnDatafono)}</span>
                            </div>
                            <div style={{borderTop:`1px solid ${C.border}`,marginTop:4,paddingTop:4,display:"flex",justifyContent:"space-between",fontWeight:800,color:C.t1}}>
                              <span>Total tienda recibe</span><span>{fmt(posBase)}</span>
                            </div>
                          </div>
                        )}
                        <div style={{borderRadius:7,padding:"6px 10px",marginBottom:8,fontSize:12,fontWeight:700,background:posFalta>0.01?C.rdBg:C.grBg,color:posFalta>0.01?C.rd:C.gr,display:"flex",justifyContent:"space-between"}}>
                          {posFalta>0.01?<><span>⚠️ Falta</span><span>{fmt(posFalta)}</span></>:<><span>✓ Completo</span>{posVuelto>0.01&&<span>Vuelto: {fmt(posVuelto)}</span>}</>}
                        </div>
                      </>)}
                      <button onClick={posCobrar} disabled={!posCarrito.length||posFalta>0.01}
                        style={{width:"100%",padding:"14px",borderRadius:10,border:"none",background:posCarrito.length&&posFalta<=0.01?"#27ae60":C.bg3,color:posCarrito.length&&posFalta<=0.01?"#fff":C.t4,fontSize:16,fontWeight:800,cursor:posCarrito.length&&posFalta<=0.01?"pointer":"not-allowed",fontFamily:"inherit",letterSpacing:"0.5px",position:"sticky",bottom:8,zIndex:10,boxShadow:posCarrito.length&&posFalta<=0.01?"0 4px 16px #27ae6055":"none"}}>
                        ✅ COBRAR S/ {posTotal.toFixed(2)}
                      </button>
                    </div>
                  </Card>
                </div>
              )}
              </>
            );
          })()}
          {/* ═══════ VENTAS ═══════ */}
          {modActual === "ventas" && (() => {
            // Vendedores solo ven sus propias ventas
            const ventasBase = isAdmin ? ventas : ventas.filter(v => v.vend === user.nombre);
            // Separar ventas reales de anticipos
            const ventasReales = ventasBase.filter(v => !v.esAnticipo);
            const anticiposBase = ventasBase.filter(v => v.esAnticipo);
            const vendedoresDisp = isAdmin ? ["Todos", ...new Set(ventas.map(v => v.vend))] : [user.nombre];
            const ventasFilt = ventasReales.filter(v => {
              const dOk = !vFiltDesde || v.f >= vFiltDesde;
              const hOk = !vFiltHasta || v.f <= vFiltHasta;
              const veOk = !isAdmin || vFiltVend === "Todos" || v.vend === vFiltVend;
              const cOk  = vFiltComp === "Todos" || v.comp === vFiltComp;
              const bOk  = !vFiltBusq || v.cli.toLowerCase().includes(vFiltBusq.toLowerCase()) || v.num.toLowerCase().includes(vFiltBusq.toLowerCase()) || v.items.some(i => i.n.toLowerCase().includes(vFiltBusq.toLowerCase()));
              return dOk && hOk && veOk && cOk && bOk;
            });
            const totalFilt = ventasFilt.reduce((a,v) => a+v.tot, 0);
            const hayFiltro = vFiltDesde || vFiltHasta || vFiltVend !== "Todos" || vFiltComp !== "Todos" || vFiltBusq;
            return (<>
              <PageTitle title="💰 Ventas" sub={`${ventasFilt.length} de ${ventas.length} registros`} action={
                <div style={{display:"flex",gap:8}}>
                  <Btn onClick={() => exportXLSX("ventas", ["Fecha","N° Comprobante","Cliente","Vendedor","Productos","Método Pago","Comprobante","Total"], ventasFilt.map(v => [v.f, v.num, v.cli, v.vend, v.items.map(i=>`${i.n} x${i.q}`).join(" / "), v.mp, v.comp, v.tot]))}>⬇️ Excel</Btn>
                  <Btn variant="primary" onClick={() => setModal("venta")}>+ Nueva Venta</Btn>
                </div>
              } />
              {/* KPIs dinámicos del filtro */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:8, marginBottom:10 }}>
                <KPI icon="💰" label={hayFiltro?"Total filtrado":"Ingresos del mes"} value={fmt(hayFiltro?totalFilt:kpi.ing)} color={C.grL} sub={`${ventasFilt.length} ventas cerradas`}/>
                {isAdmin
                  ? <KPI icon="🎫" label="Ticket promedio" value={fmt(ventasFilt.length ? totalFilt/ventasFilt.length : 0)} color={C.ac}/>
                  : <div style={{ background:"linear-gradient(135deg,#fff7ed,#fed7aa)", border:"1px solid #fdba74", borderRadius:14, padding:"12px 14px", display:"flex", flexDirection:"column", justifyContent:"center" }}>
                      <div style={{ fontSize:10, fontWeight:700, color:"#c05621", textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:4 }}>💬 Frase del día</div>
                      <div style={{ fontSize:12, color:"#7c2d12", fontWeight:600, lineHeight:1.4 }}>{[
                        "💪 ¡Cada cliente es una oportunidad! Sonríe, escucha y cierra esa venta — tu comisión te lo agradecerá.",
                        "🌟 Un espacio limpio y ordenado genera confianza. ¡El cliente que se siente bien, compra más!",
                        "🔥 ¡Hoy puede ser tu mejor día! Cada 'no' te acerca al próximo 'sí'. ¡Sigue adelante!",
                        "😊 Tu actitud es tu mejor herramienta. Un vendedor amable y atento siempre supera sus metas.",
                        "🚀 ¡Las comisiones no se esperan, se conquistan! Da el 100% en cada atención.",
                        "🏆 Los mejores vendedores no nacen — se forjan siendo amables, honestos y persistentes.",
                        "✨ ¡Mantén tu área impecable! Un ambiente agradable invita al cliente a quedarse y comprar.",
                        "🎯 Conoce cada producto, responde con seguridad y el cliente confiará en ti. ¡Tú puedes!",
                        "💡 Una buena atención hoy es un cliente fiel mañana. ¡Invierte en cada interacción!",
                        "🤝 Sé el vendedor que a ti te gustaría encontrar: honesto, amable y siempre dispuesto a ayudar.",
                        "⭐ ¡Tu energía se contagia! Si tú crees en el producto, el cliente también creerá en él.",
                        "🌺 Un 'gracias, fue un placer atenderle' puede traer al cliente de vuelta. ¡Cuida cada detalle!",
                      ][new Date().getDate() % 12]}</div>
                    </div>
                }
                <KPI icon="📄" label="Facturas" value={ventasFilt.filter(v=>v.comp==="Factura").length} color={C.bl}/>
                <KPI icon="🧾" label="Boletas" value={ventasFilt.filter(v=>v.comp==="Boleta").length} color={C.grL}/>
              </div>
              {/* Anticipos pendientes */}
              {anticiposBase.length > 0 && (
                <div style={{ background:"#fffbeb", border:"1px solid #f6e05e66", borderRadius:10, padding:"10px 16px", marginBottom:12, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <div>
                    <div style={{ fontSize:12, fontWeight:700, color:"#744210" }}>📥 Anticipos / Adelantos de pedidos ({anticiposBase.length})</div>
                    <div style={{ fontSize:11, color:"#a0aec0", marginTop:2 }}>Dinero recibido de pedidos en curso — no se cuentan como venta hasta que el pedido se entrega</div>
                  </div>
                  <div style={{ fontSize:16, fontWeight:800, color:"#c05621", flexShrink:0, marginLeft:12 }}>
                    {fmt(anticiposBase.reduce((a,v)=>a+v.tot,0))}
                  </div>
                </div>
              )}
              {/* Filtros */}
              <Card style={{ padding:"10px 14px", marginBottom:12 }}>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"flex-end" }}>
                  <div style={{ position:"relative", flex:1, minWidth:180 }}>
                    <span style={{ position:"absolute", left:9, top:"50%", transform:"translateY(-50%)", color:C.t4, fontSize:13 }}>🔍</span>
                    <input value={vFiltBusq} onChange={e => setVFiltBusq(e.target.value)} placeholder="Buscar cliente, comprobante, producto..."
                      style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px 8px 28px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
                  </div>
                  <Inp label="Desde" type="date" value={vFiltDesde} onChange={e => setVFiltDesde(e.target.value)} style={{ width:140 }}/>
                  <Inp label="Hasta" type="date" value={vFiltHasta} onChange={e => setVFiltHasta(e.target.value)} style={{ width:140 }}/>
                  {isAdmin && <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Vendedor</label>
                    <select value={vFiltVend} onChange={e => setVFiltVend(e.target.value)}
                      style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                      {vendedoresDisp.map(v => <option key={v}>{v}</option>)}
                    </select>
                  </div>}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Comprobante</label>
                    <select value={vFiltComp} onChange={e => setVFiltComp(e.target.value)}
                      style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                      {["Todos","Boleta","Factura","Nota de Venta"].map(c => <option key={c}>{c}</option>)}
                    </select>
                  </div>
                  {hayFiltro && (
                    <button onClick={() => { setVFiltDesde(""); setVFiltHasta(""); setVFiltVend("Todos"); setVFiltComp("Todos"); setVFiltBusq(""); }}
                      style={{ background:"none", border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:12, color:C.t3, cursor:"pointer", fontFamily:"inherit", marginBottom:0, alignSelf:"flex-end" }}>✕ Limpiar</button>
                  )}
                </div>
                {hayFiltro && (
                  <div style={{ marginTop:8, paddingTop:8, borderTop:`1px solid ${C.border}`, fontSize:12, color:C.t2, display:"flex", gap:16 }}>
                    <span>📊 {ventasFilt.length} ventas filtradas</span>
                    <span style={{ fontWeight:700, color:C.grL }}>Total: {fmt(totalFilt)}</span>
                    <span>Ticket prom: {fmt(ventasFilt.length ? totalFilt/ventasFilt.length : 0)}</span>
                    {vFiltVend !== "Todos" && <span>Vendedor: <strong>{vFiltVend}</strong></span>}
                  </div>
                )}
              </Card>
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Fecha</TH><TH>Cliente</TH>{isAdmin && <TH>Vendedor</TH>}<TH>Productos</TH><TH>Pago</TH><TH>Comprobante</TH><TH right>Total</TH><TH></TH></tr></thead>
                  <tbody>
                    {ventasFilt.map(v => (
                      <tr key={v.id}>
                        <TD>{v.f}</TD>
                        <TD bold color={C.t1}>{v.cli}</TD>
                        {isAdmin && <TD>{v.vend}</TD>}
                        <TD sm>{v.items.map(i => `${i.n} ×${i.q}`).join(", ")}</TD>
                        <TD><Badge color={v.mp.includes("POS") ? C.or : C.bl}>{v.mp}</Badge></TD>
                        <TD><Badge color={v.comp === "Factura" ? C.bl : v.comp === "Nota de Venta" ? C.pu : C.grL}>{v.num}</Badge></TD>
                        <TD right bold color={C.ac}>{fmt(v.tot)}</TD>
                        <TD>
                          <button onClick={() => setVoucher({ items:v.items, subtotal:v.tot, desc:0, recPOS:0, total:v.tot, pagos:[{met:v.mp,monto:v.tot}], vuelto:0, cliente:v.cli, vend:v.vend, comp:v.comp, num:v.num, fecha:v.f })}
                            style={{ padding:"3px 8px", borderRadius:6, border:`1px solid ${C.border}`, background:C.bg, color:C.t2, fontSize:11, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}>
                            🧾 Ver
                          </button>
                        </TD>
                      </tr>
                    ))}
                    {ventasFilt.length === 0 && (
                      <tr><td colSpan={isAdmin ? 8 : 7} style={{ padding:"30px", textAlign:"center", color:C.t4, fontSize:13 }}>Sin resultados con los filtros aplicados</td></tr>
                    )}
                    {ventasFilt.length > 0 && (
                      <tr style={{ background:C.bg }}>
                        <td colSpan={isAdmin ? 7 : 6} style={{ padding:"10px 12px", fontWeight:700, fontSize:13, color:C.t2, textAlign:"right" }}>TOTAL ({ventasFilt.length} ventas)</td>
                        <td style={{ padding:"10px 12px", fontWeight:800, fontSize:15, color:C.grL, textAlign:"right" }}>{fmt(totalFilt)}</td>
                        <td/>
                      </tr>
                    )}
                  </tbody>
                </table>
              </Card>
            </>);
          })()}

          {/* ═══════ INVENTARIO ═══════ */}
          {modActual === "caja" && (() => {
            const ventasDia = ventas.filter(v => v.f === HOY && !v.esAnticipo && (isAdmin || v.vend === user.nombre));
            const anticiposDia = ventas.filter(v => v.f === HOY && v.esAnticipo && (isAdmin || v.vend === user.nombre));
            const totalDia = ventasDia.reduce((a,v) => a+v.tot, 0);
            const totalAnt = anticiposDia.reduce((a,v) => a+v.tot, 0);
            const metodos = {};
            ventasDia.forEach(v => {
              v.mp.split("+").forEach(mp => {
                const m = mp.trim();
                metodos[m] = (metodos[m]||0) + v.tot / v.mp.split("+").length;
              });
            });
            const yaCerrado = cierres.some(c => c.f === HOY && (isAdmin ? true : c.vendedor === user.nombre));

            // Admin: resumen por vendedor
            const vendedoresDia = isAdmin ? [...new Set(ventasDia.map(v => v.vend))] : [user.nombre];
            const resumenPorVendedor = vendedoresDia.map(vend => {
              const vs = ventasDia.filter(v => v.vend === vend);
              const total = vs.reduce((a,v) => a+v.tot, 0);
              const mets = {};
              vs.forEach(v => v.mp.split("+").forEach(mp => { const m=mp.trim(); mets[m]=(mets[m]||0)+v.tot/v.mp.split("+").length; }));
              return { vend, total, count: vs.length, mets };
            });

            const hacerCierre = () => {
              if (yaCerrado) { showToast("Ya existe un cierre para hoy", "err"); return; }
              if (ventasDia.length === 0 && anticiposDia.length === 0) { showToast("No hay ventas hoy para cerrar", "err"); return; }
              const nc = {
                id: "CJ"+uid(), f: HOY, hora: new Date().toLocaleTimeString("es-PE"),
                vendedor: isAdmin ? "Admin" : user.nombre,
                ventas: ventasDia.map(v => v.id),
                anticipos: anticiposDia.map(v => v.id),
                total: totalDia, totalAnticipo: totalAnt,
                metodos, totalGeneral: totalDia + totalAnt
              };
              setCierres(cs => { sbSave("cierres", nc); return [nc, ...cs]; });
              showToast(`✓ Cierre registrado · S/${(totalDia+totalAnt).toFixed(2)} total del día`);
            };

            return (<>
              <PageTitle title="🏦 Cierre de Caja" sub={`${HOY} · ${isAdmin ? "Todas las ventas" : user.nombre}`} />
              {yaCerrado && (
                <div style={{ background:"#f0fff4", border:"1px solid #68d391", borderRadius:10, padding:"10px 16px", marginBottom:16, color:"#276749", fontSize:13 }}>
                  ✅ Ya realizaste el cierre de caja de hoy
                </div>
              )}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:12, marginBottom:20 }}>
                <KPI icon="🧾" label="Ventas hoy" value={ventasDia.length} color={C.grL} sub={`S/${totalDia.toFixed(2)}`}/>
                <KPI icon="💰" label="Total ventas" value={`S/${totalDia.toFixed(2)}`} color={C.ac}/>
                <KPI icon="📥" label="Anticipos" value={`S/${totalAnt.toFixed(2)}`} color={C.bl} sub={`${anticiposDia.length} pagos`}/>
                <KPI icon="💵" label="Total en caja" value={`S/${(totalDia+totalAnt).toFixed(2)}`} color={C.gr}/>
              </div>

              {/* Admin: resumen por vendedor */}
              {isAdmin && resumenPorVendedor.length > 0 && (
                <Card title="📊 Resumen por vendedor" style={{ marginBottom:16 }}>
                  {resumenPorVendedor.map(r => (
                    <div key={r.vend} style={{ padding:"10px 0", borderBottom:`1px solid ${C.border}` }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                        <span style={{ fontWeight:700, color:C.t1, fontSize:14 }}>{r.vend}</span>
                        <span style={{ fontWeight:800, color:C.ac, fontSize:14 }}>S/ {r.total.toFixed(2)}</span>
                      </div>
                      <div style={{ display:"flex", gap:12, flexWrap:"wrap" }}>
                        <span style={{ fontSize:11, color:C.t3 }}>{r.count} ventas</span>
                        {Object.entries(r.mets).map(([mp, tot]) => (
                          <span key={mp} style={{ fontSize:11, color:C.t2, background:C.bg, padding:"2px 8px", borderRadius:5 }}>{mp}: S/{(tot as number).toFixed(0)}</span>
                        ))}
                      </div>
                    </div>
                  ))}
                </Card>
              )}

              <Card title="Desglose por método de pago" style={{ marginBottom:16 }}>
                {Object.keys(metodos).length === 0
                  ? <div style={{ color:C.t3, fontSize:13, padding:"8px 0" }}>No hay ventas registradas hoy</div>
                  : Object.entries(metodos).map(([mp, tot]) => (
                    <div key={mp} style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:`1px solid ${C.border}`, fontSize:14 }}>
                      <span style={{ color:C.t2 }}>{mp}</span>
                      <strong style={{ color:C.ac }}>S/ {(tot as number).toFixed(2)}</strong>
                    </div>
                  ))
                }
              </Card>
              <Card title={`Ventas del día (${ventasDia.length})`} style={{ marginBottom:16 }}>
                {ventasDia.length === 0
                  ? <div style={{ color:C.t3, fontSize:13 }}>Sin ventas hoy</div>
                  : ventasDia.map(v => (
                    <div key={v.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:13 }}>
                      <div>
                        <span style={{ fontWeight:600, color:C.t1 }}>{v.num}</span>
                        {isAdmin && <span style={{ fontSize:11, background:C.acBg, color:C.ac, padding:"1px 6px", borderRadius:4, marginLeft:6 }}>{v.vend}</span>}
                        <span style={{ color:C.t3, marginLeft:8 }}>{v.cli}</span>
                        <span style={{ color:C.t4, marginLeft:8, fontSize:11 }}>{v.items.map(i=>i.n).join(", ")}</span>
                      </div>
                      <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                        <span style={{ fontSize:11, color:C.t3 }}>{v.mp}</span>
                        <strong style={{ color:C.ac }}>S/ {v.tot.toFixed(2)}</strong>
                      </div>
                    </div>
                  ))
                }
              </Card>
              {cierres.filter(c => isAdmin || c.vendedor === user.nombre).length > 0 && (
                <Card title="Historial de cierres anteriores">
                  {cierres.filter(c => isAdmin || c.vendedor === user.nombre).slice(0,15).map(c => (
                    <div key={c.id} style={{ display:"flex", justifyContent:"space-between", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:13 }}>
                      <span style={{ color:C.t2 }}>{c.f} · {c.hora}</span>
                      <span style={{ color:C.t3 }}>{c.vendedor}</span>
                      <strong style={{ color:C.ac }}>S/ {(c.totalGeneral||c.total||0).toFixed(2)}</strong>
                    </div>
                  ))}
                </Card>
              )}
              {!yaCerrado && (
                <div style={{ marginTop:20 }}>
                  <Btn variant="primary" full onClick={hacerCierre}>🔒 Realizar Cierre de Caja del Día</Btn>
                </div>
              )}
            </>);
          })()}

          {/* ═══════ INVENTARIO ═══════ */}
          {modActual === "inventario" && (() => {
            const invFiltrado = prods.filter(p => {
              const stk = getTotalStk(p);
              const bOk  = !invBusq || p.n.toLowerCase().includes(invBusq.toLowerCase()) || p.id.toLowerCase().includes(invBusq.toLowerCase()) || p.cols.some(c => c.toLowerCase().includes(invBusq.toLowerCase()));
              const cOk  = invCat === "Todos" || p.cat === invCat;
              const eOk  = invEst === "Todos" || (invEst === "Crítico" && stk <= p.min && stk > 0) || (invEst === "Sin stock" && stk === 0) || (invEst === "OK" && stk > p.min);
              const uOk  = !invUbi || invUbi === "Todas" || (getLocs(p)[invUbi]||0) > 0;
              return bOk && cOk && eOk && uOk;
            });
            const CATS_INV = ["Ropero","Zapatero","Cómoda","Mesa de Noche","Cabecera","Otro"];
            const guardarNuevoProd = (np) => {
              const nuevo = { ...np, id: np.id || "P"+String(prods.length+1).padStart(3,"0"), v:[0,0,0,0,0,0,0,0,0,0,0,0] };
              setProds(ps => [...ps, nuevo]);
              sbSave("productos", nuevo);
              const nk = { id:"K"+uid(), f:HOY, pid:nuevo.id, prod:nuevo.n, col:"Todos", tipo:"Ajuste", desc:"Alta de producto", ent:nuevo.stk, sal:0, saldo:nuevo.stk, costo:nuevo.c };
              setKardex(k => { sbSave("kardex", nk); return [nk, ...k]; });
              showToast(`✓ ${nuevo.n} registrado con código ${nuevo.id}`);
              setModal(null);
            };
            return (<>
              <PageTitle title="📦 Inventario" sub={`${invFiltrado.length} de ${prods.length} productos`}
                action={
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    <div style={{ display:"flex", background:C.bg2, borderRadius:8, padding:2, gap:2 }}>
                      {[["tabla","☰"],["cards","⊞"]].map(([v,ico]) => (
                        <button key={v} onClick={() => setInvVista(v)}
                          style={{ padding:"5px 10px", borderRadius:6, border:"none", background:invVista===v?C.white:"transparent", color:invVista===v?C.ac:C.t3, fontSize:14, cursor:"pointer", fontWeight:700, boxShadow:invVista===v?C.sh:"none" }}>
                          {ico}
                        </button>
                      ))}
                    </div>
              {/* Banner de stock bajo con botón WA */}
              {(() => {
                const critProds = prods.filter(p => getTotalStk(p) <= p.min);
                if (critProds.length === 0) return null;
                const msg = `🚨 *ALERTA STOCK BAJO - ${new Date().toLocaleDateString("es-PE")}*\n\n` +
                  critProds.map(p => {
                    const stk = getTotalStk(p);
                    const locs = getLocs(p);
                    const detalle = UBICACIONES.map(u => {
                      const tot = Object.entries(locs).filter(([k])=>k.startsWith(u+"|")).reduce((a,[,v])=>a+v,0);
                      return tot>0 ? `${u.replace("Taller/Almacén","Almacén")}: ${tot}` : null;
                    }).filter(Boolean).join(", ");
                    return `• ${p.ico} *${p.n}*\n  Stock: ${stk} uds (mín: ${p.min}) | ${detalle||"Sin stock"}`;
                  }).join("\n\n") +
                  `\n\n_Por favor reabastecer a la brevedad_`;
                return (
                  <a href={`https://wa.me/?text=${encodeURIComponent(msg)}`} target="_blank" rel="noopener noreferrer"
                    style={{ display:"flex", alignItems:"center", gap:6, padding:"7px 12px", borderRadius:8, background:"#25D366", color:"#fff", textDecoration:"none", fontSize:12, fontWeight:700, whiteSpace:"nowrap" }}>
                    📲 Alertar stock bajo ({critProds.length})
                  </a>
                );
              })()}
                    <Btn onClick={() => exportXLSX("inventario", ["Código","Nombre","Categoría","Tipo","Colores","Ubicación","Stock","Mínimo","Precio venta","Costo","Margen %","Estado"], prods.map(p => { const mg=Math.round(((p.p-p.c)/p.p)*100); const est=p.stk===0?"Sin stock":p.stk<=p.min?"Crítico":p.stk<=p.min*1.5?"Bajo":"OK"; return [p.id,p.n,p.cat,p.tipo,p.cols.join("/"),p.ubi||"—",p.stk,p.min,p.p,p.c,mg,est]; }))}>⬇️ Excel</Btn>
                    <Btn variant="primary" onClick={() => setModal({ tipo:"nuevoProd", form:{ id:"", n:"", cat:"Ropero", tipo:"Fabricado", ico:"🪞", img:"", barcode:"", cols:["Blanco"], stk:0, min:2, p:0, c:0, ubi:UBICACIONES[0] } })}>+ Nuevo Producto</Btn>
                  </div>
                }
              />
              {/* Modal nuevo producto */}
              {modal?.tipo === "nuevoProd" && (
                <Modal onClose={() => setModal(null)} wide>
                  <ModalTitle>📦 Registrar Nuevo Producto</ModalTitle>
                  <NuevoProdForm form={modal.form} cats={CATS_INV} onSave={guardarNuevoProd} onCancel={() => setModal(null)}/>
                </Modal>
              )}
              {/* Modal edición de producto */}
              {prodEdit && (
                <ProdEditModal prod={prodEdit} onSave={pActual => {
                  const viejo = prods.find(p => p.id === pActual.id);
                  setProds(ps => ps.map(p => p.id === pActual.id ? { ...p, ...pActual } : p));
                  sbSave("productos", pActual);
                  if (viejo && viejo.c !== pActual.c) {
                    const nk = { id:"K"+uid(), f:HOY, pid:pActual.id, prod:pActual.n, col:"Todos", tipo:"Ajuste",
                      desc:`Ajuste de costo: S/${viejo.c} → S/${pActual.c}`, ent:0, sal:0, saldo:pActual.stk, costo:pActual.c };
                    setKardex(k => { sbSave("kardex", nk); return [nk, ...k]; });
                  }
                  showToast(`✓ ${pActual.n} actualizado`);
                  setProdEdit(null);
                }} onCancel={() => setProdEdit(null)}/>
              )}
              {/* Escáner de código de barras */}
              {invScan && (
                <Card style={{ padding:"12px 16px", marginBottom:10, background:C.acBg, border:`1px solid ${C.ac}44` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
                    <div style={{ fontSize:11, fontWeight:700, color:C.ac }}>📷 ESCÁNER · Ingresa el código del producto</div>
                    <Btn sm onClick={() => { setInvScan(false); setInvScanVal(""); setCamScan(false); }}>✕</Btn>
                  </div>
                  <div style={{ display:"flex", gap:8, marginBottom:8 }}>
                    <input ref={invScanRef} autoFocus value={invScanVal} onChange={e => setInvScanVal(e.target.value)}
                      onKeyDown={e => { if (e.key === "Enter" && invScanVal) { const busq = invScanVal.toUpperCase(); const porBarcode = prods.find(p => p.barcode === busq); setInvBusq(porBarcode ? porBarcode.n : busq); setInvScan(false); setCamScan(false); setInvScanVal(""); }}}
                      style={{ flex:1, background:C.white, border:`2px solid ${C.ac}`, borderRadius:8, padding:"8px 12px", fontSize:14, fontWeight:700, outline:"none", fontFamily:"monospace", letterSpacing:"2px", textAlign:"center", color:C.t1 }}
                      placeholder="Escanea con lector o escribe código → Enter"/>
                  </div>
                  <button onClick={() => setCamScan(c => !c)}
                    style={{ width:"100%", padding:"8px", borderRadius:8, border:`1px solid ${camScan?C.ac:C.border}`, background:camScan?C.acBg:C.bg, color:camScan?C.ac:C.t2, fontSize:12, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
                    {camScan ? "⏹ Cerrar cámara" : "📱 Usar cámara del teléfono"}
                  </button>
                  {camScan && (
                    <BarcodeCamera onDetect={code => {
                      const porBarcode = prods.find(p => p.barcode === code);
                      setInvBusq(porBarcode ? porBarcode.n : code);
                      setInvScan(false); setCamScan(false); setInvScanVal("");
                      showToast(porBarcode ? `✓ ${porBarcode.n}` : `Código: ${code}`);
                    }} onClose={() => setCamScan(false)}/>
                  )}
                </Card>
              )}
              <Card style={{ padding:"10px 14px", marginBottom:12 }}>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
                  <div style={{ flex:1, minWidth:200, position:"relative" }}>
                    <span style={{ position:"absolute", left:10, top:"50%", transform:"translateY(-50%)", color:C.t4, fontSize:14, pointerEvents:"none" }}>🔍</span>
                    <input value={invBusq} onChange={e => setInvBusq(e.target.value)} placeholder="Buscar por nombre, código o color..."
                      style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px 8px 32px", fontSize:13, outline:"none", boxSizing:"border-box", color:C.t1, fontFamily:"inherit" }}/>
                  </div>
                  <button onClick={() => { setInvScan(s => !s); setInvScanVal(""); }}
                    style={{ padding:"8px 12px", borderRadius:8, border:`1px solid ${invScan?C.ac:C.border}`, background:invScan?C.acBg:C.white, color:invScan?C.ac:C.t3, cursor:"pointer", fontSize:18, title:"Modo escáner" }}>
                    📷
                  </button>
                  <select value={invCat} onChange={e => setInvCat(e.target.value)}
                    style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                    {["Todos","Ropero","Zapatero","Cómoda","Mesa de Noche","Cabecera"].map(c => <option key={c}>{c}</option>)}
                  </select>
                  <select value={invEst} onChange={e => setInvEst(e.target.value)}
                    style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                    {["Todos","OK","Crítico","Sin stock"].map(e => <option key={e}>{e}</option>)}
                  </select>
                  <select value={invUbi} onChange={e => setInvUbi(e.target.value)}
                    style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                    <option value="Todas">📍 Todas las ubicaciones</option>
                    {UBICACIONES.map(u => <option key={u}>{u}</option>)}
                  </select>
                  {(invBusq || invCat !== "Todos" || invEst !== "Todos" || invUbi !== "Todas") && (
                    <button onClick={() => { setInvBusq(""); setInvCat("Todos"); setInvEst("Todos"); setInvUbi("Todas"); setInvScan(false); }}
                      style={{ background:"none", border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:12, color:C.t3, cursor:"pointer", fontFamily:"inherit" }}>✕ Limpiar</button>
                  )}
                </div>
              </Card>
              {/* VISTA TABLA */}
              {invVista === "tabla" && (
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Producto</TH>{isAdmin&&<TH>Tipo</TH>}<TH>Colores</TH><TH>Stock por ubicación</TH><TH>Cód. Barras</TH><TH right>Total</TH><TH right>Mín</TH><TH right>Precio venta</TH>{isAdmin&&<TH right>Costo unit.</TH>}{isAdmin&&<TH right>Margen</TH>}<TH>Estado</TH>{isAdmin&&<TH></TH>}</tr></thead>
                  <tbody>
                    {invFiltrado.map(p => {
                      const stk  = getTotalStk(p);
                      const mg   = p.p > 0 ? Math.round(((p.p-p.c)/p.p)*100) : 0;
                      const est  = stk === 0 ? ["Sin stock",C.rd] : stk <= p.min ? ["Crítico",C.rd] : stk <= p.min*1.5 ? ["Bajo",C.or] : ["OK",C.grL];
                      return (
                        <tr key={p.id}>
                          <TD bold color={C.t1}>{p.ico} {p.n} <span style={{fontSize:10,color:C.t4,fontWeight:400}}>{p.id}</span></TD>
                          {isAdmin && <TD><Badge color={p.tipo === "Fabricado" ? C.grL : C.bl}>{p.tipo}</Badge></TD>}
                          <TD sm color={C.t3}>{p.cols.join(", ")}</TD>
                          <TD>
                            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                              {UBICACIONES.map(ubi => {
                                const colsAqui = (p.cols||[]).filter(col => (getLocs(p)[locKey(ubi,col)]||0) > 0);
                                if (colsAqui.length === 0) return null;
                                const ubiColor = ubi==="Tienda Principal"?C.bl:ubi==="Tienda Pasaje"?C.pu:C.or;
                                const totalUbi = colsAqui.reduce((a,col)=>a+(getLocs(p)[locKey(ubi,col)]||0),0);
                                return (
                                  <details key={ubi} style={{ cursor:"pointer" }}>
                                    <summary style={{ listStyle:"none", display:"flex", alignItems:"center", gap:6, padding:"3px 0" }}>
                                      <span style={{ fontSize:10, fontWeight:700, color:"#fff", background:ubiColor, padding:"2px 8px", borderRadius:99, whiteSpace:"nowrap" }}>
                                        📍 {ubi.replace("Taller/Almacén","Almacén")} · {totalUbi} uds
                                      </span>
                                      <span style={{ fontSize:10, color:C.t4 }}>▾</span>
                                    </summary>
                                    <div style={{ paddingLeft:8, marginTop:4, display:"flex", gap:6, flexWrap:"wrap" }}>
                                      {colsAqui.map(col => (
                                        <span key={col} style={{ fontSize:10, color:C.t2, display:"flex", alignItems:"center", gap:3, background:C.bg, padding:"2px 6px", borderRadius:5 }}>
                                          <span style={{ width:7, height:7, borderRadius:"50%", background:HEX_COLOR[col]||"#888", display:"inline-block" }}/>
                                          {col}: <strong>{getLocs(p)[locKey(ubi,col)]}</strong>
                                        </span>
                                      ))}
                                    </div>
                                  </details>
                                );
                              })}
                              {getTotalStk(p)===0 && <span style={{fontSize:11,color:C.rd,fontWeight:700}}>Sin stock</span>}
                            </div>
                          </TD>
                          <TD sm color={C.t3}>{p.barcode ? <span style={{fontFamily:"monospace",fontSize:11,background:C.bg,padding:"2px 6px",borderRadius:4}}>▌{p.barcode}</span> : <span style={{color:C.t4,fontStyle:"italic",fontSize:11}}>sin código</span>}</TD>
                          <TD right bold color={est[1]}>{stk}</TD>
                          <TD right color={C.t3}>{p.min}</TD>
                          <TD right bold color={C.t1}>{fmt(p.p)}</TD>
                          {isAdmin && <TD right><span style={{ fontWeight:700, color:C.or, background:C.orBg, padding:"2px 7px", borderRadius:5, fontSize:12 }}>{fmt(p.c)}</span></TD>}
                          {isAdmin && <TD right><Badge color={mg >= 40 ? C.grL : mg >= 25 ? C.or : C.rd}>{mg}%</Badge></TD>}
                          <TD><Badge color={est[1]}>{est[0]}</Badge></TD>
                          {isAdmin && <TD>
                            <button onClick={() => setProdEdit({ ...p })}
                              style={{ padding:"4px 10px", borderRadius:6, border:`1px solid ${C.border}`, background:C.bg, color:C.t2, fontSize:11, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}>
                              ✏️ Editar
                            </button>
                          </TD>}
                        </tr>
                      );
                    })}
                    {invFiltrado.length === 0 && (
                      <tr><td colSpan={12} style={{ padding:"30px", textAlign:"center", color:C.t4, fontSize:13 }}>Sin resultados{invBusq ? ` para "${invBusq}"` : ""}</td></tr>
                    )}
                  </tbody>
                </table>
              </Card>
              )}

              {/* VISTA CARDS — como el POS */}
              {invVista === "cards" && (
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))", gap:12 }}>
                  {invFiltrado.map(p => {
                    const mg  = Math.round(((p.p-p.c)/p.p)*100);
                    const est = p.stk === 0 ? ["Sin stock",C.rd] : p.stk <= p.min ? ["Crítico",C.rd] : p.stk <= p.min*1.5 ? ["Bajo",C.or] : ["OK",C.grL];
                    const color = CAT_C[p.cat] || C.ac;
                    const ubiColor = p.ubi === "Tienda Principal" ? C.bl : p.ubi === "Tienda Pasaje" ? C.pu : C.or;
                    return (
                      <div key={p.id} style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:14, overflow:"hidden", boxShadow:C.sh }}>
                        {/* Imagen */}
                        <div style={{ height:120, background: p.img ? "transparent" : color+"10", display:"flex", alignItems:"center", justifyContent:"center", overflow:"hidden", position:"relative" }}>
                          {p.img ? <img src={p.img} alt={p.n} style={{ width:"100%", height:"100%", objectFit:"cover" }}/> : <span style={{ fontSize:44 }}>{p.ico}</span>}
                          <div style={{ position:"absolute", top:8, left:8, background:color+"dd", color:"#fff", fontSize:9, fontWeight:700, padding:"2px 7px", borderRadius:99 }}>{p.cat}</div>
                          <div style={{ position:"absolute", top:8, right:8 }}><Badge color={est[1]}>{est[0]}</Badge></div>
                        </div>
                        <div style={{ padding:"10px 12px 12px" }}>
                          <div style={{ fontSize:12, fontWeight:700, color:C.t1, marginBottom:3, lineHeight:1.3 }}>{p.n}</div>
                          <div style={{ fontSize:10, color:C.t4, marginBottom:6, fontFamily:"monospace" }}>{p.id}</div>
                          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                            <span style={{ fontSize:14, fontWeight:800, color }}>{fmt(p.p)}</span>
                            <span style={{ fontSize:11, fontWeight:700, color: est[1] }}>{p.stk} uds</span>
                          </div>
                          <div style={{ fontSize:10, color:C.t3, marginBottom:8 }}><Badge color={ubiColor} style={{fontSize:9}}>📍 {p.ubi||"—"}</Badge></div>
                          <div style={{ display:"flex", gap:4, flexWrap:"wrap", marginBottom:8 }}>
                            {p.cols.map(c => <span key={c} style={{ display:"flex", alignItems:"center", gap:3, fontSize:10, color:C.t3 }}><div style={{width:8,height:8,borderRadius:"50%",background:HEX_COLOR[c]||"#888",flexShrink:0}}/>{c}</span>)}
                          </div>
                          <button onClick={() => setProdEdit({ ...p })}
                            style={{ width:"100%", padding:"6px", borderRadius:7, border:`1px solid ${C.border}`, background:C.bg, color:C.t2, fontSize:11, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>
                            ✏️ Editar
                          </button>
                        </div>
                      </div>
                    );
                  })}
                  {invFiltrado.length === 0 && (
                    <div style={{ gridColumn:"1/-1", textAlign:"center", padding:"40px 0", color:C.t4 }}>Sin resultados</div>
                  )}
                </div>
              )}
            </>);
          })()}

          {/* ═══════ KARDEX ═══════ */}
          {modActual === "kardex" && (() => {
            const kardexFilt = kardex.filter(k => {
              const dOk  = !kFiltDesde || k.f >= kFiltDesde;
              const hOk  = !kFiltHasta || k.f <= kFiltHasta;
              const pOk  = !kFiltProd  || k.pid === kFiltProd;
              const cOk  = !kFiltCol   || k.col === kFiltCol;
              const tOk  = kFiltTipo === "Todos" || k.tipo === kFiltTipo;
              const bOk  = !kFiltBusq  || k.prod.toLowerCase().includes(kFiltBusq.toLowerCase()) || (k.desc||"").toLowerCase().includes(kFiltBusq.toLowerCase());
              const uOk  = !kFiltUbi   || (k.desc||"").includes(kFiltUbi);
              return dOk && hOk && pOk && cOk && tOk && bOk && uOk;
            });
            const totalEnt = kardexFilt.reduce((a,k) => a+(k.ent||0), 0);
            const totalSal = kardexFilt.reduce((a,k) => a+(k.sal||0), 0);
            const prodSel  = prods.find(p => p.id === kFiltProd);
            const limpiar  = () => { setKFiltProd(""); setKFiltCol(""); setKFiltTipo("Todos"); setKFiltBusq(""); setKFiltDesde(""); setKFiltHasta(""); setKFiltUbi(""); };
            return (
              <>
              <PageTitle title="📒 Kardex" sub="Historial de movimientos por producto, color y local"
                action={
                  <Btn onClick={() => exportXLSX("kardex", ["Fecha","Producto","Color","Local","Tipo","Descripción","Entrada","Salida","Saldo","Costo unit."],
                    kardexFilt.map(k=>[k.f,k.prod,k.col||"",k.ubi||"",k.tipo,k.desc,k.ent||"",k.sal||"",k.saldo,k.costo]))}>
                    ⬇️ Excel
                  </Btn>
                }/>

              {/* Panel de filtros estilo profesional */}
              <Card style={{ marginBottom:14, padding:"14px 16px" }}>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))", gap:10, alignItems:"end" }}>

                  {/* Búsqueda por texto */}
                  <div style={{ gridColumn: "1 / -1" }}>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>🔍 Buscar por nombre de artículo</label>
                    <input type="text" value={kFiltBusq||""} onChange={e => setKFiltBusq(e.target.value)}
                      placeholder="Escribe el nombre del mueble... ej: Ropero, Zapatero, Cómoda"
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltBusq?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
                  </div>

                  {/* Rango de fechas */}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Desde</label>
                    <input type="date" value={kFiltDesde} onChange={e => setKFiltDesde(e.target.value)}
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltDesde?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
                  </div>
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Hasta</label>
                    <input type="date" value={kFiltHasta} onChange={e => setKFiltHasta(e.target.value)}
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltHasta?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
                  </div>

                  {/* Producto */}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Artículo</label>
                    <select value={kFiltProd} onChange={e => { setKFiltProd(e.target.value); setKFiltCol(""); }}
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltProd?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, boxSizing:"border-box" }}>
                      <option value="">Todos los artículos</option>
                      {prods.map(p => <option key={p.id} value={p.id}>{p.ico} {p.n}</option>)}
                    </select>
                  </div>

                  {/* Color */}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Color</label>
                    <select value={kFiltCol} onChange={e => setKFiltCol(e.target.value)} disabled={!kFiltProd}
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltCol?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:kFiltProd?"pointer":"not-allowed", color:C.t1, boxSizing:"border-box", opacity:kFiltProd?1:0.6 }}>
                      <option value="">Todos los colores</option>
                      {(prodSel?.cols||[]).map(c => <option key={c}>{c}</option>)}
                    </select>
                  </div>

                  {/* Local */}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Local / Almacén</label>
                    <select value={kFiltUbi} onChange={e => setKFiltUbi(e.target.value)}
                      style={{ width:"100%", background:C.white, border:`1px solid ${kFiltUbi?C.ac:C.border}`, borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, boxSizing:"border-box" }}>
                      <option value="">Todos los locales</option>
                      {UBICACIONES.map(u => <option key={u} value={u}>{u}</option>)}
                    </select>
                  </div>

                  {/* Tipo movimiento */}
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Tipo movimiento</label>
                    <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                      {["Todos","Venta","Compra","Fabricación","Ajuste"].map(t => (
                        <button key={t} onClick={() => setKFiltTipo(t)}
                          style={{ padding:"6px 10px", borderRadius:6, border:`1px solid ${kFiltTipo===t?C.ac:C.border}`, background:kFiltTipo===t?C.acBg:"transparent", color:kFiltTipo===t?C.ac:C.t3, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                          {t}
                        </button>
                      ))}
                    </div>
                  </div>

                </div>

                {/* Botones acción */}
                <div style={{ display:"flex", gap:8, marginTop:12, alignItems:"center" }}>
                  <Btn variant="primary" onClick={() => {/* ya filtra en tiempo real */}}>🔍 Buscar</Btn>
                  <Btn onClick={limpiar}>✕ Limpiar</Btn>
                  {(kFiltDesde||kFiltHasta||kFiltProd||kFiltCol||kFiltTipo!=="Todos"||kFiltUbi) && (
                    <span style={{ fontSize:11, color:C.t3 }}>{kardexFilt.length} resultado(s)</span>
                  )}
                </div>
              </Card>

              {/* Tabla */}
              <Card style={{ padding:0, overflow:"hidden" }}>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                    <thead>
                      <tr style={{ background:"#1e2a3a" }}>
                        <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px", whiteSpace:"nowrap" }}>Fecha</th>
                        <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Artículo</th>
                        <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Color</th>
                        <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Local</th>
                        <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Movimiento</th>
                        <th style={{ padding:"10px 12px", textAlign:"right", color:"#4ade80", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Entrada</th>
                        <th style={{ padding:"10px 12px", textAlign:"right", color:"#f87171", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Salida</th>
                        <th style={{ padding:"10px 12px", textAlign:"right", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Stock</th>
                        <th style={{ padding:"10px 12px", textAlign:"right", color:"#fbbf24", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:"0.5px" }}>Costo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {kardexFilt.length === 0 ? (
                        <tr><td colSpan={9} style={{ padding:"40px", textAlign:"center", color:C.t4 }}>
                          <div style={{ fontSize:32, marginBottom:8 }}>📒</div>
                          Sin movimientos para los filtros seleccionados
                        </td></tr>
                      ) : kardexFilt.map((k, idx) => (
                        <tr key={k.id} style={{ background: idx%2===0?C.white:C.bg, borderBottom:`1px solid ${C.border}` }}>
                          <td style={{ padding:"9px 12px", color:C.t3, fontSize:12, whiteSpace:"nowrap" }}>{k.f}</td>
                          <td style={{ padding:"9px 12px", fontWeight:700, color:C.t1 }}>{k.prod}</td>
                          <td style={{ padding:"9px 12px" }}>
                            <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                              <div style={{ width:8, height:8, borderRadius:"50%", background:HEX_COLOR[k.col]||"#888", flexShrink:0 }}/>
                              <span style={{ fontSize:12, color:C.t2 }}>{k.col||"—"}</span>
                            </div>
                          </td>
                          <td style={{ padding:"9px 12px" }}>
                            {(() => {
                              const ubi = UBICACIONES.find(u => (k.desc||"").includes(u));
                              const uColor = ubi==="Tienda Principal"?C.bl:ubi==="Tienda Pasaje"?C.pu:ubi?C.or:C.t4;
                              return <span style={{ fontSize:11, fontWeight:600, color:uColor }}>{ubi?ubi.replace("Taller/Almacén","Almacén"):"—"}</span>;
                            })()}
                          </td>
                          <td style={{ padding:"9px 12px" }}>
                            <span style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"3px 8px", borderRadius:99, fontSize:11, fontWeight:700,
                              background: k.tipo==="Venta"?C.acBg:k.tipo==="Compra"?C.blBg:k.tipo==="Fabricación"?C.grBg:C.bg,
                              color: k.tipo==="Venta"?C.ac:k.tipo==="Compra"?C.bl:k.tipo==="Fabricación"?C.grL:C.t3 }}>
                              {k.tipo==="Venta"?"🛒":k.tipo==="Compra"?"📦":k.tipo==="Fabricación"?"🔨":"🔄"} {k.tipo}
                            </span>
                            <div style={{ fontSize:10, color:C.t4, marginTop:2 }}>{k.desc}</div>
                          </td>
                          <td style={{ padding:"9px 12px", textAlign:"right", fontWeight:800, color:(k.ent||0)>0?C.grL:C.t4, fontSize:14 }}>
                            {(k.ent||0)>0 ? `+${k.ent}` : "—"}
                          </td>
                          <td style={{ padding:"9px 12px", textAlign:"right", fontWeight:800, color:(k.sal||0)>0?C.rdL:C.t4, fontSize:14 }}>
                            {(k.sal||0)>0 ? `-${k.sal}` : "—"}
                          </td>
                          <td style={{ padding:"9px 12px", textAlign:"right", fontWeight:800, color:C.t1, fontSize:14 }}>{k.saldo}</td>
                          <td style={{ padding:"9px 12px", textAlign:"right", color:C.or, fontWeight:600 }}>{fmt(k.costo)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Totales en el pie de tabla */}
                <div style={{ background:"#1e2a3a", padding:"10px 16px", display:"flex", gap:20, flexWrap:"wrap", justifyContent:"flex-end" }}>
                  <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                    <span style={{ fontSize:11, color:"rgba(255,255,255,0.6)", textTransform:"uppercase", letterSpacing:"0.4px" }}>Total ingresos:</span>
                    <span style={{ fontSize:15, fontWeight:800, color:"#4ade80" }}>+{totalEnt}</span>
                  </div>
                  <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                    <span style={{ fontSize:11, color:"rgba(255,255,255,0.6)", textTransform:"uppercase", letterSpacing:"0.4px" }}>Total salidas:</span>
                    <span style={{ fontSize:15, fontWeight:800, color:"#f87171" }}>-{totalSal}</span>
                  </div>
                  <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                    <span style={{ fontSize:11, color:"rgba(255,255,255,0.6)", textTransform:"uppercase", letterSpacing:"0.4px" }}>Neto:</span>
                    <span style={{ fontSize:15, fontWeight:800, color:"#fbbf24" }}>{totalEnt-totalSal > 0 ? "+" : ""}{totalEnt-totalSal}</span>
                  </div>
                  <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                    <span style={{ fontSize:11, color:"rgba(255,255,255,0.6)", textTransform:"uppercase", letterSpacing:"0.4px" }}>Registros:</span>
                    <span style={{ fontSize:15, fontWeight:800, color:"#fff" }}>{kardexFilt.length}</span>
                  </div>
                </div>
              </Card>
              </>
            );
          })()}

          {/* ═══════ PRODUCCIÓN ═══════ */}
          {modActual === "produccion" && (
            <>
              <PageTitle title="🔨 Producción" sub="Control de lotes de fabricación" action={<Btn variant="primary" onClick={() => setModal("lote")}>+ Nuevo Lote</Btn>}/>
              <Card style={{ padding:"10px 14px", marginBottom:12 }}>
                <div style={{ display:"flex", gap:10, alignItems:"flex-end", flexWrap:"wrap" }}>
                  <Inp label="Desde" type="date" value={prodFiltDesde} onChange={e => setProdFiltDesde(e.target.value)} style={{ width:145 }}/>
                  <Inp label="Hasta" type="date" value={prodFiltHasta} onChange={e => setProdFiltHasta(e.target.value)} style={{ width:145 }}/>
                  {(prodFiltDesde || prodFiltHasta) && <Btn onClick={() => { setProdFiltDesde(""); setProdFiltHasta(""); }}>✕ Limpiar</Btn>}
                </div>
              </Card>
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Fecha</TH><TH>Producto</TH><TH right>Cantidad</TH><TH right>Costo total</TH><TH>Estado</TH><TH>Acción</TH></tr></thead>
                  <tbody>
                    {lotes.filter(l =>
                      (!prodFiltDesde || l.f >= prodFiltDesde) &&
                      (!prodFiltHasta || l.f <= prodFiltHasta)
                    ).map(l => (
                      <tr key={l.id}>
                        <TD>{l.f}</TD>
                        <TD bold color={C.t1}>{l.prod}</TD>
                        <TD right>{l.qty} uds</TD>
                        <TD right bold>{fmt(l.total)}</TD>
                        <TD><Badge color={l.estado === "Completado" ? C.grL : C.or}>{l.estado}</Badge></TD>
                        <TD>
                          {l.estado === "En proceso" && (
                            <Btn variant="green" sm onClick={() => {
                              const loteCompleto = { ...l, estado:"Completado" };
                              setLotes(ls => { sbSave("lotes", loteCompleto); return ls.map(x => x.id === l.id ? loteCompleto : x); });
                              setProds(ps => ps.map(p => {
                                if (p.id !== l.pid) return p;
                                const np = { ...p, stk: p.stk + l.qty };
                                sbSave("productos", np);
                                return np;
                              }));
                              const nk = { id:"K"+uid(), f:HOY, pid:l.pid, prod:l.prod, col:"Varios", tipo:"Fabricación", desc:`Lote ${l.id}`, ent:l.qty, sal:0, saldo:(prods.find(p=>p.id===l.pid)?.stk||0)+l.qty, costo:l.costo };
                              setKardex(k => { sbSave("kardex", nk); return [nk, ...k]; });
                              showToast(`✓ +${l.qty} unidades al stock`);
                            }}>✓ Completar</Btn>
                          )}
                        </TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </>
          )}

          {/* ═══════ PEDIDOS ═══════ */}
          {modActual === "pedidos" && (() => {
            const EST_COLOR = { "Pendiente entrega":C.or, "En producción":C.bl, "Listo para entrega":C.grL, "Entregado":C.t4, "Cancelado":C.rd };
            const pedidosBase = isAdmin ? pedidos : pedidos.filter(p => p.vend === user.nombre);
            const hoy = new Date(HOY);

            const cerrarPedido = (pd) => {
              const pfx = "B001";
              const num = `${pfx}-${String(numB).padStart(5,"0")}`;
              const nv = {
                id:"V"+uid(), f:HOY, cli:pd.cli, vend:pd.vend,
                items:[{ id:pd.pid||"CUSTOM", n:pd.prod, col:pd.col, q:pd.qty, p:pd.precioAcordado,
                  c: prods.find(p=>p.id===pd.pid)?.c || 0 }],
                tot:pd.precioAcordado, mp:"Efectivo", comp:"Boleta", num,
                esPedido:true, pedidoId:pd.id
              };
              setVentas(v => { sbSave("ventas", nv); return [nv, ...v]; });
              if (pd.tipo === "separacion" && pd.pid) {
                setProds(ps => ps.map(p => {
                  if (p.id !== pd.pid) return p;
                  const np = { ...p, stk: Math.max(0, p.stk - pd.qty), stkRes: Math.max(0,(p.stkRes||0)-pd.qty) };
                  sbSave("productos", np);
                  return np;
                }));
                const nk = { id:"K"+uid(), f:HOY, pid:pd.pid, prod:pd.prod, col:pd.col,
                  tipo:"Venta", desc:`Venta pedido ${pd.id} - ${pd.cli}`, ent:0, sal:pd.qty,
                  saldo: (prods.find(p=>p.id===pd.pid)?.stk||0) - pd.qty, costo: prods.find(p=>p.id===pd.pid)?.c||0 };
                setKardex(k => { sbSave("kardex", nk); return [nk, ...k]; });
              }
              const pdCerrado = { ...pd, est:"Entregado", saldoPendiente:0, numVenta:num };
              setPedidos(ps => { sbSave("pedidos", pdCerrado); return ps.map(p => p.id === pd.id ? pdCerrado : p); });
              setNumB(n => { const nv2=n+1; sbSetConfig("numB",nv2); return nv2; });
              showToast(`✓ Entregado · Venta ${num} registrada · Suma a comisión de ${pd.vend}`);
            };

            const registrarAbono = (pd, monto, mp="Efectivo") => {
              const nuevoSaldo = Math.max(0, pd.saldoPendiente - monto);
              const pdActualizado = {
                ...pd,
                saldoPendiente: nuevoSaldo,
                abonos: [...(pd.abonos||[]), { f:HOY, monto, mp, id:"AB"+uid() }],
                est: nuevoSaldo === 0 ? "Listo para entrega" : pd.est
              };
              setPedidos(ps => { sbSave("pedidos", pdActualizado); return ps.map(p => p.id === pd.id ? pdActualizado : p); });
              const abonoNum = `ANTICIPO-${pd.id}-${Date.now().toString().slice(-4)}`;
              const anticipo = {
                id:"V"+uid(), f:HOY, cli:pd.cli, vend:pd.vend,
                items:[{ id:pd.pid||"ANTICIPO", n:`Anticipo pedido ${pd.id} − ${pd.prod}`, col:pd.col, q:1, p:monto, c:0 }],
                tot:monto, mp, comp:"Anticipo", num:abonoNum,
                esAnticipo:true
              };
              setVentas(vs => { sbSave("ventas", anticipo); return [anticipo, ...vs]; });
              showToast(`✓ Anticipo S/${monto} vía ${mp} · Saldo: S/${nuevoSaldo}`);
            };

            const nuevoPedido = (pd) => {
              const id = "PD"+String(pedidos.length+1).padStart(3,"0");
              if (pd.tipo === "separacion" && pd.pid) {
                setProds(ps => ps.map(p => {
                  if (p.id !== pd.pid) return p;
                  const np = { ...p, stkRes:(p.stkRes||0)+pd.qty };
                  sbSave("productos", np);
                  return np;
                }));
              }
              if (pd.adelanto > 0) {
                const adelantoV = { id:"V"+uid(), f:HOY, cli:pd.cli, vend:pd.vend,
                  items:[{ id:pd.pid||"ADELANTO", n:`Adelanto pedido ${id} - ${pd.prod}`, col:pd.col, q:1, p:pd.adelanto, c:0 }],
                  tot:pd.adelanto, mp:pd.mpAdelanto||"Efectivo", comp:"Anticipo", num:`ANTICIPO-${id}`,
                  esAnticipo:true };
                setVentas(vs => { sbSave("ventas", adelantoV); return [adelantoV, ...vs]; });
              }
              const nuevoPd = { ...pd, id, f:HOY, abonos:[], saldoPendiente: pd.precioAcordado - pd.adelanto };
              setPedidos(ps => { sbSave("pedidos", nuevoPd); return [nuevoPd, ...ps]; });
              showToast(`✓ Pedido ${id} registrado · Adelanto S/${pd.adelanto} en caja`);
              setPedidoForm(null);
            };

            // Días restantes
            const diasRestantes = (fEnt) => {
              const d = Math.round((new Date(fEnt) - hoy) / 86400000);
              return d;
            };

            const pendientes = pedidosBase.filter(p => p.est !== "Entregado" && p.est !== "Cancelado");
            const cerrados   = pedidosBase.filter(p => p.est === "Entregado" || p.est === "Cancelado");
            const totalSaldo = pendientes.reduce((a,p) => a + p.saldoPendiente, 0);

            return (<>
              <PageTitle title="📬 Pedidos" sub={`${pendientes.length} activos · S/${totalSaldo.toFixed(0)} saldo pendiente`}
                action={
                  <div style={{display:"flex",gap:8}}>
                    <Btn onClick={() => exportXLSX("pedidos", ["ID","Fecha","Tipo","Cliente","Cel","Vendedor","Producto","Color","Precio","Adelanto","Saldo","Estado","F.Entrega","Nota"],
                      pedidosBase.map(p=>[p.id,p.f,p.tipo==="separacion"?"Separación":"Pedido a medida",p.cli,p.cel||"",p.vend,p.prod,p.col,p.precioAcordado,p.adelanto,p.saldoPendiente,p.est,p.fEnt,p.nota||""]))}>⬇️ Excel</Btn>
                    <Btn variant="primary" onClick={() => setPedidoForm({ tipo:"separacion",cli:"",cel:"",vend:user.nombre,pid:"",prod:"",col:"",qty:1,precioAcordado:0,adelanto:0,fEnt:"",nota:"",diasAlerta:30 })}>+ Nuevo Pedido</Btn>
                  </div>
                }/>

              {/* KPIs */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
                <KPI icon="⏳" label="Pendiente entrega" value={pendientes.filter(p=>p.est==="Pendiente entrega").length} color={C.or}/>
                <KPI icon="🔨" label="En producción"     value={pendientes.filter(p=>p.est==="En producción").length} color={C.bl}/>
                <KPI icon="✅" label="Listo para entrega" value={pendientes.filter(p=>p.est==="Listo para entrega").length} color={C.grL}/>
                <KPI icon="💰" label="Saldo por cobrar"  value={`S/ ${totalSaldo.toFixed(0)}`} color={C.rd}/>
              </div>

              {/* Alertas vencidos */}
              {pendientes.filter(p => p.fEnt && diasRestantes(p.fEnt) < 0).length > 0 && (
                <div style={{ background:C.rdBg, border:`1px solid ${C.rd}44`, borderRadius:10, padding:"10px 16px", marginBottom:14, fontSize:13, color:C.rd, fontWeight:600 }}>
                  ⚠️ {pendientes.filter(p => p.fEnt && diasRestantes(p.fEnt) < 0).length} pedido(s) con fecha de entrega vencida
                </div>
              )}

              {/* Formulario nuevo pedido */}
              {pedidoForm !== null && (
                <PedidoFormModal form={pedidoForm} setForm={setPedidoForm} prods={prods} usuarios={usuarios} onSave={nuevoPedido} onCancel={() => setPedidoForm(null)}/>
              )}

              {/* Modal registro de abono */}
              {abonoModal && (
                <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1100, padding:16 }}>
                  <div style={{ background:"#fff", borderRadius:14, padding:24, width:340, maxWidth:"100%", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
                    <div style={{ fontSize:16, fontWeight:700, color:C.t1, marginBottom:4 }}>💵 Registrar Abono</div>
                    <div style={{ fontSize:13, color:C.t3, marginBottom:16, paddingBottom:14, borderBottom:`1px solid ${C.border}` }}>
                      {abonoModal.pd.cli} — {abonoModal.pd.prod}<br/>
                      <span style={{ color:C.rd, fontWeight:700 }}>Saldo pendiente: {fmt(abonoModal.pd.saldoPendiente)}</span>
                    </div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Monto del abono S/</label>
                    <input autoFocus inputMode="decimal" value={abonoModal.monto}
                      onChange={e => setAbonoModal(a => ({ ...a, monto: e.target.value.replace(/[^0-9.]/g,"") }))}
                      placeholder="S/ 0.00"
                      style={{ width:"100%", background:"#fff", border:`2px solid ${C.grL}`, borderRadius:8, padding:"12px 14px", fontSize:20, fontWeight:800, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.gr, textAlign:"center", marginBottom:8 }}/>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:6, marginBottom:12 }}>
                      {[50,100,200].map(v => (
                        <button key={v} onClick={() => setAbonoModal(a => ({ ...a, monto: String(Math.min(v, a.pd.saldoPendiente)) }))}
                          style={{ padding:"7px", borderRadius:7, border:`1px solid ${C.border}`, background:C.bg, fontSize:13, fontWeight:700, cursor:"pointer", color:C.t2 }}>
                          S/{v}
                        </button>
                      ))}
                    </div>
                    {/* Método de pago */}
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Método de pago</label>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:12 }}>
                      {["Efectivo","Yape/Plin","Trans. BCP","Trans. Interbank","POS Tarjeta"].map(mp => (
                        <button key={mp} onClick={() => setAbonoModal(a => ({ ...a, mp }))}
                          style={{ padding:"8px 4px", borderRadius:8, border:`2px solid ${(abonoModal.mp||"Efectivo")===mp?C.ac:C.border}`, background:(abonoModal.mp||"Efectivo")===mp?C.acBg:"transparent", color:(abonoModal.mp||"Efectivo")===mp?C.ac:C.t2, fontSize:11, fontWeight:700, cursor:"pointer", textAlign:"center" }}>
                          {{"Efectivo":"💵","Yape/Plin":"📱","Trans. BCP":"🏦","Trans. Interbank":"🏦","POS Tarjeta":"💳"}[mp]} {mp}
                        </button>
                      ))}
                    </div>
                    {parseFloat(abonoModal.monto) > abonoModal.pd.saldoPendiente && (
                      <div style={{ fontSize:11, color:C.rd, marginBottom:8, fontWeight:600 }}>⚠️ El monto supera el saldo pendiente</div>
                    )}
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={() => setAbonoModal(null)} style={{ flex:1 }}>Cancelar</Btn>
                      <Btn variant="green" onClick={() => {
                        const monto = parseFloat(abonoModal.monto);
                        if (!monto || monto <= 0) return;
                        if (monto > abonoModal.pd.saldoPendiente) return;
                        registrarAbono(abonoModal.pd, monto, abonoModal.mp||"Efectivo");
                        setAbonoModal(null);
                      }} style={{ flex:1 }} disabled={!parseFloat(abonoModal.monto) || parseFloat(abonoModal.monto) > abonoModal.pd.saldoPendiente}>
                        ✓ Confirmar
                      </Btn>
                    </div>
                  </div>
                </div>
              )}

              {/* Lista pedidos activos */}
              <div style={{ marginBottom:10, fontSize:12, fontWeight:700, color:C.t3, textTransform:"uppercase", letterSpacing:"0.5px" }}>Pedidos activos</div>
              <div style={{ display:"flex", flexDirection:"column", gap:10, marginBottom:20 }}>
                {pendientes.length === 0 && <Card style={{padding:24,textAlign:"center",color:C.t4}}>Sin pedidos activos</Card>}
                {pendientes.map(pd => {
                  const dias = pd.fEnt ? diasRestantes(pd.fEnt) : null;
                  const vencido = dias !== null && dias < 0;
                  const urgente = dias !== null && dias <= 3 && dias >= 0;
                  const pctPagado = pd.precioAcordado > 0 ? Math.round(((pd.precioAcordado - pd.saldoPendiente)/pd.precioAcordado)*100) : 0;
                  return (
                    <Card key={pd.id} style={{ padding:16, border: vencido ? `2px solid ${C.rd}` : urgente ? `2px solid ${C.or}` : undefined }}>
                      <div style={{ display:"flex", gap:12, alignItems:"flex-start" }}>
                        <div style={{ flexShrink:0 }}>
                          <div style={{ width:44, height:44, borderRadius:10, background: pd.tipo==="separacion"?C.blBg:C.acBg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22 }}>
                            {pd.tipo === "separacion" ? "📦" : "🔨"}
                          </div>
                          <div style={{ fontSize:9, textAlign:"center", color:C.t4, marginTop:3, fontWeight:600 }}>
                            {pd.tipo === "separacion" ? "STOCK" : "A MEDIDA"}
                          </div>
                        </div>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:3, flexWrap:"wrap" }}>
                            <span style={{ fontWeight:700, color:C.t1, fontSize:14 }}>{pd.cli}</span>
                            <Badge color={EST_COLOR[pd.est]||C.t3}>{pd.est}</Badge>
                            <span style={{ fontSize:11, color:C.t4 }}>{pd.id} · {pd.f}</span>
                          </div>
                          <div style={{ fontSize:13, color:C.t2, marginBottom:6 }}>
                            {pd.prod} — <span style={{ color:C.t3 }}>{pd.col}</span>
                            {pd.cel && <span style={{ color:C.t4, marginLeft:8 }}>📱 {pd.cel}</span>}
                          </div>
                          {/* Barra de pago */}
                          <div style={{ marginBottom:8 }}>
                            <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.t3, marginBottom:3 }}>
                              <span>Pagado: {fmt(pd.precioAcordado - pd.saldoPendiente)} / {fmt(pd.precioAcordado)}</span>
                              <span style={{ fontWeight:700, color: pd.saldoPendiente>0?C.rd:C.grL }}>Saldo: {fmt(pd.saldoPendiente)}</span>
                            </div>
                            <div style={{ height:6, borderRadius:99, background:C.bg3, overflow:"hidden" }}>
                              <div style={{ height:"100%", width:`${pctPagado}%`, background: pctPagado===100?C.grL:C.ac, borderRadius:99, transition:"width 0.3s" }}/>
                            </div>
                          </div>
                          {/* Fecha entrega */}
                          {pd.fEnt && (
                            <div style={{ fontSize:12, color: vencido?C.rd:urgente?C.or:C.t3, fontWeight: vencido||urgente?700:400 }}>
                              {vencido ? `⚠️ Venció hace ${Math.abs(dias)} días` : urgente ? `🔥 Entrega en ${dias} días` : `📅 Entrega: ${pd.fEnt} (${dias} días)`}
                            </div>
                          )}
                          {/* Vendedor y nota */}
                          <div style={{ fontSize:11, color:C.t4, marginTop:3 }}>
                            Vendedor: {pd.vend} {pd.nota && `· ${pd.nota}`}
                          </div>
                          {/* Historial abonos */}
                          {pd.abonos?.length > 0 && (
                            <div style={{ marginTop:6, display:"flex", gap:5, flexWrap:"wrap" }}>
                              {pd.abonos.map((a,i) => (
                                <span key={i} style={{ fontSize:10, background:C.grBg, color:C.grL, padding:"2px 7px", borderRadius:5, fontWeight:600 }}>+{fmt(a.monto)} ({a.f})</span>
                              ))}
                            </div>
                          )}
                        </div>
                        {/* Acciones */}
                        <div style={{ display:"flex", flexDirection:"column", gap:6, flexShrink:0 }}>
                          {/* Selector de estado — bloquea Entregado si hay saldo */}
                          <select value={pd.est} onChange={e => {
                            if (e.target.value === "Entregado" && pd.saldoPendiente > 0) {
                              showToast(`⚠️ Saldo pendiente: ${fmt(pd.saldoPendiente)} · Registra el abono primero`, "err");
                              return;
                            }
                            setPedidos(ps => ps.map(p => {
                              if (p.id !== pd.id) return p;
                              const np = {...p, est:e.target.value};
                              sbSave("pedidos", np);
                              return np;
                            }));
                          }}
                            style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:7, padding:"5px 8px", fontSize:11, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                            {Object.keys(EST_COLOR).filter(e => e !== "Entregado").map(e => <option key={e}>{e}</option>)}
                            <option value="Entregado" disabled={pd.saldoPendiente > 0}>
                              {pd.saldoPendiente > 0 ? `🔒 Entregado (falta ${fmt(pd.saldoPendiente)})` : "Entregado"}
                            </option>
                          </select>
                          {/* Registrar abono */}
                          {pd.saldoPendiente > 0 && (
                            <button onClick={() => setAbonoModal({ pd, monto:"" })}
                              style={{ padding:"5px 10px", borderRadius:7, border:`1px solid ${C.grL}`, background:C.grBg, color:C.grL, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                              💵 Registrar abono
                            </button>
                          )}
                          {/* Cerrar pedido */}
                          {pd.saldoPendiente === 0 && pd.est !== "Entregado" && (
                            <Btn sm variant="green" onClick={() => cerrarPedido(pd)}>✅ Cerrar pedido</Btn>
                          )}
                          {pd.cel && (
                            <a href={`https://wa.me/51${pd.cel}?text=${encodeURIComponent(`Hola ${pd.cli}, le recordamos que su pedido "${pd.prod}" tiene un saldo pendiente de S/${pd.saldoPendiente}. ${pd.fEnt?`Fecha de entrega estimada: ${pd.fEnt}.`:""}`)}`}
                              target="_blank" rel="noopener noreferrer"
                              style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:4, background:"#25D366", color:"#fff", borderRadius:7, padding:"5px 10px", fontSize:11, fontWeight:700, textDecoration:"none" }}>
                              💬 WA
                            </a>
                          )}
                        </div>
                      </div>
                    </Card>
                  );
                })}
              </div>

              {/* Pedidos cerrados */}
              {cerrados.length > 0 && (<>
                <div style={{ marginBottom:10, fontSize:12, fontWeight:700, color:C.t3, textTransform:"uppercase", letterSpacing:"0.5px" }}>Historial ({cerrados.length})</div>
                <Card>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead><tr><TH>ID</TH><TH>Fecha</TH><TH>Cliente</TH><TH>Producto</TH><TH right>Total</TH><TH>Estado</TH><TH>Vendedor</TH></tr></thead>
                    <tbody>
                      {cerrados.map(pd => (
                        <tr key={pd.id}>
                          <TD sm color={C.t4}>{pd.id}</TD>
                          <TD>{pd.f}</TD>
                          <TD bold color={C.t1}>{pd.cli}</TD>
                          <TD sm>{pd.prod} · {pd.col}</TD>
                          <TD right bold color={C.grL}>{fmt(pd.precioAcordado)}</TD>
                          <TD><Badge color={EST_COLOR[pd.est]||C.t3}>{pd.est}</Badge></TD>
                          <TD sm color={C.t3}>{pd.vend}</TD>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </Card>
              </>)}
            </>);
          })()}

          {/* ═══════ TRASLADOS ═══════ */}
          {modActual === "traslados" && (() => {
            const realizarTraslado = (t) => {
              const prod = prods.find(p => p.id === t.pid);
              if (!prod) return;
              const stkOrig = getLocs(prod)[locKey(t.ubiOrig, t.col)]||0;
              if (stkOrig < t.qty) {
                showToast(`⚠️ Solo hay ${stkOrig} uds de ${t.col} en ${t.ubiOrig}`, "err"); return;
              }
              setProds(ps => ps.map(p => {
                if (p.id !== t.pid) return p;
                const np = trasladarStk(p, t.qty, t.ubiOrig, t.ubiDest, t.col);
                sbSave("productos", np);
                return np;
              }));
              const nk = { id:"K"+uid(), f:t.f||HOY, pid:t.pid, prod:prod.n, col:t.col, tipo:"Ajuste",
                desc:`Traslado ${t.qty} uds ${t.col}: ${t.ubiOrig} → ${t.ubiDest}`, ent:0, sal:0, saldo:getTotalStk(prod), costo:prod.c };
              setKardex(k => { sbSave("kardex", nk); return [nk, ...k]; });
              const nuevoTraslado = { ...t, id:"TR"+uid(), f:t.f||HOY, prod:prod.n };
              setTraslados(ts => { sbSave("traslados", nuevoTraslado); return [nuevoTraslado, ...ts]; });
              showToast(`✓ ${t.qty} × ${prod.n} (${t.col}) → ${t.ubiDest}`);
              setTrasladoForm(null);
            };

            return (<>
              <PageTitle title="🚚 Traslados" sub="Mover productos entre ubicaciones"
                action={<Btn variant="primary" onClick={() => setTrasladoForm({ f:HOY, pid:"", col:"", qty:1, ubiOrig:UBICACIONES[0], ubiDest:UBICACIONES[1], nota:"" })}>+ Nuevo Traslado</Btn>}/>

              {/* KPIs */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:16 }}>
                {UBICACIONES.map(ubi => (
                  <KPI key={ubi} icon={ubi==="Tienda Principal"?"🏪":ubi==="Tienda Pasaje"?"🛍️":"🏭"}
                    label={ubi}
                    value={`${prods.filter(p=>(getLocs(p)[ubi]||0)>0).length} productos`}
                    color={ubi==="Tienda Principal"?C.bl:ubi==="Tienda Pasaje"?C.pu:C.or}
                    sub={`Stock: ${prods.reduce((a,p)=>a+(getLocs(p)[ubi]||0),0)} uds`}/>
                ))}
              </div>

              {/* Formulario traslado — nuevo diseño */}
              {trasladoForm !== null && (
                <Card style={{ padding:20, marginBottom:16, border:`2px solid ${C.ac}` }}>
                  <div style={{ fontSize:15, fontWeight:700, color:C.t1, marginBottom:20 }}>🚚 Nuevo Traslado</div>
                  {/* Buscador + Producto */}
                  <div style={{ marginBottom:16 }}>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>¿Qué producto trasladar?</label>
                    <input value={trasladoForm._busq||""} onChange={e => setTrasladoForm(f=>({...f,_busq:e.target.value}))}
                      placeholder="🔍 Escribe el nombre del producto..."
                      style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1, marginBottom:6 }}/>
                    <div style={{ maxHeight:180, overflowY:"auto", border:`1px solid ${C.border}`, borderRadius:8, background:C.white }}>
                      {prods.filter(p => !trasladoForm._busq || p.n.toLowerCase().includes((trasladoForm._busq||"").toLowerCase())).map(p => {
                        const locs = getLocs(p);
                        return (
                          <button key={p.id} onClick={() => setTrasladoForm(f=>({...f, pid:p.id, col:p.cols[0]||"",
                            ubiOrig: UBICACIONES.find(u=>(locs[u]||0)>0)||UBICACIONES[0],
                            ubiDest: UBICACIONES.find(u=>(locs[u]||0)===0 || u!==UBICACIONES.find(u2=>(locs[u2]||0)>0))||UBICACIONES[1],
                            _busq:p.n }))}
                            style={{ width:"100%", textAlign:"left", padding:"9px 14px", border:"none", borderBottom:`1px solid ${C.border}`, background:trasladoForm.pid===p.id?C.acBg:"transparent", cursor:"pointer", fontFamily:"inherit" }}>
                            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                              <span style={{ fontSize:13, fontWeight:trasladoForm.pid===p.id?700:400, color:trasladoForm.pid===p.id?C.ac:C.t1 }}>{p.ico} {p.n}</span>
                              <div style={{ fontSize:10, color:C.t3, textAlign:"right", flexShrink:0, marginLeft:8 }}>
                                {UBICACIONES.map(u => (locs[u]||0) > 0 && (
                                  <div key={u} style={{ color:u==="Tienda Principal"?C.bl:u==="Tienda Pasaje"?C.pu:C.or, fontWeight:600 }}>
                                    {u.replace("Taller/Almacén","Almacén")}: {locs[u]}
                                  </div>
                                ))}
                              </div>
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  {trasladoForm.pid && (() => {
                    const prod = prods.find(p=>p.id===trasladoForm.pid);
                    if (!prod) return null;
                    const ubisDisp = ubisConStock(prod); // ubicaciones con stock (cualquier color)
                    const colsDisp = colsConStockEnUbi(prod, trasladoForm.ubiOrig); // colores con stock en el origen
                    const stkDisp  = getLocs(prod)[locKey(trasladoForm.ubiOrig, trasladoForm.col)]||0;
                    return (<>
                      {/* Color — solo muestra colores con stock EN EL ORIGEN */}
                      <div style={{ marginBottom:16 }}>
                        <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>Color a trasladar</label>
                        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                          {(prod.cols||[]).map(col => {
                            const stkCol = getLocs(prod)[locKey(trasladoForm.ubiOrig, col)]||0;
                            const activo = stkCol > 0;
                            return (
                              <button key={col} onClick={() => activo && setTrasladoForm(f=>({...f,col}))}
                                title={activo ? `${stkCol} uds en ${trasladoForm.ubiOrig}` : `Sin stock de ${col} en ${trasladoForm.ubiOrig}`}
                                style={{ padding:"8px 16px", borderRadius:99, border:`2px solid ${trasladoForm.col===col?C.ac:activo?C.border:"#eee"}`,
                                  background:trasladoForm.col===col?C.acBg:activo?"transparent":"#f9f9f9",
                                  color:trasladoForm.col===col?C.ac:activo?C.t2:C.t4,
                                  fontSize:13, fontWeight:700, cursor:activo?"pointer":"not-allowed",
                                  opacity:activo?1:0.4, position:"relative" }}>
                                <span style={{ display:"inline-block", width:10, height:10, borderRadius:"50%", background:HEX_COLOR[col]||"#888", marginRight:6, verticalAlign:"middle" }}/>{col}
                                {activo && <span style={{ marginLeft:6, fontSize:10, color:activo&&trasladoForm.col===col?C.ac:C.t3 }}>({stkCol})</span>}
                                {!activo && <span style={{ marginLeft:4, fontSize:9 }}>✕</span>}
                              </button>
                            );
                          })}
                        </div>
                        {colsDisp.length === 0 && <div style={{ fontSize:12, color:C.rd, marginTop:6 }}>⚠️ Sin stock disponible en {trasladoForm.ubiOrig}</div>}
                      </div>
                      {/* Origen → Destino — origen solo muestra ubicaciones con stock del color */}
                      <div style={{ marginBottom:16 }}>
                        <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>¿De dónde a dónde?</label>
                        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                          <div style={{ flex:1 }}>
                            <div style={{ fontSize:10, color:C.t4, fontWeight:600, textTransform:"uppercase", marginBottom:4 }}>📍 Origen</div>
                            <select value={trasladoForm.ubiOrig} onChange={e => {
                              const nuevoOrigen = e.target.value;
                              const colsEnNuevo = colsConStockEnUbi(prod, nuevoOrigen);
                              const nuevoColor = colsEnNuevo.includes(trasladoForm.col) ? trasladoForm.col : (colsEnNuevo[0]||"");
                              setTrasladoForm(f=>({...f, ubiOrig:nuevoOrigen, col:nuevoColor,
                                ubiDest: UBICACIONES.find(u=>u!==nuevoOrigen)||UBICACIONES[1]}));
                            }}
                              style={{ width:"100%", background:C.rdBg, border:`2px solid ${C.rdL}44`, borderRadius:10, padding:"10px 12px", fontSize:13, fontWeight:700, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                              {UBICACIONES.map(u => {
                                const stkUbi = getStkEnUbi(prod,u);
                                return <option key={u} value={u} disabled={stkUbi===0}>{u} {stkUbi>0?`(${stkUbi} uds)`:""} {stkUbi===0?"— sin stock":""}</option>;
                              })}
                            </select>
                          </div>
                          <div style={{ fontSize:28, color:C.ac, fontWeight:700, paddingTop:18 }}>→</div>
                          <div style={{ flex:1 }}>
                            <div style={{ fontSize:10, color:C.t4, fontWeight:600, textTransform:"uppercase", marginBottom:4 }}>📍 Destino</div>
                            <select value={trasladoForm.ubiDest} onChange={e => setTrasladoForm(f=>({...f,ubiDest:e.target.value}))}
                              style={{ width:"100%", background:C.grBg, border:`2px solid ${C.grL}44`, borderRadius:10, padding:"10px 12px", fontSize:13, fontWeight:700, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                              {UBICACIONES.filter(u=>u!==trasladoForm.ubiOrig).map(u => <option key={u}>{u}</option>)}
                            </select>
                          </div>
                        </div>
                      </div>
                      {/* Cantidad con botones +/- */}
                      <div style={{ marginBottom:16 }}>
                        <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>¿Cuántas unidades? <span style={{color:C.t4,fontWeight:400,textTransform:"none"}}>(hay {stkDisp} de {trasladoForm.col||"este color"} en {trasladoForm.ubiOrig})</span></label>
                        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                          <button onClick={() => setTrasladoForm(f=>({...f,qty:Math.max(1,f.qty-1)}))}
                            style={{ width:44, height:44, borderRadius:10, border:`1px solid ${C.border}`, background:C.bg, fontSize:22, fontWeight:700, cursor:"pointer", color:C.t1, flexShrink:0 }}>−</button>
                          <input inputMode="numeric"
                            value={trasladoForm.qty === 0 ? "" : trasladoForm.qty}
                            onChange={e => { const raw = e.target.value.replace(/[^0-9]/g,""); setTrasladoForm(f=>({...f,qty:raw===""?0:Math.min(parseInt(raw),stkDisp)})); }}
                            onBlur={() => { if (!trasladoForm.qty||trasladoForm.qty<1) setTrasladoForm(f=>({...f,qty:1})); }}
                            style={{ flex:1, textAlign:"center", fontSize:28, fontWeight:800, background:C.white, border:`2px solid ${stkDisp>0?C.ac:C.rd}`, borderRadius:10, padding:"8px", outline:"none", color:C.t1, fontFamily:"inherit" }}/>
                          <button onClick={() => setTrasladoForm(f=>({...f,qty:Math.min(stkDisp,f.qty+1)}))}
                            style={{ width:44, height:44, borderRadius:10, border:`1px solid ${C.border}`, background:C.bg, fontSize:22, fontWeight:700, cursor:"pointer", color:C.t1, flexShrink:0 }}>+</button>
                          <span style={{ fontSize:12, color:C.t3, flexShrink:0 }}>de {stkDisp} uds</span>
                        </div>
                        <div style={{ display:"flex", gap:6, marginTop:8 }}>
                          {[1,2,3,5,10].filter(v=>v<=stkDisp).map(v => (
                            <button key={v} onClick={() => setTrasladoForm(f=>({...f,qty:v}))}
                              style={{ padding:"5px 12px", borderRadius:99, border:`1px solid ${trasladoForm.qty===v?C.ac:C.border}`, background:trasladoForm.qty===v?C.acBg:"transparent", fontSize:12, fontWeight:700, cursor:"pointer", color:trasladoForm.qty===v?C.ac:C.t3 }}>
                              {v}
                            </button>
                          ))}
                          {stkDisp > 1 && <button onClick={() => setTrasladoForm(f=>({...f,qty:stkDisp}))}
                            style={{ padding:"5px 12px", borderRadius:99, border:`1px solid ${trasladoForm.qty===stkDisp?C.ac:C.border}`, background:trasladoForm.qty===stkDisp?C.acBg:"transparent", fontSize:12, fontWeight:700, cursor:"pointer", color:trasladoForm.qty===stkDisp?C.ac:C.t3 }}>
                            Todo ({stkDisp})
                          </button>}
                        </div>
                      </div>
                      {/* Resumen */}
                      <div style={{ background:C.acBg, border:`1px solid ${C.ac}33`, borderRadius:10, padding:"12px 16px", marginBottom:14, fontSize:13 }}>
                        <div style={{ fontWeight:700, color:C.ac, marginBottom:4 }}>{trasladoForm.qty} × {prod?.ico} {prod?.n} — {trasladoForm.col}</div>
                        <div><span style={{color:C.rd}}>📍 {trasladoForm.ubiOrig}</span><span style={{margin:"0 8px",color:C.ac}}>→</span><span style={{color:C.grL}}>📍 {trasladoForm.ubiDest}</span></div>
                      </div>
                      <Inp label="Nota (opcional)" value={trasladoForm.nota} onChange={e => setTrasladoForm(f=>({...f,nota:e.target.value}))} placeholder="Ej: Para exhibición en vitrina"/>
                    </>);
                  })()}
                  <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:16, paddingTop:14, borderTop:`1px solid ${C.border}` }}>
                    <Btn onClick={() => setTrasladoForm(null)}>Cancelar</Btn>
                    <Btn variant="primary" onClick={() => {
                      if (!trasladoForm.pid) { showToast("Selecciona un producto","err"); return; }
                      if (trasladoForm.ubiOrig===trasladoForm.ubiDest) { showToast("Origen y destino deben ser distintos","err"); return; }
                      if (!trasladoForm.qty||trasladoForm.qty<1) { showToast("Ingresa una cantidad válida","err"); return; }
                      realizarTraslado(trasladoForm);
                    }}>🚚 Confirmar Traslado</Btn>
                  </div>
                </Card>
              )}


              {/* Historial traslados */}
              {traslados.length > 0 ? (
                <Card>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead><tr><TH>Fecha</TH><TH>Producto</TH><TH right>Cant.</TH><TH>Origen</TH><TH></TH><TH>Destino</TH><TH>Nota</TH></tr></thead>
                    <tbody>
                      {traslados.map(t => (
                        <tr key={t.id}>
                          <TD>{t.f}</TD>
                          <TD bold color={C.t1}>{t.prod} {t.col && <span style={{color:C.t3,fontWeight:400}}>({t.col})</span>}</TD>
                          <TD right bold color={C.ac}>{t.qty} uds</TD>
                          <TD><Badge color={C.t4}>{t.ubiOrig}</Badge></TD>
                          <TD>→</TD>
                          <TD><Badge color={C.grL}>{t.ubiDest}</Badge></TD>
                          <TD sm color={C.t3}>{t.nota||"—"}</TD>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </Card>
              ) : (
                <Card style={{ padding:40, textAlign:"center", color:C.t4 }}>
                  <div style={{ fontSize:32, marginBottom:8 }}>🚚</div>
                  <div>Sin traslados registrados aún</div>
                  <div style={{ fontSize:12, marginTop:4 }}>Usa "+ Nuevo Traslado" para mover productos entre ubicaciones</div>
                </Card>
              )}
            </>);
          })()}

          {/* ═══════ GASTOS ═══════ */}
          {modActual === "gastos" && (
            <>
              <PageTitle title="💸 Gastos" sub="Control de egresos del negocio" action={
                <div style={{display:"flex",gap:8}}>
                  <Btn onClick={() => exportXLSX("gastos", ["Fecha","Categoría","Descripción","Proveedor","Monto"],
                    gastos.filter(g => (!gastoFiltDesde||g.f>=gastoFiltDesde) && (!gastoFiltHasta||g.f<=gastoFiltHasta) && (gastoFiltCat==="Todas"||g.cat===gastoFiltCat))
                    .map(g=>[g.f,g.cat,g.desc,g.prov||"",g.monto]))}>⬇️ Excel</Btn>
                  <Btn variant="primary" onClick={() => setModal("gasto")}>+ Registrar</Btn>
                </div>
              }/>
              <Card style={{ padding:"10px 14px", marginBottom:12 }}>
                <div style={{ display:"flex", gap:10, alignItems:"flex-end", flexWrap:"wrap" }}>
                  <Inp label="Desde" type="date" value={gastoFiltDesde} onChange={e => setGastoFiltDesde(e.target.value)} style={{ width:145 }}/>
                  <Inp label="Hasta" type="date" value={gastoFiltHasta} onChange={e => setGastoFiltHasta(e.target.value)} style={{ width:145 }}/>
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Categoría</label>
                    <select value={gastoFiltCat} onChange={e => setGastoFiltCat(e.target.value)}
                      style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                      {["Todas","Alquiler","Servicios taller","Sueldo base","Mantenimiento","Materiales","Armador por obra","Proveedor","Comisión","Otro"].map(c => <option key={c}>{c}</option>)}
                    </select>
                  </div>
                  {(gastoFiltDesde || gastoFiltHasta || gastoFiltCat !== "Todas") && (
                    <Btn onClick={() => { setGastoFiltDesde(""); setGastoFiltHasta(""); setGastoFiltCat("Todas"); }}>✕ Limpiar</Btn>
                  )}
                </div>
              </Card>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
                {[
                  {cat:"Alquiler",      icon:"🏠"},
                  {cat:"Sueldo base",   icon:"👤"},
                  {cat:"Materiales",    icon:"🪵"},
                  {cat:"Comisión",      icon:"🏆"},
                ].map(({ cat, icon }) => {
                  const tot = gastos.filter(g => g.cat === cat && (!gastoFiltDesde||g.f>=gastoFiltDesde) && (!gastoFiltHasta||g.f<=gastoFiltHasta)).reduce((a,g) => a + g.monto, 0);
                  return (<KPI key={cat} icon={icon} label={cat} value={fmt(tot)} color={C.or}/>);
                })}
              </div>
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Fecha</TH><TH>Categoría</TH><TH>Descripción</TH><TH>Proveedor</TH><TH right>Monto</TH></tr></thead>
                  <tbody>
                    {[...gastos]
                      .filter(g => (!gastoFiltDesde||g.f>=gastoFiltDesde) && (!gastoFiltHasta||g.f<=gastoFiltHasta) && (gastoFiltCat==="Todas"||g.cat===gastoFiltCat))
                      .sort((a,b) => b.f.localeCompare(a.f)).map(g => (
                      <tr key={g.id}>
                        <TD>{g.f}</TD>
                        <TD><Badge color={C.or}>{g.cat}</Badge></TD>
                        <TD bold color={C.t1}>{g.desc}</TD>
                        <TD>{g.prov || "—"}</TD>
                        <TD right bold color={C.rd}>{fmt(g.monto)}</TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </>
          )}

          {/* ═══════ FINANZAS ═══════ */}
          {modActual === "finanzas" && (() => {
            // mesesDisp calculado inline (no hooks dentro de IIFE)
            const mesesSet = new Set([...ventas.map(v => v.f.slice(0,7)), ...gastos.map(g => g.f.slice(0,7))]);
            const mesesDisp = [...mesesSet].sort().reverse();
            const ventasMes  = ventas.filter(v => v.f.startsWith(finMes));
            const gastosMes  = gastos.filter(g => g.f.startsWith(finMes));
            const fIng  = ventasMes.reduce((a,v) => a+v.tot, 0);
            const fCmv  = ventasMes.reduce((a,v) => a + v.items.reduce((b,i) => b + i.c*i.q, 0), 0);
            const fGm   = gastosMes.reduce((a,g) => a+g.monto, 0);
            const fUt   = fIng - fCmv - fGm;
            const fMg   = fIng > 0 ? Math.round((fUt/fIng)*100) : 0;
            const [yy, mm] = finMes.split("-");
            const nomMes = `${MESES[parseInt(mm)-1]} ${yy}`;
            // Comparativo vs mes anterior
            const [prevYY, prevMM] = (() => { const d = new Date(parseInt(yy), parseInt(mm)-2, 1); return [d.getFullYear(), String(d.getMonth()+1).padStart(2,"0")]; })();
            const prevKey = `${prevYY}-${prevMM}`;
            const prevIng = ventas.filter(v=>v.f.startsWith(prevKey)).reduce((a,v)=>a+v.tot,0);
            const varIng = prevIng > 0 ? Math.round(((fIng-prevIng)/prevIng)*100) : null;
            return (<>
              <PageTitle title="📊 Finanzas" sub={`Estado financiero · ${nomMes}`}
                action={
                  <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                    <select value={finMes} onChange={e => setFinMes(e.target.value)}
                      style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"7px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, fontWeight:600 }}>
                      {mesesDisp.map(m => { const [y,mo]=m.split("-"); return <option key={m} value={m}>{MESES[parseInt(mo)-1]} {y}</option>; })}
                    </select>
                    <Btn onClick={() => exportXLSX("finanzas_"+finMes, ["Concepto","Monto"], [["Ventas brutas",fIng],["Costo mercadería",-fCmv],["Utilidad bruta",fIng-fCmv],["Gastos operativos",-fGm],["Utilidad neta",fUt],["Margen %",fMg]])}>⬇️ Excel</Btn>
                  </div>
                }
              />
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
                <KPI icon="💰" label="Ingresos" value={fmtK(fIng)} color={C.grL} trend={varIng} sub={`${ventasMes.length} ventas`}/>
                <KPI icon="📉" label="Gastos" value={fmtK(fGm)} color={C.rdL} sub={`${gastosMes.length} registros`}/>
                <KPI icon="✨" label="Utilidad neta" value={fmtK(fUt)} color={fUt>=0?C.ac:C.rd} sub={`Margen ${fMg}%`}/>
                <KPI icon="🏷️" label="Costo mercadería" value={fmtK(fCmv)} color={C.or} sub="CMV"/>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
                <Card style={{ padding:20 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:16, textTransform:"uppercase", letterSpacing:"0.5px" }}>Estado de Resultados · {nomMes}</div>
                  {[
                    { label:"Ventas brutas",        val: fIng,        color:C.grL,  bold:true },
                    { label:"(-) Costo mercadería", val:-fCmv,        color:C.rd },
                    { label:"= Utilidad bruta",     val: fIng-fCmv,  color:C.ac,   bold:true, line:true },
                    { label:"(-) Gastos operativos",val:-fGm,         color:C.rd },
                    { label:"= Utilidad neta",      val: fUt,         color: fUt >= 0 ? C.grL : C.rd, bold:true, line:true },
                  ].map((r,i) => (
                    <div key={i} style={{ ...(r.line ? { borderTop:`1px solid ${C.border}`, paddingTop:8, marginTop:4 } : {}), display:"flex", justifyContent:"space-between", marginBottom:8, fontSize:13 }}>
                      <span style={{ color: r.bold ? C.t1 : C.t2, fontWeight: r.bold ? 700 : 400 }}>{r.label}</span>
                      <span style={{ color:r.color, fontWeight:700 }}>{fmt(Math.abs(r.val))}</span>
                    </div>
                  ))}
                  <div style={{ borderTop:`2px solid ${C.borderD}`, paddingTop:12, marginTop:8, display:"flex", justifyContent:"space-between" }}>
                    <span style={{ fontWeight:700, color:C.t1 }}>Margen neto</span>
                    <span style={{ fontWeight:800, color: fMg >= 20 ? C.grL : C.or, fontSize:22 }}>{fMg}%</span>
                  </div>
                  {prevIng > 0 && (
                    <div style={{ marginTop:12, padding:"10px 12px", background:C.bg, borderRadius:8, fontSize:12, color:C.t3 }}>
                      vs {MESES[parseInt(prevMM)-1]}: <strong style={{color:C.t2}}>{fmt(prevIng)}</strong>
                      <span style={{ marginLeft:8, fontWeight:700, color: varIng >= 0 ? C.grL : C.rdL }}>{varIng >= 0 ? "↑" : "↓"}{Math.abs(varIng)}%</span>
                    </div>
                  )}
                  {/* Estimado SUNAT */}
                  {(() => {
                    // SUNAT: solo boletas y facturas tienen valor fiscal, NO notas de venta ni anticipos
                    const ventasFiscales = ventas.filter(v => !v.esAnticipo && (v.comp==="Boleta"||v.comp==="Factura") && v.f.startsWith(finMes));
                    const ingFiscal = ventasFiscales.reduce((a,v)=>a+v.tot,0);
                    const igvEstim  = +(ingFiscal * 0.18 / 1.18).toFixed(2);
                    const rerEstim  = +(ingFiscal * 0.015).toFixed(2);
                    const irEstim   = fUt > 0 ? +(fUt * 0.10).toFixed(2) : 0;
                    return (
                      <div style={{ marginTop:14, padding:"14px 16px", background:"#fffbeb", border:"1px solid #f6e05e44", borderRadius:10 }}>
                        <div style={{ fontSize:12, fontWeight:700, color:"#744210", marginBottom:10, display:"flex", justifyContent:"space-between" }}>
                          <span>🏛️ Estimado SUNAT · {nomMes}</span>
                          <span style={{ fontSize:10, fontWeight:400, color:"#a0aec0" }}>Base: {fmt(ingFiscal)} en Boletas+Facturas</span>
                        </div>
                        {[
                          { label:"IGV (18% incluido en ventas)", val:igvEstim, tip:"Si emites comprobantes con IGV" },
                          { label:"Impuesto RER (1.5% ingresos)", val:rerEstim, tip:"Régimen Especial Renta" },
                          { label:"IR estimado (10% utilidad)", val:irEstim, tip:"Régimen MYPE Tributario" },
                        ].map(r => (
                          <div key={r.label} style={{ display:"flex", justifyContent:"space-between", marginBottom:6, fontSize:12 }}>
                            <div>
                              <div style={{ color:"#744210", fontWeight:600 }}>{r.label}</div>
                              <div style={{ fontSize:10, color:"#a0aec0" }}>{r.tip}</div>
                            </div>
                            <span style={{ fontWeight:800, color:"#c05621", fontSize:14, flexShrink:0, marginLeft:12 }}>{fmt(r.val)}</span>
                          </div>
                        ))}
                        <div style={{ borderTop:"1px solid #f6e05e", paddingTop:8, marginTop:4, display:"flex", justifyContent:"space-between", fontSize:13, fontWeight:800 }}>
                          <span style={{ color:"#744210" }}>Total estimado a apartar</span>
                          <span style={{ color:"#c05621" }}>{fmt(Math.max(igvEstim, rerEstim) + irEstim)}</span>
                        </div>
                        <div style={{ fontSize:10, color:"#a0aec0", marginTop:6, fontStyle:"italic" }}>
                          ⚠️ Estos valores son referenciales. Tu contador determina el régimen correcto y los montos exactos a declarar.
                        </div>
                      </div>
                    );
                  })()}
                </Card>

                {/* PUNTO DE EQUILIBRIO */}
                <Card style={{ padding:20 }}>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>⚖️ Punto de Equilibrio · {nomMes}</div>
                  <div style={{ fontSize:11, color:C.t4, marginBottom:14 }}>¿Cuánto necesitas vender para no perder?</div>
                  {(() => {
                    const gastosFijos = gastos.filter(g => (g.esFijo || ["Alquiler","Servicios taller","Sueldo base","Mantenimiento"].includes(g.cat)) && g.f.startsWith(finMes)).reduce((a,g)=>a+g.monto,0);
                    const gastoVar    = gastos.filter(g => (!g.esFijo && ["Materiales","Armador por obra","Proveedor","Comisión"].includes(g.cat)) && g.f.startsWith(finMes)).reduce((a,g)=>a+g.monto,0);
                    const margenProm  = fIng > 0 ? (fUt + gastosFijos) / fIng : 0.35;
                    const peVentas    = margenProm > 0 ? gastosFijos / margenProm : 0;
                    const peDiario    = peVentas / 26; // 26 días hábiles
                    const avance      = fIng > 0 ? Math.min(100, Math.round((fIng / peVentas) * 100)) : 0;
                    return (<>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
                        <div style={{ background:C.bg, borderRadius:8, padding:"10px 12px", textAlign:"center" }}>
                          <div style={{ fontSize:18, fontWeight:800, color:C.ac }}>{fmt(peVentas)}</div>
                          <div style={{ fontSize:10, color:C.t3, marginTop:2 }}>Ventas mínimas/mes</div>
                        </div>
                        <div style={{ background:C.bg, borderRadius:8, padding:"10px 12px", textAlign:"center" }}>
                          <div style={{ fontSize:18, fontWeight:800, color:C.bl }}>{fmt(peDiario)}</div>
                          <div style={{ fontSize:10, color:C.t3, marginTop:2 }}>Ventas mínimas/día</div>
                        </div>
                      </div>
                      <div style={{ marginBottom:6, display:"flex", justifyContent:"space-between", fontSize:12 }}>
                        <span style={{ color:C.t3 }}>Progreso este mes</span>
                        <span style={{ fontWeight:700, color:avance>=100?C.grL:avance>=70?C.or:C.rd }}>{avance}%</span>
                      </div>
                      <div style={{ height:10, borderRadius:99, background:C.bg3, overflow:"hidden", marginBottom:10 }}>
                        <div style={{ height:"100%", width:`${avance}%`, background:avance>=100?C.grL:avance>=70?C.or:C.rd, borderRadius:99, transition:"width 0.3s" }}/>
                      </div>
                      <div style={{ fontSize:11, color:C.t3, marginBottom:6 }}>Gastos fijos considerados: <strong>{fmt(gastosFijos)}</strong> (alquileres + sueldos base + servicios taller)</div>
                      {gastosFijos === 0 && <div style={{ fontSize:11, color:C.or, padding:"6px 10px", background:C.orBg, borderRadius:7 }}>⚠️ Registra tus gastos fijos (alquiler, sueldos) en el módulo Gastos para un cálculo preciso.</div>}
                    </>);
                  })()}
                </Card>

                <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
                  <Card style={{ padding:20 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:14, textTransform:"uppercase", letterSpacing:"0.5px" }}>Métodos de pago · {nomMes}</div>
                    {METODOS_PAGO.map(m => {
                      const tot = ventasMes.filter(v => v.mp === m).reduce((a,v) => a+v.tot, 0);
                      if (!tot) return null;
                      const pct = fIng > 0 ? Math.round((tot/fIng)*100) : 0;
                      return (
                        <div key={m} style={{ marginBottom:10 }}>
                          <div style={{ display:"flex", justifyContent:"space-between", fontSize:13, marginBottom:4 }}>
                            <span style={{ color:C.t2, fontWeight:600 }}>{m}</span>
                            <span style={{ fontWeight:700, color:C.t1 }}>{fmt(tot)} <span style={{ color:C.t4, fontWeight:400 }}>({pct}%)</span></span>
                          </div>
                          <div style={{ height:6, borderRadius:99, background:C.bg3, overflow:"hidden" }}>
                            <div style={{ height:"100%", width:`${pct}%`, background:C.ac, borderRadius:99 }}/>
                          </div>
                        </div>
                      );
                    })}
                    {ventasMes.length === 0 && <div style={{ fontSize:12, color:C.t4, fontStyle:"italic" }}>Sin ventas en {nomMes}</div>}
                  </Card>
                  <Card style={{ padding:16 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>Gastos por categoría · {nomMes}</div>
                    {["Alquiler","Servicios taller","Sueldo base","Mantenimiento","Materiales","Armador por obra","Proveedor","Comisión","Otro"].map(cat => {
                      const tot = gastosMes.filter(g => g.cat === cat).reduce((a,g)=>a+g.monto, 0);
                      if (!tot) return null;
                      const pct = fGm > 0 ? Math.round((tot/fGm)*100) : 0;
                      return (
                        <div key={cat} style={{ marginBottom:8 }}>
                          <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3 }}>
                            <span style={{ color:C.t2 }}>{cat}</span>
                            <span style={{ fontWeight:700, color:C.rd }}>{fmt(tot)} <span style={{color:C.t4,fontWeight:400}}>({pct}%)</span></span>
                          </div>
                          <div style={{ height:5, borderRadius:99, background:C.bg3, overflow:"hidden" }}>
                            <div style={{ height:"100%", width:`${pct}%`, background:C.or, borderRadius:99 }}/>
                          </div>
                        </div>
                      );
                    })}
                    {gastosMes.length === 0 && <div style={{ fontSize:12, color:C.t4, fontStyle:"italic" }}>Sin gastos en {nomMes}</div>}
                  </Card>
                </div>
              </div>
            </>);
          })()}

          {/* ═══════ ANÁLISIS ABC ═══════ */}
          {modActual === "abc" && (
            <>
              <PageTitle title="🏆 Análisis ABC" sub="Rentabilidad real por producto incluyendo costo de local"/>
              {/* Filtros */}
              <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
                {[
                  { k:"Todos",    ico:"🔍", label:`Todos (${abc.length})` },
                  { k:"A",        ico:"⭐", label:`Clase A (${abc.filter(p=>p.clase==="A").length})` },
                  { k:"B",        ico:"📈", label:`Clase B (${abc.filter(p=>p.clase==="B").length})` },
                  { k:"C",        ico:"📉", label:`Clase C (${abc.filter(p=>p.clase==="C").length})` },
                  { k:"—",        ico:"❓", label:`Sin clasif. (${abc.filter(p=>p.clase==="—").length})` },
                  { k:"Obsoletos",ico:"💀", label:`Obsoletos (${abc.filter(p=>p.obsoleto).length})` },
                ].map(f => (
                  <button key={f.k} onClick={() => setAbcFiltro(f.k)}
                    style={{ padding:"7px 16px", borderRadius:99, border:`1px solid ${abcFiltro === f.k ? ({"A":C.ac,"B":C.bl,"C":C.t4,"Obsoletos":C.rd,"Todos":C.ac,"—":C.t3}[f.k]||C.ac) : C.border}`, background: abcFiltro === f.k ? ({"A":C.ac,"B":C.bl,"C":C.t4,"Obsoletos":C.rd,"Todos":C.ac,"—":C.t3}[f.k]||C.ac)+"18" : "transparent", color: abcFiltro === f.k ? ({"A":C.ac,"B":C.bl,"C":C.t4,"Obsoletos":C.rd,"Todos":C.ac,"—":C.t3}[f.k]||C.ac) : C.t3, fontSize:13, fontWeight:700, cursor:"pointer" }}>
                    {f.ico} {f.label}
                  </button>
                ))}
              </div>
              {/* Info */}
              {abcFiltro === "Obsoletos" && abc.filter(p=>p.obsoleto).length > 0 && (
                <div style={{ background:C.rdBg, border:`1px solid ${C.rdL}44`, borderRadius:10, padding:"12px 16px", marginBottom:16, fontSize:13, color:C.t2 }}>
                  💀 Capital inmovilizado: <strong style={{color:C.rd}}>{fmt(abc.filter(p=>p.obsoleto).reduce((a,p)=>a+p.capInmov,0))}</strong> · 
                  Costo operativo proporcional: <strong style={{color:C.rd}}>{fmt(abc.filter(p=>p.obsoleto).reduce((a,p)=>a+p.costoPorProd,0))}</strong>/mes sin retorno
                </div>
              )}
              {/* KPIs filtro */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
                <KPI icon="📦" label="Productos"            value={abcFiltrado.length} color={C.ac}/>
                <KPI icon="💰" label="Ingresos acumulados"  value={fmtK(abcFiltrado.reduce((a,p)=>a+p.ing,0))} color={C.grL}/>
                <KPI icon="💸" label="Capital inmovilizado" value={fmtK(abcFiltrado.reduce((a,p)=>a+p.capInmov,0))} color={C.rd} sub="Stock sin vender"/>
                <KPI icon="📊" label="Margen promedio"      value={`${Math.round(abcFiltrado.reduce((a,p)=>a+p.margen,0)/(abcFiltrado.length||1))}%`} color={C.bl}/>
              </div>
              {/* Tabla */}
              <Card style={{ marginBottom:16 }}>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr>
                    <TH>Clase</TH><TH>Producto</TH><TH>Tipo</TH>
                    <TH right>Prom/mes</TH><TH right>Ingreso/mes</TH>
                    <TH right>Costo c/local</TH><TH right>Ganancia real</TH>
                    <TH right>Capital inmov.</TH><TH right>Margen</TH><TH>Acción</TH>
                  </tr></thead>
                  <tbody>
                    {abcFiltrado.map(p => {
                      const accion = p.obsoleto?"🚫 Liquidar":p.clase==="A"?"⭐ Priorizar":p.clase==="B"?"✅ Mantener":p.clase==="—"?"📊 Sin ventas aún":p.avg<2?"⚠️ Descontinuar":"📉 Revisar precio";
                      const ganReal = p.ganReal;
                      return (
                        <tr key={p.id} style={{ background: p.obsoleto ? C.rdBg : "transparent" }}>
                          <TD><span style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", width:28, height:28, borderRadius:8, background: ({A:C.ac,B:C.bl,C:C.t4,"—":C.t3}[p.clase]||C.t3)+"18", fontWeight:800, color:({A:C.ac,B:C.bl,C:C.t4,"—":C.t3}[p.clase]||C.t3), fontSize:13 }}>{p.obsoleto ? "💀" : p.clase}</span></TD>
                          <TD bold color={C.t1}>{p.ico} {p.n}</TD>
                          <TD><Badge color={p.tipo==="Fabricado"?C.grL:C.bl}>{p.tipo}</Badge></TD>
                          <TD right bold color={p.avg<1?C.rd:C.t1}>{p.avg} uds</TD>
                          <TD right bold color={C.grL}>{fmt(p.ingMes)}</TD>
                          <TD right color={C.rd}>{fmt(p.costMes)}</TD>
                          <TD right bold color={p.ganReal>0?C.grL:C.rd}>{p.ganReal>=0?"+":""}{fmt(p.ganReal)}</TD>
                          <TD right color={p.capInmov>500?C.or:C.t2}>{fmt(p.capInmov)}</TD>
                          <TD right><Badge color={p.margen>=40?C.grL:p.margen>=25?C.or:C.rd}>{p.margen}%</Badge></TD>
                          <TD sm color={p.obsoleto?C.rd:C.t2}>{accion}</TD>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {abcFiltrado.length === 0 && <div style={{ padding:"30px", textAlign:"center", color:C.t4 }}>No hay productos en esta categoría</div>}
              </Card>
              {/* Gráfico */}
              <Card style={{ padding:16 }}>
                <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:14, textTransform:"uppercase", letterSpacing:"0.5px" }}>
                  Ganancia bruta mensual estimada
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={[...abc].sort((a,b) => b.ganReal - a.ganReal)} margin={{left:10}}>
                    <XAxis dataKey="n" tick={{fill:C.t3,fontSize:9}} tickFormatter={v => v.split(" ").slice(0,2).join(" ")} axisLine={false} tickLine={false}/>
                    <YAxis tick={{fill:C.t4,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v => `S/${Math.round(v)}`}/>
                    <Tooltip content={<ChartTip/>}/>
                    <Bar dataKey="ganReal" name="Ganancia real/mes" radius={[4,4,0,0]}>
                      {[...abc].sort((a,b) => b.ganReal-a.ganReal).map((p,i) => (
                        <Cell key={i} fill={p.ganReal <= 0 ? C.rd : p.obsoleto ? C.rd : {A:C.ac,B:C.bl,C:C.t4}[p.clase]}/>
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </>
          )}

          {/* ═══════ PREDICCIÓN ═══════ */}
          {modActual === "prediccion" && (
            <>
              <PageTitle title="🔮 Predicción de Demanda" sub="Media móvil ponderada · Los meses recientes tienen más peso"/>
              <div style={{ background:C.blBg, border:`1px solid ${C.bl}44`, borderRadius:10, padding:"12px 16px", marginBottom:16, fontSize:13, color:C.t2 }}>
                🤖 <strong style={{color:C.bl}}>IA activa</strong> — Predicción basada en historial de 5 meses con pesos progresivos. Toma en cuenta patrones de quincena y fin de mes.
              </div>
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Producto</TH><TH right>Stock actual</TH><TH right>Pred. próx. mes</TH><TH right>Días de stock</TH><TH right>Déficit</TH><TH>Semáforo</TH><TH>Recomendación</TH></tr></thead>
                  <tbody>
                    {pred.map(p => {
                      const col = p.dias <= 7 ? C.rd : p.dias <= 15 ? C.or : p.dias <= 30 ? C.ac : C.grL;
                      const sem = p.dias <= 7 ? "🔴 Urgente" : p.dias <= 15 ? "🟠 Pronto" : p.dias <= 30 ? "🟡 Atención" : "🟢 OK";
                      const rec = p.deficit > 0 ? (p.tipo === "Fabricado" ? `Fabricar ${p.deficit + p.min} uds` : `Comprar ${p.deficit + p.min} uds`) : "Stock suficiente";
                      return (
                        <tr key={p.id}>
                          <TD bold color={C.t1}>{p.ico} {p.n}</TD>
                          <TD right bold color={p.stk <= p.min ? C.rd : C.t1}>{p.stk}</TD>
                          <TD right bold color={C.bl}>{p.pred}</TD>
                          <TD right><span style={{ fontWeight:700, color:col, background:col+"15", padding:"2px 8px", borderRadius:6, fontSize:12 }}>{p.dias >= 999 ? "∞" : p.dias} días</span></TD>
                          <TD right bold color={p.deficit > 0 ? C.rd : C.grL}>{p.deficit > 0 ? `-${p.deficit}` : "✓ OK"}</TD>
                          <TD>{sem}</TD>
                          <TD sm color={p.deficit > 0 ? C.ac : C.t3}>{rec}</TD>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </Card>
            </>
          )}

          {/* ═══════ USUARIOS ═══════ */}
          {modActual === "usuarios" && (() => {
            const ROLES = ["admin","vendedor","taller"];
            const ICOSET = ["👑","👩","👨","👧","👦","🧑","👤","🔨","💼","🌟"];
            const guardarUser = (u) => {
              const conModulos = { ...u, modulos: u.modulos || modulosPorRol(u.rol) };
              if (u.id) {
                setUsuarios(us => { sbSave("usuarios", conModulos); return us.map(x => x.id === u.id ? conModulos : x); });
                showToast("✓ Usuario actualizado");
              } else {
                const nuevoU = { ...conModulos, id: Date.now(), activo:true };
                setUsuarios(us => { sbSave("usuarios", nuevoU); return [...us, nuevoU]; });
                showToast("✓ Usuario creado");
              }
              setModal(null);
            };
            const toggleActivo = (id) => setUsuarios(us => us.map(u => {
              if (u.id !== id) return u;
              const nu = { ...u, activo:!u.activo };
              sbSave("usuarios", nu);
              return nu;
            }));
            const ROL_COLOR = { admin:C.ac, vendedor:C.bl, taller:C.grL };

            const limpiarDatosPrueba = async () => {
              if (!window.confirm("⚠️ ¿Limpiar TODOS los datos de prueba?\n\nEsto borrará permanentemente:\n• Todas las ventas\n• Todos los gastos\n• Todo el kardex\n• Todos los clientes\n• Todas las cotizaciones\n• Todos los pedidos\n• Todas las compras\n\nLos productos e inventario NO se borran.\n\nEsta acción no se puede deshacer.")) return;
              showToast("Limpiando datos...", "ok");
              // Borrar de Supabase
              const tablas = ["ventas","gastos","kardex","clientes","cotizaciones","compras","pedidos","cierres","traslados"];
              await Promise.all(tablas.map(async (tabla) => {
                const { error } = await (sb as any).from(tabla).delete().neq("id","__never__");
                if (error) console.error(`Error limpiando ${tabla}:`, error);
              }));
              // Limpiar estado local
              setVentas([]);
              setGastos([]);
              setKardex([]);
              setClientes([]);
              setCotizs([]);
              setCompras([]);
              setPedidos([]);
              setCierres([]);
              setTraslados([]);
              showToast("✓ Datos de prueba eliminados correctamente");
            };

            return (<>
              <PageTitle title="👤 Usuarios" sub={`${usuarios.length} usuarios · ${usuarios.filter(u=>u.activo).length} activos`}
                action={
                  <div style={{ display:"flex", gap:8 }}>
                    <button onClick={limpiarDatosPrueba}
                      style={{ padding:"7px 14px", borderRadius:8, border:`1px solid ${C.rdL}`, background:C.rdBg, color:C.rd, fontSize:12, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
                      🗑️ Limpiar datos prueba
                    </button>
                    <Btn variant="primary" onClick={() => setModal({ tipo:"nuevoUser", form:{ nombre:"", pass:"", rol:"vendedor", ico:"👧", modulos:modulosPorRol("vendedor") } })}>+ Nuevo Usuario</Btn>
                  </div>
                }/>
              {/* Lista usuarios */}
              <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                {usuarios.map(u => {
                  const uMods = u.modulos || modulosPorRol(u.rol);
                  return (
                    <Card key={u.id} style={{ padding:16, opacity: u.activo ? 1 : 0.55 }}>
                      <div style={{ display:"flex", gap:14, alignItems:"flex-start" }}>
                        <div style={{ width:44, height:44, borderRadius:12, background:C.acBg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:24, flexShrink:0 }}>{u.ico}</div>
                        <div style={{ flex:1 }}>
                          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                            <span style={{ fontSize:15, fontWeight:700, color:C.t1 }}>{u.nombre}</span>
                            <Badge color={ROL_COLOR[u.rol]}>{u.rol}</Badge>
                            {!u.activo && <Badge color={C.t4}>Inactivo</Badge>}
                            {u.cumple && (() => {
                              const esCumple = u.cumple.slice(5) === HOY.slice(5);
                              const fecha = new Date(u.cumple+"T12:00:00").toLocaleDateString("es-PE",{day:"2-digit",month:"long"});
                              return <span style={{ fontSize:10, color:esCumple?"#854d0e":C.t3, background:esCumple?"#fef9c3":C.bg, padding:"2px 7px", borderRadius:5 }}>{esCumple?"🎂 ¡Hoy!":"📅"} {fecha}</span>;
                            })()}
                          </div>
                          {/* Módulos habilitados */}
                          <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                            {TODOS_MODULOS.filter(m => uMods.includes(m.id)).map(m => (
                              <span key={m.id} style={{ fontSize:10, background:C.bg, border:`1px solid ${C.border}`, borderRadius:5, padding:"2px 7px", color:C.t2 }}>{m.label}</span>
                            ))}
                          </div>
                        </div>
                        <div style={{ display:"flex", gap:6, flexShrink:0, alignItems:"center" }}>
                          <button onClick={() => { if(u.id===1){showToast("No puedes desactivar al Admin","err");return;} toggleActivo(u.id); }}
                            style={{ padding:"5px 12px", borderRadius:6, border:`1px solid ${u.activo?C.grL:C.t4}`, background:u.activo?C.grBg:C.bg3, color:u.activo?C.grL:C.t3, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                            {u.activo ? "✓ Activo" : "○ Inactivo"}
                          </button>
                          <Btn sm onClick={() => setModal({ tipo:"nuevoUser", form:{ ...u, modulos: uMods } })}>✏️ Editar</Btn>
                          <Btn sm onClick={() => setModal({ tipo:"cambiarPass", userId: u.id, nombre: u.nombre })}>🔑 Contraseña</Btn>
                        </div>
                      </div>
                    </Card>
                  );
                })}
              </div>
              {/* Modal crear/editar usuario */}
              {modal?.tipo === "nuevoUser" && (
                <Modal onClose={() => setModal(null)} wide>
                  <ModalTitle>{modal.form.id ? "✏️ Editar Usuario" : "👤 Nuevo Usuario"}</ModalTitle>
                  <UsuarioForm form={modal.form} onSave={guardarUser} onCancel={() => setModal(null)} icoset={ICOSET} roles={ROLES} todosModulos={TODOS_MODULOS} modulosPorRol={modulosPorRol}/>
                </Modal>
              )}
              {modal?.tipo === "cambiarPass" && (
                <Modal onClose={() => setModal(null)}>
                  <ModalTitle>🔑 Cambiar contraseña · {modal.nombre}</ModalTitle>
                  <Inp label="Nueva contraseña" type="password" id="np1" style={{ marginBottom:10 }}/>
                  <Inp label="Confirmar contraseña" type="password" id="np2" style={{ marginBottom:16 }}/>
                  <div style={{ display:"flex", gap:8 }}>
                    <Btn variant="primary" full onClick={() => {
                      const np = (document.getElementById("np1") as HTMLInputElement)?.value || "";
                      const np2 = (document.getElementById("np2") as HTMLInputElement)?.value || "";
                      if (!np) { showToast("Escribe la nueva contraseña", "err"); return; }
                      if (np !== np2) { showToast("Las contraseñas no coinciden", "err"); return; }
                      if (np.length < 4) { showToast("Mínimo 4 caracteres", "err"); return; }
                      setUsuarios(us => us.map(u => {
                        if (u.id !== modal.userId) return u;
                        const nu = { ...u, pass: np };
                        sbSave("usuarios", nu);
                        return nu;
                      }));
                      showToast(`✓ Contraseña de ${modal.nombre} actualizada`);
                      setModal(null);
                    }}>Guardar</Btn>
                    <Btn full onClick={() => setModal(null)}>Cancelar</Btn>
                  </div>
                </Modal>
              )}
            </>);
          })()}

          {/* ═══════ ASISTENTE IA ═══════ */}
          {modActual === "ia" && (
            <>
              <PageTitle title="🤖 Asistente MoblaMel" sub="Analizo tus datos en tiempo real — sin internet externo, sin costo"/>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 300px", gap:16 }}>
                <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
                  <Card style={{ display:"flex", flexDirection:"column", height:"calc(100vh - 280px)", minHeight:380 }}>
                    <div style={{ flex:1, overflowY:"auto", padding:16, display:"flex", flexDirection:"column", gap:12 }}>
                      {iaMsgs.map((m, i) => (
                        <div key={i} style={{ display:"flex", gap:10, justifyContent: m.rol === "user" ? "flex-end" : "flex-start" }}>
                          {m.rol === "ai" && <div style={{ width:32, height:32, borderRadius:10, background:C.acBg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, flexShrink:0 }}>🤖</div>}
                          <div style={{ maxWidth:"80%", padding:"10px 14px", borderRadius:12, background: m.rol === "user" ? C.ac : C.bg, color: m.rol === "user" ? "#fff" : C.t1, fontSize:13, lineHeight:1.7, borderBottomRightRadius: m.rol === "user" ? 2 : 12, borderBottomLeftRadius: m.rol === "ai" ? 2 : 12, whiteSpace:"pre-wrap" }}>
                            {m.txt}
                          </div>
                        </div>
                      ))}
                      {iaLoad && (
                        <div style={{ display:"flex", gap:10 }}>
                          <div style={{ width:32, height:32, borderRadius:10, background:C.acBg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:18 }}>🤖</div>
                          <div style={{ padding:"10px 14px", borderRadius:12, background:C.bg, color:C.t3, fontSize:13 }}>Analizando tu negocio...</div>
                        </div>
                      )}
                    </div>
                    <div style={{ padding:"12px 14px", borderTop:`1px solid ${C.border}`, display:"flex", gap:8 }}>
                      <input value={iaInput} onChange={e => setIaInput(e.target.value)} onKeyDown={e => e.key === "Enter" && !e.shiftKey && enviarIA()}
                        style={{ flex:1, background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", color:C.t1 }}
                        placeholder="Ej: Muéstrame las ventas por vendedor este mes" />
                      <Btn variant="primary" onClick={enviarIA} disabled={iaLoad || !iaInput.trim()}>Enviar</Btn>
                    </div>
                  </Card>
                  {/* Gráfico generado por IA */}
                  {iaGrafico && (
                    <Card style={{ padding:16 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                        <div style={{ fontWeight:700, color:C.t1, fontSize:14 }}>📊 {iaGrafico.titulo}</div>
                        <button onClick={() => setIaGrafico(null)} style={{ background:"none", border:"none", color:C.t3, cursor:"pointer", fontSize:18 }}>✕</button>
                      </div>
                      {iaGrafico.tipo === "bar" && (
                        <ResponsiveContainer width="100%" height={220}>
                          <BarChart data={iaGrafico.datos} margin={{left:10,right:10,top:5,bottom:5}}>
                            <XAxis dataKey="nombre" tick={{fontSize:11}} />
                            <YAxis tick={{fontSize:11}} tickFormatter={v => `S/${v}`}/>
                            <Tooltip formatter={v => [`S/ ${v}`,""]} />
                            <Bar dataKey="valor" fill={C.ac} radius={[5,5,0,0]} label={{ position:"top", fontSize:10, fill:C.t3, formatter: v => `S/${v}` }}/>
                          </BarChart>
                        </ResponsiveContainer>
                      )}
                      {iaGrafico.tipo === "line" && (
                        <ResponsiveContainer width="100%" height={220}>
                          <LineChart data={iaGrafico.datos} margin={{left:10,right:10,top:5,bottom:5}}>
                            <XAxis dataKey="nombre" tick={{fontSize:11}}/>
                            <YAxis tick={{fontSize:11}} tickFormatter={v => `S/${v}`}/>
                            <Tooltip formatter={v => [`S/ ${v}`,""]}/>
                            <Line type="monotone" dataKey="valor" stroke={C.ac} strokeWidth={2.5} dot={{ fill:C.ac, r:4 }}/>
                          </LineChart>
                        </ResponsiveContainer>
                      )}
                      {iaGrafico.tipo === "pie" && (
                        <ResponsiveContainer width="100%" height={220}>
                          <PieChart>
                            <Pie data={iaGrafico.datos} dataKey="valor" nameKey="nombre" cx="50%" cy="50%" outerRadius={80} label={({nombre,percent}) => `${nombre} ${Math.round(percent*100)}%`} labelLine={false} fontSize={10}>
                              {iaGrafico.datos.map((_,i) => <Cell key={i} fill={[C.ac,C.bl,C.grL,C.or,C.pu,C.rd,C.acL,C.t4][i%8]}/>)}
                            </Pie>
                            <Tooltip formatter={v => [`S/ ${v}`,""]}/>
                          </PieChart>
                        </ResponsiveContainer>
                      )}
                    </Card>
                  )}
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                  <Card style={{ padding:14 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:8, textTransform:"uppercase", letterSpacing:"0.5px" }}>📊 Informes rápidos</div>
                    {[
                      { label:"Ventas por vendedor (barras)", grafico:"bar", generar: () => ({ titulo:"Ventas por vendedor", tipo:"bar", datos: usuarios.filter(u=>u.rol!=="taller").map(u => ({ nombre:u.nombre, valor: Math.round(ventas.filter(v=>v.vend===u.nombre).reduce((a,v)=>a+v.tot,0)) })).filter(d=>d.valor>0) }) },
                      { label:"Ventas por mes (línea)", grafico:"line", generar: () => ({ titulo:"Ingresos por mes", tipo:"line", datos: MESES.map((mes,i) => ({ nombre:mes, valor:Math.round(prods.reduce((a,p)=>a+p.v[i]*p.p,0)) })) }) },
                      { label:"Ingresos por categoría (torta)", grafico:"pie", generar: () => ({ titulo:"Ingresos por categoría", tipo:"pie", datos: Object.keys(CAT_C).map(cat => ({ nombre:cat, valor:Math.round(prods.filter(p=>p.cat===cat).reduce((a,p)=>a+p.v.reduce((x,y)=>x+y,0)*p.p,0)) })).filter(d=>d.valor>0) }) },
                      { label:"Stock por ubicación (barras)", grafico:"bar", generar: () => ({ titulo:"Stock total por ubicación", tipo:"bar", datos: UBICACIONES.map(u => ({ nombre:u.split("/")[0], valor:prods.filter(p=>p.ubi===u).reduce((a,p)=>a+p.stk,0) })) }) },
                      { label:"Márgenes por producto (barras)", grafico:"bar", generar: () => ({ titulo:"Margen % por producto", tipo:"bar", datos: prods.map(p => ({ nombre:p.n.split(" ").slice(0,2).join(" "), valor:Math.round(((p.p-p.c)/p.p)*100) })).sort((a,b)=>b.valor-a.valor) }) },
                      { label:"Gastos por categoría (torta)", grafico:"pie", generar: () => ({ titulo:"Gastos por categoría", tipo:"pie", datos: ["Alquiler","Servicios taller","Sueldo base","Mantenimiento","Materiales","Armador por obra","Proveedor","Comisión","Otro"].map(cat => ({ nombre:cat, valor:Math.round(gastos.filter(g=>g.cat===cat).reduce((a,g)=>a+g.monto,0)) })).filter(d=>d.valor>0) }) },
                    ].map((inf,i) => (
                      <button key={i} onClick={() => { setIaGrafico(inf.generar()); setIaMsgs(m => [...m, {rol:"user",txt:`📊 ${inf.label}`},{rol:"ai",txt:`Aquí tienes el informe: **${inf.label}** 👇`}]); }}
                        style={{ width:"100%", textAlign:"left", padding:"8px 10px", marginBottom:5, borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, color:C.t2, fontSize:12, cursor:"pointer", fontFamily:"inherit", lineHeight:1.4 }}>
                        📊 {inf.label}
                      </button>
                    ))}
                  </Card>
                  <Card style={{ padding:14 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:8, textTransform:"uppercase", letterSpacing:"0.5px" }}>💡 Preguntas al asistente</div>
                    {["¿Qué debo fabricar esta semana?","¿Cuáles son mis productos más rentables?","¿Cuándo se agotará el stock de roperos?","¿Qué productos debería descontinuar?","Dame un resumen del negocio"].map((q,i) => (
                      <button key={i} onClick={() => setIaInput(q)} style={{ width:"100%", textAlign:"left", padding:"8px 10px", marginBottom:5, borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, color:C.t2, fontSize:12, cursor:"pointer", fontFamily:"inherit", lineHeight:1.4 }}>
                        💬 {q}
                      </button>
                    ))}
                  </Card>
                </div>
              </div>
            </>
          )}


          {/* ═══════ ETIQUETAS ═══════ */}
          {modActual === "etiquetas" && (() => {
            const prodSel = prods.find(p => p.id === etProd);

            const TAMAÑOS = {
              small:    { w:62,  h:38,  fontSize:6,  barcodeH:16, label:"Pequeña A4 (6×3.8cm) · 21/hoja",   ticketera:false },
              medium:   { w:74,  h:52,  fontSize:7,  barcodeH:22, label:"Mediana A4 (7.4×5.2cm) · 10/hoja",  ticketera:false },
              large:    { w:105, h:74,  fontSize:9,  barcodeH:30, label:"Grande A4 (10.5×7.4cm) · 4/hoja",   ticketera:false },
              ticket58: { w:58,  h:40,  fontSize:7,  barcodeH:20, label:"🖨️ Ticketera 58mm",                  ticketera:true  },
              ticket80: { w:80,  h:50,  fontSize:8,  barcodeH:24, label:"🖨️ Ticketera 80mm",                  ticketera:true  },
            } as any;
            const tam = TAMAÑOS[etTamaño];

            const colsDisp = prodSel?.cols || [];
            const codeVal = prodSel ? (etCol ? `${prodSel.id}-${etCol.slice(0,3).toUpperCase()}` : prodSel.id) : "MOBLAMEL";

            // Genera barras Code-128 visual — más densas y nítidas para impresión térmica
            const genBars = (code, totalW, totalH) => {
              // Patrón determinista por caracter — barras más anchas para mayor legibilidad
              const charWidths = [2,1,3,2,1,3,2,1,2,3,1,2,3,1,2,3,2,1,2,1];
              let segs: {x:number,w:number}[] = [];
              let x = 0;
              // Barra de inicio
              segs.push({x, w:3}); x+=4;
              segs.push({x, w:1}); x+=3;
              for (let i = 0; i < code.length; i++) {
                const c = code.charCodeAt(i);
                for (let j = 0; j < 4; j++) {
                  const bw = charWidths[(c * 3 + j * 7) % charWidths.length];
                  if ((c + j) % 2 === 0) segs.push({x, w:bw});
                  x += bw + 1;
                }
              }
              // Barra de fin
              segs.push({x, w:3}); x+=4;
              segs.push({x, w:1}); x+=2;
              const totalBarsW = x;
              const scaleX = (totalW - 8) / totalBarsW;
              const bars = segs.map(s =>
                `<rect x="${4 + s.x * scaleX}" y="0" width="${Math.max(1, s.w * scaleX)}" height="${totalH}" fill="#000"/>`
              ).join("");
              return bars;
            };

            const renderEtiqueta = (_idx) => {
              const esTicket = tam.ticketera;
              const nombreCorto = prodSel ? prodSel.n.substring(0, 26) : "Producto";
              const colLabel = etCol || (prodSel?.cols[0] || "");
              const precio = prodSel ? `S/ ${prodSel.p.toFixed(2)}` : "S/ 0.00";

              if (esTicket) {
                // ── DISEÑO TICKETERA 58mm ─────────────────────────────────────
                // Ancho fijo 58mm = 219.24px @ 96dpi, usamos unidades mm directas
                const W = 58; // mm
                // Calcular altura dinámica según elementos activos
                let lines: string[] = [];
                if (etLogo)   lines.push("logo");
                if (etNombre) lines.push("nombre", "color");
                lines.push("barcode", "code");
                if (etPrecio) lines.push("precio");

                // SVG en mm para máxima precisión de impresión
                const barcodeH = 18; // mm — código de barras grande
                const lineH = 4;     // mm por línea de texto
                let totalH = 4;      // margen top
                if (etLogo)   totalH += 7;
                if (etNombre) totalH += lineH * 2;
                totalH += barcodeH + 2; // barcode
                totalH += lineH;        // código texto
                if (etPrecio) totalH += lineH + 1;
                totalH += 3; // margen bottom

                const px = W * 3.7795;    // mm → px (96dpi)
                const py = totalH * 3.7795;

                let y = 4;
                let svgParts = `<rect width="${px}" height="${py}" fill="white"/>`;

                if (etLogo) {
                  svgParts += `<rect x="0" y="${y*3.7795}" width="${px}" height="${7*3.7795}" fill="#a0714f"/>`;
                  svgParts += `<text x="${px/2}" y="${(y+5)*3.7795}" font-family="Georgia,serif" font-size="14" font-weight="bold" fill="white" text-anchor="middle" letter-spacing="3">MOBLAMEL</text>`;
                  y += 7 + 1;
                }
                if (etNombre) {
                  svgParts += `<text x="${px/2}" y="${(y+lineH-1)*3.7795}" font-family="Arial,sans-serif" font-size="13" font-weight="bold" fill="#1a1a1a" text-anchor="middle">${nombreCorto}</text>`;
                  y += lineH;
                  svgParts += `<text x="${px/2}" y="${(y+lineH-1)*3.7795}" font-family="Arial,sans-serif" font-size="11" fill="#666" text-anchor="middle">${colLabel}</text>`;
                  y += lineH;
                }
                // Código de barras — ocupa casi todo el ancho
                const bY = y * 3.7795;
                const bW = px - 8;
                svgParts += `<g transform="translate(4, ${bY})">${genBars(codeVal, bW, barcodeH * 3.7795)}</g>`;
                y += barcodeH + 2;
                // Texto del código
                svgParts += `<text x="${px/2}" y="${(y+lineH-1)*3.7795}" font-family="Courier New,monospace" font-size="11" fill="#333" text-anchor="middle" letter-spacing="1">${codeVal}</text>`;
                y += lineH + 1;
                if (etPrecio) {
                  svgParts += `<rect x="4" y="${y*3.7795}" width="${px-8}" height="${lineH*3.7795}" rx="3" fill="#a0714f"/>`;
                  svgParts += `<text x="${px/2}" y="${(y+lineH-0.8)*3.7795}" font-family="Arial,sans-serif" font-size="14" font-weight="bold" fill="white" text-anchor="middle">${precio}</text>`;
                }

                return `<svg width="${W}mm" height="${totalH}mm" viewBox="0 0 ${px} ${py}" xmlns="http://www.w3.org/2000/svg" style="display:block;max-width:100%">${svgParts}</svg>`;

              } else {
                // ── DISEÑO A4 ─────────────────────────────────────────────────
                const w = tam.w; const h = tam.h; const fs = tam.fontSize; const bh = tam.barcodeH;
                const px2 = w * 3.78; const py2 = h * 3.78;
                let curY = etLogo ? 20 : 8;
                const nombreY = curY + (etNombre ? 12 : 0);
                const colorY = nombreY + (etNombre ? 10 : 0);
                const barcodeY = colorY + 4;
                return `
<svg width="${w}mm" height="${h}mm" viewBox="0 0 ${px2} ${py2}" xmlns="http://www.w3.org/2000/svg" style="display:block">
  <rect width="${px2}" height="${py2}" fill="white" stroke="#ddd" stroke-width="0.5"/>
  ${etLogo ? `<rect width="${px2}" height="16" fill="#a0714f"/>
  <text x="8" y="11" font-family="Georgia,serif" font-size="8" font-weight="bold" fill="white" letter-spacing="1">MOBLAMEL</text>
  <circle cx="${px2-12}" cy="8" r="4" fill="${HEX_COLOR[colLabel]||"#888"}" stroke="white" stroke-width="0.5"/>` : ""}
  ${etNombre ? `<text x="${px2/2}" y="${nombreY}" font-family="Arial,sans-serif" font-size="${fs+1}" font-weight="bold" fill="#2c2016" text-anchor="middle">${nombreCorto}</text>
  <text x="${px2/2}" y="${colorY}" font-family="Arial,sans-serif" font-size="${fs}" fill="#8a7560" text-anchor="middle">${colLabel}</text>` : ""}
  <g transform="translate(8, ${barcodeY})">${genBars(codeVal, px2-16, bh)}</g>
  <text x="${px2/2}" y="${barcodeY+bh+6}" font-family="Courier New,monospace" font-size="${fs-1}" fill="#5c4a38" text-anchor="middle">${codeVal}</text>
  ${etPrecio ? `<rect x="${px2-42}" y="${py2-18}" width="38" height="14" rx="4" fill="#a0714f"/>
  <text x="${px2-23}" y="${py2-8}" font-family="Arial,sans-serif" font-size="${fs+2}" font-weight="bold" fill="white" text-anchor="middle">${precio}</text>` : ""}
</svg>`;
              }
            };

            const imprimirEtiquetas = () => {
              if (!prodSel) { alert("Selecciona un producto"); return; }
              const esTicketera = tam.ticketera;
              const ventana = window.open("", "_blank");
              if (!ventana) return;
              if (esTicketera) {
                // Modo ticketera: lista vertical continua
                const etqs = Array.from({length: etQty}, (_, i) => `<div class="etq">${renderEtiqueta(i)}</div>`).join("\n");
                ventana.document.write(`<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Etiquetas MoblaMel</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#fff; font-family:Arial,sans-serif; }
  .etq { page-break-after:always; display:flex; justify-content:center; padding:2mm 0; }
  .etq:last-child { page-break-after:avoid; }
  @media print {
    @page { margin:0; size:${etTamaño==="ticket58"?"58mm":"80mm"} auto; }
    .no-print { display:none; }
  }
</style></head>
<body>
<div class="no-print" style="padding:12px;background:#f4f6f9;display:flex;gap:10px;align-items:center;border-bottom:1px solid #e2e6ed;position:sticky;top:0">
  <strong>🖨️ ${etQty} etiqueta(s) · ${tam.label}</strong>
  <button onclick="window.print()" style="padding:7px 14px;background:#a0714f;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold">Imprimir</button>
  <button onclick="window.close()" style="padding:7px 14px;background:#fff;border:1px solid #ccc;border-radius:6px;cursor:pointer">Cerrar</button>
</div>
<div>${etqs}</div>
<script>setTimeout(()=>{window.print();},400);<\/script>
</body></html>`);
              } else {
                // Modo A4: grilla
                const etqs = Array.from({length: etQty}, (_, i) => `<div class="etq">${renderEtiqueta(i)}</div>`).join("\n");
                ventana.document.write(`<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Etiquetas MoblaMel</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#fff; font-family:Arial,sans-serif; }
  .grid { display:flex; flex-wrap:wrap; gap:3mm; padding:10mm; justify-content:flex-start; }
  .etq { page-break-inside:avoid; }
  @media print {
    @page { margin:8mm; size:A4; }
    .no-print { display:none; }
  }
</style></head>
<body>
<div class="no-print" style="padding:12px;background:#f4f6f9;display:flex;gap:10px;align-items:center;border-bottom:1px solid #e2e6ed">
  <strong>🏷️ ${etQty} etiqueta(s) de ${prodSel.n}${etCol?" · "+etCol:""}</strong>
  <button onclick="window.print()" style="padding:7px 14px;background:#a0714f;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold">🖨️ Imprimir</button>
  <button onclick="window.close()" style="padding:7px 14px;background:#fff;border:1px solid #ccc;border-radius:6px;cursor:pointer">Cerrar</button>
</div>
<div class="grid">${etqs}</div>
<script>setTimeout(()=>{window.print();},400);<\/script>
</body></html>`);
              }
              ventana.document.close();
            };

            // Toggle helper
            const Toggle = ({ val, onToggle, label, sub }) => (
              <div style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", background:C.bg, borderRadius:8 }}>
                <button onClick={onToggle}
                  style={{ width:40, height:22, borderRadius:11, border:"none", background:val?C.grL:C.t4, cursor:"pointer", position:"relative", transition:"all 0.2s", flexShrink:0 }}>
                  <div style={{ position:"absolute", top:2, left:val?20:2, width:18, height:18, borderRadius:"50%", background:"#fff", transition:"all 0.2s" }}/>
                </button>
                <div>
                  <div style={{ fontSize:12, fontWeight:700, color:C.t1 }}>{label}</div>
                  {sub && <div style={{ fontSize:10, color:C.t4 }}>{sub}</div>}
                </div>
              </div>
            );

            return (<>
              <PageTitle title="🏷️ Etiquetas" sub="Genera etiquetas para hoja A4 o ticketera térmica"/>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 320px", gap:20 }}>

                {/* Panel izquierdo */}
                <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

                  {/* Producto */}
                  <Card style={{ padding:18 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:14, textTransform:"uppercase", letterSpacing:"0.5px" }}>📦 Producto</div>
                    <div style={{ marginBottom:12 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Seleccionar producto</label>
                      <select value={etProd} onChange={e => { setEtProd(e.target.value); setEtCol(""); }}
                        style={{ width:"100%", background:C.white, border:`2px solid ${etProd?C.ac:C.border}`, borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, boxSizing:"border-box" }}>
                        <option value="">— Seleccionar producto —</option>
                        {prods.map(p => <option key={p.id} value={p.id}>{p.ico} {p.n} · {p.id}</option>)}
                      </select>
                    </div>
                    {prodSel && (
                      <div>
                        <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>Color (opcional)</label>
                        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                          <button onClick={() => setEtCol("")}
                            style={{ padding:"6px 12px", borderRadius:99, border:`2px solid ${etCol===""?C.ac:C.border}`, background:etCol===""?C.acBg:"transparent", color:etCol===""?C.ac:C.t3, fontSize:12, fontWeight:700, cursor:"pointer" }}>
                            Todos
                          </button>
                          {prodSel.cols.map(col => (
                            <button key={col} onClick={() => setEtCol(col)}
                              style={{ display:"flex", alignItems:"center", gap:6, padding:"6px 12px", borderRadius:99, border:`2px solid ${etCol===col?C.ac:C.border}`, background:etCol===col?C.acBg:"transparent", color:etCol===col?C.ac:C.t2, fontSize:12, fontWeight:700, cursor:"pointer" }}>
                              <div style={{ width:10, height:10, borderRadius:"50%", background:HEX_COLOR[col]||"#888" }}/>
                              {col}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </Card>

                  {/* Tamaño */}
                  <Card style={{ padding:18 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>📐 Tamaño / Formato</div>
                    <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                      {/* A4 */}
                      <div style={{ fontSize:10, color:C.t4, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.5px", padding:"4px 0 2px" }}>Hoja A4</div>
                      {(["small","medium","large"] as string[]).map(k => (
                        <button key={k} onClick={() => setEtTamaño(k)}
                          style={{ padding:"10px 14px", borderRadius:8, border:`2px solid ${etTamaño===k?C.ac:C.border}`, background:etTamaño===k?C.acBg:"transparent", cursor:"pointer", textAlign:"left" }}>
                          <div style={{ fontSize:12, fontWeight:700, color:etTamaño===k?C.ac:C.t2 }}>{k==="small"?"Pequeña":k==="medium"?"Mediana":"Grande"}</div>
                          <div style={{ fontSize:10, color:C.t4, marginTop:2 }}>{TAMAÑOS[k].label}</div>
                        </button>
                      ))}
                      {/* Ticketera */}
                      <div style={{ fontSize:10, color:C.t4, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.5px", padding:"8px 0 2px" }}>Ticketera térmica</div>
                      {(["ticket58","ticket80"] as string[]).map(k => (
                        <button key={k} onClick={() => setEtTamaño(k)}
                          style={{ padding:"10px 14px", borderRadius:8, border:`2px solid ${etTamaño===k?C.ac:C.border}`, background:etTamaño===k?C.acBg:"transparent", cursor:"pointer", textAlign:"left" }}>
                          <div style={{ fontSize:12, fontWeight:700, color:etTamaño===k?C.ac:C.t2 }}>{k==="ticket58"?"58mm":"80mm"}</div>
                          <div style={{ fontSize:10, color:C.t4, marginTop:2 }}>{TAMAÑOS[k].label}</div>
                        </button>
                      ))}
                    </div>
                  </Card>

                  {/* Cantidad */}
                  <Card style={{ padding:18 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>🔢 Cantidad</div>
                    <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
                      <button onClick={() => setEtQty(q => Math.max(1, q-1))}
                        style={{ width:38, height:38, borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, fontSize:20, fontWeight:700, cursor:"pointer" }}>−</button>
                      <input type="number" min="1" max="200" value={etQty} onChange={e => setEtQty(Math.max(1, Math.min(200, parseInt(e.target.value)||1)))}
                        style={{ flex:1, textAlign:"center", fontSize:22, fontWeight:800, background:C.white, border:`2px solid ${C.ac}`, borderRadius:8, padding:"6px", outline:"none", color:C.t1, fontFamily:"inherit" }}/>
                      <button onClick={() => setEtQty(q => Math.min(200, q+1))}
                        style={{ width:38, height:38, borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, fontSize:20, fontWeight:700, cursor:"pointer" }}>+</button>
                    </div>
                    <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                      {[1,2,5,10,20,50,100].map(n => (
                        <button key={n} onClick={() => setEtQty(n)}
                          style={{ padding:"4px 10px", borderRadius:99, border:`1px solid ${etQty===n?C.ac:C.border}`, background:etQty===n?C.acBg:"transparent", fontSize:11, fontWeight:700, cursor:"pointer", color:etQty===n?C.ac:C.t3 }}>
                          {n}
                        </button>
                      ))}
                    </div>
                  </Card>

                  {/* Opciones de contenido */}
                  <Card style={{ padding:18 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>⚙️ Contenido de la etiqueta</div>
                    <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                      <Toggle val={etLogo}   onToggle={() => setEtLogo(v=>!v)}   label="Logo MOBLAMEL" sub="Franja café con el nombre de la marca"/>
                      <Toggle val={etNombre} onToggle={() => setEtNombre(v=>!v)} label="Nombre del producto" sub="Muestra el nombre y color del mueble"/>
                      <Toggle val={etPrecio} onToggle={() => setEtPrecio(v=>!v)} label="Precio de venta" sub={`Muestra S/ ${prodSel?prodSel.p.toFixed(2):"—"} en la etiqueta`}/>
                    </div>
                  </Card>

                  {/* Botón imprimir */}
                  <button onClick={imprimirEtiquetas} disabled={!prodSel}
                    style={{ padding:"16px", borderRadius:12, border:"none", background:prodSel?"linear-gradient(135deg,#a0714f,#7d5538)":"#e2e6ed", color:prodSel?"#fff":"#a0aec0", fontSize:16, fontWeight:800, cursor:prodSel?"pointer":"not-allowed", fontFamily:"inherit", boxShadow:prodSel?"0 4px 16px rgba(100,60,20,0.3)":"none" }}>
                    🖨️ Generar e Imprimir {etQty} Etiqueta{etQty!==1?"s":""} {tam.ticketera?"(Ticketera)":"(A4)"}
                  </button>
                </div>

                {/* Panel derecho: previsualización */}
                <div>
                  <Card style={{ padding:18, position:"sticky", top:16 }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.t3, marginBottom:14, textTransform:"uppercase", letterSpacing:"0.5px" }}>👁️ Previsualización</div>
                    {prodSel ? (
                      <>
                        <div style={{ background:"#f4f6f9", borderRadius:10, padding:16, display:"flex", justifyContent:"center", alignItems:"center", minHeight:160, marginBottom:12, overflow:"hidden" }}
                          dangerouslySetInnerHTML={{ __html: renderEtiqueta(0) }}/>
                        <div style={{ fontSize:10, color:C.t4, textAlign:"center", marginBottom:12 }}>
                          {tam.ticketera ? "Vista previa ticketera — escala aproximada" : "Vista previa A4 — la impresión es más precisa"}
                        </div>
                        <div style={{ background:C.acBg, borderRadius:8, padding:"10px 12px", fontSize:11, color:C.t2 }}>
                          <div style={{ fontWeight:700, color:C.ac, marginBottom:6 }}>📋 Resumen</div>
                          <div>Producto: <strong>{prodSel.n}</strong></div>
                          {etCol && <div>Color: <strong>{etCol}</strong></div>}
                          <div>Cantidad: <strong>{etQty} etiqueta{etQty!==1?"s":""}</strong></div>
                          <div>Formato: <strong>{TAMAÑOS[etTamaño].label.split("·")[0].trim()}</strong></div>
                          <div>Logo: <strong>{etLogo?"Incluido":"Oculto"}</strong></div>
                          <div>Nombre: <strong>{etNombre?"Incluido":"Oculto"}</strong></div>
                          <div>Precio: <strong>{etPrecio?"Incluido":"Oculto"}</strong></div>
                        </div>
                      </>
                    ) : (
                      <div style={{ textAlign:"center", padding:"60px 20px", color:C.t4 }}>
                        <div style={{ fontSize:48, marginBottom:12 }}>🏷️</div>
                        <div style={{ fontSize:13 }}>Selecciona un producto para ver la previsualización</div>
                      </div>
                    )}
                  </Card>
                </div>
              </div>
            </>);
          })()}


          {/* ═══════ VENDEDORES ═══════ */}
          {modActual === "vendedores" && (() => {
            const vendUsers = usuarios.filter(u => u.rol !== "taller" && u.activo);
            const pct = (comisionPct || 0) / 100;
            const ventasFiltradas = ventas.filter(v => v.f >= vendFiltroDesde && v.f <= vendFiltroHasta && !v.esAnticipo);
            const ranking = vendUsers.map(u => {
              const mis = ventasFiltradas.filter(v => v.vend === u.nombre);
              const ingresos = mis.reduce((a,v) => a+v.tot, 0);
              const costos   = mis.reduce((a,v) => a + v.items.reduce((b,i) => b + (i.c||0)*i.q, 0), 0);
              const utilidad = ingresos - costos;
              const margen   = ingresos > 0 ? Math.round(((ingresos-costos)/ingresos)*100) : 0;
              const comision = +(ingresos * pct).toFixed(2);
              // Sueldo base viene del perfil del vendedor (4 semanas del mes)
              const sueldoBase = (u.sueldoSemanal||0) * 4;
              const utilidadNeta = utilidad - comision - sueldoBase;
              return { ...u, qty: mis.length, ingresos, costos, utilidad, utilidadNeta, margen, comision, sueldoBase };
            }).sort((a,b) => b.ingresos - a.ingresos);
            const rankingFilt = ranking.filter(v => !vendBusq || v.nombre.toLowerCase().includes(vendBusq.toLowerCase()));
            const maxIng = ranking[0]?.ingresos || 1;
            const pagarComision = (vend) => {
              if (vend.comision <= 0) { showToast("Sin comisión pendiente", "err"); return; }
              const ng = { id:"G"+uid(), f:HOY, cat:"Comisión", desc:`Comisión ${comisionPct}% → ${vend.nombre} (${vendFiltroDesde} al ${vendFiltroHasta})`, prov:vend.nombre, monto:vend.comision };
              setGastos(gs => { sbSave("gastos", ng); return [ng, ...gs]; });
              showToast(`✓ Comisión S/ ${vend.comision.toFixed(2)} de ${vend.nombre} registrada en gastos`);
            };
            return (<>
              <PageTitle title="🏆 Vendedores" sub="Ranking, márgenes y comisiones del periodo"/>
              <Card style={{ padding:"12px 16px", marginBottom:16 }}>
                <div style={{ display:"flex", gap:12, alignItems:"flex-end", flexWrap:"wrap" }}>
                  {/* Buscador */}
                  <div style={{ flex:1, minWidth:160, position:"relative" }}>
                    <span style={{ position:"absolute", left:9, top:"50%", transform:"translateY(-50%)", color:C.t4, fontSize:13 }}>🔍</span>
                    <input value={vendBusq} onChange={e => setVendBusq(e.target.value)}
                      placeholder="Buscar vendedor..."
                      style={{ width:"100%", background:C.white, border:`1px solid ${vendBusq?C.ac:C.border}`, borderRadius:8, padding:"8px 10px 8px 28px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
                  </div>
                  <Inp label="Desde" type="date" value={vendFiltroDesde} onChange={e => setVendFiltroDesde(e.target.value)} style={{ width:150 }}/>
                  <Inp label="Hasta" type="date" value={vendFiltroHasta} onChange={e => setVendFiltroHasta(e.target.value)} style={{ width:150 }}/>
                  <div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Comisión %</label>
                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                      <input type="number" min="0" max="100" step="0.5" value={comisionPct}
                        onChange={e => setComisionPct(parseFloat(e.target.value) || 0)}
                        style={{ width:72, background:C.white, border:`2px solid ${C.ac}`, borderRadius:8, padding:"7px 10px", fontSize:14, fontWeight:700, outline:"none", textAlign:"center", color:C.ac, fontFamily:"inherit" }}/>
                      <span style={{ fontSize:14, fontWeight:700, color:C.ac }}>%</span>
                      <div style={{ display:"flex", gap:4 }}>
                        {[1,2,3,5].map(v => (
                          <button key={v} onClick={() => setComisionPct(v)}
                            style={{ padding:"4px 8px", borderRadius:6, border:`1px solid ${comisionPct===v?C.ac:C.border}`, background:comisionPct===v?C.acBg:"transparent", color:comisionPct===v?C.ac:C.t3, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                            {v}%
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                  <div style={{ fontSize:12, color:C.t3, paddingBottom:4 }}>
                    {ventasFiltradas.length} ventas · Total <strong style={{color:C.ac}}>{fmt(ventasFiltradas.reduce((a,v)=>a+v.tot,0))}</strong>
                  </div>
                </div>
              </Card>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))", gap:14, marginBottom:20 }}>
                {rankingFilt.map((vend, i) => (
                  <Card key={vend.id} style={{ padding:18, border: i===0 ? `2px solid ${C.ac}` : undefined, position:"relative", overflow:"hidden" }}>
                    {i === 0 && <div style={{ position:"absolute", top:10, right:12, fontSize:18 }}>🥇</div>}
                    {i === 1 && <div style={{ position:"absolute", top:10, right:12, fontSize:18 }}>🥈</div>}
                    {i === 2 && <div style={{ position:"absolute", top:10, right:12, fontSize:18 }}>🥉</div>}
                    <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
                      <div style={{ width:42, height:42, borderRadius:12, background:C.acBg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22 }}>{vend.ico}</div>
                      <div>
                        <div style={{ fontSize:15, fontWeight:700, color:C.t1 }}>{vend.nombre}</div>
                        <div style={{ fontSize:11, color:C.t3, textTransform:"capitalize" }}>{vend.rol} · {vend.qty} ventas</div>
                      </div>
                    </div>
                    <div style={{ marginBottom:8 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4, fontSize:13 }}>
                        <span style={{ color:C.t3 }}>Ingresos</span>
                        <span style={{ fontWeight:800, color:C.ac, fontSize:16 }}>{fmt(vend.ingresos)}</span>
                      </div>
                      <div style={{ height:6, borderRadius:99, background:C.bg3, overflow:"hidden", marginBottom:10 }}>
                        <div style={{ height:"100%", width:`${(vend.ingresos/maxIng)*100}%`, background: i===0?C.ac:C.t4, borderRadius:99, transition:"width 0.5s" }}/>
                      </div>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8 }}>
                        <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", textAlign:"center" }}>
                          <div style={{ fontSize:11, color:C.t3 }}>Margen</div>
                          <div style={{ fontSize:14, fontWeight:700, color: vend.margen>=40?C.grL:vend.margen>=25?C.or:C.rd }}>{vend.margen}%</div>
                        </div>
                        <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", textAlign:"center" }}>
                          <div style={{ fontSize:11, color:C.t3 }}>Comisión {comisionPct}%</div>
                          <div style={{ fontSize:14, fontWeight:700, color:C.bl }}>{fmt(vend.comision)}</div>
                        </div>
                        <div style={{ background:C.bg, borderRadius:8, padding:"8px 10px", textAlign:"center" }}>
                          <div style={{ fontSize:11, color:C.t3 }}>Ventas</div>
                          <div style={{ fontSize:14, fontWeight:700, color:C.t1 }}>{vend.qty}</div>
                        </div>
                      </div>
                      {/* Reporte utilidad */}
                      <div style={{ marginTop:10, background:vend.utilidadNeta>=0?C.grBg:C.rdBg, borderRadius:8, padding:"10px 12px" }}>
                        <div style={{ fontSize:11, fontWeight:700, color:C.t3, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Reporte de Utilidad</div>
                        {[
                          ["Ingresos", vend.ingresos, C.grL],
                          ["− Costo productos", -vend.costos, C.rd],
                          ["− Comisión", -vend.comision, C.or],
                          ...(vend.sueldoBase>0 ? [["− Sueldo base", -vend.sueldoBase, C.or]] : []),
                        ].map(([label, val, clr])=>(
                          <div key={label} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3 }}>
                            <span style={{ color:C.t3 }}>{label}</span>
                            <span style={{ fontWeight:700, color:clr }}>{val<0?"−":""}{fmt(Math.abs(val))}</span>
                          </div>
                        ))}
                        <div style={{ borderTop:`1px solid ${C.border}`, marginTop:6, paddingTop:6, display:"flex", justifyContent:"space-between", fontSize:13, fontWeight:800 }}>
                          <span style={{ color:C.t2 }}>Utilidad neta</span>
                          <span style={{ color:vend.utilidadNeta>=0?C.grL:C.rd }}>{fmt(vend.utilidadNeta)}</span>
                        </div>
                      </div>
                    </div>
                    <Btn variant="green" full onClick={() => pagarComision(vend)} disabled={vend.comision <= 0}>
                      💸 Pagar comisión {fmt(vend.comision)}
                    </Btn>
                  </Card>
                ))}
              </div>
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>#</TH><TH>Vendedor</TH><TH right>Ventas</TH><TH right>Ingresos</TH><TH right>Costo</TH><TH right>Margen</TH><TH right>Comisión {comisionPct}%</TH><TH right>Utilidad neta</TH></tr></thead>
                  <tbody>
                    {rankingFilt.map((vend, i) => (
                      <tr key={vend.id}>
                        <TD bold color={i===0?C.ac:C.t3}>{i+1}</TD>
                        <TD bold color={C.t1}>{vend.ico} {vend.nombre}</TD>
                        <TD right>{vend.qty}</TD>
                        <TD right bold color={C.grL}>{fmt(vend.ingresos)}</TD>
                        <TD right color={C.rd}>{fmt(vend.costos)}</TD>
                        <TD right><Badge color={vend.margen>=40?C.grL:vend.margen>=25?C.or:C.rd}>{vend.margen}%</Badge></TD>
                        <TD right bold color={C.bl}>{fmt(vend.comision)}</TD>
                        <TD right bold color={vend.utilidadNeta>=0?C.grL:C.rd}>{fmt(vend.utilidadNeta)}</TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </>);
          })()}

          {/* ═══════ CLIENTES ═══════ */}
          {modActual === "clientes" && (() => {
            const VIP_UMBRAL = 3;
            // Vendedores solo ven clientes que ellos han atendido
            const clientesBase = isAdmin ? clientes : clientes.filter(c =>
              ventas.some(v => v.vend === user.nombre && v.cli === c.nombre)
            );
            const cliConHistorial = clientesBase.map(c => {
              const compras_cli = (isAdmin ? ventas : ventas.filter(v => v.vend === user.nombre)).filter(v => v.cli === c.nombre);
              const totalGastado = compras_cli.reduce((a,v) => a+v.tot, 0);
              const esVip = compras_cli.length >= VIP_UMBRAL || totalGastado >= 1000;
              return { ...c, compras_cli, totalGastado, esVip };
            }).sort((a,b) => b.totalGastado - a.totalGastado);
            return (<>
              <PageTitle title="👥 Clientes" sub={`${clientesBase.length} clientes${!isAdmin ? ` atendidos por ${user.nombre}` : " registrados"}`}
                action={
                  <div style={{display:"flex",gap:8}}>
                    <Btn onClick={() => exportXLSX("clientes", ["Nombre","Celular","Dirección","Notas","VIP","N° Compras","Total Gastado"], cliConHistorial.map(c=>[c.nombre,c.cel||"",c.dir||"",c.notas||"",c.esVip?"Sí":"No",c.compras_cli.length,c.totalGastado]))}>⬇️ Excel</Btn>
                    <Btn variant="primary" onClick={() => setModal("cliente")}>+ Nuevo Cliente</Btn>
                  </div>
                }/>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:16 }}>
                <KPI icon="👥" label="Total clientes"   value={clientes.length} color={C.bl}/>
                <KPI icon="⭐" label="Clientes VIP"     value={cliConHistorial.filter(c=>c.esVip).length} color={C.ac} sub={`≥${VIP_UMBRAL} compras o S/1000`}/>
                <KPI icon="💰" label="Ticket promedio"  value={fmt(cliConHistorial.reduce((a,c)=>a+c.totalGastado,0)/(cliConHistorial.filter(c=>c.totalGastado>0).length||1))} color={C.grL}/>
              </div>
              <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                {cliConHistorial.map(c => (
                  <Card key={c.id} style={{ padding:16 }}>
                    <div style={{ display:"flex", gap:14, alignItems:"flex-start" }}>
                      <div style={{ width:44, height:44, borderRadius:12, background: c.esVip ? C.acBg : C.bg, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22, flexShrink:0, border: c.esVip ? `2px solid ${C.ac}` : undefined }}>
                        {c.esVip ? "⭐" : "👤"}
                      </div>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:2 }}>
                          <span style={{ fontSize:15, fontWeight:700, color:C.t1 }}>{c.nombre}</span>
                          {c.esVip && <Badge color={C.ac}>VIP</Badge>}
                        </div>
                        <div style={{ fontSize:12, color:C.t3, marginBottom:8 }}>
                          📍 {c.dir || "—"} {c.notas && `· ${c.notas}`}
                        </div>
                        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                          {c.compras_cli.slice(0,3).map(v => (
                            <div key={v.id} style={{ background:C.bg, borderRadius:8, padding:"5px 10px", fontSize:11 }}>
                              <span style={{ color:C.t3 }}>{v.f} </span>
                              <span style={{ fontWeight:700, color:C.ac }}>{fmt(v.tot)}</span>
                            </div>
                          ))}
                          {c.compras_cli.length === 0 && <span style={{ fontSize:12, color:C.t4, fontStyle:"italic" }}>Sin compras aún</span>}
                        </div>
                      </div>
                      <div style={{ display:"flex", flexDirection:"column", gap:6, alignItems:"flex-end", flexShrink:0 }}>
                        <div style={{ fontSize:16, fontWeight:800, color:C.t1 }}>{fmt(c.totalGastado)}</div>
                        <div style={{ fontSize:11, color:C.t3 }}>{c.compras_cli.length} compras</div>
                        {c.cel && (
                          <a href={`https://wa.me/51${c.cel}`} target="_blank" rel="noopener noreferrer"
                            style={{ display:"flex", alignItems:"center", gap:5, background:"#25D366", color:"#fff", borderRadius:8, padding:"6px 12px", fontSize:12, fontWeight:700, textDecoration:"none" }}>
                            💬 WhatsApp
                          </a>
                        )}
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </>);
          })()}

          {/* ═══════ COTIZACIONES ═══════ */}
          {modActual === "cotiz" && (() => {
            const EST_COLOR = { Pendiente:C.or, Vista:C.bl, Aceptada:C.grL, Rechazada:C.rd, "Convertida":C.pu };
            // Vendedores solo ven sus propias cotizaciones
            const cotizBase = isAdmin ? cotizs : cotizs.filter(q => q.vend === user.nombre);
            const guardarCotiz = (q) => {
              if (q.id) {
                setCotizs(qs => { sbSave("cotizaciones", q); return qs.map(x => x.id === q.id ? q : x); });
                showToast("✓ Cotización actualizada");
              } else {
                const nq = { ...q, id:"Q"+uid(), f:HOY };
                setCotizs(qs => { sbSave("cotizaciones", nq); return [nq, ...qs]; });
                showToast("✓ Cotización creada");
              }
              setCotizForm(null);
            };
            const convertirVenta = (q) => {
              setCotizPagoModal({ q, pagos:[{met:"Efectivo", monto:String(q.tot)}] });
            };
            const confirmarConversionVenta = (q, mpLabel, pagosArr) => {
              const pfx = "B001";
              const num = `${pfx}-${String(numB).padStart(5,"0")}`;
              const nv = { id:"V"+uid(), f:HOY, cli:q.cli, vend:q.vend,
                items: q.items.map(i => ({ id:i.id, n:i.n, col:i.col, q:i.q, p:i.p, c: prods.find(p=>p.id===i.id)?.c || 0 })),
                tot:q.tot, mp:mpLabel, comp:"Boleta", num };
              setVentas(v => { sbSave("ventas", nv); return [nv, ...v]; });
              const cotizConvertida = { ...cotizs.find(x => x.id === q.id), est:"Convertida" };
              setCotizs(qs => { sbSave("cotizaciones", cotizConvertida); return qs.map(x => x.id === q.id ? cotizConvertida : x); });
              setNumB(n => { const nv2=n+1; sbSetConfig("numB",nv2); return nv2; });
              setVoucher({ items:nv.items, subtotal:q.tot, desc:0, recPOS:0, total:q.tot,
                pagos: pagosArr||[{met:mpLabel, monto:q.tot}], vuelto:0, cliente:q.cli, vend:q.vend,
                comp:"Boleta", num, fecha:HOY });
              setCotizPagoModal(null);
              showToast(`✓ Venta ${num} · ${mpLabel}`);
            };
            const generarTextoWA = (q) => {
              const lineas = q.items.map(i => `  • ${i.ico||""} ${i.n} (${i.col}) × ${i.q} = S/ ${(i.q*i.p).toFixed(2)}`).join("\n");
              return `*COTIZACIÓN - Moblamel*\n\nFecha: ${q.f}\nCliente: ${q.cli}\n\n${lineas}\n\n*TOTAL: S/ ${q.tot.toFixed(2)}*\n\n${q.nota ? `Nota: ${q.nota}` : ""}\n\n¡Gracias por contactarnos! 🪵`;
            };
            return (<>
              <PageTitle title="📋 Cotizaciones" sub={`${cotizBase.length} cotizaciones${!isAdmin ? ` de ${user.nombre}` : ""}`}
                action={
                  <div style={{display:"flex",gap:8}}>
                    <Btn onClick={() => exportXLSX("cotizaciones", ["Fecha","Cliente","Celular","Vendedor","Productos","Total","Estado","Nota"], cotizBase.map(q=>[q.f,q.cli,q.cel||"",q.vend,q.items.map(i=>`${i.n} x${i.q}`).join(" / "),q.tot,q.est,q.nota||""]))}>⬇️ Excel</Btn>
                    <Btn variant="primary" onClick={() => setCotizForm({cli:"",cel:"",vend:user.nombre,items:[],tot:0,est:"Pendiente",nota:""})}>+ Nueva Cotización</Btn>
                  </div>
                }/>
              {/* Modal global: método de pago al convertir cotización */}
              {cotizPagoModal && (
                <div onClick={e=>{if(e.target===e.currentTarget)setCotizPagoModal(null)}}
                  style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1200, padding:16 }}>
                  <div style={{ background:"#fff", borderRadius:14, padding:24, width:420, maxWidth:"100%", boxShadow:"0 20px 60px rgba(0,0,0,0.2)", maxHeight:"90vh", overflowY:"auto" }}>
                    <div style={{ fontSize:15, fontWeight:700, color:C.t1, marginBottom:4 }}>💰 Convertir a venta</div>
                    <div style={{ fontSize:13, color:C.t3, marginBottom:16, paddingBottom:12, borderBottom:`1px solid ${C.border}` }}>
                      {cotizPagoModal.q?.cli} · Total: <strong>{fmt(cotizPagoModal.q?.tot||0)}</strong>
                    </div>
                    <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>Formas de pago (puede ser más de una)</label>
                    {/* Lista de pagos */}
                    {(cotizPagoModal.pagos||[{met:"Efectivo",monto:""}]).map((pago, i) => (
                      <div key={i} style={{ display:"flex", gap:8, marginBottom:8, alignItems:"center" }}>
                        <select value={pago.met} onChange={e => setCotizPagoModal(m => ({ ...m, pagos: m.pagos.map((p,j)=>j===i?{...p,met:e.target.value}:p) }))}
                          style={{ flex:1, background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1 }}>
                          {["Efectivo","Yape/Plin","Trans. BCP","Trans. Interbank","POS Tarjeta"].map(mp=><option key={mp}>{mp}</option>)}
                        </select>
                        <div style={{ position:"relative", width:100 }}>
                          <span style={{ position:"absolute", left:8, top:"50%", transform:"translateY(-50%)", fontSize:11, color:C.t4 }}>S/</span>
                          <input inputMode="decimal" value={pago.monto}
                            onChange={e => setCotizPagoModal(m => ({ ...m, pagos: m.pagos.map((p,j)=>j===i?{...p,monto:e.target.value.replace(/[^0-9.]/g,"")}:p) }))}
                            placeholder="0.00"
                            style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 8px 8px 22px", fontSize:13, fontWeight:700, outline:"none", color:C.t1, boxSizing:"border-box" }}/>
                        </div>
                        {i > 0 && <button onClick={() => setCotizPagoModal(m=>({...m,pagos:m.pagos.filter((_,j)=>j!==i)}))} style={{background:"none",border:"none",color:C.rdL,cursor:"pointer",fontSize:18}}>✕</button>}
                      </div>
                    ))}
                    <button onClick={() => setCotizPagoModal(m=>({...m,pagos:[...(m.pagos||[]),{met:"Efectivo",monto:""}]}))}
                      style={{ width:"100%", padding:"7px", borderRadius:8, border:`1px dashed ${C.ac}`, background:C.acBg, color:C.ac, fontSize:12, fontWeight:700, cursor:"pointer", marginBottom:12 }}>
                      + Agregar otro método
                    </button>
                    {/* Resumen */}
                    {(() => {
                      const pagos = cotizPagoModal.pagos||[];
                      const totalPagado = pagos.reduce((a,p)=>a+(parseFloat(p.monto)||0),0);
                      const falta = (cotizPagoModal.q?.tot||0) - totalPagado;
                      return (
                        <div style={{ background:Math.abs(falta)<0.01?C.grBg:C.rdBg, borderRadius:8, padding:"8px 12px", marginBottom:14, display:"flex", justifyContent:"space-between", fontSize:13, fontWeight:700 }}>
                          <span style={{color:C.t2}}>Pagado: {fmt(totalPagado)}</span>
                          <span style={{color:Math.abs(falta)<0.01?C.grL:C.rd}}>{Math.abs(falta)<0.01?"✓ Completo":`Falta: ${fmt(Math.abs(falta))}`}</span>
                        </div>
                      );
                    })()}
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={() => setCotizPagoModal(null)} style={{flex:1}}>Cancelar</Btn>
                      <Btn variant="primary" onClick={() => {
                        const pagos = cotizPagoModal.pagos||[];
                        const totalPagado = pagos.reduce((a,p)=>a+(parseFloat(p.monto)||0),0);
                        if (Math.abs(totalPagado-(cotizPagoModal.q?.tot||0))>0.01) { showToast("El total pagado no coincide","err"); return; }
                        const mpLabel = pagos.length===1 ? pagos[0].met : pagos.map(p=>`${p.met} S/${parseFloat(p.monto).toFixed(0)}`).join(" + ");
                        confirmarConversionVenta(cotizPagoModal.q, mpLabel, pagos);
                      }} style={{flex:1}}>✓ Confirmar</Btn>
                    </div>
                  </div>
                </div>
              )}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:16 }}>
                {["Pendiente","Vista","Aceptada","Convertida"].map(est => (
                  <KPI key={est} icon={{Pendiente:"⏳",Vista:"👁️",Aceptada:"✅",Convertida:"💰"}[est]} label={est}
                    value={cotizBase.filter(q=>q.est===est).length} color={EST_COLOR[est]}/>
                ))}
              </div>
              {cotizForm !== null && (
                <CotizFormModal form={cotizForm} setForm={setCotizForm} prods={prods} usuarios={usuarios} onSave={guardarCotiz} onCancel={() => setCotizForm(null)}/>
              )}
              <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                {cotizBase.map(q => (
                  <Card key={q.id} style={{ padding:16 }}>
                    <div style={{ display:"flex", gap:12, alignItems:"flex-start" }}>
                      <div style={{ flex:1 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:4 }}>
                          <span style={{ fontSize:14, fontWeight:700, color:C.t1 }}>{q.cli}</span>
                          <Badge color={EST_COLOR[q.est]||C.t3}>{q.est}</Badge>
                          <span style={{ fontSize:11, color:C.t4, marginLeft:4 }}>{q.f}</span>
                        </div>
                        <div style={{ fontSize:12, color:C.t3, marginBottom:6 }}>Vendedor: {q.vend} {q.cel && `· 📱 ${q.cel}`}</div>
                        <div style={{ fontSize:12, color:C.t2 }}>{q.items.map(i => `${i.n} ×${i.q}`).join(", ")}</div>
                        {q.nota && <div style={{ fontSize:11, color:C.t4, marginTop:4, fontStyle:"italic" }}>{q.nota}</div>}
                      </div>
                      <div style={{ display:"flex", flexDirection:"column", gap:6, alignItems:"flex-end", flexShrink:0 }}>
                        <div style={{ fontSize:18, fontWeight:800, color:C.ac }}>{fmt(q.tot)}</div>
                        <div style={{ display:"flex", gap:6, flexWrap:"wrap", justifyContent:"flex-end" }}>
                          <Btn sm onClick={() => setCotizForm(q)}>✏️ Editar</Btn>
                          {q.cel && (
                            <a href={`https://wa.me/51${q.cel}?text=${encodeURIComponent(generarTextoWA(q))}`} target="_blank" rel="noopener noreferrer"
                              style={{ display:"inline-flex", alignItems:"center", gap:4, background:"#25D366", color:"#fff", borderRadius:7, padding:"4px 10px", fontSize:11, fontWeight:700, textDecoration:"none" }}>
                              📲 Enviar
                            </a>
                          )}
                          {q.est === "Aceptada" && (
                            <Btn sm variant="green" onClick={() => convertirVenta(q)}>💰 Convertir venta</Btn>
                          )}
                          {["Pendiente","Vista"].includes(q.est) && (
                            <Btn sm onClick={() => setCotizs(qs => qs.map(x => { if (x.id!==q.id) return x; const nq={...x,est:"Aceptada"}; sbSave("cotizaciones",nq); return nq; }))}>✅ Aceptar</Btn>
                          )}
                        </div>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </>);
          })()}

          {/* ═══════ COMPRAS ═══════ */}
          {modActual === "compras" && (() => {
            const registrarCompra = (cm) => {
              const items = cm.items || [{ pid:cm.pid, prod:cm.prod, col:cm.col, qty:parseInt(cm.qty)||0, cu:parseFloat(cm.cu)||0, ubiDest:cm.ubiDest||"Taller/Almacén" }];
              let nuevosProds = [...prods];
              let nuevosKardex = [...kardex];
              const comprasReg = [];
              let totalCompra = 0;
              items.forEach(item => {
                if (!item.pid || !item.qty || !item.cu) return;
                const prod = nuevosProds.find(p => p.id === item.pid);
                if (!prod) return;
                const qty = parseInt(item.qty)||0;
                const cu  = parseFloat(item.cu)||0;
                const ubiDest = item.ubiDest||"Taller/Almacén";
                const col = item.col||prod.cols[0]||"";
                const stkActual = getTotalStk(prod);
                const costoNuevo = stkActual > 0 ? +((stkActual*prod.c+qty*cu)/(stkActual+qty)).toFixed(2) : cu;
                nuevosProds = nuevosProds.map(p => p.id === item.pid ? { ...agregarStk(prod, qty, ubiDest, col), c:costoNuevo } : p);
                nuevosKardex = [{ id:"K"+uid(), f:cm.f||HOY, pid:item.pid, prod:item.prod||prod.n, col,
                  tipo:"Compra", desc:`Compra ${cm.prov} → ${ubiDest}`, ent:qty, sal:0, saldo:stkActual+qty, costo:costoNuevo }, ...nuevosKardex];
                const subtotal = +(qty*cu).toFixed(2);
                totalCompra += subtotal;
                comprasReg.push({ id:"CM"+uid(), f:cm.f||HOY, prov:cm.prov, pid:item.pid, prod:item.prod||prod.n, col, qty, cu, total:subtotal, ubiDest, nota:cm.nota||"" });
              });
              setProds(nuevosProds);
              nuevosProds.filter(p => comprasReg.some(c => c.pid === p.id)).forEach(p => sbSave("productos", p));
              setKardex(nuevosKardex);
              nuevosKardex.filter(k => k.tipo === "Compra").forEach(k => sbSave("kardex", k));
              setCompras(cs => { comprasReg.forEach(c => sbSave("compras", c)); return [...comprasReg, ...cs]; });
              showToast(`✓ ${comprasReg.length} producto(s) registrados · S/${totalCompra.toFixed(2)}`);
              setCompraForm(null);
            };
            return (<>
              <PageTitle title="🏭 Compras" sub="Registro de compras a proveedores · Actualiza stock y costo promedio"
                action={
                  <div style={{display:"flex",gap:8}}>
                    <Btn onClick={() => exportXLSX("compras", ["Fecha","Proveedor","Producto","Color","Cantidad","Costo unit.","Total","Nota"],
                      compras.filter(c => (!compraFiltDesde||c.f>=compraFiltDesde) && (!compraFiltHasta||c.f<=compraFiltHasta))
                      .map(c=>[c.f,c.prov,c.prod,c.col||"Varios",c.qty,c.cu,c.total,c.nota||""]))}>⬇️ Excel</Btn>
                    <Btn variant="primary" onClick={() => setCompraForm({ f:HOY, prov:"", nota:"", ubiDest:"Taller/Almacén", items:[] })}>+ Registrar Compra</Btn>
                  </div>
                }/>
              <Card style={{ padding:"10px 14px", marginBottom:12 }}>
                <div style={{ display:"flex", gap:10, alignItems:"flex-end", flexWrap:"wrap" }}>
                  <Inp label="Desde" type="date" value={compraFiltDesde} onChange={e => setCompraFiltDesde(e.target.value)} style={{ width:145 }}/>
                  <Inp label="Hasta" type="date" value={compraFiltHasta} onChange={e => setCompraFiltHasta(e.target.value)} style={{ width:145 }}/>
                  {(compraFiltDesde || compraFiltHasta) && <Btn onClick={() => { setCompraFiltDesde(""); setCompraFiltHasta(""); }}>✕ Limpiar</Btn>}
                  <div style={{ fontSize:12, color:C.t3, paddingBottom:4 }}>
                    {compras.filter(c => (!compraFiltDesde||c.f>=compraFiltDesde) && (!compraFiltHasta||c.f<=compraFiltHasta)).length} compras ·
                    Total: <strong style={{color:C.rd}}>{fmt(compras.filter(c => (!compraFiltDesde||c.f>=compraFiltDesde) && (!compraFiltHasta||c.f<=compraFiltHasta)).reduce((a,c)=>a+c.total,0))}</strong>
                  </div>
                </div>
              </Card>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:16 }}>
                <KPI icon="🏭" label="Compras este mes" value={compras.filter(c=>c.f.startsWith("2025-05")).length} color={C.bl}/>
                <KPI icon="💸" label="Total invertido"  value={fmt(compras.reduce((a,c)=>a+c.total,0))} color={C.rd}/>
                <KPI icon="📦" label="Productos reabastecidos" value={new Set(compras.map(c=>c.pid)).size} color={C.grL}/>
              </div>
              {compraForm !== null && (
                <Modal onClose={() => setCompraForm(null)}>
                  <ModalTitle>🏭 Registrar Compra</ModalTitle>
                  <CompraFormInner form={compraForm} setForm={setCompraForm} prods={prods} onSave={registrarCompra} onCancel={() => setCompraForm(null)}/>
                </Modal>
              )}
              <Card>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr><TH>Fecha</TH><TH>Proveedor</TH><TH>Producto</TH><TH>Color</TH><TH right>Cant.</TH><TH right>C. Unit.</TH><TH right>Total</TH><TH>Nota</TH></tr></thead>
                  <tbody>
                    {[...compras].filter(c => (!compraFiltDesde||c.f>=compraFiltDesde) && (!compraFiltHasta||c.f<=compraFiltHasta)).sort((a,b) => b.f.localeCompare(a.f)).map(cm => (
                      <tr key={cm.id}>
                        <TD>{cm.f}</TD>
                        <TD bold color={C.t1}>{cm.prov}</TD>
                        <TD>{cm.prod}</TD>
                        <TD sm color={C.t3}>{cm.col||"Varios"}</TD>
                        <TD right bold color={C.grL}>+{cm.qty}</TD>
                        <TD right>{fmt(cm.cu)}</TD>
                        <TD right bold color={C.rd}>{fmt(cm.total)}</TD>
                        <TD sm color={C.t3}>{cm.nota||"—"}</TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
            </>);
          })()}

        </div>
      </div>

      {/* ═══════ MODAL CLIENTE ═══════ */}
      {modal === "cliente" && (
        <Modal onClose={() => setModal(null)}>
          <ModalTitle>👤 Nuevo Cliente</ModalTitle>
          <ClienteForm onSave={c => { const nc = { id:"C"+uid(), ...c }; setClientes(cs => { sbSave("clientes", nc); return [nc, ...cs]; }); showToast("✓ Cliente registrado"); setModal(null); }} onCancel={() => setModal(null)}/>
        </Modal>
      )}


      {modal === "venta" && (
        <Modal onClose={() => setModal(null)}>
          <ModalTitle>💰 Nueva Venta</ModalTitle>
          <div style={{ textAlign:"center", color:C.t3, padding:"16px 0" }}>
            <div style={{ fontSize:32, marginBottom:8 }}>🛒</div>
            Para una venta completa con escáner, colores y pagos mixtos usa el <strong style={{color:C.ac}}>Punto de Venta</strong>.
            <div style={{ marginTop:16 }}><Btn variant="primary" onClick={() => { setModal(null); setMod("pos"); }}>Ir al Punto de Venta →</Btn></div>
          </div>
        </Modal>
      )}

      {/* ═══════ MODAL GASTO ═══════ */}
      {modal === "gasto" && (
        <Modal onClose={() => setModal(null)}>
          <ModalTitle>📋 Registrar Gasto</ModalTitle>
          <GastoForm onSave={g => { const ng = { id:"G"+uid(), ...g, monto:parseFloat(g.monto) }; setGastos(gs => { sbSave("gastos", ng); return [ng, ...gs]; }); showToast("✓ Gasto registrado"); setModal(null); }} onCancel={() => setModal(null)} usuarios={usuarios} />
        </Modal>
      )}

      {/* ═══════ MODAL LOTE ═══════ */}
      {modal === "lote" && (
        <Modal onClose={() => setModal(null)}>
          <ModalTitle>🔨 Nuevo Lote de Producción</ModalTitle>
          <LoteForm prods={prods} onSave={l => { const nl = { id:"L"+uid(), ...l, total: l.qty * (prods.find(p=>p.id===l.pid)?.c||0), estado:"En proceso" }; setLotes(ls => { sbSave("lotes", nl); return [nl, ...ls]; }); showToast("✓ Lote creado"); setModal(null); }} onCancel={() => setModal(null)} />
        </Modal>
      )}

      {/* ═══════ VOUCHER ═══════ */}
      {voucher && (
        <div onClick={() => setVoucher(null)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:16 }}>
          <div onClick={e => e.stopPropagation()} style={{ background:C.white, borderRadius:14, width:320, overflow:"hidden", boxShadow:"0 20px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ padding:"14px 16px", display:"flex", justifyContent:"space-between", alignItems:"center", borderBottom:`1px solid ${C.border}` }}>
              <div>
                <div style={{ fontSize:14, fontWeight:800, color:C.grL }}>✅ Venta registrada</div>
                <div style={{ fontSize:11, color:C.t3 }}>{voucher.num} · {voucher.vend}</div>
              </div>
              <button onClick={() => setVoucher(null)} style={{ background:"none", border:`1px solid ${C.border}`, borderRadius:7, padding:"4px 10px", color:C.t3, cursor:"pointer", fontSize:12 }}>✕</button>
            </div>
            <div id="voucher-print" style={{ background:"#fff", color:"#111", padding:"18px 16px", fontFamily:"'Courier New',monospace", fontSize:12 }}>
              <div style={{ textAlign:"center", marginBottom:6 }}>
                <img src={LOGO_MED} alt="MoblaMel" style={{ width:56, height:56, objectFit:"contain" }}/>
              </div>
              <div style={{ textAlign:"center", fontWeight:700, fontSize:13, marginBottom:2 }}>MOBLAMEL</div>
              <div style={{ textAlign:"center", fontSize:10, color:"#555", marginBottom:8 }}>RUC: 10402654703 · Villa El Salvador, Lima</div>
              <div style={{ borderTop:"1px dashed #ccc", margin:"6px 0" }}/>
              <div style={{ textAlign:"center", fontWeight:700, marginBottom:2 }}>{voucher.comp === "Nota de Venta" ? "NOTA DE VENTA" : voucher.comp.toUpperCase()}</div>
              <div style={{ textAlign:"center", fontSize:10, color:"#555", marginBottom:8 }}>{voucher.num} · {voucher.fecha}</div>
              {voucher.cliente !== "Consumidor final" && <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}><span>Cliente:</span><span style={{fontWeight:700}}>{voucher.cliente}</span></div>}
              <div style={{ borderTop:"1px dashed #ccc", margin:"6px 0" }}/>
              {voucher.items.map((it, i) => {
                const precio = parseFloat(it.pr) || parseFloat(it.p) || 0;
                const qty = parseInt(it.q) || 1;
                return (
                  <div key={i} style={{ marginBottom:6 }}>
                    <div style={{ fontWeight:700 }}>{it.n}</div>
                    <div style={{ color:"#777", fontSize:10 }}>Color: {it.col}</div>
                    <div style={{ display:"flex", justifyContent:"space-between" }}>
                      <span>{qty} × {fmt(precio)}</span>
                      <span style={{fontWeight:700}}>{fmt(qty*precio)}</span>
                    </div>
                  </div>
                );
              })}
              <div style={{ borderTop:"1px dashed #ccc", margin:"6px 0" }}/>
              {voucher.desc > 0 && <div style={{ display:"flex", justifyContent:"space-between", color:"#888", marginBottom:2 }}><span>Descuento:</span><span>-{fmt(voucher.desc)}</span></div>}
              <div style={{ display:"flex", justifyContent:"space-between", fontSize:16, fontWeight:800, marginTop:4 }}><span>TOTAL</span><span>{fmt(voucher.total)}</span></div>
              <div style={{ borderTop:"1px dashed #ccc", margin:"6px 0" }}/>
              {voucher.pagos.map((p, i) => (
                <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:11, marginBottom:2 }}><span>{p.met}</span><span>{fmt(parseFloat(p.monto)||0)}</span></div>
              ))}
              {voucher.vuelto > 0.01 && <div style={{ display:"flex", justifyContent:"space-between", color:"#276749", fontWeight:700, marginTop:3 }}><span>Vuelto:</span><span>{fmt(voucher.vuelto)}</span></div>}
              {voucher.comp === "Nota de Venta" && <div style={{ textAlign:"center", fontSize:10, color:"#888", marginTop:6, fontStyle:"italic" }}>Documento de uso interno · Sin valor fiscal</div>}
              <div style={{ textAlign:"center", fontSize:10, color:"#888", marginTop:8, borderTop:"1px dashed #ccc", paddingTop:6 }}>¡Gracias por su compra! · MoblaMel</div>
            </div>
            <div style={{ padding:"10px 14px", display:"flex", gap:8, flexDirection:"column" }}>
              <div style={{ display:"flex", gap:6, marginBottom:2 }}>
                {["ticket","a4"].map(f => (
                  <button key={f} onClick={() => setVoucherFmt(f)}
                    style={{ flex:1, padding:"6px", borderRadius:6, border:`1px solid ${voucherFmt===f?C.ac:C.border}`, background:voucherFmt===f?C.acBg:"transparent", color:voucherFmt===f?C.ac:C.t3, fontSize:11, fontWeight:700, cursor:"pointer" }}>
                    {f==="ticket"?"🧾 Ticket (80mm)":"📄 A4"}
                  </button>
                ))}
              </div>
              <div style={{ display:"flex", gap:8 }}>
              {/* IMPRIMIR — método que funciona en todos los navegadores */}
              <button onClick={() => {
                const el = document.getElementById('voucher-print');
                if (!el) return;
                const contenido = el.innerHTML;
                const ventana = window.open('', '_blank');
                if (!ventana) { showToast("Permite ventanas emergentes en tu navegador","err"); return; }
                const isA4 = voucherFmt === "a4";
                ventana.document.write(`<!DOCTYPE html>
<html><head>
  <meta charset="utf-8">
  <title>Comprobante ${voucher.num}</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family: ${isA4 ? "'Arial',sans-serif" : "'Courier New',monospace"}; font-size: ${isA4?"14px":"12px"}; padding: ${isA4?"40px":"10px"}; ${isA4?"max-width:210mm;margin:0 auto;":""} }
    img { display:block; margin:0 auto; }
    @media print { body { padding: ${isA4?"20mm 20mm":"0"}; } @page { margin: ${isA4?"10mm":"3mm"}; size: ${isA4?"A4":"80mm auto"}; } }
  </style>
</head><body>${contenido}
<script>window.onload=function(){window.print();setTimeout(function(){window.close();},500);}<\/script>
</body></html>`);
                ventana.document.close();
              }} style={{ flex:1, padding:"12px", borderRadius:8, border:"none", background:C.ac, color:"#fff", fontWeight:700, cursor:"pointer", fontSize:13, fontFamily:"inherit" }}>🖨️ Imprimir</button>

              {/* WHATSAPP — abre con mensaje de texto del voucher */}
              <button onClick={() => {
                const v = voucher;
                const items = (v.items||[]).map(i => `• ${i.n} (${i.col||''}) x${i.q} = S/${(i.q*(i.p||i.pr||0)).toFixed(2)}`).join('\n');
                const pagos = (v.pagos||[]).map(p => `  ${p.met}: S/${parseFloat(p.monto||0).toFixed(2)}`).join('\n');
                const msg =
`🧾 *MOBLAMEL*
RUC: 10402654703 · VES, Lima

📄 *${v.comp?.toUpperCase()||'NOTA DE VENTA'}*
N°: ${v.num}
Fecha: ${v.fecha||HOY}
Cliente: ${v.cliente||'Consumidor final'}
Vendedor: ${v.vend||''}

${items}
${(v.desc||0)>0?`Descuento: -S/${parseFloat(v.desc||0).toFixed(2)}\n`:''}
*TOTAL: S/${parseFloat(v.total||0).toFixed(2)}*

Pago:
${pagos}

✅ ¡Gracias por su compra!
_MoblaMel · Muebles en Melamina_`;
                window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
              }} style={{ flex:1, padding:"12px", borderRadius:8, border:"1px solid #25D366", background:"#25D366", color:"#fff", fontWeight:700, cursor:"pointer", fontSize:13, fontFamily:"inherit" }}>📲 WhatsApp</button>
            </div>
            </div>
          </div>
        </div>
      )}

      {/* TOAST */}
      {toast && (
        <div style={{ position:"fixed", bottom:20, left:"50%", transform:"translateX(-50%)", background: toast.t === "ok" ? C.grL : C.rdL, color:"#fff", padding:"10px 20px", borderRadius:99, fontSize:13, fontWeight:700, zIndex:9999, boxShadow:"0 4px 20px rgba(0,0,0,0.15)", whiteSpace:"nowrap" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// ─── FORMULARIOS EXTERNOS ─────────────────────────────────────────────────────
function ProdEditModal({ prod, onSave, onCancel }) {
  const [imgError, setImgError] = useState("");
  const [nuevoColor, setNuevoColor] = useState("");
  const [f, setF] = useState({
    p:    String(prod.p),
    c:    String(prod.c),
    min:  String(prod.min),
    stk:  String(prod.stk),
    ubi:  prod.ubi || UBICACIONES[0],
    tipo: prod.tipo,
    cat:  prod.cat || "Ropero",
    img:  prod.img || "",
    barcode: prod.barcode || "",
    cols: prod.cols || [],
  });
  const U = (k, v) => setF(x => ({ ...x, [k]: v }));
  const addColor = () => {
    const c = nuevoColor.trim();
    if (!c || f.cols.includes(c)) return;
    U("cols", [...f.cols, c]);
    setNuevoColor("");
  };
  const removeColor = (col) => {
    if (f.cols.length <= 1) return; // siempre al menos 1
    U("cols", f.cols.filter(c => c !== col));
  };
  const pV = parseFloat(f.p) || 0;
  const cV = parseFloat(f.c) || 0;
  const mg = pV > 0 ? Math.round(((pV - cV) / pV) * 100) : 0;
  const mgColor = mg >= 40 ? "#38a169" : mg >= 25 ? "#c05621" : "#c53030";

  const handleImg = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) { setImgError("La imagen debe ser menor a 2MB"); return; }
    const reader = new FileReader();
    reader.onload = ev => U("img", ev.target.result);
    reader.readAsDataURL(file);
  };

  return (
    <div onClick={e => { if (e.target === e.currentTarget) onCancel(); }}
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1100, padding:16 }}>
      <div style={{ background:"#fff", borderRadius:14, padding:24, width:520, maxWidth:"100%", maxHeight:"92vh", overflowY:"auto", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ fontSize:17, fontWeight:700, color:"#1a202c", marginBottom:4 }}>✏️ Editar Producto</div>
        <div style={{ fontSize:13, color:"#718096", marginBottom:18, paddingBottom:14, borderBottom:"1px solid #e2e6ed" }}>
          {prod.ico} {prod.n} <span style={{ fontSize:11, color:"#a0aec0" }}>· {prod.id}</span>
        </div>

        {/* Foto del producto */}
        <div style={{ marginBottom:16 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>
            📸 Foto del producto <span style={{ color:"#a0aec0", fontWeight:400, textTransform:"none" }}>(opcional · máx 2MB)</span>
          </label>
          <div style={{ display:"flex", gap:12, alignItems:"center" }}>
            <div style={{ width:90, height:90, borderRadius:10, background:"#f4f6f9", border:"2px dashed #e2e6ed", overflow:"hidden", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
              {f.img
                ? <img src={f.img} alt="preview" style={{ width:"100%", height:"100%", objectFit:"cover" }}/>
                : <span style={{ fontSize:32 }}>{prod.ico}</span>
              }
            </div>
            <div style={{ flex:1 }}>
              <label style={{ display:"block", padding:"9px 14px", borderRadius:8, border:"1px solid #e2e6ed", background:"#f4f6f9", color:"#4a5568", fontSize:13, fontWeight:600, cursor:"pointer", textAlign:"center", marginBottom:6 }}>
                📂 Seleccionar foto
                <input type="file" accept="image/*" onChange={handleImg} style={{ display:"none" }}/>
              </label>
              {f.img && (
                <button onClick={() => U("img", "")} style={{ width:"100%", padding:"6px", borderRadius:7, border:"1px solid #fed7d7", background:"#fff5f5", color:"#c53030", fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
                  🗑 Quitar foto
                </button>
              )}
              <div style={{ fontSize:10, color:"#a0aec0", marginTop:4 }}>JPG, PNG o WebP · La foto aparece en el POS</div>
            </div>
          </div>
          {imgError && <div style={{fontSize:11,color:"#c53030",marginTop:4,fontWeight:600}}>⚠️ {imgError}</div>}
        </div>

        {/* Categoría */}
        <div style={{ marginBottom:16 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Categoría</label>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {[["Ropero","🪞"],["Zapatero","👟"],["Cómoda","🗄️"],["Mesa de Noche","🛏️"],["Cabecera","🛋️"],["Otro","📦"]].map(([c,ico]) => (
              <button key={c} type="button" onClick={() => U("cat",c)}
                style={{ display:"flex", alignItems:"center", gap:5, padding:"6px 12px", borderRadius:99, border:`2px solid ${f.cat===c?"#e07b39":"#e2e6ed"}`, background:f.cat===c?"#fef3ec":"transparent", cursor:"pointer", fontFamily:"inherit" }}>
                <span style={{ fontSize:14 }}>{ico}</span>
                <span style={{ fontSize:12, fontWeight:700, color:f.cat===c?"#c4622a":"#718096" }}>{c}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Colores */}
        <div style={{ marginBottom:16 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>
            Colores disponibles
          </label>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:8 }}>
            {f.cols.map(col => (
              <div key={col} style={{ display:"flex", alignItems:"center", gap:5, padding:"5px 10px", borderRadius:99, background:"#f4f6f9", border:"1px solid #e2e6ed" }}>
                <div style={{ width:9, height:9, borderRadius:"50%", background:HEX_COLOR[col]||"#888", flexShrink:0 }}/>
                <span style={{ fontSize:12, fontWeight:600, color:"#1a202c" }}>{col}</span>
                {f.cols.length > 1 && (
                  <button onClick={() => removeColor(col)}
                    style={{ background:"none", border:"none", color:"#e53e3e", cursor:"pointer", fontSize:13, padding:"0 0 0 2px", lineHeight:1 }}>✕</button>
                )}
              </div>
            ))}
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <input value={nuevoColor} onChange={e => setNuevoColor(e.target.value)}
              onKeyDown={e => e.key === "Enter" && addColor()}
              placeholder="Ej: Cerezo, Caoba, Gris..."
              style={{ flex:1, background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"7px 10px", fontSize:13, outline:"none", fontFamily:"inherit", color:"#1a202c" }}/>
            <button onClick={addColor} disabled={!nuevoColor.trim() || f.cols.includes(nuevoColor.trim())}
              style={{ padding:"7px 14px", borderRadius:8, border:"none", background:nuevoColor.trim()&&!f.cols.includes(nuevoColor.trim())?"#e07b39":"#e2e6ed", color:nuevoColor.trim()&&!f.cols.includes(nuevoColor.trim())?"#fff":"#a0aec0", fontSize:12, fontWeight:700, cursor:"pointer" }}>
              + Agregar
            </button>
          </div>
          <div style={{ fontSize:10, color:"#a0aec0", marginTop:4 }}>Presiona Enter o toca "+ Agregar" · Mínimo 1 color requerido</div>
        </div>

        {/* Código de barras */}
        <div style={{ marginBottom:16 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
            Código de barras <span style={{ color:"#a0aec0", fontWeight:400, textTransform:"none" }}>(escanea o escribe)</span>
          </label>
          <input value={f.barcode} onChange={e => U("barcode", e.target.value)}
            placeholder="Ej: 7501234567890"
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"9px 12px", fontSize:14, fontWeight:600, outline:"none", boxSizing:"border-box", fontFamily:"monospace", letterSpacing:"1px", color:"#1a202c" }}/>
          <div style={{ fontSize:10, color:"#a0aec0", marginTop:3 }}>Conecta tu lector USB/Bluetooth y escanea el código — se escribe automáticamente</div>
        </div>

        {/* Precios */}
        <div style={{ background:"#fef3ec", border:"1px solid #f5a56e44", borderRadius:10, padding:"14px 16px", marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#c4622a", marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>
            💰 Precios
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:10 }}>
            <div>
              <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Precio de venta S/</label>
              <input type="number" min="0" step="0.5" value={f.p} onChange={e => U("p", e.target.value)}
                style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"9px 12px", fontSize:16, fontWeight:700, outline:"none", boxSizing:"border-box", color:"#1a202c", fontFamily:"inherit" }}/>
            </div>
            <div>
              <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
                Costo unitario S/ <span style={{ color:"#c05621" }}>← lo que pagaste</span>
              </label>
              <input type="number" min="0" step="0.5" value={f.c} onChange={e => U("c", e.target.value)}
                style={{ width:"100%", background:"#fff", border:"2px solid #c05621", borderRadius:8, padding:"9px 12px", fontSize:16, fontWeight:700, outline:"none", boxSizing:"border-box", color:"#c05621", fontFamily:"inherit" }}/>
            </div>
          </div>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", background:"#fff", borderRadius:8, padding:"8px 12px" }}>
            <span style={{ fontSize:13, color:"#718096" }}>Margen resultante</span>
            <span style={{ fontSize:20, fontWeight:800, color:mgColor }}>{mg}%</span>
          </div>
          {cV > pV && <div style={{ marginTop:8, fontSize:11, color:"#c53030", fontWeight:600 }}>⚠️ El costo supera el precio — vendiendo a pérdida</div>}
        </div>

        {/* Stock */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14 }}>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Stock actual</label>
            <input type="number" min="0" value={f.stk} onChange={e => U("stk", e.target.value)}
              style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 12px", fontSize:14, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
          </div>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Stock mínimo (alerta)</label>
            <input type="number" min="0" value={f.min} onChange={e => U("min", e.target.value)}
              style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 12px", fontSize:14, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
          </div>
        </div>

        {/* Ubicación y tipo */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14 }}>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Ubicación principal</label>
            <select value={f.ubi} onChange={e => U("ubi", e.target.value)}
              style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
              {UBICACIONES.map(u => <option key={u}>{u}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Tipo</label>
            <select value={f.tipo} onChange={e => U("tipo", e.target.value)}
              style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
              <option>Fabricado</option>
              <option>Comprado</option>
            </select>
          </div>
        </div>
        {/* Segunda ubicación (cuando hay stock en dos lugares) */}
        <div style={{ background:C.bg, borderRadius:8, padding:"10px 12px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:600, color:C.t3, marginBottom:8 }}>📍 ¿También hay stock en otra ubicación?</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 80px", gap:10 }}>
            <select value={f.ubi2||""} onChange={e => U("ubi2", e.target.value)}
              style={{ background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:12, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
              <option value="">— Sin segunda ubicación —</option>
              {UBICACIONES.filter(u => u !== f.ubi).map(u => <option key={u}>{u}</option>)}
            </select>
            <input type="number" min="0" value={f.stk2||0} onChange={e => U("stk2", parseInt(e.target.value)||0)}
              placeholder="Stock" disabled={!f.ubi2}
              style={{ background: f.ubi2?"#fff":"#f4f6f9", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
          </div>
        </div>

        <div style={{ display:"flex", gap:8, justifyContent:"flex-end", borderTop:"1px solid #e2e6ed", paddingTop:16 }}>
          <button onClick={onCancel}
            style={{ padding:"9px 18px", borderRadius:8, border:"1px solid #e2e6ed", background:"#fff", color:"#718096", fontSize:13, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>
            Cancelar
          </button>
          <button onClick={() => onSave({ ...prod, p:pV, c:cV, min:parseInt(f.min)||0, stk:parseInt(f.stk)||0, ubi:f.ubi, ubi2:f.ubi2||"", stk2:parseInt(f.stk2)||0, tipo:f.tipo, cat:f.cat||prod.cat, cols:f.cols.length?f.cols:prod.cols, img:f.img, barcode:f.barcode?.trim()||"" })}
            style={{ padding:"9px 20px", borderRadius:8, border:"none", background:"#e07b39", color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
            ✓ Guardar cambios
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── FORMULARIO PEDIDO ───────────────────────────────────────────────────────
function PedidoFormModal({ form, setForm, prods, usuarios, onSave, onCancel }) {
  const [saveErr, setSaveErr] = useState("");
  const F = (k,v) => setForm(f => ({ ...f, [k]:v }));
  const prodSel = prods.find(p => p.id === form.pid);
  const saldo = (parseFloat(form.precioAcordado)||0) - (parseFloat(form.adelanto)||0);
  const vendedores = usuarios.filter(u => u.rol !== "taller" && u.activo);

  // Al cambiar producto en separación, auto-rellenar nombre y precio
  const selProd = (pid) => {
    const p = prods.find(x => x.id === pid);
    F("pid", pid);
    if (p) { setForm(f => ({ ...f, pid, prod:p.n, col:p.cols[0], precioAcordado:p.p })); }
  };

  return (
    <div onClick={e => { if (e.target === e.currentTarget) onCancel(); }}
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1100, padding:16 }}>
      <div style={{ background:"#fff", borderRadius:14, padding:24, width:560, maxWidth:"100%", maxHeight:"93vh", overflowY:"auto", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>

        {/* Tipo */}
        <div style={{ fontSize:17, fontWeight:700, color:"#1a202c", marginBottom:16, paddingBottom:14, borderBottom:"1px solid #e2e6ed" }}>
          📬 Nuevo Pedido
        </div>
        <div style={{ display:"flex", gap:8, marginBottom:18 }}>
          {[["separacion","📦 Separación de stock","El mueble ya existe en inventario"],
            ["medida","🔨 Pedido a medida","Se fabrica según especificación del cliente"]
          ].map(([val,label,desc]) => (
            <button key={val} onClick={() => setForm(f => ({ ...f, tipo:val, pid:"", prod:"" }))}
              style={{ flex:1, padding:"10px 8px", borderRadius:10, border:`2px solid ${form.tipo===val?"#e07b39":"#e2e6ed"}`, background:form.tipo===val?"#fef3ec":"transparent", cursor:"pointer", textAlign:"left" }}>
              <div style={{ fontSize:13, fontWeight:700, color:form.tipo===val?"#c4622a":"#4a5568" }}>{label}</div>
              <div style={{ fontSize:11, color:"#a0aec0", marginTop:2 }}>{desc}</div>
            </button>
          ))}
        </div>

        {/* Cliente */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
          <Inp label="Nombre del cliente *" value={form.cli} onChange={e => F("cli",e.target.value)} placeholder="Ej: Juan Quispe"/>
          <Inp label="Celular" value={form.cel||""} onChange={e => F("cel",e.target.value.replace(/\D/g,"").slice(0,9))} placeholder="987654321"/>
        </div>

        {/* Vendedor */}
        <div style={{ marginBottom:14 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Vendedor</label>
          <select value={form.vend} onChange={e => F("vend",e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
            {vendedores.map(u => <option key={u.id}>{u.nombre}</option>)}
          </select>
        </div>

        {/* Producto */}
        {form.tipo === "separacion" ? (
          <div style={{ marginBottom:12 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Producto del inventario *</label>
            <select value={form.pid} onChange={e => selProd(e.target.value)}
              style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"9px 12px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box", marginBottom:8 }}>
              <option value="">— Seleccionar producto —</option>
              {prods.filter(p => p.stk > (p.stkRes||0)).map(p => <option key={p.id} value={p.id}>{p.ico} {p.n} — Stock disponible: {p.stk-(p.stkRes||0)}</option>)}
            </select>
            {prodSel && (
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                <div>
                  <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Color</label>
                  <select value={form.col} onChange={e => F("col",e.target.value)}
                    style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
                    {prodSel.cols.map(c => <option key={c}>{c}</option>)}
                  </select>
                </div>
                <Inp label="Cantidad" type="number" min="1" max={prodSel.stk-(prodSel.stkRes||0)} value={form.qty} onChange={e => F("qty",parseInt(e.target.value)||1)}/>
              </div>
            )}
          </div>
        ) : (
          <div style={{ marginBottom:12 }}>
            <Inp label="Descripción del pedido *" value={form.prod} onChange={e => F("prod",e.target.value)} placeholder="Ej: Ropero 3P color Cerezo oscuro, 2.20m alto"/>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:10 }}>
              <Inp label="Color / Acabado" value={form.col||""} onChange={e => F("col",e.target.value)} placeholder="Ej: Cerezo oscuro"/>
              <Inp label="Cantidad" type="number" min="1" value={form.qty} onChange={e => F("qty",parseInt(e.target.value)||1)}/>
            </div>
          </div>
        )}

        {/* Nota */}
        <div style={{ marginBottom:14 }}>
          <Inp label="Nota / especificaciones" value={form.nota||""} onChange={e => F("nota",e.target.value)} placeholder="Medidas, detalles especiales, observaciones..."/>
        </div>

        {/* Precios y adelanto */}
        <div style={{ background:"#fef3ec", borderRadius:10, padding:"14px 16px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#c4622a", marginBottom:12, textTransform:"uppercase", letterSpacing:"0.5px" }}>💰 Precio y adelanto</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:10 }}>
            <div>
              <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
                Precio acordado S/ {form.tipo==="medida"&&<span style={{color:"#c05621",fontWeight:400,textTransform:"none"}}>(especial)</span>}
              </label>
              <input
                inputMode="decimal"
                value={form.precioAcordado === 0 ? "" : String(form.precioAcordado)}
                onChange={e => { const v = e.target.value.replace(/[^0-9.]/g,""); F("precioAcordado", v === "" ? 0 : parseFloat(v) || 0); }}
                placeholder="S/ 0.00"
                style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"9px 12px", fontSize:18, fontWeight:800, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
              {form.tipo==="medida" && <div style={{ fontSize:10, color:"#a0aec0", marginTop:3 }}>Precio distinto al catálogo — se fabrica por unidad</div>}
            </div>
            <div>
              <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Adelanto S/ <span style={{color:"#c05621",fontWeight:400,textTransform:"none"}}>(entra a caja hoy)</span></label>
              <input
                inputMode="decimal"
                value={form.adelanto === 0 ? "" : String(form.adelanto)}
                onChange={e => { const v = e.target.value.replace(/[^0-9.]/g,""); F("adelanto", v === "" ? 0 : parseFloat(v) || 0); }}
                placeholder="S/ 0.00"
                style={{ width:"100%", background:"#fff", border:"2px solid #38a169", borderRadius:8, padding:"9px 12px", fontSize:18, fontWeight:800, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#276749" }}/>
            </div>
          </div>
          {/* Método de pago del adelanto */}
          {(form.adelanto||0) > 0 && (
            <div style={{ marginBottom:10 }}>
              <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Método de pago del adelanto</label>
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                {["Efectivo","Yape/Plin","Trans. BCP","POS Tarjeta"].map(mp => (
                  <button key={mp} type="button" onClick={() => F("mpAdelanto", mp)}
                    style={{ padding:"6px 12px", borderRadius:99, border:`2px solid ${(form.mpAdelanto||"Efectivo")===mp?"#e07b39":"#e2e6ed"}`, background:(form.mpAdelanto||"Efectivo")===mp?"#fef3ec":"transparent", color:(form.mpAdelanto||"Efectivo")===mp?"#c4622a":"#718096", fontSize:12, fontWeight:700, cursor:"pointer" }}>
                    {{"Efectivo":"💵","Yape/Plin":"📱","Trans. BCP":"🏦","POS Tarjeta":"💳"}[mp]} {mp}
                  </button>
                ))}
              </div>
            </div>
          )}
          <div style={{ display:"flex", justifyContent:"space-between", background:"#fff", borderRadius:8, padding:"8px 12px", fontSize:13 }}>
            <span style={{ color:"#718096" }}>Saldo pendiente que quedará</span>
            <span style={{ fontWeight:800, color: saldo>0?"#c53030":"#38a169", fontSize:16 }}>S/ {saldo.toFixed(2)}</span>
          </div>
        </div>

        {/* Fecha entrega y alerta */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:20 }}>
          <Inp label="Fecha estimada de entrega" type="date" value={form.fEnt||""} onChange={e => F("fEnt",e.target.value)}/>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Alerta si no recoge en (días)</label>
            <input type="number" min="1" value={form.diasAlerta||30} onChange={e => F("diasAlerta",parseInt(e.target.value)||30)}
              style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, fontWeight:600, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
          </div>
        </div>

        {saveErr && <div style={{fontSize:11,color:"#c53030",padding:"6px 10px",background:"#fff5f5",borderRadius:7,marginBottom:8,fontWeight:600}}>⚠️ {saveErr}</div>}
        <div style={{ display:"flex", gap:8, justifyContent:"flex-end", borderTop:"1px solid #e2e6ed", paddingTop:16 }}>
          <Btn onClick={onCancel}>Cancelar</Btn>
          <Btn variant="primary" onClick={() => {
            setSaveErr("");
            if (!form.cli) { setSaveErr("El nombre del cliente es obligatorio"); return; }
            if (!form.prod && form.tipo==="medida") { setSaveErr("Describe el pedido"); return; }
            if (form.tipo==="separacion" && !form.pid) { setSaveErr("Selecciona el producto del inventario"); return; }
            if ((form.precioAcordado||0) <= 0) { setSaveErr("El precio acordado es obligatorio"); return; }
            onSave({ ...form, qty:parseInt(form.qty)||1, precioAcordado:parseFloat(form.precioAcordado)||0, adelanto:parseFloat(form.adelanto)||0,
              est: form.tipo==="separacion" ? "Pendiente entrega" : "En producción" });
          }}>✓ Registrar Pedido</Btn>
        </div>
      </div>
    </div>
  );
}

// ─── ESCÁNER DE CÁMARA (zxing-js — funciona en iPhone Safari) ────────────────
function BarcodeCamera({ onDetect, onClose }) {
  const videoRef = useRef(null);
  const [err, setErr] = useState("");
  const [hint, setHint] = useState("Iniciando cámara...");
  const streamRef = useRef(null);
  const readerRef = useRef(null);

  useEffect(() => {
    let active = true;
    const start = async () => {
      try {
        // Cargar zxing desde CDN
        if (!window.ZXingBrowser) {
          setHint("Cargando librería de escaneo...");
          await new Promise((res, rej) => {
            const s = document.createElement("script");
            s.src = "https://cdnjs.cloudflare.com/ajax/libs/zxing-browser/0.0.10/index.min.js";
            s.onload = res; s.onerror = rej;
            document.head.appendChild(s);
          });
        }
        if (!active) return;
        const stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: "environment", width: { ideal: 1280 } }
        });
        if (!active) { stream.getTracks().forEach(t => t.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) videoRef.current.srcObject = stream;
        setHint("Apunta el código de barras hacia la cámara");

        // Intentar con BarcodeDetector nativo primero (Android/Chrome)
        if ("BarcodeDetector" in window) {
          const det = new window.BarcodeDetector({ formats:["ean_13","ean_8","code_128","code_39","qr_code","upc_a"] });
          const scan = async () => {
            if (!active) return;
            try {
              if (videoRef.current?.readyState >= 2) {
                const barcodes = await det.detect(videoRef.current);
                if (barcodes.length > 0) { stop(); onDetect(barcodes[0].rawValue); return; }
              }
            } catch(_){}
            requestAnimationFrame(scan);
          };
          requestAnimationFrame(scan);
          return;
        }

        // Fallback: zxing-js (iPhone Safari + todos los demás)
        if (window.ZXingBrowser) {
          const reader = new window.ZXingBrowser.BrowserMultiFormatReader();
          readerRef.current = reader;
          reader.decodeFromVideoElement(videoRef.current, (result, err) => {
            if (!active) return;
            if (result) { stop(); onDetect(result.getText()); }
          });
        }
      } catch(e) {
        if (e.name === "NotAllowedError") setErr("Permiso de cámara denegado. Ve a Ajustes → Safari → Cámara y actívalo.");
        else setErr("No se pudo abrir la cámara: " + e.message);
      }
    };
    const stop = () => {
      active = false;
      readerRef.current?.reset?.();
      streamRef.current?.getTracks().forEach(t => t.stop());
    };
    start();
    return stop;
  }, []);

  return (
    <div style={{ marginTop:10, borderRadius:10, overflow:"hidden", background:"#000", position:"relative" }}>
      {err ? (
        <div style={{ padding:"20px", textAlign:"center", color:"#fff", fontSize:13 }}>
          <div style={{ fontSize:24, marginBottom:8 }}>📵</div>
          <div style={{ marginBottom:10 }}>{err}</div>
          <button onClick={onClose} style={{ padding:"7px 14px", borderRadius:8, border:"none", background:"#e07b39", color:"#fff", fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>Cerrar</button>
        </div>
      ) : (<>
        <video ref={videoRef} autoPlay playsInline muted style={{ width:"100%", maxHeight:220, objectFit:"cover", display:"block" }}/>
        <div style={{ position:"absolute", inset:0, display:"flex", alignItems:"center", justifyContent:"center", pointerEvents:"none" }}>
          <div style={{ width:"65%", height:"35%", border:"2px solid #e07b39", borderRadius:8, boxShadow:"0 0 0 9999px rgba(0,0,0,0.4)" }}/>
        </div>
        <div style={{ position:"absolute", bottom:0, left:0, right:0, background:"rgba(0,0,0,0.6)", color:"#fff", padding:"6px 10px", fontSize:11, textAlign:"center" }}>{hint}</div>
        <button onClick={onClose} style={{ position:"absolute", top:8, right:8, padding:"4px 10px", borderRadius:7, border:"none", background:"rgba(0,0,0,0.6)", color:"#fff", fontSize:12, cursor:"pointer" }}>✕</button>
      </>)}
    </div>
  );
}

function NuevoProdForm({ form, cats, onSave, onCancel }) {
  const [saveErr, setSaveErr] = useState("");
  const [imgError, setImgError] = useState("");
  const ICOS = ["🪞","👟","🗄️","🛏️","🛋️","🪑","🚪","📦","🖼️","🛠️"];
  const [f, setF] = useState({ ...form, colsStr: form.cols?.join(", ") || "Blanco", barcode: form.barcode || "" });
  const U = (k, v) => setF(x => ({ ...x, [k]: v }));
  const pV = parseFloat(f.p) || 0, cV = parseFloat(f.c) || 0;
  const mg = pV > 0 ? Math.round(((pV-cV)/pV)*100) : 0;
  const mgColor = mg >= 40 ? C.grL : mg >= 25 ? C.or : C.rd;

  const handleImg = (e) => {
    const file = e.target.files[0]; if (!file) return;
    if (file.size > 2*1024*1024) { setImgError("Imagen menor a 2MB"); return; }
    const reader = new FileReader();
    reader.onload = ev => U("img", ev.target.result);
    reader.readAsDataURL(file);
  };

  const handleSave = () => {
    if (!f.n) { setSaveErr("El nombre es obligatorio"); return; }
    const cols = f.colsStr.split(",").map(c => c.trim()).filter(Boolean);
    if (!cols.length) { setSaveErr("Agrega al menos un color"); return; }
    onSave({ ...f, p:pV, c:cV, min:parseInt(f.min)||1, stk:parseInt(f.stk)||0, cols,
      id: f.id ? f.id.toUpperCase() : "", ubi:f.ubi||UBICACIONES[0], barcode:f.barcode?.trim()||"" });
  };

  return (
    <div>
      {/* Código + nombre */}
      <div style={{ display:"grid", gridTemplateColumns:"120px 1fr", gap:12, marginBottom:14 }}>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
            Código <span style={{ color:"#a0aec0", fontWeight:400 }}>(opcional)</span>
          </label>
          <input value={f.id} onChange={e => U("id", e.target.value.toUpperCase())} placeholder="Ej: P011"
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"monospace", color:"#1a202c" }}/>
          <div style={{ fontSize:10, color:"#a0aec0", marginTop:3 }}>Se genera solo si está vacío</div>
        </div>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Nombre del producto *</label>
          <input value={f.n} onChange={e => U("n", e.target.value)} placeholder="Ej: Ropero 5 Puertas con Luna"
            style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"9px 10px", fontSize:13, fontWeight:600, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
        </div>
      </div>

      {/* Foto + ícono */}
      <div style={{ display:"flex", gap:14, marginBottom:14, alignItems:"flex-start" }}>
        <div style={{ flex:1 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>📸 Foto</label>
          <div style={{ display:"flex", gap:10, alignItems:"center" }}>
            <div style={{ width:72, height:72, borderRadius:10, background:"#f4f6f9", border:"2px dashed #e2e6ed", overflow:"hidden", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
              {f.img ? <img src={f.img} alt="" style={{ width:"100%", height:"100%", objectFit:"cover" }}/> : <span style={{ fontSize:28 }}>{f.ico}</span>}
            </div>
            <div>
              <label style={{ display:"block", padding:"7px 12px", borderRadius:8, border:"1px solid #e2e6ed", background:"#f4f6f9", color:"#4a5568", fontSize:12, fontWeight:600, cursor:"pointer", marginBottom:5 }}>
                📂 Subir foto
                <input type="file" accept="image/*" onChange={handleImg} style={{ display:"none" }}/>
              </label>
              {f.img && <button onClick={() => U("img","")} style={{ width:"100%", padding:"5px", borderRadius:7, border:"1px solid #fed7d7", background:"#fff5f5", color:"#c53030", fontSize:11, cursor:"pointer", fontFamily:"inherit" }}>🗑 Quitar</button>}
            </div>
          </div>
        </div>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Ícono</label>
          <div style={{ display:"flex", gap:5, flexWrap:"wrap", maxWidth:180 }}>
            {ICOS.map(ic => (
              <button key={ic} onClick={() => U("ico",ic)} style={{ width:34, height:34, borderRadius:8, border:`2px solid ${f.ico===ic?"#e07b39":"#e2e6ed"}`, background:f.ico===ic?"#fef3ec":"transparent", fontSize:18, cursor:"pointer" }}>{ic}</button>
            ))}
          </div>
        </div>
      </div>

      {/* Categoría, tipo, ubicación */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10, marginBottom:14 }}>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Categoría *</label>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:6, marginBottom:4 }}>
            {[["Ropero","🪞"],["Zapatero","👟"],["Cómoda","🗄️"],["Mesa de Noche","🛏️"],["Cabecera","🛋️"],["Otro","📦"]].map(([c,ico]) => (
              <button key={c} type="button" onClick={() => U("cat",c)}
                style={{ padding:"8px 4px", borderRadius:9, border:`2px solid ${f.cat===c?"#e07b39":"#e2e6ed"}`, background:f.cat===c?"#fef3ec":"transparent", cursor:"pointer", textAlign:"center", fontFamily:"inherit" }}>
                <div style={{ fontSize:20, marginBottom:2 }}>{ico}</div>
                <div style={{ fontSize:10, fontWeight:700, color:f.cat===c?"#c4622a":"#718096" }}>{c}</div>
              </button>
            ))}
          </div>
        </div>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Tipo</label>
          <select value={f.tipo} onChange={e => U("tipo",e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
            <option>Fabricado</option><option>Comprado</option>
          </select>
        </div>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Ubicación</label>
          <select value={f.ubi} onChange={e => U("ubi",e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", boxSizing:"border-box" }}>
            {UBICACIONES.map(u => <option key={u}>{u}</option>)}
          </select>
        </div>
      </div>

      {/* Código de barras */}
      <div style={{ marginBottom:14 }}>
        <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
          Código de barras <span style={{ color:"#a0aec0", fontWeight:400, textTransform:"none" }}>(opcional · escanea o escribe)</span>
        </label>
        <input value={f.barcode||""} onChange={e => U("barcode",e.target.value)}
          placeholder="Ej: 7501234567890 — conecta tu lector y escanea"
          style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 12px", fontSize:13, fontWeight:600, outline:"none", boxSizing:"border-box", fontFamily:"monospace", letterSpacing:"1px", color:"#1a202c" }}/>
      </div>

      {/* Colores */}
      <div style={{ marginBottom:14 }}>
        <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>
          Colores <span style={{ color:"#a0aec0", fontWeight:400, textTransform:"none" }}>(separados por coma)</span>
        </label>
        <input value={f.colsStr} onChange={e => U("colsStr",e.target.value)} placeholder="Blanco, Nogal, Wengué"
          style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 12px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
        <div style={{ display:"flex", gap:5, flexWrap:"wrap", marginTop:6 }}>
          {f.colsStr.split(",").map(c=>c.trim()).filter(Boolean).map((col,i) => (
            <span key={i} style={{ display:"flex", alignItems:"center", gap:4, padding:"3px 8px", borderRadius:99, background:C.bg, border:`1px solid ${C.border}`, fontSize:11 }}>
              <div style={{ width:8, height:8, borderRadius:"50%", background:HEX_COLOR[col]||"#888" }}/>
              {col}
            </span>
          ))}
        </div>
      </div>

      {/* Precios */}
      <div style={{ background:"#fef3ec", borderRadius:10, padding:"12px 14px", marginBottom:14 }}>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:8 }}>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Precio de venta S/</label>
            <input type="number" min="0" value={f.p} onChange={e => U("p",e.target.value)}
              style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"8px 10px", fontSize:15, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
          </div>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Costo unitario S/</label>
            <input type="number" min="0" value={f.c} onChange={e => U("c",e.target.value)}
              style={{ width:"100%", background:"#fff", border:"2px solid #c05621", borderRadius:8, padding:"8px 10px", fontSize:15, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#c05621" }}/>
          </div>
        </div>
        <div style={{ display:"flex", justifyContent:"space-between", background:"#fff", borderRadius:8, padding:"6px 10px", fontSize:13 }}>
          <span style={{ color:"#718096" }}>Margen</span>
          <span style={{ fontWeight:800, color:mgColor }}>{mg}%</span>
        </div>
      </div>

      {/* Stock */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:20 }}>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Stock inicial</label>
          <input type="number" min="0" value={f.stk} onChange={e => U("stk",e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:14, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
        </div>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Stock mínimo (alerta)</label>
          <input type="number" min="0" value={f.min} onChange={e => U("min",e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:14, fontWeight:700, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
        </div>
      </div>

      <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" onClick={handleSave}>✓ Registrar Producto</Btn>
      </div>
    </div>
  );
}

function GastoForm({ onSave, onCancel, usuarios }) {
  const CATS_FIJAS = ["Alquiler","Servicios taller","Mantenimiento"];
  const CATS_VARS  = ["Materiales","Armador por obra","Proveedor","Comisión","Otro"];
  const CAT_INFO = {
    "Alquiler":          { ico:"🏠", tipo:"fijo",     hint:"Alquiler mensual de local o taller" },
    "Servicios taller":  { ico:"⚡", tipo:"fijo",     hint:"Luz, agua del taller" },
    "Mantenimiento":     { ico:"🔧", tipo:"fijo",     hint:"Herramientas y máquinas" },
    "Materiales":        { ico:"🪵", tipo:"variable", hint:"Melamina, tableros, accesorios" },
    "Armador por obra":  { ico:"👷", tipo:"variable", hint:"Pago por trabajo específico" },
    "Proveedor":         { ico:"🏭", tipo:"variable", hint:"Muebles comprados a proveedor" },
    "Comisión":          { ico:"🏆", tipo:"variable", hint:"% comisión pagada a vendedor" },
    "Otro":              { ico:"📦", tipo:"variable", hint:"Otros gastos" },
  };
  const [form, setF] = useState({ f:new Date().toISOString().split("T")[0], cat:"Alquiler", desc:"", prov:"", monto:"", esFijo:true, vendedor:"" });
  const F = (k, v) => setF(f => ({ ...f, [k]: v }));
  const info = CAT_INFO[form.cat] || { ico:"📦", tipo:"variable", hint:"" };
  const vendedores = (usuarios||[]).filter(u => u.rol !== "taller" && u.activo);
  return (
    <div>
      <div style={{ background:"#eff6ff", borderRadius:8, padding:"8px 12px", marginBottom:10, fontSize:11, color:"#1e40af" }}>
        💡 <strong>Sueldos base de vendedores:</strong> se configuran en <strong>Usuarios → Editar vendedor</strong>. No se registran aquí.
      </div>
      {/* Fijo vs Variable */}
      <div style={{ display:"flex", gap:8, marginBottom:10 }}>
        {[["fijo","🔒 Gasto Fijo","Se paga siempre"],["variable","📈 Variable","Según operación"]].map(([tipo,label,desc]) => (
          <button key={tipo} type="button" onClick={() => { const c=tipo==="fijo"?"Alquiler":"Materiales"; setF(f=>({...f,cat:c,esFijo:tipo==="fijo"})); }}
            style={{ flex:1, padding:"10px 8px", borderRadius:10, border:`2px solid ${info.tipo===tipo?C.ac:C.border}`, background:info.tipo===tipo?C.acBg:"transparent", cursor:"pointer", textAlign:"left" }}>
            <div style={{ fontSize:13, fontWeight:700, color:info.tipo===tipo?C.ac:C.t2 }}>{label}</div>
            <div style={{ fontSize:10, color:C.t4, marginTop:2 }}>{desc}</div>
          </button>
        ))}
      </div>
      {/* Categorías */}
      <div style={{ marginBottom:14 }}>
        <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Categoría</label>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
          {(info.tipo==="fijo" ? CATS_FIJAS : CATS_VARS).map(c => {
            const ci = CAT_INFO[c];
            return (
              <button key={c} type="button" onClick={() => F("cat",c)}
                style={{ padding:"8px 10px", borderRadius:8, border:`2px solid ${form.cat===c?C.ac:C.border}`, background:form.cat===c?C.acBg:"transparent", cursor:"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontSize:16 }}>{ci.ico}</span>
                <div>
                  <div style={{ fontSize:12, fontWeight:700, color:form.cat===c?C.ac:C.t2 }}>{c}</div>
                  <div style={{ fontSize:9, color:C.t4 }}>{ci.hint}</div>
                </div>
              </button>
            );
          })}
        </div>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
        <Inp label="Fecha" type="date" value={form.f} onChange={e => F("f", e.target.value)} />
        <Inp label="Monto S/ *" type="number" inputMode="decimal" value={form.monto} onChange={e => F("monto", e.target.value)} placeholder="0.00"/>
      </div>
      {(form.cat==="Sueldo base"||form.cat==="Comisión") && (
        <div style={{ marginBottom:12 }}>
          <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>Vendedor</label>
          <select value={form.vendedor} onChange={e => F("vendedor",e.target.value)}
            style={{ width:"100%", background:"#fff", border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, boxSizing:"border-box" }}>
            <option value="">— Seleccionar —</option>
            {vendedores.map(u => <option key={u.id}>{u.nombre}</option>)}
          </select>
        </div>
      )}
      <div style={{ marginBottom:12 }}>
        <Inp label="Descripción *" value={form.desc} onChange={e => F("desc",e.target.value)}
          placeholder={form.cat==="Alquiler"?"Ej: Alquiler Tienda Principal · junio":form.cat==="Sueldo base"?"Ej: Sueldo base · semana 1":form.cat==="Materiales"?"Ej: Planchas melamina 18mm x20":"Descripción del gasto"}/>
      </div>
      {form.cat!=="Sueldo base"&&form.cat!=="Comisión" && (
        <div style={{ marginBottom:14 }}>
          <Inp label="Proveedor / Quién cobró" value={form.prov} onChange={e => F("prov",e.target.value)} placeholder="Ej: Maderco SAC"/>
        </div>
      )}
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" onClick={() => { if(!form.desc||!form.monto) return; onSave({...form,esFijo:CATS_FIJAS.includes(form.cat)}); }}>Registrar gasto</Btn>
      </div>
    </div>
  );
}


function LoteForm({ prods, onSave, onCancel }) {
  const [form, setF] = useState({ f: new Date().toISOString().split("T")[0], pid:"", qty:1 });
  const F = (k, v) => setF(f => ({ ...f, [k]: v }));
  const prod = prods.find(p => p.id === form.pid);
  return (
    <div>
      <div style={{ marginBottom:12 }}>
        <Sel label="Producto a fabricar" value={form.pid} onChange={e => F("pid", e.target.value)}>
          <option value="">— Seleccionar —</option>
          {prods.filter(p => p.tipo === "Fabricado").map(p => <option key={p.id} value={p.id}>{p.ico} {p.n}</option>)}
        </Sel>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        <Inp label="Fecha inicio" type="date" value={form.f} onChange={e => F("f", e.target.value)} />
        <Inp label="Cantidad" type="number" min="1" value={form.qty} onChange={e => F("qty", e.target.value)} />
      </div>
      {prod && (
        <div style={{ background:"#f4f6f9", borderRadius:10, padding:12, marginTop:12, fontSize:13 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}><span style={{color:"#718096"}}>Costo unitario</span><span style={{fontWeight:700}}>S/ {prod.c}</span></div>
          <div style={{ display:"flex", justifyContent:"space-between", fontWeight:800, fontSize:14 }}><span>Costo total del lote</span><span style={{color:"#e07b39"}}>S/ {prod.c * parseInt(form.qty||1)}</span></div>
        </div>
      )}
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:14 }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" onClick={() => { if (!form.pid || !form.qty) return; onSave({ ...form, qty: parseInt(form.qty), prod: prod?.n || "", costo: prod?.c || 0 }); }}>Crear lote</Btn>
      </div>
    </div>
  );
}

function ClienteForm({ onSave, onCancel }) {
  const [form, setF] = useState({ nombre:"", cel:"", dir:"", notas:"", vip:false });
  const F = (k, v) => setF(f => ({ ...f, [k]: v }));
  return (
    <div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14 }}>
        <div style={{ gridColumn:"1/-1" }}><Inp label="Nombre completo *" value={form.nombre} onChange={e => F("nombre", e.target.value)} placeholder="Ej: Juan Quispe"/></div>
        <Inp label="Celular" type="tel" value={form.cel} onChange={e => F("cel", e.target.value.replace(/\D/g,"").slice(0,9))} placeholder="Ej: 987654321"/>
        <Inp label="Dirección / Distrito" value={form.dir} onChange={e => F("dir", e.target.value)} placeholder="Ej: Ate Vitarte"/>
        <div style={{ gridColumn:"1/-1" }}><Inp label="Notas" value={form.notas} onChange={e => F("notas", e.target.value)} placeholder="Ej: Prefiere roperos en nogal"/></div>
      </div>
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" onClick={() => { if (!form.nombre) return; onSave(form); }}>Guardar</Btn>
      </div>
    </div>
  );
}

function CotizFormModal({ form, setForm, prods, usuarios, onSave, onCancel }) {
  const F = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [busqProd, setBusqProd] = useState("");
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { id:"", n:"", ico:"", col:"", q:1, p:0 }] }));
  const updItem = (i, k, v) => setForm(f => {
    const items = f.items.map((x, j) => {
      if (j !== i) return x;
      let upd = { ...x, [k]: v };
      if (k === "id") {
        const p = prods.find(p => p.id === v);
        if (p) upd = { ...upd, n:p.n, ico:p.ico, col:p.cols[0], p:p.p };
      }
      return upd;
    });
    const tot = items.reduce((a,it) => a + (it.q||1)*(it.p||0), 0);
    return { ...f, items, tot };
  });
  const remItem = (i) => setForm(f => { const items = f.items.filter((_,j) => j !== i); return { ...f, items, tot: items.reduce((a,it) => a+(it.q||1)*(it.p||0),0) }; });
  const vendedores = usuarios.filter(u => u.rol !== "taller" && u.activo).map(u => u.nombre);
  const prodsFiltrados = busqProd ? prods.filter(p => p.n.toLowerCase().includes(busqProd.toLowerCase()) || p.cat.toLowerCase().includes(busqProd.toLowerCase()) || p.id.toLowerCase().includes(busqProd.toLowerCase())) : prods;

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.4)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:16 }}>
      <div style={{ background:"#fff", borderRadius:14, padding:24, width:660, maxWidth:"100%", maxHeight:"92vh", overflowY:"auto", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ fontSize:17, fontWeight:700, marginBottom:20, paddingBottom:14, borderBottom:"1px solid #e2e6ed" }}>📋 {form.id ? "Editar" : "Nueva"} Cotización</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:14 }}>
          <Inp label="Cliente *" value={form.cli} onChange={e => F("cli", e.target.value)} placeholder="Nombre completo"/>
          <Inp label="Celular" value={form.cel} onChange={e => F("cel", e.target.value.replace(/\D/g,"").slice(0,9))} placeholder="987654321"/>
          <Sel label="Vendedor" value={form.vend} onChange={e => F("vend", e.target.value)}>
            {vendedores.map(v => <option key={v}>{v}</option>)}
          </Sel>
        </div>
        <div style={{ marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:600, color:"#718096", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            Productos
            <button onClick={addItem} style={{ background:"#fef3ec", border:"1px solid #e07b39", color:"#e07b39", borderRadius:7, padding:"3px 10px", fontSize:11, fontWeight:700, cursor:"pointer" }}>+ Agregar</button>
          </div>
          {/* Buscador de productos */}
          <div style={{ position:"relative", marginBottom:10 }}>
            <span style={{ position:"absolute", left:9, top:"50%", transform:"translateY(-50%)", color:"#a0aec0", fontSize:13 }}>🔍</span>
            <input value={busqProd} onChange={e => setBusqProd(e.target.value)}
              placeholder="Buscar producto por nombre, categoría o código..."
              style={{ width:"100%", background:"#f4f6f9", border:"1px solid #e2e6ed", borderRadius:8, padding:"7px 10px 7px 28px", fontSize:12, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#1a202c" }}/>
            {busqProd && <button onClick={() => setBusqProd("")} style={{ position:"absolute", right:8, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", color:"#a0aec0", cursor:"pointer", fontSize:14 }}>✕</button>}
          </div>
          {form.items.map((it, i) => (
            <div key={i} style={{ background:"#f9fafb", borderRadius:8, padding:"10px 10px 8px", marginBottom:8 }}>
              <div style={{ display:"grid", gridTemplateColumns:"1fr auto", gap:8, marginBottom:6, alignItems:"center" }}>
                <select value={it.id} onChange={e => updItem(i,"id",e.target.value)}
                  style={{ background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", fontFamily:"inherit", cursor:"pointer", color:"#1a202c", width:"100%" }}>
                  <option value="">— Seleccionar producto —</option>
                  {(busqProd ? prodsFiltrados : prods).map(p => <option key={p.id} value={p.id}>{p.ico} {p.n} — {p.cat} · S/{p.p}</option>)}
                </select>
                <button onClick={() => remItem(i)} style={{ background:"none", border:"none", color:"#e53e3e", cursor:"pointer", fontSize:18, padding:"0 4px" }}>✕</button>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 80px 100px", gap:8 }}>
                <Sel value={it.col} onChange={e => updItem(i,"col",e.target.value)}>
                  {it.id ? (prods.find(p=>p.id===it.id)?.cols||[]).map(c => <option key={c}>{c}</option>) : <option>— Color —</option>}
                </Sel>
                <Inp type="number" min="1" value={it.q} onChange={e => updItem(i,"q",parseInt(e.target.value)||1)} placeholder="Cant."/>
                <Inp type="number" value={it.p} onChange={e => updItem(i,"p",parseFloat(e.target.value)||0)} placeholder="S/ Precio"/>
              </div>
              {it.id && <div style={{ fontSize:10, color:"#a0aec0", marginTop:4 }}>Subtotal: S/ {((it.q||1)*(it.p||0)).toFixed(2)}</div>}
            </div>
          ))}
          {form.items.length === 0 && <div style={{ color:"#a0aec0", fontSize:12, fontStyle:"italic", padding:"8px 0", textAlign:"center" }}>Usa el buscador de arriba y toca "+ Agregar" para añadir productos</div>}
        </div>
        <Inp label="Nota para el cliente" value={form.nota} onChange={e => F("nota", e.target.value)} placeholder="Ej: Incluye instalación, entrega en 5 días hábiles..."/>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:16, paddingTop:14, borderTop:"1px solid #e2e6ed" }}>
          <div style={{ fontSize:20, fontWeight:800, color:"#e07b39" }}>Total: S/ {form.tot.toFixed(2)}</div>
          <div style={{ display:"flex", gap:8 }}>
            <Btn onClick={onCancel}>Cancelar</Btn>
            <Btn variant="primary" onClick={() => { if (!form.cli || !form.items.length) return; onSave(form); }}>Guardar cotización</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

function UsuarioForm({ form, onSave, onCancel, icoset, roles, todosModulos, modulosPorRol }) {
  const [f, setF] = useState({ modulos: modulosPorRol(form.rol || "vendedor"), ...form });
  const U = (k, v) => setF(x => ({ ...x, [k]: v }));
  const cambiarRol = (rol) => setF(x => ({ ...x, rol, modulos: modulosPorRol(rol) }));
  const toggleMod = (id) => setF(x => ({
    ...x,
    modulos: x.modulos.includes(id) ? x.modulos.filter(m => m !== id) : [...x.modulos, id]
  }));
  const ROL_DESC = { admin:"Acceso total", vendedor:"Venta y atención al cliente", taller:"Producción e inventario" };
  // Módulos disponibles para el rol seleccionado
  const modsDisp = todosModulos.filter(m => m.roles.includes(f.rol));
  return (
    <div>
      {/* Ícono */}
      <div style={{ marginBottom:14 }}>
        <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Ícono</label>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
          {icoset.map(ico => (
            <button key={ico} onClick={() => U("ico", ico)}
              style={{ width:38, height:38, borderRadius:10, border:`2px solid ${f.ico===ico?"#e07b39":"#e2e6ed"}`, background:f.ico===ico?"#fef3ec":"transparent", fontSize:20, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}>
              {ico}
            </button>
          ))}
        </div>
      </div>
      {/* Nombre y contraseña */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14 }}>
        <Inp label="Nombre *" value={f.nombre} onChange={e => U("nombre", e.target.value)} placeholder="Ej: María"/>
        <Inp label="Contraseña *" value={f.pass} onChange={e => U("pass", e.target.value)} placeholder="Mínimo 3 caracteres"/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14 }}>
        <Inp label="Celular" value={f.cel||""} onChange={e => U("cel", e.target.value)} placeholder="9XXXXXXXX"/>
        <div>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>🎂 Fecha de cumpleaños</label>
          <input type="date" value={f.cumple||""} onChange={e => U("cumple", e.target.value)}
            style={{ width:"100%", background:"#fff", border:"1px solid #e2e6ed", borderRadius:8, padding:"8px 10px", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#2d3748" }}/>
        </div>
      </div>
      {f.rol === "vendedor" && (
        <div style={{ marginBottom:14 }}>
          <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>💰 Sueldo base semanal S/</label>
          <input type="number" min="0" value={f.sueldoSemanal||""} onChange={e => U("sueldoSemanal", parseFloat(e.target.value)||0)}
            placeholder="Ej: 300"
            style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"8px 10px", fontSize:16, fontWeight:800, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#c4622a" }}/>
          <div style={{ fontSize:10, color:"#a0aec0", marginTop:3 }}>Solo admin puede ver/editar · Se usa en el punto de equilibrio</div>
        </div>
      )}
      {/* Rol */}
      <div style={{ marginBottom:16 }}>
        <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Rol base</label>
        <div style={{ display:"flex", gap:8, marginBottom:5 }}>
          {roles.map(rol => (
            <button key={rol} onClick={() => cambiarRol(rol)}
              style={{ flex:1, padding:"9px 8px", borderRadius:8, border:`2px solid ${f.rol===rol?"#e07b39":"#e2e6ed"}`, background:f.rol===rol?"#fef3ec":"transparent", color:f.rol===rol?"#e07b39":"#718096", fontSize:12, fontWeight:700, cursor:"pointer", textTransform:"capitalize" }}>
              {rol}
            </button>
          ))}
        </div>
        <div style={{ fontSize:11, color:"#a0aec0", fontStyle:"italic" }}>✓ {ROL_DESC[f.rol]} · Cambia el rol para resetear módulos</div>
      </div>
      {/* Módulos con checkboxes */}
      <div style={{ marginBottom:16 }}>
        <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>
          Módulos habilitados ({f.modulos.length}/{modsDisp.length})
        </label>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
          {modsDisp.map(m => {
            const activo = f.modulos.includes(m.id);
            const esEssencial = m.id === "pos"; // POS siempre activo para vendedores
            return (
              <button key={m.id} onClick={() => { if(esEssencial) return; toggleMod(m.id); }}
                style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 10px", borderRadius:8, border:`1px solid ${activo?"#e07b39":"#e2e6ed"}`, background:activo?"#fef3ec":"transparent", cursor:esEssencial?"default":"pointer", textAlign:"left" }}>
                <div style={{ width:16, height:16, borderRadius:4, border:`2px solid ${activo?"#e07b39":"#cbd5e0"}`, background:activo?"#e07b39":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                  {activo && <span style={{ color:"#fff", fontSize:10, fontWeight:900 }}>✓</span>}
                </div>
                <span style={{ fontSize:12, color: activo?"#c4622a":"#718096", fontWeight: activo?600:400 }}>{m.label}</span>
              </button>
            );
          })}
        </div>
      </div>
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" onClick={() => { if (!f.nombre || !f.pass || f.pass.length < 3) return; onSave(f); }}>
          {f.id ? "Guardar cambios" : "Crear usuario"}
        </Btn>
      </div>
    </div>
  );
}

function CompraFormInner({ form, setForm, prods, onSave, onCancel }) {
  const F = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [busqProd, setBusqProd] = useState("");

  const items = form.items || [];

  // Agregar producto (con colQtys vacío)
  const addItem = (prod) => {
    const ya = items.find(it => it.pid === prod.id);
    if (ya) return; // ya está en la lista
    const colQtys = {};
    prod.cols.forEach(c => { colQtys[c] = 0; });
    setForm(f => ({ ...f, items: [...(f.items||[]), { pid:prod.id, prod:prod.n, cu:String(prod.c), colQtys }] }));
    setBusqProd("");
  };

  const updCu = (i, v) => setForm(f => ({ ...f, items:f.items.map((it,j)=>j===i?{...it,cu:v}:it) }));
  const updColQty = (i, col, qty) => setForm(f => ({ ...f, items:f.items.map((it,j)=>j===i?{...it,colQtys:{...it.colQtys,[col]:Math.max(0,parseInt(qty)||0)}}:it) }));
  const remItem = i => setForm(f => ({ ...f, items:f.items.filter((_,j)=>j!==i) }));

  const prodsFiltrados = busqProd
    ? prods.filter(p => p.n.toLowerCase().includes(busqProd.toLowerCase()) || p.cat.toLowerCase().includes(busqProd.toLowerCase()))
    : prods;

  // Calcular totales
  const itemTotals = items.map(it => {
    const totalQty = Object.values(it.colQtys||{}).reduce((a,b)=>a+b,0);
    return { totalQty, subtotal: totalQty * (parseFloat(it.cu)||0) };
  });
  const totalGeneral = itemTotals.reduce((a,t)=>a+t.subtotal, 0);
  const canSave = form.prov && items.length > 0
    && items.every(it => it.pid && parseFloat(it.cu) > 0 && Object.values(it.colQtys||{}).some(q=>q>0));

  return (
    <div>
      {/* Encabezado */}
      <div style={{ background:C.bg, borderRadius:10, padding:"12px 14px", marginBottom:14 }}>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:10 }}>
          <Inp label="Fecha" type="date" value={form.f} onChange={e => F("f", e.target.value)}/>
          <Inp label="Proveedor *" value={form.prov} onChange={e => F("prov", e.target.value)} placeholder="Ej: Maderco SAC"/>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:C.t3, display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>📍 Todo ingresa a</label>
            <select value={form.ubiDest||"Taller/Almacén"} onChange={e => F("ubiDest", e.target.value)}
              style={{ width:"100%", background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px", fontSize:13, fontWeight:600, outline:"none", fontFamily:"inherit", cursor:"pointer", color:C.t1, boxSizing:"border-box" }}>
              {UBICACIONES.map(u => <option key={u}>{u}</option>)}
            </select>
            <div style={{ fontSize:10, color:C.t4, marginTop:2 }}>Recomendado: Taller/Almacén</div>
          </div>
          <Inp label="Nota general" value={form.nota||""} onChange={e => F("nota", e.target.value)} placeholder="Ej: Lote mayo..."/>
        </div>
      </div>

      {/* Buscador */}
      <div style={{ marginBottom:10 }}>
        <div style={{ position:"relative" }}>
          <span style={{ position:"absolute", left:9, top:"50%", transform:"translateY(-50%)", color:C.t4, fontSize:13 }}>🔍</span>
          <input value={busqProd} onChange={e => setBusqProd(e.target.value)}
            placeholder="Buscar modelo para agregar a la lista..."
            style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px 8px 28px", fontSize:12, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:C.t1 }}/>
        </div>
        {busqProd && (
          <div style={{ border:`1px solid ${C.border}`, borderRadius:8, background:C.white, maxHeight:160, overflowY:"auto", marginTop:4 }}>
            {prodsFiltrados.filter(p => !items.find(it=>it.pid===p.id)).map(p => (
              <button key={p.id} onClick={() => addItem(p)}
                style={{ width:"100%", textAlign:"left", padding:"9px 12px", border:"none", borderBottom:`1px solid ${C.border}`, background:"transparent", cursor:"pointer", fontFamily:"inherit", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <span style={{ fontSize:13, color:C.t1 }}>{p.ico} {p.n}</span>
                <span style={{ fontSize:11, color:C.t3 }}>Stock: {getTotalStk(p)} · Costo: S/{p.c}</span>
              </button>
            ))}
            {prodsFiltrados.filter(p=>!items.find(it=>it.pid===p.id)).length===0 && (
              <div style={{ padding:"10px 12px", fontSize:12, color:C.t4 }}>Sin resultados o ya están todos agregados</div>
            )}
          </div>
        )}
      </div>

      {/* Lista de modelos */}
      <div style={{ marginBottom:10 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
          <span style={{ fontSize:11, fontWeight:700, color:C.t3, textTransform:"uppercase", letterSpacing:"0.4px" }}>Modelos a comprar ({items.length})</span>
        </div>
        {items.length === 0 && (
          <div style={{ padding:"20px", textAlign:"center", color:C.t4, fontSize:13, border:`1px dashed ${C.border}`, borderRadius:8 }}>
            Busca un modelo arriba para agregarlo
          </div>
        )}
        {items.map((it, i) => {
          const prod = prods.find(p => p.id === it.pid);
          if (!prod) return null;
          const cu = parseFloat(it.cu)||0;
          const totalQty = itemTotals[i].totalQty;
          const subtotal = itemTotals[i].subtotal;
          return (
            <div key={it.pid} style={{ background:C.bg, border:`1px solid ${C.border}`, borderRadius:12, padding:"12px 14px", marginBottom:10 }}>
              {/* Header del modelo */}
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ fontSize:20 }}>{prod.ico}</span>
                  <div>
                    <div style={{ fontSize:14, fontWeight:700, color:C.t1 }}>{prod.n}</div>
                    <div style={{ fontSize:11, color:C.t3 }}>Stock actual: {getTotalStk(prod)} uds · Costo actual S/{prod.c}</div>
                  </div>
                </div>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ textAlign:"right" }}>
                    <div style={{ fontSize:13, fontWeight:800, color:subtotal>0?C.rd:C.t4 }}>{subtotal>0?`S/${subtotal.toFixed(2)}`:"—"}</div>
                    <div style={{ fontSize:10, color:C.t4 }}>{totalQty} uds total</div>
                  </div>
                  <button onClick={() => remItem(i)} style={{ background:"none", border:"none", color:C.rdL, cursor:"pointer", fontSize:20, padding:"0 4px" }}>✕</button>
                </div>
              </div>

              {/* Costo unitario (aplica a todos los colores de este modelo) */}
              <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:12, padding:"8px 10px", background:C.orBg, borderRadius:8 }}>
                <span style={{ fontSize:12, color:C.or, fontWeight:600 }}>Costo unit. S/</span>
                <input inputMode="decimal" value={it.cu}
                  onChange={e => updCu(i, e.target.value.replace(/[^0-9.]/g,""))}
                  placeholder="0.00"
                  style={{ width:100, background:C.white, border:`2px solid ${C.or}`, borderRadius:7, padding:"6px 10px", fontSize:16, fontWeight:800, outline:"none", color:C.or, textAlign:"center", fontFamily:"inherit" }}/>
                <span style={{ fontSize:11, color:C.t3 }}>← mismo precio para todos los colores</span>
              </div>

              {/* Colores con cantidades */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))", gap:8 }}>
                {prod.cols.map(col => {
                  const qty = it.colQtys?.[col]||0;
                  const colActivo = qty > 0;
                  return (
                    <div key={col} style={{ background:colActivo?C.acBg:C.white, border:`2px solid ${colActivo?C.ac:C.border}`, borderRadius:10, padding:"10px 12px", transition:"all 0.15s" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:8 }}>
                        <div style={{ width:12, height:12, borderRadius:"50%", background:HEX_COLOR[col]||"#888", border:"1px solid rgba(0,0,0,0.15)", flexShrink:0 }}/>
                        <span style={{ fontSize:12, fontWeight:700, color:colActivo?C.ac:C.t2 }}>{col}</span>
                      </div>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <button onClick={() => updColQty(i, col, Math.max(0,qty-1))}
                          style={{ width:28, height:28, borderRadius:7, border:`1px solid ${C.border}`, background:C.bg, fontSize:16, fontWeight:700, cursor:"pointer", color:C.t1 }}>−</button>
                        <input type="number" min="0" value={qty||""}
                          onChange={e => updColQty(i, col, parseInt(e.target.value)||0)}
                          placeholder="0"
                          style={{ flex:1, textAlign:"center", fontSize:16, fontWeight:800, background:C.white, border:`1px solid ${colActivo?C.ac:C.border}`, borderRadius:7, padding:"4px 0", outline:"none", color:colActivo?C.ac:C.t3, fontFamily:"inherit", width:"100%" }}/>
                        <button onClick={() => updColQty(i, col, qty+1)}
                          style={{ width:28, height:28, borderRadius:7, border:`1px solid ${C.border}`, background:C.bg, fontSize:16, fontWeight:700, cursor:"pointer", color:C.t1 }}>+</button>
                      </div>
                      {colActivo && cu > 0 && <div style={{ fontSize:10, color:C.ac, marginTop:4, textAlign:"center" }}>S/{(qty*cu).toFixed(2)}</div>}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>

      {/* Total */}
      {items.length > 0 && (
        <div style={{ background:C.rdBg, borderRadius:8, padding:"10px 14px", marginBottom:14, display:"flex", justifyContent:"space-between", fontSize:15, fontWeight:800 }}>
          <span style={{ color:C.t2 }}>Total ({items.length} modelo{items.length>1?"s":""} · {items.reduce((a,_,i)=>a+itemTotals[i].totalQty,0)} uds)</span>
          <span style={{ color:C.rd }}>S/ {totalGeneral.toFixed(2)}</span>
        </div>
      )}

      {!canSave && items.length > 0 && (
        <div style={{ fontSize:11, color:C.or, marginBottom:10, padding:"6px 10px", background:C.orBg, borderRadius:7 }}>
          ⚠️ {!form.prov?"Falta el proveedor · ":""}
          {items.some(it=>!(parseFloat(it.cu)>0))?"Ingresa el costo en todos los modelos · ":""}
          {items.some(it=>!Object.values(it.colQtys||{}).some(q=>q>0))?"Agrega al menos 1 unidad de algún color en cada modelo":""}
        </div>
      )}

      <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
        <Btn onClick={onCancel}>Cancelar</Btn>
        <Btn variant="primary" disabled={!canSave} onClick={() => {
          if (!canSave) return;
          // Expandir a items planos por color para el registrarCompra
          const itemsPlanos = [];
          items.forEach(it => {
            Object.entries(it.colQtys||{}).forEach(([col, qty]) => {
              if (qty > 0) itemsPlanos.push({ pid:it.pid, prod:it.prod, col, qty, cu:parseFloat(it.cu)||0, ubiDest:form.ubiDest||"Taller/Almacén" });
            });
          });
          onSave({ ...form, items:itemsPlanos });
        }}>
          ✓ Registrar {items.length > 0 ? `(${items.reduce((_,__,i)=>_+itemTotals[i].totalQty,0)} uds · S/${totalGeneral.toFixed(0)})` : ""}
        </Btn>
      </div>
    </div>
  );
}
