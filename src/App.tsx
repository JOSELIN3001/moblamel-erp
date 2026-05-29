// MoblaMel ERP v45 - Build 2
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
const LOGO_SMALL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAABWGlDQ1BJQ0MgUHJvZmlsZQAAeJx9kLFLw1AQxr9WpaB1EB0cHDKJQ5SSCro4tBVEcQhVweqUvqapkMZHkiIFN/+Bgv+BCs5uFoc6OjgIopPo5uSk4KLleS+JpCJ6j+N+fO+74zggOW5wbvcDqDu+W1zKK5ulLSX1jAS9IAzm8Zyur0r+rj/j/T703k7LWb///43Biukxqp+UGcZdH0ioxPqezyXvE4+5tBRxS7IV8onkcsjngWe9WCC+JlZYzagQvxCr5R7d6uG63WDRDnL7tOlsrMk5lBNYxA48cNgw0IQCHdk//LOBv4BdcjfhUp+FGnzqyZEiJ5jEy3DAMAOVWEOGUpN3ju53F91PjbWDJ2ChI4S4iLWVDnA2Rydrx9rUPDAyBFy1ueEagdRHmaxWgddTYLgEjN5Qz7ZXzWrh9uk8MPAoxNskkDoEui0hPo6E6B5T8wNw6XwBA6diE8HYWhMAAAY4SURBVHja7ZrPb1xXFce/33Pvex577NiOx6ntOj+coARKUqPSBYQfAiGEWLCBBWLRRcWWNX8BfwI7trBjgeiiUFS1KrRxIE2dND8bO3HiOE7suPZ4xjPz5r17Dos3dhyJBcrE8USaOxppFjNP93Pv+Z7zPUdDM8PLvAQv+eoCdAG6AF2ALkBby7f/CDMD/p9qSJLPHYAvthIbwE66AYNqujD3eZJsmcHUVFtv6pPLoSBkoaf/wGvTZ9lRIWSEUa9f+rheXvNRrCGPJRpMd90xBSFtDo+++vXp73SWBgjAXKF3IKSJj7ypmgGkwSxHIAkTYQhRVCjsQQS1K2IDoRaCZgiEGgT5DYBqZgQBqFLVtrXOjstC21cBo6VpFiwjSRrJPDkRzLK0maYdCWAAxcRRJGTpmTd/MDh0CDARD4A0UggazUXxHkRQ2wDM0ztJwJSjrxw7ODL+UlViPsnvSppCVUNomob/dV3aiZUYrfoqQHi0fHtkdDwosqDrKw/jnlhVATOFuGhweFg6D4BCgjBDj4+vzX5QKPQfPzmtofnuX/987eLHBw4cCGpJUj80ceQ3v/0dnHReCFEAgGpEFMUXz797Z/5SFMW/fOvXp06/3mxUaSkthaXPX8JtAxiftmgULwizFz6p1WqF3v4f/+wXKjEozkXiPPbAdrUfQtayPDDQVANY+Pb3ftLX17eyvPjHP/y+mdSQ+Waz3mhsdZydNjAHAAwmhpBkOPvdn45NHG8mjQ/ee6cQRVMnvqrQoNlQaWIv7HS7bhQgKQCETJJ0+o0fHZ56LZg6H/38V29Hce82IffIuvvnkIQAAAoD3MHSmKrClBQfFVSVrWbHABORjgMgQGkdrZkRIiI7uSGPGWIP13Mxcw5mAJzI0t1rlc21LM1ajlrNYATNVLw/ceoN731HacBAAi4/fuf9/M1ZVc3P3WA7eVZDiHqLx75ypsMAWnHyRJs+jg1CKAgChAGk0FTj3h5Kp4mYludSEVDIllqznbDPlQGFmcEcTDpRA2mzXquV454+07wFk5am2VK5UDRkaui4NEoKDEeOnxkeHYujHpKkUJxQdk+BSMLMRT3eRx03F2rJ+KWbzJnp9hmrtiKn5YrygzdTablUmCopLdshrcLXGhlR9vMGFCrbNctM893sZgDYegEKk5zTDFDS7f7yi7PTOfD9hbmN9RWCy8t3VleXcj1Uqpu3526QrFYrd+aukEIS4PwXlxv1Cox3F65vlldJkm7p3vzqo+WdB75IAAVw7dLMf/75PsGP/v7O0sIXqmFpcT6O/YVz7y/dm/NOr8yer1a+XLw3R9NzH/7t1vXPhPjHX/60VSlvVSsb66tblbWbn/+bpLXXKPtnyTzAia+dnj1/rrZVLhSK/cXixfMfbm6sP16+XyqN3Lpx+cH9weHhkc3y2p1bV1YfLU1/81uLd28fPzU9ODySpslnF/5V3ywPDBT7+gd2zQVebEfWN3BwZLQ089F7k0ePlDcqG+srZ7//w4X5q4VC3+vfeHP5/j0CaSOBYWV5cfzwlKbh8sWZyWMnHj98sLayUBo/nIYAC/vWUoYsPXbytFIHhobSLBQHh2dmPpk4cqyR1K9f/bR0qEQXX738aaNejuMIEh+eOukk6usdcD7u7RscmzxaGp9M0nTf6kCaNn3kCNTrtdWHy6PjEw8fLE1NHd8sb25srE9MTCbNhgFrjx+XRkfjqMeJ0EflL1dr9Xqx2F+tVkZKpZBlfcXBNrPQswCYWZo2VZUgxUQkUxTiKEma3jvnfJomIo5CoVNVIJhBQS/OEEghHKC56WgTwD9TFbPKZiWvYts7YFWVIoCGrOm8JxBCILk9bTehgDAzUxMRMxPpGRkZ2Z8QClnz5o1LpE28OtloJPVaXcR558TJgaFSbWszaTSGhkuNRi1Jao4+SZIozqd00KA+8hR5ZexI+23+M1qJoKHRrIi5xXu3m0nTeS8iWZoChPi11QchS6OosFWv1qobzUY9jntqjXp/sQgTMzOqKsbGj+6biHdiV0Mmzu2yD7lNkNxPbx+w7vpJq8s3GOH22Y3aznaebvPtyWd7urlpWVdrv4A9Lzu976v7V4MuQBegC9AF6AJ0AboAXYAuwP6t/wLPCmK83LVbLgAAAABJRU5ErkJggg==";
const LOGO_MED   = "data:image/png;base64,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";



// ─── PALETA ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#f4f6f9", bg2:"#eef0f4", bg3:"#e4e8ef",
  white:"#ffffff",
  border:"#e2e6ed", borderD:"#cdd2db",
  ac:"#e07b39", acD:"#c4622a", acL:"#f5a56e", acBg:"#fef3ec",
  gr:"#276749", grBg:"#e6f4ed", grL:"#38a169",
  rd:"#c53030", rdBg:"#fff0f0", rdL:"#e53e3e",
  or:"#c05621", orBg:"#fffaf0",
  bl:"#2b6cb0", blBg:"#ebf4ff",
  pu:"#6b46c1", puBg:"#f3f0ff",
  t1:"#1a202c", t2:"#4a5568", t3:"#718096", t4:"#a0aec0",
  sh:"0 1px 3px rgba(0,0,0,0.08)",
  sh2:"0 4px 12px rgba(0,0,0,0.1)",
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
];

// Módulos por defecto según rol
const modulosPorRol = (rol) => TODOS_MODULOS.filter(m => m.roles.includes(rol)).map(m => m.id);

// ─── USUARIOS (mutable via estado) ───────────────────────────────────────────
const USUARIOS_INIT = [
  { id:"U000", nombre:"Admin",    rol:"admin",    cel:"",          ico:"👑", activo:true, pass:"011207", sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia"] },
  { id:"U001", nombre:"Adolfo",   rol:"admin",    cel:"998649169", ico:"👨‍💼", activo:true, pass:"1234",   sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia"] },
  { id:"U002", nombre:"Joselin",  rol:"admin",    cel:"907569211", ico:"👩‍💼", activo:true, pass:"1234",   sueldoSemanal:0, modulos:["dashboard","pos","ventas","caja","pedidos","traslados","vendedores","clientes","cotiz","inventario","kardex","produccion","compras","gastos","finanzas","abc","prediccion","usuarios","ia"] },
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
const PRODS_INIT = [
  {id:"P001",n:"Ropero 2 Puertas Económico", cat:"Ropero", tipo:"Fabricado",img:"",barcode:"",ico:"🪞",cols:["Blanco","Nogal","Wengué"],
    locs:{"Tienda Principal|Blanco":2,"Tienda Principal|Nogal":1,"Tienda Principal|Wengué":1,"Taller/Almacén|Blanco":1,"Taller/Almacén|Nogal":1},
    stk:6,min:3,p:300,c:220,ubi:"Tienda Principal",v:[12,15,18,14,20,22,19,25,21,17,23,24]},
  {id:"P002",n:"Ropero 3P con Espejo",  cat:"Ropero",      tipo:"Fabricado",img:"",barcode:"",ico:"🪞",cols:["Blanco","Nogal"],
    locs:{"Tienda Pasaje|Blanco":1,"Tienda Pasaje|Nogal":1,"Taller/Almacén|Blanco":1},
    stk:3,min:3,p:680,c:310,ubi:"Tienda Pasaje",v:[8,10,9,12,11,14,10,13,15,12,14,16]},
  {id:"P003",n:"Ropero 4P Lujo",        cat:"Ropero",      tipo:"Comprado", img:"",barcode:"",ico:"🪞",cols:["Blanco","Cerezo"],
    locs:{"Tienda Principal|Blanco":1,"Tienda Principal|Cerezo":1},
    stk:2,min:2,p:950,c:600,ubi:"Tienda Principal",v:[3,4,3,5,4,6,4,5,6,4,5,7]},
  {id:"P004",n:"Zapatero 12 Pares",     cat:"Zapatero",    tipo:"Fabricado",img:"",barcode:"",ico:"👟",cols:["Blanco","Nogal"],
    locs:{"Tienda Principal|Blanco":2,"Tienda Principal|Nogal":1,"Tienda Pasaje|Blanco":2,"Tienda Pasaje|Nogal":2,"Taller/Almacén|Blanco":2,"Taller/Almacén|Nogal":1},
    stk:10,min:4,p:185,c:80,ubi:"Tienda Pasaje",v:[20,22,25,28,24,30,26,32,28,25,29,33]},
  {id:"P005",n:"Zapatero 24 Pares",     cat:"Zapatero",    tipo:"Fabricado",img:"",barcode:"",ico:"👟",cols:["Blanco","Nogal"],
    locs:{"Tienda Pasaje|Blanco":1,"Tienda Pasaje|Nogal":1,"Taller/Almacén|Blanco":2,"Taller/Almacén|Nogal":1},
    stk:5,min:3,p:270,c:120,ubi:"Taller/Almacén",v:[10,12,11,14,12,15,13,16,14,12,15,17]},
  {id:"P006",n:"Cómoda 5 Cajones",      cat:"Cómoda",      tipo:"Fabricado",img:"",barcode:"",ico:"🗄️",cols:["Blanco","Nogal","Wengué"],
    locs:{"Tienda Principal|Blanco":1,"Tienda Principal|Nogal":1,"Tienda Pasaje|Wengué":1,"Taller/Almacén|Blanco":1},
    stk:4,min:3,p:390,c:190,ubi:"Tienda Principal",v:[15,17,16,19,18,21,17,22,20,18,21,23]},
  {id:"P007",n:"Cómoda 3 Cajones",      cat:"Cómoda",      tipo:"Comprado", img:"",barcode:"",ico:"🗄️",cols:["Blanco","Nogal"],
    locs:{"Tienda Principal|Blanco":2,"Tienda Principal|Nogal":2,"Tienda Pasaje|Blanco":1,"Tienda Pasaje|Nogal":1,"Taller/Almacén|Blanco":1,"Taller/Almacén|Nogal":1},
    stk:8,min:2,p:250,c:155,ubi:"Tienda Pasaje",v:[2,1,3,2,1,2,1,2,2,1,2,1]},
  {id:"P008",n:"Mesa de Noche Simple",  cat:"Mesa de Noche",tipo:"Fabricado",img:"",barcode:"",ico:"🛏️",cols:["Blanco","Nogal"],
    locs:{"Tienda Principal|Blanco":3,"Tienda Principal|Nogal":3,"Tienda Pasaje|Blanco":2,"Tienda Pasaje|Nogal":1,"Taller/Almacén|Blanco":2,"Taller/Almacén|Nogal":1},
    stk:12,min:5,p:98,c:40,ubi:"Tienda Principal",v:[30,35,32,38,36,40,35,42,38,34,39,43]},
  {id:"P009",n:"Mesa Noche con Cajón",  cat:"Mesa de Noche",tipo:"Fabricado",img:"",barcode:"",ico:"🛏️",cols:["Blanco","Nogal","Cerezo"],
    locs:{"Tienda Principal|Blanco":1,"Tienda Principal|Nogal":1,"Tienda Pasaje|Cerezo":1,"Taller/Almacén|Blanco":2,"Taller/Almacén|Nogal":1},
    stk:6,min:4,p:135,c:60,ubi:"Taller/Almacén",v:[18,20,19,22,21,24,20,25,23,20,24,26]},
  {id:"P010",n:"Cabecera Tapizada 2P",  cat:"Cabecera",    tipo:"Comprado", img:"",barcode:"",ico:"🛋️",cols:["Gris","Beige"],
    locs:{"Tienda Principal|Gris":1,"Tienda Pasaje|Gris":1,"Tienda Pasaje|Beige":1,"Taller/Almacén|Beige":1},
    stk:4,min:3,p:330,c:210,ubi:"Tienda Pasaje",v:[8,9,8,10,9,11,9,12,10,9,11,12]},
];

const VENTAS_INIT = [
  {id:"V001",f:"2025-05-05",cli:"Juan Quispe",  vend:"Rosa",      items:[{id:"P001",n:"Ropero 2 Puertas",     col:"Blanco",q:1,p:480,c:220}],tot:480, mp:"Efectivo",     comp:"Boleta", num:"B001-00021"},
  {id:"V002",f:"2025-05-08",cli:"Rosa Mamani",  vend:"Adolfo",      items:[{id:"P004",n:"Zapatero 12 Pares",    col:"Nogal", q:2,p:185,c:80}], tot:350, mp:"Yape",         comp:"Boleta", num:"B001-00022"},
  {id:"V003",f:"2025-05-12",cli:"Carlos Flores",vend:"Joselin",     items:[{id:"P008",n:"Mesa Noche Simple",    col:"Blanco",q:2,p:98, c:40},{id:"P006",n:"Cómoda 5 Cajones",col:"Nogal",q:1,p:390,c:190}],tot:615,mp:"POS Tarjeta",comp:"Factura",num:"F001-00001"},
  {id:"V004",f:"2025-05-15",cli:"Lucía Torres", vend:"Rosa",      items:[{id:"P002",n:"Ropero 3P con Espejo", col:"Blanco",q:1,p:680,c:310}],tot:650, mp:"Trans. BCP",    comp:"Boleta", num:"B001-00023"},
  {id:"V005",f:"2025-05-19",cli:"Pedro Condori",vend:"Joselin",     items:[{id:"P009",n:"Mesa Noche con Cajón", col:"Cerezo",q:2,p:135,c:60}], tot:270, mp:"Yape",         comp:"Boleta", num:"B001-00024"},
  {id:"V006",f:"2025-05-21",cli:"Ana Huanca",   vend:"Maricielo", items:[{id:"P004",n:"Zapatero 12 Pares",    col:"Blanco",q:1,p:185,c:80}], tot:185, mp:"Efectivo",     comp:"Boleta", num:"B001-00025"},
];

const GASTOS_INIT = [
  // ── GASTOS FIJOS ──────────────────────────────────────────────────────────
  {id:"G001",f:"2025-05-01",cat:"Alquiler",      esFijo:true,  desc:"Alquiler Tienda Principal · mayo", prov:"",               monto:1700},
  {id:"G002",f:"2025-05-01",cat:"Alquiler",      esFijo:true,  desc:"Alquiler Tienda Pasaje · mayo",    prov:"",               monto:2000},
  {id:"G003",f:"2025-05-01",cat:"Servicios taller", esFijo:true, desc:"Luz y agua taller · mayo",      prov:"",               monto:120},
  {id:"G004",f:"2025-05-05",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Rosa · semana 1",     vendedor:"Rosa",     prov:"",               monto:300},
  {id:"G005",f:"2025-05-05",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Adolfo · semana 1",   vendedor:"Adolfo",     prov:"",               monto:280},
  {id:"G006",f:"2025-05-05",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Maricielo · semana 1",vendedor:"Maricielo",prov:"",               monto:260},
  {id:"G007",f:"2025-05-12",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Rosa · semana 2",     vendedor:"Rosa",     prov:"",               monto:300},
  {id:"G008",f:"2025-05-12",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Adolfo · semana 2",   vendedor:"Adolfo",     prov:"",               monto:280},
  {id:"G009",f:"2025-05-12",cat:"Sueldo base",   esFijo:true,  desc:"Sueldo base Maricielo · semana 2",vendedor:"Maricielo",prov:"",               monto:260},
  // ── GASTOS VARIABLES ─────────────────────────────────────────────────────
  {id:"G010",f:"2025-05-03",cat:"Materiales",    esFijo:false, desc:"Planchas melamina 18mm x20",       prov:"Maderco SAC",    monto:1200},
  {id:"G011",f:"2025-05-08",cat:"Armador por obra",esFijo:false,desc:"Armado lote roperos mayo",        prov:"Luis Quispe",    monto:450},
  {id:"G012",f:"2025-05-15",cat:"Materiales",    esFijo:false, desc:"Jaladores y bisagras x50",         prov:"Ferretería VES", monto:180},
  {id:"G013",f:"2025-05-20",cat:"Comisión",      esFijo:false, desc:"Comisión Rosa · ventas sem 3",    vendedor:"Rosa",     prov:"",               monto:96},
  {id:"G014",f:"2025-05-20",cat:"Comisión",      esFijo:false, desc:"Comisión Adolfo · ventas sem 3",  vendedor:"Adolfo",     prov:"",               monto:72},
  {id:"G015",f:"2025-05-22",cat:"Mantenimiento", esFijo:false, desc:"Afilado sierras y mantenimiento",  prov:"Taller técnico", monto:85},
];

const KARDEX_INIT = [
  {id:"K001",f:"2025-05-01",pid:"P001",prod:"Ropero 2 Puertas",    col:"Blanco",tipo:"Fabricación",desc:"Lote fabricación",   ent:5,sal:0,saldo:11,costo:220},
  {id:"K002",f:"2025-05-05",pid:"P001",prod:"Ropero 2 Puertas",    col:"Blanco",tipo:"Venta",      desc:"Venta B001-00021",  ent:0,sal:1,saldo:10,costo:220},
  {id:"K003",f:"2025-05-08",pid:"P004",prod:"Zapatero 12 Pares",   col:"Nogal", tipo:"Fabricación",desc:"Lote fabricación",   ent:8,sal:0,saldo:18,costo:80},
  {id:"K004",f:"2025-05-08",pid:"P004",prod:"Zapatero 12 Pares",   col:"Nogal", tipo:"Venta",      desc:"Venta B001-00022",  ent:0,sal:2,saldo:16,costo:80},
  {id:"K005",f:"2025-05-12",pid:"P008",prod:"Mesa Noche Simple",   col:"Blanco",tipo:"Venta",      desc:"Venta F001-00001",  ent:0,sal:2,saldo:14,costo:40},
  {id:"K006",f:"2025-05-12",pid:"P006",prod:"Cómoda 5 Cajones",    col:"Nogal", tipo:"Venta",      desc:"Venta F001-00001",  ent:0,sal:1,saldo:5, costo:190},
  {id:"K007",f:"2025-05-15",pid:"P003",prod:"Ropero 4P Lujo",      col:"Blanco",tipo:"Compra",     desc:"Compra proveedor",  ent:3,sal:0,saldo:5, costo:600},
  {id:"K008",f:"2025-05-15",pid:"P002",prod:"Ropero 3P con Espejo",col:"Blanco",tipo:"Venta",      desc:"Venta B001-00023",  ent:0,sal:1,saldo:4, costo:310},
];

const LOTES_INIT = [
  {id:"L001",f:"2025-05-01",pid:"P001",prod:"Ropero 2 Puertas",   qty:5,costo:220,total:1100,estado:"Completado"},
  {id:"L002",f:"2025-05-08",pid:"P004",prod:"Zapatero 12 Pares",  qty:8,costo:80, total:640, estado:"Completado"},
  {id:"L003",f:"2025-05-15",pid:"P006",prod:"Cómoda 5 Cajones",   qty:4,costo:190,total:760, estado:"En proceso"},
];

const CLIENTES_INIT = [
  {id:"C001",nombre:"Juan Quispe",  cel:"987654321",dir:"Ate Vitarte",    notas:"Prefiere roperos nogal",  vip:false},
  {id:"C002",nombre:"Rosa Mamani",  cel:"912345678",dir:"San Juan de Lurigancho",notas:"",               vip:false},
  {id:"C003",nombre:"Carlos Flores",cel:"956123456",dir:"Miraflores",     notas:"Cliente corporativo",    vip:true },
  {id:"C004",nombre:"Lucía Torres", cel:"934567890",dir:"San Borja",      notas:"Interesada en dormitorio completo", vip:false},
  {id:"C005",nombre:"Pedro Condori",cel:"967890123",dir:"Villa El Salvador",notas:"",                    vip:false},
  {id:"C006",nombre:"Ana Huanca",   cel:"945678901",dir:"Chorrillos",     notas:"",                      vip:false},
];

const COTIZ_INIT = [
  {id:"Q001",f:"2025-05-10",cli:"Lucía Torres",cel:"934567890",vend:"Rosa",items:[{id:"P001",n:"Ropero 2 Puertas",barcode:"",ico:"🪞",col:"Blanco",q:1,p:480}],tot:480,est:"Pendiente",nota:"Quiere entrega en 1 semana"},
  {id:"Q002",f:"2025-05-14",cli:"Marco Ríos",  cel:"923456789",vend:"Adolfo",items:[{id:"P006",n:"Cómoda 5 Cajones",barcode:"",ico:"🗄️",col:"Nogal",q:2,p:390},{id:"P008",n:"Mesa de Noche Simple",barcode:"",ico:"🛏️",col:"Nogal",q:2,p:98}],tot:976,est:"Vista",nota:"Pidió descuento, pendiente confirmar"},
  {id:"Q003",f:"2025-05-18",cli:"Sara Mejía",  cel:"978901234",vend:"Mia",items:[{id:"P003",n:"Ropero 4P Lujo",barcode:"",ico:"🪞",col:"Cerezo",q:1,p:950}],tot:950,est:"Aceptada",nota:""},
];

const COMPRAS_INIT = [
  {id:"CM001",f:"2025-05-03",prov:"Maderco SAC",      pid:"P001",prod:"Ropero 2 Puertas",    col:"Blanco",qty:5,  cu:210, total:1050,nota:"Lote melamina importada"},
  {id:"CM002",f:"2025-05-15",prov:"Muebles Lima Import",pid:"P003",prod:"Ropero 4P Lujo",   col:"Blanco",qty:3,  cu:595, total:1785,nota:"Stock navideño"},
  {id:"CM003",f:"2025-05-20",prov:"Cabeceras Perú",   pid:"P010",prod:"Cabecera Tapizada 2P",col:"Gris",  qty:2,  cu:205, total:410, nota:""},
];

const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

const PEDIDOS_INIT = [
  { id:"PD001", f:"2025-05-10", tipo:"separacion", cli:"Lucía Torres", cel:"934567890", vend:"Rosa",
    pid:"P001", prod:"Ropero 2 Puertas", col:"Blanco", qty:1, precioAcordado:480,
    adelanto:200, abonos:[], saldoPendiente:280, est:"Pendiente entrega",
    fEnt:"2025-05-25", nota:"Cliente viene el sábado a recoger", diasAlerta:30 },
  { id:"PD002", f:"2025-05-15", tipo:"medida", cli:"Marco Ríos", cel:"923456789", vend:"Adolfo",
    pid:"", prod:"Ropero 3P Cerezo especial 2.20m", col:"Cerezo", qty:1, precioAcordado:780,
    adelanto:300, abonos:[], saldoPendiente:480, est:"En producción",
    fEnt:"2025-06-05", nota:"Medida especial 2.20m alto, color cerezo oscuro", diasAlerta:30 },
];
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
      <div style={{ background:C.white, borderRadius:14, padding:24, width: wide ? 620 : 520, maxWidth:"100%", maxHeight:"90vh", overflowY:"auto", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
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
  const [numB,   setNumB]   = useState(26);
  const [numF,   setNumF]   = useState(2);

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
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg,${C.bg} 0%,${C.bg2} 100%)`, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'Segoe UI',sans-serif" }}>
      <div style={{ width:380, background:C.white, border:`1px solid ${C.border}`, borderRadius:16, padding:36, boxShadow:C.sh2 }}>
        <div style={{ textAlign:"center", marginBottom:28 }}>
          <div style={{ width:80, height:80, borderRadius:16, overflow:"hidden", margin:"0 auto 12px", display:"flex", alignItems:"center", justifyContent:"center", background:"#f9f6f0" }}>
            <img src={LOGO_MED} alt="MoblaMel" style={{ width:76, height:76, objectFit:"contain" }}/>
          </div>
          <div style={{ fontSize:22, fontWeight:800, color:C.ac }}>MoblaMel</div>
          <div style={{ fontSize:12, color:C.t3, marginTop:2 }}>Sistema de gestión · Mueblería de melamina</div>
        </div>
        <div style={{ marginBottom:12 }}><Inp label="Usuario" placeholder="Ej: Admin, Mamá..." value={lU} onChange={e => setLU(e.target.value)} onKeyDown={e => e.key === "Enter" && login()} /></div>
        <div style={{ marginBottom:16 }}><Inp label="Contraseña" type="password" value={lP} onChange={e => setLP(e.target.value)} onKeyDown={e => e.key === "Enter" && login()} /></div>
        {lErr && <div style={{ color:C.rd, fontSize:12, marginBottom:10, textAlign:"center" }}>{lErr}</div>}
        <Btn variant="primary" full onClick={login}>Ingresar →</Btn>
        
      </div>
    </div>
  );

  const isAdmin   = user.rol === "admin";
  // Módulos que este usuario tiene habilitados (admin siempre ve todo)
  const userModulos = user.modulos || modulosPorRol(user.rol);
  const navItems = TODOS_MODULOS.filter(m => userModulos.includes(m.id)).map(m => ({
    id: m.id, ico: { dashboard:"🏠",pos:"🛒",ventas:"💰",pedidos:"📬",traslados:"🚚",vendedores:"🏆",clientes:"👥",cotiz:"📋",inventario:"📦",kardex:"📒",produccion:"🔨",compras:"🏭",gastos:"💸",finanzas:"📊",abc:"🏆",prediccion:"🔮",usuarios:"👤",ia:"🤖" }[m.id], label: m.label.replace(/^.{2}/,"") // strip emoji for label
  })).map(m => ({ ...m, label: TODOS_MODULOS.find(x=>x.id===m.id)?.label.slice(2).trim() || m.label }));

  const modActual = navItems.find(n => n.id === mod) ? mod : navItems[0]?.id || "pos";

  // Auto-colapsar sidebar en POS para maximizar espacio en móvil
  useEffect(() => { setSidebarOpen(modActual !== "pos"); }, [modActual]);

  // ─── RENDER PRINCIPAL ────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.t1, fontFamily:"'Segoe UI','Helvetica Neue',sans-serif", fontSize:14, display:"flex", flexDirection:"column" }}>

      {/* TOPBAR */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.border}`, height:52, display:"flex", alignItems:"center", padding:"0 20px", gap:16, flexShrink:0, boxShadow:C.sh, position:"sticky", top:0, zIndex:50 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <div style={{ width:32, height:32, borderRadius:8, overflow:"hidden", background:"#f9f6f0", display:"flex", alignItems:"center", justifyContent:"center" }}>
            <img src={LOGO_SMALL} alt="M" style={{ width:30, height:30, objectFit:"contain" }}/>
          </div>
          <span style={{ fontSize:15, fontWeight:800, color:C.ac }}>MoblaMel</span>
        </div>
        <div style={{ flex:1 }} />
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

      <div style={{ display:"flex", flex:1, overflow:"hidden", height:"calc(100vh - 52px)" }}>

        {/* SIDEBAR */}
        <div style={{ width: sidebarOpen ? 200 : 52, background:C.white, borderRight:`1px solid ${C.border}`, padding: sidebarOpen ? "12px 8px" : "12px 6px", flexShrink:0, overflowY:"auto", overflowX:"hidden", display:"flex", flexDirection:"column", transition:"width 0.2s ease", position:"relative" }}>
          {/* Toggle collapse button */}
          <button onClick={() => setSidebarOpen(o => !o)}
            style={{ display:"flex", alignItems:"center", justifyContent:"center", width:"100%", padding:"6px 4px", marginBottom:10, borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, cursor:"pointer", color:C.t3, fontSize:14, flexShrink:0 }}
            title={sidebarOpen ? "Colapsar menú" : "Expandir menú"}>
            {sidebarOpen ? "◀" : "▶"}
          </button>
          {(isAdmin || user.rol === "vendedor") && (
            <button onClick={() => { setModal("venta"); }} style={{ width:"100%", padding:"9px", borderRadius:8, border:`1px solid ${C.ac}`, background:C.acBg, color:C.ac, fontSize:sidebarOpen?12:18, fontWeight:700, cursor:"pointer", fontFamily:"inherit", textAlign:"center", marginBottom:12, whiteSpace:"nowrap", overflow:"hidden" }}
              title="Nueva Venta">
              {sidebarOpen ? "🧾 + Nueva Venta" : "🧾"}
            </button>
          )}
          {sidebarOpen && <div style={{ fontSize:10, color:C.t4, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.8px", padding:"0 8px", marginBottom:6 }}>Módulos</div>}
          {navItems.map(n => (
            <button key={n.id} onClick={() => { setMod(n.id); if(n.id==="pos") setSidebarOpen(false); }} title={n.label}
              style={{ width:"100%", padding: sidebarOpen?"8px 10px":"8px 4px", borderRadius:8, border:"none", background: modActual === n.id ? C.acBg : "transparent", color: modActual === n.id ? C.ac : C.t3, fontSize:12, fontWeight: modActual === n.id ? 700 : 400, cursor:"pointer", textAlign:sidebarOpen?"left":"center", fontFamily:"inherit", marginBottom:1, display:"flex", alignItems:"center", gap:sidebarOpen?8:0, justifyContent:sidebarOpen?"flex-start":"center", transition:"all 0.12s", overflow:"hidden" }}>
              <span style={{ fontSize:sidebarOpen?15:18, flexShrink:0 }}>{n.ico}</span>
              {sidebarOpen && <span style={{ whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{n.label}</span>}
              {sidebarOpen && (n.id === "ia" || n.id === "prediccion") && <span style={{ marginLeft:"auto", fontSize:9, background: n.id === "ia" ? C.ac : C.bl, color:"#fff", padding:"1px 5px", borderRadius:99, fontWeight:700 }}>IA</span>}
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

        {/* CONTENIDO */}
        <div style={{ flex:1, overflowY:"auto", padding:22, background:C.bg }}>

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
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:20 }}>
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
                  return (
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:20 }}>
                      <KPI icon="💰" label="Mis ventas del mes" value={fmt(miTotal)} color={C.grL} sub={`${misVentasMes.length} ventas`}/>
                      <KPI icon="📊" label="Ticket promedio" value={fmt(misVentasMes.length?miTotal/misVentasMes.length:0)} color={C.ac}/>
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
                <div style={{display:"flex",height:"calc(100vh - 120px)",minHeight:500,borderRadius:14,overflow:"hidden",border:`1px solid ${C.border}`,boxShadow:"0 4px 24px rgba(0,0,0,0.08)"}}>

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
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:14 }}>
                <KPI icon="💰" label={hayFiltro?"Total filtrado":"Ingresos del mes"} value={fmt(hayFiltro?totalFilt:kpi.ing)} color={C.grL} sub={`${ventasFilt.length} ventas cerradas`}/>
                <KPI icon="🎫" label="Ticket promedio" value={fmt(ventasFilt.length ? totalFilt/ventasFilt.length : 0)} color={C.ac}/>
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
            // Ventas del día del usuario actual (vendedor ve solo las suyas, admin ve todas)
            const ventasDia = ventas.filter(v => v.f === HOY && !v.esAnticipo && (isAdmin || v.vend === user.nombre));
            const anticiposDia = ventas.filter(v => v.f === HOY && v.esAnticipo && (isAdmin || v.vend === user.nombre));
            const totalDia = ventasDia.reduce((a,v) => a+v.tot, 0);
            const totalAnt = anticiposDia.reduce((a,v) => a+v.tot, 0);
            // Agrupar por método de pago
            const metodos = {};
            ventasDia.forEach(v => {
              v.mp.split("+").forEach(mp => {
                const m = mp.trim();
                metodos[m] = (metodos[m]||0) + v.tot / v.mp.split("+").length;
              });
            });
            const yaCerrado = cierres.some(c => c.f === HOY && (isAdmin ? true : c.vendedor === user.nombre));

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
              {/* Resumen del día */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:12, marginBottom:20 }}>
                <KPI icon="🧾" label="Ventas hoy" value={ventasDia.length} color={C.grL} sub={`S/${totalDia.toFixed(2)}`}/>
                <KPI icon="💰" label="Total ventas" value={`S/${totalDia.toFixed(2)}`} color={C.ac}/>
                <KPI icon="📥" label="Anticipos" value={`S/${totalAnt.toFixed(2)}`} color={C.bl} sub={`${anticiposDia.length} pagos`}/>
                <KPI icon="💵" label="Total en caja" value={`S/${(totalDia+totalAnt).toFixed(2)}`} color={C.gr}/>
              </div>
              {/* Desglose por método */}
              <Card title="Desglose por método de pago" style={{ marginBottom:16 }}>
                {Object.keys(metodos).length === 0
                  ? <div style={{ color:C.t3, fontSize:13, padding:"8px 0" }}>No hay ventas registradas hoy</div>
                  : Object.entries(metodos).map(([mp, tot]) => (
                    <div key={mp} style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:`1px solid ${C.border}`, fontSize:14 }}>
                      <span style={{ color:C.t2 }}>{mp}</span>
                      <strong style={{ color:C.ac }}>S/ {tot.toFixed(2)}</strong>
                    </div>
                  ))
                }
              </Card>
              {/* Lista de ventas del día */}
              <Card title={`Ventas del día (${ventasDia.length})`} style={{ marginBottom:16 }}>
                {ventasDia.length === 0
                  ? <div style={{ color:C.t3, fontSize:13 }}>Sin ventas hoy</div>
                  : ventasDia.map(v => (
                    <div key={v.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:13 }}>
                      <div>
                        <span style={{ fontWeight:600, color:C.t1 }}>{v.num}</span>
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
              {/* Historial de cierres */}
              {cierres.length > 0 && (
                <Card title="Historial de cierres anteriores">
                  {cierres.slice(0,10).map(c => (
                    <div key={c.id} style={{ display:"flex", justifyContent:"space-between", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:13 }}>
                      <span style={{ color:C.t2 }}>{c.f} · {c.hora}</span>
                      <span style={{ color:C.t3 }}>{c.vendedor}</span>
                      <strong style={{ color:C.ac }}>S/ {(c.totalGeneral||c.total||0).toFixed(2)}</strong>
                    </div>
                  ))}
                </Card>
              )}
              {/* Botón cierre */}
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
            return (<>
              <PageTitle title="👤 Usuarios" sub={`${usuarios.length} usuarios · ${usuarios.filter(u=>u.activo).length} activos`}
                action={<Btn variant="primary" onClick={() => setModal({ tipo:"nuevoUser", form:{ nombre:"", pass:"", rol:"vendedor", ico:"👧", modulos:modulosPorRol("vendedor") } })}>+ Nuevo Usuario</Btn>}/>
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
        {f.rol === "vendedor" && (
          <div>
            <label style={{ fontSize:11, fontWeight:600, color:"#718096", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.4px" }}>💰 Sueldo base semanal S/</label>
            <input type="number" min="0" value={f.sueldoSemanal||""} onChange={e => U("sueldoSemanal", parseFloat(e.target.value)||0)}
              placeholder="Ej: 300"
              style={{ width:"100%", background:"#fff", border:"2px solid #e07b39", borderRadius:8, padding:"8px 10px", fontSize:16, fontWeight:800, outline:"none", boxSizing:"border-box", fontFamily:"inherit", color:"#c4622a" }}/>
            <div style={{ fontSize:10, color:"#a0aec0", marginTop:3 }}>Solo admin puede ver/editar · Se usa en el punto de equilibrio</div>
          </div>
        )}
      </div>
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
