import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, Legend, ResponsiveContainer, CartesianGrid } from "recharts";

/* ── IPADE brand palette ── */
const IPADE = {
  navy: "#1B2A4A", darkNavy: "#0F1B33", gold: "#C8A951", lightGold: "#E8D48B",
  white: "#FFFFFF", offWhite: "#F5F3EE", warmGray: "#E8E4DD", midGray: "#9A9590",
  accent1: "#2E5090", accent2: "#4A7AB5",
};

const SEDE_PALETTE_DEFAULT = {
  MEX:"#1B2A4A", GDL:"#2E7D32", MTY:"#C62828", HMO:"#E65100", LEO:"#6A1B9A",
  CUL:"#00838F", GUA:"#AD1457", MOR:"#4E342E", STAFE:"#1565C0", AGU:"#558B2F",
  Virtual:"#455A64", Vir:"#607D8B",
};
const PROF_PALETTE_DEFAULT = {
  JRMC:"#1B5E20", JSG:"#B71C1C", ACFz:"#0D47A1", AAM:"#E65100",
  VTP:"#4A148C", AJGO:"#006064", CERM:"#880E4F", "(Vacío)":"#9E9E9E",
};
const PROG_PALETTE_DEFAULT = {
  InCompany:"#C8A951", Perfeccionamiento:"#1B2A4A", CA:"#2E5090",
  SEPOS:"#C62828", Enfocados:"#2E7D32", "Máster":"#6A1B9A", MEDEX:"#00695C",
};
const SPECIAL_PROGRAMS = ["InCompany", "Enfocados", "SEPOS"];

/* ── helpers ── */
function excelDateToJS(serial) {
  if (!serial || typeof serial !== "number") return null;
  return new Date((Math.floor(serial - 25569)) * 86400 * 1000);
}
function excelTimeToStr(t) {
  if (t == null || t === "" || typeof t !== "number") return "";
  const totalMin = Math.round(t * 24 * 60);
  return `${String(Math.floor(totalMin / 60)).padStart(2,"0")}:${String(totalMin % 60).padStart(2,"0")}`;
}
function fmtDate(d) {
  if (!d) return "";
  return d.toLocaleDateString("es-MX", { day:"2-digit", month:"2-digit", year:"numeric" });
}
function ColorDot({ color, size = 10 }) {
  return <span style={{ display:"inline-block", width:size, height:size, borderRadius:"50%", backgroundColor:color, marginRight:6, flexShrink:0 }} />;
}

/* ── hex color utils ── */
function hexToRgb(hex) {
  const h = (hex||"#999999").replace("#","");
  return { r: parseInt(h.substring(0,2),16), g: parseInt(h.substring(2,4),16), b: parseInt(h.substring(4,6),16) };
}
function lighten(hex, amt = 0.82) {
  const {r,g,b} = hexToRgb(hex);
  const mix = c => Math.round(c * (1-amt) + 255 * amt);
  return {r:mix(r), g:mix(g), b:mix(b)};
}
function rgbToHex({r,g,b}) {
  return [r,g,b].map(c => c.toString(16).padStart(2,"0")).join("").toUpperCase();
}

/* ── color picker ── */
function ColorPicker({ color, onChange, label }) {
  const [open, setOpen] = useState(false);
  return (
    <span style={{ position:"relative", display:"inline-flex", alignItems:"center", gap:4 }}>
      <span onClick={()=>setOpen(!open)} style={{ width:18, height:18, borderRadius:4, backgroundColor:color, border:"2px solid #ccc", cursor:"pointer", display:"inline-block" }} />
      <span style={{ fontSize:12, color:"#555" }}>{label}</span>
      {open && (
        <span style={{ position:"absolute", top:24, left:0, zIndex:999, background:"#fff", border:"1px solid #ddd", borderRadius:6, padding:6, boxShadow:"0 4px 16px rgba(0,0,0,.15)" }}>
          <input type="color" value={color} onChange={e=>onChange(e.target.value)} style={{ width:50, height:30, border:"none", cursor:"pointer" }} />
          <div style={{ textAlign:"center", marginTop:4 }}><button onClick={()=>setOpen(false)} style={{ fontSize:10, cursor:"pointer", border:"none", background:IPADE.navy, color:"#fff", borderRadius:3, padding:"2px 8px" }}>OK</button></div>
        </span>
      )}
    </span>
  );
}

/* ── multiselect filter dropdown ── */
function FilterDropdown({ values, selected, onChange, label }) {
  const [open, setOpen] = useState(false);
  const ref = useRef();
  useEffect(() => {
    const handler = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);
  const allSelected = selected.length === values.length;
  return (
    <div ref={ref} style={{ position:"relative", display:"inline-block" }}>
      <button onClick={()=>setOpen(!open)} style={{ fontSize:11, padding:"3px 8px", background:selected.length<values.length?IPADE.gold:"#e8e4dd", color:selected.length<values.length?"#fff":"#333", border:"1px solid #ccc", borderRadius:4, cursor:"pointer", whiteSpace:"nowrap", maxWidth:140, overflow:"hidden", textOverflow:"ellipsis" }}>
        {label} {selected.length<values.length?`(${selected.length})`:"▾"}
      </button>
      {open && (
        <div style={{ position:"absolute", top:"100%", left:0, zIndex:100, background:"#fff", border:"1px solid #ddd", borderRadius:6, maxHeight:260, overflowY:"auto", minWidth:180, boxShadow:"0 4px 16px rgba(0,0,0,.12)", padding:4 }}>
          <label style={{ display:"flex", alignItems:"center", gap:4, padding:"4px 8px", fontSize:12, fontWeight:600, cursor:"pointer", borderBottom:"1px solid #eee" }}>
            <input type="checkbox" checked={allSelected} onChange={()=>onChange(allSelected?[]:[...values])} /> Todos
          </label>
          {values.map(v => (
            <label key={v} style={{ display:"flex", alignItems:"center", gap:4, padding:"3px 8px", fontSize:12, cursor:"pointer" }}>
              <input type="checkbox" checked={selected.includes(v)} onChange={()=>onChange(selected.includes(v)?selected.filter(x=>x!==v):[...selected,v])} /> {v}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

/* ── KPI card ── */
function KPI({ label, value, sub, color }) {
  return (
    <div style={{ background:"#fff", borderRadius:10, padding:"18px 20px", minWidth:160, flex:"1 1 160px", borderLeft:`5px solid ${color||IPADE.navy}`, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
      <div style={{ fontSize:28, fontWeight:800, color:IPADE.navy, fontFamily:"'DM Serif Display', Georgia, serif" }}>{value}</div>
      <div style={{ fontSize:13, color:"#666", marginTop:2, fontWeight:500 }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:IPADE.midGray, marginTop:2 }}>{sub}</div>}
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   COLORED EXCEL GENERATOR — builds SpreadsheetML XML 
   directly (no dependency on SheetJS Pro styles)
   ══════════════════════════════════════════════════════════ */
function generateColoredExcelBlob(data, sedeColors, profColors, progColors) {
  const headers = ["Ciclo","Programa","Sede Sesión","Modalidad","Grupo","Profesor","Secuencia","Módulo","Fecha Sesión","Hora inicio","Hora fin","Tema","Caso/Nota"];
  const COL_PROG = 1, COL_SEDE = 2, COL_PROF = 5;

  /* Collect unique colors needed */
  const colorSet = new Set();
  colorSet.add("1B2A4A"); colorSet.add("FFFFFF"); colorSet.add("F5F3EE"); colorSet.add("E0DDD6");
  data.forEach(r => {
    const add = hex => { const c = (hex||"#999").replace("#","").toUpperCase(); colorSet.add(c); colorSet.add(rgbToHex(lighten(hex||"#999"))); };
    add(progColors[r._progPadre]); add(sedeColors[r._sede]); add(profColors[r._profesor]);
  });

  /* Build style index maps */
  const fills = ["none","gray125"]; // 0=none, 1=gray125 (defaults required by Excel)
  const fillMap = {};
  const addFill = (fgHex) => {
    if (fillMap[fgHex] !== undefined) return fillMap[fgHex];
    fills.push(fgHex);
    fillMap[fgHex] = fills.length - 1;
    return fills.length - 1;
  };
  // Pre-add fills
  const hdrFillIdx = addFill("1B2A4A");
  const evenFillIdx = addFill("F5F3EE");
  const whiteFillIdx = addFill("FFFFFF");
  // add color fills
  data.forEach(r => {
    addFill(rgbToHex(lighten(progColors[r._progPadre]||"#999")));
    addFill(rgbToHex(lighten(sedeColors[r._sede]||"#999")));
    addFill(rgbToHex(lighten(profColors[r._profesor]||"#999")));
  });

  const fonts = [];
  const fontMap = {};
  const addFont = (key, sz, bold, colorHex) => {
    if (fontMap[key] !== undefined) return fontMap[key];
    fonts.push({ sz, bold, color: colorHex });
    fontMap[key] = fonts.length - 1;
    return fonts.length - 1;
  };
  const hdrFontIdx = addFont("hdr", 11, true, "FFFFFF");
  const baseFontIdx = addFont("base", 10, false, "1B2A4A");
  data.forEach(r => {
    const pc = (progColors[r._progPadre]||"#999").replace("#","").toUpperCase();
    const sc = (sedeColors[r._sede]||"#999").replace("#","").toUpperCase();
    const prc = (profColors[r._profesor]||"#999").replace("#","").toUpperCase();
    addFont("bold_"+pc, 10, true, pc);
    addFont("bold_"+sc, 10, true, sc);
    addFont("bold_"+prc, 10, true, prc);
  });

  /* xfs (cell formats): combination of font + fill */
  const xfs = [];
  const xfMap = {};
  const addXf = (fontIdx, fillIdx) => {
    const key = `${fontIdx}_${fillIdx}`;
    if (xfMap[key] !== undefined) return xfMap[key];
    xfs.push({ fontIdx, fillIdx });
    xfMap[key] = xfs.length - 1;
    return xfs.length - 1;
  };
  const defaultXfIdx = addXf(baseFontIdx, 0);
  const hdrXfIdx = addXf(hdrFontIdx, hdrFillIdx);

  /* Pre-compute xf for each data cell */
  const rowXfs = data.map(r => {
    const isEven = false; // will be set per-row
    const progHex = progColors[r._progPadre]||"#999";
    const sedeHex = sedeColors[r._sede]||"#999";
    const profHex = profColors[r._profesor]||"#999";
    return { progHex, sedeHex, profHex };
  });

  /* Build shared strings */
  const sst = [];
  const sstMap = {};
  const addStr = s => {
    const v = String(s ?? "");
    if (sstMap[v] !== undefined) return sstMap[v];
    sst.push(v);
    sstMap[v] = sst.length - 1;
    return sst.length - 1;
  };
  headers.forEach(h => addStr(h));
  const rowVals = data.map(r => [
    r["Ciclo"]||"", r._programa||"", r._sede||"", r["Modalidad"]||"", r["Grupo"]||"",
    r._profesor||"", r["Secuencia"]!=null?String(r["Secuencia"]):"", r["Módulo"]!=null?String(r["Módulo"]):"",
    fmtDate(r._fecha), r._horaInicio||"", r._horaFin||"", r["Tema"]||"", r["Caso/Nota"]||"",
  ]);
  rowVals.forEach(row => row.forEach(v => addStr(v)));

  /* escape XML */
  const esc = s => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  /* Build XML parts */
  const colName = c => { let s = ""; let n = c; while(n >= 0){ s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n/26) - 1; } return s; };

  let sheetRows = "";
  // Header row
  sheetRows += `<row r="1" spans="1:${headers.length}">`;
  headers.forEach((h,ci) => {
    sheetRows += `<c r="${colName(ci)}1" t="s" s="${hdrXfIdx}"><v>${addStr(h)}</v></c>`;
  });
  sheetRows += `</row>`;

  // Data rows
  rowVals.forEach((vals, ri) => {
    const rowNum = ri + 2;
    const rd = data[ri];
    const isEven = ri % 2 === 0;
    const baseFillIdx2 = isEven ? whiteFillIdx : evenFillIdx;
    const progHex = progColors[rd._progPadre]||"#999";
    const sedeHex = sedeColors[rd._sede]||"#999";
    const profHex = profColors[rd._profesor]||"#999";

    sheetRows += `<row r="${rowNum}" spans="1:${headers.length}">`;
    vals.forEach((v, ci) => {
      let xfIdx;
      if (ci === COL_PROG) {
        const fIdx = addFont("bold_"+(progHex.replace("#","").toUpperCase()), 10, true, progHex.replace("#","").toUpperCase());
        const flIdx = addFill(rgbToHex(lighten(progHex)));
        xfIdx = addXf(fIdx, flIdx);
      } else if (ci === COL_SEDE) {
        const fIdx = addFont("bold_"+(sedeHex.replace("#","").toUpperCase()), 10, true, sedeHex.replace("#","").toUpperCase());
        const flIdx = addFill(rgbToHex(lighten(sedeHex)));
        xfIdx = addXf(fIdx, flIdx);
      } else if (ci === COL_PROF) {
        const fIdx = addFont("bold_"+(profHex.replace("#","").toUpperCase()), 10, true, profHex.replace("#","").toUpperCase());
        const flIdx = addFill(rgbToHex(lighten(profHex)));
        xfIdx = addXf(fIdx, flIdx);
      } else {
        xfIdx = addXf(baseFontIdx, baseFillIdx2);
      }
      sheetRows += `<c r="${colName(ci)}${rowNum}" t="s" s="${xfIdx}"><v>${addStr(v)}</v></c>`;
    });
    sheetRows += `</row>`;
  });

  const colsXml = headers.map((_,i) => {
    const w = [8,22,10,14,6,10,8,8,12,10,10,40,40][i] || 12;
    return `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
  }).join("");

  const fontsXml = fonts.map(f =>
    `<font><sz val="${f.sz}"/>${f.bold?"<b/>":""}<color rgb="FF${f.color}"/><name val="Calibri"/></font>`
  ).join("");

  const fillsXml = fills.map((f,i) => {
    if (i === 0) return `<fill><patternFill patternType="none"/></fill>`;
    if (i === 1) return `<fill><patternFill patternType="gray125"/></fill>`;
    return `<fill><patternFill patternType="solid"><fgColor rgb="FF${f}"/></patternFill></fill>`;
  }).join("");

  const xfsXml = xfs.map(x =>
    `<xf fontId="${x.fontIdx}" fillId="${x.fillIdx}" borderId="0" numFmtId="0" applyFont="1" applyFill="1"/>`
  ).join("");

  const sstXml = sst.map(s => `<si><t>${esc(s)}</t></si>`).join("");

  /* Build the XLSX zip manually using SheetJS's zip writer */
  // Actually, we'll use SheetJS to create the zip structure but inject our own XML
  // Simpler approach: generate the XML files, then use the built-in XLSX zip tools

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Colorigrama" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="${fonts.length}">${fontsXml}</fonts>
  <fills count="${fills.length}">${fillsXml}</fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${xfs.length}">${xfsXml}</cellXfs>
</styleSheet>`;

  const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.length}" uniqueCount="${sst.length}">${sstXml}</sst>`;

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>${colsXml}</cols>
  <sheetData>${sheetRows}</sheetData>
</worksheet>`;

  /* Use JSZip-style approach via SheetJS internal zip (cfb) — but actually
     the simplest reliable approach in a browser sandbox is to use the XLSX 
     utility to write a zip. SheetJS exposes XLSX.utils and can write raw. */
  // We'll build a minimal zip manually using Uint8Arrays
  
  const files = {
    "[Content_Types].xml": contentTypes,
    "_rels/.rels": rels,
    "xl/_rels/workbook.xml.rels": xlRels,
    "xl/workbook.xml": workbook,
    "xl/styles.xml": styles,
    "xl/sharedStrings.xml": sharedStrings,
    "xl/worksheets/sheet1.xml": sheet,
  };

  // Minimal ZIP implementation
  function createZip(fileEntries) {
    const enc = new TextEncoder();
    const entries = Object.entries(fileEntries).map(([name, content]) => ({
      name: enc.encode(name),
      data: enc.encode(content),
    }));
    const parts = [];
    const centralDir = [];
    let offset = 0;
    entries.forEach(({name, data}) => {
      // Local file header
      const header = new Uint8Array(30 + name.length);
      const view = new DataView(header.buffer);
      view.setUint32(0, 0x04034b50, true); // sig
      view.setUint16(4, 20, true); // version needed
      view.setUint16(6, 0, true); // flags
      view.setUint16(8, 0, true); // compression: store
      view.setUint16(10, 0, true); // mod time
      view.setUint16(12, 0, true); // mod date
      // CRC32
      const crc = crc32(data);
      view.setUint32(14, crc, true);
      view.setUint32(18, data.length, true); // compressed size
      view.setUint32(22, data.length, true); // uncompressed size
      view.setUint16(26, name.length, true);
      view.setUint16(28, 0, true); // extra field length
      header.set(name, 30);

      // Central directory entry
      const cdEntry = new Uint8Array(46 + name.length);
      const cdView = new DataView(cdEntry.buffer);
      cdView.setUint32(0, 0x02014b50, true);
      cdView.setUint16(4, 20, true);
      cdView.setUint16(6, 20, true);
      cdView.setUint16(8, 0, true);
      cdView.setUint16(10, 0, true);
      cdView.setUint16(12, 0, true);
      cdView.setUint16(14, 0, true);
      cdView.setUint32(16, crc, true);
      cdView.setUint32(20, data.length, true);
      cdView.setUint32(24, data.length, true);
      cdView.setUint16(28, name.length, true);
      cdView.setUint16(30, 0, true);
      cdView.setUint16(32, 0, true);
      cdView.setUint16(34, 0, true);
      cdView.setUint16(36, 0, true);
      cdView.setUint32(38, 0, true);
      cdView.setUint32(42, offset, true);
      cdEntry.set(name, 46);
      centralDir.push(cdEntry);

      parts.push(header, data);
      offset += header.length + data.length;
    });

    const cdOffset = offset;
    let cdSize = 0;
    centralDir.forEach(cd => { parts.push(cd); cdSize += cd.length; });

    // End of central directory
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, entries.length, true);
    eocdView.setUint16(10, entries.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, cdOffset, true);
    eocdView.setUint16(20, 0, true);
    parts.push(eocd);

    const totalLen = parts.reduce((a, p) => a + p.length, 0);
    const result = new Uint8Array(totalLen);
    let pos = 0;
    parts.forEach(p => { result.set(p, pos); pos += p.length; });
    return result;
  }

  function crc32(data) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) {
      crc ^= data[i];
      for (let j = 0; j < 8; j++) {
        crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
      }
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  const zipBytes = createZip(files);
  return new Blob([zipBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}


/* ══════════════ MAIN APP ══════════════ */
export default function Colorigrama() {
  const [rawData, setRawData] = useState(null);
  const [fileName, setFileName] = useState("");
  const [sedeColors, setSedeColors] = useState({ ...SEDE_PALETTE_DEFAULT });
  const [profColors, setProfColors] = useState({ ...PROF_PALETTE_DEFAULT });
  const [progColors, setProgColors] = useState({ ...PROG_PALETTE_DEFAULT });
  const [showColorPanel, setShowColorPanel] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filters, setFilters] = useState({});
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [exportStatus, setExportStatus] = useState("");

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const parsed = json.map(r => ({
        ...r,
        _profesor: r["Profesores"] ? String(r["Profesores"]).trim() : "(Vacío)",
        _sede: r["Sede Sesión"] ? String(r["Sede Sesión"]).trim() : "",
        _progPadre: r["Programa Padre"] ? String(r["Programa Padre"]).trim() : "",
        _programa: r["Programa"] ? String(r["Programa"]).trim() : "",
        _fecha: excelDateToJS(r["Fecha sesión (Día/Mes/Año)"]),
        _horaInicio: excelTimeToStr(r["Hora inicio"]),
        _horaFin: excelTimeToStr(r["Hora fin"]),
      }));
      parsed.sort((a,b) => (a._fecha||0) - (b._fecha||0));
      const esc2 = { ...SEDE_PALETTE_DEFAULT }, epc = { ...PROF_PALETTE_DEFAULT }, eprc = { ...PROG_PALETTE_DEFAULT };
      const fb = ["#37474F","#5D4037","#1A237E","#BF360C","#004D40","#33691E","#263238","#F57F17","#880E4F","#311B92"];
      let ci2 = 0;
      parsed.forEach(r => {
        if (r._sede && !esc2[r._sede]) esc2[r._sede] = fb[ci2++ % fb.length];
        if (r._profesor && !epc[r._profesor]) epc[r._profesor] = fb[ci2++ % fb.length];
        if (r._progPadre && !eprc[r._progPadre]) eprc[r._progPadre] = fb[ci2++ % fb.length];
      });
      setSedeColors(esc2); setProfColors(epc); setProgColors(eprc);
      setRawData(parsed); setFilters({});
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const DISPLAY_COLS = useMemo(() => [
    { key:"Ciclo", label:"Ciclo" }, { key:"_progPadre", label:"Programa Padre" },
    { key:"_programa", label:"Programa" }, { key:"_sede", label:"Sede Sesión" },
    { key:"Modalidad", label:"Modalidad" }, { key:"Grupo", label:"Grupo" },
    { key:"_profesor", label:"Profesor" }, { key:"Secuencia", label:"Secuencia" },
    { key:"Módulo", label:"Módulo" }, { key:"_fecha", label:"Fecha Sesión", fmt:fmtDate },
    { key:"_horaInicio", label:"Hora inicio" }, { key:"_horaFin", label:"Hora fin" },
    { key:"Tema", label:"Tema" }, { key:"Caso/Nota", label:"Caso/Nota" },
  ], []);

  const filterableValues = useMemo(() => {
    if (!rawData) return {};
    const out = {};
    DISPLAY_COLS.forEach(c => {
      if (["_fecha","_horaInicio","_horaFin","Tema","Caso/Nota","Secuencia","Módulo"].includes(c.key)) return;
      out[c.key] = [...new Set(rawData.map(r => String(r[c.key] ?? "")))].sort();
    });
    return out;
  }, [rawData, DISPLAY_COLS]);

  useEffect(() => {
    if (!rawData) return;
    const init = {};
    Object.entries(filterableValues).forEach(([k,v]) => { init[k] = [...v]; });
    setFilters(init);
  }, [rawData, filterableValues]);

  const filteredData = useMemo(() => {
    if (!rawData) return [];
    let d = rawData.filter(r => Object.entries(filters).every(([k,sel]) => {
      if (!sel || sel.length === 0) return false;
      return sel.includes(String(r[k] ?? ""));
    }));
    if (sortCol) {
      d = [...d].sort((a,b) => {
        let va = a[sortCol] ?? "", vb = b[sortCol] ?? "";
        if (sortCol === "_fecha") { va = va||new Date(0); vb = vb||new Date(0); return sortDir==="asc"?va-vb:vb-va; }
        va = String(va); vb = String(vb);
        return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
      });
    }
    return d;
  }, [rawData, filters, sortCol, sortDir]);

  const metrics = useMemo(() => {
    if (!filteredData.length) return null;
    const total = filteredData.length;
    const local = filteredData.filter(r => r._sede === "MEX").length;
    const foranea = total - local;
    const byProf = {}, specialByProf = {}, bySede = {};
    filteredData.forEach(r => { byProf[r._profesor] = (byProf[r._profesor]||0)+1; bySede[r._sede] = (bySede[r._sede]||0)+1; });
    filteredData.filter(r => SPECIAL_PROGRAMS.includes(r._progPadre)).forEach(r => { specialByProf[r._profesor] = (specialByProf[r._profesor]||0)+1; });
    return { total, local, foranea, byProf, specialByProf, bySede };
  }, [filteredData]);

  const sedeChartData = useMemo(() => metrics ? Object.entries(metrics.bySede).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);
  const profChartData = useMemo(() => metrics ? Object.entries(metrics.byProf).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);
  const specialChartData = useMemo(() => metrics ? Object.entries(metrics.specialByProf).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);

  /* ══════ EXPORT: COLORED EXCEL (real .xlsx with cell fills) ══════ */
  const exportExcel = useCallback(() => {
    if (!filteredData.length) return;
    setExportStatus("Generando Excel coloreado...");
    try {
      const blob = generateColoredExcelBlob(filteredData, sedeColors, profColors, progColors);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Colorigrama_IPADE.xlsx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setExportStatus("✅ Excel descargado");
    } catch(e) {
      console.error(e);
      setExportStatus("⚠️ Error, descargando sin estilos...");
      // Fallback: plain xlsx
      const rows2 = filteredData.map(r => ({
        Ciclo:r["Ciclo"], Programa:r._programa, "Sede Sesión":r._sede,
        Modalidad:r["Modalidad"], Grupo:r["Grupo"], Profesor:r._profesor,
        Secuencia:r["Secuencia"], "Módulo":r["Módulo"], "Fecha Sesión":fmtDate(r._fecha),
        "Hora inicio":r._horaInicio, "Hora fin":r._horaFin, Tema:r["Tema"], "Caso/Nota":r["Caso/Nota"],
      }));
      const ws2 = XLSX.utils.json_to_sheet(rows2);
      const wb2 = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb2, ws2, "Colorigrama");
      XLSX.writeFile(wb2, "Colorigrama_IPADE.xlsx");
    }
    setTimeout(() => setExportStatus(""), 3000);
  }, [filteredData, sedeColors, profColors, progColors]);

  /* ══════ EXPORT: PDF (direct download as .html — user prints to PDF) ══════ */
  const exportPDF = useCallback(() => {
    if (!filteredData.length || !metrics) return;
    setExportStatus("Generando reporte...");
    const esc = s => String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    const kpi = metrics;
    const rows = filteredData.map((r,i) => {
      const pC=progColors[r._progPadre]||"#999", sC=sedeColors[r._sede]||"#999", prC=profColors[r._profesor]||"#999";
      const bg = i%2===0?"#fff":"#f9f8f5";
      return `<tr style="background:${bg}"><td>${esc(r["Ciclo"])}</td><td style="background:${pC}22;border-left:3px solid ${pC};font-weight:600;color:${pC}">${esc(r._programa)}</td><td style="background:${sC}22;border-left:3px solid ${sC};font-weight:600;color:${sC}">${esc(r._sede)}</td><td>${esc(r["Modalidad"])}</td><td>${esc(r["Grupo"])}</td><td style="background:${prC}22;border-left:3px solid ${prC};font-weight:600;color:${prC}">${esc(r._profesor)}</td><td>${esc(r["Secuencia"])}</td><td>${esc(r["Módulo"])}</td><td style="white-space:nowrap">${esc(fmtDate(r._fecha))}</td><td>${esc(r._horaInicio)}</td><td>${esc(r._horaFin)}</td><td class="w">${esc(r["Tema"])}</td><td class="w">${esc(r["Caso/Nota"])}</td></tr>`;
    }).join("");
    const sedeKeys=[...new Set(filteredData.map(r=>r._sede))].sort();
    const profKeys=[...new Set(filteredData.map(r=>r._profesor))].sort();
    const progKeys=[...new Set(filteredData.map(r=>r._progPadre))].sort();
    const dot=(c,l)=>`<span style="display:inline-flex;align-items:center;gap:3px;margin-right:10px"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${c}"></span>${l}</span>`;

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Colorigrama IPADE</title>
<style>@page{size:landscape;margin:10mm}*{margin:0;padding:0;box-sizing:border-box}body{font-family:Helvetica,Arial,sans-serif;color:#1B2A4A;font-size:9px}.hdr{background:#1B2A4A;color:#fff;padding:14px 18px;display:flex;justify-content:space-between;align-items:center;-webkit-print-color-adjust:exact;print-color-adjust:exact}.hdr h1{font-size:16px}.hdr .g{color:#C8A951;font-size:9px;letter-spacing:3px;text-transform:uppercase}.kpis{display:flex;gap:8px;padding:10px 18px;flex-wrap:wrap}.kpi{border:1px solid #ddd;border-radius:5px;padding:6px 12px;border-left:4px solid #1B2A4A}.kpi .n{font-size:18px;font-weight:800;color:#1B2A4A}.kpi .l{font-size:8px;color:#666}.leg{padding:6px 18px;font-size:8px;display:flex;gap:16px;flex-wrap:wrap;border-bottom:1px solid #eee}.leg b{margin-right:6px}table{width:100%;border-collapse:collapse;margin-top:2px}th{background:#1B2A4A;color:#fff;padding:4px;text-align:left;font-size:7.5px;font-weight:600;-webkit-print-color-adjust:exact;print-color-adjust:exact}td{padding:3px 4px;border-bottom:1px solid #eee;font-size:7.5px;vertical-align:top;-webkit-print-color-adjust:exact;print-color-adjust:exact}td.w{max-width:160px;word-wrap:break-word}.note{padding:12px 18px;font-size:9px;color:#888;text-align:center;margin-top:8px}@media print{table{page-break-inside:auto}tr{page-break-inside:avoid}}</style></head><body>
<div class="hdr"><div><div class="g">IPADE Business School</div><h1>Colorigrama — Reporte de Sesiones</h1></div><div style="text-align:right;font-size:9px;opacity:.7">${new Date().toLocaleDateString("es-MX")}</div></div>
<div class="kpis"><div class="kpi"><div class="n">${kpi.total}</div><div class="l">Total</div></div><div class="kpi" style="border-left-color:#2E7D32"><div class="n">${kpi.local}</div><div class="l">Locales</div></div><div class="kpi" style="border-left-color:#C62828"><div class="n">${kpi.foranea}</div><div class="l">Foráneas</div></div><div class="kpi" style="border-left-color:#C8A951"><div class="n">${Object.values(kpi.specialByProf).reduce((a,b)=>a+b,0)}</div><div class="l">Especiales</div></div></div>
<div class="leg"><div><b>Sedes:</b>${sedeKeys.map(k=>dot(sedeColors[k]||"#999",k)).join("")}</div><div><b>Profesores:</b>${profKeys.map(k=>dot(profColors[k]||"#999",k)).join("")}</div><div><b>Programa:</b>${progKeys.map(k=>dot(progColors[k]||"#999",k)).join("")}</div></div>
<table><thead><tr><th>Ciclo</th><th>Programa</th><th>Sede</th><th>Modalidad</th><th>Grupo</th><th>Profesor</th><th>Sec.</th><th>Mód.</th><th>Fecha</th><th>H.Ini</th><th>H.Fin</th><th>Tema</th><th>Caso/Nota</th></tr></thead><tbody>${rows}</tbody></table>
<div class="note">Para guardar como PDF: usa Ctrl+P (o Cmd+P) → "Guardar como PDF" con orientación horizontal</div>
</body></html>`;

    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Colorigrama_IPADE.html";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setExportStatus("✅ Reporte descargado — ábrelo y usa Ctrl+P para guardar como PDF");
    setTimeout(() => setExportStatus(""), 5000);
  }, [filteredData, metrics, sedeColors, profColors, progColors]);

  const handleSort = col => {
    if (sortCol === col) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  /* ══════════ RENDER ══════════ */
  if (!rawData) {
    return (
      <div style={{ minHeight:"100vh", background:`linear-gradient(145deg, ${IPADE.darkNavy} 0%, ${IPADE.navy} 50%, ${IPADE.accent1} 100%)`, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'Instrument Sans', sans-serif", padding:20 }}>
        <div style={{ textAlign:"center", maxWidth:520 }}>
          <div style={{ fontSize:11, letterSpacing:6, color:IPADE.gold, textTransform:"uppercase", marginBottom:12, fontWeight:700 }}>IPADE Business School</div>
          <h1 style={{ fontFamily:"'DM Serif Display', Georgia, serif", fontSize:42, color:IPADE.white, marginBottom:8, lineHeight:1.15 }}>Colorigrama</h1>
          <p style={{ color:"rgba(255,255,255,.6)", fontSize:15, marginBottom:32, lineHeight:1.6 }}>Sistema de análisis visual de sesiones académicas.<br/>Carga tu archivo Excel para comenzar.</p>
          <label style={{ display:"inline-flex", alignItems:"center", gap:10, background:IPADE.gold, color:IPADE.darkNavy, fontWeight:700, fontSize:15, padding:"14px 32px", borderRadius:8, cursor:"pointer", boxShadow:"0 4px 20px rgba(200,169,81,.35)" }}>
            📁 Cargar archivo Excel
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display:"none" }} />
          </label>
          <p style={{ color:"rgba(255,255,255,.35)", fontSize:12, marginTop:16 }}>Formatos aceptados: .xlsx, .xls</p>
        </div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily:"'Instrument Sans', sans-serif", background:IPADE.offWhite, minHeight:"100vh", color:IPADE.navy }}>
      {/* HEADER */}
      <header style={{ background:IPADE.navy, padding:"14px 24px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
        <div>
          <span style={{ fontSize:10, letterSpacing:4, color:IPADE.gold, textTransform:"uppercase", fontWeight:700 }}>IPADE Business School</span>
          <h1 style={{ fontFamily:"'DM Serif Display', Georgia, serif", fontSize:22, color:"#fff", margin:0 }}>Colorigrama</h1>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ color:"rgba(255,255,255,.5)", fontSize:12 }}>📄 {fileName}</span>
          <label style={{ fontSize:12, color:IPADE.gold, cursor:"pointer", padding:"5px 12px", border:`1px solid ${IPADE.gold}`, borderRadius:5 }}>
            Cambiar archivo <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display:"none" }} />
          </label>
          <button onClick={()=>setShowColorPanel(!showColorPanel)} style={{ fontSize:12, padding:"5px 12px", background:showColorPanel?IPADE.gold:"transparent", color:showColorPanel?IPADE.darkNavy:IPADE.gold, border:`1px solid ${IPADE.gold}`, borderRadius:5, cursor:"pointer" }}>
            🎨 Colores
          </button>
        </div>
      </header>

      {/* TABS */}
      <div style={{ background:IPADE.darkNavy, display:"flex", gap:0, paddingLeft:24 }}>
        {[["dashboard","📊 Dashboard"],["tabla","📋 Tabla Detallada"]].map(([id,label]) => (
          <button key={id} onClick={()=>setActiveTab(id)} style={{ padding:"10px 20px", fontSize:13, fontWeight:activeTab===id?700:400, background:activeTab===id?IPADE.offWhite:"transparent", color:activeTab===id?IPADE.navy:"rgba(255,255,255,.6)", border:"none", borderRadius:"8px 8px 0 0", cursor:"pointer" }}>
            {label}
          </button>
        ))}
      </div>

      {/* COLOR PANEL */}
      {showColorPanel && (
        <div style={{ background:"#fff", borderBottom:"1px solid #ddd", padding:"14px 24px", display:"flex", gap:32, flexWrap:"wrap", fontSize:12 }}>
          <div>
            <div style={{ fontWeight:700, marginBottom:6 }}>Colores de Sede</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              {Object.entries(sedeColors).map(([k,v])=>(<ColorPicker key={k} color={v} label={k} onChange={c=>setSedeColors(p=>({...p,[k]:c}))} />))}
            </div>
          </div>
          <div>
            <div style={{ fontWeight:700, marginBottom:6 }}>Colores de Profesor</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              {Object.entries(profColors).map(([k,v])=>(<ColorPicker key={k} color={v} label={k} onChange={c=>setProfColors(p=>({...p,[k]:c}))} />))}
            </div>
          </div>
          <div>
            <div style={{ fontWeight:700, marginBottom:6 }}>Colores de Programa Padre</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              {Object.entries(progColors).map(([k,v])=>(<ColorPicker key={k} color={v} label={k} onChange={c=>setProgColors(p=>({...p,[k]:c}))} />))}
            </div>
          </div>
        </div>
      )}

      {/* FILTERS BAR */}
      <div style={{ background:"#fff", borderBottom:"1px solid #eee", padding:"10px 24px", display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, fontWeight:700, marginRight:8 }}>Filtros:</span>
        {Object.entries(filterableValues).map(([k,vals]) => {
          const colDef = DISPLAY_COLS.find(c=>c.key===k);
          return <FilterDropdown key={k} values={vals} selected={filters[k]||[]} onChange={sel=>setFilters(p=>({...p,[k]:sel}))} label={colDef?.label||k} />;
        })}
        <button onClick={()=>{const init={};Object.entries(filterableValues).forEach(([k,v])=>{init[k]=[...v];});setFilters(init);}} style={{ fontSize:11, padding:"3px 10px", background:IPADE.navy, color:"#fff", border:"none", borderRadius:4, cursor:"pointer", marginLeft:8 }}>
          Limpiar filtros
        </button>
        <div style={{ marginLeft:"auto", display:"flex", gap:6, alignItems:"center" }}>
          {exportStatus && <span style={{ fontSize:11, color:IPADE.gold, fontWeight:600 }}>{exportStatus}</span>}
          <button onClick={exportExcel} style={{ fontSize:11, padding:"5px 14px", background:"#217346", color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontWeight:600 }}>📥 Excel coloreado</button>
          <button onClick={exportPDF} style={{ fontSize:11, padding:"5px 14px", background:"#C62828", color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontWeight:600 }}>📄 Reporte PDF</button>
        </div>
      </div>
      {/* Info bar */}
      <div style={{ background:"#FFFDE7", padding:"6px 24px", fontSize:11, color:"#5D4037", borderBottom:"1px solid #eee" }}>
        💡 Las descargas contienen <strong>solo los {filteredData.length} registros filtrados</strong> que se muestran en la tabla. Ajusta los filtros arriba para controlar qué se descarga.
      </div>

      {/* MAIN CONTENT */}
      <div style={{ padding:24 }}>
        {activeTab === "dashboard" && metrics && (
          <>
            <div style={{ display:"flex", gap:14, flexWrap:"wrap", marginBottom:24 }}>
              <KPI label="Total Sesiones" value={metrics.total} color={IPADE.navy} />
              <KPI label="Sesiones Locales (MEX)" value={metrics.local} color="#2E7D32" sub={`${((metrics.local/metrics.total)*100).toFixed(1)}%`} />
              <KPI label="Sesiones Foráneas" value={metrics.foranea} color="#C62828" sub={`${((metrics.foranea/metrics.total)*100).toFixed(1)}%`} />
              <KPI label="Sesiones Especiales" value={Object.values(metrics.specialByProf).reduce((a,b)=>a+b,0)} color={IPADE.gold} sub="InCompany + Enfocados + SEPOS" />
              <KPI label="Profesores Activos" value={Object.keys(metrics.byProf).filter(p=>p!=="(Vacío)").length} color={IPADE.accent1} />
              <KPI label="Sesiones sin Profesor" value={metrics.byProf["(Vacío)"]||0} color={IPADE.midGray} />
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(340px, 1fr))", gap:20, marginBottom:24 }}>
              <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones por Sede</h3>
                <ResponsiveContainer width="100%" height={280}>
                  <PieChart><Pie data={sedeChartData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={100} label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} style={{fontSize:11}}>
                    {sedeChartData.map(d=>(<Cell key={d.name} fill={sedeColors[d.name]||"#999"} />))}
                  </Pie><Tooltip/></PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones por Profesor</h3>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={profChartData} layout="vertical" margin={{left:40,right:20}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis type="number" style={{fontSize:11}}/><YAxis type="category" dataKey="name" width={50} style={{fontSize:11}}/><Tooltip/>
                    <Bar dataKey="value" radius={[0,6,6,0]}>{profChartData.map(d=>(<Cell key={d.name} fill={profColors[d.name]||"#999"}/>))}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones Especiales por Profesor</h3>
                <p style={{ fontSize:11, color:"#888", marginBottom:10 }}>InCompany, Enfocados, SEPOS</p>
                {specialChartData.length>0 ? (
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={specialChartData} layout="vertical" margin={{left:40,right:20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis type="number" style={{fontSize:11}}/><YAxis type="category" dataKey="name" width={50} style={{fontSize:11}}/><Tooltip/>
                      <Bar dataKey="value" radius={[0,6,6,0]}>{specialChartData.map(d=>(<Cell key={d.name} fill={profColors[d.name]||IPADE.gold}/>))}</Bar>
                    </BarChart>
                  </ResponsiveContainer>
                ) : <p style={{color:"#999",fontSize:13,textAlign:"center",padding:40}}>Sin sesiones especiales con los filtros actuales</p>}
              </div>
              <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Local vs. Foráneas</h3>
                <ResponsiveContainer width="100%" height={260}>
                  <PieChart><Pie data={[{name:"Local (MEX)",value:metrics.local},{name:"Foráneas",value:metrics.foranea}]} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={55} outerRadius={95} label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} style={{fontSize:11}}>
                    <Cell fill="#2E7D32"/><Cell fill="#C62828"/>
                  </Pie><Tooltip/><Legend/></PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </>
        )}

        {/* TABLE */}
        {(activeTab==="tabla"||activeTab==="dashboard") && (
          <div style={{ background:"#fff", borderRadius:12, boxShadow:"0 2px 8px rgba(0,0,0,.06)", overflow:"hidden" }}>
            {activeTab==="dashboard" && <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, padding:"16px 20px 0" }}>Vista de Datos ({filteredData.length} registros)</h3>}
            <div style={{ overflowX:"auto", maxHeight:activeTab==="tabla"?"calc(100vh - 260px)":"none", overflowY:activeTab==="tabla"?"auto":"visible" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr>{DISPLAY_COLS.map(c=>(
                    <th key={c.key} onClick={()=>handleSort(c.key)} style={{ background:IPADE.navy, color:"#fff", padding:"10px 8px", textAlign:"left", fontSize:11, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap", userSelect:"none", position:"sticky", top:0, zIndex:10 }}>
                      {c.label} {sortCol===c.key?(sortDir==="asc"?"▲":"▼"):""}
                    </th>
                  ))}</tr>
                </thead>
                <tbody>
                  {(activeTab==="dashboard"?filteredData.slice(0,50):filteredData).map((r,i)=>(
                    <tr key={i} style={{ background:i%2===0?"#fff":"#f9f8f5" }} onMouseEnter={e=>e.currentTarget.style.background="#f0efe9"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#fff":"#f9f8f5"}>
                      {DISPLAY_COLS.map(c=>{
                        let val = c.fmt?c.fmt(r[c.key]):(r[c.key]??"");
                        let style = { padding:"8px", borderBottom:"1px solid #f0efea", fontSize:11, maxWidth:c.key==="Tema"||c.key==="Caso/Nota"?220:"auto", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" };
                        if (c.key==="_progPadre"||c.key==="_programa") { const pc=progColors[r._progPadre]||"#999"; style.borderLeft=`4px solid ${pc}`; style.background=`${pc}15`; }
                        if (c.key==="_sede") { const sc=sedeColors[r._sede]||"#999"; style.borderLeft=`4px solid ${sc}`; style.background=`${sc}15`; val=<span style={{display:"flex",alignItems:"center"}}><ColorDot color={sc}/>{val}</span>; }
                        if (c.key==="_profesor") { const prc=profColors[r._profesor]||"#999"; style.borderLeft=`4px solid ${prc}`; style.background=`${prc}15`; val=<span style={{display:"flex",alignItems:"center"}}><ColorDot color={prc}/>{val}</span>; }
                        return <td key={c.key} style={style} title={String(c.fmt?c.fmt(r[c.key]):(r[c.key]??""))}>{val}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {activeTab==="dashboard"&&filteredData.length>50&&(
              <div style={{textAlign:"center",padding:12}}>
                <button onClick={()=>setActiveTab("tabla")} style={{fontSize:12,padding:"6px 16px",background:IPADE.navy,color:"#fff",border:"none",borderRadius:5,cursor:"pointer"}}>
                  Ver todos los {filteredData.length} registros →
                </button>
              </div>
            )}
          </div>
        )}
      </div>

      <footer style={{ background:IPADE.navy, padding:"16px 24px", textAlign:"center", marginTop:20 }}>
        <span style={{ color:IPADE.gold, fontSize:10, letterSpacing:3, textTransform:"uppercase", fontWeight:700 }}>IPADE Business School</span>
        <span style={{ color:"rgba(255,255,255,.3)", fontSize:11, display:"block", marginTop:4 }}>Colorigrama — Sistema de Análisis de Sesiones</span>
      </footer>
    </div>
  );
}
