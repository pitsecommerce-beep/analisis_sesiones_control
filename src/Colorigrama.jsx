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
const SPECIAL_PROGRAMS = ["InCompany", "Enfocados", "SEPOS", "Certificados"];

/* ── Simulación de Programación ── */
// Cursos especiales: cuando la columna "Curso" del Excel inicial coincide con
// alguno de estos valores, la sesión se asigna al profesor configurado en la
// opción "PROP CF" de la plantilla (independiente del reparto 60/40).
const SPECIAL_CF_CURSOS = ["PROP CF", "CF", "PROP CF MJ", "FCF", "PROP CF M"];
const SPECIAL_CF_SET = new Set(SPECIAL_CF_CURSOS.map(s => s.toUpperCase()));
const PROP_CF_LABEL = "PROP CF";
const SIM_PROG_PADRE_OPTIONS = ["MEDEX", "Máster", "Perfeccionamiento", "CA"];
const SIM_CURSO_OPTIONS = ["Contabilidad", "Costos", "Control II", "Riesgos", "Dirección", "Alta Dirección", "Alta Dirección 2"];
// Reparto del curso: la Dupla imparte el 60% inicial, el Titular el 40% final.
const DUPLA_SHARE = 0.6;

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
/* Una sesión cuenta como Virtual si su Modalidad dice "Virtual"
   o si su Sede Sesión dice "Virtual" / "Vir". */
function isVirtualSession(r) {
  const mod = String(r["Modalidad"] ?? "").trim().toLowerCase();
  const sede = String(r._sede ?? r["Sede Sesión"] ?? "").trim().toLowerCase();
  return mod === "virtual" || sede === "virtual" || sede === "vir";
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

/* ── editable text cell (commits on blur / Enter to avoid heavy re-renders) ── */
function EditableCell({ value, onCommit, style }) {
  const [v, setV] = useState(value);
  useEffect(() => { setV(value); }, [value]);
  return (
    <input
      value={v}
      onChange={e => setV(e.target.value)}
      onBlur={() => { if (v !== value) onCommit(v); }}
      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
      style={{ width:"100%", border:"1px solid transparent", background:"transparent", font:"inherit", color:"inherit", padding:"2px 4px", borderRadius:3, ...style }}
      onFocus={e => { e.target.style.border = "1px solid #C8A951"; e.target.style.background = "#fff"; }}
    />
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


/* ══════════════════════════════════════════════════════════
   MINIMAL XLSX ZIP WRITER (shared, no compression / "store")
   ══════════════════════════════════════════════════════════ */
function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function buildZip(fileEntries) {
  const enc = new TextEncoder();
  const entries = Object.entries(fileEntries).map(([name, content]) => ({
    name: enc.encode(name), data: enc.encode(content),
  }));
  const parts = [], centralDir = [];
  let offset = 0;
  entries.forEach(({ name, data }) => {
    const header = new Uint8Array(30 + name.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x04034b50, true); view.setUint16(4, 20, true);
    view.setUint16(6, 0, true); view.setUint16(8, 0, true);
    view.setUint16(10, 0, true); view.setUint16(12, 0, true);
    const crc = crc32(data);
    view.setUint32(14, crc, true); view.setUint32(18, data.length, true);
    view.setUint32(22, data.length, true); view.setUint16(26, name.length, true);
    view.setUint16(28, 0, true); header.set(name, 30);
    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true); cv.setUint16(10, 0, true); cv.setUint16(12, 0, true);
    cv.setUint16(14, 0, true); cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true); cv.setUint16(28, name.length, true);
    cv.setUint16(30, 0, true); cv.setUint16(32, 0, true); cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true); cv.setUint32(38, 0, true); cv.setUint32(42, offset, true);
    cd.set(name, 46); centralDir.push(cd);
    parts.push(header, data); offset += header.length + data.length;
  });
  const cdOffset = offset;
  let cdSize = 0;
  centralDir.forEach(cd => { parts.push(cd); cdSize += cd.length; });
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(4, 0, true); ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true); ev.setUint16(10, entries.length, true);
  ev.setUint32(12, cdSize, true); ev.setUint32(16, cdOffset, true); ev.setUint16(20, 0, true);
  parts.push(eocd);
  const total = parts.reduce((a, p) => a + p.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  parts.forEach(p => { out.set(p, pos); pos += p.length; });
  return out;
}

/* ══════════════════════════════════════════════════════════
   PLANTILLA DE TITULARIDADES — genera un .xlsx con menús
   desplegables (data validation) para configurar la simulación.
   ══════════════════════════════════════════════════════════ */
function generateTemplateBlob(programIds, professors, existingConfig) {
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colName = c => { let s = "", n = c; while (n >= 0) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } return s; };

  const headers = ["Programa Padre", "Curso", "Programa", "Titular", "Dupla"];
  // Opciones del desplegable de "Programa": todos los IDs únicos + la etiqueta especial PROP CF
  const programaOpts = [...programIds, PROP_CF_LABEL];
  const profOpts = [...professors, PROP_CF_LABEL];

  // Filas pre-cargadas: una por cada programa del Excel inicial, más una fila PROP CF.
  let rows;
  if (existingConfig && existingConfig.length) {
    rows = existingConfig.map(c => [c.progPadre || "", c.curso || "", c.programa || "", c.titular || "", c.dupla || ""]);
  } else {
    rows = programaOpts.map(id => ["", "", id, "", ""]);
  }
  const MAX_ROWS = Math.max(rows.length + 50, 300); // rango amplio para validaciones

  /* ── Hoja 1: Titularidades ── */
  const cell = (col, rowNum, val, styleAttr = "") => {
    if (val === "" || val == null) return "";
    return `<c r="${colName(col)}${rowNum}" t="inlineStr"${styleAttr}><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
  };
  let sd1 = `<row r="1">${headers.map((h, i) => cell(i, 1, h, ' s="1"')).join("")}</row>`;
  rows.forEach((r, ri) => {
    const rn = ri + 2;
    sd1 += `<row r="${rn}">${r.map((v, ci) => cell(ci, rn, v)).join("")}</row>`;
  });

  const dvs = [
    { sq: `A2:A${MAX_ROWS}`, f: "progPadreList" },
    { sq: `B2:B${MAX_ROWS}`, f: "cursoList" },
    { sq: `C2:C${MAX_ROWS}`, f: "programaList" },
    { sq: `D2:D${MAX_ROWS}`, f: "profList" },
    { sq: `E2:E${MAX_ROWS}`, f: "profList" },
  ].map(d => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${d.sq}"><formula1>${d.f}</formula1></dataValidation>`).join("");

  const colsXml1 = [26, 18, 28, 14, 14].map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
  const sheet1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${colsXml1}</cols><sheetData>${sd1}</sheetData><dataValidations count="5">${dvs}</dataValidations></worksheet>`;

  /* ── Hoja 2: Listas (fuente de los desplegables) ── */
  const listCols = [SIM_PROG_PADRE_OPTIONS, SIM_CURSO_OPTIONS, programaOpts, profOpts];
  const listHeaders = ["Programa Padre", "Curso", "Programa", "Profesores"];
  const maxLen = Math.max(...listCols.map(c => c.length));
  let sd2 = `<row r="1">${listHeaders.map((h, i) => cell(i, 1, h, ' s="1"')).join("")}</row>`;
  for (let ri = 0; ri < maxLen; ri++) {
    const rn = ri + 2;
    const cells = listCols.map((col, ci) => cell(ci, rn, col[ri] ?? "")).join("");
    if (cells) sd2 += `<row r="${rn}">${cells}</row>`;
  }
  const sheet2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sd2}</sheetData></worksheet>`;

  /* ── Defined names (rangos de las listas) ── */
  const defNames = `<definedNames>
    <definedName name="progPadreList">Listas!$A$2:$A$${SIM_PROG_PADRE_OPTIONS.length + 1}</definedName>
    <definedName name="cursoList">Listas!$B$2:$B$${SIM_CURSO_OPTIONS.length + 1}</definedName>
    <definedName name="programaList">Listas!$C$2:$C$${programaOpts.length + 1}</definedName>
    <definedName name="profList">Listas!$D$2:$D$${profOpts.length + 1}</definedName>
  </definedNames>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Titularidades" sheetId="1" r:id="rId1"/><sheet name="Listas" sheetId="2" r:id="rId2"/></sheets>${defNames}</workbook>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color rgb="FF1B2A4A"/><name val="Calibri"/></font><font><sz val="11"/><b/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1B2A4A"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;

  const files = {
    "[Content_Types].xml": contentTypes,
    "_rels/.rels": rels,
    "xl/_rels/workbook.xml.rels": xlRels,
    "xl/workbook.xml": workbook,
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": sheet1,
    "xl/worksheets/sheet2.xml": sheet2,
  };
  return new Blob([buildZip(files)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

/* Reparte una cantidad de sesiones entre Dupla (60% inicial) y Titular (40% final).
   Devuelve el límite de secuencia (k): secuencia<=k → Dupla, secuencia>k → Titular. */
function duplaCutoff(maxSec) {
  return Math.round((Number(maxSec) || 0) * DUPLA_SHARE);
}

/* ══════════════════════════════════════════════════════════
   EXPORT SIMULACIÓN — .xlsx con formato de tabla (autofiltro)
   y colores por Profesor Final y por ID de programa.
   ══════════════════════════════════════════════════════════ */
const SIM_EXPORT_PALETTE = ["#1B5E20","#B71C1C","#0D47A1","#E65100","#4A148C","#006064","#880E4F","#37474F","#5D4037","#1A237E","#BF360C","#004D40","#33691E","#263238","#F57F17","#311B92","#2E7D32","#C62828","#6A1B9A","#00695C","#283593","#AD1457","#00695C","#4E342E"];
function generateSimExcelBlob(rows, profColors) {
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colName = c => { let s = "", n = c; while (n >= 0) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } return s; };
  const dayName = d => d ? d.toLocaleDateString("es-MX", { weekday: "long" }) : "";

  const COLS = [
    { h: "ID", g: r => r["Id Sesión"] ?? "", w: 10 },
    { h: "Ciclo", g: r => r["Ciclo"] ?? "", w: 8 },
    { h: "Fecha", g: r => fmtDate(r._fecha), w: 12 },
    { h: "Hora inicio", g: r => r._horaInicio || "", w: 10 },
    { h: "Hora fin", g: r => r._horaFin || "", w: 10 },
    { h: "Día", g: r => dayName(r._fecha), w: 11 },
    { h: "Modalidad", g: r => r["Modalidad"] ?? "", w: 14 },
    { h: "Profesor Final", g: r => r._profSim || "", c: "prof", w: 14 },
    { h: "Programa Padre", g: r => r._progPadre || "", w: 16 },
    { h: "Curso", g: r => r._curso || "", w: 12 },
    { h: "Programa", g: r => r._idPrograma || "", c: "prog", w: 22 },
    { h: "Secuencia", g: r => r._secuencia ?? "", w: 9 },
    { h: "Sede Sesión", g: r => r._sede || "", w: 11 },
    { h: "Módulo", g: r => r["Módulo"] ?? "", w: 8 },
    { h: "Tema", g: r => r["Tema"] ?? "", w: 40 },
    { h: "Caso/Nota", g: r => r["Caso/Nota"] ?? "", w: 40 },
  ];
  const nCols = COLS.length, nRows = rows.length;
  const lastCol = colName(nCols - 1), lastRow = nRows + 1;
  const ref = `A1:${lastCol}${lastRow}`;

  /* ── color maps ── */
  const GRAY = "#9E9E9E";
  const profColor = {}; let pi = 0;
  [...new Set(rows.map(r => r._profSim || "").filter(Boolean))].forEach(p => {
    profColor[p] = (profColors && profColors[p]) || SIM_EXPORT_PALETTE[pi++ % SIM_EXPORT_PALETTE.length];
  });
  const progColor = {}; let gi = 0;
  [...new Set(rows.map(r => r._idPrograma).filter(Boolean))].forEach(id => {
    progColor[id] = SIM_EXPORT_PALETTE[gi++ % SIM_EXPORT_PALETTE.length];
  });

  /* ── style registries ── */
  const fonts = [], fontMap = {};
  const addFont = (key, sz, bold, color) => { if (fontMap[key] != null) return fontMap[key]; fonts.push({ sz, bold, color }); return fontMap[key] = fonts.length - 1; };
  const fills = ["none", "gray125"], fillMap = {};
  const addFill = hex => { const k = hex.replace("#", "").toUpperCase(); if (fillMap[k] != null) return fillMap[k]; fills.push(k); return fillMap[k] = fills.length - 1; };
  const xfs = [], xfMap = {};
  const addXf = (f, fl) => { const k = f + "_" + fl; if (xfMap[k] != null) return xfMap[k]; xfs.push({ f, fl }); return xfMap[k] = xfs.length - 1; };

  const baseFont = addFont("base", 10, false, "1B2A4A");
  const hdrFont = addFont("hdr", 11, true, "FFFFFF");
  const hdrFill = addFill("1B2A4A"), evenFill = addFill("F5F3EE"), whiteFill = addFill("FFFFFF");
  const baseXfWhite = addXf(baseFont, whiteFill), baseXfEven = addXf(baseFont, evenFill), hdrXf = addXf(hdrFont, hdrFill);
  const coloredXf = hex => {
    const ch = (hex || GRAY).replace("#", "").toUpperCase();
    return addXf(addFont("b_" + ch, 10, true, ch), addFill(rgbToHex(lighten(hex || GRAY))));
  };

  /* ── sheet rows ── */
  const cellStr = (col, rowNum, val, xf) => `<c r="${colName(col)}${rowNum}" t="inlineStr" s="${xf}"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
  let sheetRows = `<row r="1">${COLS.map((c, ci) => cellStr(ci, 1, c.h, hdrXf)).join("")}</row>`;
  rows.forEach((r, ri) => {
    const rowNum = ri + 2;
    const baseXf = ri % 2 === 0 ? baseXfWhite : baseXfEven;
    sheetRows += `<row r="${rowNum}">` + COLS.map((c, ci) => {
      const val = c.g(r);
      let xf = baseXf;
      if (c.c === "prof" && val) xf = coloredXf(profColor[val]);
      else if (c.c === "prog" && val) xf = coloredXf(progColor[val]);
      return cellStr(ci, rowNum, val, xf);
    }).join("") + `</row>`;
  });

  const colsXml = COLS.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.w}" customWidth="1"/>`).join("");
  const fontsXml = fonts.map(f => `<font><sz val="${f.sz}"/>${f.bold ? "<b/>" : ""}<color rgb="FF${f.color}"/><name val="Calibri"/></font>`).join("");
  const fillsXml = fills.map((f, i) => i === 0 ? `<fill><patternFill patternType="none"/></fill>` : i === 1 ? `<fill><patternFill patternType="gray125"/></fill>` : `<fill><patternFill patternType="solid"><fgColor rgb="FF${f}"/></patternFill></fill>`).join("");
  const xfsXml = xfs.map(x => `<xf fontId="${x.f}" fillId="${x.fl}" borderId="0" numFmtId="0" applyFont="1" applyFill="1"/>`).join("");

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="${fonts.length}">${fontsXml}</fonts><fills count="${fills.length}">${fillsXml}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfsXml}</cellXfs></styleSheet>`;

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cols>${colsXml}</cols><sheetData>${sheetRows}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;

  const tableCols = COLS.map((c, i) => `<tableColumn id="${i + 1}" name="${esc(c.h)}"/>`).join("");
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Colorigrama" displayName="Colorigrama" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/><tableColumns count="${nCols}">${tableCols}</tableColumns><tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>`;

  const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Colorigrama" sheetId="1" r:id="rId1"/></sheets></workbook>`;

  const files = {
    "[Content_Types].xml": contentTypes,
    "_rels/.rels": rels,
    "xl/_rels/workbook.xml.rels": xlRels,
    "xl/workbook.xml": workbook,
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": sheet,
    "xl/worksheets/_rels/sheet1.xml.rels": sheetRels,
    "xl/tables/table1.xml": table,
  };
  return new Blob([buildZip(files)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
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
  const [dateRange, setDateRange] = useState({ from: "", to: "" });
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [exportStatus, setExportStatus] = useState("");
  /* ── Simulación de Programación ── */
  const [simConfig, setSimConfig] = useState([]);   // [{progPadre, curso, programa, titular, dupla}]
  const [simOverrides, setSimOverrides] = useState({}); // { idSesion: profesorEditadoAMano }
  const [simFilters, setSimFilters] = useState({});
  const [simDateRange, setSimDateRange] = useState({ from: "", to: "" });
  const [simSortCol, setSimSortCol] = useState("_fecha");
  const [simSortDir, setSimSortDir] = useState("asc");
  const [simStatus, setSimStatus] = useState("");
  const simFileRef = useRef();

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];

      // Detect merged cells to identify sessions with multiple professors.
      // SheetJS stores merges as [{s:{r,c}, e:{r,c}}] (0-indexed, row 0 = header).
      // When "Id Sesión" (col 0) is vertically merged, each sub-row is a different professor.
      const merges = ws["!merges"] || [];
      const ID_COL = 0; // "Id Sesión" is always column A (index 0)
      // Map: data-row-index (0-based, after header) -> index of the group leader row
      const rowToLeader = {};
      merges.forEach(({ s, e }) => {
        if (s.c <= ID_COL && e.c >= ID_COL && e.r > s.r) {
          const leaderData = s.r - 1; // subtract 1 for header row
          for (let r = s.r; r <= e.r; r++) rowToLeader[r - 1] = leaderData;
        }
      });

      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

      // Group rows that share the same merged "Id Sesión" cell
      const groups = {};
      json.forEach((r, i) => {
        const leader = rowToLeader[i] !== undefined ? rowToLeader[i] : i;
        if (!groups[leader]) groups[leader] = { row: r, profs: [] };
        const p = r["Profesores"] ? String(r["Profesores"]).trim() : "";
        if (p && !groups[leader].profs.includes(p)) groups[leader].profs.push(p);
      });

      const parsed = Object.values(groups).map(({ row: r, profs }) => {
        const profArr = profs.length ? profs : ["(Vacío)"];
        const sede = r["Sede Sesión"] ? String(r["Sede Sesión"]).trim() : "";
        const programa = r["Programa"] ? String(r["Programa"]).trim() : "";
        const grupo = r["Grupo"] != null ? String(r["Grupo"]).trim() : "";
        return {
          ...r,
          _profesoresArr: profArr,
          _profesor: profArr.join(" / "),
          _sede: sede,
          _progPadre: r["Programa Padre"] ? String(r["Programa Padre"]).trim() : "",
          _programa: programa,
          _grupo: grupo,
          // ID por programa = Programa + " " + Sede Sesión + " " + Grupo
          _idPrograma: `${programa} ${sede} ${grupo}`.replace(/\s+/g, " ").trim(),
          _curso: r["Curso"] != null ? String(r["Curso"]).trim() : "",
          _secuencia: Number(r["Secuencia"]) || 0,
          _fecha: excelDateToJS(r["Fecha sesión (Día/Mes/Año)"]),
          _horaInicio: excelTimeToStr(r["Hora inicio"]),
          _horaFin: excelTimeToStr(r["Hora fin"]),
        };
      });

      parsed.sort((a,b) => (a._fecha||0) - (b._fecha||0));
      const esc2 = { ...SEDE_PALETTE_DEFAULT }, epc = { ...PROF_PALETTE_DEFAULT }, eprc = { ...PROG_PALETTE_DEFAULT };
      const fb = ["#37474F","#5D4037","#1A237E","#BF360C","#004D40","#33691E","#263238","#F57F17","#880E4F","#311B92"];
      let ci2 = 0;
      parsed.forEach(r => {
        if (r._sede && !esc2[r._sede]) esc2[r._sede] = fb[ci2++ % fb.length];
        r._profesoresArr.forEach(p => { if (p && p !== "(Vacío)" && !epc[p]) epc[p] = fb[ci2++ % fb.length]; });
        if (r._progPadre && !eprc[r._progPadre]) eprc[r._progPadre] = fb[ci2++ % fb.length];
      });
      setSedeColors(esc2); setProfColors(epc); setProgColors(eprc);
      setRawData(parsed); setFilters({}); setDateRange({ from: "", to: "" });
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

  // Columnas de la tabla detallada simulada
  const SIM_COLS = useMemo(() => [
    { key:"Ciclo", label:"Ciclo" }, { key:"_progPadre", label:"Programa Padre" },
    { key:"_programa", label:"Programa" }, { key:"_idPrograma", label:"ID Programa" },
    { key:"_sede", label:"Sede" }, { key:"Modalidad", label:"Modalidad" },
    { key:"_grupo", label:"Grupo" }, { key:"_curso", label:"Curso" },
    { key:"_secuencia", label:"Secuencia" }, { key:"_profSim", label:"Profesor (Simulado)" },
    { key:"_rolSim", label:"Rol" }, { key:"_fecha", label:"Fecha Sesión", fmt:fmtDate },
    { key:"_horaInicio", label:"Hora inicio" }, { key:"_horaFin", label:"Hora fin" },
    { key:"Tema", label:"Tema" }, { key:"Caso/Nota", label:"Caso/Nota" },
  ], []);

  const filterableValues = useMemo(() => {
    if (!rawData) return {};
    const out = {};
    DISPLAY_COLS.forEach(c => {
      if (["_fecha","_horaInicio","_horaFin","Tema","Caso/Nota","Secuencia","Módulo"].includes(c.key)) return;
      if (c.key === "_profesor") {
        // List individual professors, not the combined string
        out[c.key] = [...new Set(rawData.flatMap(r => r._profesoresArr || [r._profesor]))].sort();
      } else {
        out[c.key] = [...new Set(rawData.map(r => String(r[c.key] ?? "")))].sort();
      }
    });
    return out;
  }, [rawData, DISPLAY_COLS]);

  const dateBounds = useMemo(() => {
    if (!rawData) return { min: "", max: "" };
    const ts = rawData.map(r => r._fecha?.getTime()).filter(Boolean);
    if (!ts.length) return { min: "", max: "" };
    const toISO = t => new Date(t).toISOString().slice(0, 10);
    return { min: toISO(Math.min(...ts)), max: toISO(Math.max(...ts)) };
  }, [rawData]);

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
      // For professor filter, include the session if ANY of its professors is selected
      if (k === "_profesor") return (r._profesoresArr || [r._profesor]).some(p => sel.includes(p));
      return sel.includes(String(r[k] ?? ""));
    }));
    if (dateRange.from || dateRange.to) {
      const from = dateRange.from ? new Date(dateRange.from).getTime() : -Infinity;
      const to   = dateRange.to   ? new Date(dateRange.to + "T23:59:59").getTime() : Infinity;
      d = d.filter(r => r._fecha && r._fecha.getTime() >= from && r._fecha.getTime() <= to);
    }
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
    // Virtual: sesiones cuya Modalidad es "Virtual" o cuya Sede Sesión es "Virtual"/"Vir"
    const virtual = filteredData.filter(isVirtualSession).length;
    const local = filteredData.filter(r => !isVirtualSession(r) && r._sede === "MEX").length;
    const foranea = total - local - virtual;
    const byProf = {}, specialByProf = {}, bySede = {};
    filteredData.forEach(r => {
      (r._profesoresArr || [r._profesor]).forEach(p => { byProf[p] = (byProf[p]||0)+1; });
      bySede[r._sede] = (bySede[r._sede]||0)+1;
    });
    filteredData.filter(r => SPECIAL_PROGRAMS.includes(r._progPadre)).forEach(r => {
      (r._profesoresArr || [r._profesor]).forEach(p => { specialByProf[p] = (specialByProf[p]||0)+1; });
    });
    return { total, local, foranea, virtual, byProf, specialByProf, bySede };
  }, [filteredData]);

  const sedeChartData = useMemo(() => metrics ? Object.entries(metrics.bySede).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);
  const profChartData = useMemo(() => metrics ? Object.entries(metrics.byProf).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);
  const specialChartData = useMemo(() => metrics ? Object.entries(metrics.specialByProf).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value) : [], [metrics]);

  /* ══════════ SIMULACIÓN DE PROGRAMACIÓN ══════════ */
  // IDs únicos de programa (Programa + " " + Sede Sesión + " " + Grupo) del Excel inicial
  const programIds = useMemo(() => {
    if (!rawData) return [];
    return [...new Set(rawData.map(r => r._idPrograma).filter(Boolean))].sort();
  }, [rawData]);

  // Lista de profesores únicos del Excel inicial (códigos individuales)
  const professorsAll = useMemo(() => {
    if (!rawData) return [];
    return [...new Set(rawData.flatMap(r => r._profesoresArr || []).filter(p => p && p !== "(Vacío)"))].sort();
  }, [rawData]);

  // Secuencia máxima (= número de sesiones) por ID de programa
  const maxSecByProg = useMemo(() => {
    const m = {};
    (rawData || []).forEach(r => {
      const id = r._idPrograma;
      if (!id) return;
      if (!(id in m) || r._secuencia > m[id]) m[id] = r._secuencia;
    });
    return m;
  }, [rawData]);

  // Mapa de configuración por ID de programa + profesor especial PROP CF
  const simLookup = useMemo(() => {
    const byProg = {};
    let propCfProf = "";
    simConfig.forEach(c => {
      const id = String(c.programa || "").trim();
      if (!id) return;
      if (id.toUpperCase() === PROP_CF_LABEL) { propCfProf = String(c.titular || c.dupla || "").trim(); return; }
      byProg[id] = c;
    });
    return { byProg, propCfProf };
  }, [simConfig]);

  // Resuelve el profesor simulado de una sesión según el reparto 60/40 (o el caso PROP CF).
  // Si no hay profesor asignado para ese programa, devuelve "" (queda en blanco) para
  // señalar que todavía hace falta asignarle profesor.
  const resolveSimProf = useCallback((r) => {
    if (SPECIAL_CF_SET.has(String(r._curso || "").toUpperCase())) {
      return { prof: simLookup.propCfProf || "", rol: "PROP CF" };
    }
    const cfg = simLookup.byProg[r._idPrograma];
    if (!cfg) return { prof: "", rol: "" };
    const k = duplaCutoff(maxSecByProg[r._idPrograma] || 0);
    const isDupla = r._secuencia <= k;
    const prof = (isDupla ? cfg.dupla : cfg.titular) || "";
    return { prof, rol: isDupla ? "Dupla" : "Titular" };
  }, [simLookup, maxSecByProg]);

  // Sesiones con profesor simulado (respeta ediciones manuales)
  const simulatedAll = useMemo(() => {
    if (!rawData) return [];
    return rawData.map(r => {
      const { prof, rol } = resolveSimProf(r);
      const idSes = r["Id Sesión"];
      const override = simOverrides[idSes];
      return { ...r, _profSim: override != null ? override : prof, _rolSim: override != null ? "Manual" : rol };
    });
  }, [rawData, resolveSimProf, simOverrides]);

  // Valores filtrables para la pestaña de simulación
  const simFilterableValues = useMemo(() => {
    if (!simulatedAll.length) return {};
    const cols = ["Ciclo", "_progPadre", "_programa", "_sede", "Modalidad", "_grupo", "_profSim"];
    const out = {};
    cols.forEach(k => { out[k] = [...new Set(simulatedAll.map(r => String(r[k] ?? "")))].sort(); });
    return out;
  }, [simulatedAll]);

  useEffect(() => {
    if (!simulatedAll.length) return;
    const init = {};
    Object.entries(simFilterableValues).forEach(([k, v]) => { init[k] = [...v]; });
    setSimFilters(init);
  }, [simulatedAll.length, simFilterableValues]);

  // Sesiones simuladas filtradas + ordenadas (default: más antiguo → más nuevo)
  const simFiltered = useMemo(() => {
    let d = simulatedAll.filter(r => Object.entries(simFilters).every(([k, sel]) => {
      if (!sel || sel.length === 0) return false;
      return sel.includes(String(r[k] ?? ""));
    }));
    if (simDateRange.from || simDateRange.to) {
      const from = simDateRange.from ? new Date(simDateRange.from).getTime() : -Infinity;
      const to = simDateRange.to ? new Date(simDateRange.to + "T23:59:59").getTime() : Infinity;
      d = d.filter(r => r._fecha && r._fecha.getTime() >= from && r._fecha.getTime() <= to);
    }
    const col = simSortCol || "_fecha";
    d = [...d].sort((a, b) => {
      let va = a[col] ?? "", vb = b[col] ?? "";
      if (col === "_fecha") { va = va || new Date(0); vb = vb || new Date(0); return simSortDir === "asc" ? va - vb : vb - va; }
      va = String(va); vb = String(vb);
      return simSortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    return d;
  }, [simulatedAll, simFilters, simDateRange, simSortCol, simSortDir]);

  // Métricas / datos de gráficos de la simulación
  const simCharts = useMemo(() => {
    const byProf = {}, byProfSede = {}, byProfPrograma = {}, sedeSet = new Set(), padreSet = new Set();
    let local = 0, foranea = 0, virtual = 0;
    simFiltered.forEach(r => {
      const p = r._profSim || "(Sin asignar)";
      byProf[p] = (byProf[p] || 0) + 1;
      const sede = r._sede || "(Vacío)";
      sedeSet.add(sede);
      (byProfSede[p] = byProfSede[p] || {})[sede] = (byProfSede[p][sede] || 0) + 1;
      const padre = r._progPadre || "(Vacío)";
      padreSet.add(padre);
      (byProfPrograma[p] = byProfPrograma[p] || {})[padre] = (byProfPrograma[p][padre] || 0) + 1;
      if (isVirtualSession(r)) virtual++;
      else if (r._sede === "MEX") local++;
      else foranea++;
    });
    const profOrder = Object.entries(byProf).sort((a, b) => b[1] - a[1]).map(([p]) => p);
    const profData = profOrder.map(name => ({ name, value: byProf[name] }));
    const sedeKeys = [...sedeSet].sort();
    const padreKeys = [...padreSet].sort();
    const profSedeData = profOrder.map(p => ({ name: p, ...byProfSede[p] }));
    const profProgramaData = profOrder.map(p => ({ name: p, ...byProfPrograma[p] }));
    return { profData, profSedeData, profProgramaData, sedeKeys, padreKeys, local, foranea, virtual, total: simFiltered.length };
  }, [simFiltered]);

  const simDateBounds = useMemo(() => {
    if (!simulatedAll.length) return { min: "", max: "" };
    const ts = simulatedAll.map(r => r._fecha?.getTime()).filter(Boolean);
    if (!ts.length) return { min: "", max: "" };
    const toISO = t => new Date(t).toISOString().slice(0, 10);
    return { min: toISO(Math.min(...ts)), max: toISO(Math.max(...ts)) };
  }, [simulatedAll]);

  /* Descarga la plantilla de titularidades (.xlsx con desplegables).
     Solo incluye los programas visibles según los filtros activos de la pestaña. */
  const downloadTemplate = useCallback(() => {
    if (!programIds.length) { setSimStatus("⚠️ Carga primero el Excel inicial"); return; }
    // IDs de programa visibles con los filtros actuales
    const visibleIds = [...new Set(simFiltered.map(r => r._idPrograma).filter(Boolean))].sort();
    // Conserva las asignaciones existentes de la configuración cargada
    const cfgByProg = {};
    simConfig.forEach(c => { const id = String(c.programa || "").trim(); if (id) cfgByProg[id] = c; });
    const rows = visibleIds.map(id => cfgByProg[id] || { progPadre: "", curso: "", programa: id, titular: "", dupla: "" });
    // Siempre incluir la fila especial PROP CF
    const propRow = simConfig.find(c => String(c.programa || "").trim().toUpperCase() === PROP_CF_LABEL)
      || { progPadre: "", curso: "", programa: PROP_CF_LABEL, titular: "", dupla: "" };
    rows.push(propRow);
    setSimStatus("Generando plantilla...");
    try {
      const blob = generateTemplateBlob(visibleIds, professorsAll, rows);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = "Plantilla_Titularidades.xlsx";
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setSimStatus(`✅ Plantilla descargada (${visibleIds.length} programas visibles) — llénala y vuelve a cargarla`);
    } catch (e) { console.error(e); setSimStatus("⚠️ Error al generar la plantilla"); }
    setTimeout(() => setSimStatus(""), 4000);
  }, [programIds, professorsAll, simConfig, simFiltered]);

  /* Descarga "Colorigrama <fecha>" con lo que se ve en la tabla de detalle simulada,
     en formato de tabla y coloreado por Profesor Final y por ID de programa. */
  const exportSimExcel = useCallback(() => {
    if (!simFiltered.length) { setSimStatus("⚠️ Sin sesiones para exportar"); setTimeout(() => setSimStatus(""), 3000); return; }
    setSimStatus("Generando Colorigrama...");
    try {
      const blob = generateSimExcelBlob(simFiltered, profColors);
      const url = URL.createObjectURL(blob);
      const today = new Date().toLocaleDateString("es-MX").replace(/\//g, "-");
      const a = document.createElement("a");
      a.href = url; a.download = `Colorigrama ${today}.xlsx`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setSimStatus(`✅ Colorigrama descargado (${simFiltered.length} sesiones)`);
    } catch (e) { console.error(e); setSimStatus("⚠️ Error al generar el archivo"); }
    setTimeout(() => setSimStatus(""), 4000);
  }, [simFiltered, profColors]);

  /* Carga la plantilla de titularidades llenada */
  const handleTemplateUpload = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSimStatus("Leyendo plantilla...");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["Titularidades"] || wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const norm = s => String(s ?? "").trim();
        const cfg = json.map(r => ({
          progPadre: norm(r["Programa Padre"]),
          curso: norm(r["Curso"]),
          programa: norm(r["Programa"]),
          titular: norm(r["Titular"]),
          dupla: norm(r["Dupla"]),
        })).filter(c => c.programa);
        setSimConfig(cfg);
        setSimOverrides({});
        setSimStatus(`✅ ${cfg.length} titularidades cargadas`);
      } catch (err) { console.error(err); setSimStatus("⚠️ No se pudo leer la plantilla"); }
      setTimeout(() => setSimStatus(""), 4000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }, []);

  const updateSimConfigRow = (idx, field, value) => {
    setSimConfig(prev => prev.map((c, i) => i === idx ? { ...c, [field]: value } : c));
  };
  const addSimConfigRow = () => setSimConfig(prev => [...prev, { progPadre: "", curso: "", programa: "", titular: "", dupla: "" }]);
  const removeSimConfigRow = (idx) => setSimConfig(prev => prev.filter((_, i) => i !== idx));
  const handleSimSort = col => {
    if (simSortCol === col) setSimSortDir(simSortDir === "asc" ? "desc" : "asc");
    else { setSimSortCol(col); setSimSortDir("asc"); }
  };

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
        {[["dashboard","📊 Dashboard"],["tabla","📋 Tabla Detallada"],["simulacion","🧩 Simulación de Programación"]].map(([id,label]) => (
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
      {activeTab!=="simulacion" && (<>
      <div style={{ background:"#fff", borderBottom:"1px solid #eee", padding:"10px 24px", display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, fontWeight:700, marginRight:8 }}>Filtros:</span>
        {Object.entries(filterableValues).map(([k,vals]) => {
          const colDef = DISPLAY_COLS.find(c=>c.key===k);
          return <FilterDropdown key={k} values={vals} selected={filters[k]||[]} onChange={sel=>setFilters(p=>({...p,[k]:sel}))} label={colDef?.label||k} />;
        })}
        {/* Date range picker */}
        <div style={{ display:"inline-flex", alignItems:"center", gap:4, border:"1px solid #ccc", borderRadius:4, padding:"2px 8px", background: dateRange.from||dateRange.to ? IPADE.gold : "#e8e4dd" }}>
          <span style={{ fontSize:11, fontWeight:600, color: dateRange.from||dateRange.to ? "#fff" : "#333", whiteSpace:"nowrap" }}>📅 Fecha</span>
          <input
            type="date" value={dateRange.from} min={dateBounds.min} max={dateRange.to||dateBounds.max}
            onChange={e => setDateRange(p => ({ ...p, from: e.target.value }))}
            style={{ fontSize:11, border:"none", background:"transparent", cursor:"pointer", color: dateRange.from ? "#1B2A4A" : "#888" }}
          />
          <span style={{ fontSize:11, color:"#666" }}>—</span>
          <input
            type="date" value={dateRange.to} min={dateRange.from||dateBounds.min} max={dateBounds.max}
            onChange={e => setDateRange(p => ({ ...p, to: e.target.value }))}
            style={{ fontSize:11, border:"none", background:"transparent", cursor:"pointer", color: dateRange.to ? "#1B2A4A" : "#888" }}
          />
          {(dateRange.from||dateRange.to) && (
            <span onClick={() => setDateRange({ from:"", to:"" })} style={{ fontSize:12, cursor:"pointer", color:"#fff", fontWeight:700, marginLeft:2 }} title="Quitar filtro de fecha">✕</span>
          )}
        </div>
        <button onClick={()=>{const init={};Object.entries(filterableValues).forEach(([k,v])=>{init[k]=[...v];});setFilters(init);setDateRange({from:"",to:"" });}} style={{ fontSize:11, padding:"3px 10px", background:IPADE.navy, color:"#fff", border:"none", borderRadius:4, cursor:"pointer", marginLeft:8 }}>
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
      </>)}

      {/* MAIN CONTENT */}
      <div style={{ padding:24 }}>
        {activeTab === "dashboard" && metrics && (
          <>
            <div style={{ display:"flex", gap:14, flexWrap:"wrap", marginBottom:24 }}>
              <KPI label="Total Sesiones" value={metrics.total} color={IPADE.navy} />
              <KPI label="Sesiones Locales (MEX)" value={metrics.local} color="#2E7D32" sub={`${((metrics.local/metrics.total)*100).toFixed(1)}%`} />
              <KPI label="Sesiones Foráneas" value={metrics.foranea} color="#C62828" sub={`${((metrics.foranea/metrics.total)*100).toFixed(1)}%`} />
              <KPI label="Sesiones Virtuales" value={metrics.virtual} color="#455A64" sub={`${((metrics.virtual/metrics.total)*100).toFixed(1)}%`} />
              <KPI label="Sesiones Especiales" value={Object.values(metrics.specialByProf).reduce((a,b)=>a+b,0)} color={IPADE.gold} sub="InCompany + Enfocados + SEPOS + Certificados" />
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
                <p style={{ fontSize:11, color:"#888", marginBottom:10 }}>InCompany, Enfocados, SEPOS, Certificados</p>
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
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Local vs. Foráneas vs. Virtual</h3>
                <ResponsiveContainer width="100%" height={280}>
                  <PieChart margin={{top:20, right:40, bottom:0, left:40}}><Pie data={[{name:"Local (MEX)",value:metrics.local},{name:"Foráneas",value:metrics.foranea},{name:"Virtual",value:metrics.virtual}].filter(d=>d.value>0)} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={55} outerRadius={90} label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={true} style={{fontSize:11}}>
                    {[{name:"Local (MEX)",value:metrics.local,fill:"#2E7D32"},{name:"Foráneas",value:metrics.foranea,fill:"#C62828"},{name:"Virtual",value:metrics.virtual,fill:"#455A64"}].filter(d=>d.value>0).map(d=>(<Cell key={d.name} fill={d.fill}/>))}
                  </Pie><Tooltip/><Legend/></PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </>
        )}

        {/* ══════════ SIMULACIÓN DE PROGRAMACIÓN ══════════ */}
        {activeTab==="simulacion" && (
          <>
            {/* ─ Sección Titularidades ─ */}
            <div style={{ background:"#fff", borderRadius:12, boxShadow:"0 2px 8px rgba(0,0,0,.06)", padding:20, marginBottom:24 }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10, marginBottom:6 }}>
                <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:18, margin:0 }}>Titularidades</h3>
                <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                  {simStatus && <span style={{ fontSize:11, color:IPADE.gold, fontWeight:600 }}>{simStatus}</span>}
                  <button onClick={downloadTemplate} style={{ fontSize:11, padding:"5px 14px", background:"#217346", color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontWeight:600 }}>📥 Descargar plantilla</button>
                  <button onClick={()=>simFileRef.current?.click()} style={{ fontSize:11, padding:"5px 14px", background:IPADE.accent1, color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontWeight:600 }}>📤 Cargar plantilla</button>
                  <input ref={simFileRef} type="file" accept=".xlsx,.xls" onChange={handleTemplateUpload} style={{ display:"none" }} />
                  <button onClick={addSimConfigRow} style={{ fontSize:11, padding:"5px 12px", background:IPADE.navy, color:"#fff", border:"none", borderRadius:5, cursor:"pointer" }}>+ Fila</button>
                </div>
              </div>
              <p style={{ fontSize:11.5, color:"#666", marginBottom:12, lineHeight:1.5 }}>
                Simula la programación de profesores por programa. La <strong>Dupla</strong> imparte el 60% inicial de las sesiones y el <strong>Titular</strong> el 40% final
                (según la <em>Secuencia</em> de cada ID de programa <code>Programa + Sede + Grupo</code>). Las filas con <strong>Programa = "PROP CF"</strong> definen al profesor
                asignado a las sesiones cuyo <em>Curso</em> sea {SPECIAL_CF_CURSOS.map(c=>`"${c}"`).join(", ")}. Descarga la plantilla, llénala y cárgala, o edita directamente abajo.
              </p>
              <div style={{ overflowX:"auto", maxHeight:360, overflowY:"auto", border:"1px solid #eee", borderRadius:8 }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                  <thead>
                    <tr>{["Programa Padre","Curso","Programa (ID)","Titular (40% final)","Dupla (60% inicial)","Sesiones",""].map(h=>(
                      <th key={h} style={{ background:IPADE.navy, color:"#fff", padding:"8px", textAlign:"left", fontSize:11, fontWeight:600, whiteSpace:"nowrap", position:"sticky", top:0, zIndex:5 }}>{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {simConfig.length===0 && (
                      <tr><td colSpan={7} style={{ padding:20, textAlign:"center", color:"#999" }}>Sin titularidades. Descarga la plantilla y cárgala, o agrega filas con “+ Fila”.</td></tr>
                    )}
                    {[...simConfig.map((c,idx)=>({c,idx}))].sort((a,b)=>String(a.c.curso||"").localeCompare(String(b.c.curso||""))).map(({c,idx})=>{
                      const N = maxSecByProg[String(c.programa||"").trim()];
                      const k = N!=null?duplaCutoff(N):null;
                      return (
                      <tr key={idx} style={{ background:idx%2===0?"#fff":"#f9f8f5" }}>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea" }}>
                          <select value={c.progPadre||""} onChange={e=>updateSimConfigRow(idx,"progPadre",e.target.value)} style={{ fontSize:11, border:"1px solid #ddd", borderRadius:4, padding:"3px 4px", width:"100%" }}>
                            <option value="">—</option>{SIM_PROG_PADRE_OPTIONS.map(o=><option key={o} value={o}>{o}</option>)}
                          </select>
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea" }}>
                          <select value={c.curso||""} onChange={e=>updateSimConfigRow(idx,"curso",e.target.value)} style={{ fontSize:11, border:"1px solid #ddd", borderRadius:4, padding:"3px 4px", width:"100%" }}>
                            <option value="">—</option>{SIM_CURSO_OPTIONS.map(o=><option key={o} value={o}>{o}</option>)}
                          </select>
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea", minWidth:160 }}>
                          <select value={c.programa||""} onChange={e=>updateSimConfigRow(idx,"programa",e.target.value)} style={{ fontSize:11, border:"1px solid #ddd", borderRadius:4, padding:"3px 4px", width:"100%" }}>
                            <option value="">—</option>
                            <option value={PROP_CF_LABEL}>{PROP_CF_LABEL}</option>
                            {programIds.map(o=><option key={o} value={o}>{o}</option>)}
                          </select>
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea" }}>
                          <input list="sim-prof-list" value={c.titular||""} onChange={e=>updateSimConfigRow(idx,"titular",e.target.value)} style={{ fontSize:11, border:"1px solid #ddd", borderRadius:4, padding:"3px 4px", width:90 }} />
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea" }}>
                          <input list="sim-prof-list" value={c.dupla||""} onChange={e=>updateSimConfigRow(idx,"dupla",e.target.value)} style={{ fontSize:11, border:"1px solid #ddd", borderRadius:4, padding:"3px 4px", width:90 }} />
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea", fontSize:11, color:"#666", whiteSpace:"nowrap" }}>
                          {String(c.programa||"").toUpperCase()===PROP_CF_LABEL ? "—" : N!=null ? `${N} ses · D 1-${k} · T ${k+1}-${N}` : <span style={{color:"#C62828"}}>no encontrado</span>}
                        </td>
                        <td style={{ padding:"3px 6px", borderBottom:"1px solid #f0efea", textAlign:"center" }}>
                          <button onClick={()=>removeSimConfigRow(idx)} style={{ fontSize:13, border:"none", background:"transparent", color:"#C62828", cursor:"pointer" }} title="Eliminar fila">✕</button>
                        </td>
                      </tr>
                    );})}
                  </tbody>
                </table>
                <datalist id="sim-prof-list">
                  <option value={PROP_CF_LABEL} />
                  {professorsAll.map(p=><option key={p} value={p} />)}
                </datalist>
              </div>
            </div>

            {/* ─ Gráficos de simulación ─ */}
            {simConfig.length>0 && (
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(340px, 1fr))", gap:20, marginBottom:24 }}>
                <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                  <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones por Profesor</h3>
                  <ResponsiveContainer width="100%" height={Math.max(260, simCharts.profData.length*26)}>
                    <BarChart data={simCharts.profData} layout="vertical" margin={{left:50,right:20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis type="number" style={{fontSize:11}}/><YAxis type="category" dataKey="name" width={70} style={{fontSize:11}}/><Tooltip/>
                      <Bar dataKey="value" radius={[0,6,6,0]}>{simCharts.profData.map(d=>(<Cell key={d.name} fill={profColors[d.name]||IPADE.accent1}/>))}</Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                  <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones por Profesor por Sede</h3>
                  <ResponsiveContainer width="100%" height={Math.max(260, simCharts.profSedeData.length*26)}>
                    <BarChart data={simCharts.profSedeData} layout="vertical" margin={{left:50,right:20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis type="number" style={{fontSize:11}}/><YAxis type="category" dataKey="name" width={70} style={{fontSize:11}}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/>
                      {simCharts.sedeKeys.map(s=>(<Bar key={s} dataKey={s} stackId="a" fill={sedeColors[s]||"#999"}/>))}
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                  <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Sesiones por Profesor por Programa</h3>
                  <ResponsiveContainer width="100%" height={Math.max(260, simCharts.profProgramaData.length*26)}>
                    <BarChart data={simCharts.profProgramaData} layout="vertical" margin={{left:50,right:20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#eee"/><XAxis type="number" style={{fontSize:11}}/><YAxis type="category" dataKey="name" width={70} style={{fontSize:11}}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/>
                      {simCharts.padreKeys.map(p=>(<Bar key={p} dataKey={p} stackId="a" fill={progColors[p]||"#999"}/>))}
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background:"#fff", borderRadius:12, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                  <h3 style={{ fontFamily:"'DM Serif Display', serif", fontSize:16, marginBottom:14 }}>Foráneas vs. Local (MEX) vs. Virtual</h3>
                  <ResponsiveContainer width="100%" height={280}>
                    <PieChart margin={{top:20, right:40, bottom:0, left:40}}>
                      <Pie data={[{name:"Local (MEX)",value:simCharts.local,fill:"#2E7D32"},{name:"Foráneas",value:simCharts.foranea,fill:"#C62828"},{name:"Virtual",value:simCharts.virtual,fill:"#455A64"}].filter(d=>d.value>0)} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={55} outerRadius={90} label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={true} style={{fontSize:11}}>
                        {[{name:"Local (MEX)",value:simCharts.local,fill:"#2E7D32"},{name:"Foráneas",value:simCharts.foranea,fill:"#C62828"},{name:"Virtual",value:simCharts.virtual,fill:"#455A64"}].filter(d=>d.value>0).map(d=>(<Cell key={d.name} fill={d.fill}/>))}
                      </Pie><Tooltip/><Legend/>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {/* ─ Filtros de la tabla simulada ─ */}
            <div style={{ background:"#fff", borderRadius:12, boxShadow:"0 2px 8px rgba(0,0,0,.06)", padding:"10px 16px", marginBottom:14, display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
              <span style={{ fontSize:12, fontWeight:700, marginRight:8 }}>Filtros:</span>
              {Object.entries(simFilterableValues).map(([k,vals])=>{
                const labels={ Ciclo:"Ciclo", _progPadre:"Programa Padre", _programa:"Programa", _sede:"Sede", Modalidad:"Modalidad", _grupo:"Grupo", _profSim:"Profesor (sim.)" };
                return <FilterDropdown key={k} values={vals} selected={simFilters[k]||[]} onChange={sel=>setSimFilters(p=>({...p,[k]:sel}))} label={labels[k]||k} />;
              })}
              <div style={{ display:"inline-flex", alignItems:"center", gap:4, border:"1px solid #ccc", borderRadius:4, padding:"2px 8px", background: simDateRange.from||simDateRange.to ? IPADE.gold : "#e8e4dd" }}>
                <span style={{ fontSize:11, fontWeight:600, color: simDateRange.from||simDateRange.to ? "#fff" : "#333", whiteSpace:"nowrap" }}>📅 Fecha</span>
                <input type="date" value={simDateRange.from} min={simDateBounds.min} max={simDateRange.to||simDateBounds.max} onChange={e=>setSimDateRange(p=>({...p,from:e.target.value}))} style={{ fontSize:11, border:"none", background:"transparent", cursor:"pointer" }} />
                <span style={{ fontSize:11, color:"#666" }}>—</span>
                <input type="date" value={simDateRange.to} min={simDateRange.from||simDateBounds.min} max={simDateBounds.max} onChange={e=>setSimDateRange(p=>({...p,to:e.target.value}))} style={{ fontSize:11, border:"none", background:"transparent", cursor:"pointer" }} />
                {(simDateRange.from||simDateRange.to) && <span onClick={()=>setSimDateRange({from:"",to:""})} style={{ fontSize:12, cursor:"pointer", color:"#fff", fontWeight:700 }}>✕</span>}
              </div>
              <button onClick={()=>{const init={};Object.entries(simFilterableValues).forEach(([k,v])=>{init[k]=[...v];});setSimFilters(init);setSimDateRange({from:"",to:""});}} style={{ fontSize:11, padding:"3px 10px", background:IPADE.navy, color:"#fff", border:"none", borderRadius:4, cursor:"pointer", marginLeft:8 }}>Limpiar filtros</button>
              <div style={{ marginLeft:"auto", display:"flex", gap:10, alignItems:"center" }}>
                <span style={{ fontSize:11, color:"#666" }}>{simFiltered.length} sesiones · edita el profesor a mano</span>
                <button onClick={exportSimExcel} style={{ fontSize:11, padding:"5px 14px", background:"#217346", color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontWeight:600 }}>📥 Colorigrama</button>
              </div>
            </div>

            {/* ─ Tabla detallada simulada (editable) ─ */}
            <div style={{ background:"#fff", borderRadius:12, boxShadow:"0 2px 8px rgba(0,0,0,.06)", overflow:"hidden" }}>
              <div style={{ overflowX:"auto", maxHeight:"calc(100vh - 130px)", minHeight:480, overflowY:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                  <thead>
                    <tr>{SIM_COLS.map(c=>(
                      <th key={c.key} onClick={()=>handleSimSort(c.key)} style={{ background:IPADE.navy, color:"#fff", padding:"10px 8px", textAlign:"left", fontSize:11, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap", userSelect:"none", position:"sticky", top:0, zIndex:10 }}>
                        {c.label} {simSortCol===c.key?(simSortDir==="asc"?"▲":"▼"):""}
                      </th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {simFiltered.map((r,i)=>(
                      <tr key={r["Id Sesión"]??i} style={{ background:i%2===0?"#fff":"#f9f8f5" }}>
                        {SIM_COLS.map(c=>{
                          if (c.key==="_profSim") {
                            const pc=profColors[r._profSim]||"#999";
                            return <td key={c.key} style={{ padding:"4px 6px", borderBottom:"1px solid #f0efea", borderLeft:`4px solid ${pc}`, background:`${pc}15`, fontWeight:600, minWidth:120 }}>
                              <EditableCell value={r._profSim} onCommit={val=>setSimOverrides(p=>({...p,[r["Id Sesión"]]:val}))} style={{fontWeight:600,color:pc}} />
                            </td>;
                          }
                          let val=c.fmt?c.fmt(r[c.key]):(r[c.key]??"");
                          const wrap = c.key==="Caso/Nota" || c.key==="Tema";
                          let style = wrap
                            ? { padding:"8px", borderBottom:"1px solid #f0efea", fontSize:11, verticalAlign:"top", whiteSpace:"normal", wordBreak:"break-word", minWidth:200, maxWidth:c.key==="Caso/Nota"?420:300 }
                            : { padding:"8px", borderBottom:"1px solid #f0efea", fontSize:11, verticalAlign:"top", maxWidth:"auto", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" };
                          if (c.key==="_progPadre"||c.key==="_programa") { const pc=progColors[r._progPadre]||"#999"; style.borderLeft=`4px solid ${pc}`; style.background=`${pc}15`; }
                          if (c.key==="_sede") { const sc=sedeColors[r._sede]||"#999"; style.borderLeft=`4px solid ${sc}`; style.background=`${sc}15`; val=<span style={{display:"flex",alignItems:"center"}}><ColorDot color={sc}/>{val}</span>; }
                          if (c.key==="_rolSim") { const rc=r._rolSim==="Titular"?"#C62828":r._rolSim==="Dupla"?"#2E7D32":r._rolSim==="PROP CF"?IPADE.gold:"#999"; style.color=rc; style.fontWeight=600; }
                          return <td key={c.key} style={style} title={String(c.fmt?c.fmt(r[c.key]):(r[c.key]??""))}>{val}</td>;
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
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
                        if (c.key==="_profesor") {
                          const profs = r._profesoresArr || [r._profesor];
                          const firstColor = profColors[profs[0]] || "#999";
                          style.borderLeft = `4px solid ${firstColor}`;
                          style.background = `${firstColor}15`;
                          val = <span style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                            {profs.map(p => <span key={p} style={{display:"inline-flex",alignItems:"center",whiteSpace:"nowrap"}}><ColorDot color={profColors[p]||"#999"}/>{p}</span>)}
                          </span>;
                        }
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
