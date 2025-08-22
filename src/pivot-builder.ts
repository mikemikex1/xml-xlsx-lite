// src/pivot-builder.ts
import JSZip from "jszip";

// 為 Node.js 環境提供 Buffer 類型
declare global {
  interface Buffer extends Uint8Array {}
}

const esc = (s: string) =>
  String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
           .replace(/"/g,"&quot;").replace(/'/g,"&apos;");

export type PivotAgg = "sum" | "count" | "average" | "max" | "min" | "product";

export interface PivotFieldSpec { 
  name: string; 
}

export interface PivotValueSpec { 
  name: string; 
  agg?: PivotAgg; 
  displayName?: string; 
  numFmtId?: number; 
}

export interface PivotLayout { 
  rows?: PivotFieldSpec[]; 
  cols?: PivotFieldSpec[]; 
  values: PivotValueSpec[]; 
}

export interface CreatePivotOptions {
  sourceSheet: string; 
  sourceRange: string;
  targetSheet: string; 
  anchorCell: string;
  layout: PivotLayout; 
  refreshOnLoad?: boolean; 
  styleName?: string;
}

type FieldType = "text" | "number";

/** 你套件對外：Buffer in → 新 Buffer out */
export async function addPivotToWorkbookBuffer(buf: Buffer, opt: CreatePivotOptions): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buf);
  await PivotBuilder.attach(zip, opt);
  return await zip.generateAsync({ type: "nodebuffer" }) as Buffer;
}

export class PivotBuilder {
  static async attach(zip: JSZip, opt: CreatePivotOptions) {
    const refreshOnLoad = opt.refreshOnLoad ?? true;
    const styleName = opt.styleName ?? "PivotStyleMedium9";

    // ① 建立 workbook 映射（名稱→sheetN.xml）
    const map = await WorkbookMap.build(zip);

    // ② 解析來源範圍
    const { firstRow, lastRow } = rangeRows(opt.sourceRange);
    const recordCount = lastRow - firstRow;
    if (recordCount <= 0) throw new Error("sourceRange 必須包含至少 1 筆資料（含標題列）");

    const srcSheetPath = map.sheetFileByName(opt.sourceSheet);
    const srcSheetXml = await mustRead(zip, srcSheetPath);

    const headers = extractHeadersFromSheetXml(srcSheetXml, opt.sourceRange);
    const types = inferFieldTypesFromSheetXml(srcSheetXml, opt.sourceRange, headers);

    const name2idx = new Map(headers.map((h,i)=>[h,i]));
    const need = (n: string) => {
      if (!name2idx.has(n)) throw new Error(`找不到欄位：${n}`);
      return name2idx.get(n)!;
    };

    const rowIdxs = (opt.layout.rows ?? []).map(f=>need(f.name));
    const colIdxs = (opt.layout.cols ?? []).map(f=>need(f.name));
    const valIdxs = opt.layout.values.map(v=>need(v.name));
    
    // 值欄必須為數字
    opt.layout.values.forEach(v => {
      const i = need(v.name);
      if (types[i] !== "number") throw new Error(`值欄位必須為數值：${v.name}`);
    });

    // ③ 產生新 id / 檔名
    const cacheId = await map.nextPivotCacheId();
    const ptId = await map.nextPivotTableId();
    const cacheDefPath = `xl/pivotCache/pivotCacheDefinition${cacheId}.xml`;
    const cacheRecPath = `xl/pivotCache/pivotCacheRecords${cacheId}.xml`;
    const ptPath       = `xl/pivotTables/pivotTable${ptId}.xml`;

    // ④ 寫 cacheDefinition / cacheRecords（records 可留空，Excel 會自動重建）
    zip.file(cacheDefPath, genCacheDefinitionXml({
      sourceSheet: opt.sourceSheet, 
      sourceRange: opt.sourceRange,
      headers, 
      types, 
      recordCount, 
      refreshOnLoad
    }));
    zip.file(cacheRecPath, genEmptyCacheRecordsXml(recordCount));

    // ⑤ 掛到 workbook：rels + <pivotCaches>
    await map.addPivotCache(cacheId, cacheDefPath);

    // ⑥ 寫 pivotTableDefinition
    zip.file(ptPath, genPivotTableXml({
      cacheId, 
      headers, 
      rowIdxs, 
      colIdxs, 
      valIdxs,
      values: opt.layout.values, 
      anchorCell: opt.anchorCell, 
      styleName
    }));

    // ⑦ 把 pivot 掛到目標工作表（rels）
    const tgtSheetPath = map.sheetFileByName(opt.targetSheet);
    await map.linkPivotToSheet(tgtSheetPath, ptPath);

    // ⑧ 補 Content_Types
    await ensureContentTypes(zip, { cacheDefPath, cacheRecPath, ptPath });
  }
}

/* ───────── Workbook 映射 & rels 編輯 ───────── */

class WorkbookMap {
  constructor(
    private zip: JSZip,
    private wbXml: string,
    private wbRelsXml: string,
    private sheetIdToTarget: Map<string,string>,     // r:id → worksheets/sheetN.xml
    private nameToRid: Map<string,string>,           // sheet name → r:id
  ) {}
  
  static async build(zip: JSZip) {
    const wbXml = await mustRead(zip, "xl/workbook.xml");
    const wbRelsXml = await mustRead(zip, "xl/_rels/workbook.xml.rels");
    const nameToRid = new Map<string,string>();
    const sheetIdToTarget = new Map<string,string>();

    // 解析 sheet name → r:id
    for (const m of wbXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
      nameToRid.set(m[1], m[2]);
    }
    
    // 解析 r:id → target
    for (const m of wbRelsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Type="[^"]*worksheet"[^>]*Target="([^"]+)"/g)) {
      sheetIdToTarget.set(m[1], `xl/${m[2].replace(/^\.\.\//,'')}`);
    }
    
    return new WorkbookMap(zip, wbXml, wbRelsXml, sheetIdToTarget, nameToRid);
  }
  
  sheetFileByName(name: string) {
    const rid = this.nameToRid.get(name);
    if (!rid) throw new Error(`workbook 找不到工作表：${name}`);
    const target = this.sheetIdToTarget.get(rid);
    if (!target) throw new Error(`rels 找不到 target：${name}(${rid})`);
    return target;
  }
  
  async nextPivotCacheId() {
    const ids = Array.from(this.wbXml.matchAll(/<pivotCache[^>]*cacheId="(\d+)"/g)).map(m=>+m[1]);
    return (ids.length? Math.max(...ids):0) + 1;
  }
  
  async nextPivotTableId() {
    const files = this.zip.folder("xl/pivotTables")?.file(/.*/g) ?? [];
    const ids = files.map(f => +(f.name.match(/pivotTable(\d+)\.xml$/)?.[1] ?? 0));
    return (ids.length? Math.max(...ids):0) + 1;
  }
  
  async addPivotCache(cacheId: number, cacheDefPath: string) {
    // 1) workbook.xml.rels 新增 rel → cacheDefinition
    let rel = await mustRead(this.zip, "xl/_rels/workbook.xml.rels");
    const newRelId = nextRelId(rel);
    rel = rel.replace(/<\/Relationships>\s*$/i,
      `  <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/${cacheDefPath.split('/').pop()}"/>\n</Relationships>`
    );
    this.zip.file("xl/_rels/workbook.xml.rels", rel);

    // 2) workbook.xml 新增/擴充 <pivotCaches>
    let wb = await mustRead(this.zip, "xl/workbook.xml");
    if (/<pivotCaches>/.test(wb)) {
      wb = wb.replace(/<pivotCaches>([\s\S]*?)<\/pivotCaches>/,
        (_m, inner) => `<pivotCaches>${inner}<pivotCache cacheId="${cacheId}" r:id="${newRelId}"/></pivotCaches>`);
    } else {
      wb = wb.replace(/<\/workbook>\s*$/i,
        `  <pivotCaches><pivotCache cacheId="${cacheId}" r:id="${newRelId}"/></pivotCaches>\n</workbook>`);
    }
    this.zip.file("xl/workbook.xml", wb);
  }
  
  async linkPivotToSheet(sheetPath: string, ptPath: string) {
    const relPath = sheetPath.replace(/worksheets\/(sheet\d+\.xml)$/,'worksheets/_rels/$1.rels');
    let rel = (await readOrNull(this.zip, relPath)) ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n</Relationships>`;
    const newRelId = nextRelId(rel);
    rel = rel.replace(/<\/Relationships>\s*$/i,
      `  <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/${ptPath.split('/').pop()}"/>\n</Relationships>`
    );
    this.zip.file(relPath, rel);
    
    // 確保目標 sheet 至少有 <sheetData/>
    const ws = await mustRead(this.zip, sheetPath);
    if (!/<sheetData/.test(ws)) {
      this.zip.file(sheetPath, ws.replace(/<\/worksheet>\s*$/i, `  <sheetData/>\n</worksheet>`));
    }
  }
}

function nextRelId(relXml: string) {
  const ids = Array.from(relXml.matchAll(/Id="rId(\d+)"/g)).map(m=>+m[1]);
  return `rId${(ids.length? Math.max(...ids):0)+1}`;
}

/* ───────── Cache / Pivot XML 生成 ───────── */

function genCacheDefinitionXml(p:{
  sourceSheet:string; 
  sourceRange:string; 
  headers:string[]; 
  types:FieldType[];
  recordCount:number; 
  refreshOnLoad:boolean;
}) {
  const fields = p.headers.map((h,i)=>{
    const isNum = p.types[i]==="number";
    const shared = isNum ? `<sharedItems containsNumber="1"/>` : `<sharedItems/>`;
    return `<cacheField name="${esc(h)}" numFmtId="0">${shared}</cacheField>`;
  }).join("");
  
  const refresh = p.refreshOnLoad ? ` refreshOnLoad="1"` : ``;
  
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  invalid="1" recordCount="${p.recordCount}"${refresh}>
  <cacheSource type="worksheet">
    <worksheetSource sheet="${esc(p.sourceSheet)}" ref="${esc(p.sourceRange)}"/>
  </cacheSource>
  <cacheFields count="${p.headers.length}">
    ${fields}
  </cacheFields>
</pivotCacheDefinition>`;
}

function genEmptyCacheRecordsXml(recordCount:number) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${recordCount}">
</pivotCacheRecords>`;
}

function genPivotTableXml(p:{
  cacheId:number; 
  headers:string[]; 
  rowIdxs:number[]; 
  colIdxs:number[]; 
  valIdxs:number[];
  values:PivotValueSpec[]; 
  anchorCell:string; 
  styleName:string;
}) {
  const fields = p.headers.map((_,i)=>{
    if (p.rowIdxs.includes(i)) return `<pivotField axis="axisRow" showAll="0"/>`;
    if (p.colIdxs.includes(i)) return `<pivotField axis="axisCol" showAll="0"/>`;
    if (p.valIdxs.includes(i)) return `<pivotField dataField="1"/>`;
    return `<pivotField/>`;
  }).join("");
  
  const rowFields = p.rowIdxs.length ? `<rowFields count="${p.rowIdxs.length}">${p.rowIdxs.map(x=>`<field x="${x}"/>`).join("")}</rowFields>` : "";
  const colFields = p.colIdxs.length ? `<colFields count="${p.colIdxs.length}">${p.colIdxs.map(x=>`<field x="${x}"/>`).join("")}</colFields>` : "";
  
  const dataFields = p.values.map((v,i)=>{
    const fld = p.valIdxs[i];
    const subtotal = v.agg ?? "sum";
    const name = esc(v.displayName ?? v.name);
    const numFmtId = v.numFmtId ?? 0;
    return `<dataField fld="${fld}" baseField="0" baseItem="0" subtotal="${subtotal}" name="${name}" numFmtId="${numFmtId}"/>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable${p.cacheId}" cacheId="${p.cacheId}" dataCaption="值" updatedVersion="3" createdVersion="3"
  useAutoFormatting="1" applyNumberFormats="1" applyBorderFormats="1"
  applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="1" applyWidthHeightFormats="1">
  <location ref="${esc(p.anchorCell)}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="${p.headers.length}">
    ${fields}
  </pivotFields>
  ${rowFields}
  ${colFields}
  <dataFields count="${p.values.length}">
    ${dataFields}
  </dataFields>
  <pivotTableStyleInfo name="${esc(p.styleName)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
</pivotTableDefinition>`;
}

/* ───────── Content_Types 編輯 ───────── */

async function ensureContentTypes(zip:JSZip, p:{cacheDefPath:string; cacheRecPath:string; ptPath:string;}) {
  const path = "[Content_Types].xml";
  let xml = await mustRead(zip, path);
  const ensure = (part:string, type:string) => {
    const tag = `<Override PartName="/${part}" ContentType="${type}"/>`;
    if (!xml.includes(tag)) xml = xml.replace(/<\/Types>\s*$/i, `  ${tag}\n</Types>`);
  };
  ensure(p.cacheDefPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
  ensure(p.cacheRecPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
  ensure(p.ptPath,       "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
  zip.file(path, xml);
}

/* ───────── 來源表解析（標題/型別） ───────── */

function extractHeadersFromSheetXml(sheetXml:string, range:string):string[] {
  const { firstCol, lastCol, firstRow } = rangeColsRows(range);
  const headers:string[] = [];
  for (let c=firstCol; c<=lastCol; c++) {
    const addr = colNumToName(c)+firstRow;
    const cell = findCell(sheetXml, addr);
    headers.push(readCellText(cell) || `F${c-firstCol+1}`);
  }
  return headers;
}

function inferFieldTypesFromSheetXml(sheetXml:string, range:string, headers:string[]):FieldType[] {
  const { firstCol, lastCol, firstRow, lastRow } = rangeColsRows(range);
  const types:FieldType[] = new Array(headers.length).fill("text");
  for (let r=firstRow+1; r<=Math.min(lastRow, firstRow+20); r++) {
    for (let c=firstCol; c<=lastCol; c++) {
      const m = findCell(sheetXml, colNumToName(c)+r);
      if (!m) continue;
      if (/<v>\s*-?\d+(\.\d+)?\s*<\/v>/.test(m)) types[c-firstCol] = "number";
    }
  }
  return types;
}

function findCell(sheetXml:string, addr:string):string|null {
  const re = new RegExp(`<c[^>]*\\br="${addr}"[^>]*>([\\s\\S]*?)</c>`, "i");
  const m = sheetXml.match(re);
  return m ? m[0] : null;
}

function readCellText(cellXml:string|null):string {
  if (!cellXml) return "";
  let m = cellXml.match(/<is>\s*<t>([\s\S]*?)<\/t>\s*<\/is>/);
  if (m) return decodeXml(m[1]);
  m = cellXml.match(/<v>([\s\S]*?)<\/v>/);
  return m ? decodeXml(m[1]) : "";
}

function decodeXml(s:string){
  return s.replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&amp;/g,"&").replace(/&quot;/g,'"').replace(/&apos;/g,"'");
}

/* ───────── A1 / 欄列工具 & IO ───────── */

function colNameToNum(name:string){
  let n=0;
  for(let i=0;i<name.length;i++)n=n*26+(name.charCodeAt(i)-64);
  return n;
}

function colNumToName(n:number){
  let s="";
  while(n>0){
    n--;
    s=String.fromCharCode(65+(n%26))+s;
    n=Math.floor(n/26);
  }
  return s;
}

function rangeColsRows(a1:string){
  const m = a1.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i); 
  if(!m) throw new Error(`非法範圍：${a1}`);
  const firstCol=colNameToNum(m[1].toUpperCase()), 
        firstRow=+m[2], 
        lastCol=colNameToNum(m[3].toUpperCase()), 
        lastRow=+m[4];
  return { firstCol, firstRow, lastCol, lastRow };
}

function rangeRows(a1:string){ 
  const {firstRow,lastRow}=rangeColsRows(a1); 
  return {firstRow,lastRow}; 
}

async function mustRead(zip:JSZip, path:string){ 
  const f=zip.file(path); 
  if(!f) throw new Error(`找不到檔案：${path}`); 
  return await f.async("string"); 
}

async function readOrNull(zip:JSZip, path:string){ 
  const f=zip.file(path); 
  return f? await f.async("string"): null; 
}
