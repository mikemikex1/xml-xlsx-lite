import JSZip from "jszip";

/*** Type Definitions ***/
export interface CellOptions {
  numFmt?: string;
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    name?: string;
    color?: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  fill?: {
    type?: 'pattern' | 'gradient';
    color?: string;
    patternType?: string;
  };
  border?: {
    style?: string;
    color?: string;
  };
}

export interface Cell {
  address: string;
  value: number | string | boolean | Date | null;
  type: 'n' | 's' | 'b' | 'd' | null;
  options: CellOptions;
}

export interface Worksheet {
  name: string;
  getCell(address: string): Cell;
  setCell(address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell;
  rows(): Generator<[number, Map<number, Cell>]>;
}

export interface Workbook {
  getWorksheet(nameOrIndex: string | number): Worksheet;
  getCell(worksheet: string | Worksheet, address: string): Cell;
  setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell;
  writeBuffer(): Promise<ArrayBuffer>;
}

/*** Utilities ***/
const COL_A_CODE = "A".charCodeAt(0);
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30 (Excel's 1900 date system, including the 1900 leap-year bug)

function colToNumber(col: string): number {
  // e.g., A -> 1, Z -> 26, AA -> 27
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - COL_A_CODE + 1);
  }
  return n;
}

function numberToCol(n: number): string {
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(COL_A_CODE + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function parseAddress(addr: string): { col: number; row: number } {
  // "B12" -> { col: 2, row: 12 }
  const m = /^([A-Z]+)(\d+)$/.exec(addr.toUpperCase());
  if (!m) throw new Error(`Invalid cell address: ${addr}`);
  return { col: colToNumber(m[1]), row: parseInt(m[2], 10) };
}

function addrFromRC(row: number, col: number): string {
  return `${numberToCol(col)}${row}`;
}

function isDate(val: any): val is Date {
  return val instanceof Date;
}

function excelSerialFromDate(d: Date): number {
  // Convert JS Date (UTC) to Excel serial number (1900 system)
  const msPerDay = 24 * 60 * 60 * 1000;
  const diff = (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - EXCEL_EPOCH.getTime()) / msPerDay;
  return diff;
}

/*** Internal Cell Model ***/
class CellModel implements Cell {
  address: string;
  value: number | string | boolean | Date | null;
  type: 'n' | 's' | 'b' | 'd' | null;
  options: CellOptions;

  constructor(address: string) {
    this.address = address; // "A1"
    this.value = null;      // number | string | boolean | Date | null
    this.type = null;       // 'n' | 's' | 'b' | 'd' | null (internal hint)
    this.options = {};      // placeholder for exceljs-like options (numFmt, font, alignment, etc.)
  }
}

/*** Worksheet Model ***/
class WorksheetImpl implements Worksheet {
  name: string;
  private _grid: Map<number, Map<number, CellModel>>;
  private _maxRow: number;
  private _maxCol: number;

  constructor(name: string) {
    this.name = name;
    // row -> (col -> CellModel)
    this._grid = new Map();
    this._maxRow = 0;
    this._maxCol = 0;
  }

  /** exceljs-like */
  getCell(address: string): Cell {
    const { row, col } = parseAddress(address);
    return this._ensureCell(row, col);
  }

  /** exceljs-like */
  setCell(address: string, value: number | string | boolean | Date | null, options: CellOptions = {}): Cell {
    const cell = this.getCell(address);
    cell.value = value;
    cell.type = inferCellType(value);
    cell.options = { ...cell.options, ...options };
    return cell;
  }

  private _ensureCell(row: number, col: number): CellModel {
    if (!this._grid.has(row)) this._grid.set(row, new Map());
    const rowMap = this._grid.get(row)!;
    if (!rowMap.has(col)) {
      const address = addrFromRC(row, col);
      rowMap.set(col, new CellModel(address));
      if (row > this._maxRow) this._maxRow = row;
      if (col > this._maxCol) this._maxCol = col;
    }
    return rowMap.get(col)!;
  }

  /** Iterate rows ascending */
  *rows(): Generator<[number, Map<number, Cell>]> {
    const rows = Array.from(this._grid.keys()).sort((a, b) => a - b);
    for (const r of rows) {
      yield [r, this._grid.get(r)!];
    }
  }
}

function inferCellType(value: any): 'n' | 's' | 'b' | 'd' | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return "n";
  if (typeof value === "boolean") return "b";
  if (isDate(value)) return "n"; // we will write as serial number for now
  return "s"; // default: string
}

/*** Workbook ***/
export class WorkbookImpl implements Workbook {
  private _sheets: WorksheetImpl[];
  private _sheetByName: Map<string, WorksheetImpl>;
  // shared strings handling (Excel prefers sharedStrings.xml for strings)
  private _sst: Map<string, number>; // string -> idx
  private _sstArr: string[];     // idx -> string

  constructor() {
    this._sheets = [];
    this._sheetByName = new Map();
    // shared strings handling (Excel prefers sharedStrings.xml for strings)
    this._sst = new Map(); // string -> idx
    this._sstArr = [];     // idx -> string
  }

  /** exceljs-like */
  getWorksheet(nameOrIndex: string | number): Worksheet {
    if (typeof nameOrIndex === "number") {
      const idx0 = nameOrIndex - 1; // exceljs is 1-based index; we accept both? We'll treat numbers as 1-based.
      const ws = this._sheets[idx0];
      if (!ws) throw new Error(`Worksheet index out of bounds: ${nameOrIndex}`);
      return ws;
    }
    if (this._sheetByName.has(nameOrIndex)) return this._sheetByName.get(nameOrIndex)!;
    const ws = new WorksheetImpl(nameOrIndex);
    this._sheets.push(ws);
    this._sheetByName.set(nameOrIndex, ws);
    return ws;
  }

  /** Convenience passthroughs */
  getCell(worksheet: string | Worksheet, address: string): Cell {
    const ws = typeof worksheet === "string" ? this.getWorksheet(worksheet) : worksheet;
    return ws.getCell(address);
  }

  setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options: CellOptions = {}): Cell {
    const ws = typeof worksheet === "string" ? this.getWorksheet(worksheet) : worksheet;
    return ws.setCell(address, value, options);
  }

  /** Build .xlsx as ArrayBuffer */
  async writeBuffer(): Promise<ArrayBuffer> {
    const zip = new JSZip();

    // Prepare XML parts
    const contentTypes = buildContentTypes(this._sheets.length, /*hasStyles*/ true, /*hasSharedStrings*/ true);
    const rootRels = buildRootRels();
    const { workbookXml, workbookRelsXml } = buildWorkbookXml(this._sheets);

    const sheetsXml = this._sheets.map((ws, i) => buildSheetXml(ws, i + 1, this._sst));

    const sharedStringsXml = buildSharedStringsXml(this._sst, this._sstArr);
    const stylesXml = buildStylesXml();

    // Add to zip
    zip.file("[Content_Types].xml", contentTypes);
    const rels = zip.folder("_rels");
    rels.file(".rels", rootRels);

    const xl = zip.folder("xl");
    xl.file("workbook.xml", workbookXml);
    const xlrels = xl.folder("_rels");
    xlrels.file("workbook.xml.rels", workbookRelsXml);

    const wsFolder = xl.folder("worksheets");
    for (let i = 0; i < sheetsXml.length; i++) {
      wsFolder.file(`sheet${i + 1}.xml`, sheetsXml[i]);
    }

    xl.file("sharedStrings.xml", sharedStringsXml);
    xl.file("styles.xml", stylesXml);

    // Generate and return ArrayBuffer
    return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  }

  /** Internal: called by buildSheetXml to register shared strings */
  private _sstIndex(str: string): number {
    if (this._sst.has(str)) return this._sst.get(str)!;
    const idx = this._sst.size;
    this._sst.set(str, idx);
    this._sstArr[idx] = str;
    return idx;
  }
}

/*** XML builders ***/
function xmlHeader(): string {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
}

function buildContentTypes(sheetCount: number, hasStyles: boolean, hasSharedStrings: boolean): string {
  const parts = [
    xmlHeader(),
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>'
  ];
  parts.push('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' );
  for (let i = 1; i <= sheetCount; i++) {
    parts.push(`<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  }
  if (hasStyles) parts.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' );
  if (hasSharedStrings) parts.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' );
  parts.push("</Types>");
  return parts.join("");
}

function buildRootRels(): string {
  return [
    xmlHeader(),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    "</Relationships>"
  ].join("");
}

function buildWorkbookXml(sheets: WorksheetImpl[]): { workbookXml: string; workbookRelsXml: string } {
  const workbookXml = [
    xmlHeader(),
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    "<sheets>"
  ];
  const workbookRels = [
    xmlHeader(),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  ];
  for (let i = 0; i < sheets.length; i++) {
    const sheetId = i + 1;
    const name = escapeXmlAttr(sheets[i].name || `Sheet${sheetId}`);
    workbookXml.push(`<sheet name="${name}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`);
    workbookRels.push(`<Relationship Id="rId${sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetId}.xml"/>`);
  }
  workbookXml.push("</sheets>", "</workbook>");
  workbookRels.push("</Relationships>");
  return { workbookXml: workbookXml.join(""), workbookRelsXml: workbookRels.join("") };
}

function buildSheetXml(ws: WorksheetImpl, index: number, sstMap: Map<string, number>): string {
  // Build <sheetData> with rows and cells
  const parts = [
    xmlHeader(),
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  ];

  // dimension if any cell exists
  if ((ws as any)._maxRow > 0 && (ws as any)._maxCol > 0) {
    const dim = `A1:${addrFromRC((ws as any)._maxRow, (ws as any)._maxCol)}`;
    parts.push(`<dimension ref="${dim}"/>`);
  }

  parts.push("<sheetData>");

  for (const [r, rowMap] of ws.rows()) {
    parts.push(`<row r="${r}">`);
    // cells sorted by col
    const cols = Array.from(rowMap.keys()).sort((a, b) => a - b);
    for (const c of cols) {
      const cell = rowMap.get(c)!;
      const raddr = cell.address; // e.g., "B12"
      const { t, v } = buildCellValue(cell, sstMap);
      const tAttr = t ? ` t="${t}"` : "";
      parts.push(`<c r="${raddr}"${tAttr}><v>${v}</v></c>`);
    }
    parts.push("</row>");
  }

  parts.push("</sheetData>");
  parts.push("</worksheet>");
  return parts.join("");
}

function buildCellValue(cell: CellModel, sstMap: Map<string, number>): { t: string | null; v: string } {
  const val = cell.value;
  if (val === null || val === undefined) return { t: null, v: "" };
  if (typeof val === "number") return { t: "n", v: String(val) };
  if (typeof val === "boolean") return { t: "b", v: val ? "1" : "0" };
  if (isDate(val)) return { t: "n", v: String(excelSerialFromDate(val)) };
  // string: add to shared strings
  let sIdx: number;
  const key = String(val);
  if (sstMap.has(key)) sIdx = sstMap.get(key)!;
  else {
    sIdx = sstMap.size;
    sstMap.set(key, sIdx);
  }
  return { t: "s", v: String(sIdx) };
}

function buildSharedStringsXml(sstMap: Map<string, number>, sstArr: string[]): string {
  // sstArr may be sparse if we built with map-only during sheets; rebuild from map in order
  const arr = new Array(sstMap.size);
  for (const [str, idx] of sstMap.entries()) arr[idx] = str;

  const parts = [
    xmlHeader(),
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${arr.length}" uniqueCount="${arr.length}">`
  ];
  for (const s of arr) {
    parts.push(`<si><t>${escapeXmlText(s)}</t></si>`);
  }
  parts.push("</sst>");
  return parts.join("");
}

function buildStylesXml(): string {
  // Minimal styles part so Excel is happy. No custom formats yet.
  return [
    xmlHeader(),
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    // number formats
    "<numFmts count=\"0\"/>",
    // fonts, fills, borders (at least one default required)
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>",
    "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
    "<borders count=\"1\"><border/></borders>",
    // cell style formats (one default xf)
    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
    // cell formats (one default xf)
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
    // stylesheet cell styles (Normal)
    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
    "</styleSheet>"
  ].join("");
}

/*** XML helpers ***/
function escapeXmlText(str: any): string {
  // 使用 replace 搭配正則表達式以支援較舊的 JavaScript 版本
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}


function escapeXmlAttr(str: any): string {
  return escapeXmlText(str);
}

// Export the main class
export const Workbook = WorkbookImpl;

// Default export for convenience
export default { Workbook: WorkbookImpl };
