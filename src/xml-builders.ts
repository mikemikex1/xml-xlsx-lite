import { Worksheet, Workbook, Cell } from './types';
import { CellModel } from './cell';
import { escapeXmlText, escapeXmlAttr, excelSerialFromDate, isDate, addrFromRC } from './utils';

/**
 * XML 標頭
 */
function xmlHeader(): string {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
}

/**
 * 建立 Content Types XML
 */
export function buildContentTypes(sheetCount: number, hasStyles: boolean, hasSharedStrings: boolean): string {
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

/**
 * 建立根關係 XML
 */
export function buildRootRels(): string {
  return [
    xmlHeader(),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    "</Relationships>"
  ].join("");
}

/**
 * 建立工作簿 XML
 */
export function buildWorkbookXml(sheets: Worksheet[]): { workbookXml: string; workbookRelsXml: string } {
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

/**
 * 建立工作表 XML
 */
export function buildSheetXml(ws: Worksheet, index: number, sstMap: Map<string, number>, workbook: Workbook): string {
  // Build <sheetData> with rows and cells
  const parts = [
    xmlHeader(),
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  ];

  // dimension if any cell exists
  if ((ws as any)._maxRowValue > 0 && (ws as any)._maxColValue > 0) {
    const dim = `A1:${addrFromRC((ws as any)._maxRowValue, (ws as any)._maxColValue)}`;
    parts.push(`<dimension ref="${dim}"/>`);
  }

  // Phase 3: 欄寬設定
  if ((ws as any)._columnWidthsValue && (ws as any)._columnWidthsValue.size > 0) {
    parts.push('<cols>');
    const cols = Array.from((ws as any)._columnWidthsValue.entries() as Iterable<[number, number]>).sort((a, b) => a[0] - b[0]);
    for (const [colNum, width] of cols) {
      parts.push(`<col min="${colNum}" max="${colNum}" width="${width}" customWidth="1"/>`);
    }
    parts.push('</cols>');
  }

  // Phase 3: 凍結窗格
  const freezePanes = (ws as any).getFreezePanes();
  if (freezePanes.row || freezePanes.column) {
    parts.push('<sheetViews>');
    parts.push('<sheetView workbookViewId="0">');
    if (freezePanes.row || freezePanes.column) {
      const topLeftCell = addrFromRC(
        freezePanes.row || 0,
        freezePanes.column || 0
      );
      parts.push(`<pane xSplit="${freezePanes.column || 0}" ySplit="${freezePanes.row || 0}" topLeftCell="${topLeftCell}" state="frozen"/>`);
    }
    parts.push('</sheetView>');
    parts.push('</sheetViews>');
  }

  parts.push("<sheetData>");

  for (const [r, rowMap] of ws.rows()) {
    // Phase 3: 列高設定
    const rowHeight = (ws as any).getRowHeight(r);
    const rowHeightAttr = rowHeight !== 15 ? ` ht="${rowHeight}" customHeight="1"` : '';
    
    parts.push(`<row r="${r}"${rowHeightAttr}>`);
    // cells sorted by col
    const cols = Array.from(rowMap.keys()).sort((a, b) => a - b);
    for (const c of cols) {
      const cell = rowMap.get(c)!;
      const raddr = cell.address; // e.g., "B12"
      const { t, v } = buildCellValue(cell, sstMap);
      const tAttr = t ? ` t="${t}"` : "";
      
      // 添加樣式索引
      const styleId = (workbook as any)._getStyleIndex(cell.options);
      const styleAttr = styleId > 0 ? ` s="${styleId}"` : "";
      
      // Phase 3: 公式支援
      const formulaAttr = cell.options.formula ? ` f="${cell.options.formula}"` : "";
      
      parts.push(`<c r="${raddr}"${tAttr}${styleAttr}${formulaAttr}><v>${v}</v></c>`);
    }
    parts.push("</row>");
  }

  parts.push("</sheetData>");

  // Phase 3: 合併儲存格
  const mergedRanges = (ws as any).getMergedRanges();
  if (mergedRanges.length > 0) {
    parts.push('<mergeCells count="' + mergedRanges.length + '">');
    for (const range of mergedRanges) {
      parts.push(`<mergeCell ref="${range}"/>`);
    }
    parts.push('</mergeCells>');
  }

  parts.push("</worksheet>");
  return parts.join("");
}

/**
 * 建立儲存格值
 */
function buildCellValue(cell: Cell, sstMap: Map<string, number>): { t: string | null; v: string } {
  const val = cell.value;
  
  // Phase 3: 公式支援
  if (cell.options.formula) {
    // 如果有公式，優先使用公式
    return { t: null, v: "" }; // 公式儲存格不需要值，Excel 會自動計算
  }
  
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

/**
 * 建立共享字串 XML
 */
export function buildSharedStringsXml(sstMap: Map<string, number>, sstArr: string[]): string {
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

/**
 * 建立樣式 XML
 */
export function buildStylesXml(workbook: Workbook): string {
  const parts = [
    xmlHeader(),
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  ];

  // 生成字體 XML
  const fonts = Array.from((workbook as any)._fontsMap.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<fonts count="${fonts.length}">`);
  for (const [fontKey, fontId] of fonts) {
    if (fontId === 0) {
      // 預設字體
      parts.push('<font><sz val="11"/><name val="Calibri"/></font>');
    } else {
      const font = JSON.parse(fontKey);
      const fontParts = ['<font>'];
      if (font.size) fontParts.push(`<sz val="${font.size}"/>`);
      if (font.name) fontParts.push(`<name val="${font.name}"/>`);
      if (font.bold) fontParts.push('<b/>');
      if (font.italic) fontParts.push('<i/>');
      if (font.underline) fontParts.push('<u/>');
      if (font.strike) fontParts.push('<strike/>');
      if (font.color) fontParts.push(`<color rgb="${font.color.replace('#', '')}"/>`);
      fontParts.push('</font>');
      parts.push(fontParts.join(''));
    }
  }
  parts.push('</fonts>');

  // 生成填滿 XML
  const fills = Array.from((workbook as any)._fillsMap.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<fills count="${fills.length}">`);
  for (const [fillKey, fillId] of fills) {
    if (fillId === 0) {
      // 預設填滿
      parts.push('<fill><patternFill patternType="none"/></fill>');
    } else {
      const fill = JSON.parse(fillKey);
      const fillParts = ['<fill>'];
      if (fill.type === 'pattern') {
        fillParts.push('<patternFill');
        if (fill.patternType) fillParts.push(`patternType="${fill.patternType}"`);
        fillParts.push('>');
        if (fill.fgColor) fillParts.push(`<fgColor rgb="${fill.fgColor.replace('#', '')}"/>`);
        if (fill.bgColor) fillParts.push(`<bgColor rgb="${fill.bgColor.replace('#', '')}"/>`);
        fillParts.push('</patternFill>');
      }
      fillParts.push('</fill>');
      parts.push(fillParts.join(''));
    }
  }
  parts.push('</fills>');

  // 生成邊框 XML
  const borders = Array.from((workbook as any)._bordersMap.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<borders count="${borders.length}">`);
  for (const [borderKey, borderId] of borders) {
    if (borderId === 0) {
      // 預設邊框
      parts.push('<border/>');
    } else {
      const border = JSON.parse(borderKey);
      const borderParts = ['<border>'];
      
      // 處理各個邊的樣式
      const sides = ['left', 'right', 'top', 'bottom'];
      for (const side of sides) {
        if (border[side]) {
          const sideBorder = border[side];
          borderParts.push(`<${side}`);
          if (sideBorder.style) borderParts.push(`style="${sideBorder.style}"`);
          borderParts.push('>');
          if (sideBorder.color) borderParts.push(`<color rgb="${sideBorder.color.replace('#', '')}"/>`);
          borderParts.push(`</${side}>`);
        }
      }
      
      borderParts.push('</border>');
      parts.push(borderParts.join(''));
    }
  }
  parts.push('</borders>');

  // 生成對齊 XML
  const alignments = Array.from((workbook as any)._alignmentsMap.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<cellStyleXfs count="${alignments.length}">`);
  for (const [alignmentKey, alignmentId] of alignments) {
    if (alignmentId === 0) {
      // 預設對齊
      parts.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    } else {
      const alignment = JSON.parse(alignmentKey);
      const xfParts = ['<xf'];
      if (alignment.font) xfParts.push('fontId="0"');
      if (alignment.fill) xfParts.push('fillId="0"');
      if (alignment.border) xfParts.push('borderId="0"');
      xfParts.push('>');
      
      // 對齊設定
      if (alignment.horizontal || alignment.vertical || alignment.wrapText || alignment.indent || alignment.rotation) {
        const alignParts = ['<alignment'];
        if (alignment.horizontal) alignParts.push(`horizontal="${alignment.horizontal}"`);
        if (alignment.vertical) alignParts.push(`vertical="${alignment.vertical}"`);
        if (alignment.wrapText) alignParts.push('wrapText="1"');
        if (alignment.indent) alignParts.push(`indent="${alignment.indent}"`);
        if (alignment.rotation) alignParts.push(`textRotation="${alignment.rotation}"`);
        alignParts.push('/>');
        xfParts.push(alignParts.join(' '));
      }
      
      xfParts.push('</xf>');
      parts.push(xfParts.join(' '));
    }
  }
  parts.push('</cellStyleXfs>');

  // 生成儲存格樣式 XML
  const styles = Array.from((workbook as any)._stylesMap.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<cellXfs count="${styles.length}">`);
  for (const [styleKey, styleId] of styles) {
    if (styleId === 0) {
      // 預設樣式
      parts.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    } else {
      const style = JSON.parse(styleKey);
      const xfParts = ['<xf'];
      if (style.font) xfParts.push('fontId="0"');
      if (style.fill) xfParts.push('fillId="0"');
      if (style.border) xfParts.push('borderId="0"');
      xfParts.push('xfId="0"');
      xfParts.push('>');
      
      // 對齊設定
      if (style.alignment) {
        const alignParts = ['<alignment'];
        if (style.alignment.horizontal) alignParts.push(`horizontal="${style.alignment.horizontal}"`);
        if (style.alignment.vertical) alignParts.push(`vertical="${style.alignment.vertical}"`);
        if (style.alignment.wrapText) alignParts.push('wrapText="1"');
        if (style.alignment.indent) alignParts.push(`indent="${style.alignment.indent}"`);
        if (style.alignment.rotation) alignParts.push(`textRotation="${style.alignment.rotation}"`);
        alignParts.push('/>');
        xfParts.push(alignParts.join(' '));
      }
      
      xfParts.push('</xf>');
      parts.push(xfParts.join(' '));
    }
  }
  parts.push('</cellXfs>');

  // 樣式名稱
  parts.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
  
  parts.push("</styleSheet>");
  return parts.join("");
}
