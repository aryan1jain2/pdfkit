import JSZip from "jszip";
import { jsPDF } from "jspdf";
async function loadZip(input) {
  try {
    const zip = await JSZip.loadAsync(input);
    return {
      async getText(path) {
        const file = zip.file(path);
        if (!file) return null;
        return file.async("string");
      },
      async getBytes(path) {
        const file = zip.file(path);
        if (!file) return null;
        return file.async("uint8array");
      },
      async getDataUrl(path, mimeType) {
        const file = zip.file(path);
        if (!file) return null;
        const b64 = await file.async("base64");
        return `data:${mimeType};base64,${b64}`;
      },
      listFiles() {
        return Object.keys(zip.files).filter((k) => {
          var _a;
          return !((_a = zip.files[k]) == null ? void 0 : _a.dir);
        });
      }
    };
  } catch (e) {
    return {
      code: "CORRUPT_ZIP",
      message: "Failed to open file as ZIP archive",
      detail: String(e)
    };
  }
}
function isParseError(x) {
  return typeof x === "object" && x !== null && "code" in x;
}
function parseXml(xmlString) {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "application/xml");
    if (doc.querySelector("parsererror")) return null;
    return doc;
  } catch {
    return null;
  }
}
function getEl(parent, localName) {
  var _a;
  const results = ((_a = parent.getElementsByTagNameNS) == null ? void 0 : _a.call(parent, "*", localName)) ?? parent.getElementsByTagNameNS("*", localName);
  return results.length > 0 ? results[0] ?? null : null;
}
function getEls(parent, localName) {
  var _a;
  const results = ((_a = parent.getElementsByTagNameNS) == null ? void 0 : _a.call(parent, "*", localName)) ?? parent.getElementsByTagNameNS("*", localName);
  return Array.from(results);
}
function attr(el, name) {
  const direct = el.getAttribute(name);
  if (direct !== null) return direct;
  const localName = name.includes(":") ? name.split(":")[1] : name;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.localName === localName) return a.value;
  }
  return "";
}
function numAttr(el, name, fallback = 0) {
  const v = parseFloat(el.getAttribute(name) ?? "");
  return isNaN(v) ? fallback : v;
}
const PT_PER_INCH = 72;
const EMU_PER_INCH = 914400;
function emuToPt(emu) {
  return emu / EMU_PER_INCH * PT_PER_INCH;
}
function halfPtToPt(halfPt) {
  return halfPt / 100;
}
function cmToPt(cm) {
  return cm * 28.3465;
}
function hexToRgba(hex, alpha = 1) {
  const clean = hex.replace("#", "");
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  return { r, g, b, a: alpha };
}
const BLACK = { r: 0, g: 0, b: 0, a: 1 };
const WHITE = { r: 255, g: 255, b: 255, a: 1 };
const TRANSPARENT = { r: 0, g: 0, b: 0, a: 0 };
const PAGE_SIZES = {
  SLIDE_WIDESCREEN: { width: 720, height: 405 }
};
async function parsePptx(file) {
  const zip = await loadZip(file);
  if (isParseError(zip)) return { ok: false, error: zip };
  const presXml = await zip.getText("ppt/presentation.xml");
  if (!presXml) return { ok: false, error: { code: "MISSING_ENTRY", message: "ppt/presentation.xml not found" } };
  const presDoc = parseXml(presXml);
  if (!presDoc) return { ok: false, error: { code: "XML_PARSE_FAILED", message: "Could not parse presentation.xml" } };
  const slideSize = readSlideSize(presDoc);
  const themeXml = await zip.getText("ppt/theme/theme1.xml");
  const themeColors = readThemeColors(themeXml);
  const relsXml = await zip.getText("ppt/_rels/presentation.xml.rels");
  const slideOrder = readSlideOrder(relsXml);
  const pages = [];
  for (const slidePath of slideOrder) {
    const slideXml = await zip.getText(`ppt/${slidePath}`);
    if (!slideXml) continue;
    const slideDoc = parseXml(slideXml);
    if (!slideDoc) continue;
    const slideFile = slidePath.split("/").pop();
    const slideRelsXml = await zip.getText(`ppt/slides/_rels/${slideFile}.rels`);
    const { images: imageMap, hyperlinks } = await buildSlideRels(zip, slideRelsXml);
    pages.push(parseSlide(slideDoc, slideSize, imageMap, hyperlinks, themeColors));
  }
  return { ok: true, doc: { format: "pptx", pages, metadata: {} } };
}
function readSlideSize(presDoc) {
  const sldSz = getEl(presDoc, "sldSz");
  if (!sldSz) return PAGE_SIZES.SLIDE_WIDESCREEN;
  return {
    width: emuToPt(numAttr(sldSz, "cx", 6858e3)),
    height: emuToPt(numAttr(sldSz, "cy", 3858750))
  };
}
function readThemeColors(themeXml) {
  const map = /* @__PURE__ */ new Map();
  if (!themeXml) return map;
  const doc = parseXml(themeXml);
  if (!doc) return map;
  const clrScheme = getEl(doc, "clrScheme");
  if (!clrScheme) return map;
  for (let i = 0; i < clrScheme.children.length; i++) {
    const slot = clrScheme.children[i];
    const name = slot.localName;
    const srgb = getEl(slot, "srgbClr");
    if (srgb) {
      map.set(name, hexToRgba(attr(srgb, "val")));
      continue;
    }
    const sys = getEl(slot, "sysClr");
    if (sys) {
      const lastClr = attr(sys, "lastClr");
      if (lastClr) map.set(name, hexToRgba(lastClr));
    }
  }
  return map;
}
function resolveSchemeColor(el, themeColors) {
  const val = attr(el, "val");
  const aliasMap = {
    tx1: "dk1",
    tx2: "dk2",
    bg1: "lt1",
    bg2: "lt2"
  };
  const key = aliasMap[val] ?? val;
  let color = themeColors.get(key) ?? BLACK;
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i];
    const name = child.localName;
    if (name === "lumMod") {
      const mod = numAttr(child, "val", 1e5) / 1e5;
      color = { ...color, r: Math.min(255, Math.round(color.r * mod)), g: Math.min(255, Math.round(color.g * mod)), b: Math.min(255, Math.round(color.b * mod)) };
    } else if (name === "lumOff") {
      const off = numAttr(child, "val", 0) / 1e5;
      color = { ...color, r: Math.min(255, Math.round(color.r + 255 * off)), g: Math.min(255, Math.round(color.g + 255 * off)), b: Math.min(255, Math.round(color.b + 255 * off)) };
    } else if (name === "alpha") {
      color = { ...color, a: Math.min(1, numAttr(child, "val", 1e5) / 1e5) };
    } else if (name === "shade") {
      const shade = numAttr(child, "val", 1e5) / 1e5;
      color = { ...color, r: Math.round(color.r * shade), g: Math.round(color.g * shade), b: Math.round(color.b * shade) };
    } else if (name === "tint") {
      const tint = numAttr(child, "val", 1e5) / 1e5;
      color = { ...color, r: Math.round(color.r + (255 - color.r) * (1 - tint)), g: Math.round(color.g + (255 - color.g) * (1 - tint)), b: Math.round(color.b + (255 - color.b) * (1 - tint)) };
    }
  }
  return color;
}
function readSlideOrder(relsXml) {
  if (!relsXml) return [];
  const doc = parseXml(relsXml);
  if (!doc) return [];
  return getEls(doc, "Relationship").filter((el) => attr(el, "Type").includes('/slide"') || attr(el, "Type").endsWith("/slide")).sort((a, b) => {
    const aId = parseInt(attr(a, "Id").replace("rId", ""));
    const bId = parseInt(attr(b, "Id").replace("rId", ""));
    return aId - bId;
  }).map((el) => attr(el, "Target"));
}
async function buildSlideRels(zip, relsXml) {
  var _a;
  const images = /* @__PURE__ */ new Map();
  const hyperlinks = /* @__PURE__ */ new Map();
  if (!relsXml) return { images, hyperlinks };
  const doc = parseXml(relsXml);
  if (!doc) return { images, hyperlinks };
  const MIME = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
    webp: "image/webp",
    emf: "image/emf",
    wmf: "image/wmf"
  };
  for (const rel of getEls(doc, "Relationship")) {
    const type = attr(rel, "Type");
    const rId = attr(rel, "Id");
    const target = attr(rel, "Target");
    if (type.includes("image")) {
      const path = target.replace("../", "ppt/");
      const ext = ((_a = path.split(".").pop()) == null ? void 0 : _a.toLowerCase()) ?? "";
      const mime = MIME[ext] ?? "image/png";
      const dataUrl = await zip.getDataUrl(path, mime);
      if (dataUrl) images.set(rId, dataUrl);
    } else if (type.includes("hyperlink")) {
      if (target.startsWith("http") || attr(rel, "TargetMode") === "External") {
        hyperlinks.set(rId, target);
      }
    }
  }
  return { images, hyperlinks };
}
function parseSlide(slideDoc, size, imageMap, hyperlinks, themeColors) {
  const elements = [];
  let background = WHITE;
  const bgEl = getEl(slideDoc, "bg");
  if (bgEl) {
    const solidFill = getEl(bgEl, "solidFill");
    if (solidFill) background = readColour(solidFill, themeColors) ?? WHITE;
  }
  const spTree = getEl(slideDoc, "spTree");
  if (!spTree) return { ...size, background, elements };
  for (const sp of getEls(spTree, "sp")) {
    const el = parseShape(sp, themeColors, hyperlinks);
    if (el) elements.push(el);
  }
  for (const cxn of getEls(spTree, "cxnSp")) {
    const el = parseConnector(cxn, themeColors);
    if (el) elements.push(el);
  }
  for (const pic of getEls(spTree, "pic")) {
    const el = parsePicture(pic, imageMap);
    if (el) elements.push(el);
  }
  for (const gf of getEls(spTree, "graphicFrame")) {
    const tbl = getEl(gf, "tbl");
    if (tbl) {
      const el = parseTable(gf, tbl, themeColors);
      if (el) elements.push(el);
    }
  }
  return { width: size.width, height: size.height, background, elements };
}
function parseShape(sp, themeColors, hyperlinks = /* @__PURE__ */ new Map()) {
  const xfrm = getEl(sp, "xfrm");
  if (!xfrm) return null;
  const off = getEl(xfrm, "off");
  const ext = getEl(xfrm, "ext");
  if (!off || !ext) return null;
  const x = emuToPt(numAttr(off, "x"));
  const y = emuToPt(numAttr(off, "y"));
  const w = emuToPt(numAttr(ext, "cx"));
  const h = emuToPt(numAttr(ext, "cy"));
  const txBody = getEl(sp, "txBody");
  const spPr = getEl(sp, "spPr");
  if (txBody) {
    return parseTextBox(txBody, x, y, w, h, themeColors, spPr, hyperlinks);
  }
  if (!spPr) return null;
  const noFill = getEl(spPr, "noFill");
  const solidFill = getEl(spPr, "solidFill");
  const fill = noFill ? TRANSPARENT : solidFill ? readColour(solidFill, themeColors) ?? TRANSPARENT : TRANSPARENT;
  const ln = getEl(spPr, "ln");
  const lnNoFill = ln ? getEl(ln, "noFill") : null;
  const lnSolidFill = ln && !lnNoFill ? getEl(ln, "solidFill") : null;
  const stroke = lnSolidFill ? readColour(lnSolidFill, themeColors) ?? void 0 : void 0;
  const strokeWidth = ln && !lnNoFill ? Math.max(emuToPt(numAttr(ln, "w", 12700)), 0.5) : 0;
  const prstGeom = getEl(spPr, "prstGeom");
  const prst = prstGeom ? attr(prstGeom, "prst") : "";
  const kind = prst === "ellipse" ? "ellipse" : "rect";
  return { type: "shape", kind, x, y, width: w, height: h, fill, stroke, strokeWidth };
}
function parseConnector(cxn, themeColors) {
  const xfrm = getEl(cxn, "xfrm");
  if (!xfrm) return null;
  const off = getEl(xfrm, "off");
  const ext = getEl(xfrm, "ext");
  if (!off || !ext) return null;
  const x = emuToPt(numAttr(off, "x"));
  const y = emuToPt(numAttr(off, "y"));
  const w = emuToPt(numAttr(ext, "cx"));
  const h = emuToPt(numAttr(ext, "cy"));
  const flipH = attr(xfrm, "flipH") === "1";
  const flipV = attr(xfrm, "flipV") === "1";
  const spPr = getEl(cxn, "spPr");
  const ln = spPr ? getEl(spPr, "ln") : null;
  const lnSolidFill = ln ? getEl(ln, "solidFill") : null;
  const stroke = lnSolidFill ? readColour(lnSolidFill, themeColors) ?? BLACK : BLACK;
  const strokeWidth = ln ? Math.max(emuToPt(numAttr(ln, "w", 12700)), 0.5) : 0.75;
  const tailEnd = ln ? getEl(ln, "tailEnd") : null;
  const headEnd = ln ? getEl(ln, "headEnd") : null;
  const hasArrow = tailEnd && attr(tailEnd, "type") !== "none" || headEnd && attr(headEnd, "type") !== "none";
  const shape = {
    type: "shape",
    kind: hasArrow ? "arrow" : "line",
    x,
    y,
    width: w,
    height: h,
    fill: TRANSPARENT,
    stroke,
    strokeWidth
  };
  if (flipH) shape.flipH = true;
  if (flipV) shape.flipV = true;
  return shape;
}
function parseTextBox(txBody, x, y, w, h, themeColors, spPr, hyperlinks = /* @__PURE__ */ new Map()) {
  const runs = [];
  let bgFill;
  if (spPr) {
    const solidFill = getEl(spPr, "solidFill");
    if (solidFill) bgFill = readColour(solidFill, themeColors) ?? void 0;
  }
  getEl(txBody, "lstStyle");
  getEl(txBody, "bodyPr");
  const paragraphs = getEls(txBody, "p");
  let listCounter = 0;
  let lastIndent = -1;
  for (const para of paragraphs) {
    const pPr = getEl(para, "pPr");
    const indent = pPr ? numAttr(pPr, "lvl", 0) : 0;
    if (indent !== lastIndent) {
      listCounter = 0;
      lastIndent = indent;
    }
    let prefix = "";
    if (pPr) {
      const buNone = getEl(pPr, "buNone");
      const buChar = getEl(pPr, "buChar");
      const buAutoNum = getEl(pPr, "buAutoNum");
      getEl(pPr, "buFont");
      if (!buNone) {
        if (buChar) {
          const raw = attr(buChar, "char");
          const isSafeChar = raw && raw.charCodeAt(0) >= 32 && raw.charCodeAt(0) <= 126;
          prefix = isSafeChar ? raw + " " : "• ";
        } else if (buAutoNum) {
          listCounter++;
          const type = attr(buAutoNum, "type");
          if (type.startsWith("alphaLc")) {
            prefix = String.fromCharCode(96 + listCounter) + ". ";
          } else if (type.startsWith("alphaUc")) {
            prefix = String.fromCharCode(64 + listCounter) + ". ";
          } else if (type.includes("ParenR")) {
            prefix = listCounter + ") ";
          } else {
            prefix = listCounter + ". ";
          }
        }
      }
    }
    const firstRpr = getEl(getEls(para, "r")[0] ?? para, "rPr");
    const defaultSize = firstRpr ? halfPtToPt(numAttr(firstRpr, "sz", 1800)) : 18;
    const defaultColor = readRunColour(firstRpr, themeColors) ?? (bgFill ? WHITE : BLACK);
    if (prefix) {
      runs.push({
        text: prefix,
        fontSize: defaultSize,
        fontFamily: "sans-serif",
        fontWeight: "normal",
        fontStyle: "normal",
        color: defaultColor,
        underline: false,
        strikethrough: false
      });
    }
    for (const r of getEls(para, "r")) {
      const rPr = getEl(r, "rPr");
      const t = getEl(r, "t");
      if (!(t == null ? void 0 : t.textContent)) continue;
      let color = readRunColour(rPr, themeColors);
      if (!color) {
        color = bgFill ? WHITE : BLACK;
      }
      let url;
      if (rPr) {
        const hlink = getEl(rPr, "hlinkClick");
        if (hlink) {
          const rId = attr(hlink, "r:id") || attr(hlink, "id");
          url = hyperlinks.get(rId);
        }
      }
      runs.push({
        text: t.textContent,
        fontSize: rPr ? halfPtToPt(numAttr(rPr, "sz", 1800)) : 18,
        fontFamily: readFontFamily(rPr) ?? "sans-serif",
        fontWeight: rPr && attr(rPr, "b") === "1" ? "bold" : "normal",
        fontStyle: rPr && attr(rPr, "i") === "1" ? "italic" : "normal",
        color,
        underline: rPr ? attr(rPr, "u") !== "none" && attr(rPr, "u") !== "" : false,
        strikethrough: rPr ? attr(rPr, "strike") === "sngStrike" : false,
        ...url ? { url } : {}
      });
    }
    runs.push({
      text: "\n",
      fontSize: defaultSize,
      fontFamily: "sans-serif",
      fontWeight: "normal",
      fontStyle: "normal",
      color: TRANSPARENT,
      underline: false,
      strikethrough: false
    });
  }
  return {
    type: "text",
    x,
    y,
    width: w,
    height: h,
    align: "left",
    runs,
    lineHeight: 1.2,
    // Pass background colour as hint so generator can draw box bg
    ...bgFill ? { bgFill } : {}
  };
}
function parsePicture(pic, imageMap) {
  const blip = getEl(pic, "blip");
  if (!blip) return null;
  const rId = attr(blip, "r:embed") || attr(blip, "embed");
  if (!rId) return null;
  const src = imageMap.get(rId);
  if (!src) return null;
  const xfrm = getEl(pic, "xfrm");
  if (!xfrm) return null;
  const off = getEl(xfrm, "off");
  const ext = getEl(xfrm, "ext");
  if (!off || !ext) return null;
  return {
    type: "image",
    x: emuToPt(numAttr(off, "x")),
    y: emuToPt(numAttr(off, "y")),
    width: emuToPt(numAttr(ext, "cx")),
    height: emuToPt(numAttr(ext, "cy")),
    src
  };
}
function parseTable(graphicFrame, tbl, themeColors) {
  const xfrm = getEl(graphicFrame, "xfrm");
  if (!xfrm) return null;
  const off = getEl(xfrm, "off");
  const ext = getEl(xfrm, "ext");
  if (!off || !ext) return null;
  const x = emuToPt(numAttr(off, "x"));
  const y = emuToPt(numAttr(off, "y"));
  const w = emuToPt(numAttr(ext, "cx"));
  const colWidths = getEls(tbl, "gridCol").map((c) => emuToPt(numAttr(c, "w")));
  const rows = [];
  for (const tr of getEls(tbl, "tr")) {
    const rowHeight = emuToPt(numAttr(tr, "h"));
    const cells = [];
    for (const tc of getEls(tr, "tc")) {
      const texts = getEls(tc, "t").map((t) => t.textContent ?? "").join("");
      const tcPr = getEl(tc, "tcPr");
      const colspan = numAttr(tc, "gridSpan", 1);
      const rowspan = numAttr(tc, "rowSpan", 1);
      const solidFill = tcPr ? getEl(tcPr, "solidFill") : null;
      const style = {
        fill: solidFill ? readColour(solidFill, themeColors) ?? void 0 : void 0,
        align: "left"
      };
      cells.push({ value: texts, colspan, rowspan, style });
    }
    rows.push({ cells, height: rowHeight });
  }
  return { type: "table", x, y, width: w, colWidths, rows };
}
function readColour(parent, themeColors = /* @__PURE__ */ new Map()) {
  const srgb = getEl(parent, "srgbClr");
  if (srgb) {
    const rgba = hexToRgba(attr(srgb, "val"));
    return applyAlpha(rgba, srgb);
  }
  const schemeClr = getEl(parent, "schemeClr");
  if (schemeClr) {
    const rgba = resolveSchemeColor(schemeClr, themeColors);
    return applyAlpha(rgba, schemeClr);
  }
  return null;
}
function applyAlpha(color, el) {
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i];
    if (child.localName === "alpha") {
      const val = numAttr(child, "val", 1e5);
      return { ...color, a: Math.min(1, val / 1e5) };
    }
  }
  return color;
}
function readRunColour(rPr, themeColors = /* @__PURE__ */ new Map()) {
  if (!rPr) return null;
  const solidFill = getEl(rPr, "solidFill");
  return solidFill ? readColour(solidFill, themeColors) : null;
}
function readFontFamily(rPr) {
  if (!rPr) return null;
  const latin = getEl(rPr, "latin");
  return latin ? attr(latin, "typeface") || null : null;
}
async function parseXlsx(file) {
  const zip = await loadZip(file);
  if (isParseError(zip)) return { ok: false, error: zip };
  const sharedStrings = await loadSharedStrings(zip);
  const styles = await loadStyles(zip);
  const wbXml = await zip.getText("xl/workbook.xml");
  if (!wbXml) return { ok: false, error: { code: "MISSING_ENTRY", message: "xl/workbook.xml not found" } };
  const wbDoc = parseXml(wbXml);
  if (!wbDoc) return { ok: false, error: { code: "XML_PARSE_FAILED", message: "Could not parse workbook.xml" } };
  const sheets = getEls(wbDoc, "sheet").map((s) => ({
    name: attr(s, "name"),
    rId: attr(s, "r:id") || attr(s, "id")
  }));
  const relsXml = await zip.getText("xl/_rels/workbook.xml.rels");
  const pathMap = buildSheetPathMap(relsXml);
  const pages = [];
  for (const sheet of sheets) {
    const path = pathMap.get(sheet.rId);
    if (!path) continue;
    const sheetXml = await zip.getText(`xl/${path}`);
    if (!sheetXml) continue;
    const sheetDoc = parseXml(sheetXml);
    if (!sheetDoc) continue;
    const page = parseSheet(sheetDoc, sharedStrings, styles, sheet.name);
    pages.push(page);
  }
  return {
    ok: true,
    doc: { format: "xlsx", pages, metadata: {} }
  };
}
async function loadSharedStrings(zip) {
  const xml = await zip.getText("xl/sharedStrings.xml");
  if (!xml) return [];
  const doc = parseXml(xml);
  if (!doc) return [];
  return getEls(doc, "si").map((si) => {
    const runs = getEls(si, "t");
    return runs.map((t) => t.textContent ?? "").join("");
  });
}
async function loadStyles(zip) {
  const empty = { fills: [], fonts: [], borders: [], cellXfs: [] };
  const xml = await zip.getText("xl/styles.xml");
  if (!xml) return empty;
  const doc = parseXml(xml);
  if (!doc) return empty;
  const fills = getEls(doc, "fill").map((fill) => {
    const fgColor = getEl(fill, "fgColor");
    const rgb = fgColor ? attr(fgColor, "rgb") : "";
    return { fgColor: rgb ? hexToRgba(rgb.slice(2)) : void 0 };
  });
  const fonts = getEls(doc, "font").map((font) => ({
    bold: !!getEl(font, "b"),
    italic: !!getEl(font, "i"),
    color: (() => {
      const c = getEl(font, "color");
      const rgb = c ? attr(c, "rgb") : "";
      return rgb ? hexToRgba(rgb.slice(2)) : void 0;
    })(),
    size: numAttr(getEl(font, "sz"), "val", 11),
    name: attr(getEl(font, "name"), "val") || "Calibri"
  }));
  const borders = getEls(doc, "border").map((b) => ({
    top: !!getEl(b, "top"),
    bottom: !!getEl(b, "bottom"),
    left: !!getEl(b, "left"),
    right: !!getEl(b, "right")
  }));
  const cellXfs = getEls(getEl(doc, "cellXfs"), "xf").map((xf) => ({
    fontId: numAttr(xf, "fontId"),
    fillId: numAttr(xf, "fillId"),
    borderId: numAttr(xf, "borderId"),
    numFmtId: numAttr(xf, "numFmtId"),
    alignH: (() => {
      const align = getEl(xf, "alignment");
      return align ? attr(align, "horizontal") : void 0;
    })()
  }));
  return { fills, fonts, borders, cellXfs };
}
function buildSheetPathMap(relsXml) {
  const map = /* @__PURE__ */ new Map();
  if (!relsXml) return map;
  const doc = parseXml(relsXml);
  if (!doc) return map;
  for (const rel of getEls(doc, "Relationship")) {
    if (!attr(rel, "Type").includes("worksheet")) continue;
    map.set(attr(rel, "Id"), attr(rel, "Target"));
  }
  return map;
}
function parseSheet(sheetDoc, sharedStrings, styles, sheetName) {
  var _a;
  const cellGrid = /* @__PURE__ */ new Map();
  let maxRow = 0;
  let maxCol = 0;
  for (const row of getEls(sheetDoc, "row")) {
    const rowIdx = numAttr(row, "r", 0) - 1;
    for (const c of getEls(row, "c")) {
      const ref = attr(c, "r");
      const { col } = cellRefToIndex(ref);
      const type = attr(c, "t");
      const styleId = numAttr(c, "s");
      const vEl = getEl(c, "v");
      const raw = (vEl == null ? void 0 : vEl.textContent) ?? "";
      let value = raw;
      if (type === "s") {
        value = sharedStrings[parseInt(raw)] ?? "";
      } else if (type === "b") {
        value = raw === "1" ? "TRUE" : "FALSE";
      } else if (type === "e") {
        value = raw;
      }
      const cellStyle = resolveStyle(styles, styleId);
      if (!cellGrid.has(rowIdx)) cellGrid.set(rowIdx, /* @__PURE__ */ new Map());
      cellGrid.get(rowIdx).set(col, { value, style: cellStyle });
      maxRow = Math.max(maxRow, rowIdx);
      maxCol = Math.max(maxCol, col);
    }
  }
  const numCols = maxCol + 1;
  const colWidth = cmToPt(2.5);
  const tableRows = [];
  for (let r = 0; r <= maxRow; r++) {
    const cells = [];
    for (let c = 0; c < numCols; c++) {
      const cell = (_a = cellGrid.get(r)) == null ? void 0 : _a.get(c);
      cells.push({
        value: (cell == null ? void 0 : cell.value) ?? "",
        colspan: 1,
        rowspan: 1,
        style: (cell == null ? void 0 : cell.style) ?? {}
      });
    }
    tableRows.push({ cells, height: cmToPt(0.6) });
  }
  const tableWidth = numCols * colWidth;
  const tableHeight = tableRows.reduce((s, r) => s + r.height, 0);
  const margin = cmToPt(1);
  const pageWidth = Math.max(tableWidth + margin * 2, 595.28);
  const pageHeight = Math.max(tableHeight + margin * 2, 841.89);
  const table = {
    type: "table",
    x: margin,
    y: margin,
    width: tableWidth,
    colWidths: Array(numCols).fill(colWidth),
    rows: tableRows
  };
  return {
    width: pageWidth,
    height: pageHeight,
    background: WHITE,
    elements: [table],
    label: sheetName
  };
}
function cellRefToIndex(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { col: 0, row: 0 };
  const colStr = match[1] ?? "";
  const rowStr = match[2] ?? "1";
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: parseInt(rowStr) - 1 };
}
function resolveStyle(styles, styleId) {
  const xf = styles.cellXfs[styleId];
  if (!xf) return {};
  const font = styles.fonts[xf.fontId];
  const fill = styles.fills[xf.fillId];
  const border = styles.borders[xf.borderId];
  const alignMap = {
    left: "left",
    center: "center",
    right: "right",
    general: "left"
  };
  return {
    fill: fill == null ? void 0 : fill.fgColor,
    color: font == null ? void 0 : font.color,
    fontSize: font == null ? void 0 : font.size,
    fontWeight: (font == null ? void 0 : font.bold) ? "bold" : "normal",
    fontStyle: (font == null ? void 0 : font.italic) ? "italic" : "normal",
    align: alignMap[xf.alignH ?? "general"] ?? "left",
    borderTop: border == null ? void 0 : border.top,
    borderBottom: border == null ? void 0 : border.bottom,
    borderLeft: border == null ? void 0 : border.left,
    borderRight: border == null ? void 0 : border.right
  };
}
function generatePDF(doc, opts = {}) {
  if (doc.pages.length === 0) throw new Error("Document has no pages");
  const firstPage = doc.pages[0];
  const pdf = new jsPDF({
    orientation: firstPage.width > firstPage.height ? "landscape" : "portrait",
    unit: "pt",
    format: [firstPage.width, firstPage.height]
  });
  if (opts.title) pdf.setProperties({ title: opts.title });
  if (opts.author) pdf.setProperties({ author: opts.author });
  doc.pages.forEach((page, i) => {
    if (i > 0) pdf.addPage([page.width, page.height], page.width > page.height ? "landscape" : "portrait");
    renderPage(pdf, page);
    if (opts.pageNumbers) renderPageNumber(pdf, i + 1, doc.pages.length, page);
  });
  return {
    download(filename = "output.pdf") {
      pdf.save(filename);
    },
    toBlob() {
      return pdf.output("blob");
    },
    toDataUrl() {
      return pdf.output("datauristring");
    }
  };
}
function renderPage(pdf, page) {
  if (page.background.a > 0) {
    pdf.setFillColor(page.background.r, page.background.g, page.background.b);
    pdf.rect(0, 0, page.width, page.height, "F");
  }
  const bgMap = buildBackgroundMap(page);
  for (const el of page.elements) {
    if (el.type === "shape") renderShape(pdf, el);
  }
  for (const el of page.elements) {
    if (el.type === "image") renderImage(pdf, el);
  }
  for (const el of page.elements) {
    if (el.type === "table") renderTable(pdf, el);
  }
  for (let i = 0; i < page.elements.length; i++) {
    const el = page.elements[i];
    if (el.type === "text") {
      renderText(pdf, el, bgMap.get(i) ?? page.background);
    }
  }
}
function buildBackgroundMap(page) {
  const map = /* @__PURE__ */ new Map();
  const filledShapes = page.elements.filter((el) => el.type === "shape").map((el) => el).filter((s) => s.fill && s.fill.a > 0.1 && s.kind !== "line" && s.kind !== "arrow");
  for (let i = 0; i < page.elements.length; i++) {
    const el = page.elements[i];
    if (el.type !== "text") continue;
    const t = el;
    let bestFill = null;
    let bestArea = 0;
    for (const shape of filledShapes) {
      const area = overlapArea(
        t.x,
        t.y,
        t.x + t.width,
        t.y + t.height,
        shape.x,
        shape.y,
        shape.x + shape.width,
        shape.y + shape.height
      );
      if (area > bestArea) {
        bestArea = area;
        bestFill = shape.fill;
      }
    }
    if (bestFill && bestArea > 0) map.set(i, bestFill);
  }
  return map;
}
function overlapArea(ax1, ay1, ax2, ay2, bx1, by1, bx2, by2) {
  const ix1 = Math.max(ax1, bx1);
  const iy1 = Math.max(ay1, by1);
  const ix2 = Math.min(ax2, bx2);
  const iy2 = Math.min(ay2, by2);
  if (ix2 <= ix1 || iy2 <= iy1) return 0;
  return (ix2 - ix1) * (iy2 - iy1);
}
function linearize(c) {
  const s = c / 255;
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
}
function luminance(c) {
  return 0.2126 * linearize(c.r) + 0.7152 * linearize(c.g) + 0.0722 * linearize(c.b);
}
function contrastRatio(a, b) {
  const la = luminance(a);
  const lb = luminance(b);
  const lighter = Math.max(la, lb);
  const darker = Math.min(la, lb);
  return (lighter + 0.05) / (darker + 0.05);
}
const RGBA_BLACK = { r: 0, g: 0, b: 0, a: 1 };
const RGBA_WHITE = { r: 255, g: 255, b: 255, a: 1 };
function ensureContrast(textColor, bg) {
  if (textColor.a < 0.5) return textColor;
  if (contrastRatio(textColor, bg) >= 3) return textColor;
  const blackContrast = contrastRatio(RGBA_BLACK, bg);
  const whiteContrast = contrastRatio(RGBA_WHITE, bg);
  return whiteContrast >= blackContrast ? RGBA_WHITE : RGBA_BLACK;
}
function renderText(pdf, el, effectiveBg = RGBA_WHITE) {
  let result = el.y;
  try {
    result = _renderTextInner(pdf, el, effectiveBg);
  } catch (e) {
    console.error("[pdfkit-client] renderText failed:", e, "\nRuns:", el.runs.length, "at", el.x, el.y);
  }
  return result;
}
function _renderTextInner(pdf, el, effectiveBg = RGBA_WHITE) {
  if (el.runs.length === 0) return el.y;
  const bgFill = el.bgFill;
  if (bgFill && bgFill.a > 0.05) {
    pdf.setFillColor(bgFill.r, bgFill.g, bgFill.b);
    pdf.rect(el.x, el.y, el.width, el.height, "F");
  }
  const activeBg = bgFill && bgFill.a > 0.05 ? bgFill : effectiveBg;
  const textRuns = el.runs.filter((r) => r.text !== "\n");
  const allFaded = textRuns.length > 0 && textRuns.every((r) => r.color.a < 0.5);
  if (allFaded && !bgFill) {
    console.debug(`[faded] skipping element at y=${el.y} — all runs alpha<0.5`);
    return el.y;
  }
  const paragraphs = [];
  let current = [];
  for (const run of el.runs) {
    if (run.text === "\n") {
      paragraphs.push(current);
      current = [];
    } else current.push(run);
  }
  if (current.length > 0) paragraphs.push(current);
  let cursorY = el.y;
  for (const para of paragraphs) {
    if (para.length === 0) {
      cursorY += 14 * el.lineHeight;
      continue;
    }
    const paraSize = para.reduce((max, r) => Math.max(max, r.fontSize), 12);
    const lineH = paraSize * el.lineHeight;
    const baseline = cursorY + paraSize;
    let cursorX = el.x;
    let maxLinesInPara = 1;
    for (const run of para) {
      if (!run.text) continue;
      const text = sanitizeText(run.text);
      if (!text) continue;
      const fontStyle = getFontStyle(run.fontWeight, run.fontStyle);
      try {
        pdf.setFont(mapFont(run.fontFamily), fontStyle);
      } catch {
        pdf.setFont("helvetica", fontStyle);
      }
      const safeFontSize = run.fontSize > 0 && isFinite(run.fontSize) ? run.fontSize : 12;
      pdf.setFontSize(safeFontSize);
      if (run.color.a < 0.5) continue;
      const safeColor = ensureContrast(run.color, activeBg);
      pdf.setTextColor(safeColor.r, safeColor.g, safeColor.b);
      const availWidth = el.width - (cursorX - el.x);
      if (availWidth < 4) {
        cursorX = el.x;
        cursorY += lineH;
        continue;
      }
      const lines = pdf.splitTextToSize(text, availWidth).filter((l) => typeof l === "string" && l.length > 0);
      if (lines.length === 0) continue;
      maxLinesInPara = Math.max(maxLinesInPara, lines.length);
      if (lines.length === 1) {
        if (typeof lines[0] !== "string" || lines[0].length === 0) {
          continue;
        }
        if (!isFinite(cursorX) || !isFinite(baseline)) {
          continue;
        }
        pdf.text(lines[0], cursorX, baseline);
        const lineW = pdf.getStringUnitWidth(lines[0]) * (safeFontSize / pdf.internal.scaleFactor);
        if (run.url) {
          pdf.link(cursorX, baseline - run.fontSize, lineW, run.fontSize * 1.2, { url: run.url });
        }
        cursorX += lineW;
      } else {
        if (typeof lines[0] === "string" && lines[0].length > 0 && isFinite(cursorX) && isFinite(baseline)) {
          pdf.text(lines[0], cursorX, baseline);
        }
        for (let li = 1; li < lines.length; li++) {
          const lineText = lines[li];
          const lineY = baseline + li * lineH;
          if (typeof lineText === "string" && lineText.length > 0 && isFinite(el.x) && isFinite(lineY)) {
            pdf.text(lineText, el.x, lineY);
          }
        }
        const lastLine = lines[lines.length - 1];
        cursorX = el.x + pdf.getStringUnitWidth(lastLine) * (safeFontSize / pdf.internal.scaleFactor);
        maxLinesInPara = Math.max(maxLinesInPara, lines.length);
      }
    }
    cursorY += lineH * maxLinesInPara;
  }
  pdf.setTextColor(0, 0, 0);
  return cursorY;
}
function sanitizeText(text) {
  return text.replace(/[\uF000-\uF0FF\u0080-\u009F]/g, (ch) => {
    if (ch >= "" && ch <= "") return "• ";
    return "";
  });
}
function getFontStyle(weight, style) {
  if (weight === "bold" && style === "italic") return "bolditalic";
  if (weight === "bold") return "bold";
  if (style === "italic") return "italic";
  return "normal";
}
function mapFont(family) {
  const lower = family.toLowerCase();
  if (lower.includes("arial") || lower.includes("helvetica") || lower.includes("calibri")) return "helvetica";
  if (lower.includes("times") || lower.includes("georgia") || lower.includes("garamond")) return "times";
  if (lower.includes("courier") || lower.includes("consolas") || lower.includes("mono")) return "courier";
  return "helvetica";
}
function renderImage(pdf, el) {
  try {
    const format = el.src.includes("png") ? "PNG" : el.src.includes("jpeg") || el.src.includes("jpg") ? "JPEG" : "PNG";
    pdf.addImage(el.src, format, el.x, el.y, el.width, el.height);
  } catch (e) {
    console.warn("[pdfkit-client] Failed to render image:", e);
  }
}
function renderShape(pdf, el) {
  const hasFill = el.fill && el.fill.a > 0.05;
  const hasStroke = el.stroke && el.stroke.a > 0.05;
  if (!hasFill && !hasStroke) return;
  if (hasFill) pdf.setFillColor(el.fill.r, el.fill.g, el.fill.b);
  if (hasStroke) {
    pdf.setDrawColor(el.stroke.r, el.stroke.g, el.stroke.b);
    pdf.setLineWidth(el.strokeWidth || 0.75);
  }
  const drawMode = hasFill && hasStroke ? "FD" : hasFill ? "F" : "D";
  switch (el.kind) {
    case "ellipse":
      pdf.ellipse(el.x + el.width / 2, el.y + el.height / 2, el.width / 2, el.height / 2, drawMode);
      break;
    case "line":
    case "arrow": {
      if (!hasStroke) break;
      const flipH = el.flipH === true;
      const flipV = el.flipV === true;
      const sx = flipH ? el.x + el.width : el.x;
      const sy = flipV ? el.y + el.height : el.y;
      const ex = flipH ? el.x : el.x + el.width;
      const ey = flipV ? el.y : el.y + el.height;
      pdf.line(sx, sy, ex, ey);
      if (el.kind === "arrow") {
        const angle = Math.atan2(ey - sy, ex - sx);
        const headLen = Math.min(8, Math.sqrt((ex - sx) ** 2 + (ey - sy) ** 2) * 0.25);
        pdf.line(ex, ey, ex - headLen * Math.cos(angle - 0.4), ey - headLen * Math.sin(angle - 0.4));
        pdf.line(ex, ey, ex - headLen * Math.cos(angle + 0.4), ey - headLen * Math.sin(angle + 0.4));
      }
      break;
    }
    default:
      pdf.rect(el.x, el.y, el.width, el.height, drawMode);
  }
}
function renderTable(pdf, el) {
  let rowY = el.y;
  for (const row of el.rows) {
    let cellX = el.x;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      const cellWidth = el.colWidths[c] ?? 50;
      if (cell.style.fill && cell.style.fill.a > 0.05) {
        pdf.setFillColor(cell.style.fill.r, cell.style.fill.g, cell.style.fill.b);
        pdf.rect(cellX, rowY, cellWidth, row.height, "F");
      }
      pdf.setDrawColor(200, 200, 200);
      pdf.setLineWidth(0.5);
      pdf.rect(cellX, rowY, cellWidth, row.height, "D");
      if (cell.value) {
        const fontSize = cell.style.fontSize ?? 10;
        const fontWeight = cell.style.fontWeight ?? "normal";
        const fontStyle = cell.style.fontStyle ?? "normal";
        pdf.setFont("helvetica", getFontStyle(fontWeight, fontStyle));
        pdf.setFontSize(fontSize);
        const color = cell.style.color;
        if (color) pdf.setTextColor(color.r, color.g, color.b);
        else pdf.setTextColor(0, 0, 0);
        const padding = 3;
        pdf.text(
          cell.value,
          cellX + padding,
          rowY + row.height / 2 + fontSize / 3,
          { maxWidth: cellWidth - padding * 2 }
        );
      }
      cellX += cellWidth;
    }
    rowY += row.height;
  }
}
function renderPageNumber(pdf, current, total, page) {
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(9);
  pdf.setTextColor(150, 150, 150);
  pdf.text(`${current} / ${total}`, page.width / 2, page.height - 12, { align: "center" });
  pdf.setTextColor(0, 0, 0);
}
async function toPDF(file, opts = {}) {
  try {
    const parseResult = await toIR(file);
    if (!parseResult.ok) {
      return { ok: false, error: parseResult.error.message };
    }
    let doc = parseResult.doc;
    if (opts.pages && opts.pages.length > 0) {
      doc = {
        ...doc,
        pages: opts.pages.filter((i) => i >= 0 && i < doc.pages.length).map((i) => doc.pages[i])
      };
    }
    const pdf = generatePDF(doc, opts);
    return { ok: true, pdf };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
async function toIR(file) {
  var _a;
  const ext = (_a = file.name.split(".").pop()) == null ? void 0 : _a.toLowerCase();
  switch (ext) {
    case "pptx":
      return parsePptx(file);
    case "xlsx":
      return parseXlsx(file);
    default:
      return {
        ok: false,
        error: {
          code: "UNSUPPORTED_FORMAT",
          message: `Unsupported file format: .${ext}`,
          detail: "Supported formats: .pptx, .xlsx"
        }
      };
  }
}
const SUPPORTED_FORMATS = [".pptx", ".xlsx"];
function isSupportedFile(file) {
  return SUPPORTED_FORMATS.some((ext) => file.name.toLowerCase().endsWith(ext));
}
export {
  SUPPORTED_FORMATS,
  isSupportedFile,
  toIR,
  toPDF
};
//# sourceMappingURL=index.esm.js.map
