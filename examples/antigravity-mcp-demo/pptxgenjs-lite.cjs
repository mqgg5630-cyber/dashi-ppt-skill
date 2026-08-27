/*
 * Tiny, dependency-free subset of the PptxGenJS API used by the sample
 * exporter. It writes Open XML directly so the checked-in examples can be
 * regenerated in a fresh clone even when npm cannot reach a registry.
 *
 * This is not a general-purpose PowerPoint library. It intentionally covers
 * only native text boxes, rectangles, rounded rectangles, ellipses, and lines.
 */

'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

const ShapeType = Object.freeze({
  rect: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  line: 'line',
});

class PptxGenJS {
  constructor() {
    this.ShapeType = ShapeType;
    this.layout = 'LAYOUT_WIDE';
    this.author = '';
    this.company = '';
    this.subject = '';
    this.title = '';
    this.lang = 'zh-CN';
    this.theme = {};
    this._masters = new Map();
    this._slides = [];
    this._nextShapeId = 2;
  }

  defineSlideMaster(master) {
    this._masters.set(String(master.title || 'MASTER'), master);
  }

  addSlide(masterName) {
    const master = this._masters.get(String(masterName || '')) || {};
    const slide = new LiteSlide(this, master);
    this._slides.push(slide);
    return slide;
  }

  async writeFile({ fileName }) {
    const output = path.resolve(String(fileName));
    fs.mkdirSync(path.dirname(output), { recursive: true });
    const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'dashi-pptx-'));
    try {
      this._writePackage(temp);
      writeZip(temp, output);
    } finally {
      fs.rmSync(temp, { recursive: true, force: true });
    }
  }

  _writePackage(root) {
    const master = this._masters.values().next().value || {};
    const dirs = [
      '_rels', 'docProps', 'ppt', 'ppt/_rels', 'ppt/slideMasters',
      'ppt/slideMasters/_rels', 'ppt/slideLayouts', 'ppt/slideLayouts/_rels',
      'ppt/slides', 'ppt/slides/_rels', 'ppt/theme',
    ];
    dirs.forEach((dir) => fs.mkdirSync(path.join(root, dir), { recursive: true }));

    write(root, '[Content_Types].xml', contentTypes(this._slides.length));
    write(root, '_rels/.rels', rootRels());
    write(root, 'docProps/core.xml', coreProps(this));
    write(root, 'docProps/app.xml', appProps(this._slides.length));
    write(root, 'ppt/presentation.xml', presentationXml(this._slides.length));
    write(root, 'ppt/_rels/presentation.xml.rels', presentationRels(this._slides.length));
    write(root, 'ppt/theme/theme1.xml', themeXml());
    write(root, 'ppt/slideMasters/slideMaster1.xml', slideMasterXml(master));
    write(root, 'ppt/slideMasters/_rels/slideMaster1.xml.rels', slideMasterRels());
    write(root, 'ppt/slideLayouts/slideLayout1.xml', slideLayoutXml());
    write(root, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', slideLayoutRels());

    this._slides.forEach((slide, index) => {
      const number = index + 1;
      write(root, `ppt/slides/slide${number}.xml`, slideXml(slide, number));
      write(root, `ppt/slides/_rels/slide${number}.xml.rels`, slideRels());
    });
  }
}

class LiteSlide {
  constructor(pptx, master) {
    this._pptx = pptx;
    this._master = master || {};
    this._objects = [];
  }

  addText(text, options = {}) {
    this._objects.push({ kind: 'text', text: String(text == null ? '' : text), options: { ...options } });
  }

  addShape(shapeType, options = {}) {
    this._objects.push({ kind: 'shape', shapeType: String(shapeType), options: { ...options } });
  }
}

PptxGenJS.ShapeType = ShapeType;
module.exports = PptxGenJS;

function write(root, relative, content) {
  const target = path.join(root, relative);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, content);
}

function esc(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function safeColor(value, fallback = '000000') {
  const color = String(value || fallback).replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(color) ? color : fallback;
}

function emu(value) {
  return Math.round(Number(value || 0) * 914400);
}

function alpha(transparency) {
  const t = Math.max(0, Math.min(100, Number(transparency || 0)));
  return Math.round((100 - t) * 1000);
}

function xfrm(options, isLine = false) {
  let x = Number(options.x || 0);
  let y = Number(options.y || 0);
  let w = Number(options.w || 0);
  let h = Number(options.h || 0);
  const flipH = w < 0;
  const flipV = h < 0;
  if (flipH) { x += w; w = Math.abs(w); }
  if (flipV) { y += h; h = Math.abs(h); }
  const flips = `${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
  return `<a:xfrm${flips}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${Math.max(1, emu(w))}" cy="${Math.max(1, emu(h))}"/></a:xfrm>`;
}

function solidFill(fill, fallback, defaultTransparency = 0) {
  if (!fill || fill.color === undefined) return `<a:solidFill><a:srgbClr val="${safeColor(fallback)}"/></a:solidFill>`;
  const color = safeColor(fill.color, safeColor(fallback));
  const t = fill.transparency === undefined ? defaultTransparency : fill.transparency;
  return `<a:solidFill><a:srgbClr val="${color}"><a:alpha val="${alpha(t)}"/></a:srgbClr></a:solidFill>`;
}

function lineXml(line, fallback) {
  if (!line || Number(line.transparency) >= 100) return '<a:ln w="0"><a:noFill/></a:ln>';
  const width = Math.max(1, Math.round(Number(line.width || 1) * 12700));
  const color = safeColor(line.color, safeColor(fallback));
  const t = Number(line.transparency || 0);
  const dash = line.dash ? ` prstDash="${dashName(line.dash)}"` : '';
  const head = line.beginArrowType ? `<a:headEnd type="${arrowName(line.beginArrowType)}" w="sm" len="sm"/>` : '';
  const tail = line.endArrowType ? `<a:tailEnd type="${arrowName(line.endArrowType)}" w="sm" len="sm"/>` : '';
  return `<a:ln w="${width}" cap="flat"${dash}><a:solidFill><a:srgbClr val="${color}"><a:alpha val="${alpha(t)}"/></a:srgbClr></a:solidFill>${head}${tail}</a:ln>`;
}

function dashName(value) {
  const s = String(value).toLowerCase();
  if (s.includes('dot')) return 'sysDot';
  if (s.includes('dash')) return 'dash';
  return 'solid';
}

function arrowName(value) {
  const s = String(value).toLowerCase();
  if (s.includes('triangle')) return 'triangle';
  if (s.includes('stealth')) return 'stealth';
  return 'none';
}

function geometry(shapeType) {
  if (shapeType === ShapeType.roundRect) return 'roundRect';
  if (shapeType === ShapeType.ellipse) return 'ellipse';
  if (shapeType === ShapeType.line) return 'line';
  return 'rect';
}

function shapeXml(obj, id, name) {
  const o = obj.options || {};
  const type = geometry(obj.shapeType);
  const isLine = type === 'line';
  const fill = o.fill || {};
  const line = o.line || {};
  const fillXml = isLine ? '<a:noFill/>' : solidFill(fill, 'FFFFFF');
  const lineXmlValue = lineXml(line, fill.color || 'FFFFFF');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(o, isLine)}<a:prstGeom prst="${type}"><a:avLst/></a:prstGeom>${fillXml}${lineXmlValue}</p:spPr></p:sp>`;
}

function textXml(obj, id, name) {
  const o = obj.options || {};
  const text = obj.text;
  const color = safeColor(o.color, '000000');
  const fontSize = Math.max(1, Math.round(Number(o.fontSize || 12) * 100));
  const bold = o.bold ? ' b="1"' : '';
  const italic = o.italic ? ' i="1"' : '';
  const align = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[o.align] || 'l';
  const valign = { top: 't', mid: 'ctr', bottom: 'b' }[o.valign] || 'ctr';
  const face = esc(o.fontFace || 'Aptos');
  const alphaValue = alpha(o.transparency || 0);
  const colorXml = `<a:solidFill><a:srgbClr val="${color}"><a:alpha val="${alphaValue}"/></a:srgbClr></a:solidFill>`;
  const runs = String(text).split(/\n/).map((part, index) => (
    `${index ? '<a:br/>' : ''}<a:r><a:rPr lang="zh-CN" sz="${fontSize}"${bold}${italic}>${colorXml}<a:latin typeface="${face}"/><a:ea typeface="${face}"/><a:cs typeface="${face}"/></a:rPr><a:t>${esc(part)}</a:t></a:r>`
  )).join('');
  const margin = Number(o.margin || 0);
  const mar = Math.max(0, Math.round(margin * 914400));
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="0"><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${valign}" lIns="${mar}" rIns="${mar}" tIns="${mar}" bIns="${mar}"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${align}"><a:defRPr sz="${fontSize}"${bold}${italic}>${colorXml}<a:latin typeface="${face}"/><a:ea typeface="${face}"/><a:cs typeface="${face}"/></a:defRPr></a:pPr>${runs}</a:p></p:txBody></p:sp>`;
}

function groupXml(slide) {
  const objects = slide._objects.map((obj, index) => {
    const id = slide._pptx._nextShapeId++;
    return obj.kind === 'text'
      ? textXml(obj, id, `Text ${id}`)
      : shapeXml(obj, id, `Shape ${id}`);
  }).join('');
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${objects}</p:spTree>`;
}

function slideXml(slide, number) {
  const masterBg = slide._master.background && slide._master.background.color;
  const bg = masterBg ? `<p:bg><p:bgPr>${solidFill({ color: masterBg }, masterBg)}<a:effectLst/></p:bgPr></p:bg>` : '';
  const slideNumber = slide._master.slideNumber;
  let extra = '';
  if (slideNumber) {
    extra = textXml({ kind: 'text', text: String(number), options: { ...slideNumber, color: slideNumber.color || '808080', fontSize: slideNumber.fontSize || 8, x: slideNumber.x, y: slideNumber.y, w: 0.45, h: 0.16, align: 'right', valign: 'mid' } }, 100000 + number, `Slide number ${number}`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">${bg}<p:cSld name="Slide ${number}">${groupXml(slide).replace('</p:spTree>', `${extra}</p:spTree>`)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function contentTypes(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${slides}</Types>`;
}

function rootRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

function presentationXml(slideCount) {
  const ids = Array.from({ length: slideCount }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr/><a:lvl1pPr marL="0" algn="l"/></p:defaultTextStyle></p:presentation>`;
}

function presentationRels(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>${slides}</Relationships>`;
}

function slideMasterXml(master) {
  const color = safeColor(master.background && master.background.color, 'FFFFFF');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld name="Master"><p:bg><p:bgPr>${solidFill({ color }, color)}<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>`;
}

function slideMasterRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`;
}

function slideLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function slideLayoutRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;
}

function slideRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`;
}

function themeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${NS.a}" name="Dashi PPT Native Theme"><a:themeElements><a:clrScheme name="Dashi"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F1F1F"/></a:dk2><a:lt2><a:srgbClr val="F7F7F7"/></a:lt2><a:accent1><a:srgbClr val="5B8DEF"/></a:accent1><a:accent2><a:srgbClr val="46B083"/></a:accent2><a:accent3><a:srgbClr val="7A5AE0"/></a:accent3><a:accent4><a:srgbClr val="E0A23A"/></a:accent4><a:accent5><a:srgbClr val="E95078"/></a:accent5><a:accent6><a:srgbClr val="3BB6EC"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Dashi"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Aptos Display"/><a:cs typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/></a:minorFont></a:fontScheme><a:fmtScheme name="Dashi"><a:fillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst/><a:bgFillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

function coreProps(pptx) {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(pptx.title || 'Dashi PPT')}</dc:title><dc:subject>${esc(pptx.subject || '')}</dc:subject><dc:creator>${esc(pptx.author || 'Dashi PPT')}</dc:creator><cp:lastModifiedBy>${esc(pptx.author || 'Dashi PPT')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
}

function appProps(slideCount) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>${slideCount}</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant><vt:variant><vt:i4>${slideCount}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${slideCount}" baseType="lpstr">${Array.from({ length: slideCount }, (_, i) => `<vt:lpstr>Slide ${i + 1}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts></Properties>`;
}

function writeZip(root, output) {
  const files = [];
  walk(root, root, files);
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const now = new Date();
  const { dosDate, dosTime } = dosDateTime(now);
  for (const file of files) {
    const name = Buffer.from(file.name.replaceAll(path.sep, '/'), 'utf8');
    const data = file.data;
    const compressed = zlib.deflateRawSync(data, { level: 6 });
    const checksum = crc32(data);
    const local = Buffer.concat([
      u32(0x04034b50), u16(20), u16(0), u16(8), u16(dosTime), u16(dosDate),
      u32(checksum), u32(compressed.length), u32(data.length), u16(name.length), u16(0), name, compressed,
    ]);
    localParts.push(local);
    const central = Buffer.concat([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(8), u16(dosTime), u16(dosDate),
      u32(checksum), u32(compressed.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0),
      u16(0), u32(0), u32(offset), name,
    ]);
    centralParts.push(central);
    offset += local.length;
  }
  const central = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(central.length), u32(offset), u16(0),
  ]);
  fs.writeFileSync(output, Buffer.concat([...localParts, central, end]));
}

function walk(root, current, files) {
  for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
    const full = path.join(current, entry.name);
    if (entry.isDirectory()) walk(root, full, files);
    else files.push({ name: path.relative(root, full), data: fs.readFileSync(full) });
  }
}

function u16(value) {
  const buffer = Buffer.alloc(2);
  buffer.writeUInt16LE(value >>> 0, 0);
  return buffer;
}

function u32(value) {
  const buffer = Buffer.alloc(4);
  buffer.writeUInt32LE(value >>> 0, 0);
  return buffer;
}

function dosDateTime(date) {
  return {
    dosTime: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    dosDate: ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

const CRC_TABLE = Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) value = (value & 1) ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function crc32(data) {
  let value = 0xffffffff;
  for (const byte of data) value = CRC_TABLE[(value ^ byte) & 0xff] ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}
