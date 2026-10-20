/**
 * Card arredondado no Word, pela técnica da Lex Design (cartilha da Anglo
 * American): forma inline de canto arredondado cuja caixa de texto guarda a
 * tabela do conteúdo. Fica no fluxo (acompanha o texto), cresce quando
 * alguém digita (spAutoFit) e a borda, o fundo e a sombra são da forma.
 *
 * O build.ts deixa o conteúdo entre dois parágrafos-marcador
 * (LexRoundOpen-<id> / LexRoundClose-<id>); aqui eles viram a forma.
 */
import type { RoundShape } from './build.js';

const emu = (px: number) => Math.round(px * 9525);

const shapeXml = (s: RoundShape, inner: string, id: number): string => {
  const cx = emu(s.w);
  const cy = emu(s.h);
  // roundRect: raio = adj/100000 × o lado menor (máximo 50000, a pílula).
  const adj = s.adj;
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill.toUpperCase()}"/></a:solidFill>` : '<a:noFill/>';
  const dash = s.line?.style === 'dashed' ? '<a:prstDash val="dash"/>' : s.line?.style === 'dotted' ? '<a:prstDash val="sysDot"/>' : '';
  const line = s.line
    ? `<a:ln w="${emu(s.line.w)}"><a:solidFill><a:srgbClr val="${s.line.color.toUpperCase()}"/></a:solidFill>${dash}</a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  const sh = s.shadow;
  const shadow = sh
    ? `<a:effectLst><a:outerShdw blurRad="${emu(sh.blur)}" dist="${emu(Math.hypot(sh.x, sh.y))}" dir="${Math.round((((Math.atan2(sh.y, sh.x) * 180) / Math.PI + 360) % 360) * 60000)}" algn="tl" rotWithShape="0"><a:srgbClr val="${sh.color.toUpperCase()}">${sh.alpha < 0.999 ? `<a:alpha val="${Math.round(sh.alpha * 100000)}"/>` : ''}</a:srgbClr></a:outerShdw></a:effectLst>`
    : '';
  // Espaço que a sombra ocupa fora da forma (não encosta no que vem depois).
  const out = (v: number) => Math.max(0, emu(v + (sh?.blur ?? 0)));
  const extent = sh ? `<wp:effectExtent l="${out(-sh.x)}" t="${out(-sh.y)}" r="${out(sh.x)}" b="${out(sh.y)}"/>` : '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
  // Caixa de texto termina em parágrafo: o menor possível.
  const tail = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  return (
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>${extent}<wp:docPr id="${id}" name="Card ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:cNvSpPr txBox="1"/>' +
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>${fill}${line}${shadow}</wps:spPr>` +
    `<wps:txbx><w:txbxContent>${inner}${tail}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${emu(s.inset[0])}" tIns="${emu(s.inset[2])}" rIns="${emu(s.inset[1])}" bIns="${emu(s.inset[3])}" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr>` +
    '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
  );
};

/** Parágrafo inteiro que contém o marcador. */
const paragraphAround = (xml: string, mark: string): [number, number] | null => {
  const at = xml.indexOf(`w:val="${mark}"`);
  if (at < 0) return null;
  const start = Math.max(xml.lastIndexOf('<w:p>', at), xml.lastIndexOf('<w:p ', at));
  const end = xml.indexOf('</w:p>', at);
  return start < 0 || end < 0 ? null : [start, end + 6];
};

export const roundShapes = (xml: string, shapes: RoundShape[]): string => {
  let out = xml;
  shapes.forEach((s, i) => {
    const open = paragraphAround(out, `LexRoundOpen-${i}`);
    const close = paragraphAround(out, `LexRoundClose-${i}`);
    if (!open || !close || close[0] < open[1]) return;
    out = out.slice(0, open[0]) + shapeXml(s, out.slice(open[1], close[0]), 7_000_000 + i) + out.slice(close[1]);
  });
  return out;
};
