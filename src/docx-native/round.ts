/**
 * Card arredondado no Word, pela técnica da Lex Design (cartilha da Anglo
 * American): forma inline de canto arredondado cuja caixa de texto guarda a
 * tabela do conteúdo. Fica no fluxo (acompanha o texto), cresce quando
 * alguém digita (spAutoFit) e a borda, o fundo e a sombra são da forma.
 *
 * O build.ts deixa o conteúdo entre dois parágrafos-marcador
 * (LexRoundOpen-<id> / LexRoundClose-<id>); aqui eles viram a forma.
 *
 * Plano B, pro que a forma não comporta (card com outros cards ou arte
 * dentro, numeração automática, célula mesclada, peça pequena): o conteúdo
 * fica na tabela e a forma vai solta atrás do texto, no canto da célula
 * (run marcado LexRoundBack-<id> no primeiro parágrafo). Não cresce com o
 * texto, mas não mexe em nada do fluxo.
 */
import type { RoundShape } from './build.js';

const emu = (px: number) => Math.round(px * 9525);

/** Geometria, fundo, contorno e sombra da forma. */
const spPr = (s: RoundShape): string => {
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill.toUpperCase()}"/></a:solidFill>` : '<a:noFill/>';
  const dash = s.line?.style === 'dashed' ? '<a:prstDash val="dash"/>' : s.line?.style === 'dotted' ? '<a:prstDash val="sysDot"/>' : '';
  const line = s.line
    ? `<a:ln w="${emu(s.line.w)}"><a:solidFill><a:srgbClr val="${s.line.color.toUpperCase()}"/></a:solidFill>${dash}</a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  const sh = s.shadow;
  const shadow = sh
    ? `<a:effectLst><a:outerShdw blurRad="${emu(sh.blur)}" dist="${emu(Math.hypot(sh.x, sh.y))}" dir="${Math.round((((Math.atan2(sh.y, sh.x) * 180) / Math.PI + 360) % 360) * 60000)}" algn="tl" rotWithShape="0"><a:srgbClr val="${sh.color.toUpperCase()}">${sh.alpha < 0.999 ? `<a:alpha val="${Math.round(sh.alpha * 100000)}"/>` : ''}</a:srgbClr></a:outerShdw></a:effectLst>`
    : '';
  // roundRect: raio = adj/100000 × o lado menor (máximo 50000, a pílula).
  return (
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(s.w)}" cy="${emu(s.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${s.adj}"/></a:avLst></a:prstGeom>${fill}${line}${shadow}</wps:spPr>`
  );
};

/** Espaço que a sombra ocupa fora da forma (não encosta no que vem depois). */
const effectExtent = (s: RoundShape): string => {
  const sh = s.shadow;
  const out = (v: number) => Math.max(0, emu(v + (sh?.blur ?? 0)));
  return sh ? `<wp:effectExtent l="${out(-sh.x)}" t="${out(-sh.y)}" r="${out(sh.x)}" b="${out(sh.y)}"/>` : '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
};

const graphic = (body: string) =>
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  `<wps:wsp>${body}</wps:wsp></a:graphicData></a:graphic>`;

const header = (s: RoundShape) =>
  `<wp:extent cx="${emu(s.w)}" cy="${emu(s.h)}"/>${effectExtent(s)}`;

const inlineXml = (s: RoundShape, inner: string, id: number): string => {
  const [l, r, t, b] = s.inset ?? [0, 0, 0, 0];
  // Caixa de texto termina em parágrafo: o menor possível.
  const tail = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  return (
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">${header(s)}<wp:docPr id="${id}" name="Card ${id}"/><wp:cNvGraphicFramePr/>` +
    graphic(
      '<wps:cNvSpPr txBox="1"/>' +
        spPr(s) +
        `<wps:txbx><w:txbxContent>${inner}${tail}</w:txbxContent></wps:txbx>` +
        `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${emu(l)}" tIns="${emu(t)}" rIns="${emu(r)}" bIns="${emu(b)}" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr>`,
    ) +
    '</wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
  );
};

const backXml = (s: RoundShape, id: number): string =>
  '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(s.x ?? 0)}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(s.y ?? 0)}</wp:posOffset></wp:positionV>` +
  `${header(s)}<wp:wrapNone/><wp:docPr id="${id}" name="Card ${id}"/><wp:cNvGraphicFramePr/>` +
  graphic(`<wps:cNvSpPr/>${spPr(s)}<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`) +
  '</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>';

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
    const id = 7_000_000 + i;
    if (s.back) {
      out = out.replace(new RegExp(`<w:r><w:rPr><w:rStyle w:val="LexRoundBack-${i}"/>(?:(?!</w:r>)[\\s\\S])*</w:r>`), () => backXml(s, id));
      return;
    }
    const open = paragraphAround(out, `LexRoundOpen-${i}`);
    const close = paragraphAround(out, `LexRoundClose-${i}`);
    if (!open || !close || close[0] < open[1]) return;
    out = out.slice(0, open[0]) + inlineXml(s, out.slice(open[1], close[0]), id) + out.slice(close[1]);
  });
  return out;
};
