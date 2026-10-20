/**
 * Layout do documento (office/extract.js) → slides do PowerPoint.
 *
 * O PowerPoint posiciona tudo por coordenada, como o layout do Studio: cada
 * página vira um slide e cada peça vai no mesmo lugar — texto em caixa de
 * texto editável (fontes reais, espaçamento ajustado à largura do Studio),
 * fundo/borda/canto em forma nativa, imagem como imagem (SVG com PNG de
 * reserva) e o que o PowerPoint não desenha (gradiente, máscara, ícone em
 * SVG inline) capturado do próprio Chromium. Fundo de página fica no fundo do
 * slide (não se seleciona sem querer).
 */
import { createHash } from 'node:crypto';

import type { Assets, ImageData } from '../office/assets.js';
import { fitLetterSpacing, type Face, type FontResolver } from '../office/fonts.js';
import { isOneLine, type Border, type Box, type DocumentLayout, type LayoutNode, type Para, type Run, type Shadow, type Tint } from '../office/layout.js';
import { pageNumbering } from '../office/pages.js';

// ---------------------------------------------------------------------------
// Unidades e XML
// ---------------------------------------------------------------------------

/** px CSS → EMU (96 px = 914400 EMU). */
export const emu = (px: number) => Math.round(px * 9525);
/** px CSS → centésimos de ponto (tamanho de fonte, entrelinha, espaçamento). */
const pt100 = (px: number) => Math.round(px * 75);

export const esc = (s: string) =>
  s
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const attrs = (a: Record<string, string | number | undefined | false>) =>
  Object.entries(a)
    .filter(([, v]) => v !== undefined && v !== false)
    .map(([k, v]) => ` ${k}="${typeof v === 'string' ? esc(v) : v}"`)
    .join('');

/** Cor, com transparência quando tem (alfa de 0 a 1). */
const clr = (hex: string, alpha?: number) =>
  alpha !== undefined && alpha < 0.995
    ? `<a:srgbClr val="${hex.toUpperCase()}"><a:alpha val="${Math.round(alpha * 100000)}"/></a:srgbClr>`
    : `<a:srgbClr val="${hex.toUpperCase()}"/>`;
const solid = (hex: string, alpha?: number) => `<a:solidFill>${clr(hex, alpha)}</a:solidFill>`;
const fillOf = (f: string | Tint) => (typeof f === 'string' ? solid(f) : solid(f.color, f.alpha));

/** Tracejado do CSS (o Chromium desenha traço e vão de ~3× a espessura). */
const dashOf = (style?: Border['style']) =>
  style === 'dashed' ? '<a:custDash><a:ds d="300000" sp="300000"/></a:custDash>' : style === 'dotted' ? '<a:prstDash val="sysDot"/>' : '';
const lineXml = (b: Border) => `<a:ln w="${emu(b.w)}"${b.style === 'dotted' ? ' cap="rnd"' : ''}>${b.raw ? solid(b.raw.color, b.raw.alpha) : solid(b.color)}${dashOf(b.style)}</a:ln>`;

/**
 * Sombra suave do CSS → sombra do PowerPoint (só cabe uma): o desfoque da
 * maior, as transparências somadas (o CSS empilha uma curta e uma longa).
 */
const shadowXml = (list: Shadow[] | undefined, b: Box): string => {
  const soft = (list ?? []).filter((s) => s.blur > 0);
  if (!soft.length) return '';
  const main = soft.reduce((a, c) => (c.blur > a.blur ? c : a));
  const alpha = Math.min(0.6, soft.reduce((sum, c) => sum + c.alpha, 0));
  const dist = Math.hypot(main.x, main.y);
  const dir = dist ? Math.round((((Math.atan2(main.y, main.x) * 180) / Math.PI + 360) % 360) * 60000) : 0;
  const sx = main.spread && b.w ? Math.round(((b.w + 2 * main.spread) / b.w) * 100000) : undefined;
  const sy = main.spread && b.h ? Math.round(((b.h + 2 * main.spread) / b.h) * 100000) : undefined;
  return `<a:effectLst><a:outerShdw${attrs({ blurRad: emu(main.blur), dist: emu(dist), dir, sx, sy, algn: 'ctr', rotWithShape: 0 })}>${clr(main.color, alpha)}</a:outerShdw></a:effectLst>`;
};

// ---------------------------------------------------------------------------
// Saída
// ---------------------------------------------------------------------------

export type Media = { name: string; data: Buffer };
export type SlideRel = { id: string; type: 'image' | 'hyperlink' | 'slide'; target: string };
export type BuiltSlide = { xml: string; rels: SlideRel[] };
export type Deck = {
  slides: BuiltSlide[];
  media: Media[];
  used: Set<Face>;
  /** Face mais usada: vira a fonte do tema (caixa de texto nova já sai nela). */
  mainFace: Face | null;
  size: { w: number; h: number };
  /** Número do primeiro slide: o Studio pode não contar a capa ("02/28" na 3ª página). */
  firstSlideNum: number;
};

export type BuildInput = {
  layout: DocumentLayout;
  fonts: FontResolver;
  assets: Assets;
  /** Imagem de fundo de cada página. */
  background: (pageIndex: number) => Promise<ImageData | null>;
};

/** Trecho já cortado onde a linha quebrava no Studio. */
type PRun = Run & { wrap?: boolean; spc?: number };

/**
 * Espaços como o navegador mostra: em sequência viram um, e no começo/fim do
 * parágrafo somem (com white-space normal). Parágrafo só de espaço (o vão
 * entre tags do HTML) não aparece no Studio — nem aqui.
 */
const visibleParas = (paras: Para[], numbered = false): Para[] =>
  paras
    .map((p, pi) => {
      const runs: PRun[] = p.runs.map((r) => (r.t && r.collapse ? { ...r, t: r.t.replace(/[ \t\r\f]+/g, ' ') } : { ...r }));
      // Depois do número da cláusula o espaço fica (é ele que separa "2.1" do texto).
      let atStart = !(numbered && pi === 0);
      for (const [i, r] of runs.entries()) {
        if (r.br) {
          atStart = true;
          continue;
        }
        if (!r.t) continue;
        if (r.wrap) atStart = true;
        if (r.collapse && atStart) r.t = r.t.replace(/^ +/, '');
        const next = runs.slice(i + 1).find((x) => x.br || x.t);
        if (r.collapse && (!next || next.br || next.wrap)) r.t = r.t.replace(/ +$/, '');
        if (r.collapse && next?.t && next.collapse && r.t.endsWith(' ')) next.t = next.t.replace(/^ +/, '');
        if (r.t) atStart = false;
      }
      return { ...p, runs: runs.filter((r) => r.br || r.t) };
    })
    .filter((p) => p.runs.some((r) => r.br || r.t?.trim()));

/** Link na cor do próprio texto (sem isso o PowerPoint pinta todo link de azul). */
const LINK_COLOR =
  '<a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst>';

/** Guid fixo do campo "número do slide" (o PowerPoint usa um por tipo de campo). */
const SLIDENUM_FIELD = '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}';

// ---------------------------------------------------------------------------
// Montagem
// ---------------------------------------------------------------------------

const covers = (b: Box, page: Box) =>
  Math.abs(b.x - page.x) < 2 && Math.abs(b.y - page.y) < 2 && Math.abs(b.w - page.w) < 2 && Math.abs(b.h - page.h) < 2;
const hasText = (n: LayoutNode): boolean => n.k === 'text' || (n.k === 'box' && n.kids.some(hasText));

export const buildDeck = async ({ layout, fonts, assets, background }: BuildInput): Promise<Deck> => {
  /**
   * Fundo do slide a partir de um design de página inteira (frame do
   * Figma): a moldura que cobre a página e as camadas de baixo dela — o que
   * vem antes do primeiro texto (painel, ilustração, arcos). Só a cor da
   * moldura: cor de fundo do slide; com camadas ou pintura: uma imagem de
   * fundo (fica no lugar, não se seleciona sem querer).
   */
  const slideBackdrop = async (page: DocumentLayout['pages'][number], trees: LayoutNode[]) => {
    for (const tree of trees) {
      const frames: Extract<LayoutNode, { k: 'box' | 'shape' }>[] = [];
      let n: LayoutNode = tree;
      let container: Extract<LayoutNode, { k: 'box' }> | null = null;
      while ((n.k === 'box' || n.k === 'shape') && covers(n.box, page.box)) {
        if (n.fill || n.fillRaw || n.paint || n.bgImage) frames.push(n);
        if (n.k !== 'box') break;
        container = n;
        if (n.kids.length === 1 && n.kids[0].k === 'box' && covers(n.kids[0].box, page.box)) n = n.kids[0];
        else break;
      }
      if (!frames.length) continue;
      const layers: LayoutNode[] = [];
      for (const kid of container?.kids ?? []) {
        if (hasText(kid)) break;
        layers.push(kid);
      }
      const top = frames[frames.length - 1];
      // Cor lisa (sem borda, sombra, canto): vira a cor do slide; o resto, imagem.
      const plain = frames.length === 1 && !layers.length && top.fill && !top.paint && !top.bgImage && !top.fillRaw && !top.border && !top.shadow && !top.radius && !page.bgImage;
      if (plain) return { frames, layers, color: top.fill!, image: null };
      const image = await assets.layers(page.index, frames.map((f) => f.id), layers.map((l) => l.id), page.box);
      if (image) return { frames, layers, color: null, image };
    }
    return null;
  };

  const used = new Set<Face>();
  const usage = new Map<Face, number>();
  const media: Media[] = [];
  const mediaByHash = new Map<string, string>();
  // Bloco → slide, pros links internos (menu de seções, sumário).
  const blockPage = new Map(layout.blocks.map((b) => [b.blockId, b.page]));
  const slideOf = (pageIndex: number) => layout.pages.findIndex((p) => p.index === pageIndex);

  const mediaName = (data: Buffer, ext: string) => {
    const hash = createHash('sha1').update(data).digest('hex');
    let name = mediaByHash.get(hash);
    if (!name) {
      name = `image${media.length + 1}.${ext}`;
      media.push({ name, data });
      mediaByHash.set(hash, name);
    }
    return name;
  };

  const faceOf = (r: Pick<Run, 'font' | 'weight' | 'italic'>, chars = 0) => {
    const face = fonts.resolve(r.font, r.weight, r.italic);
    if (face) {
      used.add(face);
      usage.set(face, (usage.get(face) ?? 0) + chars);
    }
    return face;
  };

  // Número do primeiro slide: o Studio pode não contar a capa ("02/28" na 3ª
  // página). O PowerPoint só aceita de 0 a 9999 — fora disso o número vai
  // como texto (não acompanha se mudar a ordem dos slides, mas aparece certo).
  const numbered = layout.pages.findIndex((p) => p.footer?.number);
  const firstOffsets = numbered >= 0 ? pageNumbering(layout, layout.pages[numbered].index) : null;
  const firstNum = 1 + (firstOffsets?.current ?? 0);
  const fieldOk = firstNum >= 0 && firstNum <= 9999;

  const slides: BuiltSlide[] = [];
  for (const [slideIndex, page] of layout.pages.entries()) {
    const origin = { x: page.box.x, y: page.box.y };
    const shapes: string[] = [];
    const rels: SlideRel[] = [];
    let nextId = 2;

    const rel = (type: SlideRel['type'], target: string) => {
      const hit = rels.find((r) => r.type === type && r.target === target);
      if (hit) return hit.id;
      const id = `rId${rels.length + 2}`; // rId1 = layout
      rels.push({ id, type, target });
      return id;
    };
    const imageRel = (img: ImageData) => rel('image', `../media/${mediaName(img.data, img.type === 'jpg' ? 'jpeg' : img.type)}`);

    const xfrm = (b: Box) =>
      `<a:xfrm><a:off x="${emu(b.x - origin.x)}" y="${emu(b.y - origin.y)}"/><a:ext cx="${emu(Math.max(0, b.w))}" cy="${emu(Math.max(0, b.h))}"/></a:xfrm>`;

    /** Link do Studio → clique no PowerPoint: site/e-mail, ou o slide do bloco. */
    const linkXml = (link: string | null | undefined): string => {
      if (!link) return '';
      if (/^(https?:|mailto:|tel:)/i.test(link)) return `<a:hlinkClick r:id="${rel('hyperlink', link)}">${LINK_COLOR}</a:hlinkClick>`;
      const target = blockPage.get(link.replace(/^#/, ''));
      if (target === undefined) return '';
      const slide = slideOf(target);
      if (slide < 0 || slide === slideIndex) return '';
      return `<a:hlinkClick r:id="${rel('slide', `slide${slide + 1}.xml`)}" action="ppaction://hlinksldjump">${LINK_COLOR}</a:hlinkClick>`;
    };

    // --- imagens ----------------------------------------------------------

    const picture = (img: ImageData, b: Box, opts: { name?: string; radius?: number; locked?: boolean } = {}) => {
      let box = b;
      // SVG: o PowerPoint estica até a caixa; o navegador desenha sem distorcer.
      if (img.svg && img.aspect) {
        const boxAspect = b.w / b.h;
        if (Math.abs(boxAspect - img.aspect) > 0.01 * img.aspect) {
          const w = boxAspect > img.aspect ? b.h * img.aspect : b.w;
          const h = boxAspect > img.aspect ? b.h : b.w / img.aspect;
          box = { x: b.x + (b.w - w) / 2, y: b.y + (b.h - h) / 2, w, h };
        }
      }
      const png = imageRel(img);
      const svg = img.svg
        ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${rel('image', `../media/${mediaName(img.svg, 'svg')}`)}"/></a:ext></a:extLst>`
        : '';
      const id = nextId++;
      const locks = opts.locked
        ? '<a:picLocks noGrp="1" noSelect="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noChangeArrowheads="1"/>'
        : '<a:picLocks noChangeAspect="1"/>';
      shapes.push(
        `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(opts.name ?? `Imagem ${id}`)}"/><p:cNvPicPr>${locks}</p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="${png}">${svg}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr>${xfrm(box)}${geometry(box, opts.radius)}</p:spPr></p:pic>`,
      );
    };

    // --- formas -----------------------------------------------------------

    const geometry = (b: Box, radius?: number) => {
      if (!radius || radius < 0.5) return '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
      const adj = Math.min(50000, Math.round((radius / Math.max(1, Math.min(b.w, b.h))) * 100000));
      return `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>`;
    };

    const rect = (b: Box, fill: string | Tint | null, line?: Border, radius?: number, name = 'Forma', effects = '') => {
      const id = nextId++;
      const ln = line ? lineXml(line) : '<a:ln><a:noFill/></a:ln>';
      shapes.push(
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(b)}${geometry(b, radius)}${fill ? fillOf(fill) : '<a:noFill/>'}${ln}${effects}</p:spPr></p:sp>`,
      );
    };

    /** Linha reta (borda tracejada de um lado só). */
    const line = (x1: number, y1: number, x2: number, y2: number, border: Border) => {
      const id = nextId++;
      const b = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
      shapes.push(
        `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Linha ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
          `<p:spPr>${xfrm(b)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lineXml(border)}</p:spPr></p:cxnSp>`,
      );
    };

    /** Fundo, borda e cantos da caixa. Pintura que o PowerPoint não tem (gradiente, imagem de fundo) vai capturada. */
    const decorate = async (n: Extract<LayoutNode, { k: 'box' | 'shape' }>) => {
      if (n.paint) {
        const img = await assets.paint(n.id);
        if (img) return picture(img, n.box, { name: `Fundo ${n.id}` });
      }
      const sides = n.border ?? {};
      const list = (['top', 'right', 'bottom', 'left'] as const).map((s) => sides[s]);
      const fill = n.fillRaw ?? n.fill ?? null;
      if (!fill && !list.some(Boolean)) return;
      // Sombra sem desfoque (anel "0 0 0 6px" em volta do ícone): forma maior atrás, na cor dela.
      for (const s of n.shadow ?? []) {
        if (s.blur > 0 || (!s.spread && !s.x && !s.y)) continue;
        const b = { x: n.box.x + s.x - s.spread, y: n.box.y + s.y - s.spread, w: n.box.w + 2 * s.spread, h: n.box.h + 2 * s.spread };
        rect(b, { color: s.color, alpha: s.alpha }, undefined, n.radius ? n.radius + s.spread : undefined, 'Contorno');
      }
      const effects = fill ? shadowXml(n.shadow, n.box) : '';
      const same = (a?: Border, b?: Border) => a && b && a.w === b.w && a.color === b.color && a.style === b.style && a.raw?.alpha === b.raw?.alpha;
      if (list.every((s) => same(s, list[0]))) {
        // A borda do CSS fica por dentro da caixa; a do PowerPoint, centrada na linha.
        const w = list[0]!.w;
        const b = { x: n.box.x + w / 2, y: n.box.y + w / 2, w: n.box.w - w, h: n.box.h - w };
        return rect(b, fill, list[0], n.radius ? Math.max(0, n.radius - w / 2) : undefined, 'Forma', effects);
      }
      // Borda grossa de um lado só num card arredondado (faixa de destaque à
      // esquerda): em camadas, pra faixa acompanhar a curva como no Studio —
      // a cor da faixa no card inteiro, o card deslocado por cima.
      const names = ['top', 'right', 'bottom', 'left'] as const;
      const thick = names.filter((k) => sides[k] && !sides[k]!.style);
      const accent = thick.length ? thick.reduce((a, k) => (sides[k]!.w > sides[a]!.w ? k : a)) : null;
      const others = names.filter((k) => k !== accent && sides[k]);
      const thin = others.length ? sides[others[0]]! : undefined;
      if (
        accent && n.radius && n.radius >= 4 && n.fill && sides[accent]!.w >= 2 &&
        others.every((k) => same(sides[k], thin) && sides[k]!.w < sides[accent]!.w)
      ) {
        const a = sides[accent]!;
        const inset = (b: Box, d: Partial<Record<(typeof names)[number], number>>): Box => ({
          x: b.x + (d.left ?? 0),
          y: b.y + (d.top ?? 0),
          w: b.w - (d.left ?? 0) - (d.right ?? 0),
          h: b.h - (d.top ?? 0) - (d.bottom ?? 0),
        });
        rect(n.box, a.raw ?? a.color, undefined, n.radius, 'Forma', effects);
        const card = inset(n.box, { [accent]: a.w });
        // Por cima da faixa, só cor opaca (a transparente deixaria a faixa aparecer).
        if (thin && others.length === 3) {
          rect(card, thin.color, undefined, n.radius, 'Borda');
          const w = thin.w;
          const inner = inset(card, Object.fromEntries(names.filter((k) => k !== accent).map((k) => [k, w])));
          return rect(inner, n.fill, undefined, Math.max(0, n.radius - w), 'Forma');
        }
        return rect(card, n.fill, undefined, n.radius, 'Forma');
      }
      if (fill) rect(n.box, fill, undefined, n.radius, 'Forma', effects);
      // Bordas soltas (divisória embaixo do título, linha do tempo tracejada): uma por lado.
      const { x, y, w, h } = n.box;
      const side = (border: Border | undefined, bar: Box, from: [number, number], to: [number, number]) => {
        if (!border) return;
        if (border.style) return line(from[0], from[1], to[0], to[1], border);
        rect(bar, border.raw ?? border.color, undefined, undefined, 'Linha');
      };
      const { top, bottom, left, right } = sides;
      if (top) side(top, { x, y, w, h: top.w }, [x, y + top.w / 2], [x + w, y + top.w / 2]);
      if (bottom) side(bottom, { x, y: y + h - bottom.w, w, h: bottom.w }, [x, y + h - bottom.w / 2], [x + w, y + h - bottom.w / 2]);
      if (left) side(left, { x, y, w: left.w, h }, [x + left.w / 2, y], [x + left.w / 2, y + h]);
      if (right) side(right, { x: x + w - right.w, y, w: right.w, h }, [x + w - right.w / 2, y], [x + w - right.w / 2, y + h]);
    };

    // --- texto ------------------------------------------------------------

    const runXml = (r: Run, text: string, face: Face | null, spacing: number | undefined, tag: 'r' | 'fld' = 'r') => {
      const bold = face ? face.bold : r.weight >= 600;
      const italic = face ? face.italic : r.italic;
      const typeface = esc(face?.family ?? r.font);
      const props =
        `<a:rPr${attrs({
          lang: 'pt-BR',
          sz: Math.max(100, pt100(r.size)),
          b: bold ? 1 : undefined,
          i: italic ? 1 : undefined,
          u: r.underline ? 'sng' : undefined,
          strike: r.strike ? 'sngStrike' : undefined,
          cap: r.caps ? 'all' : undefined,
          spc: spacing ? Math.round(spacing * 75) : undefined,
          baseline: r.verticalAlign === 'super' ? 30000 : r.verticalAlign === 'sub' ? -25000 : undefined,
          dirty: 0,
        })}>` +
        solid(r.color, r.alpha) +
        (r.highlight ? `<a:highlight><a:srgbClr val="${r.highlight.toUpperCase()}"/></a:highlight>` : '') +
        `<a:latin typeface="${typeface}"/><a:ea typeface="${typeface}"/><a:cs typeface="${typeface}"/>` +
        linkXml(r.link) +
        '</a:rPr>';
      const t = `<a:t>${esc(r.lower ? text.toLowerCase() : text)}</a:t>`;
      return tag === 'fld' ? `<a:fld id="${SLIDENUM_FIELD}" type="slidenum">${props}${t}</a:fld>` : `<a:r>${props}${t}</a:r>`;
    };

    /** Trecho de texto com o espaçamento que dá a largura do Studio. */
    const textRun = (r: PRun, justified = false) => {
      const face = faceOf(r, r.t?.length ?? 0);
      return runXml(r, r.t ?? '', face, r.spc ?? fitLetterSpacing(r, face, fonts.data, r.size, justified));
    };

    const BULLET: Record<string, string> = {
      decimal: 'arabicPeriod', 'lower-alpha': 'alphaLcPeriod', 'lower-latin': 'alphaLcPeriod',
      'upper-alpha': 'alphaUcPeriod', 'upper-latin': 'alphaUcPeriod', 'lower-roman': 'romanLcPeriod', 'upper-roman': 'romanUcPeriod',
    };

    const paragraph = (para: Para, left: number, prefix = '', before = 0) => {
      const algn = { left: 'l', center: 'ctr', right: 'r', both: 'just' }[para.align];
      let list = '<a:buNone/>';
      let indent: Record<string, number> = {};
      if (para.list) {
        // Item de lista: o texto começa onde começava no Studio; o marcador fica no vão antes dele.
        const gap = (para.indent ?? 0) >= 4 ? para.indent! : para.fontSize * 1.1;
        const marL = para.x !== undefined ? para.x - left : gap;
        indent = { marL: emu(Math.max(0, marL)), indent: -emu(Math.min(gap, Math.max(0, marL))), lvl: Math.min(8, para.list.level) };
        // Marcador na cor, no tamanho e na fonte do Studio (o "•" do Chromium vem da fonte do texto).
        const first = para.runs.find((r) => r.t);
        const mk = para.marker;
        const look =
          (mk ? `<a:buClr>${clr(mk.color, mk.alpha)}</a:buClr>` : '') +
          (mk?.size && first ? `<a:buSzPct val="${Math.round(Math.min(4, Math.max(0.25, mk.size / first.size)) * 100000)}"/>` : '');
        const font = first ? esc(faceOf(first)?.family ?? first.font) : 'Arial';
        list =
          para.list.kind === 'bullet'
            ? `${look}<a:buFont typeface="${font}"/><a:buChar char="•"/>`
            : `${look}<a:buFont typeface="${font}"/><a:buAutoNum type="${BULLET[para.list.kind] ?? 'arabicPeriod'}"/>`;
      }
      const lnSpc = para.lineHeight
        ? `<a:lnSpc><a:spcPts val="${pt100(para.lineHeight)}"/></a:lnSpc>`
        : '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>';
      const runs = (para.runs as PRun[])
        .map((r) => (r.br ? '<a:br><a:rPr lang="pt-BR" dirty="0"/></a:br>' : r.t ? (r.wrap ? `<a:br>${breakProps(r)}</a:br>` : '') + textRun(r, para.align === 'both') : ''))
        .join('');
      const first = para.runs.find((r) => r.t);
      const end = first
        ? `<a:endParaRPr lang="pt-BR" sz="${Math.max(100, pt100(first.size))}" dirty="0">${solid(first.color)}<a:latin typeface="${esc(faceOf(first)?.family ?? first.font)}"/></a:endParaRPr>`
        : `<a:endParaRPr lang="pt-BR" sz="${Math.max(100, pt100(para.fontSize))}" dirty="0"/>`;
      return `<a:p><a:pPr${attrs({ ...indent, algn })}>${lnSpc}<a:spcBef><a:spcPts val="${pt100(before)}"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>${list}</a:pPr>${prefix}${runs}${end}</a:p>`;
    };

    /** Quebra no lugar onde o Studio quebrou: com o tamanho do texto (senão a linha vazia teria 18pt). */
    const breakProps = (r: Run) => `<a:rPr lang="pt-BR" sz="${Math.max(100, pt100(r.size))}" dirty="0"/>`;

    /**
     * Corta os trechos onde a linha quebrava no Studio: o PowerPoint mede um
     * tiquinho diferente e quebraria em outro lugar ("Onde nasce o aroma"
     * numa linha só). O espaçamento é o do trecho inteiro. Justificado fica
     * de fora (a linha antes da quebra não esticaria).
     */
    const atStudioBreaks = (paras: Para[]): Para[] =>
      paras.map((p) =>
        p.align === 'both'
          ? p
          : {
              ...p,
              runs: p.runs.flatMap((r): PRun[] => {
                if (!r.t || !r.breaks?.length) return [r];
                const spc = fitLetterSpacing(r, fonts.resolve(r.font, r.weight, r.italic), fonts.data, r.size);
                const cuts = [...new Set([0, ...r.breaks.filter((i) => i > 0 && i < r.t!.length), r.t.length])];
                return cuts.slice(0, -1).map((a, k) => ({ ...r, t: r.t!.slice(a, cuts[k + 1]), spc, breaks: undefined, wrap: k > 0 || r.breaks![0] === 0 }));
              }),
            },
      );

    const textBox = (b: Box, body: string, opts: { wrap: boolean; middle?: boolean; link?: string | null; name?: string }) => {
      const id = nextId++;
      shapes.push(
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(opts.name ?? `Texto ${id}`)}">${linkXml(opts.link)}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(b)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
          `<p:txBody><a:bodyPr${attrs({ wrap: opts.wrap ? 'square' : 'none', lIns: 0, tIns: 0, rIns: 0, bIns: 0, anchor: opts.middle ? 'ctr' : 't', rtlCol: 0 })}><a:noAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>`,
      );
    };

    const text = (node: Extract<LayoutNode, { k: 'text' }>) => {
      const n = { ...node, paras: visibleParas(atStudioBreaks(node.paras), Boolean(node.autonumber && node.autonumberStyle)) };
      const wrapped = n.paras.some((p) => (p.runs as PRun[]).some((r) => r.wrap));
      const hasRuns = n.paras.length > 0;
      const number = n.autonumber ? n.autonumberText ?? n.autonumber : null;
      const style = n.autonumberStyle;
      // Só o número (selo da cláusula): no lugar exato dele, numa linha.
      if (!hasRuns) {
        if (!number || !style) return;
        const b = n.autonumberBox ?? n.box;
        const r: Run = { ...style, t: number, link: null };
        const lh = style.lineHeight ?? style.size * 1.2;
        const para: Para = { align: 'left', lineHeight: lh, fontSize: style.size, runs: [r] };
        return textBox({ ...b, w: b.w + 4 }, paragraph(para, b.x), { wrap: false, name: `Número ${number}` });
      }
      // Número no começo do texto: vai como texto, com o vão até as palavras.
      let prefix = '';
      if (number && style) {
        prefix = textRun({ ...style, t: number, link: null, width: n.autonumberBox?.w, lines: 1 });
        // Margem do CSS entre o número e o texto: um espaço do tamanho dela.
        const gap = n.autonumberGap ?? 0;
        if (gap > 1) {
          const r: Run = { ...style, t: ' ', link: null };
          const face = faceOf(r);
          const space = (face && fonts.data.metrics.get(face)?.[32]) ?? 0.25;
          prefix += runXml(r, ' ', face, gap - space * style.size);
        }
      }
      const oneLine = isOneLine(n);
      const align = n.paras[0]?.align ?? 'left';
      let b = n.box;
      if (oneLine) {
        // Uma linha: nunca quebra (o PowerPoint mede um tiquinho diferente); a caixa ancora o alinhamento.
        if (n.inkRight != null && align === 'left') b = { ...b, w: Math.max(b.w, n.inkRight - b.x) };
      } else {
        // Várias linhas: folga mínima pra a quebra cair no mesmo lugar; com a
        // quebra do Studio marcada, folga maior (a linha nunca quebra de novo).
        const slack = wrapped ? Math.max(2, b.w * 0.04) : 2;
        b = align === 'center' ? { ...b, x: b.x - slack / 2, w: b.w + slack } : align === 'right' ? { ...b, x: b.x - slack, w: b.w + slack } : { ...b, w: b.w + slack };
      }
      const links = new Set(n.paras.flatMap((p) => p.runs.filter((r) => r.t?.trim()).map((r) => r.link ?? '')));
      const link = links.size === 1 ? [...links][0] || null : null;
      // Espaço entre parágrafos (margem do <p>, do <ul>): o que sobra entre as linhas além da entrelinha.
      const lh = (p: Para) => p.lineHeight ?? p.fontSize * 1.2;
      const before = (p: Para, i: number) => {
        if (p.first === undefined) return 0;
        const prev = n.paras[i - 1];
        const gap = i === 0 ? (n.middle ? 0 : p.first - n.box.y - lh(p) / 2) : prev?.last !== undefined ? p.first - prev.last - (lh(prev) + lh(p)) / 2 : 0;
        return gap > 0.5 ? gap : 0;
      };
      // Entrelinha menor que a altura da fonte ("80+" com line-height 1): o
      // PowerPoint sobe o texto — medido: ~metade da diferença, menos 1px.
      const p0 = n.paras[0];
      const lift = !n.middle && p0?.area ? 0.49 * (p0.area - lh(p0)) - 1 : 0;
      if (lift > 0.5) b = { ...b, y: b.y + lift };
      const body = n.paras.map((p, i) => paragraph(p, b.x, i === 0 ? prefix : '', before(p, i))).join('');
      textBox(b, body, { wrap: !oneLine, middle: n.middle, link });
    };

    // --- árvore -----------------------------------------------------------

    // Peças que foram pro fundo do slide: a moldura (só a pintura) e as camadas de baixo (inteiras).
    const inBackground = new Set<LayoutNode>();
    const frameOnly = new Set<LayoutNode>();

    const walk = async (n: LayoutNode, parent?: Extract<LayoutNode, { k: 'box' }>): Promise<void> => {
      if (inBackground.has(n)) return;
      if (n.k === 'text') return text(n);
      // Passa da borda de um card que recorta: só o pedaço visível, capturado.
      if ((n.k === 'img' || n.k === 'raster') && n.clip) {
        if (n.clip.w < 0.5 || n.clip.h < 0.5) return; // toda fora da área que recorta
        const img = await assets.raster(n.id, n.clip);
        return img ? picture(img, n.clip) : undefined;
      }
      if (n.k === 'img') {
        const img = await assets.image(n.src, n.box.w, n.box.h);
        if (!img) return;
        // Foto que preenche um card de canto arredondado: os mesmos cantos.
        const fills = parent?.radius && Math.abs(n.box.w - parent.box.w) < 2 && Math.abs(n.box.h - parent.box.h) < 2;
        return picture(img, n.box, { radius: n.radius ?? (fills ? parent!.radius : undefined) });
      }
      if (n.k === 'raster') {
        const img = await assets.raster(n.id, n.box);
        return img ? picture(img, n.box) : undefined;
      }
      if (!frameOnly.has(n)) await decorate(n);
      if (n.k === 'box') for (const kid of n.kids) await walk(kid, n);
    };

    // --- página -----------------------------------------------------------

    const backdrop = await slideBackdrop(page, layout.blocks.filter((b) => b.page === page.index && b.tree).map((b) => b.tree!));
    let bgXml = '';
    if (backdrop) {
      for (const n of backdrop.frames) frameOnly.add(n);
      for (const n of backdrop.layers) inBackground.add(n);
    }
    const bg = backdrop ? null : await background(page.index);
    if (backdrop?.color) {
      bgXml = `<p:bg><p:bgPr>${solid(backdrop.color)}<a:effectLst/></p:bgPr></p:bg>`;
    } else if (backdrop?.image) {
      bgXml = `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${imageRel(backdrop.image)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
    } else if (bg && page.bgColor && page.bgColor !== 'ffffff') {
      // Cor + imagem (que pode ser transparente): cor no fundo, imagem travada por cima.
      bgXml = `<p:bg><p:bgPr>${solid(page.bgColor)}<a:effectLst/></p:bgPr></p:bg>`;
      picture(bg, page.box, { name: 'Fundo da página', locked: true });
    } else if (bg) {
      bgXml = `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${imageRel(bg)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
    } else if (page.bgColor) {
      bgXml = `<p:bg><p:bgPr>${solid(page.bgColor)}<a:effectLst/></p:bgPr></p:bg>`;
    }

    if (page.nav) await walk(page.nav.tree);
    for (const block of layout.blocks.filter((b) => b.page === page.index && b.tree)) await walk(block.tree!);

    const footer = page.footer;
    if (footer?.text) text(footer.text);
    if (footer?.number) {
      const num = footer.number;
      const offsets = pageNumbering(layout, page.index);
      const current = slideIndex + 1 + (offsets?.current ?? 0);
      const total = layout.pages.length + (offsets?.total ?? 0);
      const format = num.format ?? 'page-total';
      const padded = format === 'page-total';
      const r: Run = { ...num.style, t: '', link: null };
      const face = faceOf(r);
      const plain = (t: string) => runXml(r, t, face, undefined);
      // Página: campo do PowerPoint (acompanha se mudar a ordem dos slides). O total é texto.
      const field = (padded && current >= 0 && current < 10 ? plain('0') : '') + (fieldOk ? runXml(r, String(current), face, undefined, 'fld') : plain(String(current)));
      const runs =
        format === 'page' ? field
        : format === 'page-of-total' ? plain('Página ') + field + plain(` de ${total}`)
        : field + plain(`/${String(total).padStart(2, '0')}`);
      const algn = num.align === 'right' || num.align === 'end' ? 'r' : num.align === 'center' ? 'ctr' : 'l';
      // Folga pros lados (o número pode crescer: "9" → "10").
      const extra = 24;
      const b = algn === 'r' ? { ...num.box, x: num.box.x - extra, w: num.box.w + extra } : algn === 'ctr' ? { ...num.box, x: num.box.x - extra / 2, w: num.box.w + extra } : { ...num.box, w: num.box.w + extra };
      const body = `<a:p><a:pPr algn="${algn}"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr>${runs}</a:p>`;
      textBox(b, body, { wrap: false, middle: true, name: 'Número da página' });
    }

    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      `<p:cSld>${bgXml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      `${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
    slides.push({ xml, rels });
  }

  const mainFace = [...usage.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  return {
    slides,
    media,
    used,
    mainFace,
    size: { w: layout.meta.pageWidthPx, h: layout.meta.pageHeightPx },
    firstSlideNum: fieldOk ? firstNum : 1,
  };
};
