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
import { isOneLine, type Box, type DocumentLayout, type LayoutNode, type Para, type Run } from '../office/layout.js';
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

const solid = (hex: string) => `<a:solidFill><a:srgbClr val="${hex.toUpperCase()}"/></a:solidFill>`;

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

/**
 * Espaços como o navegador mostra: em sequência viram um, e no começo/fim do
 * parágrafo somem (com white-space normal). Parágrafo só de espaço (o vão
 * entre tags do HTML) não aparece no Studio — nem aqui.
 */
const visibleParas = (paras: Para[], numbered = false): Para[] =>
  paras
    .map((p, pi) => {
      const runs = p.runs.map((r) => (r.t && r.collapse ? { ...r, t: r.t.replace(/[ \t\r\f]+/g, ' ') } : { ...r }));
      // Depois do número da cláusula o espaço fica (é ele que separa "2.1" do texto).
      let atStart = !(numbered && pi === 0);
      for (const [i, r] of runs.entries()) {
        if (r.br) {
          atStart = true;
          continue;
        }
        if (!r.t) continue;
        if (r.collapse && atStart) r.t = r.t.replace(/^ +/, '');
        const next = runs.slice(i + 1).find((x) => x.br || x.t);
        if (r.collapse && (!next || next.br)) r.t = r.t.replace(/ +$/, '');
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

export const buildDeck = async ({ layout, fonts, assets, background }: BuildInput): Promise<Deck> => {
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

    const rect = (b: Box, fill: string | null, line?: { w: number; color: string }, radius?: number, name = 'Forma') => {
      const id = nextId++;
      const ln = line ? `<a:ln w="${emu(line.w)}">${solid(line.color)}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
      shapes.push(
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(b)}${geometry(b, radius)}${fill ? solid(fill) : '<a:noFill/>'}${ln}</p:spPr></p:sp>`,
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
      const uniform = list.every((s) => s && s.w === list[0]!.w && s.color === list[0]!.color);
      if (!n.fill && !list.some(Boolean)) return;
      if (uniform) {
        // A borda do CSS fica por dentro da caixa; a do PowerPoint, centrada na linha.
        const w = list[0]!.w;
        const b = { x: n.box.x + w / 2, y: n.box.y + w / 2, w: n.box.w - w, h: n.box.h - w };
        return rect(b, n.fill ?? null, list[0], n.radius ? Math.max(0, n.radius - w / 2) : undefined);
      }
      if (n.fill) rect(n.box, n.fill, undefined, n.radius);
      // Bordas soltas (divisória embaixo do título): uma barra por lado.
      const { x, y, w, h } = n.box;
      if (sides.top) rect({ x, y, w, h: sides.top.w }, sides.top.color, undefined, undefined, 'Linha');
      if (sides.bottom) rect({ x, y: y + h - sides.bottom.w, w, h: sides.bottom.w }, sides.bottom.color, undefined, undefined, 'Linha');
      if (sides.left) rect({ x, y, w: sides.left.w, h }, sides.left.color, undefined, undefined, 'Linha');
      if (sides.right) rect({ x: x + w - sides.right.w, y, w: sides.right.w, h }, sides.right.color, undefined, undefined, 'Linha');
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
        solid(r.color) +
        (r.highlight ? `<a:highlight><a:srgbClr val="${r.highlight.toUpperCase()}"/></a:highlight>` : '') +
        `<a:latin typeface="${typeface}"/><a:ea typeface="${typeface}"/><a:cs typeface="${typeface}"/>` +
        linkXml(r.link) +
        '</a:rPr>';
      const t = `<a:t>${esc(r.lower ? text.toLowerCase() : text)}</a:t>`;
      return tag === 'fld' ? `<a:fld id="${SLIDENUM_FIELD}" type="slidenum">${props}${t}</a:fld>` : `<a:r>${props}${t}</a:r>`;
    };

    /** Trecho de texto com o espaçamento que dá a largura do Studio. */
    const textRun = (r: Run, justified = false) => {
      const face = faceOf(r, r.t?.length ?? 0);
      return runXml(r, r.t ?? '', face, fitLetterSpacing(r, face, fonts.data, r.size, justified));
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
        list = para.list.kind === 'bullet' ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : `<a:buAutoNum type="${BULLET[para.list.kind] ?? 'arabicPeriod'}"/>`;
      }
      const lnSpc = para.lineHeight
        ? `<a:lnSpc><a:spcPts val="${pt100(para.lineHeight)}"/></a:lnSpc>`
        : '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>';
      const runs = para.runs
        .map((r) => (r.br ? '<a:br><a:rPr lang="pt-BR" dirty="0"/></a:br>' : r.t ? textRun(r, para.align === 'both') : ''))
        .join('');
      const first = para.runs.find((r) => r.t);
      const end = first
        ? `<a:endParaRPr lang="pt-BR" sz="${Math.max(100, pt100(first.size))}" dirty="0">${solid(first.color)}<a:latin typeface="${esc(faceOf(first)?.family ?? first.font)}"/></a:endParaRPr>`
        : `<a:endParaRPr lang="pt-BR" sz="${Math.max(100, pt100(para.fontSize))}" dirty="0"/>`;
      return `<a:p><a:pPr${attrs({ ...indent, algn })}>${lnSpc}<a:spcBef><a:spcPts val="${pt100(before)}"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>${list}</a:pPr>${prefix}${runs}${end}</a:p>`;
    };

    const textBox = (b: Box, body: string, opts: { wrap: boolean; middle?: boolean; link?: string | null; name?: string }) => {
      const id = nextId++;
      shapes.push(
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(opts.name ?? `Texto ${id}`)}">${linkXml(opts.link)}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(b)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
          `<p:txBody><a:bodyPr${attrs({ wrap: opts.wrap ? 'square' : 'none', lIns: 0, tIns: 0, rIns: 0, bIns: 0, anchor: opts.middle ? 'ctr' : 't', rtlCol: 0 })}><a:noAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>`,
      );
    };

    const text = (node: Extract<LayoutNode, { k: 'text' }>) => {
      const n = { ...node, paras: visibleParas(node.paras, Boolean(node.autonumber && node.autonumberStyle)) };
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
        // Várias linhas: folga mínima pra a quebra cair no mesmo lugar.
        const slack = 2;
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
      const body = n.paras.map((p, i) => paragraph(p, b.x, i === 0 ? prefix : '', before(p, i))).join('');
      textBox(b, body, { wrap: !oneLine, middle: n.middle, link });
    };

    // --- árvore -----------------------------------------------------------

    const walk = async (n: LayoutNode, parent?: Extract<LayoutNode, { k: 'box' }>): Promise<void> => {
      if (n.k === 'text') return text(n);
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
      await decorate(n);
      if (n.k === 'box') for (const kid of n.kids) await walk(kid, n);
    };

    // --- página -----------------------------------------------------------

    const bg = await background(page.index);
    let bgXml = '';
    if (bg && page.bgColor && page.bgColor !== 'ffffff') {
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
      const field = (padded && current >= 0 && current < 10 ? plain('0') : '') + runXml(r, String(current), face, undefined, 'fld');
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

  const first = layout.pages.findIndex((p) => p.footer?.number);
  const offsets = first >= 0 ? pageNumbering(layout, layout.pages[first].index) : null;
  const mainFace = [...usage.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  return {
    slides,
    media,
    used,
    mainFace,
    size: { w: layout.meta.pageWidthPx, h: layout.meta.pageHeightPx },
    firstSlideNum: 1 + (offsets?.current ?? 0),
  };
};
