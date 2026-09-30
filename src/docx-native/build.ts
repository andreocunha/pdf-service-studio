/**
 * Layout do documento (extract.js) → Word no padrão da Lex
 * (LEXDESIGN-ReceitinhasDaLex): tudo que é visual é tabela sem borda, nada
 * flutua no corpo, numeração é automática, títulos têm estilo, fundo e capa
 * moram no cabeçalho (não editáveis), seção nova só quando o fundo muda
 * (quebra de página + seção contínua).
 *
 * Como a estrutura sai de uma árvore com posições reais (inclusive modelos
 * com posição absoluta), a montagem é por fatiamento geométrico: faixas
 * horizontais vazias separam linhas, faixas verticais separam colunas
 * (células de tabela). Caixa com fundo/borda/padding vira célula com
 * sombreamento/borda/margem.
 */
import {
  AlignmentType,
  BorderStyle,
  Bookmark,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  HeightRule,
  HorizontalPositionRelativeFrom,
  ImageRun,
  InternalHyperlink,
  LevelFormat,
  LevelSuffix,
  LineRuleType,
  PageBreak,
  Paragraph,
  SectionType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  VerticalMergeType,
  VerticalPositionRelativeFrom,
  WidthType,
  type IBorderOptions,
  type ISectionOptions,
  type ParagraphChild,
} from 'docx';

import type { Face } from '../docx-fonts.js';
import type { FontResolver } from './fonts.js';
import type { Box, DocumentLayout, LayoutNode, Para, Run } from './layout.js';

// ---------------------------------------------------------------------------
// Unidades
// ---------------------------------------------------------------------------

/** 1 px CSS = 15 twips (1/20 pt; 96 px = 72 pt). */
const tw = (px: number) => Math.max(0, Math.round(px * 15));
/** Tamanho de fonte: px CSS → meio-ponto. */
const halfPt = (px: number) => Math.max(2, Math.round(px * 1.5));
/** px CSS → EMU (DrawingML). */
const emu = (px: number) => Math.round(px * 9525);

// ---------------------------------------------------------------------------
// Representação intermediária (mutável; vira docx no fim)
// ---------------------------------------------------------------------------

type ImageData = { data: Buffer; type: 'png' | 'jpg' | 'gif' | 'bmp'; svg?: Buffer; aspect?: number };
type RunIR =
  | { kind: 'text'; text: string; run: Run; face: Face | null; spacing?: number }
  | { kind: 'br' }
  | { kind: 'img'; image: ImageData; w: number; h: number };
type PIR = {
  kind: 'p';
  runs: RunIR[];
  align?: 'left' | 'center' | 'right' | 'both';
  indentLeft?: number;
  indentRight?: number;
  /** Deslocamento do marcador de lista (o texto quebra alinhado, o marcador fica no vão). */
  hanging?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  line?: { value: number; rule: 'auto' | 'exact' | 'atLeast' };
  numbering?: { ref: string; level: number; instance?: number };
  bookmark?: string;
  borderBottom?: { w: number; color: string };
  borderTop?: { w: number; color: string };
  pageBreakBefore?: boolean;
  /** Formatação da marca de parágrafo — é dela que o número automático herda a fonte. */
  mark?: { run: Run; face: Face | null };
  heading?: 1 | 2 | 3;
  /** Mantém com o próximo (título de capítulo não fica sozinho no pé da página). */
  keepNext?: boolean;
  spacer?: number;
  /** Parágrafo que só segura imagem: fonte mínima na marca (sem folga de descendente). */
  tiny?: boolean;
  /** Sombreamento do parágrafo (fundo do selo do número). */
  shade?: string;
  /** Parágrafo-vão que cresce: tamanho da fonte da marca (meio-ponto). */
  grow?: number;
  /** Linha em branco no estilo Normal (fonte e entrelinha do corpo de texto): o espaço entre tópicos. */
  blank?: boolean;
};
type CellIR = {
  width: number;
  /** Espaço vazio entre colunas (recebe enfeites tirados do corte). */
  gap?: boolean;
  /** Arte atrás do conteúdo (banner com gradiente/ilustração): imagem ancorada na célula, atrás do texto. */
  backdrop?: { image: ImageData; w: number; h: number; x?: number; y?: number; front?: boolean; center?: boolean };
  fill?: string;
  borders?: Partial<Record<'top' | 'right' | 'bottom' | 'left', { w: number; color: string }>>;
  margins?: [number, number, number, number];
  vAlign?: 'top' | 'center' | 'bottom';
  /** Mesclagem vertical: começa aqui / continua a de cima. */
  vmerge?: 'restart' | 'continue';
  children: OutIR[];
};
type TIR = { kind: 't'; width: number; indent: number; rows: { height?: number; cells: CellIR[] }[] };
type OutIR = PIR | TIR;

// ---------------------------------------------------------------------------
// Contexto
// ---------------------------------------------------------------------------

export type Assets = {
  image: (src: string, w: number, h: number) => Promise<ImageData | null>;
  raster: (id: number, box: Box) => Promise<ImageData | null>;
  /** Só a pintura da caixa (fundo, gradiente, borda, sombra, cantos), sem o conteúdo. */
  paint: (id: number) => Promise<ImageData | null>;
  /** Tudo do elemento menos o texto (bloco complexo). */
  backdrop: (id: number, clip: Box) => Promise<ImageData | null>;
};

type Ctx = {
  /** Quantos caracteres saíram em cada face/tamanho — a mais usada vira o estilo Normal. */
  usage: Map<string, number>;
  fonts: FontResolver;
  assets: Assets;
  used: Set<Face>;
  pending: Promise<void>[];
  /** Formatos de lista manual ("i)", "(a)", "1.") → numeração do Word. */
  lists: Map<string, { format: ListFormat; before: string; after: string; suffix: 'nothing' | 'space' }>;
  /** Sequência em andamento: o próximo "ii)" continua a lista do "i)". */
  seq: { key: string; value: number; instance: number } | null;
  instances: number;
  /** Onde a numeração de cláusulas do Word está (1.2.3 → [1, 2, 3]). */
  clauses: number[];
};

type ListFormat = 'lowerRoman' | 'upperRoman' | 'lowerLetter' | 'upperLetter' | 'decimal';

// ---------------------------------------------------------------------------
// Geometria
// ---------------------------------------------------------------------------

const TOL = 0.75;
const bottom = (b: Box) => b.y + b.h;
const right = (b: Box) => b.x + b.w;
/**
 * Só fundo ou borda viram célula. Padding sozinho não: a posição dos filhos
 * já vem medida, e uma tabela a mais só somaria margem.
 */
const decorated = (n: LayoutNode) =>
  (n.k === 'box' || n.k === 'text') && Boolean((n as { fill?: string }).fill || (n as { border?: object }).border || (n as { paint?: boolean }).paint);

type Decorated = { id: number; box: Box; fill?: string; border?: CellIR['borders']; paint?: boolean };

/**
 * Como a caixa se pinta no Word: sempre no preenchimento/borda da própria
 * célula (padrão da equipe: a estrutura acompanha o texto quando alguém
 * digita; forma atrás da tabela desmonta). Canto arredondado é ignorado —
 * tabela do Word não arredonda. Só o que célula não pinta (gradiente, sombra,
 * imagem) vai como a pintura capturada, atrás do texto, colada no canto da
 * célula (margem esquerda descontada).
 */
const paintOf = async (n: Decorated, margins: number[], ctx: Ctx): Promise<Partial<CellIR>> => {
  if (n.paint) {
    const image = await ctx.assets.paint(n.id);
    if (image) return { backdrop: { image, w: n.box.w, h: n.box.h, x: -(margins[3] ?? 0), y: 0 } };
  }
  return { fill: n.fill, borders: n.border };
};

const isLine = (n: LayoutNode) => n.k === 'shape' && n.box.h <= 2.5 && n.box.w > n.box.h * 4;
const contains = (outer: Box, inner: Box) =>
  inner.x >= outer.x - TOL && inner.y >= outer.y - TOL && right(inner) <= right(outer) + TOL && bottom(inner) <= bottom(outer) + TOL;

const hasText = (n: LayoutNode): boolean => n.k === 'text' || (n.k === 'box' && n.kids.some(hasText));

/**
 * Padding da caixa como margem da célula — encolhido onde o conteúdo vaza
 * (texto "nowrap" maior que a caixa, como o valor "R$ 9.800,00"): no Studio
 * ele passa por cima do padding; no Word quebraria a linha.
 */
const cellPadding = (box: Box, pad: number[], kids: LayoutNode[]): [number, number, number, number] => {
  const items = leaves(kids);
  if (!items.length) return [pad[0], pad[1], pad[2], pad[3]];
  const content = union(items);
  const leftRoom = content.x - box.x;
  const rightRoom = right(box) - right(content);
  // Conteúdo encostado na borda (texto que ocupa a largura exata): 3px de
  // folga — qualquer arredondamento do Word quebraria a linha.
  // Quando passa do padding (texto "nowrap" vazando), ou texto de uma linha
  // encostado (aba do menu). Texto que quebra linha encostado é o normal —
  // tirar 3px ali fazia caber uma palavra a mais por linha no Word.
  const oneLine = items.every((n) => n.k !== 'text' || (n.paras.length === 1 && n.box.h < (n.paras[0].lineHeight ?? n.paras[0].fontSize * 1.2) * 1.5));
  const tight = rightRoom < pad[1] - 0.5 || (oneLine && rightRoom <= pad[1] + 1) ? 3 : 0;
  return [pad[0], Math.max(0, Math.min(pad[1], rightRoom) - tight), pad[2], Math.max(0, Math.min(pad[3], leftRoom))];
};

/** Caixas sem decoração não existem no Word: os filhos sobem pro pai. */
const flatten = (nodes: LayoutNode[]): LayoutNode[] =>
  nodes.flatMap((n) => {
    if (n.k !== 'box' || decorated(n) || n.bgImage) return [n];
    // Composição ocupa, na vertical, o lugar do conteúdo dela (a coluna de
    // 216px com uma bola de 48px no meio); na horizontal, o lugar dela no
    // fluxo (a bola vaza pros cards vizinhos, mas a coluna tem 24px).
    if (complexRoot(n) || rowFlex(n)) {
      const c = union(leaves(n.kids).length ? leaves(n.kids) : [n]);
      return [{ ...n, box: { x: n.box.x, w: n.box.w, y: c.y, h: c.h } }];
    }
    return flatten(n.kids);
  });

/** Só as folhas e caixas pintadas — a geometria do conteúdo, atravessando tudo. */
const allParagraphs = (items: OutIR[]): PIR[] =>
  items.flatMap((item) => (item.kind === 'p' ? [item] : item.rows.flatMap((r) => r.cells.flatMap((c) => allParagraphs(c.children)))));

const isEmptyP = (k: OutIR) => k.kind === 'p' && !k.runs.length && !k.numbering && !k.bookmark && !k.borderBottom && !k.borderTop;
const painted = (c: CellIR) => Boolean(c.fill || c.backdrop || c.borders);
const filler = (width: number): CellIR => ({ width, children: [] });

/** Encosta `w` px de espaço vazio num lado da linha: alarga a célula da ponta se ela não tem pintura, senão célula vazia. */
const padRow = (cells: CellIR[], w: number, side: 'left' | 'right') => {
  if (w < 0.5) return;
  const edge = side === 'left' ? cells[0] : cells[cells.length - 1];
  // Poucos px: alarga a própria célula da ponta (mesmo pintada) — uma
  // célula-fiapo sem cor deixava um recorte branco no canto do card.
  if (edge && !edge.vmerge && (!painted(edge) || w < 6)) {
    const m = edge.margins ?? [0, 0, 0, 0];
    m[side === 'left' ? 3 : 1] += w;
    edge.margins = m;
    edge.width += w;
  } else if (side === 'left') cells.unshift(filler(w));
  else cells.push(filler(w));
};

/**
 * Tabela dentro de tabela vira uma grade só (padrão da equipe: uma célula por
 * elemento, mesclando colunas/linhas). Linha única cujas células guardam cada
 * uma UMA tabela com o mesmo número de linhas (cards lado a lado, cabeçalho +
 * corpo): as linhas de dentro viram linhas da tabela de fora; célula sem
 * tabela (a seta entre os cards) atravessa as linhas mesclada.
 */
const flattenGrid = (t: TIR): TIR => {
  if (t.rows.length !== 1) return t;
  const outer = t.rows[0];
  const nestedOf = (c: CellIR): TIR | null => {
    const real = c.children.filter((k) => !(k.kind === 'p' && (k.spacer !== undefined ? k.spacer < 1 : isEmptyP(k))));
    // Tabela deslocada pra baixo dentro da célula (espaçador antes) perderia o deslocamento.
    const lead = c.children[0];
    if (lead?.kind === 'p' && lead.spacer !== undefined && lead.spacer >= 1) return null;
    if (real.length !== 1 || real[0].kind !== 't' || painted(c) || c.vmerge) return null;
    return real[0];
  };
  const nested = outer.cells.map(nestedOf);
  const n = nested.find(Boolean)?.rows.length;
  if (!n || nested.some((x) => x && x.rows.length !== n)) return t;
  // Só quando a tabela de dentro ocupa a célula toda na altura (cabeçalho +
  // corpo do card); um selo pequeno no topo da célula esticaria até o pé.
  const tall = (x: TIR) => x.rows.reduce((a, r) => a + (r.height ?? 0), 0);
  // Card com várias linhas (cabeçalho + corpo) pode ser um pouco mais baixo que o vizinho.
  if (!outer.height || nested.some((x) => x && tall(x) < outer.height! * (n > 1 ? 0.8 : 1) - 2)) return t;
  const rows = Array.from({ length: n }, (_, r) => ({ height: Math.max(0, ...nested.map((x) => x?.rows[r].height ?? 0)) || undefined, cells: [] as CellIR[] }));
  // A última linha completa a altura da linha de fora (coluna mais alta que a soma das de dentro).
  const inner = rows.slice(0, -1).reduce((a, r) => a + (r.height ?? 0), 0);
  if (outer.height) rows[n - 1].height = Math.max(rows[n - 1].height ?? 0, outer.height - inner);
  outer.cells.forEach((c, i) => {
    const x = nested[i];
    if (!x) {
      // Mesclada na vertical: o Word desenha borda e fundo por pedaço — cada
      // pedaço leva o fundo e as laterais; o de cima a borda de cima, o
      // último a de baixo (senão o card ficava sem borda embaixo e do lado).
      const b = c.borders ?? {};
      const part = (r: number): CellIR['borders'] =>
        c.borders ? { left: b.left, right: b.right, ...(r === 0 ? { top: b.top } : {}), ...(r === n - 1 ? { bottom: b.bottom } : {}) } : undefined;
      rows.forEach((row, r) =>
        row.cells.push(
          r === 0
            ? { ...c, borders: part(0), vmerge: n > 1 ? 'restart' : undefined }
            : { width: c.width, margins: c.margins, fill: c.fill, borders: part(r), vmerge: 'continue', children: [] },
        ),
      );
      return;
    }
    const left = (c.margins?.[3] ?? 0) + x.indent;
    const right = c.width - left - x.width;
    x.rows.forEach((xr, r) => {
      const cells = xr.cells.map((k) => ({ ...k, margins: k.margins ? [...k.margins] as CellIR['margins'] : undefined }));
      padRow(cells, left, 'left');
      padRow(cells, right, 'right');
      rows[r].cells.push(...cells);
    });
  });
  return { ...t, rows };
};

/**
 * Tabela de uma célula que só embrulha outra tabela (a caixa do item com a
 * linha embaixo, em volta de "ícone | texto"): some a de fora — fundo,
 * bordas e margens passam pra de dentro. Uma tabela por estrutura, sem
 * linhas fantasma em cima nem tabela dentro de tabela.
 */
const unwrap = (t: TIR): TIR => {
  if (t.rows.length !== 1 || t.rows[0].cells.length !== 1) return t;
  const outer = t.rows[0].cells[0];
  if (outer.backdrop || outer.vmerge) return t;
  const real = outer.children.filter((k) => !(k.kind === 'p' && (k.spacer !== undefined || isEmptyP(k)) && !k.bookmark));
  if (real.length !== 1 || real[0].kind !== 't') return t;
  const inner = real[0];
  // Espaços antes/depois da tabela de dentro viram margem de cima/baixo.
  const at = outer.children.indexOf(inner);
  const gapOf = (ks: OutIR[]) => ks.reduce((a, k) => a + (k.kind === 'p' && k.spacer !== undefined ? k.spacer : 0), 0);
  const [mt0, , mb0, ml] = outer.margins ?? [0, 0, 0, 0];
  const mt = mt0 + gapOf(outer.children.slice(0, at));
  const mb = mb0 + gapOf(outer.children.slice(at + 1));
  const b = outer.borders ?? {};
  const rows = inner.rows.map((r, ri) => {
    const cells = r.cells.map((c) => ({ ...c, margins: c.margins ? ([...c.margins] as CellIR['margins']) : undefined, fill: c.fill ?? outer.fill }));
    padRow(cells, ml + inner.indent, 'left');
    padRow(cells, outer.width - ml - inner.indent - inner.width, 'right');
    for (const c of cells) if (!c.fill && outer.fill) c.fill = outer.fill;
    const first = ri === 0;
    const last = ri === inner.rows.length - 1;
    cells.forEach((c, ci) => {
      const m = c.margins ?? [0, 0, 0, 0];
      c.margins = [m[0] + (first ? mt : 0), m[1], m[2] + (last ? mb : 0), m[3]];
      // A arte fica ancorada no topo da célula: acompanha a margem nova.
      if (first && c.backdrop && mt) c.backdrop = { ...c.backdrop, y: (c.backdrop.y ?? 0) + mt };
      c.borders = {
        ...c.borders,
        ...(first && b.top ? { top: b.top } : {}),
        ...(last && b.bottom ? { bottom: b.bottom } : {}),
        ...(ci === 0 && b.left ? { left: b.left } : {}),
        ...(ci === cells.length - 1 && b.right ? { right: b.right } : {}),
      };
    });
    return { ...r, cells };
  });
  const innerH = inner.rows.reduce((a, r) => a + (r.height ?? 0), 0);
  const outerH = t.rows[0].height ?? 0;
  const lastRow = rows[rows.length - 1];
  if (outerH > innerH + mt + mb) lastRow.height = (lastRow.height ?? 0) + (outerH - innerH - mt - mb);
  return { ...t, rows };
};

/**
 * Célula cujo conteúdo é só uma tabela de uma linha (card colorido com
 * "selo | texto" dentro, ao lado de um ícone): as células de dentro entram
 * na própria linha de fora, herdando fundo, bordas e margens.
 */
const inlineCells = (t: TIR): TIR => {
  const rows = t.rows.map((row) => {
    const cells: CellIR[] = [];
    for (const c of row.cells) {
      const real = c.children.filter((k) => !(k.kind === 'p' && (k.spacer !== undefined || isEmptyP(k)) && !k.bookmark));
      const inner = real.length === 1 && real[0].kind === 't' && real[0].rows.length === 1 ? real[0] : null;
      // Tabelinha 1×1 (a bola da seta na coluna do meio, mesclada): o conteúdo
      // e a arte passam pra própria célula.
      if (inner && inner.rows[0].cells.length === 1 && !painted(c) && !inner.rows[0].cells[0].vmerge) {
        const ic = inner.rows[0].cells[0];
        const at = c.children.indexOf(inner);
        const before = c.children.slice(0, at).reduce((a, k) => a + (k.kind === 'p' && k.spacer !== undefined ? k.spacer : 0), 0);
        const dx = inner.indent;
        cells.push({
          ...c,
          fill: ic.fill,
          borders: ic.borders,
          backdrop: ic.backdrop ? { ...ic.backdrop, x: (ic.backdrop.x ?? 0) + dx, y: (ic.backdrop.y ?? 0) + before } : undefined,
          children: [...c.children.slice(0, at), ...ic.children, ...c.children.slice(at + 1)],
        });
        continue;
      }
      if (!inner || c.backdrop || c.vmerge || inner.rows[0].cells.some((k) => k.vmerge)) {
        cells.push(c);
        continue;
      }
      const at = c.children.indexOf(inner);
      const gapOf = (ks: OutIR[]) => ks.reduce((a, k) => a + (k.kind === 'p' && k.spacer !== undefined ? k.spacer : 0), 0);
      const [mt0, , mb0, ml] = c.margins ?? [0, 0, 0, 0];
      const mt = mt0 + gapOf(c.children.slice(0, at));
      const mb = mb0 + gapOf(c.children.slice(at + 1));
      const b = c.borders ?? {};
      const got = inner.rows[0].cells.map((k) => ({ ...k, margins: k.margins ? ([...k.margins] as CellIR['margins']) : undefined, fill: k.fill ?? c.fill }));
      padRow(got, ml + inner.indent, 'left');
      padRow(got, c.width - ml - inner.indent - inner.width, 'right');
      got.forEach((k, ki) => {
        if (!k.fill && c.fill) k.fill = c.fill;
        const m = k.margins ?? [0, 0, 0, 0];
        k.margins = [m[0] + mt, m[1], m[2] + mb, m[3]];
        if (k.backdrop && mt) k.backdrop = { ...k.backdrop, y: (k.backdrop.y ?? 0) + mt };
        k.vAlign = k.vAlign ?? c.vAlign;
        k.borders = {
          ...k.borders,
          ...(b.top ? { top: b.top } : {}),
          ...(b.bottom ? { bottom: b.bottom } : {}),
          ...(ki === 0 && b.left ? { left: b.left } : {}),
          ...(ki === got.length - 1 && b.right ? { right: b.right } : {}),
        };
      });
      cells.push(...got);
    }
    return { ...row, cells };
  });
  return { ...t, rows };
};

/** Mesma faixa horizontal (uma dentro da outra ou quase): dá pra ser uma tabela só. */
const canJoin = (a: TIR, b: TIR) => Math.min(a.indent + a.width, b.indent + b.width) - Math.max(a.indent, b.indent) > 0.5 * Math.min(a.width, b.width);

/** Junta `b` embaixo de `a`: as duas passam a cobrir a mesma faixa (células vazias/alargadas nas pontas). */
const join = (a: TIR, b: TIR) => {
  const L = Math.min(a.indent, b.indent);
  const R = Math.max(a.indent + a.width, b.indent + b.width);
  for (const t of [a, b]) {
    for (const row of t.rows) {
      padRow(row.cells, t.indent - L, 'left');
      padRow(row.cells, R - (t.indent + t.width), 'right');
    }
    t.indent = L;
    t.width = R - L;
  }
  a.rows.push(...b.rows);
};

/**
 * Vão entre duas estruturas vira espaço dentro da tabela: no fim da linha de
 * cima se ela não tem fundo nem linha embaixo, no começo da de baixo se ela
 * não tem fundo nem linha em cima; senão, uma linha vazia só pro vão.
 */
const absorb = (a: TIR, b: TIR, gap: number): boolean => {
  const last = a.rows[a.rows.length - 1];
  const first = b.rows[0];
  const clear = (row: TIR['rows'][number], side: 'top' | 'bottom') => row.cells.every((c) => !c.fill && !c.backdrop && !c.borders?.[side]);
  if (clear(last, 'bottom')) {
    for (const c of last.cells) c.children = [...c.children, spacer(gap)];
    if (last.height) last.height += gap;
  } else if (clear(first, 'top')) {
    for (const c of first.cells) c.children = [spacer(gap), ...c.children];
    if (first.height) first.height += gap;
  } else {
    b.rows.unshift({ height: gap, cells: [{ width: b.width, children: [spacer(gap)] }] });
  }
  return true;
};

/**
 * Estruturas empilhadas na mesma faixa viram UMA tabela, uma linha cada
 * (padrão da equipe: "mesma estrutura visual em uma mesma tabela") — coladas
 * ou com um vão que cabe numa linha sem pintura. Some também o parágrafo
 * mínimo que o Word exige entre duas tabelas.
 */
/**
 * Primeira linha da tabela que é só um texto de ponta a ponta, sem fundo nem
 * borda ("OBRIGAÇÕES DA CONTRATADA" em cima do quadro): sai da tabela como
 * texto normal na página.
 */
const liftTitle = (t: TIR): OutIR[] => {
  if (t.rows.length < 2) return [t];
  const first = t.rows[0];
  const content = first.cells.filter((c) => c.children.some((k) => !(k.kind === 'p' && isEmptyP(k))));
  if (content.length !== 1 || first.cells.some(painted) || first.cells.some((c) => c.vmerge)) return [t];
  const c = content[0];
  if (!c.children.every((k) => k.kind === 'p')) return [t];
  const offset = t.indent + first.cells.slice(0, first.cells.indexOf(c)).reduce((a, x) => a + x.width, 0) + (c.margins?.[3] ?? 0);
  const ps = (c.children as PIR[]).map((p) => ({ ...p, indentLeft: (p.indentLeft ?? 0) + offset }));
  return [...ps, { ...t, rows: t.rows.slice(1) }];
};

const mergeTables = (items: OutIR[]): OutIR[] => {
  const out: OutIR[] = [];
  for (let item of items) {
    if (item.kind === 't') {
      item.rows.forEach((r) => r.cells.forEach((c) => (c.children = mergeTables(c.children))));
      item = inlineCells(unwrap(flattenGrid(unwrap(item))));
    }
    const prev = out[out.length - 1];
    const prev2 = out[out.length - 2];
    if (item.kind === 't' && prev?.kind === 't' && canJoin(prev, item)) {
      join(prev, item);
      continue;
    }
    if (item.kind === 't' && prev?.kind === 'p' && prev.spacer !== undefined && !prev.bookmark && prev2?.kind === 't' && canJoin(prev2, item) && absorb(prev2, item, prev.spacer)) {
      out.pop();
      join(prev2, item);
      continue;
    }
    out.push(item);
  }
  return out.flatMap((k) => (k.kind === 't' ? liftTitle(k) : [k]));
};

const everyNode = (n: LayoutNode): LayoutNode[] => [n, ...(n.k === 'box' ? n.kids.flatMap(everyNode) : [])];
/** Imagens soltas: a própria imagem ou alguma caixa em volta dela tem posição absoluta. */
const looseImages = (n: LayoutNode, inAbs = false): LayoutNode[] =>
  n.k === 'img' || n.k === 'raster' ? (inAbs || n.abs ? [n] : []) : n.k === 'box' ? n.kids.flatMap((k) => looseImages(k, inAbs || Boolean(n.abs))) : [];

/** Cópia da árvore sem os nós dados. */
const without = (n: LayoutNode, drop: Set<LayoutNode>): LayoutNode =>
  n.k === 'box'
    ? {
        ...n,
        kids: n.kids
          .filter((k) => !drop.has(k))
          .map((k) => without(k, drop))
          // Caixa que ficou vazia (só embrulhava o que saiu) some junto.
          .filter((k) => k.k !== 'box' || k.kids.length || decorated(k) || k.bgImage),
      }
    : n;

const rgb = (hex: string) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
const near = (a: string, b: string) => rgb(a).every((v, i) => Math.abs(v - rgb(b)[i]) <= 12);

/**
 * Cópia da árvore sem fundo da mesma cor da página embaixo (branco em página
 * branca, azul no Quadro resumo azul): não se vê — e viraria tabela à toa
 * (a equipe quer título como texto na página).
 */
const stripSame = (n: LayoutNode, page: string): LayoutNode => {
  if (n.k !== 'box' && n.k !== 'text' && n.k !== 'shape') return n;
  const fill = (n as { fill?: string }).fill;
  const plain = fill && near(fill, page) && !(n as { border?: object }).border && !(n as { paint?: boolean }).paint;
  const copy = plain ? { ...n, fill: undefined } : { ...n };
  // Dentro de uma caixa pintada o que está embaixo já não é a página: o selo
  // branco em cima do card colorido continua branco.
  if (copy.k !== 'box') return copy;
  return fill && !plain ? copy : { ...copy, kids: copy.kids.map((k) => stripSame(k, page)) };
};

/**
 * Conteúdo que passa da largura disponível (caixa larga dentro de uma coluna
 * estreitada): a tabela encolhe na célula mais larga — o Word corta o que
 * passa da célula, não quebra.
 */
const fit = (items: OutIR[], width: number): void => {
  for (const item of items) {
    if (item.kind !== 't') continue;
    const over = item.indent + item.width - width;
    if (over > 0.5) {
      const shift = Math.min(item.indent, over);
      item.indent -= shift;
      const cut = over - shift;
      if (cut > 0.5) {
        for (const row of item.rows) {
          const widest = row.cells.reduce((a, c) => (c.width > a.width ? c : a));
          widest.width = Math.max(1, widest.width - cut);
        }
        item.width -= cut;
      }
    }
    for (const row of item.rows) for (const c of row.cells) fit(c.children, c.width - (c.margins ? c.margins[1] + c.margins[3] : 0));
  }
};

/**
 * Entre duas tabelas do corpo, o espaço é uma linha em branco de verdade
 * (estilo Normal, cresce quando alguém digita e empura a tabela de baixo).
 * Espaçador de altura fixa ali fazia o texto digitado sumir atrás da tabela.
 */
/**
 * Estruturas coladas no Studio (vão de até 4px — as linhas do Sumário, cada
 * uma um bloco): uma tabela só, uma linha cada. Só entre estruturas com
 * espaço visível fica a linha em branco digitável (separate).
 */
const joinTight = (items: OutIR[]): OutIR[] => {
  const out: OutIR[] = [];
  for (const item of items) {
    const prev = out[out.length - 1];
    const prev2 = out[out.length - 2];
    if (item.kind === 't' && prev?.kind === 't' && canJoin(prev, item)) {
      join(prev, item);
      continue;
    }
    const tiny = prev?.kind === 'p' && prev.spacer !== undefined && prev.spacer <= 4 && !prev.bookmark && !prev.pageBreakBefore && !prev.runs.length;
    if (item.kind === 't' && tiny && prev2?.kind === 't' && canJoin(prev2, item) && absorb(prev2, item, (prev as PIR).spacer!)) {
      out.pop();
      join(prev2, item);
      continue;
    }
    out.push(item);
  }
  return out;
};

const separate = (items: OutIR[]): OutIR[] =>
  items.map((item, i) =>
    item.kind === 'p' && item.spacer !== undefined && item.spacer >= 1 && !item.runs.length && !item.pageBreakBefore && i > 0
      ? growGap(item.spacer, item)
      : item,
  );

/**
 * Vão entre tabelas que cresce: altura "pelo menos" o vão do Studio, com a
 * fonte da marca do tamanho que cabe nele. Quem digita ali (ou aumenta a
 * fonte) empurra a tabela de baixo — altura fixa escondia o texto atrás dela.
 */
const growGap = (gap: number, from: PIR): PIR => ({
  kind: 'p',
  runs: [],
  keepNext: from.keepNext,
  bookmark: from.bookmark,
  line: { value: Math.max(MIN_LINE, tw(gap)), rule: 'atLeast' },
  grow: Math.max(2, Math.min(144, Math.round((gap / simpleRatio) * 1.5))),
});

const leaves = (nodes: LayoutNode[]): LayoutNode[] =>
  nodes.flatMap((n) => (n.k === 'box' && !decorated(n) && !n.bgImage ? leaves(n.kids) : [n]));

/**
 * Caixa com peças soltas (posição absoluta) dentro: é ela — só ela — que vira
 * composição (desenho capturado + texto por cima). O resto do bloco segue em
 * tabela/forma nativa.
 */
/**
 * Linha flex do Studio (cards lado a lado): cada filho é uma coluna, inteira.
 * Fatiar por faixas horizontais separava os cabeçalhos dos corpos dos cards
 * em tabelas diferentes — as larguras não batiam e o card desalinhava.
 */
const rowFlex = (n: LayoutNode): boolean =>
  n.k === 'box' &&
  /flex/.test(n.layout.display) &&
  !/column/.test(n.layout.dir) &&
  n.kids.length >= 2 &&
  !n.kids.some((k) => k.abs) &&
  bands(n.kids, 'x').length === n.kids.length;

/**
 * Retângulo colorido solto com peças soltas em cima (padrão do Figma: fundo
 * e conteúdo posicionados à parte) → o retângulo vira a caixa pintada que
 * contém essas peças. Aí o card é tabela com fundo preenchido, não forma.
 */
const nestShapes = (n: LayoutNode): LayoutNode => {
  if (n.k !== 'box') return n;
  let kids = n.kids.map(nestShapes);
  const shapes = kids.filter((k) => k.k === 'shape' && !isLine(k) && (k as { fill?: string }).fill);
  for (const shape of shapes.sort((a, b) => a.box.w * a.box.h - b.box.w * b.box.h)) {
    if (!kids.includes(shape)) continue;
    const inside = kids.filter((k) => k !== shape && k.k !== 'shape' && contains(shape.box, k.box));
    // Só é fundo de card se tem texto em cima. Forma só com ícone (a bola
    // branca com a seta entre dois cards) é enfeite: vai como desenho.
    if (!inside.some(hasText)) continue;
    kids = kids.filter((k) => !inside.includes(k) && k !== shape);
    kids.push({
      ...(shape as Extract<LayoutNode, { k: 'shape' }>),
      k: 'box',
      kids: inside,
      layout: { display: 'block', dir: 'row', alignItems: 'normal', justify: 'normal', textAlign: 'left' },
    } as LayoutNode);
  }
  return { ...n, kids };
};

const overlaps = (a: Box, b: Box) => {
  const w = Math.min(right(a), right(b)) - Math.max(a.x, b.x);
  const h = Math.min(bottom(a), bottom(b)) - Math.max(a.y, b.y);
  return w > 1 && h > 1 && w * h > 0.1 * Math.min(a.w * a.h, b.w * b.h);
};
const overflow = (k: Box, n: Box) => Math.max(n.x - k.x, n.y - k.y, right(k) - right(n), bottom(k) - bottom(n));
/**
 * Só quando as peças soltas se sobrepõem de verdade (texto sobre desenho, a
 * bola sobre dois cards) ou vazam muito da caixa (ícone de 80px num título de
 * 22px). Posição absoluta que só posiciona (layout do Figma: número e título
 * lado a lado, linha embaixo) continua tabela nativa — a linha vira borda.
 */
const complexRoot = (n: LayoutNode): boolean => {
  if (n.k !== 'box' || !n.kids.some((k) => k.abs)) return false;
  const kids = n.kids;
  // Compara o que se vê (textos, selos, fundos), não as caixas invisíveis que
  // embrulham: o ícone no canto vazio de um card não é sobreposição.
  const seen = (k: LayoutNode) => leaves([k]).filter((l) => !isLine(l));
  if (kids.some((a, i) => kids.some((b, j) => j > i && (a.abs || b.abs) && seen(a).some((x) => seen(b).some((y) => overlaps(x.box, y.box)))))) return true;
  // Vazamento só conta pra desenho (ícone de 80px num título de 22px); texto
  // posicionado fora da caixa é só o jeito do Figma de posicionar.
  return kids.some((k) => k.abs && !hasText(k) && overflow(k.box, n.box) > Math.max(16, 0.4 * Math.min(k.box.w, k.box.h)));
};

/** Agrupa intervalos que se sobrepõem (em y ou x) em faixas. */
const bands = (items: LayoutNode[], axis: 'y' | 'x'): LayoutNode[][] => {
  const start = (n: LayoutNode) => (axis === 'y' ? n.box.y : n.box.x);
  const end = (n: LayoutNode) => (axis === 'y' ? bottom(n.box) : right(n.box));
  const sorted = [...items].sort((a, b) => start(a) - start(b));
  const out: LayoutNode[][] = [];
  let limit = -Infinity;
  for (const n of sorted) {
    if (out.length && start(n) < limit - TOL) {
      out[out.length - 1].push(n);
      limit = Math.max(limit, end(n));
    } else {
      out.push([n]);
      limit = end(n);
    }
  }
  return out;
};

const union = (items: LayoutNode[]): Box => {
  const x = Math.min(...items.map((n) => n.box.x));
  const y = Math.min(...items.map((n) => n.box.y));
  return { x, y, w: Math.max(...items.map((n) => right(n.box))) - x, h: Math.max(...items.map((n) => bottom(n.box))) - y };
};

// ---------------------------------------------------------------------------
// Texto
// ---------------------------------------------------------------------------

/**
 * Entrelinha automática do Word (a equipe edita em "Simples"): cresce e
 * encolhe com a fonte. Igual à natural da fonte → Simples; o Studio mais
 * aberto (1,65 na Caixa) → Múltiplos com o fator que dá a mesma altura. A
 * altura natural é a do Windows (usWin), onde a equipe e os clientes editam.
 */
const lineOf = (para: Para, face: Face | null): PIR['line'] => {
  if (!para.lineHeight) return undefined;
  const natural = (face?.lineHeight ?? 1.2) * para.fontSize;
  const m = para.lineHeight / natural;
  return { value: Math.abs(m - 1) < 0.08 ? 240 : Math.round(240 * Math.max(0.8, m)), rule: 'auto' };
};

/**
 * Espaçamento entre letras que faz o trecho ocupar no Word a mesma largura do
 * Studio (o Chromium arredonda o avanço de cada letra; o Word não). Inclui o
 * letter-spacing do CSS. twips por caractere; undefined sem métrica.
 */
const spacingOf = (r: Run, face: Face | null, ctx: Ctx, justified = false): number | undefined => {
  const css = r.letterSpacing ? Math.round(r.letterSpacing * 15) : undefined;
  if (!face || !r.t) return css;
  const metric = ctx.fonts.data.metrics.get(face);
  if (!metric) return css;
  const text = r.caps ? r.t.toUpperCase() : r.t;
  // Tamanho que o Word desenha: ele só aceita meio ponto (13px = 9,75pt vira
  // 10pt, 2,5% mais largo) — a diferença entra no espaçamento.
  const wordSize = halfPt(r.size) / 1.5;
  let studio = 0;
  let word = 0;
  let count = 0;
  for (const ch of text) {
    const adv = metric[ch.codePointAt(0)!];
    if (adv === undefined) return css;
    studio += adv * r.size;
    word += adv * wordSize;
    count++;
  }
  if (!count || !text.trim()) return undefined;
  // Largura de referência: a medida no Studio (pega arredondamento e kerning
  // do Chromium), somando o espaço onde a linha quebrou — ele não entra na
  // medida. Justificado com quebra estica os espaços: aí vale a largura
  // natural + letter-spacing.
  const space = (metric[32] ?? 0.25) * r.size;
  const lines = r.lines ?? 1;
  const target =
    r.width && !(justified && lines > 1) ? r.width + (lines - 1) * space : studio + (r.letterSpacing ?? 0) * count;
  const perChar = (target - word) / count;
  // Diferença grande demais não é arredondamento — é outra coisa.
  if (Math.abs(perChar) > r.size * 0.25) return css;
  return Math.round(perChar * 15);
};

// ---------------------------------------------------------------------------
// Listas manuais: "i)", "(ii)", "a.", "3)" digitados como texto no Studio
// viram numeração do Word — duplicou a linha, o número já vem certo.
// ---------------------------------------------------------------------------

const MARKER = /^\s*(\(?)([ivxlcdm]+|[a-z]|\d{1,2})([).])(\s*)/i;
const ROMAN: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
const romanValue = (s: string): number | null => {
  const v = [...s.toLowerCase()].map((c) => ROMAN[c]);
  let total = 0;
  for (let i = 0; i < v.length; i++) total += v[i] < (v[i + 1] ?? 0) ? -v[i] : v[i];
  return total > 0 && total < 40 ? total : null;
};

/** Marcador → formato e valor. Letra única que também é romano (i, v, x) decide pela sequência em andamento. */
const parseMarker = (token: string, ctx: Ctx): { format: ListFormat; value: number } | null => {
  const upper = token === token.toUpperCase() && /[a-z]/i.test(token);
  if (/^\d+$/.test(token)) return { format: 'decimal', value: Number(token) };
  const letter = { format: (upper ? 'upperLetter' : 'lowerLetter') as ListFormat, value: token.toLowerCase().charCodeAt(0) - 96 };
  const roman = /^[ivxlcdm]+$/i.test(token) ? romanValue(token) : null;
  const seq = ctx.seq && ctx.lists.get(ctx.seq.key);
  if (token.length === 1 && seq?.format.endsWith('Letter') && letter.value === ctx.seq!.value + 1) return letter;
  if (roman) return { format: upper ? 'upperRoman' : 'lowerRoman', value: roman };
  return token.length === 1 ? letter : null;
};

/**
 * Parágrafo que começa com marcador de lista: tira o marcador do texto e
 * devolve a numeração (instância nova quando a sequência recomeça no 1).
 * Marcador fora de sequência (pulou número) fica como texto.
 */
const takeMarker = (p: PIR, ctx: Ctx): boolean => {
  const first = p.runs.find((r) => r.kind === 'text') as Extract<RunIR, { kind: 'text' }> | undefined;
  if (!first || p.numbering) return false;
  const m = first.text.match(MARKER);
  if (!m) return false;
  const alone = p.runs.filter((r) => r.kind === 'text').map((r) => (r as { text: string }).text).join('').trim() === m[0].trim();
  // Marcador no começo de frase precisa de espaço depois ("(i) O pedido"); sozinho, basta ele.
  if (!alone && !m[4]) return false;
  const parsed = parseMarker(m[2], ctx);
  if (!parsed) return false;
  const suffix = alone ? 'nothing' : 'space';
  // Nome sem símbolo: a lib docx monta regex com ele.
  const key = `${parsed.format}-${m[1] ? 'paren' : ''}${m[3] === ')' ? 'close' : 'dot'}-${suffix}`;
  let instance: number;
  if (ctx.seq?.key === key && parsed.value === ctx.seq.value + 1) instance = ctx.seq.instance;
  else if (parsed.value === 1) instance = ++ctx.instances;
  else return false;
  ctx.lists.set(key, { format: parsed.format, before: m[1], after: m[3], suffix });
  ctx.seq = { key, value: parsed.value, instance };
  first.text = first.text.slice(m[0].length);
  if (!first.text) {
    p.runs.splice(p.runs.indexOf(first), 1);
    const next = p.runs.find((r) => r.kind === 'text') as { text: string } | undefined;
    if (next) next.text = next.text.replace(/^\s+/, '');
  }
  p.mark = { run: { ...first.run, letterSpacing: first.spacing !== undefined ? first.spacing / 15 : first.run.letterSpacing }, face: first.face };
  p.numbering = { ref: `list:${key}`, level: 0, instance };
  return true;
};

/**
 * Número de cláusula digitado à mão ("13.9." num selo) que continua a
 * sequência automática: vira item da lista de cláusulas do Word — senão os
 * subitens automáticos seguintes saíam "13.8.1" (o Word não sabia do 13.9).
 */
const manualClause = (node: Extract<LayoutNode, { k: 'text' }>, p: PIR, ctx: Ctx): boolean => {
  const text = p.runs.map((r) => (r.kind === 'text' ? r.text : '')).join('').trim();
  const m = node.paras.length === 1 && text.match(/^(\d+(?:\.\d+){0,2})\.?$/);
  if (!m || !ctx.clauses.length) return false;
  const parts = m[1].split('.').map(Number);
  const level = parts.length - 1;
  const expected = [...ctx.clauses.slice(0, level), (ctx.clauses[level] ?? 0) + 1];
  if (parts.join('.') !== expected.join('.')) return false;
  const first = p.runs.find((r) => r.kind === 'text') as Extract<RunIR, { kind: 'text' }>;
  p.mark = { run: { ...first.run, letterSpacing: first.spacing !== undefined ? first.spacing / 15 : first.run.letterSpacing }, face: first.face };
  p.runs = [];
  p.numbering = { ref: 'clauses', level };
  ctx.clauses = parts;
  return true;
};

const textToParagraphs = (node: Extract<LayoutNode, { k: 'text' }>, region: Box, ctx: Ctx): PIR[] => {
  const out: PIR[] = [];
  const indentLeft = node.box.x - region.x;
  // A largura só manda na quebra quando o texto quebrou linha no Studio; aí
  // vale ela (com folga de 2px pra diferença mínima de métrica). Texto de uma
  // linha fica livre — prender a largura exata quebrava palavra ("Banc|o").
  const lh = node.paras[0]?.lineHeight ?? node.paras[0]?.fontSize * 1.2;
  const wrapped = node.paras.length > 1 || (lh ? node.box.h > lh * 1.5 : false);
  const indentRight = wrapped ? Math.max(0, right(region) - right(node.box) - 2) : 0;
  // Uma linha centralizada/à direita: o centro (ou a borda direita) fica onde
  // estava, com folga dos dois lados pra não quebrar se o Word desenhar maior.
  const slack = node.box.w * 0.15 + 4;
  const rightRoom = Math.max(0, right(region) - right(node.box));
  const oneLine = (align: Para['align']) =>
    wrapped || (align !== 'center' && align !== 'right')
      ? null
      : align === 'center'
        ? { left: Math.max(0, indentLeft - slack), right: Math.max(0, rightRoom - slack) }
        : { left: Math.max(0, indentLeft - slack), right: rightRoom };
  node.paras.forEach((para) => {
    const runs: RunIR[] = [];
    let firstFace: Face | null = null;
    for (const r of para.runs) {
      if (r.br) {
        runs.push({ kind: 'br' });
        continue;
      }
      if (!r.t) continue;
      const face = ctx.fonts.resolve(r.font, r.weight, r.italic);
      if (face) ctx.used.add(face);
      if (face) {
        const key = `${face.family}|${halfPt(r.size)}`;
        ctx.usage.set(key, (ctx.usage.get(key) ?? 0) + r.t.length);
      }
      firstFace ??= face;
      runs.push({ kind: 'text', text: r.t, run: r, face, spacing: spacingOf(r, face, ctx, para.align === 'both') });
    }
    const first = para.runs.find((r) => !r.br) as Run | undefined;
    out.push({
      kind: 'p',
      runs,
      align: para.align,
      indentLeft: para.list && para.x !== undefined ? para.x - region.x : oneLine(para.align)?.left ?? indentLeft,
      // Marcador pendurado fora do item (lista sem padding): ~1 em antes do texto.
      hanging: para.list && para.x !== undefined ? Math.max(0, Math.min((para.indent ?? 0) >= 4 ? para.indent! : para.fontSize * 1.1, para.x - region.x)) : undefined,
      indentRight: oneLine(para.align)?.right ?? indentRight,
      line: lineOf(para, firstFace),
      mark: first ? { run: first, face: firstFace } : undefined,
      numbering: para.list ? { ref: para.list.kind === 'bullet' ? 'bullet' : 'decimal', level: Math.min(para.list.level, 8) } : undefined,
    });
  });
  if (out.length && !node.autonumber && !manualClause(node, out[0], ctx)) takeMarker(out[0], ctx);
  if (node.autonumber) {
    const level = node.autonumber.split('.').filter(Boolean).length - 1;
    const target = out[0] ?? { kind: 'p', runs: [] };
    if (!out.length) out.push(target);
    // O número é do Word (lista multinível), não texto: a próxima cláusula
    // criada no Word já sai numerada, com a cara do selo (a numeração herda a
    // formatação da marca de parágrafo).
    target.numbering = { ref: 'clauses', level: Math.max(0, Math.min(level, 2)) };
    ctx.clauses = node.autonumber.split('.').filter(Boolean).map(Number);
    const style = node.autonumberStyle;
    if (style) {
      const run: Run = { ...style, t: '', link: null };
      const face = ctx.fonts.resolve(style.font, style.weight, style.italic);
      // O número desenhado pelo Word também sai no meio ponto arredondado
      // ("7.1.1" 2,5% mais largo não cabia no selo): compensa no espaçamento.
      const sp = spacingOf({ ...style, t: node.autonumber, link: null }, face, ctx);
      if (sp !== undefined) run.letterSpacing = sp / 15;
      if (face) ctx.used.add(face);
      target.mark = { run, face };
      if (!target.runs.length && style.lineHeight) {
        target.line = lineOf({ lineHeight: style.lineHeight, fontSize: style.size, align: 'left', runs: [] }, face);
      }
    }
  }
  return out;
};

// ---------------------------------------------------------------------------
// Montagem por fatiamento
// ---------------------------------------------------------------------------

/**
 * Altura dos parágrafos que só existem porque o Word exige (âncora de imagem,
 * fim de célula, separador entre tabelas): 1/10 de ponto. Com 1pt cada, 16
 * linhas do Sumário somavam 60px e a página estourava.
 */
const MIN_LINE = 2;
const spacer = (px: number): PIR => ({ kind: 'p', runs: [], spacer: px, line: { value: Math.max(MIN_LINE, tw(px)), rule: 'exact' } });

/** Espaço vertical antes de um conteúdo: no parágrafo, ou um espaçador antes da tabela. */
const withGap = (out: OutIR[], gap: number): OutIR[] => {
  if (gap < 1 || !out.length) return out;
  const first = out[0];
  if (first.kind === 'p' && !first.spacer) {
    first.spaceBefore = (first.spaceBefore ?? 0) + gap;
    return out;
  }
  return [spacer(gap), ...out];
};

const renderRegion = async (nodes: LayoutNode[], region: Box, ctx: Ctx): Promise<OutIR[]> => {
  let items = flatten(nodes);
  // Linha fina solta vira borda inferior de parágrafo (o traço do Sumário).
  const lines = items.filter(isLine);
  items = items.filter((n) => !isLine(n));
  // Forma colorida que contém outros itens é o fundo deles: vira célula.
  for (const shape of items.filter((n) => n.k === 'shape' && !isLine(n))) {
    const inside = items.filter((n) => n !== shape && contains(shape.box, n.box));
    if (inside.length) {
      items = items.filter((n) => !inside.includes(n) && n !== shape);
      items.push({ ...(shape as Extract<LayoutNode, { k: 'shape' }>), k: 'box', kids: inside, layout: { display: 'block', dir: 'row', alignItems: 'normal', justify: 'normal', textAlign: 'left' } });
    }
  }
  const out: OutIR[] = [];
  let cursor = region.y;
  const strokeOf = (line: LayoutNode) => ({ w: Math.max(0.5, line.box.h), color: (line as { fill?: string }).fill ?? '000000' });
  // Linha fina vira borda do conteúdo vizinho (célula da tabela ou
  // parágrafo) — nunca parágrafo solto: a equipe quer a linha na tabela.
  const edge = (items: OutIR[], side: 'top' | 'bottom', stroke: { w: number; color: string }): boolean => {
    const item = side === 'top' ? items[0] : items[items.length - 1];
    if (!item) return false;
    if (item.kind === 't') {
      const row = side === 'top' ? item.rows[0] : item.rows[item.rows.length - 1];
      for (const cell of row.cells) cell.borders = { ...cell.borders, [side]: stroke };
      return true;
    }
    if (item.spacer) return false;
    if (side === 'bottom' ? item.borderBottom : item.borderTop) return false;
    if (side === 'bottom') item.borderBottom = stroke;
    else item.borderTop = stroke;
    return true;
  };
  const take = (line: LayoutNode) => lines.splice(lines.indexOf(line), 1);
  for (const band of bands(items, 'y')) {
    const b = union(band);
    const rendered = await renderBand(band, { x: region.x, y: b.y, w: region.w, h: b.h }, ctx);
    // Linha logo acima (entre o conteúdo anterior e este): borda de cima.
    for (const line of lines.filter((l) => l.box.y >= cursor - TOL && bottom(l.box) <= b.y + TOL && b.y - bottom(l.box) <= 12)) {
      if (edge(rendered, 'top', strokeOf(line))) take(line);
    }
    out.push(...withGap(rendered, b.y - cursor));
    cursor = bottom(b);
    // Linha dentro da faixa ou logo abaixo: borda de baixo.
    for (const line of lines.filter((l) => l.box.y >= b.y - TOL && l.box.y <= cursor + 12)) {
      if (edge(rendered, 'bottom', strokeOf(line))) {
        take(line);
        cursor = Math.max(cursor, bottom(line.box));
      }
    }
  }
  // Linha sem vizinho (divisória entre dois vãos): parágrafo com borda.
  for (const line of lines.sort((a, b) => a.box.y - b.box.y)) {
    out.push({
      kind: 'p',
      runs: [],
      indentLeft: Math.max(0, line.box.x - region.x),
      indentRight: Math.max(0, right(region) - right(line.box)),
      spaceBefore: Math.max(0, line.box.y - cursor - 1),
      line: { value: 20, rule: 'exact' },
      borderBottom: strokeOf(line),
    });
    cursor = Math.max(cursor, bottom(line.box));
  }
  return out;
};

const renderBand = async (band: LayoutNode[], region: Box, ctx: Ctx): Promise<OutIR[]> => {
  // Arte (imagem/raster) que cobre outros itens é o fundo deles: célula com a
  // imagem atrás do texto, e o conteúdo por cima continua editável.
  const backdrop = band.find(
    (n) => (n.k === 'img' || n.k === 'raster') && band.some((o) => o !== n && hasText(o) && contains(n.box, o.box)),
  );
  if (backdrop && (backdrop.k === 'img' || backdrop.k === 'raster')) {
    const image = backdrop.k === 'img' ? await ctx.assets.image(backdrop.src, backdrop.box.w, backdrop.box.h) : await ctx.assets.raster(backdrop.id, backdrop.box);
    const inside = band.filter((o) => o !== backdrop && contains(backdrop.box, o.box));
    const outside = band.filter((o) => o !== backdrop && !inside.includes(o));
    const children = await renderRegion(inside, backdrop.box, ctx);
    const table: TIR = {
      kind: 't',
      width: backdrop.box.w,
      indent: backdrop.box.x - region.x,
      rows: [{ height: backdrop.box.h, cells: [{ width: backdrop.box.w, backdrop: image ? { image, w: backdrop.box.w, h: backdrop.box.h } : undefined, children: children.length ? children : [{ kind: 'p', runs: [] }] }] }],
    };
    if (!outside.length) return [table];
    return [table, ...(await renderRegion(outside, region, ctx))];
  }
  const cols = bands(band, 'x');
  if (cols.length > 1) return [await renderColumns(cols, region, ctx, [])];
  // Enfeite com posição absoluta que invade as colunas vizinhas (o círculo
  // da seta entre dois cards) impede o corte: sai do corte e volta depois
  // no espaço entre as colunas.
  const overlays = band.filter((n) => n.abs && (n.k === 'img' || n.k === 'raster' || n.k === 'shape' || (n.k === 'box' && !hasText(n))));
  if (overlays.length && overlays.length < band.length) {
    const rest = bands(band.filter((n) => !overlays.includes(n)), 'x');
    if (rest.length > 1) return [await renderColumns(rest, region, ctx, overlays)];
  }
  if (band.length === 1) return renderItem(band[0], region, ctx);
  // Itens sobrepostos: texto por cima de imagem/forma decorativa. Mantém o
  // texto editável e empilha pela ordem vertical.
  const texts = band.filter((n) => n.k === 'text' || n.k === 'box');
  if (texts.length === band.length || texts.length === 0) {
    const out: OutIR[] = [];
    for (const n of [...band].sort((a, b) => a.box.y - b.box.y || a.box.x - b.box.x)) out.push(...(await renderItem(n, region, ctx)));
    return out;
  }
  const out: OutIR[] = [];
  for (const n of texts.sort((a, b) => a.box.y - b.box.y)) out.push(...(await renderItem(n, region, ctx)));
  return out;
};

/** Colunas lado a lado → uma tabela de uma linha, largura exata de cada célula. */
const renderColumns = async (cols: LayoutNode[][], region: Box, ctx: Ctx, overlays: LayoutNode[]): Promise<TIR> => {
  const boxes = cols.map(union);
  const left = boxes[0].x;
  const cells: CellIR[] = [];
  const rowBottom = Math.max(...boxes.map(bottom));
  const rowTop = Math.min(...boxes.map((b) => b.y));
  // Centralizado de verdade só se a geometria diz (folga igual em cima e embaixo).
  const alignOf = (b: Box): CellIR['vAlign'] => {
    const top = b.y - rowTop;
    const low = rowBottom - bottom(b);
    return top > 1.5 && Math.abs(top - low) < 2 ? 'center' : 'top';
  };
  for (let i = 0; i < cols.length; i++) {
    const start = i === 0 ? left : boxes[i].x;
    const end = i + 1 < cols.length ? boxes[i + 1].x : right(region);
    const width = Math.max(1, end - start);
    const col = cols[i];
    const colBox = boxes[i];
    const only = col.length === 1 ? col[0] : null;
    // Coluna que é uma caixa decorada ocupando a coluna: a própria célula
    // leva o fundo/borda/margem (sem tabela dentro de tabela à toa).
    // Só quando a caixa ocupa a altura da linha: um selo pequeno ao lado de
    // texto longo ("1.2") fica caixinha no topo da célula, não fundo da célula.
    // Também quando é a maior parte da altura (quadrado do ícone de 60px num
    // cabeçalho de 82px): no Word a célula pintada ocupa a linha toda e o
    // ícone fica centralizado — tabela dentro da célula desalinhava.
    if (only && only.k === 'box' && decorated(only) && !only.abs && !complexRoot(only) && only.box.h >= Math.min(rowBottom - rowTop - 1.5, (rowBottom - rowTop) * 0.6)) {
      const pad = cellPadding(only.box, only.pad ?? [0, 0, 0, 0], only.kids);
      const inner = { x: only.box.x + pad[3], y: only.box.y + pad[0], w: only.box.w - pad[1] - pad[3], h: only.box.h - pad[0] - pad[2] };
      const children = await renderRegion(only.kids, inner, ctx);
      const gapRight = end - right(only.box);
      const content = only.kids.length ? union(leaves(only.kids).length ? leaves(only.kids) : only.kids) : inner;
      const top = content.y - inner.y;
      const low = bottom(inner) - bottom(content);
      cells.push({
        width: only.box.w,
        ...(await paintOf(only as Decorated, pad, ctx)),
        margins: [pad[0], pad[1], pad[2], pad[3]],
        // Centro só se o conteúdo estava centralizado na caixa (folga igual).
        vAlign: top > 1.5 && Math.abs(top - low) < 2 ? 'center' : 'top',
        children: children.length ? children : [{ kind: 'p', runs: [] }],
      });
      if (gapRight >= 1) cells.push({ width: gapRight, children: [{ kind: 'p', runs: [] }], gap: true });
      continue;
    }
    // Coluna só com desenho (a bola da seta entre dois cards): alinha pelo
    // desenho, não pela caixa que o embrulha (que tem a altura toda).
    const vAlign = alignOf(col.length === 1 && !hasText(col[0]) && leaves(col).length ? union(leaves(col)) : colBox);
    // Célula na largura da coluna; o vão até a próxima vira célula vazia
    // (é onde enfeites como a seta entre dois cards voltam).
    const own = i + 1 < cols.length ? Math.max(1, Math.min(width, right(colBox) - start)) : width;
    const children = await renderRegion(col, { x: start, y: vAlign === 'center' ? colBox.y : rowTop, w: own, h: colBox.h }, ctx);
    fit(children, own);
    cells.push({ width: own, vAlign, children: children.length ? children : [{ kind: 'p', runs: [] }] });
    if (width - own >= 1) cells.push({ width: width - own, children: [{ kind: 'p', runs: [] }], gap: true });
  }
  // Enfeites tirados do corte: a imagem (ou a que está dentro deles) volta no
  // espaço vazio onde cai o centro dela.
  const images = (n: LayoutNode): LayoutNode[] => (n.k === 'img' || n.k === 'raster' ? [n] : n.k === 'box' ? n.kids.flatMap(images) : []);
  for (const o of overlays.flatMap(images)) {
    const cx = o.box.x + o.box.w / 2;
    let x = left;
    for (const cell of cells) {
      if (cx >= x && cx < x + cell.width && cell.gap) {
        const [image] = await renderItem(o, { x, y: rowTop, w: cell.width, h: rowBottom - rowTop }, ctx);
        if (image?.kind === 'p') {
          image.align = 'center';
          image.indentLeft = 0;
          image.spaceBefore = Math.max(0, o.box.y - rowTop);
          cell.children = [image];
        }
        break;
      }
      x += cell.width;
    }
  }
  // Vão entre colunas vira margem da célula vizinha sem pintura (a equipe via
  // "colunas fantasma" de poucos px). Vão com enfeite, ou entre duas células
  // pintadas, continua célula.
  const emptyGap = (c: CellIR) => c.gap && c.children.every((k) => k.kind === 'p' && !k.runs.length);
  for (let i = cells.length - 1; i >= 0; i--) {
    const c = cells[i];
    if (!emptyGap(c)) continue;
    const next = cells[i + 1];
    const prev = cells[i - 1];
    const widen = (t: CellIR, side: 1 | 3) => {
      const m = t.margins ?? [0, 0, 0, 0];
      m[side] += c.width;
      t.margins = m;
      t.width += c.width;
      cells.splice(i, 1);
    };
    if (next && !painted(next)) widen(next, 3);
    else if (prev && !painted(prev)) widen(prev, 1);
    else if (!next) cells.splice(i, 1);
  }
  const width = cells.reduce((s, c) => s + c.width, 0);
  return {
    kind: 't',
    width,
    indent: left - region.x,
    rows: badgeRows(cells, rowBottom - rowTop),
  };
};

/**
 * Selo pequeno no topo de uma célula ao lado de texto longo ("1.2" + a
 * cláusula): em vez de tabela dentro da célula, o número vira um parágrafo
 * com sombreamento (o fundo do selo) na própria célula — uma linha, sem
 * mesclagem (linha mesclada se partia entre páginas e o selo ficava órfão).
 */
const badgeRows = (cells: CellIR[], height: number): TIR['rows'] => {
  const lone = (c: CellIR): TIR | null => {
    const real = c.children.filter((k) => !(k.kind === 'p' && isEmptyP(k)));
    if (painted(c) || real.length !== 1 || real[0].kind !== 't') return null;
    const t = real[0];
    if (t.rows.length !== 1 || t.rows[0].cells.length !== 1 || !t.rows[0].cells[0].fill || t.rows[0].cells[0].backdrop) return null;
    return t;
  };
  const idx = cells.findIndex((c) => lone(c));
  if (idx < 0 || cells.filter((c) => lone(c)).length > 1) return [{ height, cells }];
  const col = cells[idx];
  const badge = lone(col)!;
  const bh = badge.rows[0].height ?? 0;
  const inner = badge.rows[0].cells[0];
  const texts = allParagraphs(inner.children).filter((p) => p.numbering || p.runs.some((r) => r.kind === 'text' && r.text.trim()));
  const lead = col.children[0];
  if (!bh || texts.length !== 1 || height - bh < 2 || (lead?.kind === 'p' && lead.spacer !== undefined && lead.spacer >= 1)) return [{ height, cells }];
  // Folga: o Word desenha o número mais largo que o Chromium ("7.1.1" cortava).
  const next = cells[idx + 1];
  const borrow = next && !painted(next) && (next.margins?.[3] ?? 0) >= 10 ? 6 : 0;
  const leftPad = (col.margins?.[3] ?? 0) + badge.indent;
  const w = badge.width + borrow;
  const number: PIR = {
    ...texts[0],
    shade: inner.fill,
    align: 'center',
    // Sombreamento ocupa o recuo interno: a largura do selo é a da célula.
    indentLeft: 0,
    indentRight: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    line: { value: tw(bh), rule: 'exact' },
  };
  const out = [...cells];
  out[idx] = { width: leftPad + w, margins: [0, 0, 0, leftPad], vAlign: 'top', children: [number] };
  // O que a célula do selo ganhou/perdeu sai/entra na margem esquerda do texto ao lado.
  const rest = col.width - (leftPad + w);
  if (next) {
    const m = next.margins ?? [0, 0, 0, 0];
    next.margins = [m[0], m[1], m[2], Math.max(0, m[3] + rest)];
    next.width += rest;
  } else out[idx].width = col.width;
  return [{ height, cells: out }];
};

const renderItem = async (n: LayoutNode, region: Box, ctx: Ctx): Promise<OutIR[]> => {
  if (complexRoot(n)) return renderComplex(n, region, ctx);
  if (n.k === 'box' && rowFlex(n) && !decorated(n)) return [await renderColumns([...n.kids].sort((a, b) => a.box.x - b.box.x).map((k) => [k]), region, ctx, [])];
  if (n.k === 'text') {
    if (decorated(n)) return [await decoratedCell(n, [n], region, ctx)];
    return textToParagraphs(n, region, ctx);
  }
  if (n.k === 'img' || n.k === 'raster') {
    const image = n.k === 'img' ? await ctx.assets.image(n.src, n.box.w, n.box.h) : await ctx.assets.raster(n.id, n.box);
    if (!image) return [];
    // SVG com proporção diferente da caixa: o desenho encolhe centralizado, como no navegador.
    let box = n.box;
    if (image.svg && image.aspect && Math.abs(image.aspect - box.w / box.h) > 0.02) {
      const w = Math.min(box.w, box.h * image.aspect);
      const h = w / image.aspect;
      box = { x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h };
    }
    const offset = box.x - region.x;
    const centered = Math.abs(offset + box.w / 2 - region.w / 2) < 2 && offset > 1;
    return [
      {
        kind: 'p',
        runs: [{ kind: 'img', image, w: box.w, h: box.h }],
        align: centered ? 'center' : 'left',
        indentLeft: centered ? 0 : Math.max(0, offset),
        spaceBefore: box.y - n.box.y > 0.5 ? box.y - n.box.y : undefined,
        // "Pelo menos": linha exata do tamanho da imagem o Word corta o topo
        // dela. Marca de parágrafo mínima: sem a folga da fonte embaixo.
        line: { value: tw(box.h), rule: 'atLeast' },
        tiny: true,
      },
    ];
  }
  if (n.k === 'shape') {
    // Forma sólida sem conteúdo (barra, quadrado): célula pintada no tamanho dela.
    return [
      {
        kind: 't',
        width: n.box.w,
        indent: Math.max(0, n.box.x - region.x),
        rows: [{ height: n.box.h, cells: [{ width: n.box.w, ...(await paintOf(n as Decorated, [0, 0, 0, 0], ctx)), children: [spacer(1)] }] }],
      },
    ];
  }
  if (decorated(n)) return [await decoratedCell(n, n.kids, region, ctx)];
  return renderRegion(n.kids, n.box, ctx);
};

/** Caixa com fundo/borda/padding → tabela de uma célula. */
const decoratedCell = async (n: Extract<LayoutNode, { k: 'box' | 'text' }>, kids: LayoutNode[], region: Box, ctx: Ctx): Promise<TIR> => {
  const pad = n.k === 'box' ? cellPadding(n.box, n.pad ?? [0, 0, 0, 0], kids) : n.pad ?? [0, 0, 0, 0];
  const inner = { x: n.box.x + pad[3], y: n.box.y + pad[0], w: n.box.w - pad[1] - pad[3], h: n.box.h - pad[0] - pad[2] };
  const children =
    n.k === 'text'
      ? textToParagraphs({ ...n, box: { ...n.box, x: inner.x, w: inner.w } }, inner, ctx)
      : await renderRegion(kids, inner, ctx);
  const layout = n.k === 'box' ? n.layout : null;
  const paint = await paintOf(n as Decorated, pad, ctx);
  // Caixa pintada cujo conteúdo é só uma linha de colunas (o selo "1" + o
  // título da cláusula): a tabela de dentro vira A tabela — 1 linha, uma
  // célula por elemento — e o fundo da caixa vai pras células sem cor própria.
  const real = children.filter((k) => !(k.kind === 'p' && isEmptyP(k)));
  const only = real.length === 1 && real[0].kind === 't' && real[0].rows.length === 1 ? real[0] : null;
  if (only && paint.fill && !paint.backdrop && !paint.borders) {
    const cells = only.rows[0].cells.map((c) => ({ ...c, fill: c.fill ?? (c.backdrop ? undefined : paint.fill) }));
    padRow(cells, pad[3] + only.indent, 'left');
    padRow(cells, n.box.w - pad[3] - only.indent - only.width, 'right');
    // Faixas das pontas que viraram célula vazia também levam o fundo.
    for (const c of cells) if (!c.fill && !c.backdrop && !c.children.length) c.fill = paint.fill;
    for (const c of cells) {
      const m = c.margins ?? [0, 0, 0, 0];
      c.margins = [m[0] + pad[0], m[1], m[2] + pad[2], m[3]];
    }
    return { kind: 't', width: n.box.w, indent: n.box.x - region.x, rows: [{ height: n.box.h, cells }] };
  }
  return {
    kind: 't',
    width: n.box.w,
    indent: n.box.x - region.x,
    rows: [
      {
        height: n.box.h,
        cells: [
          {
            width: n.box.w,
            ...paint,
            margins: [pad[0], pad[1], pad[2], pad[3]],
            vAlign: layout && (layout.alignItems === 'center' || layout.justify === 'center') ? 'center' : 'top',
            children: children.length ? children : [{ kind: 'p', runs: [] }],
          },
        ],
      },
    ],
  };
};

/**
 * Bloco com peças soltas (posição absoluta: a bola com a seta entre dois
 * cards, a linha do tempo com barra de progresso) não se fatia em tabela sem
 * deslocar. Guia da Lex pra "estrutura muito complexa": o visual vai como uma
 * imagem só, atrás, e o texto fica editável por cima, em tabela, no lugar.
 */

const renderComplex = async (tree: LayoutNode, region: Box, ctx: Ctx): Promise<OutIR[]> => {
  const texts: LayoutNode[] = [];
  const collect = (n: LayoutNode) => (n.k === 'text' ? texts.push(n) : n.k === 'box' ? n.kids.forEach(collect) : undefined);
  collect(tree);
  // Área de tudo que o bloco desenha — peça solta pode vazar da caixa (a
  // ponta da barra da linha do tempo).
  const all: LayoutNode[] = [];
  const every = (n: LayoutNode) => (all.push(n), n.k === 'box' && n.kids.forEach(every));
  every(tree);
  const u = union(all);
  // Na vertical, a altura que o fluxo usa (o conteúdo, como extent()): a caixa
  // sem pintura do bloco pode ir além e sobrepor o vizinho no Studio — no Word
  // não sobrepõe, e cada linha do Sumário crescia 8px.
  const v = union(leaves([tree]).length ? leaves([tree]) : [tree]);
  const x = Math.max(region.x - 20, Math.floor(u.x));
  const y = Math.floor(v.y);
  const area = { x, y, w: Math.ceil(right(u)) - x, h: Math.ceil(bottom(v)) - y };
  // A captura vai até onde o desenho vai (ícone de 80px num título de 22px
  // transborda pro bloco de baixo, como no Studio); a linha da tabela fica
  // com a altura do conteúdo.
  const shot = { ...area, h: Math.max(area.h, Math.ceil(bottom(u)) - y) };
  const image = await ctx.assets.backdrop(tree.id, shot);
  if (!image) return renderRegion([tree], region, ctx);
  // Sem texto por cima e cabendo na largura (ícones e barra da linha do
  // tempo): imagem na linha do texto, não flutuante — anda com o conteúdo.
  if (!texts.length && shot.x >= region.x - 8 && right(shot) <= right(region) + 8) {
    const w = Math.min(shot.w, region.w);
    const h = shot.h * (w / shot.w);
    return [{ kind: 'p', runs: [{ kind: 'img', image, w, h }], indentLeft: Math.max(0, shot.x - region.x), line: { value: tw(h), rule: 'atLeast' }, tiny: true }];
  }
  // Colunas primeiro: cada coluna de texto flui sozinha — fatiar por linhas
  // amarrava a altura de uma coluna na outra e desalinhava do desenho.
  // A tabela fica no lugar do fluxo; o desenho que vaza pros lados (a bola
  // entre dois cards, numa coluna de 24px) vai deslocado na âncora — tabela
  // não passa da célula, imagem flutuante passa.
  const tx = Math.max(area.x, region.x);
  const box = { x: tx, y: area.y, w: Math.max(1, Math.min(right(area), right(region)) - tx), h: area.h };
  const cols = bands(texts, 'x');
  const children = !texts.length ? [] : cols.length > 1 ? withGap([await renderColumns(cols, box, ctx, [])], union(texts).y - box.y) : await renderRegion(texts, box, ctx);
  return [
    {
      kind: 't',
      width: box.w,
      indent: box.x - region.x,
      rows: [{ height: box.h, cells: [{ width: box.w, backdrop: { image, w: shot.w, h: shot.h, x: area.x - box.x, front: !texts.length, center: !texts.length }, children: children.length ? children : [{ kind: 'p', runs: [] }] }] }],
    },
  ];
};

// ---------------------------------------------------------------------------
// IR → docx
// ---------------------------------------------------------------------------

const runProps = (run: Run, face: Face | null) => ({
  font: face?.family ?? run.font,
  bold: face ? face.bold : run.weight >= 600,
  italics: face ? face.italic : run.italic,
  size: halfPt(run.size),
  color: run.color,
  allCaps: run.caps || undefined,
  underline: run.underline ? {} : undefined,
  strike: run.strike || undefined,
  characterSpacing: run.letterSpacing ? Math.round(run.letterSpacing * 15) : undefined,
  superScript: run.verticalAlign === 'super' || undefined,
  subScript: run.verticalAlign === 'sub' || undefined,
});

/**
 * Indicador oculto ("_" na frente: fica fora da lista "Indicadores" do Word),
 * só no título do bloco que é destino de algum link.
 */
const bookmarkName = (blockId: string) => `_Toc${blockId.replace(/-/g, '').slice(0, 30)}`;
/** Blocos que ganharam indicador — link pra qualquer outro vira texto comum (sem link quebrado). */
let linkTargets = new Set<string>();

const toRuns = (p: PIR): ParagraphChild[] =>
  p.runs.flatMap((r): ParagraphChild[] => {
    if (r.kind === 'br') return [new TextRun({ break: 1 })];
    if (r.kind === 'img') {
      const size = { width: Math.round(r.w), height: Math.round(r.h) };
      return [
        r.image.svg
          ? new ImageRun({ type: 'svg', data: r.image.svg, transformation: size, fallback: { type: r.image.type, data: r.image.data } } as never)
          : new ImageRun({ type: r.image.type, data: r.image.data, transformation: size } as never),
      ];
    }
    const text = r.run.lower ? r.text.toLowerCase() : r.text;
    const run = new TextRun({ text, ...runProps(r.run, r.face), ...(r.spacing !== undefined ? { characterSpacing: r.spacing } : {}) });
    if (r.run.link) {
      if (/^https?:|^mailto:/.test(r.run.link)) return [new ExternalHyperlink({ link: r.run.link, children: [run] })];
      if (!linkTargets.has(r.run.link)) return [run];
      return [new InternalHyperlink({ anchor: bookmarkName(r.run.link), children: [run] })];
    }
    return [run];
  });

/** Altura em px da linha Simples de uma fonte de 1 meio-ponto no corpo (medida da fonte do Normal). */
let simpleRatio = 1.2;
/** Parágrafo sem texto: meio-ponto da fonte cuja linha Simples dá a altura pedida. */
const structural = (p: PIR): number | undefined => {
  if (p.runs.some((r) => r.kind === 'text') || p.numbering) return undefined;
  if (p.tiny) return 2;
  if (p.grow) return p.grow;
  if (!p.line || p.line.rule === 'auto') return undefined;
  const px = p.line.value / 15;
  return Math.max(2, Math.min(144, Math.round((px / simpleRatio) * 1.5)));
};

const toParagraph = (p: PIR): Paragraph => {
  if (p.blank) return new Paragraph({ keepNext: p.keepNext || undefined, children: p.bookmark ? [new Bookmark({ id: p.bookmark, children: [] })] : [] });
  const children = toRuns(p);
  const body = p.bookmark ? [new Bookmark({ id: p.bookmark, children })] : children;
  return new Paragraph({
    children: p.pageBreakBefore ? [new PageBreak(), ...body] : body,
    alignment:
      p.align === 'center' ? AlignmentType.CENTER : p.align === 'right' ? AlignmentType.RIGHT : p.align === 'both' ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    indent:
      p.indentLeft || p.indentRight || p.hanging !== undefined
        ? { left: Math.round((p.indentLeft ?? 0) * 15), right: tw(p.indentRight ?? 0), ...(p.hanging !== undefined ? { hanging: tw(p.hanging) } : {}) }
        : undefined,
    spacing: {
      before: tw(p.spaceBefore ?? 0),
      after: tw(p.spaceAfter ?? 0),
      // Sempre automática (Simples/Múltiplos) — nunca fixa nem "pelo menos".
      line: p.line?.rule === 'auto' ? p.line.value : 240,
      lineRule: LineRuleType.AUTO,
    },
    numbering: p.numbering ? { reference: p.numbering.ref, level: p.numbering.level, ...(p.numbering.instance ? { instance: p.numbering.instance } : {}) } : undefined,
    border:
      p.borderBottom || p.borderTop
        ? {
            ...(p.borderBottom ? { bottom: { style: BorderStyle.SINGLE, size: Math.max(2, Math.round(p.borderBottom.w * 6)), color: p.borderBottom.color, space: 0 } } : {}),
            ...(p.borderTop ? { top: { style: BorderStyle.SINGLE, size: Math.max(2, Math.round(p.borderTop.w * 6)), color: p.borderTop.color, space: 0 } } : {}),
          }
        : undefined,
    // Espaçador herda a fonte do corpo (Normal) — com a altura exata, não
    // aparece fonte 1 no meio do texto; só o parágrafo de imagem é mínimo.
    // Parágrafo sem texto (espaço, âncora de imagem) em Simples: a altura vem
    // do tamanho da fonte da marca — do tamanho do espaço do Studio.
    run: p.mark ? runProps(p.mark.run, p.mark.face) : structural(p) !== undefined ? { size: structural(p) } : undefined,
    heading: p.heading ? [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][p.heading - 1] : undefined,
    keepNext: p.keepNext || undefined,
    shading: p.shade ? { fill: p.shade, color: 'auto', type: 'clear' as never } : undefined,
  });
};

/** Parágrafo mínimo no topo da célula segurando a arte de fundo (atrás do texto, dentro da célula). */
/**
 * `centered`: a célula é centralizada na vertical e a âncora é o único
 * parágrafo dela — fica no meio da célula, e a arte (metade pra cima, metade
 * pra baixo) acompanha quando a linha cresce (a bola com a seta entre dois
 * cards continua no meio deles, como no Studio).
 */
const backdropParagraph = (b: NonNullable<CellIR['backdrop']>, keepNext: boolean, centered = false): Paragraph =>
  new Paragraph({
    keepNext: keepNext || undefined,
    spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
    run: { size: 2 },
    children: [
      new ImageRun({
        ...(b.image.svg ? { type: 'svg', data: b.image.svg, fallback: { type: b.image.type, data: b.image.data } } : { type: b.image.type, data: b.image.data }),
        transformation: { width: Math.round(b.w), height: Math.round(b.h) },
        floating: {
          horizontalPosition: { relative: HorizontalPositionRelativeFrom.COLUMN, offset: emu(b.x ?? 0) },
          verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: emu(centered ? -b.h / 2 + 1 : b.y ?? 0) },
          // Sem texto por cima (a bola entre dois cards): na frente — atrás, o
          // sombreamento das células vizinhas cobria.
          behindDocument: !b.front,
          allowOverlap: true,
          layoutInCell: true,
          lockAnchor: true,
        },
      } as never),
    ],
  });

const border = (b?: { w: number; color: string }): IBorderOptions =>
  b ? { style: BorderStyle.SINGLE, size: Math.max(2, Math.round(b.w * 6)), color: b.color } : { style: BorderStyle.NONE, size: 0, color: 'auto' };

const toBlocks = (items: OutIR[], body = false): (Paragraph | Table)[] => {
  const out: (Paragraph | Table)[] = [];
  items.forEach((item, i) => {
    if (item.kind === 'p') {
      out.push(toParagraph(item));
      return;
    }
    // Duas tabelas coladas o Word funde numa só: separa com parágrafo mínimo
    // — que segue o "manter com o próximo" da tabela de cima (título do card
    // não fica sozinho no pé da página).
    const prev = items[i - 1];
    if (prev?.kind === 't') {
      const sep: PIR = body ? { kind: 'p', runs: [], blank: true } : spacer(0.5);
      sep.keepNext = prev.rows[prev.rows.length - 1].cells.some((c) => allParagraphs(c.children).some((p) => p.keepNext));
      out.push(toParagraph(sep));
    }
    out.push(toTable(item));
  });
  return out;
};

const toTable = (t: TIR): Table => {
  // Grade comum a todas as linhas: linhas com células diferentes (título do
  // card + corpo com ícone) na mesma tabela, cada célula ocupando as colunas
  // da grade que cobre.
  const edges = new Set<number>([0]);
  for (const row of t.rows) {
    let x = 0;
    for (const c of row.cells) edges.add((x += tw(c.width)));
  }
  const grid = [...edges].sort((a, b) => a - b);
  const widths = grid.slice(1).map((e, i) => e - grid[i]);
  const spanOf = (from: number, to: number) => grid.filter((e) => e > from && e <= to).length;
  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    indent: t.indent ? { size: Math.round(t.indent * 15), type: WidthType.DXA } : undefined,
    borders: {
      top: border(), bottom: border(), left: border(), right: border(), insideHorizontal: border(), insideVertical: border(),
    },
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    rows: t.rows.map(
      (row) =>
        new TableRow({
          cantSplit: true,
          height: row.height ? { value: tw(Math.max(1, row.height)), rule: HeightRule.ATLEAST } : undefined,
          children: row.cells.map((cell, ci) => {
            const from = row.cells.slice(0, ci).reduce((a, c) => a + tw(c.width), 0);
            const span = spanOf(from, from + tw(cell.width));
            // Padding de cima/baixo vira espaçamento do parágrafo, não margem da
            // célula: o Word aplica a MAIOR margem vertical da linha a todas as
            // células (medido: ícone ao lado de card com 20px de padding ganhava
            // 40px) — com espaçamento, cada célula tem o seu, e a grade da
            // tabela pode ser uma só.
            // Célula centralizada: o Word já centraliza — o espaço de cima que
            // posicionava o conteúdo no Studio deslocava de novo (logo baixo).
            // Com arte de fundo, não: a arte é ancorada no primeiro parágrafo e
            // a centralização a levaria junto — posições do Studio, pelo topo.
            const artCentered = Boolean(cell.backdrop?.center) && cell.vAlign === 'center';
            const centered = cell.vAlign === 'center' && (!cell.backdrop || artCentered);
            const [mt, , mb] = centered ? [0, 0, 0] : cell.margins ?? [0, 0, 0, 0];
            let items = cell.children;
            if (centered) {
              items = items.filter((k, i) => !(k.kind === 'p' && k.spacer !== undefined && (i === 0 || i === items.length - 1) && !k.bookmark));
              // Só nas pontas da célula: o espaço entre uma tabela e o texto
              // de baixo (calendário → "Comunicação em até…") é conteúdo.
              const first = items[0];
              if (first?.kind === 'p') first.spaceBefore = 0;
              const last = items[items.length - 1];
              if (last?.kind === 'p') last.spaceAfter = 0;
            }
            if (mt >= 0.5) items = [spacer(mt), ...items];
            if (mb >= 0.5) {
              const last = items[items.length - 1];
              if (last?.kind === 'p' && !last.spacer && !last.blank) last.spaceAfter = (last.spaceAfter ?? 0) + mb;
              else items = [...items, spacer(mb)];
            }
            const children = toBlocks(items);
            // A âncora segue o "manter com o próximo" da célula (o Word olha todos os parágrafos da linha).
            const keep = allParagraphs(cell.children).some((p) => p.keepNext);
            // Arte centralizada sem texto: a âncora é o único parágrafo (senão o
            // parágrafo vazio de baixo deslocaria o centro).
            if (artCentered && items.every((k) => k.kind === 'p' && !k.runs.length && !k.bookmark)) children.splice(0, children.length);
            if (cell.backdrop) children.unshift(backdropParagraph(cell.backdrop, keep, artCentered));
            // Célula tem que terminar em parágrafo.
            if (!children.length || children[children.length - 1] instanceof Table) children.push(toParagraph(spacer(0.5)));
            return new TableCell({
              width: { size: tw(cell.width), type: WidthType.DXA },
              ...(span > 1 ? { columnSpan: span } : {}),
              ...(cell.vmerge ? { verticalMerge: cell.vmerge === 'restart' ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE } : {}),
              shading: cell.fill ? { fill: cell.fill, color: 'auto', type: 'clear' as never } : undefined,
              margins: cell.margins
                ? { top: 0, right: tw(cell.margins[1]), bottom: 0, left: tw(cell.margins[3]), marginUnitType: WidthType.DXA }
                : undefined,
              verticalAlign: cell.vAlign === 'center' && (!cell.backdrop || (cell.backdrop.center && cell.vAlign === 'center')) ? VerticalAlign.CENTER : cell.vAlign === 'bottom' ? VerticalAlign.BOTTOM : VerticalAlign.TOP,
              borders: {
                top: border(cell.borders?.top), bottom: border(cell.borders?.bottom), left: border(cell.borders?.left), right: border(cell.borders?.right),
              },
              children,
            });
          }),
        }),
    ),
  });
};

// ---------------------------------------------------------------------------
// Rodapé
// ---------------------------------------------------------------------------

/** Marcador trocado depois (index.ts) pelos campos PAGE/NUMPAGES do Word. */
export const PAGE_FIELD = '§LEXPAGE§';

/**
 * Rodapé do Studio (texto + "02/28") → rodapé do Word: tabela sem borda com
 * o texto e a numeração como campo (atualiza sozinha quando o conteúdo cresce).
 */
const buildFooter = (layout: DocumentLayout, pageIndex: number, pad: { left: number; right: number }, ctx: Ctx) => {
  const page = layout.pages[pageIndex];
  const footer = page?.footer;
  // Sem rodapé: parágrafo mínimo — um vazio no estilo Normal ocupava uma linha no pé da capa.
  if (!footer?.number) return { footer: new Footer({ children: [new Paragraph({ spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO }, run: { size: 2 }, children: [] })] }), distance: 0, offsets: null };
  const num = footer.number;
  const [current, total] = num.text.split('/').map((v) => Number(v));
  const offsets = { current: current - (pageIndex + 1), total: total - layout.pages.length };
  const numFace = ctx.fonts.resolve(num.style.font, num.style.weight, num.style.italic);
  if (numFace) ctx.used.add(numFace);
  const textW = Math.max(1, num.box.x - pad.left);
  const region = { x: pad.left, y: 0, w: textW, h: 0 };
  const text = footer.text ? textToParagraphs({ ...footer.text, box: { ...footer.text.box, x: pad.left } }, region, ctx) : [];
  const numberParagraph: PIR = {
    kind: 'p',
    align: num.align === 'right' ? 'right' : num.align === 'center' ? 'center' : 'left',
    runs: [{ kind: 'text', text: PAGE_FIELD, run: { ...num.style, t: PAGE_FIELD, link: null }, face: numFace }],
  };
  const table: TIR = {
    kind: 't',
    width: textW + num.box.w,
    indent: 0,
    rows: [{ cells: [
      { width: textW, vAlign: 'bottom', children: text.length ? text : [{ kind: 'p', runs: [] }] },
      { width: num.box.w, vAlign: 'bottom', children: [numberParagraph] },
    ] }],
  };
  const lowest = Math.max(bottom(num.box), footer.text ? bottom(footer.text.box) : 0);
  return {
    footer: new Footer({ children: [...toBlocks([table]), toParagraph(spacer(0.5))] }),
    distance: Math.max(0, page.box.y + page.box.h - lowest),
    offsets,
  };
};

// ---------------------------------------------------------------------------
// Documento
// ---------------------------------------------------------------------------

export type BuildInput = {
  layout: DocumentLayout;
  fonts: FontResolver;
  assets: Assets;
  /** Imagem de cada página inteira que vira fundo no cabeçalho (capa inclusive). */
  background: (pageIndex: number) => Promise<ImageData | null>;
  /** Páginas cujo conteúdo inteiro vai no cabeçalho (capa: não editável). */
  coverPages: Set<number>;
};

export const buildDocument = async ({ layout, fonts, assets, background, coverPages }: BuildInput) => {
  const ctx: Ctx = { fonts, assets, used: new Set(), pending: [], usage: new Map(), lists: new Map(), seq: null, instances: 0, clauses: [] };
  const { meta } = layout;
  const pageW = meta.pageWidthPx;
  const pageH = meta.pageHeightPx;
  const pad = {
    top: meta.paddingTopPx ?? 0,
    right: meta.paddingRightPx ?? 0,
    bottom: meta.paddingBottomPx ?? 0,
    left: meta.paddingLeftPx ?? 0,
  };

  // Seção nova quando o cabeçalho muda: fundo (capa, sumário, abertura de
  // capítulo — com quebra de página) ou aba ativa do menu (guia da Lex: uma
  // seção por tópico do menu, contínua).
  const bgKey = (i: number) => (coverPages.has(i) ? `cover:${i}` : layout.pages[i]?.bgImage ?? 'none');
  const groups: { pages: number[]; key: string; bg: string }[] = [];
  for (const p of layout.pages) {
    const bg = bgKey(p.index);
    const key = `${bg}#${coverPages.has(p.index) ? '' : p.nav?.sig ?? ''}`;
    if (groups.length && groups[groups.length - 1].key === key && !coverPages.has(p.index)) groups[groups.length - 1].pages.push(p.index);
    else groups.push({ key, bg, pages: [p.index] });
  }

  // Área que o bloco ocupa de verdade — inclusive o que transborda a caixa
  // dele (o número grande do Sumário sobe 9px além do bloco).
  const extent = (tree: LayoutNode): Box => union(leaves([tree]).length ? leaves([tree]) : [tree]);
  const firstParagraph = (items: OutIR[]): PIR | null => {
    for (const item of items) {
      if (item.kind === 'p') return item;
      for (const cell of item.rows[0]?.cells ?? []) {
        const found = firstParagraph(cell.children);
        if (found) return found;
      }
    }
    return null;
  };

  /**
   * Bloco de título (h1–h3 do Studio, ou selo de capítulo com número de
   * nível 1): o texto vira Título 1/2/3 do Word — navegação e sumário
   * funcionam — e o bloco inteiro fica com o próximo.
   */
  const markHeadings = (block: (typeof layout.blocks)[number], placed: OutIR[]) => {
    const paragraphs = allParagraphs(placed);
    const chapter = paragraphs.some((p) => p.numbering?.ref === 'clauses' && p.numbering.level === 0);
    const native = { h1: 1, h2: 2, h3: 3 }[block.type] as 1 | 2 | 3 | undefined;
    if (!chapter && !native) return;
    for (const p of paragraphs) p.keepNext = true;
    const title = paragraphs.find((p) => !p.numbering && p.runs.some((r) => r.kind === 'text' && r.text.trim()));
    if (title) title.heading = native ?? 1;
  };

  // Espaço típico entre blocos (mediana dos vãos dentro de uma mesma página):
  // é o que vale quando o Studio mudou de página e o Word continua fluindo.
  const sameGaps: number[] = [];
  for (let i = 1; i < layout.blocks.length; i++) {
    const [a, b] = [layout.blocks[i - 1], layout.blocks[i]];
    if (a.tree && b.tree && a.page === b.page) sameGaps.push(extent(b.tree).y - bottom(extent(a.tree)));
  }
  sameGaps.sort((a, b) => a - b);
  const typicalGap = Math.max(0, sameGaps[Math.floor(sameGaps.length / 2)] ?? 0);

  // Corpo de texto = fonte/tamanho/entrelinha mais usados no documento: é o
  // estilo Normal e a linha em branco entre tópicos.
  const tally = new Map<string, number>();
  const scan = (n: LayoutNode | null): void => {
    if (!n) return;
    if (n.k === 'box') return n.kids.forEach(scan);
    if (n.k !== 'text') return;
    for (const para of n.paras) {
      for (const r of para.runs) {
        if (!r.t) continue;
        const face = fonts.resolve(r.font, r.weight, r.italic);
        const key = `${face?.family ?? r.font}|${halfPt(r.size)}|${para.lineHeight ?? r.size * 1.2}|${face ? '1' : ''}`;
        tally.set(key, (tally.get(key) ?? 0) + r.t.length);
      }
    }
  };
  layout.blocks.forEach((b) => scan(b.tree));
  const [bodyFamily, bodySizeHalf, bodyLinePx] = ([...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 'Arial|20|18').split('|');
  const bodyLine = Number(bodyLinePx) || 18;
  simpleRatio = fonts.resolve(bodyFamily, 400, false)?.lineHeight ?? 1.2;

  type SectionIR = {
    gi: number;
    body: OutIR[];
    nav: OutIR[];
    headerBg: Paragraph;
    headerLogos: Paragraph[];
    foot: ReturnType<typeof buildFooter>;
    marginTop: number;
    headerDistance: number;
  };
  const sectionsIR: SectionIR[] = [];
  /** Logos tirados do corpo por seção: vão pro cabeçalho, na mesma posição da página. */
  let headerImages: { image: ImageData; box: Box }[] = [];
  /** Capa: margem de baixo até o último texto (senão o "São Paulo/SP" do pé pula de página). */
  const coverBottom = new Map<number, number>();
  /** Bloco → onde vai o indicador dele se for destino de link (o título, senão o primeiro parágrafo). */
  const targets = new Map<string, PIR>();
  let pageOffsets: { current: number; total: number } | null = null;
  // Onde o conteúdo anterior terminou. Atravessa a troca de seção contínua (só
  // a aba do menu mudou): o vão é o do Studio, não a distância do topo da página.
  let cursor: { page: number; y: number } | null = null;
  let previous: OutIR[] | null = null;
  // Espaço entre tópicos em linhas em branco do corpo de texto (guia da Lex:
  // "manter o espaçamento definido no corpo"), não em px quebrados do Figma.
  const blanksFor = (gap: number) => (gap < bodyLine * 0.6 ? 0 : Math.max(1, Math.round(gap / bodyLine)));
  for (const [gi, group] of groups.entries()) {
    const pageTop = (i: number) => layout.pages[i].box.y;
    const blocks = layout.blocks.filter((b) => group.pages.includes(b.page) && b.tree);
    const body: OutIR[] = [];
    // Margem de cima da seção: onde o conteúdo começa na primeira página dela
    // (a página do Sumário encosta no topo; as de texto respeitam o padding).
    const firstTop = blocks.length ? extent(blocks[0].tree!).y - pageTop(blocks[0].page) : pad.top;
    // Menu no cabeçalho: o corpo começa abaixo dele (senão o Word empurra o
    // corpo pra baixo do cabeçalho E ainda soma o vão até o primeiro bloco).
    const nav = group.bg.startsWith('cover:') ? null : layout.pages[group.pages[0]]?.nav;
    const navBottom = nav ? bottom(nav.tree.box) - pageTop(group.pages[0]) : 0;
    const marginTop = Math.max(0, Math.min(pad.top, firstTop), Math.min(firstTop, navBottom));
    if (gi > 0 && groups[gi - 1].bg !== group.bg) {
      cursor = null;
      previous = null;
    }
    if (group.bg.startsWith('cover:')) {
      // Capa: a arte está no cabeçalho; aqui só os textos, no lugar deles.
      const texts: LayoutNode[] = [];
      const collect = (n: LayoutNode | null): void => {
        if (!n) return;
        if (n.k === 'text') texts.push(n);
        else if (n.k === 'box') n.kids.forEach(collect);
      };
      blocks.forEach((b) => collect(b.tree));
      if (texts.length) {
        const top = pageTop(group.pages[0]);
        body.push(...(await renderRegion(texts, { x: pad.left, y: top + marginTop, w: pageW - pad.left - pad.right, h: pageH }, ctx)));
        // Folga pro parágrafo da quebra de página caber ainda na capa.
        coverBottom.set(gi, Math.max(0, top + pageH - bottom(union(texts)) - 8));
      }
    } else {
      const newBackground = gi === 0 || groups[gi - 1].bg !== group.bg;
      for (let bi = 0; bi < blocks.length; bi++) {
        const block = blocks[bi];
        const page = layout.pages[block.page];
        // Fundo branco sobre página branca não é fundo (vira texto solto, não
        // tabela). Página com imagem de fundo: vale a cor da imagem ali embaixo.
        // Só página de cor lisa: fundo em imagem vai como foi salvo, sem adivinhar cor.
        const pageColor = page?.bgImage ? null : page?.bgColor ?? 'ffffff';
        let tree = nestShapes(pageColor ? stripSame(block.tree!, pageColor) : block.tree!);
        // Logo solto no topo da primeira página da seção (Sumário, Quadro
        // resumo): vai pro cabeçalho; o título fica como texto na página.
        if (bi === 0 && newBackground) {
          const top = pageTop(block.page);
          const logos = looseImages(tree).filter((l) => l.box.y - top < pageH * 0.2 && l.box.w <= pageW * 0.35 && l.box.h <= pageH * 0.15);
          for (const logo of logos) {
            const image = logo.k === 'img' ? await ctx.assets.image(logo.src, logo.box.w, logo.box.h) : await ctx.assets.raster(logo.id, logo.box);
            if (image) headerImages.push({ image, box: { ...logo.box, x: logo.box.x - page.box.x, y: logo.box.y - top } });
          }
          if (logos.length) tree = without(tree, new Set(logos));
        }
        // Ícone do título que desce ao lado do texto seguinte (o globo da
        // "Transferência internacional"): título + texto seguinte numa coluna,
        // o ícone na outra — a imagem fica numa célula e acompanha o texto.
        const next = blocks[bi + 1];
        const hanging = next?.tree && next.page === block.page ? looseImages(tree)[0] : undefined;
        let joined: typeof next | undefined;
        if (hanging) {
          const rest = without(tree, new Set([hanging]));
          const own = extent(rest);
          const nextArea = extent(next!.tree!);
          // Direita do conteúdo de verdade (textos, imagens) — a caixa pode ter a largura toda (linha embaixo).
          const content = everyNode(next!.tree!).filter((n) => n.k === 'text' || n.k === 'img' || n.k === 'raster');
          const nextRight = content.length ? Math.max(...content.map((n) => (n.k === 'text' && n.inkRight ? n.inkRight : right(n.box)))) : right(nextArea);
            if (bottom(hanging.box) > bottom(own) + 16 && nextArea.y < bottom(hanging.box) && nextRight <= hanging.box.x + 1) {
            tree = { ...(rest as Extract<LayoutNode, { k: 'box' }>), k: 'box', kids: [rest, next!.tree!, hanging], layout: { display: 'block', dir: 'row', alignItems: 'normal', justify: 'normal', textAlign: 'left' } } as LayoutNode;
            joined = next;
            bi++;
          }
        }
        const area = extent(tree);
        const region = { x: pad.left, y: area.y, w: pageW - pad.left - pad.right, h: area.h };
        const out = mergeTables(
          joined
            ? [await renderColumns([[(tree as Extract<LayoutNode, { k: 'box' }>).kids[0], (tree as Extract<LayoutNode, { k: 'box' }>).kids[1]], [(tree as Extract<LayoutNode, { k: 'box' }>).kids[2]]], region, ctx, [])]
            : await renderRegion([tree], region, ctx),
        );
        let placed: OutIR[];
        if (!cursor) {
          // Primeiro bloco da página: a distância do topo continua exata.
          placed = withGap(out, Math.max(0, area.y - (pageTop(block.page) + marginTop)));
        } else {
          // Mudou de página no Studio: no Word o texto continua — separa como tópico (uma linha do corpo, no mínimo).
          const gap = cursor.page === block.page ? area.y - cursor.y : Math.max(typicalGap, bodyLine);
          const lastOfPrevious = previous ? allParagraphs(previous).pop() : undefined;
          const n = blanksFor(gap);
          const blanks: PIR[] = Array.from({ length: n }, () => ({ kind: 'p', runs: [], blank: true, keepNext: lastOfPrevious?.keepNext }));
          // Vão pequeno (linhas de um quadro, 8px): não é espaço entre tópicos — fica exato.
          placed = n ? [...blanks, ...out] : withGap(out, Math.max(0, gap));
        }
        markHeadings(block, placed);
        // No Studio um bloco não se parte entre páginas (vai inteiro pra
        // próxima): "manter com o próximo" em tudo menos no fim do bloco.
        // Parágrafo com "manter" numa linha de tabela prende a linha ao que
        // vem depois — a última linha do bloco fica sem, senão os blocos se
        // encadeiam.
        if (area.h < pageH * 0.6) {
          placed.forEach((item, i) => {
            const last = i === placed.length - 1;
            if (item.kind === 'p') {
              if (!last) item.keepNext = true;
              return;
            }
            item.rows.forEach((row, ri) => {
              if (!last || ri < item.rows.length - 1) row.cells.forEach((c) => allParagraphs(c.children).forEach((p) => (p.keepNext = true)));
            });
          });
        }
        const ps = allParagraphs(out);
        const target = ps.find((p) => p.heading) ?? ps.find((p) => !p.blank && !p.spacer) ?? ps[0];
        if (target) targets.set(block.blockId, target);
        if (joined) {
          const first = leaves([joined.tree!]).find((l) => l.k === 'text') as Extract<LayoutNode, { k: 'text' }> | undefined;
          const text = first?.paras[0]?.runs.find((r) => r.t)?.t;
          const p = text ? ps.find((q) => q.runs.some((r) => r.kind === 'text' && r.text === text)) : undefined;
          if (p) targets.set(joined.blockId, p);
        }
        body.push(...placed);
        cursor = { page: (joined ?? block).page, y: bottom(area) };
        previous = out;
      }
    }
    if (!body.length) body.push({ kind: 'p', runs: [] });
    // Guia da Lex: quebra de página + seção contínua (nunca seção "próxima
    // página") quando o fundo muda; só a aba do menu mudou → seção contínua.
    // Depois da capa a próxima seção já começa em página nova (abaixo): a quebra
    // aqui, com a capa cheia até o pé, caía na página seguinte e deixava uma em branco.
    if (gi < groups.length - 1 && groups[gi + 1].bg !== group.bg && !group.bg.startsWith('cover:')) {
      body.push({ kind: 'p', runs: [], pageBreakBefore: true, line: { value: MIN_LINE, rule: 'exact' } });
    }

    const logosHere = headerImages;
    headerImages = [];
    const bg = await background(group.pages[0]);
    const headerBg = bg
      ? new Paragraph({
          // Só segura a âncora do fundo: não pode ocupar altura no cabeçalho.
          spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
          run: { size: 2 },
          children: [
            new ImageRun({
              ...(bg.svg ? { type: 'svg', data: bg.svg, fallback: { type: bg.type, data: bg.data } } : { type: bg.type, data: bg.data }),
              transformation: { width: pageW, height: pageH },
              floating: {
                horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
                verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
                behindDocument: true,
                allowOverlap: true,
                lockAnchor: true,
              },
            } as never),
          ],
        })
      : new Paragraph({ spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO }, run: { size: 2 }, children: [] });

    // Menu de seções no cabeçalho, como tabela, com links pros capítulos.
    let headerDistance = 0;
    let navIR: OutIR[] = [];
    if (nav) {
      headerDistance = Math.max(0, nav.tree.box.y - pageTop(group.pages[0]));
      navIR = await renderRegion([nav.tree], { x: pad.left, y: nav.tree.box.y, w: pageW - pad.left - pad.right, h: nav.tree.box.h }, ctx);
    }

    const withFooter = group.pages.find((i) => layout.pages[i]?.footer?.number);
    const foot = withFooter !== undefined ? buildFooter(layout, withFooter, pad, ctx) : buildFooter(layout, -1, pad, ctx);
    pageOffsets ??= foot.offsets;
    const headerLogos = logosHere.map(({ image, box }) =>
      new Paragraph({
        spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
        run: { size: 2 },
        children: [
          new ImageRun({
            ...(image.svg ? { type: 'svg', data: image.svg, fallback: { type: image.type, data: image.data } } : { type: image.type, data: image.data }),
            transformation: { width: Math.round(box.w), height: Math.round(box.h) },
            floating: {
              horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: Math.round(box.x * 9525) },
              verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: Math.round(box.y * 9525) },
              // Na frente: atrás do texto, o Word desenhava o logo embaixo do fundo do cabeçalho.
              behindDocument: false,
              allowOverlap: true,
              lockAnchor: true,
            },
          } as never),
        ],
      }),
    );
    sectionsIR.push({ gi, body, nav: navIR, headerBg, headerLogos, foot, marginTop, headerDistance });
  }

  // Indicador só onde algum link aponta (menu, sumário, links do texto).
  const referenced = new Set<string>();
  for (const sec of sectionsIR) {
    for (const p of allParagraphs([...sec.body, ...sec.nav])) {
      for (const r of p.runs) if (r.kind === 'text' && r.run.link && !/^https?:|^mailto:/.test(r.run.link)) referenced.add(r.run.link);
    }
  }
  linkTargets = new Set();
  for (const id of referenced) {
    const p = targets.get(id);
    if (!p) continue;
    p.bookmark = bookmarkName(id);
    linkTargets.add(id);
  }

  const sections: ISectionOptions[] = sectionsIR.map((sec) => {
    const headerChildren: (Paragraph | Table)[] = [sec.headerBg, ...sec.headerLogos];
    if (sec.nav.length) {
      const navBlocks = toBlocks(sec.nav);
      headerChildren.push(...navBlocks);
      if (navBlocks[navBlocks.length - 1] instanceof Table) headerChildren.push(toParagraph(spacer(0.5)));
    }
    return {
      properties: {
        type: sec.gi === 0 || groups[sec.gi - 1].bg.startsWith('cover:') ? SectionType.NEXT_PAGE : SectionType.CONTINUOUS,
        page: {
          size: { width: tw(pageW), height: tw(pageH) },
          margin: {
            top: tw(sec.marginTop),
            right: tw(pad.right),
            bottom: tw(Math.min(pad.bottom, coverBottom.get(sec.gi) ?? pad.bottom)),
            left: tw(pad.left),
            header: tw(sec.headerDistance),
            footer: tw(sec.foot.distance),
          },
        },
      },
      headers: { default: new Header({ children: headerChildren }) },
      footers: { default: sec.foot.footer },
      children: toBlocks(separate(joinTight(sec.body)), true),
    };
  });

  const body = { font: bodyFamily, size: Number(bodySizeHalf) };
  // Normal com a entrelinha do corpo ("pelo menos": cresce se aumentarem a fonte).
  const bodyFace = fonts.resolve(bodyFamily, 400, false);
  const normalLine = lineOf({ lineHeight: bodyLine, fontSize: Number(bodySizeHalf) / 1.5, align: 'left', runs: [] }, bodyFace)!;
  const doc = new Document({
    features: { updateFields: false },
    numbering: {
      config: [
        {
          reference: 'clauses',
          levels: [0, 1, 2].map((level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: ['%1', '%1.%2', '%1.%2.%3'][level],
            suffix: LevelSuffix.NOTHING,
            alignment: AlignmentType.LEFT,
            // Sem recuo: o estilo "Parágrafo da Lista" põe 1,27 cm, que
            // empurrava o número pra fora do selo (cortava no Windows).
            style: { paragraph: { indent: { left: 0, hanging: 0 } } },
          })),
        },
        ...[...ctx.lists.entries()].map(([key, l]) => ({
          reference: `list:${key}`,
          levels: [{
            level: 0,
            format: { lowerRoman: LevelFormat.LOWER_ROMAN, upperRoman: LevelFormat.UPPER_ROMAN, lowerLetter: LevelFormat.LOWER_LETTER, upperLetter: LevelFormat.UPPER_LETTER, decimal: LevelFormat.DECIMAL }[l.format],
            text: `${l.before}%1${l.after}`,
            suffix: l.suffix === 'space' ? LevelSuffix.SPACE : LevelSuffix.NOTHING,
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 0, hanging: 0 } } },
          }],
        })),
        {
          reference: 'bullet',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: LevelFormat.BULLET,
            text: level % 2 ? '◦' : '•',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360 * (level + 1), hanging: 240 } } },
          })),
        },
        {
          reference: 'decimal',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: level === 1 ? LevelFormat.LOWER_LETTER : LevelFormat.DECIMAL,
            text: level === 1 ? `%${level + 1})` : `%${level + 1}.`,
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360 * (level + 1), hanging: 300 } } },
          })),
        },
      ],
    },
    styles: {
      default: {
        // Normal = fonte/tamanho mais usados no documento: texto novo digitado no Word já sai certo.
        document: {
          run: { font: body.font, size: body.size },
          paragraph: { spacing: { after: 0, line: normalLine.value, lineRule: LineRuleType.AUTO } },
        },
        // Títulos só marcam a estrutura (navegação, sumário); a aparência vem dos próprios trechos.
        heading1: { run: { font: body.font, size: body.size, color: undefined }, paragraph: { spacing: { before: 0, after: 0 } } },
        heading2: { run: { font: body.font, size: body.size }, paragraph: { spacing: { before: 0, after: 0 } } },
        heading3: { run: { font: body.font, size: body.size }, paragraph: { spacing: { before: 0, after: 0 } } },
      },
    },
    sections,
  });
  return { doc, used: ctx.used, pageOffsets };
};
