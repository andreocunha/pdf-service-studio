/**
 * Devolve ao .docx do iLovePDF as fontes reais do documento.
 *
 * O iLovePDF reconstrói a estrutura do Word com muita fidelidade, mas não
 * reconhece as fontes embutidas no PDF (subsets "ABCDEF+CAIXAStd-SemiBold") e
 * as troca por Arial MT / Arial Black / Times New Roman. Pra que o layout
 * continue batendo, ele compensa a diferença de métrica do substituto: escala
 * horizontal (<w:w w:val="105"/>), espaçamento entre letras (<w:spacing
 * w:val="-9"/>) e entrelinha como múltiplo da altura do Arial.
 *
 * Aqui só trocamos a tipografia e refazemos essas compensações pra fonte real —
 * nenhum parágrafo, posição, tabela ou imagem muda:
 *   1. lê do PDF (que nós mesmos geramos) a face e o avanço real de cada glifo;
 *   2. acha cada parágrafo/run do docx no texto do PDF e troca o substituto
 *      pela face original (família do Word + negrito/itálico);
 *   3. espaçamento entre letras = avanço real no PDF − avanço natural da fonte
 *      (o Chromium arredonda cada avanço pra pixel inteiro; o Word não);
 *   4. entrelinha "auto" reescalada pela razão de altura substituto/real;
 *   5. embute as faces usadas (word/fonts/*.odttf), pra abrir certo em quem
 *      não tem a fonte instalada.
 *
 * As faces vêm do app, em `<app>/word-fonts/` (geradas junto com as fontes do
 * editor pelo `npm run fonts:import` de lex-studio-v2 — mesma origem das fontes
 * que o PDF usa, então não tem como divergir). Texto numa fonte fora do
 * manifest fica como o iLovePDF deixou.
 */
import { randomUUID } from 'node:crypto';
import { brotliDecompressSync } from 'node:zlib';

import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';

type Face = {
  family: string;
  bold: boolean;
  italic: boolean;
  lineHeight: number;
  file?: string;
  /** Instância de fonte variável: arquivo de origem e FontBBox (milésimos de em). */
  variable?: string;
  bbox?: number[];
};

type FontData = {
  baseUrl: string;
  faces: Record<string, Face>;
  metrics: Map<Face, Record<string, number>>;
  instances: Map<string, Face[]>;
};

/** O manifest muda a cada deploy do app com fonte nova; relê de tempos em tempos. */
const MANIFEST_TTL_MS = 5 * 60_000;
const FETCH_TIMEOUT_MS = 8_000;

const fetchBytes = async (url: string): Promise<Buffer> => {
  const res = await fetch(url, { signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
  if (!res.ok) throw new Error(`word-fonts: ${res.status} em ${url}`);
  return Buffer.from(await res.arrayBuffer());
};

/**
 * Os arquivos são brotli (.br). Se algum CDN no caminho servir com
 * Content-Encoding: br, o fetch já entrega descomprimido — aceita os dois.
 */
const unbrotli = (buf: Buffer): Buffer => {
  try {
    return brotliDecompressSync(buf);
  } catch {
    return buf;
  }
};

let fontData: { data: FontData; at: number } | null = null;
const loadFonts = async (baseUrl: string): Promise<FontData> => {
  if (fontData && fontData.data.baseUrl === baseUrl && Date.now() - fontData.at < MANIFEST_TTL_MS) return fontData.data;
  try {
    const [manifest, metricsBr] = await Promise.all([
      fetchBytes(new URL('manifest.json', baseUrl).href),
      fetchBytes(new URL('metrics.json.br', baseUrl).href),
    ]);
    const faces = JSON.parse(manifest.toString('utf8')) as Record<string, Face>;
    const raw = JSON.parse(unbrotli(metricsBr).toString('utf8')) as Record<string, Record<string, number>>;
    const metrics = new Map<Face, Record<string, number>>();
    const instances = new Map<string, Face[]>();
    for (const [ps, face] of Object.entries(faces)) {
      if (raw[ps]) metrics.set(face, raw[ps]);
      if (face.variable) instances.set(face.variable, [...(instances.get(face.variable) ?? []), face]);
    }
    fontData = { data: { baseUrl, faces, metrics, instances }, at: Date.now() };
  } catch (err) {
    // App fora do ar por um instante: a última versão lida serve.
    if (fontData?.data.baseUrl !== baseUrl) throw err;
  }
  return fontData!.data;
};

/** Arquivos de fonte têm hash no nome — nunca mudam, podem ficar em memória. */
const fontFiles = new Map<string, Promise<Buffer>>();
/** ~100–300 KB por face: 60 cobrem os documentos do dia (≈15 MB no máximo). */
const MAX_CACHED_FONTS = 60;
const loadFontFile = (url: string): Promise<Buffer> => {
  let file = fontFiles.get(url);
  if (!file) {
    if (fontFiles.size >= MAX_CACHED_FONTS) fontFiles.clear();
    file = fetchBytes(url).then(unbrotli);
    file.catch(() => fontFiles.delete(url));
    fontFiles.set(url, file);
  }
  return file;
};

/**
 * Face do manifest pra uma fonte do PDF. Fonte variável o Chromium grava com o
 * nome do arquivo ("Montserrat-Thin") qualquer que seja o peso — aí vale a
 * instância cujo bbox bate com o FontBBox do PDF. Nomes "FamiliaRoman-Peso"
 * (Work Sans variável do Google) caem na família estática equivalente.
 */
const resolveFace = ({ faces, instances }: FontData, name: string | undefined, bbox?: number[]): Face | null => {
  const ps = name?.replace(/^[A-Z]{6}\+/, '');
  if (!ps) return null;
  const face = faces[ps] ?? faces[ps.replace(/(?:Roman|Upright)(?=-)/, '')];
  if (!face?.variable || !bbox || bbox.length !== 4) return face ?? null;
  let best = face;
  let bestDistance = Infinity;
  for (const candidate of instances.get(face.variable) ?? []) {
    const distance = candidate.bbox!.reduce((sum, v, i) => sum + Math.abs(v - bbox[i]), 0);
    if (distance < bestDistance) [best, bestDistance] = [candidate, distance];
  }
  return best;
};

/**
 * Altura de linha simples (usWinAscent + usWinDescent, em em) dos substitutos
 * que o iLovePDF usa. É a base do múltiplo que ele grava em w:line.
 */
const SUBSTITUTE_LINE_HEIGHT: Record<string, number> = {
  Arial: 1.1172,
  'Arial MT': 1.1172,
  'Arial Black': 1.4102,
  'Arial Narrow': 1.1172,
  'Times New Roman': 1.1074,
  'Courier New': 1.1328,
  Verdana: 1.2153,
  'Trebuchet MS': 1.1641,
  Georgia: 1.1362,
};

// ---------------------------------------------------------------------------
// PDF: face e avanço real de cada glifo
// ---------------------------------------------------------------------------

/** Espaços não entram no alinhamento: o docx e o PDF quebram/juntam diferente. */
const IGNORED = /[\s\u00ad\u200b-\u200d\ufeff]/gu;
/**
 * Forma de comparação: NFC, ligaduras de apresentação abertas ("ﬁ" → "fi") e
 * sem espaços. NFKC inteiro não serve: abriria "½" em 3 caracteres e o texto
 * corrompido pelo iLovePDF deixaria de ter o mesmo tamanho do original.
 */
const normalize = (s: string): string =>
  s
    .normalize('NFC')
    .replace(/[\ufb00-\ufb06]/g, (c) => c.normalize('NFKC'))
    .replace(IGNORED, '');

type PdfText = {
  /** Caracteres visíveis, na ordem do content stream. */
  text: string;
  faces: (Face | null)[];
  /** Avanço real − natural, em em (NaN = sem métrica). */
  extra: Float64Array;
  /** Mesmo, pros espaços logo antes do caractere i (NaN = nenhum). */
  spaceExtra: Float64Array;
  /**
   * 1 = o caractere i está colado no i−1: mesmo fragmento de texto, sem glifo
   * de espaço entre eles. Um espaço do docx nesse ponto foi inventado pelo
   * iLovePDF (ele lê letter-spacing largo como separação: "TRI MESTRE").
   */
  joined: Uint8Array;
  /** Palavras do PDF (sequências coladas, sem pontuação nas pontas). */
  words: Set<string>;
};

type PdfGlyph = { unicode?: string; width?: number } | number;

/** Palavra sem pontuação nas pontas ("serviços." → "serviços"). */
const bareWord = (w: string): string => w.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '');

const readPdfText = async (pdf: Buffer, fonts: FontData): Promise<PdfText> => {
  const { getDocument, OPS } = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await getDocument({
    data: new Uint8Array(pdf),
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    verbosity: 0,
  }).promise;
  const { metrics } = fonts;
  const chars: string[] = [];
  const faces: (Face | null)[] = [];
  const extra: number[] = [];
  const spaceExtra: number[] = [];
  const joined: number[] = [];
  let pendingSpace: number[] = [];
  let sameFragment = false;

  const push = (unicode: string, face: Face | null, actualEm: number) => {
    const metric = face ? metrics.get(face) : undefined;
    const cps = [...unicode];
    const natural = cps.length === 1 && metric ? metric[cps[0].codePointAt(0)!] : undefined;
    const delta = natural === undefined ? NaN : actualEm - natural;
    const visible = normalize(unicode);
    if (!visible) {
      if (/\s/.test(unicode)) pendingSpace.push(delta);
      return;
    }
    const spaces = pendingSpace.filter((d) => !Number.isNaN(d));
    const before = spaces.length ? spaces.reduce((a, b) => a + b, 0) / spaces.length : NaN;
    const glued = sameFragment && pendingSpace.length === 0 && chars.length > 0;
    pendingSpace = [];
    for (const [k, ch] of [...visible].entries()) {
      chars.push(ch);
      faces.push(face);
      extra.push(visible.length === 1 ? delta : NaN);
      spaceExtra.push(before);
      joined.push(k > 0 || glued ? 1 : 0);
    }
    sameFragment = true;
  };

  try {
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const { fnArray, argsArray } = await page.getOperatorList();
      let face: Face | null = null;
      let size = 1;
      let unitsPerEm = 0.001;
      for (let i = 0; i < fnArray.length; i++) {
        const fn = fnArray[i];
        const args = argsArray[i];
        if (fn === OPS.setFont) {
          size = Math.abs(args[1] as number) || 1;
          try {
            const font = page.commonObjs.get(args[0] as string) as {
              name?: string;
              fontMatrix?: number[];
              bbox?: number[];
            };
            face = resolveFace(fonts, font?.name, font?.bbox);
            unitsPerEm = font?.fontMatrix?.[0] ?? 0.001;
          } catch {
            face = null;
          }
          continue;
        }
        // Novo fragmento de texto: posição absoluta ou mudança de linha.
        if (fn === OPS.setTextMatrix || fn === OPS.beginText || fn === OPS.endText) sameFragment = false;
        if (fn === OPS.moveText && Math.abs((args as number[])[1]) > 1e-6) sameFragment = false;
        if (fn !== OPS.showText) continue;
        const glyphs = args[0] as PdfGlyph[];
        const real = glyphs.filter((g): g is { unicode?: string; width?: number } => typeof g !== 'number');
        // O Chromium posiciona glifo a glifo (showText de 1 glifo + moveText):
        // o dx do moveText seguinte é o avanço exato que ele usou.
        const next = fnArray[i + 1] === OPS.moveText ? (argsArray[i + 1] as number[]) : null;
        glyphs.forEach((g, k) => {
          if (typeof g === 'number' || !g.unicode) return;
          let advance = (g.width ?? 0) * unitsPerEm;
          // Ajuste de TJ logo depois do glifo (milésimos de em, negativo = avança).
          const adj = glyphs[k + 1];
          if (typeof adj === 'number') advance -= adj / 1000;
          // Só vale se for mesmo o passo de uma letra — na mesma linha de base
          // o moveText também pode ser um salto pra outra coluna.
          const step = next && Math.abs(next[1]) < 1e-6 ? next[0] / size : NaN;
          const jump = real.length === 1 && !Number.isNaN(step) && Math.abs(step - advance) >= 0.25;
          if (real.length === 1 && !jump && !Number.isNaN(step)) advance = step;
          push(g.unicode, face, advance);
          // Salto na mesma linha de base (tab, coluna): o próximo glifo é outro fragmento.
          if (jump) sameFragment = false;
        });
      }
      page.cleanup();
    }
  } finally {
    await doc.destroy();
  }
  const words = new Set<string>();
  let word = '';
  chars.forEach((ch, i) => {
    if (!joined[i] && word) words.add(bareWord(word));
    word = joined[i] ? word + ch : ch;
  });
  if (word) words.add(bareWord(word));
  return {
    text: chars.join(''),
    faces,
    extra: Float64Array.from(extra),
    spaceExtra: Float64Array.from(spaceExtra),
    joined: Uint8Array.from(joined),
    words,
  };
};

// ---------------------------------------------------------------------------
// DOCX: parágrafos, runs e propriedades
// ---------------------------------------------------------------------------

type Range = { start: number; end: number };
type Run = {
  rPr: Range | null;
  insertAt: number;
  text: string;
  /** Conteúdo de cada <w:t> do run (entre as tags). */
  texts: Range[];
  para: number;
  fallback: boolean;
};
type Paragraph = { id: number; openEnd: number; pPr: Range | null; markRPr: Range | null };

const TAG = /<(\/?)(w:p|w:pPr|w:r|w:rPr|w:t|mc:Fallback)(?=[\s/>])[^>]*?(\/?)>/g;

const unescapeXml = (s: string): string =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(x?)([0-9a-f]+);/gi, (_, hex, n) => String.fromCodePoint(parseInt(n, hex ? 16 : 10)))
    .replace(/&amp;/g, '&');

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeAttr = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** Varre um part (document/header/footer) só registrando posições. */
const scanPart = (xml: string): { runs: Run[]; paragraphs: Paragraph[] } => {
  const runs: Run[] = [];
  const paragraphs: Paragraph[] = [];
  const stack: { tag: string; run?: Run; para?: Paragraph; start?: number }[] = [];
  let fallbackDepth = 0;
  let textStart = -1;

  const currentPara = (): Paragraph | undefined => {
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].tag === 'w:p') return stack[i].para;
    return undefined;
  };
  const currentRun = (): Run | undefined => {
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].tag === 'w:r') return stack[i].run;
    return undefined;
  };
  const attachRPr = (owner: (typeof stack)[number] | undefined, range: Range) => {
    if (owner?.tag === 'w:r' && owner.run) owner.run.rPr = range;
    else if (owner?.tag === 'w:pPr') {
      const para = currentPara();
      if (para) para.markRPr = range;
    }
  };

  for (const m of xml.matchAll(TAG)) {
    const [whole, closing, tag, selfClosing] = m;
    const start = m.index!;
    const end = start + whole.length;

    if (tag === 'w:t') {
      if (closing) {
        const run = currentRun();
        if (run && textStart >= 0) {
          run.text += unescapeXml(xml.slice(textStart, start));
          run.texts.push({ start: textStart, end: start });
        }
        textStart = -1;
      } else if (!selfClosing) {
        textStart = end;
      }
      continue;
    }
    if (tag === 'mc:Fallback') {
      if (!selfClosing) fallbackDepth += closing ? -1 : 1;
      continue;
    }
    if (tag === 'w:rPr' && selfClosing) {
      attachRPr(stack[stack.length - 1], { start, end });
      continue;
    }
    if (tag === 'w:pPr' && selfClosing) {
      const para = currentPara();
      if (para && stack[stack.length - 1]?.tag === 'w:p') para.pPr = { start, end };
      continue;
    }
    if (selfClosing) continue;

    if (closing) {
      // Fecha até o tag correspondente (tolera aninhamento que não rastreamos).
      while (stack.length && stack[stack.length - 1].tag !== tag) stack.pop();
      const open = stack.pop();
      if (!open) continue;
      if (tag === 'w:rPr') attachRPr(stack[stack.length - 1], { start: open.start!, end });
      if (tag === 'w:pPr' && stack[stack.length - 1]?.tag === 'w:p') {
        stack[stack.length - 1].para!.pPr = { start: open.start!, end };
      }
      continue;
    }

    if (tag === 'w:p') {
      const para: Paragraph = { id: paragraphs.length, openEnd: end, pPr: null, markRPr: null };
      paragraphs.push(para);
      stack.push({ tag, para });
    } else if (tag === 'w:r') {
      const run: Run = {
        rPr: null,
        insertAt: end,
        text: '',
        texts: [],
        para: currentPara()?.id ?? -1,
        fallback: fallbackDepth > 0,
      };
      runs.push(run);
      stack.push({ tag, run });
    } else {
      stack.push({ tag, start });
    }
  }
  return { runs, paragraphs };
};

const attr = (xml: string, element: string, name: string): string | null =>
  xml.match(new RegExp(`<w:${element}\\b[^>]*\\bw:${name}="([^"]*)"`))?.[1] ?? null;

/** Conteúdo do pPr sem o rPr da marca de parágrafo (que tem seu próprio w:spacing). */
const withoutRPr = (pPr: string): string =>
  pPr.replace(/<w:rPr\s*\/>|<w:rPr\b[\s\S]*?<\/w:rPr>/g, (s) => ' '.repeat(s.length));

/** Liga/desliga de w:b, w:i...: true/false quando declarado, null quando herda. */
const toggle = (rPr: string, tag: string): boolean | null => {
  const el = rPr.match(new RegExp(`<w:${tag}(?=[\\s/>])[^>]*/>`))?.[0];
  if (!el) return null;
  return !/w:val="(?:0|false|off)"/.test(el);
};

type StyleInfo = {
  font: string | null;
  bold: boolean;
  italic: boolean;
  sz: number | null;
  line: number | null;
  lineAuto: boolean;
};

/** Propriedades de cada estilo de parágrafo, já resolvendo basedOn e docDefaults. */
const readStyles = (stylesXml: string): Map<string, StyleInfo> => {
  const defaults = stylesXml.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/)?.[0] ?? '';
  const base: StyleInfo = {
    font: null,
    bold: false,
    italic: false,
    sz: Number(attr(defaults, 'sz', 'val')) || 20,
    line: Number(attr(defaults, 'spacing', 'line')) || 240,
    lineAuto: (attr(defaults, 'spacing', 'lineRule') ?? 'auto') === 'auto',
  };
  const direct = new Map<string, { xml: string; basedOn: string | null }>();
  for (const m of stylesXml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    direct.set(m[1], { xml: m[2], basedOn: attr(m[2], 'basedOn', 'val') });
  }
  const resolved = new Map<string, StyleInfo>();
  const resolve = (id: string, depth = 0): StyleInfo => {
    const cached = resolved.get(id);
    if (cached) return cached;
    const s = direct.get(id);
    const parent = s?.basedOn && depth < 10 ? resolve(s.basedOn, depth + 1) : base;
    if (!s) return parent;
    const rPr = s.xml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
    const pPr = withoutRPr(s.xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '');
    const line = attr(pPr, 'spacing', 'line');
    const info: StyleInfo = {
      font: attr(rPr, 'rFonts', 'ascii') ?? parent.font,
      bold: toggle(rPr, 'b') ?? parent.bold,
      italic: toggle(rPr, 'i') ?? parent.italic,
      sz: Number(attr(rPr, 'sz', 'val')) || parent.sz,
      line: line ? Number(line) : parent.line,
      lineAuto: line ? (attr(pPr, 'spacing', 'lineRule') ?? 'auto') === 'auto' : parent.lineAuto,
    };
    resolved.set(id, info);
    return info;
  };
  for (const id of direct.keys()) resolve(id);
  resolved.set('', base);
  return resolved;
};

// ---------------------------------------------------------------------------
// Alinhamento parágrafo/run ↔ PDF
// ---------------------------------------------------------------------------

/** Quanto à frente do cursor um trecho ainda conta como "o próximo". */
const WINDOW = 4000;
/** Texto curto ("1", "de") casa em qualquer lugar — não serve pra ressincronizar. */
const RESYNC_MIN = 12;

/**
 * Onde cada run caiu no PDF: [início, fim) no texto visível. `reliable`: veio
 * do parágrafo inteiro (posição exata entre runs vizinhos). `repair`: o texto
 * certo, quando o do iLovePDF veio com letras trocadas ("Rescisio" → "Rescisão").
 */
type Placement =
  | {
      at: number;
      length: number;
      /** Âncora confiável pros vizinhos (parágrafo achado inteiro ou consertado). */
      reliable: boolean;
      /** Cada caractere do docx corresponde ao caractere `at + k` do PDF (sem inserções). */
      exact: boolean;
      /** Texto certo de cada caractere visível do run ('' = sobrou no docx). */
      repair?: string[];
    }
  | undefined;

/** Fração de caracteres iguais na mesma posição (strings do mesmo tamanho). */
const sameness = (a: string, b: string): number => {
  let same = 0;
  for (let i = 0; i < a.length; i++) if (a[i] === b[i]) same++;
  return same / a.length;
};

/**
 * Letras trocadas pelo iLovePDF: mesmo tamanho, maioria das letras igual
 * ("Ǫuulificução dus purtєs" × "Qualificação das partes": 71%) e trocas
 * isoladas, sempre letra por letra. Um bloco inteiro diferente ("Ocorrência"
 * × "serviços.") é outro texto, não letra trocada. Número e pontuação nunca.
 */
const LETTER = /\p{L}/u;
const looksGarbled = (docx: string, pdf: string): boolean => {
  if (docx.length !== pdf.length || docx.length < 3) return false;
  if (sameness(docx, pdf) < (docx.length >= 5 ? 0.5 : 0.66)) return false;
  let streak = 0;
  for (let i = 0; i < docx.length; i++) {
    if (docx[i] === pdf[i]) {
      streak = 0;
      continue;
    }
    if (!LETTER.test(docx[i]) || !LETTER.test(pdf[i]) || ++streak > 2) return false;
  }
  return true;
};

/**
 * Palavra do docx a que pertence cada caractere visível do parágrafo. Palavra
 * termina em espaço e em fim de run (o iLovePDF cola rótulos de caixas
 * diferentes em runs vizinhos: "Prazo" + "Consequência").
 */
const wordsOf = (raws: string[]): string[] => {
  const out: string[] = [];
  for (const raw of raws) {
    for (const token of raw.split(/\s+/)) {
      const visible = normalize(token);
      const bare = bareWord(visible);
      for (let k = 0; k < visible.length; k++) out.push(bare);
    }
  }
  return out;
};

/** Trigramas com mais ocorrências que isso não ajudam a localizar nada. */
const MAX_GRAM_HITS = 64;

class Aligner {
  private cursor = 0;
  constructor(readonly pdf: PdfText) {}

  /** `ordered`: o part segue a ordem do PDF (corpo). Header/footer/fallback não. */
  private find(text: string, ordered: boolean): number | undefined {
    const { pdf, cursor } = this;
    const ahead = pdf.text.indexOf(text, cursor);
    const behind = cursor > 0 ? pdf.text.lastIndexOf(text, cursor - 1) : -1;
    if (ahead < 0 && behind < 0) return undefined;
    const aheadGap = ahead >= 0 ? ahead - cursor : Infinity;
    const behindGap = behind >= 0 ? Math.max(0, cursor - (behind + text.length)) : Infinity;
    // Na ordem normal o trecho está logo à frente. Mas o iLovePDF às vezes
    // reordena blocos (capa, caixas de texto): aí a ocorrência certa é a de
    // trás, e a da frente é outra igual mais adiante ("Dados" × "Dados Bancários").
    if (aheadGap <= WINDOW && aheadGap <= 2 * behindGap) {
      if (ordered) this.cursor = ahead + text.length;
      return ahead;
    }
    if (behind < 0 || aheadGap < behindGap) {
      if (ordered && text.length >= RESYNC_MIN) this.cursor = ahead + text.length;
      return ahead;
    }
    return behind;
  }

  /**
   * Procura primeiro o parágrafo inteiro — longo, casa num lugar só — e
   * reparte pelos runs; palavra solta ("A", "10") casaria com a ocorrência
   * vizinha. Se o parágrafo não aparece contíguo no PDF, busca run a run.
   */
  placeParagraph(texts: string[], raws: string[], ordered: boolean): Placement[] {
    const whole = texts.join('');
    let at = whole ? this.find(whole, ordered) : undefined;
    const garbled = at === undefined && whole ? this.findGarbled(whole, wordsOf(raws), ordered) : undefined;
    at ??= garbled;
    if (at !== undefined) {
      let offset = at;
      return texts.map((t) => {
        const right = this.pdf.text.slice(offset, offset + t.length);
        const placement = {
          at: offset,
          length: t.length,
          reliable: true,
          exact: true,
          ...(right !== t ? { repair: [...right] } : {}),
        };
        offset += t.length;
        return placement;
      });
    }
    return texts.map((t) => {
      const pos = t ? this.find(t, ordered) : undefined;
      return pos === undefined ? undefined : { at: pos, length: t.length, reliable: false, exact: false };
    });
  }

  /**
   * Acha no PDF um parágrafo que o iLovePDF gravou com letras trocadas:
   * trigramas intactos ("lif", "ção") votam na posição de início; vence a
   * posição mais votada cujo trecho do PDF passa em `looksGarbled`.
   */
  private findGarbled(text: string, words: string[], ordered: boolean): number | undefined {
    const { pdf } = this;
    const votes = new Map<number, number>();
    for (let i = 0; i + 3 <= text.length; i++) {
      const gram = text.slice(i, i + 3);
      const hits: number[] = [];
      for (let p = pdf.text.indexOf(gram); p >= 0 && hits.length <= MAX_GRAM_HITS; p = pdf.text.indexOf(gram, p + 1)) {
        hits.push(p - i);
      }
      if (hits.length > MAX_GRAM_HITS) continue;
      for (const start of hits) votes.set(start, (votes.get(start) ?? 0) + 1);
    }
    const ranked = [...votes.entries()]
      .filter(([start]) => start >= 0)
      .sort((a, b) => b[1] - a[1] || Math.abs(a[0] - this.cursor) - Math.abs(b[0] - this.cursor));
    for (const [start] of ranked.slice(0, 8)) {
      const candidate = pdf.text.slice(start, start + text.length);
      if (!looksGarbled(text, candidate) || this.isRealText(text, candidate, words)) continue;
      if (ordered && start >= this.cursor && start - this.cursor <= WINDOW) this.cursor = start + text.length;
      return start;
    }
    return undefined;
  }

  /** Face majoritária do trecho; null = só fontes fora do manifest. */
  /**
   * O docx pode diferir do trecho achado por outro motivo: o iLovePDF junta
   * rótulos de caixas diferentes num parágrafo ("Prazo" + "Consequência") e o
   * trecho vizinho no PDF só se parece. Se o texto do docx em volta de uma
   * diferença existe tal e qual no PDF, ele é real — não é letra trocada.
   */
  private isRealText(text: string, candidate: string, words: string[]): boolean {
    for (let i = 0; i < text.length; i++) {
      if (text[i] !== candidate[i] && words[i].length >= 3 && this.pdf.words.has(words[i])) return true;
    }
    return false;
  }

  /**
   * Espaço que o iLovePDF inventou dentro de uma palavra: no PDF os dois
   * caracteres em volta estão colados, no mesmo fragmento. `at` = índice do
   * caractere logo depois do espaço.
   */
  spuriousSpaceAt(at: number): boolean {
    return at > 0 && at < this.pdf.text.length && this.pdf.joined[at] === 1;
  }

  faceAt({ at, length }: { at: number; length: number }): Face | null {
    const votes = new Map<Face, number>();
    for (let i = at; i < at + length; i++) {
      const face = this.pdf.faces[i];
      if (face) votes.set(face, (votes.get(face) ?? 0) + 1);
    }
    let best: Face | null = null;
    let bestVotes = 0;
    for (const [face, n] of votes) if (n > bestVotes) [best, bestVotes] = [face, n];
    return best;
  }

  /** Diferença média de avanço (em) do trecho; run só de espaço usa o espaço antes de `at`. */
  extraAt({ at, length }: { at: number; length: number }): number {
    if (length === 0) return this.pdf.spaceExtra[at] ?? NaN;
    let sum = 0;
    let n = 0;
    for (let i = at; i < at + length; i++) {
      const d = this.pdf.extra[i];
      if (!Number.isNaN(d)) {
        sum += d;
        n++;
      }
    }
    return n ? sum / n : NaN;
  }
}

// ---------------------------------------------------------------------------
// Reescrita das propriedades
// ---------------------------------------------------------------------------

/** O que o iLovePDF gravou por causa do substituto: fonte, peso, estilo e compensações. */
const STRIP_FACE = /<w:(?:rFonts|b|bCs|i|iCs|w)(?=[\s/>])[^>]*\/>/g;
const STRIP_SPACING = /<w:spacing(?=[\s/>])[^>]*\/>/g;

/** Filhos de rPr que vêm depois de w:spacing no schema (CT_RPr). */
const AFTER_SPACING = /<w:(?:w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath|rPrChange)(?=[\s/>])/;
/** Filhos de pPr que vêm depois de w:spacing no schema (CT_PPr). */
const PPR_AFTER_SPACING = /<w:(?:ind|contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)(?=[\s/>])/;

const faceProps = (face: Face): string => {
  const family = escapeAttr(face.family);
  const toggle = (tag: string, on: boolean) => (on ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`);
  return (
    `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>` +
    toggle('b', face.bold) +
    toggle('bCs', face.bold) +
    toggle('i', face.italic) +
    toggle('iCs', face.italic)
  );
};

const insertBefore = (inner: string, after: RegExp, element: string): string => {
  const m = inner.match(after);
  return m ? inner.slice(0, m.index) + element + inner.slice(m.index) : inner + element;
};

/**
 * rPr com a face real, mantendo a ordem do schema (rStyle antes de rFonts).
 * `spacing`: twips a gravar em w:spacing; undefined mantém o do iLovePDF.
 */
const rewriteRPr = (rPr: string, face: Face, spacing?: number): string => {
  const selfClosing = /^<w:rPr\s*\/>$/.test(rPr);
  const open = selfClosing ? '<w:rPr>' : rPr.match(/^<w:rPr[^>]*>/)![0];
  let inner = selfClosing ? '' : rPr.slice(open.length, -'</w:rPr>'.length);
  inner = inner.replace(STRIP_FACE, '');
  if (spacing !== undefined) {
    inner = inner.replace(STRIP_SPACING, '');
    if (spacing !== 0) inner = insertBefore(inner, AFTER_SPACING, `<w:spacing w:val="${spacing}"/>`);
  }
  const rStyle = inner.match(/^<w:rStyle[^>]*\/>/)?.[0] ?? '';
  inner = rStyle + faceProps(face) + inner.slice(rStyle.length);
  return `${open}${inner}</w:rPr>`;
};

type Edit = { start: number; end: number; text: string };
const applyEdits = (xml: string, edits: Edit[]): string => {
  const out: string[] = [];
  let pos = 0;
  // Na mesma posição, inserção (start === end) antes da substituição.
  for (const e of [...edits].sort((a, b) => a.start - b.start || a.end - b.end)) {
    if (e.start < pos) continue; // sobreposta: nunca deveria acontecer, mas não corrompe o XML
    out.push(xml.slice(pos, e.start), e.text);
    pos = e.end;
  }
  out.push(xml.slice(pos));
  return out.join('');
};

/** Reescala a entrelinha "auto" do parágrafo; devolve a edição (ou nada). */
const lineEdit = (xml: string, para: Paragraph, style: StyleInfo, ratio: number): Edit | null => {
  if (!para.pPr) {
    if (!style.lineAuto || !style.line) return null;
    const line = Math.round(style.line * ratio);
    return {
      start: para.openEnd,
      end: para.openEnd,
      text: `<w:pPr><w:spacing w:line="${line}" w:lineRule="auto"/></w:pPr>`,
    };
  }
  const pPr = xml.slice(para.pPr.start, para.pPr.end);
  const visible = withoutRPr(pPr);
  const spacing = visible.match(/<w:spacing\b[^>]*\/>/);
  if (spacing) {
    const el = spacing[0];
    const own = attr(el, 'spacing', 'line');
    const rule = attr(el, 'spacing', 'lineRule') ?? (own ? 'auto' : style.lineAuto ? 'auto' : 'exact');
    const base = own ? Number(own) : style.line;
    if (rule !== 'auto' || !base) return null;
    const line = Math.round(base * ratio);
    const next = own
      ? el.replace(/w:line="[^"]*"/, `w:line="${line}"`)
      : el.replace(/\s*\/>$/, ` w:line="${line}" w:lineRule="auto"/>`);
    const start = para.pPr.start + spacing.index!;
    return { start, end: start + el.length, text: next };
  }
  if (!style.lineAuto || !style.line) return null;
  const line = Math.round(style.line * ratio);
  const el = `<w:spacing w:line="${line}" w:lineRule="auto"/>`;
  if (/^<w:pPr\s*\/>$/.test(pPr)) return { start: para.pPr.start, end: para.pPr.end, text: `<w:pPr>${el}</w:pPr>` };
  const open = pPr.match(/^<w:pPr[^>]*>/)![0];
  // Busca no pPr original: o rPr da marca (mascarado em `visible`) também vem depois de spacing.
  const m = pPr.slice(open.length).match(PPR_AFTER_SPACING);
  const at = para.pPr.start + open.length + (m ? m.index! : pPr.length - open.length - '</w:pPr>'.length);
  return { start: at, end: at, text: el };
};

// ---------------------------------------------------------------------------
// Embutir as fontes
// ---------------------------------------------------------------------------

const SLOT = (f: Face): string =>
  f.bold && f.italic ? 'embedBoldItalic' : f.bold ? 'embedBold' : f.italic ? 'embedItalic' : 'embedRegular';
const SLOT_ORDER = ['embedRegular', 'embedBold', 'embedItalic', 'embedBoldItalic'];

/** ECMA-376 §17.8.1: XOR dos 32 primeiros bytes com a chave GUID invertida. */
const obfuscate = (ttf: Buffer, guid: string): Uint8Array => {
  const hex = guid.replace(/[{}-]/g, '');
  const key = Array.from({ length: 16 }, (_, i) => parseInt(hex.slice(30 - 2 * i, 32 - 2 * i), 16));
  const out = new Uint8Array(ttf);
  for (let i = 0; i < 32; i++) out[i] ^= key[i % 16];
  return out;
};

const embedFonts = async (files: Zippable, used: Set<Face>, baseUrl: string): Promise<void> => {
  const byFamily = new Map<string, Face[]>();
  for (const face of used) {
    if (!face.file) continue;
    const list = byFamily.get(face.family) ?? [];
    if (!list.some((f) => SLOT(f) === SLOT(face))) list.push(face);
    byFamily.set(face.family, list);
  }
  if (byFamily.size === 0) return;
  const ttfs = new Map<Face, Buffer>();
  await Promise.all(
    [...byFamily.values()].flat().map(async (face) => {
      ttfs.set(face, await loadFontFile(new URL(face.file!, baseUrl).href));
    }),
  );

  let fontTable = strFromU8(files['word/fontTable.xml'] as Uint8Array);
  const rels: string[] = [];
  let n = 0;
  for (const [family, faces] of byFamily) {
    const embeds = faces
      .sort((a, b) => SLOT_ORDER.indexOf(SLOT(a)) - SLOT_ORDER.indexOf(SLOT(b)))
      .map((face) => {
        n++;
        const guid = `{${randomUUID().toUpperCase()}}`;
        const ttf = ttfs.get(face)!;
        files[`word/fonts/font${n}.odttf`] = obfuscate(ttf, guid);
        rels.push(
          `<Relationship Id="rIdLexFont${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font${n}.odttf"/>`,
        );
        return `<w:${SLOT(face)} r:id="rIdLexFont${n}" w:fontKey="${guid}"/>`;
      })
      .join('');
    const name = escapeAttr(family);
    fontTable = fontTable.replace(
      new RegExp(`<w:font w:name="${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[\\s\\S]*?</w:font>`),
      '',
    );
    fontTable = fontTable.replace(
      '</w:fonts>',
      `<w:font w:name="${name}"><w:charset w:val="00"/><w:pitch w:val="variable"/>${embeds}</w:font></w:fonts>`,
    );
  }
  if (!fontTable.includes('xmlns:r=')) {
    fontTable = fontTable.replace(
      '<w:fonts ',
      '<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
    );
  }
  files['word/fontTable.xml'] = strToU8(fontTable);

  const relsPath = 'word/_rels/fontTable.xml.rels';
  const existing = files[relsPath]
    ? strFromU8(files[relsPath] as Uint8Array)
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  files[relsPath] = strToU8(existing.replace('</Relationships>', `${rels.join('')}</Relationships>`));

  let types = strFromU8(files['[Content_Types].xml'] as Uint8Array);
  if (!/Extension="odttf"/i.test(types)) {
    types = types.replace(
      '</Types>',
      '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>',
    );
  }
  files['[Content_Types].xml'] = strToU8(types);

  let settings = strFromU8(files['word/settings.xml'] as Uint8Array);
  if (!settings.includes('<w:embedTrueTypeFonts')) {
    // Ordem do schema (CT_Settings): só estes vêm antes de embedTrueTypeFonts.
    const before = settings.match(
      /^[\s\S]*?<w:settings[^>]*>(?:\s*<w:(?:writeProtection|view|zoom|removePersonalInformation|removeDateAndTime|doNotDisplayPageBoundaries|displayBackgroundShape|printPostScriptOverText|printFractionalCharacterWidth|printFormsData)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:(?:writeProtection|view|zoom)>))*/,
    );
    if (before) {
      settings = before[0] + '<w:embedTrueTypeFonts/>' + settings.slice(before[0].length);
    }
  }
  files['word/settings.xml'] = strToU8(settings);
};

// ---------------------------------------------------------------------------
// Entrada
// ---------------------------------------------------------------------------

/**
 * Palpite pra run sem par no PDF (texto que o iLovePDF gravou diferente):
 * a face original mais comum entre os runs achados com o mesmo substituto —
 * e, se possível, o mesmo negrito/itálico ("Arial Black" negrito ≠ não negrito).
 */
class SubstituteVotes {
  private readonly votes = new Map<string, Map<Face, number>>();

  add(substitute: string, bold: boolean, italic: boolean, face: Face, weight: number): void {
    for (const key of [substitute, `${substitute}|${bold}|${italic}`]) {
      const v = this.votes.get(key) ?? new Map<Face, number>();
      v.set(face, (v.get(face) ?? 0) + weight);
      this.votes.set(key, v);
    }
  }

  guess(substitute: string | null, bold: boolean | null, italic: boolean | null): Face | undefined {
    if (!substitute) return undefined;
    const exact = bold !== null && italic !== null ? this.votes.get(`${substitute}|${bold}|${italic}`) : undefined;
    const v = exact ?? this.votes.get(substitute);
    if (!v) return undefined;
    return [...v.entries()].sort((a, b) => b[1] - a[1])[0][0];
  }
}

export type DocxFontReport = {
  runs: number;
  matched: number;
  fallback: number;
  embedded: string[];
} & TextFixes;

const TEXT_PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

/** Ligaduras de apresentação (U+FB00–FB06): "ﬁ" num caractere só quebra busca e corretor no Word. */
const LIGATURE = /[\ufb00-\ufb06]/;

type TextFixes = { repaired: number; spacesRemoved: number; ligatures: number; split: number };

/**
 * Corrige o texto de um run contra o PDF: letras trocadas (placement.repair),
 * espaço inventado no meio da palavra e ligaduras. Anda nos caracteres do
 * docx acompanhando a posição no texto visível do PDF. null = nada muda.
 */
const fixRunText = (texts: string[], placement: Placement, aligner: Aligner, fixes: TextFixes): string[] | null => {
  let k = 0; // posição no texto visível do run
  let changed = false;
  const out = texts.map((text) => {
    let result = '';
    for (const c of text) {
      const n = normalize(c);
      if (!n) {
        // Só mexe em espaço quando a posição no PDF é exata.
        if (placement?.exact && /\s/.test(c) && aligner.spuriousSpaceAt(placement.at + k)) {
          fixes.spacesRemoved++;
          changed = true;
          continue;
        }
        result += c;
        continue;
      }
      const right = placement?.repair?.slice(k, k + n.length).join('');
      if (right !== undefined && right !== n) {
        result += right;
        fixes.repaired++;
        changed = true;
      } else if (LIGATURE.test(c)) {
        result += n;
        fixes.ligatures++;
        changed = true;
      } else {
        result += c;
      }
      k += n.length;
    }
    return result;
  });
  return changed ? out : null;
};

/**
 * Run que o iLovePDF juntou com texto de outra face (a mesma substituta pras
 * duas — SemiBold e Bold viram "Arial Black"; fim de uma linha colado no
 * começo da outra): vira um run por face. Só no caso simples e seguro:
 * posição exata no PDF, texto não consertado, `<w:r>[rPr]<w:t>…</w:t></w:r>`.
 */
const splitByFace = (
  xml: string,
  run: Run,
  rPr: string,
  placement: Placement,
  sz: number,
  aligner: Aligner,
  /** Texto já corrigido por fixRunText (mesma correspondência 1:1 com o PDF). */
  fixedText: string | undefined,
): { edit: Edit; faces: Face[] } | null => {
  if (!placement?.exact || run.texts.length !== 1) return null;
  const [t] = run.texts;
  const start = run.rPr?.start ?? run.insertAt;
  if (!/^<w:t(?:\s[^>]*)?>$/.test(xml.slice(run.rPr?.end ?? run.insertAt, t.start))) return null;
  if (!xml.startsWith('</w:t>', t.end)) return null;

  const segments: { text: string; face: Face; from: number; to: number }[] = [];
  let k = 0;
  for (const c of fixedText ?? unescapeXml(xml.slice(t.start, t.end))) {
    const n = normalize(c);
    const face = n ? aligner.pdf.faces[placement.at + k] : null;
    const last = segments.at(-1);
    if (last && (!face || face === last.face)) last.text += c;
    else if (face) segments.push({ text: c, face, from: k, to: k });
    else return null; // espaço antes da primeira letra: deixa pro caminho normal
    k += n.length;
    segments.at(-1)!.to = k;
  }
  if (segments.length < 2) return null;
  const text = segments
    .map((seg, i) => {
      const spacing = twips(aligner.extraAt({ at: placement.at + seg.from, length: seg.to - seg.from }), sz) ?? 0;
      const tail = i < segments.length - 1 ? '</w:r><w:r>' : '';
      return `${rewriteRPr(rPr || '<w:rPr/>', seg.face, spacing)}<w:t xml:space="preserve">${escapeText(seg.text)}</w:t>${tail}`;
    })
    .join('');
  return { edit: { start, end: t.end + '</w:t>'.length, text }, faces: segments.map((seg) => seg.face) };
};

const twips = (extraEm: number, halfPoints: number): number | undefined =>
  Number.isNaN(extraEm) ? undefined : Math.round(extraEm * (halfPoints / 2) * 20);

type PlacedParagraph = { group: Run[]; texts: string[]; placements: Placement[] };

/** Tamanho máximo do buraco (caracteres do PDF) em que o alinhamento é tentado. */
const SANDWICH_MAX_REGION = 20_000;
/** Mínimo de caracteres do docx alinhados a um caractere igual do PDF. */
const SANDWICH_IDENTITY = 0.5;
/** Inserção do PDF maior que isso é outro conteúdo (cabeçalho, caixa) — não entra no docx. */
const MAX_INSERT = 2;
/** Corrupção do iLovePDF troca poucos caracteres seguidos ("6çmo" → "Ação"); mais que isso é outro texto. */
const MAX_EDIT_RUN = 3;
/** Parágrafo curto ("Source Sans 3", "iv)") casa em lugar errado: não serve de borda nem é realinhado. */
const MIN_ANCHOR = 12;

/**
 * O conserto parece corrupção do iLovePDF, e não outro texto? Trocas curtas e
 * nenhuma palavra do docx que exista tal e qual no PDF é alterada ("Source"
 * existe no PDF — trocá-la por "Work" seria inventar).
 */
const plausibleRepair = (docx: string, replacements: string[], words: string[], pdfWords: Set<string>): boolean => {
  let run = 0;
  for (let k = 0; k < docx.length; k++) {
    // A contagem recomeça a cada palavra (os espaços não entram no alinhamento).
    if (k > 0 && words[k] !== words[k - 1]) run = 0;
    const replacement = replacements[k];
    // "i" → "if": o caractere continua lá, só ganhou uma inserção do lado.
    if (replacement === docx[k] || (replacement.length > 1 && replacement.includes(docx[k]))) {
      run = 0;
      continue;
    }
    if (++run > MAX_EDIT_RUN) return false;
    if (words[k].length >= 3 && pdfWords.has(words[k])) return false;
  }
  return true;
};

type Alignment = { replacements: string[]; pdfIndex: number[]; end: number; identity: number };

/**
 * Alinhamento semi-global (distância de edição): o texto do docx inteiro contra
 * um trecho qualquer de `pdf[from, to)`. Devolve, pra cada caractere do docx,
 * o texto do PDF que o substitui ('' se sobrou) e onde ele caiu.
 */
const alignText = (docx: string, pdf: string, from: number, to: number): Alignment | null => {
  const n = docx.length;
  const m = to - from;
  // Teto de 4M células (16 MB, ~30 ms): parágrafo gigante fica sem conserto
  // em vez de pesar memória e segurar o event loop das outras exportações.
  if (n === 0 || m <= 0 || n * m > 4_000_000) return null;
  const w = m + 1;
  const cost = new Int32Array((n + 1) * w);
  for (let i = 1; i <= n; i++) cost[i * w] = i;
  for (let i = 1; i <= n; i++) {
    const c = docx[i - 1];
    for (let j = 1; j <= m; j++) {
      const diag = cost[(i - 1) * w + j - 1] + (c === pdf[from + j - 1] ? 0 : 1);
      const up = cost[(i - 1) * w + j] + 1;
      const left = cost[i * w + j - 1] + 1;
      cost[i * w + j] = Math.min(diag, up, left);
    }
  }
  let endJ = 1;
  for (let j = 1; j <= m; j++) if (cost[n * w + j] < cost[n * w + endJ]) endJ = j;
  const replacements = new Array<string>(n).fill('');
  const pdfIndex = new Array<number>(n).fill(-1);
  let matches = 0;
  let pending = '';
  let i = n;
  let j = endJ;
  while (i > 0) {
    const here = cost[i * w + j];
    const same = j > 0 && docx[i - 1] === pdf[from + j - 1];
    if (j > 0 && here === cost[(i - 1) * w + j - 1] + (same ? 0 : 1)) {
      if (same) matches++;
      replacements[i - 1] = pdf[from + j - 1] + (pending.length <= MAX_INSERT ? pending : '');
      pdfIndex[i - 1] = from + j - 1;
      pending = '';
      i--;
      j--;
    } else if (j > 0 && here === cost[i * w + j - 1] + 1) {
      pending = pdf[from + j - 1] + pending;
      j--;
    } else {
      i--;
    }
  }
  return { replacements, pdfIndex, end: from + endJ, identity: matches / n };
};

/**
 * Texto que o iLovePDF corrompeu demais pra ser achado ("Montsgrrnt ®00: 6çmo",
 * "Nuчito 500 itúlico", "itfilico", "O123V56Y89" — com número trocado): entre
 * dois parágrafos vizinhos achados, cada parágrafo sem par é alinhado, em
 * ordem, ao trecho do PDF entre eles. O texto certo vem do PDF, número
 * incluído. As cópias em mc:Fallback do mesmo texto recebem o mesmo conserto.
 */
const fillSandwiches = (placed: PlacedParagraph[], pdf: PdfText): void => {
  const pdfText = pdf.text;
  const anchored = (p: PlacedParagraph) =>
    p.texts.join('').length >= MIN_ANCHOR && p.placements.some((pl) => pl?.reliable);
  const body = placed.filter((p) => !p.group[0].fallback && p.texts.join(''));
  const fixed = new Map<string, Placement[]>();
  let i = 0;
  while (i < body.length) {
    if (anchored(body[i])) {
      i++;
      continue;
    }
    let j = i;
    while (j < body.length && !anchored(body[j])) j++;
    const before = body[i - 1]?.placements.filter((pl) => pl?.reliable).at(-1);
    const after = body[j]?.placements.find((pl) => pl?.reliable);
    const gap = body.slice(i, j);
    i = j;
    if (!before || !after) continue;
    let cursor = before.at + before.length;
    const end = after.at;
    if (end - cursor > SANDWICH_MAX_REGION) continue;
    for (const p of gap) {
      const docx = p.texts.join('');
      // Curto já achado ("LEX DESIGN") fica como está; curto sem par ("iv)", "c)")
      // casaria em qualquer lugar — fica com a regra estrita de findGarbled. E
      // nenhum dos dois move o cursor: a posição de texto curto não é confiável.
      if (docx.length < MIN_ANCHOR || p.placements.some((pl) => pl?.reliable)) continue;
      const aligned = alignText(docx, pdfText, cursor, end);
      if (!aligned || aligned.identity < SANDWICH_IDENTITY) continue;
      if (!plausibleRepair(docx, aligned.replacements, wordsOf(p.group.map((r) => r.text)), pdf.words)) continue;
      let k = 0;
      p.placements = p.texts.map((t) => {
        const idx = aligned.pdfIndex.slice(k, k + t.length).filter((x) => x >= 0);
        const repair = aligned.replacements.slice(k, k + t.length);
        k += t.length;
        if (!t) return undefined;
        const at = idx.length ? idx[0] : cursor;
        return {
          at,
          length: idx.length ? idx[idx.length - 1] - at + 1 : 0,
          reliable: true,
          exact: false,
          ...(repair.join('') !== t ? { repair } : {}),
        };
      });
      fixed.set(p.texts.join('\u0000'), p.placements);
      cursor = aligned.end;
    }
  }
  for (const p of placed) {
    const copy = p.group[0].fallback ? fixed.get(p.texts.join('\u0000')) : undefined;
    if (copy) p.placements = copy;
  }
};

/**
 * `fontsUrl`: onde o app publica as fontes do Word (`<app>/word-fonts/`).
 */
export const restoreDocxFonts = async (
  docx: Buffer,
  pdf: Buffer,
  fontsUrl: string,
): Promise<{ docx: Buffer; report: DocxFontReport }> => {
  const fonts = await loadFonts(fontsUrl.endsWith('/') ? fontsUrl : `${fontsUrl}/`);
  const pdfText = await readPdfText(pdf, fonts);
  const files = unzipSync(new Uint8Array(docx)) as Zippable;
  const stylesXml = files['word/styles.xml'] ? strFromU8(files['word/styles.xml'] as Uint8Array) : '';
  const styles = readStyles(stylesXml);
  const aligner = new Aligner(pdfText);

  const parts = Object.keys(files)
    .filter((p) => TEXT_PARTS.test(p))
    .sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b)));

  type Located = {
    run: Run;
    rPr: string;
    placement: Placement;
    face: Face | null | undefined;
    substitute: string | null;
    bold: boolean;
    italic: boolean;
    sz: number;
  };
  type PartState = { xml: string; paragraphs: Paragraph[]; located: Located[]; styleOf: (p: number) => StyleInfo };
  const states = new Map<string, PartState>();
  const votes = new SubstituteVotes();

  for (const part of parts) {
    const xml = strFromU8(files[part] as Uint8Array);
    const { runs, paragraphs } = scanPart(xml);
    const styleOf = (p: number): StyleInfo => {
      const pPr = paragraphs[p]?.pPr;
      const id = pPr ? attr(xml.slice(pPr.start, pPr.end), 'pStyle', 'val') : null;
      return styles.get(id ?? 'Normal') ?? styles.get('')!;
    };
    const located: Located[] = [];
    const byParagraph = new Map<number, Run[]>();
    for (const run of runs) byParagraph.set(run.para, [...(byParagraph.get(run.para) ?? []), run]);

    const ordered = part === 'word/document.xml';
    const placed = [...byParagraph.values()].map((group) => {
      const texts = group.map((run) => normalize(run.text));
      const placements = aligner.placeParagraph(
        texts,
        group.map((run) => run.text),
        ordered && !group[0].fallback,
      );
      return { group, texts, placements };
    });
    if (ordered) fillSandwiches(placed, aligner.pdf);

    for (const { group, texts, placements } of placed) {
      group.forEach((run, i) => {
        const rPr = run.rPr ? xml.slice(run.rPr.start, run.rPr.end) : '';
        const style = styleOf(run.para);
        const placement = placements[i];
        const face = placement && texts[i] ? aligner.faceAt(placement) : undefined;
        const substitute = attr(rPr, 'rFonts', 'ascii') ?? style.font;
        const sz = Number(attr(rPr, 'sz', 'val')) || style.sz || 20;
        const bold = toggle(rPr, 'b') ?? style.bold;
        const italic = toggle(rPr, 'i') ?? style.italic;
        located.push({ run, rPr, placement, face, substitute, bold, italic, sz });
        if (face && substitute) votes.add(substitute, bold, italic, face, texts[i].length);
      });
    }
    states.set(part, { xml, paragraphs, located, styleOf });
  }

  const used = new Set<Face>();
  const report: DocxFontReport = {
    runs: 0,
    matched: 0,
    fallback: 0,
    embedded: [],
    repaired: 0,
    spacesRemoved: 0,
    ligatures: 0,
    split: 0,
  };

  for (const [part, { xml, paragraphs, located, styleOf }] of states) {
    const edits: Edit[] = [];
    // Face/substituto que definem a altura de cada parágrafo (o primeiro run com texto).
    const lead = new Map<number, { face: Face; substitute: string | null }>();
    let prev: { para: number; face: Face } | null = null;
    const justified = new Set(
      paragraphs
        .filter((p) => p.pPr && /^(both|distribute)$/.test(attr(xml.slice(p.pPr.start, p.pPr.end), 'jc', 'val') ?? ''))
        .map((p) => p.id),
    );

    for (const l of located) {
      if (!l.run.text) continue;
      const blank = normalize(l.run.text) === '';
      if (!blank) report.runs++;
      const fixed = fixRunText(
        l.run.texts.map((r) => unescapeXml(xml.slice(r.start, r.end))),
        l.placement,
        aligner,
        report,
      );
      const split = splitByFace(xml, l.run, l.rPr, l.placement, l.sz, aligner, fixed?.[0]);
      if (split) {
        edits.push(split.edit);
        split.faces.forEach((f) => used.add(f));
        report.matched++;
        report.split++;
        prev = { para: l.run.para, face: split.faces.at(-1)! };
        if (!lead.has(l.run.para)) lead.set(l.run.para, { face: split.faces[0], substitute: l.substitute });
        continue;
      }
      if (fixed) {
        l.run.texts.forEach((r, k) => edits.push({ start: r.start, end: r.end, text: escapeText(fixed[k]) }));
      }
      // Achado no PDF numa fonte que não temos: fica como o iLovePDF deixou.
      if (l.face === null) continue;
      let face = l.face;
      if (face) report.matched++;
      else if (blank && prev?.para === l.run.para) face = prev.face;
      else if ((face = votes.guess(l.substitute, l.bold, l.italic)) && !blank) report.fallback++;
      if (!face) continue;

      used.add(face);
      prev = { para: l.run.para, face };
      if (!blank && !lead.has(l.run.para)) lead.set(l.run.para, { face, substitute: l.substitute });

      // Espaçamento só quando sabemos o avanço real; em parágrafo justificado
      // o Word estica os espaços sozinho.
      const spacing =
        l.placement && !(blank && justified.has(l.run.para)) ? twips(aligner.extraAt(l.placement), l.sz) : undefined;
      edits.push(
        l.run.rPr
          ? { start: l.run.rPr.start, end: l.run.rPr.end, text: rewriteRPr(l.rPr, face, spacing ?? 0) }
          : { start: l.run.insertAt, end: l.run.insertAt, text: rewriteRPr('<w:rPr/>', face, spacing ?? 0) },
      );
    }

    for (const para of paragraphs) {
      const style = styleOf(para.id);
      const markRPr = para.markRPr ? xml.slice(para.markRPr.start, para.markRPr.end) : '';
      const markSubstitute = attr(markRPr, 'rFonts', 'ascii') ?? style.font;
      const markFace = votes.guess(
        markSubstitute,
        toggle(markRPr, 'b') ?? style.bold,
        toggle(markRPr, 'i') ?? style.italic,
      );
      const info = lead.get(para.id) ?? (markFace ? { face: markFace, substitute: markSubstitute } : undefined);
      if (!info) continue;
      used.add(info.face);
      // Marca de parágrafo: define a altura de linha vazia.
      if (para.markRPr) {
        edits.push({ start: para.markRPr.start, end: para.markRPr.end, text: rewriteRPr(markRPr, info.face) });
      }
      const subHeight = info.substitute ? SUBSTITUTE_LINE_HEIGHT[info.substitute] : undefined;
      if (subHeight && info.face.lineHeight) {
        const ratio = subHeight / info.face.lineHeight;
        const edit = Math.abs(ratio - 1) > 0.005 ? lineEdit(xml, para, style, ratio) : null;
        if (edit) edits.push(edit);
      }
    }
    if (edits.length) files[part] = strToU8(applyEdits(xml, edits));
  }

  // Estilos e numeração: rótulos de lista e o texto que o usuário digitar
  // depois já saem na fonte certa.
  for (const part of ['word/styles.xml', 'word/numbering.xml']) {
    if (!files[part]) continue;
    const xml = strFromU8(files[part] as Uint8Array);
    const edits: Edit[] = [];
    for (const m of xml.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/g)) {
      const face = votes.guess(attr(m[0], 'rFonts', 'ascii'), toggle(m[0], 'b'), toggle(m[0], 'i'));
      if (!face) continue;
      used.add(face);
      edits.push({ start: m.index!, end: m.index! + m[0].length, text: rewriteRPr(m[0], face) });
    }
    files[part] = strToU8(applyEdits(xml, edits));
  }

  await embedFonts(files, used, fonts.baseUrl);
  report.embedded = [...new Set([...used].filter((f) => f.file).map((f) => f.file!.replace(/\.\w+\.ttf\.br$/, '')))];
  return { docx: Buffer.from(zipSync(files, { level: 6 })), report };
};
