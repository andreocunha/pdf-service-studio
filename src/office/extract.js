/*
 * Roda DENTRO da página /render-pdf (mesma que gera o PDF), depois do
 * __PDF_READY. Lê o documento já diagramado e devolve uma árvore com a
 * geometria real de cada bloco — os montadores transformam isso em
 * Word (docx-native/build.ts) e PowerPoint (pptx-native/build.ts).
 *
 * JavaScript puro de propósito: é injetado como texto (page.evaluate), então
 * não pode depender de helpers que o bundler injeta.
 */
(() => {
  const TRANSPARENT = /^(transparent|rgba\([^)]*,\s*0\))$/;
  const INLINE = new Set(['inline', 'inline-block', 'contents']);
  let nextId = 0;

  const box = (el) => {
    const r = el.getBoundingClientRect();
    return { x: r.left + window.scrollX, y: r.top + window.scrollY, w: r.width, h: r.height };
  };

  /**
   * Escala de desenho do elemento: produto dos transform dos ancestrais (o
   * bloco do Figma é desenhado em 1920px com scale(0.55)). O getComputedStyle
   * dá os tamanhos antes da escala (fonte de 146px que aparece com 80px) —
   * todo comprimento do CSS passa por aqui; as caixas (getBoundingClientRect)
   * já vêm escaladas.
   */
  const scaleCache = new Map();
  const scaleOf = (el) => {
    if (!el || el === document.body || el === document.documentElement) return 1;
    if (scaleCache.has(el)) return scaleCache.get(el);
    let s = scaleOf(el.parentElement);
    const t = getComputedStyle(el).transform;
    if (t && t !== 'none') {
      try {
        const m = new DOMMatrixReadOnly(t);
        const k = Math.hypot(m.a, m.b);
        if (k > 0.01) s *= k;
      } catch {}
    }
    scaleCache.set(el, s);
    return s;
  };
  /** Girado ou espelhado por algum transform (seta do Figma reaproveitada de lado). */
  const turnCache = new Map();
  const turned = (el) => {
    if (!el || el === document.body || el === document.documentElement) return false;
    if (turnCache.has(el)) return turnCache.get(el);
    let v = turned(el.parentElement);
    const t = getComputedStyle(el).transform;
    if (!v && t && t !== 'none') {
      try {
        const m = new DOMMatrixReadOnly(t);
        v = Math.abs(m.b) > 1e-3 || Math.abs(m.c) > 1e-3 || m.a < 0 || m.d < 0;
      } catch {}
    }
    turnCache.set(el, v);
    return v;
  };
  /** Comprimento do CSS ("24px") como aparece na página. */
  const len = (el, v) => (parseFloat(v) || 0) * scaleOf(el);

  /** Área do conteúdo (sem padding e borda): onde o texto começa de verdade. */
  const contentBox = (el) => {
    const b = box(el);
    const cs = getComputedStyle(el);
    const px = (p) => len(el, cs[p]);
    const l = px('paddingLeft') + px('borderLeftWidth');
    const t = px('paddingTop') + px('borderTopWidth');
    const r = px('paddingRight') + px('borderRightWidth');
    const btm = px('paddingBottom') + px('borderBottomWidth');
    return { x: b.x + l, y: b.y + t, w: Math.max(0, b.w - l - r), h: Math.max(0, b.h - t - btm) };
  };

  const visible = (el, cs) =>
    cs.display !== 'none' && cs.visibility !== 'hidden' && Number(cs.opacity) > 0.01 && !el.hasAttribute('data-editor-only');

  /**
   * Cor do CSS como [_, r, g, b, alfa?] (strings, como o match do rgba()).
   * O Tailwind 4 escreve cores em lab()/oklch(), e o getComputedStyle devolve
   * assim mesmo: o canvas converte pra sRGB (o cinza do marcador saía preto).
   */
  const RGBA = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/;
  const colorCache = new Map();
  let paintCtx = null;
  const rgbaOf = (c) => {
    if (!c) return null;
    const m = c.match(RGBA);
    if (m) return m;
    if (colorCache.has(c)) return colorCache.get(c);
    let out = null;
    if (/^(lab|lch|oklab|oklch|hsla?|hwb|color)\(/.test(c)) {
      const a = c.match(/\/\s*([\d.]+)(%?)\s*\)$/);
      const alpha = a ? Number(a[1]) / (a[2] ? 100 : 1) : 1;
      if (!paintCtx) {
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = 1;
        paintCtx = canvas.getContext('2d', { willReadFrequently: true });
      }
      paintCtx.clearRect(0, 0, 1, 1);
      paintCtx.fillStyle = '#000';
      paintCtx.fillStyle = c.replace(/\s*\/\s*[\d.]+%?\s*\)$/, ')');
      paintCtx.fillRect(0, 0, 1, 1);
      const [r, g, b] = paintCtx.getImageData(0, 0, 1, 1).data;
      out = [c, String(r), String(g), String(b), alpha < 1 ? String(alpha) : undefined];
    }
    colorCache.set(c, out);
    return out;
  };

  const color = (c) => {
    const m = c && rgbaOf(c);
    if (!m || (m[4] !== undefined && Number(m[4]) < 0.05)) return null;
    return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('');
  };

  /**
   * Cor com a transparência de verdade (rgba × opacity do elemento e dos
   * pais): o PowerPoint tem alfa — a mistura com o que está atrás (blend)
   * fica pro Word. Só quando é transparente.
   */
  const translucent = (el, c) => {
    const m = c && rgbaOf(c);
    if (!m) return undefined;
    let alpha = m[4] === undefined ? 1 : Number(m[4]);
    for (let e = el; e && e !== document.body; e = e.parentElement) alpha *= Number(getComputedStyle(e).opacity);
    if (alpha >= 0.98 || alpha < 0.02) return undefined;
    return { color: [m[1], m[2], m[3]].map((v) => Math.round(Number(v)).toString(16).padStart(2, '0')).join(''), alpha: Math.round(alpha * 1000) / 1000 };
  };

  const border = (el, cs, side) => {
    const w = len(el, cs[`border${side}Width`]);
    const style = cs[`border${side}Style`];
    if (!w || style === 'none' || style === 'hidden') return null;
    const c = blend(el, cs[`border${side}Color`], true);
    if (!c) return null;
    const b = { w, color: c };
    const raw = translucent(el, cs[`border${side}Color`]);
    if (raw) b.raw = raw;
    if (style === 'dashed' || style === 'dotted') b.style = style;
    return b;
  };

  /** box-shadow por fora (o de dentro, "inset", é brilho de 1px — fica de fora). */
  const shadows = (el, cs) => {
    if (!cs.boxShadow || cs.boxShadow === 'none') return undefined;
    const list = cs.boxShadow
      .split(/,(?![^(]*\))/)
      .map((s) => s.trim())
      .filter((s) => !/\binset\b/.test(s))
      .map((s) => {
        const c = s.match(/(?:rgba?|lab|lch|oklab|oklch|hsla?|hwb|color)\([^)]*\)/);
        const nums = s.replace(/(?:rgba?|lab|lch|oklab|oklch|hsla?|hwb|color)\([^)]*\)/, '').trim().split(/\s+/).map(parseFloat);
        const raw = c && (translucent(el, c[0]) || { color: color(c[0]), alpha: 1 });
        if (!raw || !raw.color) return null;
        const [x = 0, y = 0, blur = 0, spread = 0] = nums.map((v) => v * scaleOf(el));
        return { x, y, blur, spread, color: raw.color, alpha: raw.alpha };
      })
      .filter(Boolean);
    return list.length ? list : undefined;
  };

  /**
   * O que está atrás do elemento: a primeira cor de fundo opaca dos pais
   * (com os fundos translúcidos do caminho por cima dela), ou null quando
   * antes dela vem uma imagem/gradiente (cor desconhecida).
   */
  const behind = (el, self = false) => {
    const layers = [];
    let base = [255, 255, 255];
    for (let e = self ? el : el.parentElement; e && e !== document.body; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.backgroundImage && cs.backgroundImage !== 'none') return null;
      const m = rgbaOf(cs.backgroundColor);
      if (!m) continue;
      const a = m[4] === undefined ? 1 : Number(m[4]);
      const rgb = [m[1], m[2], m[3]].map(Number);
      if (a >= 0.95) {
        base = rgb;
        break;
      }
      if (a > 0.02) layers.push([rgb, a]);
    }
    for (const [rgb, a] of layers.reverse()) base = base.map((b, i) => b + (rgb[i] - b) * a);
    return base;
  };

  /**
   * Cor como aparece: transparência (rgba ou opacity do elemento e dos pais)
   * misturada com o que está atrás — célula do Word não tem transparência.
   * Sobre imagem (cor de trás desconhecida) mistura com branco, menos o
   * vidro claro da capa: sem preenchimento (branco chapado escondia o texto
   * claro), e a borda dele fica na cor cheia — é o contorno do quadro.
   */
  const blend = (el, c, line = false, self = false) => {
    const m = c && rgbaOf(c);
    if (!m) return null;
    let alpha = m[4] === undefined ? 1 : Number(m[4]);
    for (let e = el; e && e !== document.body; e = e.parentElement) alpha *= Number(getComputedStyle(e).opacity);
    if (alpha < 0.05) return null;
    const rgb = [m[1], m[2], m[3]].map(Number);
    const hex = (c) => c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
    const mix = (back) => back.map((b, i) => b + (rgb[i] - b) * alpha);
    const back = alpha >= 0.95 ? [255, 255, 255] : behind(el, self);
    if (back) return hex(mix(back));
    // Atrás é imagem: vidro claro (branco translúcido) sairia branco chapado.
    if (alpha < 0.5 && rgb.every((v) => v > 235)) return line ? hex(rgb) : null;
    return hex(mix([255, 255, 255]));
  };
  const solidFill = (el, cs) => blend(el, cs.backgroundColor);

  const decoration = (el, cs) => {
    const d = {};
    const fill = solidFill(el, cs);
    if (fill) d.fill = fill;
    // Independe do `fill`: vidro claro sobre imagem não tem preenchimento no
    // Word, mas no PowerPoint tem (transparente). Só ele = `glass` (o Word tira).
    const fillRaw = translucent(el, cs.backgroundColor);
    if (fillRaw) d.fillRaw = fillRaw;
    const shadow = shadows(el, cs);
    if (shadow) d.shadow = shadow;
    if (cs.backgroundImage && cs.backgroundImage !== 'none') {
      const url = cs.backgroundImage.match(/url\("?([^")]+)"?\)/);
      if (url) d.bgImage = url[1];
      // Gradiente: a pintura sai como imagem atrás do texto (paint); a média
      // das cores fica de reserva.
      if (/gradient\(/.test(cs.backgroundImage)) d.paint = true;
      if (!d.fill && /gradient\(/.test(cs.backgroundImage)) {
        const stops = [...cs.backgroundImage.matchAll(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/g)]
          .filter((m) => m[4] === undefined || Number(m[4]) > 0.05)
          .map((m) => [Number(m[1]), Number(m[2]), Number(m[3])]);
        if (stops.length) {
          d.fill = [0, 1, 2]
            .map((i) => Math.round(stops.reduce((sum, c) => sum + c[i], 0) / stops.length).toString(16).padStart(2, '0'))
            .join('');
        }
      }
    }
    const b = {};
    for (const side of ['Top', 'Right', 'Bottom', 'Left']) {
      const v = border(el, cs, side);
      if (v) b[side.toLowerCase()] = v;
    }
    if (Object.keys(b).length) d.border = b;
    const pad = ['Top', 'Right', 'Bottom', 'Left'].map((s) => len(el, cs[`padding${s}`]));
    if (pad.some((v) => v > 0)) d.pad = pad;
    // Raio (o maior canto): círculo pequeno com ícone/texto vai como imagem redonda.
    const radius = Math.max(...['TopLeft', 'TopRight', 'BottomRight', 'BottomLeft'].map((c) => len(el, cs[`border${c}Radius`])));
    if (radius > 0.5) d.radius = radius;
    // Sombra: o Word ignora (o fundo vai no preenchimento da tabela); o PowerPoint desenha.
    return d;
  };

  /**
   * Cor do texto. Transparente (opacity, rgba) sobre fundo de cor conhecida:
   * a cor que aparece, misturada (Word não tem texto transparente e o PDF do
   * PowerPoint ignora). Sobre imagem: a cor com alfa.
   */
  const textColor = (el, cs) => {
    const t = translucent(el, cs.color);
    // O texto fica por cima do fundo do próprio elemento (não só dos pais).
    if (t && behind(el, true)) return { color: blend(el, cs.color, false, true) || color(cs.color) || '000000' };
    return { color: color(cs.color) || '000000', alpha: t?.alpha };
  };

  /** Fonte, peso, cor... de um trecho de texto. */
  const runStyle = (el) => {
    const cs = getComputedStyle(el);
    const family = cs.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
    const deco = cs.textDecorationLine || '';
    return {
      font: family,
      weight: Number(cs.fontWeight) || 400,
      italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique',
      size: len(el, cs.fontSize) || 16,
      ...textColor(el, cs),
      caps: cs.textTransform === 'uppercase',
      lower: cs.textTransform === 'lowercase',
      letterSpacing: len(el, cs.letterSpacing),
      underline: deco.includes('underline'),
      strike: deco.includes('line-through'),
      highlight: el.tagName === 'MARK' ? color(cs.backgroundColor) : null,
      verticalAlign: cs.verticalAlign === 'super' ? 'super' : cs.verticalAlign === 'sub' ? 'sub' : null,
    };
  };

  const paraStyle = (el) => {
    const cs = getComputedStyle(el);
    const lh = cs.lineHeight === 'normal' ? null : len(el, cs.lineHeight);
    return {
      align: { center: 'center', right: 'right', end: 'right', justify: 'both' }[cs.textAlign] || 'left',
      lineHeight: lh,
      fontSize: len(el, cs.fontSize) || 16,
    };
  };

  /**
   * Um bloco de texto rico (zona editável) vira parágrafos. div/p/li abrem
   * parágrafo novo, <br> é quebra de linha dentro do parágrafo.
   */
  const textLeaf = (el, link) => {
    const paras = [];
    let current = null;
    let autonumber = null;
    let autonumberStyle = null;
    let autonumberFormat = null;
    let autonumberText = null;
    // Onde o número aparece e o vão até o texto que vem depois dele (PowerPoint: o número vai como texto).
    let autonumberBox = null;
    let autonumberGap = null;
    // Só os números que aparecem (um escondido não conta).
    const multiNumber = [...el.querySelectorAll('[data-autonumber]')].filter((m) => !m.closest('[data-editor-only]') && (m.checkVisibility ? m.checkVisibility({ visibilityProperty: true, opacityProperty: true }) : true)).length > 1;
    const open = (source, list) => {
      current = { ...paraStyle(source), runs: [], list };
      // Item de lista: onde o texto do item começa, e o vão do marcador
      // (padding da lista) — vale também pro <p> dentro do <li>.
      const li = list && source.closest && source.closest('li');
      if (li && el.contains(li)) {
        const c = contentBox(li);
        current.x = c.x;
        current.indent = Math.max(0, c.x - box(li.closest('ul,ol') || li.parentElement).x);
        // Cor e tamanho do marcador (::marker), que o CSS costuma deixar mais discreto que o texto.
        const mk = getComputedStyle(li, '::marker');
        const mc = color(mk.color);
        if (mc) current.marker = { color: mc, alpha: translucent(li, mk.color)?.alpha, size: len(li, mk.fontSize) || undefined };
      }
      paras.push(current);
    };
    /**
     * Começo de cada linha nova dentro do trecho: letra cujo topo desce em
     * relação à anterior (no mesmo parágrafo, inclusive vindo do trecho de
     * antes). Só letras visíveis contam (o espaço da quebra fica no fim).
     */
    const lineBreaks = (node, text, rects) => {
      const out = [];
      // Linha da letra pela linha de base (a letra menor na mesma linha tem o topo mais baixo).
      const step = (rect, i) => {
        const base = rect.bottom - rect.height * 0.22;
        if (current.lineBase == null) {
          current.lineBase = base;
          current.lineH = rect.height;
          return;
        }
        if (base > current.lineBase + 0.45 * Math.min(rect.height, current.lineH)) {
          out.push(i);
          current.lineBase = base;
          current.lineH = rect.height;
        }
      };
      const first = text.search(/\S/);
      const solid = rects.filter((x) => x.width > 0);
      // Trecho numa linha só: basta ver se ele começou numa linha nova.
      if (new Set(solid.map((x) => Math.round(x.top))).size <= 1) {
        if (solid.length) step(solid[0], first);
        return out;
      }
      const r = document.createRange();
      for (let i = first; i < text.length; i++) {
        const code = text.charCodeAt(i);
        if (code >= 0xdc00 && code <= 0xdfff) continue;
        if (/\s/.test(text[i])) continue;
        const end = code >= 0xd800 && code <= 0xdbff ? i + 2 : i + 1;
        r.setStart(node, i);
        r.setEnd(node, Math.min(end, text.length));
        const rect = [...r.getClientRects()].find((x) => x.width > 0);
        if (rect) step(rect, i);
      }
      return out;
    };
    const walk = (node, list, linkRef) => {
      if (node.nodeType === 3) {
        const text = node.textContent;
        if (!text) return;
        if (!current) open(node.parentElement, list);
        // Largura real do trecho no Studio (soma das linhas): o builder compara
        // com a largura natural da fonte e grava a diferença como espaçamento.
        const range = document.createRange();
        range.selectNodeContents(node);
        const rects = [...range.getClientRects()];
        const width = rects.reduce((sum, r) => sum + r.width, 0);
        // Linhas do trecho: o espaço onde a linha quebrou não entra na largura.
        const lines = new Set(rects.map((r) => Math.round(r.top))).size;
        // Meio da primeira e da última linha do parágrafo: dá o espaço entre parágrafos (margem do <ul>, do <p>).
        const mid = (r) => r.top + r.height / 2 + window.scrollY;
        const solid = rects.filter((r) => r.width > 0);
        if (solid.length && text.trim()) {
          if (current.first === undefined) {
            current.first = mid(solid[0]);
            // Altura da área do texto (ascendente + descendente da fonte): com
            // entrelinha apertada o PowerPoint sobe a linha, e isso corrige.
            current.area = solid[0].height;
          }
          current.last = mid(solid[solid.length - 1]);
        }
        // Onde a linha quebrou no Studio (índices do texto): o PowerPoint quebra no mesmo lugar.
        const breaks = text.trim() ? lineBreaks(node, text, rects) : [];
        if (autonumberBox && autonumberGap === null && text.trim() && rects.length) {
          const first = rects[0];
          const top = first.top + window.scrollY;
          if (Math.abs(top + first.height / 2 - (autonumberBox.y + autonumberBox.h / 2)) < first.height) {
            autonumberGap = Math.max(0, first.left + window.scrollX - (autonumberBox.x + autonumberBox.w));
          }
        }
        // Espaço em branco que o navegador junta (white-space normal): o PowerPoint precisa saber.
        const collapse = /^(normal|nowrap)$/.test(getComputedStyle(node.parentElement).whiteSpace) || undefined;
        current.runs.push({ t: text.replace(/\n/g, ' '), ...runStyle(node.parentElement), link: linkRef || null, width, lines, collapse, breaks: breaks.length ? breaks : undefined });
        return;
      }
      if (node.nodeType !== 1) return;
      const cs = getComputedStyle(node);
      if (!visible(node, cs)) return;
      if (node.matches('a[data-zone-link]')) return; // sobreposição de clique, sem texto
      if (node.hasAttribute('data-autonumber')) {
        // Número no meio do parágrafo ("Capítulo II"): a lista do Word só põe
        // número no começo, então vai como texto, no lugar dele.
        // Vários números no mesmo bloco ("i) A … ii) B"): cada um no seu lugar, como texto.
        if (multiNumber || (current && current.runs.some((r) => r.t && r.t.trim()))) {
          const text = node.textContent;
          if (text && !current) open(node.parentElement, list);
          if (text) current.runs.push({ t: text, ...runStyle(node), link: linkRef || null, width: node.getBoundingClientRect().width, lines: 1 });
          return;
        }
        // Com estilo ("c", "IV"), o caminho real vem no atributo; o texto é o que aparece.
        autonumber = node.getAttribute('data-autonumber-path') || node.textContent.trim();
        autonumberFormat = node.getAttribute('data-autonumber-path') ? (node.getAttribute('data-autonumber') || '').trim() : null;
        autonumberText = node.textContent.trim();
        // O número vira numeração do Word, que herda a formatação da marca de
        // parágrafo: guarda como ele aparece (cor, tamanho, peso).
        autonumberStyle = { ...runStyle(node), lineHeight: paraStyle(node).lineHeight };
        autonumberBox = box(node);
        return;
      }
      if (node.tagName === 'BR') {
        if (!current) open(node.parentElement, list);
        current.lineBase = null;
        current.runs.push({ br: true });
        return;
      }
      const href = node.tagName === 'A' ? node.getAttribute('data-block-ref') || node.getAttribute('href') : null;
      if (node.tagName === 'UL' || node.tagName === 'OL') {
        const depth = list ? list.level + 1 : 0;
        const kind = node.tagName === 'OL' ? cs.listStyleType || 'decimal' : 'bullet';
        [...node.childNodes].forEach((c) => walk(c, { kind, level: depth }, linkRef));
        current = null;
        return;
      }
      // Filho de linha flex fica na mesma linha (o "º" de "Contrato nº" num
      // <span> dentro de um flex): só abre parágrafo se o pai empilha.
      const pcs = node.parentElement ? getComputedStyle(node.parentElement) : null;
      const inFlexRow = pcs && /flex/.test(pcs.display) && !/column/.test(pcs.flexDirection) && pcs.flexWrap === 'nowrap';
      const block = !INLINE.has(cs.display) && !inFlexRow;
      if (block) {
        current = null;
        if (node.tagName === 'LI' || cs.display === 'list-item') {
          open(node, list || { kind: 'bullet', level: 0 });
        }
      }
      [...node.childNodes].forEach((c) => walk(c, list, href || linkRef));
      if (block) current = null;
    };
    [...el.childNodes].forEach((c) => walk(c, null, link));
    const keep = paras.filter((p) => p.runs.length);
    for (const p of keep) {
      delete p.lineBase;
      delete p.lineH;
    }
    if (!keep.length && !autonumber) return null;
    // Texto some da captura de um bloco complexo (ele vai editável por cima).
    el.setAttribute('data-dx-text', '');
    // Marcador de lista pendurado fora do texto: a caixa inclui o espaço dele.
    const tb = contentBox(el);
    const markers = keep.filter((p) => p.x !== undefined).map((p) => p.x - (p.indent >= 4 ? p.indent : p.fontSize * 1.1));
    const left = Math.min(tb.x, ...markers);
    if (left < tb.x) { tb.w += tb.x - left; tb.x = left; }
    // Onde as linhas terminam de verdade (texto contornando um ícone numa caixa larga).
    const range = document.createRange();
    range.selectNodeContents(el);
    const ink = [...range.getClientRects()].filter((r) => r.width > 0);
    const inkRight = ink.length ? Math.max(...ink.map((r) => r.right + window.scrollX)) : null;
    // Texto centralizado na vertical dentro da própria caixa (rótulo flex
    // "align-items: center" da altura da linha da tabela): no Word a célula
    // centraliza — a caixa sozinha diz "no topo".
    const ecs = getComputedStyle(el);
    const centers = (/flex/.test(ecs.display) && (/column/.test(ecs.flexDirection) ? ecs.justifyContent === 'center' : ecs.alignItems === 'center'))
      || (ecs.display === 'table-cell' && ecs.verticalAlign === 'middle');
    const inkTop = ink.length ? Math.min(...ink.map((r) => r.top + window.scrollY)) : 0;
    const inkBottom = ink.length ? Math.max(...ink.map((r) => r.bottom + window.scrollY)) : 0;
    const middle = centers && ink.length && inkTop - tb.y > 3 && Math.abs((inkTop - tb.y) - (tb.y + tb.h - inkBottom)) < 3 ? true : undefined;
    return { k: 'text', id: nextId++, box: tb, inkRight, middle, paras: keep.length ? keep : [{ ...paraStyle(el), runs: [] }], autonumber, autonumberStyle, autonumberFormat, autonumberText, autonumberBox, autonumberGap };
  };

  const hasBlockChildren = (el) =>
    [...el.children].some((c) => {
      const cs = getComputedStyle(c);
      return visible(c, cs) && !INLINE.has(cs.display) && c.tagName !== 'BR' && !c.matches('a[data-zone-link]');
    });

  const ownText = (el) => [...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim());

  /** Link interno da zona: <a data-zone-link data-block-ref> sobreposto a ela. */
  const zoneLink = (el) => {
    const a = [...el.children].find((c) => c.matches && c.matches('a[data-zone-link]'));
    return a ? a.getAttribute('data-block-ref') || a.getAttribute('href') : null;
  };

  /** Filhos; display: contents (agrupamento do Figma) não tem caixa — entram os de dentro. */
  const kidsOf = (el, link) =>
    [...el.children].flatMap((c) => (getComputedStyle(c).display === 'contents' ? kidsOf(c, link) : [node(c, link)])).filter(Boolean);

  const node = (el, link) => {
    const cs = getComputedStyle(el);
    if (!visible(el, cs)) return null;
    const b = box(el);
    if (b.w < 0.5 || b.h < 0.5) return null;
    const tag = el.tagName;
    const abs = cs.position === 'absolute' || cs.position === 'fixed';
    if (tag === 'IMG') {
      el.setAttribute('data-dx', String(nextId));
      const r = Math.max(...['TopLeft', 'TopRight', 'BottomRight', 'BottomLeft'].map((c) => len(el, cs[`border${c}Radius`])));
      const radius = r > 0.5 ? r : undefined;
      // Girada/espelhada: o arquivo está de outro jeito — só a captura mostra igual.
      if (turned(el)) return { k: 'raster', id: nextId++, box: b, abs };
      // object-fit: a imagem desenhada não é a caixa. "contain" encolhe
      // mantendo a proporção (o Word esticaria até a caixa); "cover" recorta
      // — só a captura mostra igual.
      const fit = cs.objectFit;
      const nw = el.naturalWidth;
      const nh = el.naturalHeight;
      if (nw && nh && Math.abs(nw / nh - b.w / b.h) > 0.02) {
        if (fit === 'cover') return { k: 'raster', id: nextId++, box: b, abs };
        if (fit === 'contain' || fit === 'scale-down') {
          const k = Math.min(b.w / nw, b.h / nh, fit === 'scale-down' ? 1 : Infinity);
          const [px, py] = cs.objectPosition.split(' ').map((v) => (v.endsWith('%') ? parseFloat(v) / 100 : null));
          const w = nw * k;
          const h = nh * k;
          const drawn = { x: b.x + (b.w - w) * (px ?? 0.5), y: b.y + (b.h - h) * (py ?? 0.5), w, h };
          return { k: 'img', id: nextId++, box: drawn, src: el.currentSrc || el.src, abs, radius };
        }
      }
      return { k: 'img', id: nextId++, box: b, src: el.currentSrc || el.src, abs, radius };
    }
    // Ícone desenhado com máscara CSS (cor recortada no formato do ícone):
    // só a captura mostra o desenho — a caixa em si é só a cor.
    const mask = cs.maskImage || cs.webkitMaskImage;
    if (tag === 'svg' || tag === 'SVG' || tag === 'CANVAS' || tag === 'VIDEO' || (mask && mask !== 'none')) {
      el.setAttribute('data-dx', String(nextId));
      return { k: 'raster', id: nextId++, box: b, abs };
    }
    // Bloco aninhado (colunas): desce no conteúdo dele.
    if (el !== root && el.hasAttribute('data-block-id')) return blockContent(el);

    // Aba do menu convertida pro PDF: <a data-block-ref> envolvendo o botão.
    if (tag === 'A' && el.hasAttribute('data-block-ref')) link = el.getAttribute('data-block-ref');
    link = zoneLink(el) || link;
    const deco = decoration(el, cs);
    const positioned = cs.position === 'absolute' || cs.position === 'fixed';
    const isText =
      el.hasAttribute('data-editable') || (ownText(el) && !hasBlockChildren(el)) ||
      (!hasBlockChildren(el) && el.textContent.trim() && ![...el.querySelectorAll('img,svg')].length);
    if (isText) {
      const t = textLeaf(el, link);
      if (t && (deco.fill || deco.border || deco.pad || deco.fillRaw)) {
        if (!deco.fill && !deco.border && !deco.pad) deco.glass = true;
        return { k: 'box', id: mark(el, deco), box: b, ...deco, abs: positioned, layout: layoutOf(cs), kids: [t] };
      }
      if (t) t.abs = positioned;
      return t;
    }
    // Caixa que recorta (overflow hidden) com enfeite solto passando da
    // borda: o enfeite fica, cortado no limite da caixa (vira célula com
    // preenchimento no Word, em vez de imagem de fundo).
    const clips = cs.overflow !== 'visible';
    const clip = (n) => {
      const x0 = Math.max(n.box.x, b.x);
      const y0 = Math.max(n.box.y, b.y);
      const x1 = Math.min(n.box.x + n.box.w, b.x + b.w);
      const y1 = Math.min(n.box.y + n.box.h, b.y + b.h);
      if (x1 - x0 < 0.5 || y1 - y0 < 0.5) return null;
      if (n.k === 'text') return n;
      // Imagem que passa da borda: o pedaço de dentro vai em `clip` (o
      // PowerPoint captura só ele — o brilho do card não vaza; o Word monta
      // a tabela pela caixa inteira).
      if (n.k === 'img' || n.k === 'raster') return clipMedia(n);
      const out = { ...n, box: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } };
      if (n.kids) out.kids = n.kids.map(clip).filter(Boolean);
      return out;
    };
    // Só o `clip` das imagens (o que o PowerPoint mostra); a caixa delas fica.
    // Fora da área: clip vazio (o PowerPoint pula; o Word segue como era).
    const clipMedia = (n) => {
      if (n.k === 'img' || n.k === 'raster') {
        const r = n.clip || n.box;
        const x0 = Math.max(r.x, b.x);
        const y0 = Math.max(r.y, b.y);
        const x1 = Math.min(r.x + r.w, b.x + b.w);
        const y1 = Math.min(r.y + r.h, b.y + b.h);
        const c = { x: x0, y: y0, w: Math.max(0, x1 - x0), h: Math.max(0, y1 - y0) };
        const same = Math.abs(c.x - n.box.x) < 0.5 && Math.abs(c.y - n.box.y) < 0.5 && Math.abs(c.w - n.box.w) < 0.5 && Math.abs(c.h - n.box.h) < 0.5;
        return same ? n : { ...n, clip: c };
      }
      return n.kids ? { ...n, kids: n.kids.map(clipMedia) } : n;
    };
    let kids = kidsOf(el, link);
    if (clips && (deco.fill || deco.border)) kids = kids.map((k) => (k.abs && k.k !== 'text' ? clip(k) : k)).filter(Boolean);
    // Grupo que recorta sem cor própria ("Mask group" do Figma): só as imagens.
    else if (clips) kids = kids.map(clipMedia);
    if (!kids.length) {
      // Caixa vazia com cor = forma (barra, linha divisória, quadrado).
      if (deco.fill || deco.border || deco.bgImage) return { k: 'shape', id: mark(el, deco), box: b, ...deco, abs: positioned };
      if (deco.fillRaw) return { k: 'shape', id: mark(el, deco), box: b, ...deco, glass: true, abs: positioned };
      return null;
    }
    return { k: 'box', id: mark(el, deco), box: b, ...deco, abs: positioned, layout: layoutOf(cs), kids };
  };

  /** Id do nó; caixa que vai precisar de captura da pintura fica marcada no DOM. */
  const mark = (el, deco) => {
    const id = nextId++;
    if (deco.bgImage) deco.paint = true;
    // Toda caixa fica achável: a pintura dela, ou o bloco complexo inteiro, pode virar captura.
    el.setAttribute('data-dx', String(id));
    return id;
  };

  const layoutOf = (cs) => ({
    display: cs.display,
    dir: cs.flexDirection,
    alignItems: cs.alignItems,
    justify: cs.justifyContent,
    textAlign: cs.textAlign,
  });

  /** Coluna de conteúdo do bloco (sem as alças do editor dos lados). */
  let root = null;
  const blockContent = (blockEl) => {
    const cols = [...blockEl.children].filter((c) => !c.querySelector(':scope > [data-editor-only]') && visible(c, getComputedStyle(c)));
    const main = cols.sort((a, b) => box(b).w * box(b).h - box(a).w * box(a).h)[0];
    if (!main) return null;
    const prev = root;
    root = blockEl;
    const n = node(main, null);
    root = prev;
    return n;
  };

  const pages = [...document.querySelectorAll('[data-page-index]')];
  const topBlocks = [...document.querySelectorAll('[data-block-id]')].filter((el) => !el.parentElement.closest('[data-block-id]'));
  const meta = window.__PDF_META || {};
  return {
    meta,
    pages: pages.map((p) => {
      const cs = getComputedStyle(p);
      const url = cs.backgroundImage.match(/url\("?([^")]+)"?\)/);
      // Rodapé do Studio (texto + "02/28"): vira rodapé do Word com campo de página.
      const footerEl = p.querySelector('[data-document-footer]');
      let footer = null;
      if (footerEl) {
        const textEl = footerEl.querySelector('[data-document-footer-text]');
        const numberEl = footerEl.querySelector('[data-document-footer-number]');
        footer = {
          text: textEl ? textLeaf(textEl, null) : null,
          number: numberEl ? { text: numberEl.textContent.trim(), format: numberEl.getAttribute('data-document-footer-format') || 'page-total', box: box(numberEl), style: runStyle(numberEl), align: getComputedStyle(numberEl).textAlign } : null,
        };
      }
      // Menu de seções: vai pro cabeçalho do Word; a assinatura (aba ativa
      // incluída) decide onde começa seção nova.
      const navEl = p.querySelector('[data-section-nav]');
      let nav = null;
      if (navEl) {
        const prev = root;
        root = navEl;
        const tree = node(navEl, null);
        root = prev;
        const sig = [...navEl.querySelectorAll('a[data-block-ref]')].map((a) => a.getAttribute('data-block-ref') + ':' + a.innerHTML.length + ':' + (a.innerHTML.match(/background(?:-color)?:\s*([^;"]+)/) || [])[1]).join('|');
        nav = tree ? { tree, sig } : null;
      }
      return { index: Number(p.getAttribute('data-page-index')), box: box(p), bgImage: url ? url[1] : null, bgColor: color(cs.backgroundColor), footer, nav };
    }),
    blocks: topBlocks.map((el) => {
      root = el;
      const pageEl = el.closest('[data-page-index]');
      return {
        blockId: el.getAttribute('data-block-id'),
        type: el.getAttribute('data-block-type'),
        page: pageEl ? Number(pageEl.getAttribute('data-page-index')) : 0,
        box: box(el),
        tree: blockContent(el),
      };
    }),
  };
})()
