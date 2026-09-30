/*
 * Roda DENTRO da página /render-pdf (mesma que gera o PDF), depois do
 * __PDF_READY. Lê o documento já diagramado e devolve uma árvore com a
 * geometria real de cada bloco — o builder (build.ts) transforma isso em
 * Word seguindo o guia da Lex (tabelas, nada flutuante, numeração automática).
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

  /** Área do conteúdo (sem padding e borda): onde o texto começa de verdade. */
  const contentBox = (el) => {
    const b = box(el);
    const cs = getComputedStyle(el);
    const px = (p) => parseFloat(cs[p]) || 0;
    const l = px('paddingLeft') + px('borderLeftWidth');
    const t = px('paddingTop') + px('borderTopWidth');
    const r = px('paddingRight') + px('borderRightWidth');
    const btm = px('paddingBottom') + px('borderBottomWidth');
    return { x: b.x + l, y: b.y + t, w: Math.max(0, b.w - l - r), h: Math.max(0, b.h - t - btm) };
  };

  const visible = (el, cs) =>
    cs.display !== 'none' && cs.visibility !== 'hidden' && Number(cs.opacity) > 0.01 && !el.hasAttribute('data-editor-only');

  const color = (c) => {
    const m = c && c.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
    if (!m || (m[4] !== undefined && Number(m[4]) < 0.05)) return null;
    return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('');
  };

  const border = (cs, side) => {
    const w = parseFloat(cs[`border${side}Width`]) || 0;
    const style = cs[`border${side}Style`];
    if (!w || style === 'none' || style === 'hidden') return null;
    const c = color(cs[`border${side}Color`]);
    return c ? { w, color: c } : null;
  };

  /**
   * Cor de fundo como aparece: transparência (rgba ou opacity do elemento e
   * dos pais) misturada com branco — célula do Word não tem transparência.
   */
  const solidFill = (el, cs) => {
    const m = cs.backgroundColor.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
    if (!m) return null;
    let alpha = m[4] === undefined ? 1 : Number(m[4]);
    for (let e = el; e && e !== document.body; e = e.parentElement) alpha *= Number(getComputedStyle(e).opacity);
    if (alpha < 0.05) return null;
    return [m[1], m[2], m[3]].map((v) => Math.round(255 - (255 - Number(v)) * alpha).toString(16).padStart(2, '0')).join('');
  };

  const decoration = (el, cs) => {
    const d = {};
    const fill = solidFill(el, cs);
    if (fill) d.fill = fill;
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
      const v = border(cs, side);
      if (v) b[side.toLowerCase()] = v;
    }
    if (Object.keys(b).length) d.border = b;
    const pad = ['Top', 'Right', 'Bottom', 'Left'].map((s) => parseFloat(cs[`padding${s}`]) || 0);
    if (pad.some((v) => v > 0)) d.pad = pad;
    // Sombra é ignorada: o fundo vai no preenchimento da tabela (padrão da equipe).
    return d;
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
      size: parseFloat(cs.fontSize) || 16,
      color: color(cs.color) || '000000',
      caps: cs.textTransform === 'uppercase',
      lower: cs.textTransform === 'lowercase',
      letterSpacing: parseFloat(cs.letterSpacing) || 0,
      underline: deco.includes('underline'),
      strike: deco.includes('line-through'),
      highlight: el.tagName === 'MARK' ? color(cs.backgroundColor) : null,
      verticalAlign: cs.verticalAlign === 'super' ? 'super' : cs.verticalAlign === 'sub' ? 'sub' : null,
    };
  };

  const paraStyle = (el) => {
    const cs = getComputedStyle(el);
    const lh = cs.lineHeight === 'normal' ? null : parseFloat(cs.lineHeight);
    return {
      align: { center: 'center', right: 'right', end: 'right', justify: 'both' }[cs.textAlign] || 'left',
      lineHeight: lh,
      fontSize: parseFloat(cs.fontSize) || 16,
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
    const open = (source, list) => {
      current = { ...paraStyle(source), runs: [], list };
      // Item de lista: onde o texto do item começa, e o vão do marcador
      // (padding da lista) — vale também pro <p> dentro do <li>.
      const li = list && source.closest && source.closest('li');
      if (li && el.contains(li)) {
        const c = contentBox(li);
        current.x = c.x;
        current.indent = Math.max(0, c.x - box(li.closest('ul,ol') || li.parentElement).x);
      }
      paras.push(current);
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
        current.runs.push({ t: text.replace(/\n/g, ' '), ...runStyle(node.parentElement), link: linkRef || null, width, lines });
        return;
      }
      if (node.nodeType !== 1) return;
      const cs = getComputedStyle(node);
      if (!visible(node, cs)) return;
      if (node.matches('a[data-zone-link]')) return; // sobreposição de clique, sem texto
      if (node.hasAttribute('data-autonumber')) {
        autonumber = node.textContent.trim();
        // O número vira numeração do Word, que herda a formatação da marca de
        // parágrafo: guarda como ele aparece (cor, tamanho, peso).
        autonumberStyle = { ...runStyle(node), lineHeight: paraStyle(node).lineHeight };
        return;
      }
      if (node.tagName === 'BR') {
        if (!current) open(node.parentElement, list);
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
      const block = !INLINE.has(cs.display);
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
    return { k: 'text', id: nextId++, box: tb, inkRight, paras: keep.length ? keep : [{ ...paraStyle(el), runs: [] }], autonumber, autonumberStyle };
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

  const node = (el, link) => {
    const cs = getComputedStyle(el);
    if (!visible(el, cs)) return null;
    const b = box(el);
    if (b.w < 0.5 || b.h < 0.5) return null;
    const tag = el.tagName;
    const abs = cs.position === 'absolute' || cs.position === 'fixed';
    if (tag === 'IMG') {
      el.setAttribute('data-dx', String(nextId));
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
          return { k: 'img', id: nextId++, box: drawn, src: el.currentSrc || el.src, abs };
        }
      }
      return { k: 'img', id: nextId++, box: b, src: el.currentSrc || el.src, abs };
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
      if (t && (deco.fill || deco.border || deco.pad)) {
        return { k: 'box', id: mark(el, deco), box: b, ...deco, abs: positioned, layout: layoutOf(cs), kids: [t] };
      }
      if (t) t.abs = positioned;
      return t;
    }
    // Enfeite solto (posição absoluta, sem texto) dentro de caixa que recorta
    // ou tem pintura própria — os círculos no canto de um banner: entra na
    // captura da pintura da caixa, recortado como no Studio, em vez de virar
    // item à parte (que vazaria da caixa no Word).
    // Só o que a caixa recorta de verdade (passa da borda dela): um logo solto
    // dentro de um fundo branco é imagem própria, não parte da pintura.
    const clips = cs.overflow !== 'visible';
    const outside = (c) => {
      const r = box(c);
      return r.x < b.x - 1 || r.y < b.y - 1 || r.x + r.w > b.x + b.w + 1 || r.y + r.h > b.y + b.h + 1;
    };
    const baked = clips && (deco.fill || deco.border || deco.bgImage || deco.paint)
      ? [...el.children].filter((c) => {
          const ccs = getComputedStyle(c);
          return (ccs.position === 'absolute' || ccs.position === 'fixed') && !c.textContent.trim() && visible(c, ccs) && outside(c);
        })
      : [];
    if (baked.length) {
      deco.paint = true;
      for (const c of baked) c.setAttribute('data-dx-baked', '');
    }
    const kids = [...el.children].filter((c) => !baked.includes(c)).map((c) => node(c, link)).filter(Boolean);
    if (!kids.length) {
      // Caixa vazia com cor = forma (barra, linha divisória, quadrado).
      if (deco.fill || deco.border || deco.bgImage) return { k: 'shape', id: mark(el, deco), box: b, ...deco, abs: positioned };
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
          number: numberEl ? { text: numberEl.textContent.trim(), box: box(numberEl), style: runStyle(numberEl), align: getComputedStyle(numberEl).textAlign } : null,
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
