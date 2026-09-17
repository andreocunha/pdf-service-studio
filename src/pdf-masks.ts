import type { Page } from 'puppeteer-core';

/**
 * Skia emits layered CSS gradient masks as nested PDF tiling patterns. Quartz
 * (Preview/iOS) can repeat the holes through the content and clip whole cards.
 * Convert hard-edged gradient masks to vector clips; keep text and links intact.
 * Soft fades and URL masks are deliberately left unchanged.
 * Run after the final viewport/layout settle, immediately before printing.
 */
export async function preparePdfMasks(page: Page): Promise<number> {
  return page.evaluate(async () => {
    const properties = [
      'mask-image', 'mask-size', 'mask-position', 'mask-repeat',
      'mask-origin', 'mask-clip', 'mask-composite', 'mask-mode',
      'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
      'border-top-width', 'border-right-width', 'border-bottom-width', 'border-left-width',
      'border-top-left-radius', 'border-top-right-radius',
      'border-bottom-left-radius', 'border-bottom-right-radius',
    ];
    let converted = 0;
    elements: for (const element of Array.from(document.querySelectorAll<HTMLElement>('*'))) {
      const style = getComputedStyle(element);
      const mask = style.maskImage;
      // URL/SVG masks have different loading/CORS semantics. Leave them alone;
      // this workaround targets CSS-generated gradient tiling patterns.
      if (!mask.includes('gradient(') || mask.includes('url(') || style.clipPath !== 'none') continue;
      if (!(element instanceof HTMLElement) || !element.offsetWidth || !element.offsetHeight) continue;
      // Computed dimensions retain fractional CSS pixels and ignore transforms.
      let width = parseFloat(style.width), height = parseFloat(style.height);
      if (style.boxSizing !== 'border-box') {
        width += parseFloat(style.paddingLeft) + parseFloat(style.paddingRight)
          + parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth);
        height += parseFloat(style.paddingTop) + parseFloat(style.paddingBottom)
          + parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth);
      }
      if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) continue;

      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('width', String(width));
      svg.setAttribute('height', String(height));
      const foreign = document.createElementNS(svg.namespaceURI, 'foreignObject');
      foreign.setAttribute('width', '100%');
      foreign.setAttribute('height', '100%');
      const shape = document.createElementNS('http://www.w3.org/1999/xhtml', 'div');
      shape.style.cssText = `box-sizing:border-box;width:${width}px;height:${height}px;background:white;border-style:solid;border-color:white;`;
      for (const property of properties) {
        shape.style.setProperty(property, style.getPropertyValue(property));
      }
      foreign.append(shape);
      svg.append(foreign);
      const image = new Image();
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(new XMLSerializer().serializeToString(svg))}`;
      await image.decode();

      // 288 dpi at normal print size, bounded for large continuous documents.
      const scale = Math.min(3, Math.sqrt(16_000_000 / (width * height)), 8192 / Math.max(width, height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.ceil(width * scale));
      canvas.height = Math.max(1, Math.ceil(height * scale));
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Unable to prepare PDF gradient mask');
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      // A hard edge has at most a narrow anti-aliasing band near either
      // opaque or transparent pixels (tiny clipped holes may never reach 0).
      // Do not turn a deliberate translucent/feathered mask into a binary silhouette.
      const edgeRadius = Math.max(1, Math.ceil(scale));
      for (let y = 0; y < canvas.height; y++) {
        for (let x = 0; x < canvas.width; x++) {
          const alpha = pixels[(y * canvas.width + x) * 4 + 3];
          if (alpha <= 8 || alpha >= 247) continue;
          let transparent = false, opaque = false;
          for (let dy = -edgeRadius; dy <= edgeRadius && !(transparent && opaque); dy++) {
            for (let dx = -edgeRadius; dx <= edgeRadius; dx++) {
              const nx = x + dx, ny = y + dy;
              if (nx < 0 || ny < 0 || nx >= canvas.width || ny >= canvas.height) continue;
              const neighbor = pixels[(ny * canvas.width + nx) * 4 + 3];
              transparent ||= neighbor <= 8;
              opaque ||= neighbor >= 247;
            }
          }
          if (!transparent && !opaque) {
            canvas.width = canvas.height = 0;
            continue elements;
          }
        }
      }

      // Union of opaque horizontal runs, coalesced vertically. This avoids
      // PDF soft masks entirely (even a single PNG mask is broken in Quartz).
      // Adjacent rectangles share exact coordinates: no transparent seams.
      type Run = { left: number; right: number; top: number; bottom: number };
      let active = new Map<string, Run>();
      const rectangles: Run[] = [];
      for (let y = 0; y < canvas.height; y++) {
        const next = new Map<string, Run>();
        for (let x = 0; x < canvas.width;) {
          if (pixels[(y * canvas.width + x) * 4 + 3] < 128) { x++; continue; }
          const left = x++;
          while (x < canvas.width && pixels[(y * canvas.width + x) * 4 + 3] >= 128) x++;
          const key = `${left}:${x}`;
          const run = active.get(key) ?? { left, right: x, top: y, bottom: y };
          run.bottom = y + 1;
          // Bound path/memory growth for dense repeating patterns.
          if (rectangles.length + next.size > 20_000) {
            canvas.width = canvas.height = 0;
            continue elements;
          }
          next.set(key, run);
          active.delete(key);
        }
        rectangles.push(...active.values());
        active = next;
      }
      rectangles.push(...active.values());
      const sx = width / canvas.width, sy = height / canvas.height;
      const path = rectangles.map(r =>
        `M${Math.round(r.left * sx * 10000) / 10000} ${Math.round(r.top * sy * 10000) / 10000}H${Math.round(r.right * sx * 10000) / 10000}V${Math.round(r.bottom * sy * 10000) / 10000}H${Math.round(r.left * sx * 10000) / 10000}Z`,
      ).join('');
      if (!path) { canvas.width = canvas.height = 0; continue; }
      element.style.setProperty('clip-path', `path("${path}")`, 'important');
      element.style.setProperty('mask', 'none', 'important');
      converted++;
      canvas.width = canvas.height = 0;
    }
    if (converted) {
      await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    }
    return converted;
  });
}
