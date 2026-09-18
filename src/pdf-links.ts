import {
  PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNull, PDFNumber, PDFRef, PDFString,
  type PDFObject,
} from 'pdf-lib';

/**
 * Chromium emits /Dest /name links. Resolve those to explicit GoTo actions so
 * readers need neither legacy catalog destinations nor a name-tree lookup.
 * Run AFTER compression, which may rewrite annotation/destination objects.
 * Chromium's XYZ destinations with inherited zoom become FitH destinations:
 * fit the page width and place the section's y coordinate at the viewport top.
 * Explicit zooms, page references, hit areas and external actions stay intact.
 */
export async function normalizePdfLinks(input: Buffer): Promise<{
  buffer: Buffer;
  converted: number;
  unresolved: number;
}> {
  const pdf = await PDFDocument.load(input, { updateMetadata: false });
  const { context } = pdf;
  const name = PDFName.of;
  const lookup = (value: PDFObject | undefined): PDFObject | undefined =>
    value === undefined ? undefined : context.lookup(value);
  const keyOf = (value: PDFObject | undefined): string | undefined => {
    const object = lookup(value);
    return object instanceof PDFName || object instanceof PDFString || object instanceof PDFHexString
      ? object.decodeText() : undefined;
  };
  const destinations = new Map<string, PDFObject>();
  const legacy = lookup(pdf.catalog.get(name('Dests')));
  if (legacy instanceof PDFDict) {
    for (const [key, value] of legacy.entries()) destinations.set(key.decodeText(), value);
  }
  const visited = new Set<PDFDict>();
  const readNames = (value: PDFObject | undefined): void => {
    const node = lookup(value);
    if (!(node instanceof PDFDict) || visited.has(node)) return;
    visited.add(node);
    const entries = lookup(node.get(name('Names')));
    if (entries instanceof PDFArray) {
      for (let i = 0; i + 1 < entries.size(); i += 2) {
        const key = keyOf(entries.get(i));
        if (key !== undefined) destinations.set(key, entries.get(i + 1));
      }
    }
    const kids = lookup(node.get(name('Kids')));
    if (kids instanceof PDFArray) kids.asArray().forEach(readNames);
  };
  const names = lookup(pdf.catalog.get(name('Names')));
  if (names instanceof PDFDict) readNames(names.get(name('Dests')));

  const pages = pdf.getPages();
  const pageRefs = new Set(pages.map(page => page.ref.toString()));
  const resolve = (value: PDFObject | undefined): PDFArray | undefined => {
    const seen = new Set<PDFObject>();
    let current = lookup(value);
    while (current && !seen.has(current)) {
      seen.add(current);
      if (current instanceof PDFArray) {
        const page = current.size() >= 2 ? current.get(0) : undefined;
        return page instanceof PDFRef && pageRefs.has(page.toString()) ? current : undefined;
      }
      if (current instanceof PDFDict) current = lookup(current.get(name('D')));
      else {
        const key = keyOf(current);
        current = key === undefined ? undefined : lookup(destinations.get(key));
      }
    }
    return undefined;
  };

  let converted = 0;
  let unresolved = 0;
  for (const page of pages) {
    const annotations = lookup(page.node.get(name('Annots')));
    if (!(annotations instanceof PDFArray)) continue;
    for (const ref of annotations.asArray()) {
      const annotation = lookup(ref);
      if (!(annotation instanceof PDFDict) || keyOf(annotation.get(name('Subtype'))) !== 'Link') continue;
      const action = lookup(annotation.get(name('A')));
      // Never replace URI, GoToR, Launch or other external/non-navigation actions.
      if (action !== undefined && (!(action instanceof PDFDict) || keyOf(action.get(name('S'))) !== 'GoTo')) continue;
      const value = action instanceof PDFDict ? action.get(name('D')) : annotation.get(name('Dest'));
      if (value === undefined) continue;
      const destination = resolve(value);
      if (!destination) { unresolved++; continue; }
      let view = destination;
      if (destination.size() === 5 && keyOf(destination.get(1)) === 'XYZ') {
        const top = lookup(destination.get(3));
        const zoom = lookup(destination.get(4));
        if (top instanceof PDFNumber && (zoom === PDFNull || (zoom instanceof PDFNumber && zoom.asNumber() === 0))) {
          // Preserve the section position, including sections halfway down a
          // page. Do not mutate a destination shared by bookmarks/other links.
          view = context.obj([destination.get(0), 'FitH', top]);
        }
      }
      if (view === destination && action instanceof PDFDict && lookup(value) instanceof PDFArray && !annotation.has(name('Dest'))) continue;
      // Clone shared actions before changing them; preserve any additional keys.
      const explicit = action instanceof PDFDict ? action.clone(context) : context.obj({ S: 'GoTo' });
      explicit.set(name('D'), view);
      annotation.set(name('A'), explicit);
      annotation.delete(name('Dest')); // /Dest and /A are mutually exclusive.
      converted++;
    }
  }
  return {
    buffer: converted ? Buffer.from(await pdf.save({ useObjectStreams: false, addDefaultPage: false, updateFieldAppearances: false })) : input,
    converted,
    unresolved,
  };
}
