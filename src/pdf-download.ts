import { logger } from './logger.js';
import { normalizePdfLinks } from './pdf-links.js';

/** Link compatibility must never prevent downloading an already rendered PDF. */
export async function finalizePdfDownload(pdf: Buffer, documentId: string): Promise<Buffer> {
  try {
    const links = await normalizePdfLinks(pdf);
    logger.info({ documentId, converted: links.converted, unresolved: links.unresolved }, 'PDF internal links normalized');
    return links.buffer;
  } catch (err) {
    logger.warn({ documentId, err }, 'PDF link normalization failed — sending original');
    return pdf;
  }
}
