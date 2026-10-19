/**
 * TTF → EOT (Embedded OpenType), o formato das fontes embutidas no
 * PowerPoint (`ppt/fonts/*.fntdata`). O Word usa outro (odttf: TTF com XOR).
 *
 * EOT versão 0x00020001 sem compressão nem XOR: cabeçalho com os dados de
 * OS/2, head e name da fonte, seguido do TTF inteiro.
 * Referência: https://www.w3.org/submissions/EOT/
 */

type Table = { offset: number; length: number };

const tables = (ttf: Buffer): Map<string, Table> => {
  const out = new Map<string, Table>();
  const count = ttf.readUInt16BE(4);
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    out.set(ttf.toString('latin1', at, at + 4), { offset: ttf.readUInt32BE(at + 8), length: ttf.readUInt32BE(at + 12) });
  }
  return out;
};

/** Nome da tabela name (Windows, Unicode) em UTF-16LE. */
const nameOf = (ttf: Buffer, name: Table | undefined, id: number): Buffer => {
  if (!name) return Buffer.alloc(0);
  const count = ttf.readUInt16BE(name.offset + 2);
  const strings = name.offset + ttf.readUInt16BE(name.offset + 4);
  let fallback: Buffer | null = null;
  for (let i = 0; i < count; i++) {
    const at = name.offset + 6 + i * 12;
    const platform = ttf.readUInt16BE(at);
    const encoding = ttf.readUInt16BE(at + 2);
    const language = ttf.readUInt16BE(at + 4);
    if (ttf.readUInt16BE(at + 6) !== id) continue;
    const length = ttf.readUInt16BE(at + 8);
    const start = strings + ttf.readUInt16BE(at + 10);
    if (platform === 3 && (encoding === 1 || encoding === 0)) {
      const le = Buffer.from(ttf.subarray(start, start + length));
      le.swap16();
      if (language === 0x409) return le;
      fallback ??= le;
    } else if (platform === 1 && !fallback) {
      fallback = Buffer.from(ttf.toString('latin1', start, start + length), 'utf16le');
    }
  }
  return fallback ?? Buffer.alloc(0);
};

export const ttfToEot = (ttf: Buffer): Buffer => {
  const t = tables(ttf);
  const os2 = t.get('OS/2');
  const head = t.get('head');
  const name = t.get('name');
  const fixed = Buffer.alloc(82);
  fixed.writeUInt32LE(0x00020001, 8); // Version
  fixed.writeUInt32LE(0, 12); // Flags: sem compressão nem XOR
  if (os2) {
    ttf.copy(fixed, 16, os2.offset + 32, os2.offset + 42); // PANOSE
    fixed.writeUInt32LE(ttf.readUInt16BE(os2.offset + 4), 28); // usWeightClass
    fixed.writeUInt16LE(ttf.readUInt16BE(os2.offset + 8), 32); // fsType
    for (let i = 0; i < 4; i++) fixed.writeUInt32LE(ttf.readUInt32BE(os2.offset + 42 + i * 4), 36 + i * 4); // ulUnicodeRange1-4
    if (ttf.readUInt16BE(os2.offset) >= 1) {
      for (let i = 0; i < 2; i++) fixed.writeUInt32LE(ttf.readUInt32BE(os2.offset + 78 + i * 4), 52 + i * 4); // ulCodePageRange1-2
    }
    fixed.writeUInt8(ttf.readUInt16BE(os2.offset + 62) & 1 ? 1 : 0, 27); // fsSelection ITALIC
  }
  fixed.writeUInt8(1, 26); // Charset: DEFAULT_CHARSET
  fixed.writeUInt16LE(0x504c, 34); // MagicNumber
  if (head) fixed.writeUInt32LE(ttf.readUInt32BE(head.offset + 8), 60); // checkSumAdjustment
  // 64–79: Reserved1–4; 80: Padding1.

  const field = (data: Buffer, padded = true): Buffer => {
    const out = Buffer.alloc((padded ? 2 : 0) + 2 + data.length);
    out.writeUInt16LE(data.length, padded ? 2 : 0);
    data.copy(out, padded ? 4 : 2);
    return out;
  };
  const names = Buffer.concat([
    field(nameOf(ttf, name, 1), false), // FamilyName (Padding1 já está no fixo)
    field(nameOf(ttf, name, 2)), // StyleName
    field(nameOf(ttf, name, 5)), // VersionName
    field(nameOf(ttf, name, 4)), // FullName
    field(Buffer.alloc(0)), // RootString (vazio: qualquer documento pode usar)
  ]);
  const out = Buffer.concat([fixed, names, ttf]);
  out.writeUInt32LE(out.length, 0); // EOTSize
  out.writeUInt32LE(ttf.length, 4); // FontDataSize
  return out;
};
