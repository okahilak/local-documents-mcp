// Minimal ZIP reader/writer for OOXML packages that copies untouched entries VERBATIM.
//
// Unchanged entries are copied as raw bytes (local header + compressed data + any data
// descriptor), and their central-directory records are reused with only the local-header
// offset patched. Only replaced or added entries are re-encoded (DEFLATE via node:zlib).

import zlib from "node:zlib";

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const MAX_PART_BYTES = 512 * 1024 * 1024;

export class ZipError extends Error {}

/**
 * Parse a ZIP archive held in `buf`.
 * Returns { entries: [{ name, nameBytes, method, flags, crc, compSize, size, localOffset, raw, central }], comment }.
 * `raw` = the entry's bytes from its local header up to the next entry (or the central directory).
 */
export function readZip(buf) {
  if (buf.length >= 8 && buf.readUInt32BE(0) === 0xd0cf11e0) {
    throw new ZipError("This file is an encrypted (password-protected) or legacy binary Office file, not an OOXML ZIP package.");
  }
  // Locate the End Of Central Directory record (last 22..65557 bytes).
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new ZipError("Not a ZIP/OOXML file (no end-of-central-directory record).");
  if (eocd >= 20 && buf.readUInt32LE(eocd - 20) === SIG_ZIP64_LOCATOR) throw new ZipError("ZIP64 workbooks (over 4 GB or 65535 parts) are not supported for editing.");
  const disk = buf.readUInt16LE(eocd + 4);
  const count = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  const commentLen = buf.readUInt16LE(eocd + 20);
  if (disk !== 0 || buf.readUInt16LE(eocd + 6) !== 0) throw new ZipError("Multi-disk ZIP archives are not supported.");
  if (cdOffset + cdSize > eocd) throw new ZipError("Corrupt ZIP: central directory out of range.");

  const entries = [];
  const seen = new Set();
  let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (p + 46 > eocd || buf.readUInt32LE(p) !== SIG_CENTRAL) throw new ZipError("Corrupt ZIP: bad central directory record.");
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const crc = buf.readUInt32LE(p + 16);
    const compSize = buf.readUInt32LE(p + 20);
    const size = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const cLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const nameBytes = buf.subarray(p + 46, p + 46 + nameLen);
    const name = nameBytes.toString(flags & 0x800 ? "utf8" : "latin1");
    if (compSize === 0xffffffff || size === 0xffffffff || localOffset === 0xffffffff) throw new ZipError("ZIP64 entries are not supported for editing.");
    if (flags & 0x1) throw new ZipError(`ZIP entry ${name} is encrypted.`);
    const key = name.toLowerCase();
    if (seen.has(key)) throw new ZipError(`Corrupt package: duplicate part ${name}.`);
    seen.add(key);
    const recLen = 46 + nameLen + extraLen + cLen;
    entries.push({ name, nameBytes, flags, method, crc, compSize, size, localOffset, central: buf.subarray(p, p + recLen) });
    p += recLen;
  }

  // Each entry's raw chunk runs to the next entry's local header (covers data descriptors).
  const byOffset = [...entries].sort((a, b) => a.localOffset - b.localOffset);
  byOffset.forEach((e, i) => {
    const end = i + 1 < byOffset.length ? byOffset[i + 1].localOffset : cdOffset;
    if (e.localOffset + 30 > end || buf.readUInt32LE(e.localOffset) !== SIG_LOCAL) throw new ZipError(`Corrupt ZIP: bad local header for ${e.name}.`);
    const n = buf.readUInt16LE(e.localOffset + 26);
    const x = buf.readUInt16LE(e.localOffset + 28);
    e.dataStart = e.localOffset + 30 + n + x;
    if (e.dataStart + e.compSize > end) throw new ZipError(`Corrupt ZIP: data for ${e.name} overruns the next entry.`);
    e.raw = buf.subarray(e.localOffset, end);
    e.buf = buf;
  });
  return { entries, comment: buf.subarray(eocd + 22, eocd + 22 + commentLen) };
}

/** Decompress one entry and verify its CRC. */
export function readEntry(entry) {
  if (entry.size > MAX_PART_BYTES) throw new ZipError(`Part ${entry.name} is too large (${entry.size} bytes).`);
  const data = entry.buf.subarray(entry.dataStart, entry.dataStart + entry.compSize);
  let out;
  if (entry.method === 0) out = Buffer.from(data);
  else if (entry.method === 8) out = zlib.inflateRawSync(data, { maxOutputLength: Math.max(entry.size, 1) + 1024 });
  else throw new ZipError(`Part ${entry.name} uses unsupported compression method ${entry.method}.`);
  if (out.length !== entry.size || (zlib.crc32(out) >>> 0) !== entry.crc) throw new ZipError(`CRC/size mismatch in ${entry.name}.`);
  return out;
}

function dosDateTime(d = new Date()) {
  const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2);
  const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  return { time, date };
}

/**
 * Write a new archive.
 * @param {object} zip result of readZip
 * @param {Map<string, Buffer|null>} changes lower-cased part name -> new content (null deletes)
 * @param {Array<{name: string, data: Buffer}>} additions new parts, appended at the end
 */
export function writeZip(zip, changes, additions = []) {
  const chunks = [];
  const central = [];
  let offset = 0;
  const { time, date } = dosDateTime();
  let versionMadeBy = 20;

  const encodeNew = (nameBytes, data, flags, template) => {
    const comp = zlib.deflateRawSync(data, { level: 6 });
    const crc = zlib.crc32(data) >>> 0;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(SIG_LOCAL, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(date, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(comp.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    const comment = template ? template.central.subarray(46 + template.nameBytes.length + template.central.readUInt16LE(30)) : Buffer.alloc(0);
    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(SIG_CENTRAL, 0);
    cd.writeUInt16LE(template ? template.central.readUInt16LE(4) : versionMadeBy, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(flags, 8);
    cd.writeUInt16LE(8, 10);
    cd.writeUInt16LE(time, 12);
    cd.writeUInt16LE(date, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(comp.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBytes.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(comment.length, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(template ? template.central.readUInt16LE(36) : 0, 36);
    cd.writeUInt32LE(template ? template.central.readUInt32LE(38) : 0, 38);
    cd.writeUInt32LE(offset, 42);
    const record = Buffer.concat([local, nameBytes, comp]);
    chunks.push(record);
    central.push(Buffer.concat([cd, nameBytes, comment]));
    offset += record.length;
  };

  for (const e of zip.entries) {
    versionMadeBy = e.central.readUInt16LE(4);
    const key = e.name.toLowerCase();
    if (changes.has(key)) {
      const data = changes.get(key);
      if (data === null) continue; // deleted
      encodeNew(e.nameBytes, data, e.flags & 0x800, e);
      continue;
    }
    // Verbatim copy; patch only the local-header offset in the central record.
    const cd = Buffer.from(e.central);
    cd.writeUInt32LE(offset, 42);
    chunks.push(e.raw);
    central.push(cd);
    offset += e.raw.length;
  }
  for (const a of additions) {
    const nameBytes = Buffer.from(a.name, "utf8");
    encodeNew(nameBytes, a.data, /[^\x20-\x7e]/.test(a.name) ? 0x800 : 0, null);
  }

  const count = central.length;
  if (count > 0xffff) throw new ZipError("Too many parts for a non-ZIP64 archive.");
  const cdBuf = Buffer.concat(central);
  if (offset + cdBuf.length > 0xffffffff) throw new ZipError("Workbook would exceed 4 GB.");
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(SIG_EOCD, 0);
  eocd.writeUInt16LE(count, 8);
  eocd.writeUInt16LE(count, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(zip.comment.length, 20);
  return Buffer.concat([...chunks, cdBuf, eocd, zip.comment]);
}
