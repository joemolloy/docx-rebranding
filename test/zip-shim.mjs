// Minimal fflate-compatible ZIP shim over Node's built-in zlib, so the browser
// tool's real logic (which calls fflate.{unzipSync,zipSync,strFromU8,strToU8})
// can run headlessly in Node without npm. Reads via the central directory (so it
// tolerates data descriptors that Word sometimes emits) and writes standard
// deflate/store entries readable by Word and fflate alike.
import zlib from "node:zlib";

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function strFromU8(u8) {
  return Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength).toString("utf8");
}
export function strToU8(str) {
  return new Uint8Array(Buffer.from(str, "utf8"));
}

export function unzipSync(data) {
  const buf = Buffer.from(data.buffer || data, data.byteOffset || 0, data.byteLength ?? data.length);
  // Locate End Of Central Directory (scan backwards for signature 0x06054b50).
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("zip-shim: EOCD not found");
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const out = {};
  for (let e = 0; e < count; e++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error("zip-shim: bad central dir");
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.toString("utf8", off + 46, off + 46 + nameLen);
    // Jump to the local header to find where the actual data begins.
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const comp = buf.subarray(dataStart, dataStart + compSize);
    const raw = method === 0 ? comp : zlib.inflateRawSync(comp);
    out[name] = new Uint8Array(raw);
    off += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

export function zipSync(files, opts = {}) {
  const level = opts.level ?? 6;
  const entries = [];
  for (const [name, u8] of Object.entries(files)) {
    const data = Buffer.from(u8.buffer || u8, u8.byteOffset || 0, u8.byteLength ?? u8.length);
    const crc = crc32(data);
    let method = 8;
    let body = level === 0 ? null : zlib.deflateRawSync(data, { level });
    if (!body || body.length >= data.length) { method = 0; body = data; }
    entries.push({ name: Buffer.from(name, "utf8"), data, crc, method, body });
  }
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const en of entries) {
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);              // version needed
    local.writeUInt16LE(0, 6);               // flags
    local.writeUInt16LE(en.method, 8);
    local.writeUInt16LE(0, 10);              // mod time
    local.writeUInt16LE(0x21, 12);           // mod date (arbitrary valid)
    local.writeUInt32LE(en.crc, 14);
    local.writeUInt32LE(en.body.length, 18);
    local.writeUInt32LE(en.data.length, 22);
    local.writeUInt16LE(en.name.length, 26);
    local.writeUInt16LE(0, 28);              // extra len
    chunks.push(local, en.name, en.body);

    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0);
    cen.writeUInt16LE(20, 4);                // version made by
    cen.writeUInt16LE(20, 6);                // version needed
    cen.writeUInt16LE(0, 8);                 // flags
    cen.writeUInt16LE(en.method, 10);
    cen.writeUInt16LE(0, 12);
    cen.writeUInt16LE(0x21, 14);
    cen.writeUInt32LE(en.crc, 16);
    cen.writeUInt32LE(en.body.length, 20);
    cen.writeUInt32LE(en.data.length, 24);
    cen.writeUInt16LE(en.name.length, 28);
    cen.writeUInt16LE(0, 30);                // extra
    cen.writeUInt16LE(0, 32);                // comment
    cen.writeUInt16LE(0, 34);                // disk
    cen.writeUInt16LE(0, 36);                // internal attrs
    cen.writeUInt32LE(0, 38);                // external attrs
    cen.writeUInt32LE(offset, 42);           // local header offset
    central.push(cen, en.name);
    offset += local.length + en.name.length + en.body.length;
  }
  const centralStart = offset;
  let centralSize = 0;
  for (const c of central) centralSize += c.length;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralStart, 16);
  const all = Buffer.concat([...chunks, ...central, eocd]);
  return new Uint8Array(all);
}

export default { unzipSync, zipSync, strFromU8, strToU8 };
