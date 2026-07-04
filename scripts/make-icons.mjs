// Generate simple placeholder PNG icons (solid brand-blue squares with a white bar motif)
// without any image dependencies — raw PNG encoding via zlib.
import { deflateSync } from "node:zlib";
import { mkdirSync, writeFileSync } from "node:fs";

function crc32(buf) {
  let c,
    table = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  let crc = 0xffffffff;
  for (const b of buf) crc = table[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function png(size) {
  const bg = [0x2a, 0x78, 0xd6]; // brand blue
  const bar = [0xff, 0xff, 0xff];
  // Three white "columns" of different heights — a tiny bar chart.
  const cols = [
    { x0: 0.15, x1: 0.32, h: 0.45 },
    { x0: 0.42, x1: 0.59, h: 0.7 },
    { x0: 0.69, x1: 0.86, h: 0.95 },
  ];
  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = [0]; // filter byte
    for (let x = 0; x < size; x++) {
      const fx = x / size;
      const fy = 1 - y / size;
      const inBar = cols.some((c) => fx >= c.x0 && fx < c.x1 && fy < c.h && fy > 0.08);
      const px = inBar ? bar : bg;
      row.push(...px);
    }
    rows.push(Buffer.from(row));
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type RGB
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(Buffer.concat(rows))),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

mkdirSync("assets", { recursive: true });
for (const size of [16, 32, 64, 80]) {
  writeFileSync(`assets/icon-${size}.png`, png(size));
  console.log(`assets/icon-${size}.png`);
}
