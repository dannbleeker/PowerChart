#!/usr/bin/env node
/**
 * Draw the add-in icons in the SSF palette.
 *
 *     node scripts/build-icons.mjs           # write assets/icon-*.png + store logo
 *     node scripts/build-icons.mjs --check   # fail if they are stale
 *
 * GENERATED RATHER THAN DRAWN, because these are four sizes of one mark and the
 * 16px one is where a hand-exported PNG goes wrong: a bar landing on a half
 * pixel turns grey, and at 16px a grey bar is an invisible one. Each size is
 * drawn natively at its own resolution — never downscaled from the 80 — so
 * every edge lands on a pixel boundary.
 *
 * THE GEOMETRY IS THE OLD ICON'S, measured from `icon-80.png` before it was
 * replaced: three ascending bars at x=0.150/0.425/0.700, width 0.175, tops at
 * 0.550/0.3125/0.0625, baseline 0.925. Only the colours changed. The mark was
 * never the problem — `#2A78D6` was, a blue that is in no SSF palette.
 *
 * THE TALLEST BAR IS ORANGE, and that is the icon's one accent. It does not
 * compete with the pane's tick: the ribbon is a different surface, seen at a
 * different moment, and the rule is one orange per view rather than one per
 * product. It also gives the mark somewhere for the eye to land at 16px, where
 * three identical white bars read as a smudge.
 */
import { writeFileSync, readFileSync, existsSync } from "node:fs";
import { deflateSync } from "node:zlib";

const NAVY = [0x00, 0x25, 0x4c];
const WHITE = [0xff, 0xff, 0xff];
const ORANGE = [0xed, 0x89, 0x36];

/** Bars as fractions of the icon's side — measured from the icon this replaces. */
const BARS = [
  { x: 0.15, top: 0.55, fill: WHITE },
  { x: 0.425, top: 0.3125, fill: WHITE },
  { x: 0.7, top: 0.0625, fill: ORANGE },
];
const BAR_W = 0.175;
const BASELINE = 0.925;

const SIZES = [16, 32, 64, 80];

/**
 * The AppSource listing image, which is not a ribbon icon.
 *
 * `docs/STORE-LISTING.md` asks for a 300x300 store logo and says in as many
 * words that it is separate from the 16/32/80 the manifest carries. Generated
 * from the SAME geometry and constants anyway: a listing image drawn by hand
 * would be the one place the mark could drift, and it is the one place a drift
 * is seen by strangers rather than by us.
 *
 * Written to `assets/` beside the icons rather than into `docs/`, because it is
 * an artefact rather than documentation, and `--check` covers it too.
 */
const STORE_LOGO = 300;

/** One icon as raw RGBA pixels. */
function pixels(size) {
  const px = Buffer.alloc(size * size * 4);
  for (let i = 0; i < size * size; i++) {
    px[i * 4] = NAVY[0];
    px[i * 4 + 1] = NAVY[1];
    px[i * 4 + 2] = NAVY[2];
    px[i * 4 + 3] = 255;
  }
  // ROUNDED TO WHOLE PIXELS, and the width is floored to at least 1: at 16px
  // the bar is 2.8px, and a fractional edge is what turns a bar grey.
  const w = Math.max(1, Math.round(BAR_W * size));
  const bottom = Math.round(BASELINE * size);
  for (const bar of BARS) {
    const x0 = Math.round(bar.x * size);
    const y0 = Math.round(bar.top * size);
    for (let y = y0; y < bottom; y++) {
      for (let x = x0; x < Math.min(x0 + w, size); x++) {
        const i = (y * size + x) * 4;
        px[i] = bar.fill[0];
        px[i + 1] = bar.fill[1];
        px[i + 2] = bar.fill[2];
        px[i + 3] = 255;
      }
    }
  }
  return px;
}

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

/** A minimal 8-bit RGBA PNG. No dependency for four flat images. */
function png(size, px) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // colour type: RGBA
  // Each scanline is prefixed with filter byte 0 (None) — the images are flat,
  // so a smarter filter would buy nothing on a file this size.
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    px.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const check = process.argv.includes("--check");
let stale = 0;
const outputs = [
  ...SIZES.map((size) => ({ size, path: `assets/icon-${size}.png` })),
  { size: STORE_LOGO, path: "assets/store-logo-300.png" },
];
for (const { size, path } of outputs) {
  const bytes = png(size, pixels(size));
  if (check) {
    const current = existsSync(path) ? readFileSync(path) : Buffer.alloc(0);
    if (!current.equals(bytes)) {
      console.error(`${path} is stale — run \`node scripts/build-icons.mjs\` and commit it`);
      stale++;
    }
  } else {
    writeFileSync(path, bytes);
    console.log(`${path} (${size}x${size}, ${bytes.length} bytes)`);
  }
}
if (check && !stale) console.log("icons and the store logo are current");
process.exit(stale ? 1 : 0);
