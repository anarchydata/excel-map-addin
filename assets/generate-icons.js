/* Generates icon-16/32/64/80.png: white minimap glyph on a teal disc. */
const zlib = require("zlib");
const fs = require("fs");
const path = require("path");

function crc32(buf) {
  let c, table = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
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

function writePng(size, pixels, file) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    pixels.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0))
  ]);
  fs.writeFileSync(file, png);
  console.log("Wrote", file);
}

function render(size) {
  const px = Buffer.alloc(size * size * 4);
  const SS = 4;
  const teal = [0x0f, 0x6e, 0x56];

  function inMapFrame(u, v) {
    return u >= 0.22 && u <= 0.78 && v >= 0.2 && v <= 0.8;
  }
  function inViewport(u, v) {
    return u >= 0.34 && u <= 0.58 && v >= 0.3 && v <= 0.52;
  }
  function onViewportEdge(u, v) {
    if (!inViewport(u, v)) return false;
    const t = 0.035;
    return (
      u < 0.34 + t || u > 0.58 - t ||
      v < 0.3 + t || v > 0.52 - t
    );
  }

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let disc = 0, glyph = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const u = (x + (sx + 0.5) / SS) / size;
          const v = (y + (sy + 0.5) / SS) / size;
          const d = Math.hypot(u - 0.5, v - 0.5);
          if (d <= 0.48) {
            disc++;
            if (inMapFrame(u, v) && (onViewportEdge(u, v) || (!inViewport(u, v) && ((Math.floor(u * 18) + Math.floor(v * 18)) % 3 === 0)))) {
              glyph++;
            } else if (onViewportEdge(u, v)) {
              glyph++;
            }
          }
        }
      }
      const i = (y * size + x) * 4;
      if (disc === 0) {
        px[i] = px[i + 1] = px[i + 2] = px[i + 3] = 0;
      } else {
        const aDisc = disc / (SS * SS);
        const aGlyph = glyph / (SS * SS);
        const r = Math.round(teal[0] * (1 - aGlyph) + 255 * aGlyph);
        const g = Math.round(teal[1] * (1 - aGlyph) + 255 * aGlyph);
        const b = Math.round(teal[2] * (1 - aGlyph) + 255 * aGlyph);
        px[i] = r;
        px[i + 1] = g;
        px[i + 2] = b;
        px[i + 3] = Math.round(255 * aDisc);
      }
    }
  }
  return px;
}

const outDir = __dirname;
[16, 32, 64, 80].forEach((size) => {
  writePng(size, render(size), path.join(outDir, `icon-${size}.png`));
});
