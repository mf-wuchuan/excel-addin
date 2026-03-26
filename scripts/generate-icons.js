/**
 * Generate valid PNG icon files using raw binary PNG encoding.
 * No external dependencies needed.
 */
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

function createPNG(size, r, g, b) {
  // Create a simple solid-color PNG with a checkmark-like pattern
  const width = size;
  const height = size;

  // Build raw image data (each row: filter byte + RGBA pixels)
  const rawData = [];
  for (let y = 0; y < height; y++) {
    rawData.push(0); // filter: none
    for (let x = 0; x < width; x++) {
      // Green rounded rectangle background with white checkmark
      const margin = Math.floor(size * 0.1);
      const inRect =
        x >= margin && x < width - margin && y >= margin && y < height - margin;

      // Simple checkmark shape
      const cx = x / width;
      const cy = y / height;
      const onCheck =
        // Down stroke: from (0.2, 0.45) to (0.42, 0.72)
        (cx >= 0.18 && cx <= 0.32 &&
          cy >= 0.35 && cy <= 0.75 &&
          Math.abs(cy - (0.45 + (cx - 0.2) * 1.2)) < 0.08) ||
        // Up stroke: from (0.42, 0.72) to (0.82, 0.28)
        (cx >= 0.32 && cx <= 0.85 &&
          cy >= 0.2 && cy <= 0.78 &&
          Math.abs(cy - (0.72 - (cx - 0.42) * 1.1)) < 0.08);

      if (inRect && onCheck) {
        rawData.push(255, 255, 255, 255); // white checkmark
      } else if (inRect) {
        rawData.push(r, g, b, 255); // green background
      } else {
        rawData.push(0, 0, 0, 0); // transparent
      }
    }
  }

  const raw = Buffer.from(rawData);
  const compressed = zlib.deflateSync(raw);

  // Build PNG file
  const chunks = [];

  // Signature
  chunks.push(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));

  // IHDR
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  chunks.push(makeChunk("IHDR", ihdr));

  // IDAT
  chunks.push(makeChunk("IDAT", compressed));

  // IEND
  chunks.push(makeChunk("IEND", Buffer.alloc(0)));

  return Buffer.concat(chunks);
}

function makeChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeB = Buffer.from(type, "ascii");
  const crcData = Buffer.concat([typeB, data]);

  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(crcData) >>> 0, 0);

  return Buffer.concat([len, typeB, data, crc]);
}

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return crc ^ 0xffffffff;
}

const assetsDir = path.resolve(__dirname, "../assets");
for (const size of [16, 32, 80]) {
  const png = createPNG(size, 0x21, 0x73, 0x46); // Excel green
  const filePath = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(filePath, png);
  console.log(`Created ${filePath} (${png.length} bytes)`);
}
