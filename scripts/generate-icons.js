#!/usr/bin/env node
/**
 * Generate PNG icons for the Outlook Auto-Reply Automater add-in.
 * Creates 16x16, 32x32, 80x80 ribbon icons plus 32x32 Teams icons.
 */

const fs = require("fs");
const path = require("path");
const { PNG } = require("pngjs");

const assetsDir = path.join(__dirname, "..", "assets");
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

const COLORS = {
  transparent: [0, 0, 0, 0],
  blue: [0, 120, 212, 255],
  blueDark: [0, 90, 158, 255],
  white: [255, 255, 255, 255],
};

function putPixel(png, x, y, color) {
  if (x < 0 || y < 0 || x >= png.width || y >= png.height) {
    return;
  }
  const idx = (png.width * y + x) << 2;
  png.data[idx] = color[0];
  png.data[idx + 1] = color[1];
  png.data[idx + 2] = color[2];
  png.data[idx + 3] = color[3];
}

function fillRect(png, x, y, width, height, color) {
  for (let yy = y; yy < y + height; yy++) {
    for (let xx = x; xx < x + width; xx++) {
      putPixel(png, xx, yy, color);
    }
  }
}

function strokeRect(png, x, y, width, height, strokeWidth, color) {
  fillRect(png, x, y, width, strokeWidth, color);
  fillRect(png, x, y + height - strokeWidth, width, strokeWidth, color);
  fillRect(png, x, y, strokeWidth, height, color);
  fillRect(png, x + width - strokeWidth, y, strokeWidth, height, color);
}

function drawEnvelope(png, color) {
  const s = png.width;
  const left = Math.floor(s * 0.18);
  const top = Math.floor(s * 0.2);
  const width = Math.floor(s * 0.64);
  const height = Math.floor(s * 0.52);
  const stroke = Math.max(1, Math.floor(s * 0.08));

  strokeRect(png, left, top, width, height, stroke, color);

  for (let i = 0; i <= width / 2; i++) {
    const y = top + stroke + Math.floor((i * (height * 0.42)) / (width / 2));
    const leftX = left + stroke + i;
    const rightX = left + width - stroke - i;
    for (let t = 0; t < stroke; t++) {
      putPixel(png, leftX, y + t, color);
      putPixel(png, rightX, y + t, color);
    }
  }
}

function drawClock(png, color) {
  const s = png.width;
  const radius = Math.max(2, Math.floor(s * 0.14));
  const cx = Math.floor(s * 0.72);
  const cy = Math.floor(s * 0.76);
  const stroke = Math.max(1, Math.floor(s * 0.06));

  for (let y = -radius; y <= radius; y++) {
    for (let x = -radius; x <= radius; x++) {
      const distSq = x * x + y * y;
      if (
        distSq <= radius * radius &&
        distSq >= (radius - stroke) * (radius - stroke)
      ) {
        putPixel(png, cx + x, cy + y, color);
      }
    }
  }

  for (let i = 0; i < radius - stroke; i++) {
    putPixel(png, cx, cy - i, color);
    putPixel(png, cx + Math.floor(i * 0.6), cy, color);
  }
}

function createRibbonIcon(size) {
  const png = new PNG({ width: size, height: size });
  fillRect(png, 0, 0, size, size, COLORS.blue);

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      if (x > y + Math.floor(size * 0.45)) {
        putPixel(png, x, y, COLORS.blueDark);
      }
    }
  }

  drawEnvelope(png, COLORS.white);
  drawClock(png, COLORS.white);
  return png;
}

function createOutlineIcon(size) {
  const png = new PNG({ width: size, height: size });
  fillRect(png, 0, 0, size, size, COLORS.transparent);
  drawEnvelope(png, COLORS.white);
  drawClock(png, COLORS.white);
  return png;
}

function createColorIcon(size) {
  return createRibbonIcon(size);
}

function writePng(name, desc, png) {
  const filePath = path.join(assetsDir, name);
  fs.writeFileSync(filePath, PNG.sync.write(png));
  console.log(`✓ Created ${name} (${desc})`);
}

writePng("icon-16.png", "16x16 ribbon icon", createRibbonIcon(16));
writePng("icon-32.png", "32x32 ribbon icon", createRibbonIcon(32));
writePng("icon-80.png", "80x80 ribbon icon", createRibbonIcon(80));
writePng("icon-outline.png", "32x32 outline icon", createOutlineIcon(32));
writePng("icon-color.png", "32x32 full-color icon", createColorIcon(32));

console.log(`\n5 icon files created in ${assetsDir}/`);
