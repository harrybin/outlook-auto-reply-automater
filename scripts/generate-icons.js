#!/usr/bin/env node
/**
 * Generate PNG icons for the Outlook Auto-Reply Automater add-in.
 * Creates 16x16, 32x32, 80x80 ribbon icons plus Teams catalog icons.
 */

const fs = require("fs");
const path = require("path");
const { PNG } = require("pngjs");

const OVERSAMPLE = 8;

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
  const hiResSize = size * OVERSAMPLE;
  const png = new PNG({ width: hiResSize, height: hiResSize });
  fillRect(png, 0, 0, hiResSize, hiResSize, COLORS.blue);

  for (let y = 0; y < hiResSize; y++) {
    for (let x = 0; x < hiResSize; x++) {
      if (x > y + Math.floor(hiResSize * 0.45)) {
        putPixel(png, x, y, COLORS.blueDark);
      }
    }
  }

  drawEnvelope(png, COLORS.white);
  drawClock(png, COLORS.white);
  return downsamplePng(png, size);
}

function createOutlineIcon(size) {
  const hiResSize = size * OVERSAMPLE;
  const png = new PNG({ width: hiResSize, height: hiResSize });
  fillRect(png, 0, 0, hiResSize, hiResSize, COLORS.transparent);
  drawEnvelope(png, COLORS.white);
  drawClock(png, COLORS.white);
  return downsamplePng(png, size);
}

function createColorIcon(size) {
  return createRibbonIcon(size);
}

function downsamplePng(source, targetSize) {
  const result = new PNG({ width: targetSize, height: targetSize });
  const scaleX = source.width / targetSize;
  const scaleY = source.height / targetSize;

  for (let y = 0; y < targetSize; y++) {
    for (let x = 0; x < targetSize; x++) {
      const startX = Math.floor(x * scaleX);
      const endX = Math.floor((x + 1) * scaleX);
      const startY = Math.floor(y * scaleY);
      const endY = Math.floor((y + 1) * scaleY);

      let sumR = 0;
      let sumG = 0;
      let sumB = 0;
      let sumA = 0;
      let count = 0;

      for (let yy = startY; yy < endY; yy++) {
        for (let xx = startX; xx < endX; xx++) {
          const idx = (source.width * yy + xx) << 2;
          sumR += source.data[idx];
          sumG += source.data[idx + 1];
          sumB += source.data[idx + 2];
          sumA += source.data[idx + 3];
          count++;
        }
      }

      const outIdx = (result.width * y + x) << 2;
      result.data[outIdx] = Math.round(sumR / count);
      result.data[outIdx + 1] = Math.round(sumG / count);
      result.data[outIdx + 2] = Math.round(sumB / count);
      result.data[outIdx + 3] = Math.round(sumA / count);
    }
  }

  return result;
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
writePng("icon-color.png", "192x192 full-color icon", createColorIcon(192));

console.log(`\n5 icon files created in ${assetsDir}/`);
