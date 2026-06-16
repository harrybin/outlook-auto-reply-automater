import { writeFileSync, mkdirSync } from "fs";
import * as path from "path";
import { createHash } from "crypto";
import { PNG } from "pngjs";
import jpeg from "jpeg-js";

interface Args {
  prompt: string;
  output: string;
  format: "png" | "jpg" | "jpeg";
}

interface RenderedImage {
  width: number;
  height: number;
  rgba: Uint8Array;
}

function parseArgs(): Args {
  // Simple Node-style parser for: --prompt "..." --output "..." --format png
  const args: Args = { prompt: "", output: "", format: "png" };
  let i = 0;
  while (i < process.argv.length) {
    const key = process.argv[i];
    if (key === "--prompt" && i + 1) {
      args.prompt = process.argv[i + 1];
      i += 2;
    } else if (key === "--output" && i + 1) {
      args.output = process.argv[i + 1];
      i += 2;
    } else if (key === "--format" && i + 1) {
      const fmt = process.argv[i + 1].toLowerCase();
      args.format =
        fmt === "png" || fmt === "jpg" || fmt === "jpeg" ? fmt : "png";
      if (args.format === "jpeg") args.format = "jpg";
      i += 2;
    } else {
      i += 1;
    }
  }
  if (!args.prompt || !args.output) {
    throw new Error("--prompt and --output are required");
  }
  return args;
}

function getWindowsIdentity(): string {
  const username = process.env.USERNAME;
  const domain = process.env.USERDOMAIN;
  const identity = domain && username ? `${domain}\\${username}` : username;

  if (!identity) {
    throw new Error("Unable to determine the current Windows user credentials");
  }

  return identity;
}

function hashToSeedBytes(prompt: string, identity: string): Buffer {
  return createHash("sha256").update(`${identity}::${prompt}`).digest();
}

function renderPromptImage(prompt: string, identity: string): RenderedImage {
  const width = 768;
  const height = 768;
  const rgba = new Uint8Array(width * height * 4);
  const seed = hashToSeedBytes(prompt, identity);

  const fx = 6 + (seed[0] % 18);
  const fy = 8 + (seed[1] % 22);
  const swirl = 10 + (seed[2] % 20);
  const twist = 0.008 + (seed[3] % 8) * 0.001;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = (y * width + x) * 4;

      const nx = x / width - 0.5;
      const ny = y / height - 0.5;
      const radius = Math.sqrt(nx * nx + ny * ny);
      const angle = Math.atan2(ny, nx) + radius * swirl * twist;

      const waveA = Math.sin((x + seed[4]) * 0.02 * fx + angle);
      const waveB = Math.cos((y + seed[5]) * 0.018 * fy - angle * 0.7);
      const waveC = Math.sin((x + y + seed[6]) * 0.012 * ((seed[7] % 24) + 6));
      const blend = (waveA + waveB + waveC + 3) / 6;

      const r = Math.floor(
        blend * 255 * (0.65 + (seed[8] % 35) / 100) + radius * 75,
      );
      const g = Math.floor(
        (1 - blend) * 255 * (0.55 + (seed[9] % 40) / 100) + Math.abs(nx) * 70,
      );
      const b = Math.floor(
        ((Math.sin(angle * ((seed[10] % 9) + 2)) + 1) / 2) *
          255 *
          (0.5 + (seed[11] % 45) / 100),
      );

      rgba[i] = Math.max(0, Math.min(255, r));
      rgba[i + 1] = Math.max(0, Math.min(255, g));
      rgba[i + 2] = Math.max(0, Math.min(255, b));
      rgba[i + 3] = 255;
    }
  }

  return { width, height, rgba };
}

function encodeImage(image: RenderedImage, format: Args["format"]): Buffer {
  if (format === "png") {
    const png = new PNG({ width: image.width, height: image.height });
    png.data = Buffer.from(image.rgba);
    return PNG.sync.write(png);
  }

  const encoded = jpeg.encode(
    {
      data: Buffer.from(image.rgba),
      width: image.width,
      height: image.height,
    },
    90,
  );

  return Buffer.from(encoded.data);
}

function main() {
  const args = parseArgs();
  const windowsIdentity = getWindowsIdentity();

  const outPath = path.resolve(args.output);
  mkdirSync(path.dirname(outPath), { recursive: true });

  const image = renderPromptImage(args.prompt, windowsIdentity);
  const buffer = encodeImage(image, args.format);

  writeFileSync(outPath, buffer);
  console.log(outPath);
}

main();
