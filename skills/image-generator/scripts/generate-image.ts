import { writeFileSync, mkdirSync } from "fs";
import * as path from "path";

interface Args {
  prompt: string;
  output: string;
  format: "png" | "jpg" | "jpeg";
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

function main() {
  const args = parseArgs();
  const apiKey = process.env.IMAGE_API_KEY;
  if (!apiKey) {
    throw new Error("IMAGE_API_KEY is not set");
  }

  const outPath = path.resolve(args.output);
  mkdirSync(path.dirname(outPath), { recursive: true });

  // Replace this stub with your image backend call.
  // It should return PNG/JPEG bytes and write them to `outPath`.
  const buffer = Buffer.from([]); // stub

  writeFileSync(outPath, buffer);
  console.log(outPath);
}

main();
