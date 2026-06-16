---
name: image-generator
description: Generate PNG or JPEG images from a natural-language prompt by running the local TypeScript script. Use when the user asks to create, render, or export an image file.
---

# Image Generator Skill

Use this skill when the user wants to generate an image file.

## Workflow

1. Read the user's request and extract the prompt, format, and output filename if provided. If no output filename is provided, generate one by slugifying the first 4 words of the prompt and appending the format extension (e.g., `fantasy-forest-scene-as.png`).
2. Run `skills/image-generator/scripts/generate-image.ts` with the prompt and options via Node. If the script fails to execute for any reason (file not found, Node not installed, non-zero exit code), return the exact error message from the process and stop. Do not retry automatically.
3. Save the image as PNG by default, or JPEG if the user explicitly requests it. Always save to the `output/` directory in the project root unless the user specifies an absolute path (e.g., `output/forest.png`).
4. Return the relative output path from the project root (e.g., `output/forest.png`) and a one-sentence description of the generated image content (e.g., "Generated a PNG of a fantasy forest scene.").

## Rules

- Do not hardcode secrets.
- Do not require `IMAGE_API_KEY`; use the current Windows user context from environment credentials (`USERDOMAIN`/`USERNAME`).
- Prefer PNG unless the user explicitly asks for JPEG. If the user requests a format other than PNG or JPEG, respond with an error: "Only PNG and JPEG formats are supported. Please request one of these formats." Do not attempt to generate the image.
- If the backend is unavailable, return a clear error and stop.
- Keep filenames short and deterministic.

## Example invocation

Generate a fantasy forest scene as a PNG named `forest.png`."
This is how Copilot decides when to load the skill and what to do with it [web:46][web:59][web:60].

## TypeScript script

The implementation lives in [`scripts/generate-image.ts`](scripts/generate-image.ts).

## How to run from Copilot

- Via Node: `node skills/image-generator/scripts/generate-image.ts --prompt "..." --output "output/forest.png" --format png`
- Or with a wrapper script (e.g., `npm run generate-image`) later.

This is the exact TypeScript version of the image-generator skill, and it works the same way as the Python version but uses Node/TS instead of Python [web:46][web:60].
