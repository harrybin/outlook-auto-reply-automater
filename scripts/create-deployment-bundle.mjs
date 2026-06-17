import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const distDir = path.join(repoRoot, "dist");
const deploymentDir = path.join(distDir, "deployment");
const webDir = path.join(deploymentDir, "web");
const manifestPath = path.join(repoRoot, "manifest.json");
const iconsDir = path.join(repoRoot, "assets");
const readmePath = path.join(repoRoot, "README.md");

async function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.help) {
    printHelp();
    return;
  }

  const baseUrl = normalizeBaseUrl(args.baseUrl ?? process.env.DEPLOY_BASE_URL);
  await ensureBuildOutput();
  await fs.rm(deploymentDir, { recursive: true, force: true });
  await fs.mkdir(webDir, { recursive: true });

  await copyBuildOutput();
  await copyIcons();
  await copyReadme();

  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  const deployedManifest = rewriteLocalhostUrls(manifest, baseUrl);
  const outputManifestPath = path.join(deploymentDir, "manifest.json");

  await fs.writeFile(outputManifestPath, `${JSON.stringify(deployedManifest, null, 2)}\n`);
  await fs.writeFile(path.join(deploymentDir, "DEPLOY.md"), createDeploymentGuide(baseUrl));

  console.log(`Deployment bundle created at ${deploymentDir}`);
  console.log(`Hosted web assets: ${path.relative(repoRoot, webDir)}`);
  console.log(`Manifest: ${path.relative(repoRoot, outputManifestPath)}`);
  console.log(`Base URL: ${baseUrl}`);
}

function parseArgs(argv) {
  const parsed = {};

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--help" || arg === "-h") {
      parsed.help = true;
      continue;
    }

    if (arg === "--base-url") {
      parsed.baseUrl = argv[index + 1];
      index += 1;
      continue;
    }

    if (arg.startsWith("--base-url=")) {
      parsed.baseUrl = arg.slice("--base-url=".length);
      continue;
    }

    if (!arg.startsWith("-") && !parsed.baseUrl) {
      parsed.baseUrl = arg;
      continue;
    }

    throw new Error(`Unknown argument: ${arg}`);
  }

  return parsed;
}

function printHelp() {
  console.log("Create a production deployment bundle for Outlook on Windows and Mac.");
  console.log("");
  console.log("Usage:");
  console.log("  npm run pack -- https://your-host.example.com");
  console.log("  DEPLOY_BASE_URL=https://your-host.example.com npm run pack");
}

function normalizeBaseUrl(input) {
  if (!input) {
    throw new Error(
      "DEPLOY_BASE_URL is required. Provide an HTTPS hosting URL, for example https://addins.example.com/outlook-auto-reply-automater",
    );
  }

  const normalized = input.trim().replace(/\/$/, "");
  let parsedUrl;

  try {
    parsedUrl = new URL(normalized);
  } catch {
    throw new Error(`Invalid DEPLOY_BASE_URL: ${input}`);
  }

  if (parsedUrl.protocol !== "https:") {
    throw new Error("DEPLOY_BASE_URL must use HTTPS.");
  }

  if (["localhost", "127.0.0.1", "::1"].includes(parsedUrl.hostname)) {
    throw new Error("DEPLOY_BASE_URL must be a hosted production URL, not localhost.");
  }

  return normalized;
}

async function ensureBuildOutput() {
  const taskpaneHtmlPath = path.join(distDir, "src", "taskpane", "index.html");

  try {
    await fs.access(taskpaneHtmlPath);
  } catch {
    throw new Error(`Production build output is missing at ${taskpaneHtmlPath}. Run npm run build first.`);
  }
}

async function copyBuildOutput() {
  const entries = await fs.readdir(distDir, { withFileTypes: true });

  for (const entry of entries) {
    if (entry.name === "deployment") {
      continue;
    }

    const sourcePath = path.join(distDir, entry.name);
    const targetPath = path.join(webDir, entry.name);
    await copyEntry(sourcePath, targetPath);
  }
}

async function copyIcons() {
  const targetAssetsDir = path.join(webDir, "assets");
  await fs.mkdir(targetAssetsDir, { recursive: true });

  for (const iconName of ["icon-16.png", "icon-32.png", "icon-80.png", "icon-color.png", "icon-outline.png"]) {
    await fs.copyFile(path.join(iconsDir, iconName), path.join(targetAssetsDir, iconName));
  }
}

async function copyReadme() {
  await fs.copyFile(readmePath, path.join(webDir, "README.md"));
}

async function copyEntry(sourcePath, targetPath) {
  const stats = await fs.stat(sourcePath);

  if (stats.isDirectory()) {
    await fs.mkdir(targetPath, { recursive: true });
    const children = await fs.readdir(sourcePath);

    for (const child of children) {
      await copyEntry(path.join(sourcePath, child), path.join(targetPath, child));
    }

    return;
  }

  await fs.copyFile(sourcePath, targetPath);
}

function rewriteLocalhostUrls(value, baseUrl) {
  if (typeof value === "string") {
    return value.replace(/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?/i, baseUrl);
  }

  if (Array.isArray(value)) {
    return value.map((item) => rewriteLocalhostUrls(item, baseUrl));
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, nestedValue]) => [key, rewriteLocalhostUrls(nestedValue, baseUrl)]),
    );
  }

  return value;
}

function createDeploymentGuide(baseUrl) {
  return `# Deployment Bundle\n\nThis bundle is intended for Outlook on Windows and Mac.\n\n## Contents\n- manifest.json: Production manifest for sideloading or catalog deployment\n- web/: Static web assets that must be hosted over HTTPS\n\n## Host The Web Assets\n1. Upload the contents of web/ to an HTTPS host.\n2. Preserve the existing folder structure so the task pane stays available at ${baseUrl}/src/taskpane/index.html.\n3. Confirm these URLs are reachable before sideloading:\n   - ${baseUrl}/src/taskpane/index.html\n   - ${baseUrl}/assets/icon-16.png\n   - ${baseUrl}/assets/icon-32.png\n   - ${baseUrl}/assets/icon-80.png\n\n## Deploy In Outlook\n### Windows and Mac\n1. Open Outlook.\n2. Go to Get Add-ins or Apps.\n3. Choose My add-ins or My Apps.\n4. Use Add a custom add-in or Upload custom apps.\n5. Select manifest.json from this folder.\n\n## Centralized Deployment\nUse manifest.json after the web assets are hosted and validated. The same manifest can be used for admin deployment or AppSource submission workflows.\n`;
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});