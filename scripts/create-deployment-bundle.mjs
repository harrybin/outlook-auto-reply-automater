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
const hostedEntries = ["src/taskpane/index.html", "src/demo/index.html"];
const hostedAssetRoots = ["assets/", "src/"];

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
  await rewriteHostedHtmlFiles(baseUrl);
  await validateHostedEntries(baseUrl);
  await copyIcons();
  await copyReadme();

  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  const deployedManifest = rewriteManifestHostedUrls(manifest, baseUrl);
  const outputManifestPath = path.join(deploymentDir, "manifest.json");

  const manifestJson = `${JSON.stringify(deployedManifest, null, 2)}\n`;
  await fs.writeFile(outputManifestPath, manifestJson);
  await fs.writeFile(path.join(webDir, "manifest.json"), manifestJson);
  await fs.writeFile(
    path.join(deploymentDir, "DEPLOY.md"),
    createDeploymentGuide(baseUrl),
  );

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
  console.log(
    "Create a production deployment bundle for Outlook on Windows and Mac.",
  );
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
    throw new Error(
      "DEPLOY_BASE_URL must be a hosted production URL, not localhost.",
    );
  }

  return normalized;
}

async function ensureBuildOutput() {
  for (const hostedEntry of hostedEntries) {
    const entryPath = path.join(distDir, hostedEntry);

    try {
      await fs.access(entryPath);
    } catch {
      throw new Error(
        `Production build output is missing at ${entryPath}. Run npm run build first.`,
      );
    }
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
  const targetAssetsDirs = [
    path.join(webDir, "assets"),
    path.join(deploymentDir, "assets"),
  ];
  const iconNames = [
    "icon-16.png",
    "icon-32.png",
    "icon-80.png",
    "icon-color.png",
    "icon-outline.png",
  ];

  for (const targetAssetsDir of targetAssetsDirs) {
    await fs.mkdir(targetAssetsDir, { recursive: true });

    for (const iconName of iconNames) {
      await fs.copyFile(
        path.join(iconsDir, iconName),
        path.join(targetAssetsDir, iconName),
      );
    }
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
      await copyEntry(
        path.join(sourcePath, child),
        path.join(targetPath, child),
      );
    }

    return;
  }

  await fs.copyFile(sourcePath, targetPath);
}

async function rewriteHostedHtmlFiles(baseUrl) {
  const htmlPaths = await collectFilesByExtension(webDir, ".html");
  const basePath = normalizeHostedPathname(new URL(baseUrl).pathname);

  for (const htmlPath of htmlPaths) {
    const html = await fs.readFile(htmlPath, "utf8");
    const rewrittenHtml = rewriteHostedHtmlUrls(html, basePath);

    if (rewrittenHtml !== html) {
      await fs.writeFile(htmlPath, rewrittenHtml);
    }
  }
}

async function validateHostedEntries(baseUrl) {
  for (const hostedEntry of hostedEntries) {
    const entryPath = path.join(webDir, hostedEntry);

    try {
      await fs.access(entryPath);
    } catch {
      throw new Error(`Deployment bundle is missing ${hostedEntry}.`);
    }
  }

  const htmlPaths = await collectFilesByExtension(webDir, ".html");
  const basePath = normalizeHostedPathname(new URL(baseUrl).pathname);

  for (const htmlPath of htmlPaths) {
    const html = await fs.readFile(htmlPath, "utf8");
    const hostedUrls = extractHostedAssetUrls(html, basePath);

    for (const hostedUrl of hostedUrls) {
      const relativePath = hostedUrl.slice(basePath.length).replace(/^\/+/, "");

      if (!relativePath) {
        continue;
      }

      const assetPath = path.join(webDir, ...relativePath.split("/"));

      try {
        await fs.access(assetPath);
      } catch {
        const relativeHtmlPath = path
          .relative(webDir, htmlPath)
          .replaceAll(path.sep, "/");
        throw new Error(
          `Hosted asset ${hostedUrl} referenced by ${relativeHtmlPath} is missing from the deployment bundle.`,
        );
      }
    }
  }
}

async function collectFilesByExtension(rootDir, extension) {
  const matchingPaths = [];
  const entries = await fs.readdir(rootDir, { withFileTypes: true });

  for (const entry of entries) {
    const entryPath = path.join(rootDir, entry.name);

    if (entry.isDirectory()) {
      matchingPaths.push(
        ...(await collectFilesByExtension(entryPath, extension)),
      );
      continue;
    }

    if (entry.name.endsWith(extension)) {
      matchingPaths.push(entryPath);
    }
  }

  return matchingPaths;
}

function rewriteHostedHtmlUrls(html, basePath) {
  return html.replace(
    /(src|href)=(['"])(\/[^'"]+)\2/g,
    (match, attributeName, quote, absolutePath) => {
      const rewrittenPath = normalizeHostedAssetReference(
        absolutePath,
        basePath,
      );

      if (!rewrittenPath) {
        return match;
      }

      return `${attributeName}=${quote}${rewrittenPath}${quote}`;
    },
  );
}

function extractHostedAssetUrls(html, basePath) {
  const urls = [];

  html.replace(
    /(src|href)=(['"])(\/[^'"]+)\2/g,
    (match, _attributeName, _quote, absolutePath) => {
      const rewrittenPath = normalizeHostedAssetReference(
        absolutePath,
        basePath,
      );

      if (rewrittenPath) {
        urls.push(rewrittenPath);
      }

      return match;
    },
  );

  return urls;
}

function normalizeHostedAssetReference(absolutePath, basePath) {
  const normalizedAbsolutePath = absolutePath.replace(/^\/+/, "");
  const hostedSubpath = findHostedSubpath(normalizedAbsolutePath);

  if (!hostedSubpath) {
    return null;
  }

  return joinHostedPath(basePath, hostedSubpath);
}

function findHostedSubpath(assetPath) {
  for (const hostedAssetRoot of hostedAssetRoots) {
    if (assetPath.startsWith(hostedAssetRoot)) {
      return assetPath;
    }

    const marker = `/${hostedAssetRoot}`;
    const markerIndex = assetPath.indexOf(marker);

    if (markerIndex >= 0) {
      return assetPath.slice(markerIndex + 1);
    }
  }

  return null;
}

function normalizeHostedPathname(pathname) {
  const normalizedPathname =
    pathname === "/" ? "/" : pathname.replace(/\/+$/, "");
  return normalizedPathname.startsWith("/")
    ? normalizedPathname
    : `/${normalizedPathname}`;
}

function joinHostedPath(basePath, relativePath) {
  const trimmedRelativePath = relativePath.replace(/^\/+/, "");

  if (basePath === "/") {
    return `/${trimmedRelativePath}`;
  }

  return `${basePath}/${trimmedRelativePath}`;
}

function rewriteManifestHostedUrls(value, baseUrl) {
  if (typeof value === "string") {
    return rewriteHostedUrl(value, baseUrl);
  }

  if (Array.isArray(value)) {
    return value.map((item) => rewriteManifestHostedUrls(item, baseUrl));
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, nestedValue]) => [
        key,
        rewriteManifestHostedUrls(nestedValue, baseUrl),
      ]),
    );
  }

  return value;
}

function rewriteHostedUrl(value, baseUrl) {
  return value.replace(/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?/i, baseUrl);
}

function createDeploymentGuide(baseUrl) {
  return `# Deployment Bundle\n\nThis bundle is intended for Outlook on Windows and Mac.\n\n## Contents\n- manifest.json: Production manifest for sideloading or catalog deployment\n- web/: Static web assets that must be hosted over HTTPS\n\n## Host The Web Assets\n1. Upload the contents of web/ to an HTTPS host.\n2. Preserve the existing folder structure so the task pane stays available at ${baseUrl}/src/taskpane/index.html.\n3. Confirm these URLs are reachable before sideloading:\n   - ${baseUrl}/src/taskpane/index.html\n   - ${baseUrl}/src/demo/index.html\n   - ${baseUrl}/assets/icon-16.png\n   - ${baseUrl}/assets/icon-32.png\n   - ${baseUrl}/assets/icon-80.png\n\n## Deploy In Outlook\n### Windows and Mac\n1. Open Outlook.\n2. Go to Get Add-ins or Apps.\n3. Choose My add-ins or My Apps.\n4. Use Add a custom add-in or Upload custom apps.\n5. Select manifest.json from this folder.\n\n## Centralized Deployment\nUse manifest.json after the web assets are hosted and validated. The same manifest can be used for admin deployment or AppSource submission workflows.\n`;
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
