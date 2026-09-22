#!/usr/bin/env node
/**
 * scripts/bump-version.mjs
 *
 * Increments the build (patch) segment of the version in both
 * package.json and manifest.json when run explicitly.
 *
 * Version format: MAJOR.MINOR.PATCH
 * This script increments PATCH by 1.
 *
 * Usage:
 *   node scripts/bump-version.mjs
 *
 * Optional flag to also bump MINOR (resets PATCH to 0):
 *   node scripts/bump-version.mjs --minor
 *
 * Optional flag to also bump MAJOR (resets MINOR and PATCH to 0):
 *   node scripts/bump-version.mjs --major
 */

import { readFileSync, writeFileSync } from "fs";
import { spawnSync } from "child_process";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, "..");

const args = process.argv.slice(2);
const bumpMajor = args.includes("--major");
const bumpMinor = args.includes("--minor");

// ── Read files ────────────────────────────────────────────────────────────────

const pkgPath = resolve(root, "package.json");
const manifestPath = resolve(root, "manifest.json");

const pkg = JSON.parse(readFileSync(pkgPath, "utf8"));
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));

// ── Increment version ─────────────────────────────────────────────────────────

const [major, minor, patch] = pkg.version.split(".").map(Number);

let newVersion;
if (bumpMajor) {
  newVersion = `${major + 1}.0.0`;
} else if (bumpMinor) {
  newVersion = `${major}.${minor + 1}.0`;
} else {
  newVersion = `${major}.${minor}.${patch + 1}`;
}

// ── Write back ────────────────────────────────────────────────────────────────

pkg.version = newVersion;
writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n", "utf8");

manifest.version = newVersion;
writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n", "utf8");

// Keep package-lock.json in sync with package.json version changes.
const npmExecutable = process.platform === "win32" ? "npm.cmd" : "npm";
const lockUpdate = spawnSync(
  npmExecutable,
  ["install"],
  {
    cwd: root,
    stdio: "inherit",
  },
);

if (lockUpdate.status !== 0) {
  console.error("✖ Failed to update package-lock.json after version bump.");
  process.exit(lockUpdate.status ?? 1);
}

console.log(`✔ Version bumped to ${newVersion}`);
