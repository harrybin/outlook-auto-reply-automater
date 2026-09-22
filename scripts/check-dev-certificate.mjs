import { spawnSync } from "node:child_process";

const isWindows = process.platform === "win32";
const command = isWindows ? (process.env.ComSpec ?? "cmd.exe") : "npx";
const args = isWindows
  ? ["/d", "/s", "/c", "npx --no-install office-addin-dev-certs verify"]
  : ["--no-install", "office-addin-dev-certs", "verify"];
const result = spawnSync(command, args, {
  stdio: "inherit",
});

if (result.error || result.status !== 0) {
  console.error(
    "\nLokales HTTPS-Zertifikat ist nicht vertrauenswürdig oder nicht installiert. Outlook kann den Taskpane dadurch nicht laden.",
  );
  console.error("Bitte einmal ausführen:");
  console.error("  npx office-addin-dev-certs install");
  console.error(
    "Danach Outlook vollständig schließen und npm run start erneut ausführen.\n",
  );
  process.exit(1);
}

console.log("Lokales HTTPS-Zertifikat ist vertrauenswürdig.\n");
