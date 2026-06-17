import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";
import { resolve } from "path";
import fs from "fs";
import os from "os";
import type { ServerOptions } from "https";

// Load dev certificates created by office-addin-dev-certs (if available).
// Returns the cert config when found so basicSsl plugin is not needed.
function getHttpsConfig(): ServerOptions | undefined {
  const certPath = resolve(os.homedir(), ".office-addin-dev-certs");
  const keyFile = resolve(certPath, "localhost.key");
  const certFile = resolve(certPath, "localhost.crt");
  if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
    return { key: fs.readFileSync(keyFile), cert: fs.readFileSync(certFile) };
  }
  return undefined;
}

const httpsConfig = getHttpsConfig();

export default defineConfig(() => ({
  // basicSsl provides a self-signed cert when office-addin-dev-certs are absent;
  // Outlook add-ins require HTTPS so this ensures the dev server always uses it.
  plugins: [react(), ...(httpsConfig ? [] : [basicSsl()])],
  base: process.env.VITE_BASE_PATH ?? "/",
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "src/taskpane/index.html"),
      },
    },
    sourcemap: true,
  },
  server: {
    port: 3000,
    strictPort: false,
    // basicSsl plugin sets https when httpsConfig is absent; otherwise use explicit certs.
    https: httpsConfig,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  preview: {
    port: 3001,
    https: httpsConfig,
  },
  test: {
    globals: true,
    environment: "jsdom",
    setupFiles: ["./tests/setup.ts"],
    coverage: {
      provider: "v8",
      reporter: ["text", "html", "lcov"],
      include: ["src/**/*.{ts,tsx}"],
      exclude: ["src/**/*.d.ts", "src/taskpane/index.tsx"],
    },
  },
}));
