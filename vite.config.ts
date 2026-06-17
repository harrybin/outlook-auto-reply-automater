import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";
import fs from "fs";
import os from "os";
import type { ServerOptions } from "https";

// Load dev certificates created by office-addin-dev-certs (if available)
function getHttpsConfig(): ServerOptions | undefined {
  const certPath = resolve(os.homedir(), ".office-addin-dev-certs");
  const keyFile = resolve(certPath, "localhost.key");
  const certFile = resolve(certPath, "localhost.crt");
  if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
    return { key: fs.readFileSync(keyFile), cert: fs.readFileSync(certFile) };
  }
  return undefined;
}

export default defineConfig(() => ({
  plugins: [react()],
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
    strictPort: true,
    https: undefined,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  preview: {
    port: 3001,
    https: undefined,
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
