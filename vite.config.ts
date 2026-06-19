import { defineConfig } from "vite";
import { readFileSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { resolve } from "node:path";

/**
 * Office add-ins must be served over HTTPS, even on localhost.
 * `npm run certs` (office-addin-dev-certs) installs a trusted localhost CA at
 * ~/.office-addin-dev-certs/. If present we use it; otherwise we fall back to
 * plain HTTP so `vite build` and non-Office browser work still function.
 */
function devCerts(): { key: Buffer; cert: Buffer } | undefined {
  const dir = resolve(homedir(), ".office-addin-dev-certs");
  const key = resolve(dir, "localhost.key");
  const cert = resolve(dir, "localhost.crt");
  if (existsSync(key) && existsSync(cert)) {
    return { key: readFileSync(key), cert: readFileSync(cert) };
  }
  return undefined;
}

export default defineConfig({
  server: {
    port: 3000,
    strictPort: true,
    https: devCerts(),
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        authCallback: resolve(__dirname, "auth-callback.html"),
        authStart: resolve(__dirname, "auth-start.html"),
      },
    },
  },
});
