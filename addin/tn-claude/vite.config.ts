import path from "node:path"
import fs from "node:fs"
import os from "node:os"
import { defineConfig } from "vite"
import react from "@vitejs/plugin-react"
import tailwindcss from "@tailwindcss/vite"

// Office add-ins must be served over HTTPS: office-addin-dev-certs puts a localhost cert in ~/.office-addin-dev-certs
const certDir = path.join(os.homedir(), ".office-addin-dev-certs")
const https = fs.existsSync(path.join(certDir, "localhost.crt"))
  ? { cert: fs.readFileSync(path.join(certDir, "localhost.crt")), key: fs.readFileSync(path.join(certDir, "localhost.key")) }
  : undefined

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: { alias: { "@": path.resolve(import.meta.dirname, "./src") } },
  server: {
    port: 3000,
    strictPort: true,
    https,
    proxy: { "/api": { target: "http://127.0.0.1:3001", changeOrigin: true } },
  },
})
