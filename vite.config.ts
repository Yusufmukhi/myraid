import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { viteSingleFile } from 'vite-plugin-singlefile'

// Set SINGLEFILE=true only for standalone offline bundle distribution.
// For production deployments (Vercel, Netlify, Cloudflare), standard modular bundles
// allow strict CSP with 'script-src self' without requiring 'unsafe-inline'.
const isSingleFile = process.env.SINGLEFILE === 'true'

export default defineConfig({
  plugins: [react(), ...(isSingleFile ? [viteSingleFile()] : [])],
  build: { target: 'es2020', chunkSizeWarningLimit: 4000, sourcemap: false },
  server: { host: '0.0.0.0', port: 5173 },
  preview: { host: '0.0.0.0', port: 5173 },
})
