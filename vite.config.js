import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'

export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    host: true, // bind to all network interfaces so other devices on the LAN can connect
    proxy: {
      // All /api/* requests route to the Express server.
      // This includes /api/scan (Claude proxy), /api/admin/*, and /api/station/*.
      // The Anthropic API is no longer called directly from the browser.
      '/api': { target: 'http://localhost:3001', changeOrigin: true },
    },
  },
})
