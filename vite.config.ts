import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  root: './demo',
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  server: {
    port: 5173,
    open: false,
    // Proxy y-webrtc signaling so the browser can reach it same-origin.
    // Run `npm run signaling` (port 4444) alongside the demo.
    proxy: {
      '/signaling': {
        target: 'http://localhost:4444',
        changeOrigin: true,
        ws: true,
      },
    },
  },
  build: {
    outDir: '../dist/demo',
  },
});
