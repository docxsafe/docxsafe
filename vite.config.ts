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
    // Proxy collaboration WebSocket server (npm run collaboration:server → :1234)
    proxy: {
      '/collab-ws': {
        target: 'http://localhost:1234',
        changeOrigin: true,
        ws: true,
        rewrite: (path) => path.replace(/^\/collab-ws/, ''),
      },
    },
  },
  build: {
    outDir: '../dist/demo',
  },
});
