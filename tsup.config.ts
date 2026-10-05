import { defineConfig } from 'tsup';

const isProd = process.env.NODE_ENV === 'production';

export default defineConfig([
  // Main builds (without shebang)
  {
    entry: {
      index: 'src/index.ts',
      headless: 'src/headless.ts',
      'core-plugins': 'src/core-plugins/index.ts',
      mcp: 'src/mcp/index.ts',
      collaboration: 'src/collaboration/index.ts',
      'collaborative-docs': 'src/collaborative-docs/index.ts',
      'collaborative-docs-server': 'src/collaborative-docs/server/index.ts',
    },
    format: ['cjs', 'esm'],
    dts: true,
    splitting: true, // Enable code splitting for tree-shaking
    sourcemap: !isProd, // Disable source maps in production to reduce package size
    clean: true,
    treeshake: true, // Enable tree-shaking
    minify: true, // Minify the output
    external: [
      'react',
      'react-dom',
      // Keep Yjs stack external so consumers can dedupe versions
      'yjs',
      'y-webrtc',
      'y-websocket',
      'y-prosemirror',
      'y-protocols',
      'lib0',
      'ws',
      '@aws-sdk/client-s3',
    ],
    injectStyle: false,
  },
  // CLI build (with shebang) - bundles all deps for standalone use
  {
    entry: {
      'mcp-cli': 'src/mcp/cli.ts',
      'collab-server-cli': 'src/collaborative-docs/server/cli.ts',
    },
    format: ['esm'],
    dts: true,
    splitting: false,
    sourcemap: !isProd, // Disable source maps in production
    clean: false, // Don't clean since main build already did
    treeshake: true,
    minify: true,
    external: ['react', 'react-dom'],
    injectStyle: false,
    banner: {
      js: '#!/usr/bin/env node',
    },
  },
]);
