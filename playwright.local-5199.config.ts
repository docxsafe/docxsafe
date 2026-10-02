// Temporary config: run e2e against a dev server on port 5199
import base from './playwright.config';
import { defineConfig } from '@playwright/test';

export default defineConfig({
  ...base,
  use: { ...base.use, baseURL: 'http://localhost:5199' },
  // Use the locally installed Google Chrome (no browser download needed)
  projects: [{ name: 'chromium', use: { ...base.projects?.[0]?.use, channel: 'chrome' } }],
  webServer: {
    command: 'npx vite --port 5199 --strictPort',
    url: 'http://localhost:5199',
    reuseExistingServer: true,
    timeout: 60 * 1000,
  },
});
