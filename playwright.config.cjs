const { defineConfig } = require('@playwright/test');
module.exports = defineConfig({
  testDir: './preview', testMatch: '*.spec.cjs', timeout: 30000, use: { baseURL: 'http://127.0.0.1:4318', headless: true },
  webServer: { command: 'node preview/serve.cjs', url: 'http://127.0.0.1:4318', reuseExistingServer: !process.env.CI },
  outputDir: 'scratch/browser-results'
});
