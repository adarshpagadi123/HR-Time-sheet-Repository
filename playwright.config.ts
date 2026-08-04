import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  testMatch: '**/*.{spec,test,script}.ts',
  retries: 0,       // never restart the test from scratch
  workers: 1,       // one browser session only
  use: {
    baseURL: 'https://www.us.fieldglass.cloud.sap',
    headless: false,
    actionTimeout: 0,
    navigationTimeout: 0,
  },
  timeout: 0,
});
