import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  testMatch: '**/*.{spec,test,script}.ts',
  globalSetup: './global-setup.ts', // kills leftover Chrome before every run
  retries: 0,
  workers: 1,          // single session — all WOIDs in one browser
  fullyParallel: true,
  use: {
    baseURL: 'https://www.us.fieldglass.cloud.sap',
    headless: false,
    actionTimeout: 0,
    navigationTimeout: 0,
  },
  timeout: 0,
});
