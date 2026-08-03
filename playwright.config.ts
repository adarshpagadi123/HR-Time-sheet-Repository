import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  testMatch: '**/*.{spec,test,script}.ts',
  use: {
    baseURL: 'https://www.us.fieldglass.cloud.sap',
    headless: false,
  },
  timeout: 120000,
});
