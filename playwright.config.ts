import { resolve } from 'node:path';

import { defineConfig } from '@playwright/test';

import {
  mode,
  optionalEnvironment,
  projectRoot,
  requiredEnvironment,
  taskPaneUrl,
} from './src/config.js';

const devServerCommand =
  mode() === 'local'
    ? requiredEnvironment('WORD_E2E_DEV_SERVER_COMMAND')
    : undefined;
const devServerUrl =
  optionalEnvironment('WORD_E2E_DEV_SERVER_URL') ?? taskPaneUrl();
const devServerCwd = resolve(
  projectRoot,
  optionalEnvironment('WORD_E2E_DEV_SERVER_CWD') ?? '.',
);

export default defineConfig({
  testDir: './tests',
  fullyParallel: false,
  workers: 1,
  retries: 0,
  timeout: 120_000,
  expect: {
    timeout: 15_000,
  },
  globalSetup: './global-setup.ts',
  globalTeardown: './global-teardown.ts',
  outputDir: './test-output/playwright',
  preserveOutput:
    process.env.WORD_E2E_KEEP_ARTIFACTS === '1' ? 'always' : 'failures-only',
  reporter: [
    ['list'],
    ['html', { outputFolder: './test-output/report', open: 'never' }],
  ],
  use: {
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: devServerCommand
    ? {
        command: devServerCommand,
        cwd: devServerCwd,
        url: devServerUrl,
        ignoreHTTPSErrors:
          optionalEnvironment('WORD_E2E_IGNORE_HTTPS_ERRORS') === '1',
        reuseExistingServer: false,
        timeout: 120_000,
      }
    : undefined,
});
