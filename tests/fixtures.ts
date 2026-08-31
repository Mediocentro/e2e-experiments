import { test as base, type Page } from '@playwright/test';

import { attachToTaskPane } from '../src/word-session.js';

interface WordFixtures {
  taskPanePage: Page;
}

export const test = base.extend<WordFixtures>({
  taskPanePage: async ({ browserName }, use) => {
    if (browserName !== 'chromium') {
      throw new Error('The Word WebView2 fixture requires Chromium CDP.');
    }
    const attached = await attachToTaskPane();
    try {
      await use(attached.page);
    } finally {
      await attached.browser.close();
    }
  },
});

export { expect } from '@playwright/test';
