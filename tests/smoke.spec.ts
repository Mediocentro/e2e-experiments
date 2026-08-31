import { optionalEnvironment } from '../src/config.js';
import { expect, test } from './fixtures.js';

test('@smoke starts inside Microsoft Word', async ({ taskPanePage }) => {
  await taskPanePage.waitForFunction(
    () =>
      typeof (
        window as unknown as {
          Office?: { onReady?: unknown };
        }
      ).Office?.onReady === 'function',
  );

  const host = await taskPanePage.evaluate(async () => {
    const office = (
      window as unknown as {
        Office: {
          HostType: { Word: string };
          onReady(): Promise<{ host: string }>;
        };
      }
    ).Office;
    const information = await office.onReady();
    return { actual: information.host, expected: office.HostType.Word };
  });

  expect(host.actual).toBe(host.expected);
  await expect(taskPanePage.locator('body')).toBeVisible();

  const readySelector = optionalEnvironment('WORD_E2E_READY_SELECTOR');
  if (readySelector) {
    await expect(taskPanePage.locator(readySelector)).toBeVisible();
  }
});
