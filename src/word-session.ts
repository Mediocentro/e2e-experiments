import { chromium, type Browser, type Page } from '@playwright/test';

import { cdpEndpoint, debugLog, taskPaneUrl } from './config.js';

export interface AttachedTaskPane {
  browser: Browser;
  page: Page;
}

interface CdpTarget {
  type?: string;
  url?: string;
}

function matchesTaskPane(candidateUrl: string | undefined): boolean {
  if (!candidateUrl) {
    return false;
  }
  try {
    const expected = new URL(taskPaneUrl());
    const candidate = new URL(candidateUrl);
    return (
      candidate.origin === expected.origin &&
      candidate.pathname.startsWith(expected.pathname)
    );
  } catch {
    return false;
  }
}

async function listTargets(endpoint: string): Promise<CdpTarget[]> {
  try {
    const response = await fetch(`${endpoint}/json/list`);
    if (!response.ok) {
      return [];
    }
    return (await response.json()) as CdpTarget[];
  } catch {
    return [];
  }
}

function safeTargetSummary(targets: CdpTarget[]): string {
  return JSON.stringify(
    targets.map((target) => {
      try {
        const url = target.url ? new URL(target.url) : undefined;
        return {
          type: target.type,
          url: url ? `${url.origin}${url.pathname}` : undefined,
        };
      } catch {
        return { type: target.type, url: '<unparseable>' };
      }
    }),
  );
}

export async function attachToTaskPane(
  timeoutMs = 90_000,
): Promise<AttachedTaskPane> {
  const endpoint = cdpEndpoint();
  const deadline = Date.now() + timeoutMs;
  let targets: CdpTarget[] = [];

  while (Date.now() < deadline) {
    targets = await listTargets(endpoint);
    if (
      targets.some(
        (target) => target.type === 'page' && matchesTaskPane(target.url),
      )
    ) {
      break;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }

  if (
    !targets.some(
      (target) => target.type === 'page' && matchesTaskPane(target.url),
    )
  ) {
    throw new Error(
      `No matching Word task-pane target appeared at ${endpoint}. Targets: ${safeTargetSummary(targets)}`,
    );
  }

  debugLog(`Attaching Playwright to ${endpoint}.`);
  const browser = await chromium.connectOverCDP(endpoint);
  const page = browser
    .contexts()
    .flatMap((context) => context.pages())
    .find((candidate) => matchesTaskPane(candidate.url()));

  if (!page) {
    await browser.close();
    throw new Error(
      'CDP exposed the task-pane target, but Playwright could not resolve it.',
    );
  }

  return { browser, page };
}
