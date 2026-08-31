import { loadEnvFile } from 'node:process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const projectRoot = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '..',
);

try {
  loadEnvFile(resolve(projectRoot, '.env'));
} catch (error) {
  const code =
    typeof error === 'object' && error !== null && 'code' in error
      ? error.code
      : undefined;
  if (code !== 'ENOENT') {
    throw error;
  }
}

export type WordE2EMode = 'local' | 'deployed';

export function optionalEnvironment(name: string): string | undefined {
  const value = process.env[name]?.trim();
  return value ? value : undefined;
}

export function requiredEnvironment(name: string): string {
  const value = optionalEnvironment(name);
  if (!value) {
    throw new Error(`${name} is required.`);
  }
  return value;
}

export function mode(): WordE2EMode {
  const value = requiredEnvironment('WORD_E2E_MODE');
  if (value !== 'local' && value !== 'deployed') {
    throw new Error('WORD_E2E_MODE must be local or deployed.');
  }
  return value;
}

export function taskPaneUrl(): string {
  const value = requiredEnvironment('WORD_E2E_TASKPANE_URL');
  const parsed = new URL(value);
  if (parsed.protocol !== 'https:') {
    throw new Error('WORD_E2E_TASKPANE_URL must use HTTPS.');
  }
  return value;
}

export function cdpPort(): number {
  const value = Number(optionalEnvironment('WORD_E2E_CDP_PORT') ?? '9222');
  if (!Number.isInteger(value) || value < 1024 || value > 65_535) {
    throw new Error(
      'WORD_E2E_CDP_PORT must be an integer from 1024 through 65535.',
    );
  }
  return value;
}

export function cdpEndpoint(): string {
  return `http://127.0.0.1:${cdpPort()}`;
}

export function resolveConfiguredPath(name: string): string {
  return resolve(projectRoot, requiredEnvironment(name));
}

export function debugLog(message: string): void {
  if (optionalEnvironment('WORD_E2E_LOG_LEVEL') === 'debug') {
    console.log(`[word-e2e] ${message}`);
  }
}
