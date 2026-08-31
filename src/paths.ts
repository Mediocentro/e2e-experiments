import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

import { projectRoot } from './config.js';

export const paths = {
  projectRoot,
  closeDocumentScript: resolve(projectRoot, 'scripts/close-test-document.ps1'),
  outputRoot: resolve(projectRoot, 'test-output/word'),
  inputDocument: resolve(projectRoot, 'test-output/word/input.docx'),
  session: resolve(projectRoot, 'test-output/word/session.json'),
  testLog: resolve(projectRoot, 'test-output/word/test.log'),
} as const;

export async function ensureOutputDirectory(): Promise<void> {
  await mkdir(paths.outputRoot, { recursive: true });
}
