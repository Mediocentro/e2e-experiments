import { spawnSync } from 'node:child_process';
import { appendFile, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';

import { officeAddinDebuggingCommand } from './src/office-debugging.js';
import { paths } from './src/paths.js';

interface Session {
  mode: 'local' | 'deployed';
  launchedDocument: string;
  manifest?: string;
}

function run(command: string, arguments_: string[], cwd: string) {
  return spawnSync(command, arguments_, {
    cwd,
    env: process.env,
    encoding: 'utf8',
    windowsHide: true,
    maxBuffer: 10 * 1024 * 1024,
  });
}

export default async function globalTeardown(): Promise<void> {
  if (process.env.WORD_E2E_KEEP_OPEN === '1') {
    await appendFile(
      paths.testLog,
      '\n--- teardown ---\nWORD_E2E_KEEP_OPEN=1; Word was left open intentionally.\n',
      'utf8',
    );
    return;
  }

  let sessionText: string;
  try {
    sessionText = await readFile(paths.session, 'utf8');
  } catch (error) {
    const code =
      typeof error === 'object' && error !== null && 'code' in error
        ? error.code
        : undefined;
    if (code === 'ENOENT') {
      return;
    }
    throw error;
  }
  const session = JSON.parse(sessionText) as Session;
  const failures: string[] = [];

  if (session.mode === 'local' && session.manifest) {
    const debuggingCommand = officeAddinDebuggingCommand();
    const stop = run(
      debuggingCommand.command,
      [...debuggingCommand.prefixArguments, 'stop', session.manifest],
      dirname(session.manifest),
    );
    await appendFile(
      paths.testLog,
      `\n--- stop sideload session ---\n${stop.stdout ?? ''}${stop.stderr ?? ''}`,
      'utf8',
    );
    if (stop.error || stop.status !== 0) {
      failures.push('Unable to stop the sideload registration.');
    }
  }

  const cleanup = run(
    'powershell',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      paths.closeDocumentScript,
      '-DocumentPath',
      session.launchedDocument,
    ],
    paths.projectRoot,
  );
  await appendFile(
    paths.testLog,
    `\n--- document cleanup ---\n${cleanup.stdout ?? ''}${cleanup.stderr ?? ''}`,
    'utf8',
  );
  if (cleanup.error || cleanup.status !== 0) {
    failures.push('Unable to close the disposable Word document.');
  }

  if (failures.length > 0) {
    throw new Error(`${failures.join(' ')} See ${paths.testLog}.`);
  }
}
