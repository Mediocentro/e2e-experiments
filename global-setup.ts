import { spawn, spawnSync } from 'node:child_process';
import { appendFile, copyFile, rm, stat, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

import {
  cdpEndpoint,
  cdpPort,
  debugLog,
  mode,
  optionalEnvironment,
  requiredEnvironment,
  resolveConfiguredPath,
  type WordE2EMode,
} from './src/config.js';
import { officeAddinDebuggingCommand } from './src/office-debugging.js';
import { ensureOutputDirectory, paths } from './src/paths.js';

interface CommandResult {
  error?: Error;
  status: number | null;
  stdout: string;
  stderr: string;
}

interface Session {
  mode: WordE2EMode;
  launchedDocument: string;
  manifest?: string;
  cdpEndpoint: string;
  startedAt: string;
}

function run(
  command: string,
  arguments_: string[],
  options: { cwd?: string; environment?: NodeJS.ProcessEnv } = {},
): CommandResult {
  const result = spawnSync(command, arguments_, {
    cwd: options.cwd ?? paths.projectRoot,
    env: options.environment ?? process.env,
    encoding: 'utf8',
    windowsHide: true,
    maxBuffer: 10 * 1024 * 1024,
  });
  return {
    error: result.error,
    status: result.status,
    stdout: result.stdout ?? '',
    stderr: result.stderr ?? '',
  };
}

function ensureWindowsAndClosedWord(): string {
  if (process.platform !== 'win32') {
    throw new Error('Desktop Word E2E requires Windows.');
  }
  const result = run('tasklist', [
    '/FI',
    'IMAGENAME eq WINWORD.EXE',
    '/FO',
    'CSV',
    '/NH',
  ]);
  if (result.error || result.status !== 0) {
    throw new Error('Unable to check for running Word processes.');
  }
  if (/"WINWORD\.EXE"/i.test(result.stdout)) {
    throw new Error(
      'Close all running Microsoft Word windows before starting the E2E suite.',
    );
  }
  return result.stdout;
}

function debuggingEnvironment(): NodeJS.ProcessEnv {
  const existing = process.env.WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS?.trim();
  return {
    ...process.env,
    WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: [
      existing,
      `--remote-debugging-port=${cdpPort()}`,
      '--remote-debugging-address=127.0.0.1',
    ]
      .filter(Boolean)
      .join(' '),
  };
}

function findWordExecutable(): string {
  const configured = optionalEnvironment('WORD_E2E_WORD_EXE');
  if (configured) {
    return resolveConfiguredPath('WORD_E2E_WORD_EXE');
  }

  const result = run('reg', [
    'query',
    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE',
    '/ve',
  ]);
  const match = result.stdout.match(/REG_SZ\s+(.+?WINWORD\.EXE)\s*$/im);
  if (result.error || result.status !== 0 || !match?.[1]) {
    throw new Error(
      'Unable to locate WINWORD.EXE. Set WORD_E2E_WORD_EXE explicitly.',
    );
  }
  return match[1].trim();
}

async function launchWord(
  executable: string,
  documentPath: string,
  environment: NodeJS.ProcessEnv,
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = spawn(executable, ['/n', documentPath], {
      cwd: paths.projectRoot,
      env: environment,
      detached: true,
      stdio: 'ignore',
      windowsHide: false,
    });
    child.once('error', reject);
    child.once('spawn', () => {
      child.unref();
      resolve();
    });
  });
}

async function localSetup(environment: NodeJS.ProcessEnv): Promise<Session> {
  const manifest = resolveConfiguredPath('WORD_E2E_MANIFEST_PATH');
  const port = requiredEnvironment('WORD_E2E_DEV_SERVER_PORT');
  const debuggingCommand = officeAddinDebuggingCommand();

  // Do not pass --document. The Microsoft helper injects task-pane metadata
  // only when it creates its bundled Word document template.
  const result = run(
    debuggingCommand.command,
    [
      ...debuggingCommand.prefixArguments,
      'start',
      manifest,
      'desktop',
      '--app',
      'word',
      '--no-debug',
      '--no-live-reload',
      '--dev-server-port',
      port,
    ],
    { cwd: dirname(manifest), environment },
  );
  await appendFile(
    paths.testLog,
    `\n--- sideload and launch ---\n${result.stdout}${result.stderr}`,
    'utf8',
  );
  if (
    result.error ||
    result.status !== 0 ||
    /Unable to (start debugging|sideload)/i.test(
      `${result.stdout}\n${result.stderr}`,
    )
  ) {
    throw new Error(`Unable to sideload the add-in. See ${paths.testLog}.`);
  }

  const launchedDocument = result.stdout.match(
    /^Launching word via (.+\.docx)\r?$/m,
  )?.[1];
  if (!launchedDocument) {
    throw new Error(
      `The sideload tool did not report its document path. See ${paths.testLog}.`,
    );
  }
  await copyFile(launchedDocument, paths.inputDocument);
  return {
    mode: 'local',
    launchedDocument,
    manifest,
    cdpEndpoint: cdpEndpoint(),
    startedAt: new Date().toISOString(),
  };
}

async function deployedSetup(environment: NodeJS.ProcessEnv): Promise<Session> {
  const launchedDocument = resolveConfiguredPath('WORD_E2E_DOCUMENT_PATH');
  await stat(launchedDocument);
  await copyFile(launchedDocument, paths.inputDocument);
  await launchWord(findWordExecutable(), launchedDocument, environment);
  await appendFile(
    paths.testLog,
    `\n--- deployed document launch ---\n${launchedDocument}\n`,
    'utf8',
  );
  return {
    mode: 'deployed',
    launchedDocument,
    cdpEndpoint: cdpEndpoint(),
    startedAt: new Date().toISOString(),
  };
}

export default async function globalSetup(): Promise<void> {
  await ensureOutputDirectory();
  await rm(paths.session, { force: true });
  const preflight = ensureWindowsAndClosedWord();
  await writeFile(
    paths.testLog,
    `--- Word preflight ---\n${preflight}`,
    'utf8',
  );

  const environment = debuggingEnvironment();
  debugLog(`Starting Word with CDP at ${cdpEndpoint()}.`);
  const session =
    mode() === 'local'
      ? await localSetup(environment)
      : await deployedSetup(environment);

  await writeFile(
    paths.session,
    `${JSON.stringify(session, null, 2)}\n`,
    'utf8',
  );
}
