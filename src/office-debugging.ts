import { createRequire } from 'node:module';
import { resolve } from 'node:path';

import { optionalEnvironment, projectRoot } from './config.js';

const require = createRequire(import.meta.url);

export function officeAddinDebuggingCommand(): {
  command: string;
  prefixArguments: string[];
} {
  const configuredCli = optionalEnvironment('WORD_E2E_OFFICE_DEBUGGING_CLI');
  return {
    command: process.execPath,
    prefixArguments: [
      configuredCli
        ? resolve(projectRoot, configuredCli)
        : require.resolve('office-addin-debugging/cli.js'),
    ],
  };
}
