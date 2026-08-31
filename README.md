# Minimal Playwright E2E harness for a Word add-in

This folder is a standalone Windows test harness for attaching Playwright to a
Microsoft Word task-pane add-in running in WebView2. It contains no add-in ID,
tenant ID, credentials, production URL, document, certificate, or generated
test artifact.

It supports two modes:

- `local`: starts your development server, sideloads your manifest, and lets
  Microsoft's debugging package create a task-pane-enabled disposable DOCX.
- `deployed`: opens an existing DOCX whose metadata auto-opens an add-in that
  is already assigned to the test user.

## Prerequisites

- Windows with desktop Microsoft Word and the WebView2 runtime.
- Node.js supported by your repository.
- A dedicated non-admin Office test user.
- No existing Word process when the suite starts.
- Loopback port 9222, or another configured CDP port, must be available.
- In deployed mode, the test user must already have the add-in assigned and
  the test DOCX must contain valid auto-open web-extension metadata.

## Install

Copy this folder into your repository, open a terminal in it, and run:

```powershell
npm install
npx playwright install chromium
```

`npm install` creates the lockfile expected by your destination repository.
This template intentionally does not include a generated lockfile.

Copy the environment example:

```powershell
Copy-Item .env.example .env
```

Never commit `.env`, test documents, certificates, HAR files, authentication
tokens, or generated artifacts.

## Local sideload mode

Configure at least these values in `.env`:

```dotenv
WORD_E2E_MODE=local
WORD_E2E_TASKPANE_URL=https://localhost:3000/taskpane.html
WORD_E2E_MANIFEST_PATH=../path-to-addin/manifest.xml
WORD_E2E_DEV_SERVER_COMMAND=npm run dev
WORD_E2E_DEV_SERVER_CWD=../path-to-addin
WORD_E2E_DEV_SERVER_URL=https://localhost:3000/taskpane.html
WORD_E2E_DEV_SERVER_PORT=3000
```

Your development server must use trusted HTTPS. If the destination repository
does not already manage a localhost certificate, run:

```powershell
npx office-addin-dev-certs install
```

The harness deliberately does not pass `--document` to
`office-addin-debugging`. When a caller supplies an ordinary DOCX, Microsoft's
helper copies it unchanged and does not inject the task-pane metadata required
for automatic opening.

## Deployed add-in mode

Configure:

```dotenv
WORD_E2E_MODE=deployed
WORD_E2E_TASKPANE_URL=https://your-test-origin.example/taskpane
WORD_E2E_DOCUMENT_PATH=C:\path\to\auto-open-test-document.docx
```

Local-server settings in `.env` are ignored when `WORD_E2E_MODE=deployed`.

In this mode the harness:

1. Verifies Word is closed.
2. Configures loopback-only WebView2 remote debugging.
3. Opens the specified document in a new Word process.
4. Waits for the deployed task-pane target.
5. Attaches Playwright over CDP.
6. Closes only the exact test document during teardown.

It does not sideload a manifest or start a local web server unless
`WORD_E2E_DEV_SERVER_COMMAND` is also configured.

## Run

Run the smoke test:

```powershell
npx playwright test --grep @smoke
```

Run all specs:

```powershell
npx playwright test
```

Other useful commands:

```powershell
npx playwright test tests/smoke.spec.ts
npx playwright show-report test-output/report
npm run typecheck
```

Set `WORD_E2E_KEEP_OPEN=1` only for interactive diagnosis. In local mode, stop
the registration manually afterward:

```powershell
npx office-addin-debugging stop ..\path-to-addin\manifest.xml
```

## Add feature specs

Create separate independently runnable spec files:

```text
tests/
  smoke.spec.ts
  routing.spec.ts
  feature-a.spec.ts
  feature-b.spec.ts
  export.spec.ts
```

Import the shared Word fixture:

```ts
import { expect, test } from './fixtures.js';

test('runs feature A', async ({ taskPanePage }) => {
  await taskPanePage.getByRole('button', { name: 'Start feature A' }).click();
  await expect(taskPanePage.getByRole('status')).toContainText('Complete');
});
```

Keep every test independent. A test should create or reset its own document and
backend data instead of relying on a previous spec. The harness intentionally
uses one worker because all specs share one desktop Word host.

## Configuration reference

| Variable                        | Required | Purpose                                           |
| ------------------------------- | -------- | ------------------------------------------------- |
| `WORD_E2E_MODE`                 | Yes      | `local` or `deployed`                             |
| `WORD_E2E_TASKPANE_URL`         | Yes      | Origin and path used to find the task pane        |
| `WORD_E2E_READY_SELECTOR`       | No       | CSS selector asserted by the smoke test           |
| `WORD_E2E_CDP_PORT`             | No       | Loopback CDP port; defaults to 9222               |
| `WORD_E2E_MANIFEST_PATH`        | Local    | Development manifest path                         |
| `WORD_E2E_DEV_SERVER_COMMAND`   | Local    | Command Playwright uses to start the app          |
| `WORD_E2E_DEV_SERVER_CWD`       | Local    | Working directory for the server command          |
| `WORD_E2E_DEV_SERVER_URL`       | Local    | URL Playwright waits for before setup             |
| `WORD_E2E_DEV_SERVER_PORT`      | Local    | Port passed to `office-addin-debugging`           |
| `WORD_E2E_OFFICE_DEBUGGING_CLI` | No       | CLI override for nonstandard monorepo layouts     |
| `WORD_E2E_DOCUMENT_PATH`        | Deployed | Auto-open DOCX to launch in Word                  |
| `WORD_E2E_WORD_EXE`             | No       | Explicit Word executable if registry lookup fails |
| `WORD_E2E_KEEP_OPEN`            | No       | Set to 1 to skip automatic cleanup                |
| `WORD_E2E_KEEP_ARTIFACTS`       | No       | Set to 1 to preserve successful output            |
| `WORD_E2E_LOG_LEVEL`            | No       | Set to `debug` for lifecycle messages             |

## Troubleshooting

### No CDP endpoint

- Confirm Word was fully closed before the run.
- Confirm the WebView2 runtime is installed.
- Confirm the selected CDP port is free.
- Confirm enterprise policy permits loopback WebView2 debugging on the QA
  machine.

### CDP exists but no matching task pane appears

- Confirm `WORD_E2E_TASKPANE_URL` matches the actual origin and path.
- In local mode, confirm the manifest points to the same HTTPS URL.
- In deployed mode, confirm the add-in is assigned to the Office test user and
  the DOCX contains the correct auto-open metadata.
- Check `test-output/word/test.log` and the sanitized target list in the error.

### HTTPS or blank task pane

- Trust the development certificate.
- Open the task-pane URL directly and confirm there is no certificate warning.
- Confirm all manifest resources use HTTPS.

## Security notes

- CDP binds only to `127.0.0.1`.
- Run this on a dedicated QA machine or VM, not a normal employee workstation.
- Do not use an administrator account for add-in workflows.
- Do not commit `.env`, documents, test output, certificates, tokens, or HARs.
- Teardown targets only the exact document recorded during setup and leaves
  other Word documents open.
