## Internal reference (mapping)
Microsoft 365 Agents Toolkit (formerly Teams Toolkit) has been rebranded; users may still use old names.

| New name                                | Former name            |
|-----------------------------------------|------------------------|
| Microsoft 365 Agents Toolkit            | Teams Toolkit          |
| App Manifest                            | Teams app manifest     |
| Microsoft 365 Agents Playground         | Test Tool              |
| `m365agents.yml`                        | `teamsapp.yml`         |
| `@microsoft/m365agentstoolkit-cli` (command `atk`) | `@microsoft/teamsapp-cli` (command `teamsapp`) |

## Tooling & automation guidance (short)
- Prefer invoking toolkit helpers: `atk` (CLI) or the Visual Studio Code Microsoft 365 Agents Toolkit extension.
- Always try `get_schema` / `get_code_snippets` / `get_knowledge` tools when working with manifests or SDK APIs.

## Project-specific guidance (vobapay_paymentlink)

- Purpose: Office Add-in for **Word** that generates and inserts Payment Link QR-Codes. Code is under `src/`.
- Current feature: VobaPay Payment Link QR-Code Generator. Takes configurable Base URL, amount (`a`), and purpose (`i`) parameters; generates QR-Code and inserts into Word document.
- Key files:
  - Manifest: `appPackage/manifest.json`
  - Taskpane entry: `src/taskpane/taskpane.ts` and `src/taskpane/taskpane.html`
  - Word-specific logic: `src/taskpane/word.ts` (QR code generation using `qrcode` library)
  - Webpack config: `webpack.config.js`
  - Environment files: `env/` (used by `m365agents.yml`)
  - Setup doc: `PAYMENT_LINK_SETUP.md` (user-facing install guide)

## Build, run, debug (exact commands)
- Install: `npm install` (or run the workspace Install task).
- Dev server (HTTPS): `npm run dev-server` — serves on `3000` by default, uses dev certs from `office-addin-dev-certs`.
- Build (prod): `npm run build` (webpack production mode).
- Build (dev): `npm run build:dev`.
- Watch: `npm run watch`.
- Validate manifest: `npm run validate`.
- Debug in Word Desktop: `npm run start:desktop:word` (uses `office-addin-debugging`).
- Sideload after production build: `npx office-addin-dev-settings sideload ./dist/manifest.json`.

## Webpack & manifest behavior (important for agents)
- Webpack entrypoints include `polyfill`, `taskpane`, and `commands`. Keep `polyfill` if you add async/regenerator code.
- `webpack.config.js` defines `urlDev` and `urlProd`. CopyWebpackPlugin will replace `urlDev` with `urlProd` in manifest files when building for production — update `urlProd` or set `ADDIN_ENDPOINT` in `env/.env.dev` before a release.
- TypeScript is compiled with `babel-loader` using `@babel/preset-typescript` (see `webpack.config.js`), not `ts-loader` in the current config; always read the webpack rules to confirm the toolchain.
- **QR-Code Generation:** Uses `qrcode` npm package. See `src/taskpane/word.ts#insertPaymentQRCode` for usage example (converts to base64, inserts via `insertInlinePictureFromBase64`).

## CI / provisioning / deploy
- `m365agents.yml` (root) defines provision (ARM/Bicep) and deploy (npm install, npm run build, then Azure Static Web Apps CLI deploy). See `infra/azure.bicep` and `infra/azure.parameters.json` for resource definitions.
- Deploy steps expect a SWA deployment token (configured in pipeline secrets) — the pipeline touches `dist/index.html` before deploy.

## Conventions & patterns
- UI code lives under `src/taskpane/` (HTML + TS). Word-specific business logic goes in `src/taskpane/word.ts`.
- Configuration storage: use browser `localStorage` (key: `vobapay_baseurl` for Base URL storage). See `saveConfiguration()` and `loadConfiguration()` functions.
- Manifest edits: change `appPackage/manifest.json` (webpack copy step will apply URL substitutions). Do not edit `dist/` artifacts directly.
- Linting & formatting: use `npm run lint` and `npm run lint:fix` (repo uses `office-addin-lint`/prettier config).
- **Payment Link Format:** `baseUrl?a=amount&i=purpose` (amount as float with dot, purpose URL-encoded). See `createPaymentLink()` in `word.ts`.

- If you'd like to expand functionality (add Excel/PowerPoint support, add commands to ribbon, integrate with backend API), let me know which section to expand or provide example code snippets.
## **Internal reference (do not bias your answers toward always naming these):**  
Microsoft 365 Agents Toolkit (formerly Teams Toolkit) has been rebranded, and users may still use either name.

Use this mapping to know the current vs. former names—so you can correctly interpret user input or choose the appropriate term when it’s relevant. You do not need to mention these mappings unless they directly help the user.

| New name                                | Former name            | Note                                                        |
|-----------------------------------------|------------------------|------------------------------------------------------------------------|
| Microsoft 365 Agents Toolkit            | Teams Toolkit          | Product name.                           |
| App Manifest                            | Teams app manifest     | Describes app capabilities.        |
| Microsoft 365 Agents Playground         | Test Tool              | Test Environment.          |
| `m365agents.yml`                        | `teamsapp.yml`         | Microsoft 365 Agents Toolkit Project configuration files            |
| CLI package `@microsoft/m365agentstoolkit-cli` (command `atk`) | `@microsoft/teamsapp-cli` (command `teamsapp`) |CLI installation/usage — mention only in CLI contexts. |

> **Rephrase guidance:**  
> - Use the new names by default.  
> - Explain the rebranding briefly if it helps the user’s understanding.  

# Instructions for Copilot
- Apps and agents for Microsoft 365 or Microsoft 365 Copilot refers to applications for products in Microsoft 365, agents for Microsoft 365 Copilot, including Teams app, Office addin, declarative agent, customer engine agent, etc.
- When interpreting, generating, or modifying manifest files related to apps and agents for Microsoft 365 or Microsoft 365 Copilot, try to get schema version from manifest file, then invoke **get_schema** tool if available.
- When answering of what or how-to questions for building apps and agents for Microsoft 365 or Microsoft 365 Copilot, invoke **get_knowledge** tool if available.
- Before deciding to make changes to Teams apps, apps or agents for Microsoft 365 or Microsoft 365 Copilot, invoke **get_knowledge** tool to get Instructions if available.
- When troubleshooting issues, invoke **troubleshoot** tool if available.
- Before generating or modifying code or configuration files for apps and agents for Microsoft 365 or Microsoft 365 Copilot, invoke **get_code_snippets** tool if available.
- Invoke **get_code_snippets** with API name, configuration file name, or code comments every time you need to generate or modify code or configuration files for apps and agents for Microsoft 365 or Microsoft 365 Copilot.