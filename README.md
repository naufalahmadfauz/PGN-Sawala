# PGN WhatsApp QA Harness

This harness drives the consumer WhatsApp Web UI with Playwright. Messages follow the real route through the PGN WhatsApp Business number, LivePerson, and the PGN bot; no direct bot or messaging API is used.

## Commands

- `npm run whatsapp:login` launches headed Playwright Chromium under Xvfb, writes only a validated QR to `artifacts/whatsapp-login.png`, and waits for login.
- `npm run whatsapp:verify` verifies that `.whatsapp-profile/` opens without another QR scan.
- `npm run test:single` sends `Halo`, captures all new incoming messages after an approximately 1.8-second idle window, and writes `reports/PGN_Single_Test_Result.xlsx`.
- `npm run test:pgn` runs `data/PGN_Test_Cases.xlsx` sequentially and writes `reports/PGN_Test_Results.xlsx`.
- `npm run data:template` creates a starter test-case workbook without overwriting an existing one.
- `npm run check` runs the TypeScript compiler.

## Configuration

Local settings belong in `.env`; the supported keys and defaults are listed in `.env.example`. Configure either `PGN_WHATSAPP_PHONE` (international digits, without `+`) or `PGN_WHATSAPP_CHAT`. A phone number is preferred when both are present because the direct WhatsApp chat URL avoids ambiguous chat-name matches.

The harness verifies the open conversation header against that configured target before sending. It also confirms that WhatsApp rendered each new outgoing message before waiting for a reply.

The persistent profile is fixed at `.whatsapp-profile/`; data files are restricted to `data/`, and generated reports are restricted to `reports/`. The profile, `.env`, QR image, screenshots, diagnostics, and generated XLSX reports are gitignored. Treat the profile as an authenticated secret and do not share or commit it.

## Input And Results

The input workbook must use a `Test Cases` worksheet and these required columns:

| Test ID | Category | User Input | Expected Behaviour |
| --- | --- | --- | --- |
| POS-001 | Greeting | Halo | Bot should greet the user |

`Scenario ID` is accepted as optional metadata for future grouped multi-turn scenarios. Tests currently run in workbook order in one real WhatsApp conversation. The authenticated history shows that `reset` has received `Session deleted` from this bot, but the harness does not assume that this guarantees a complete backend reset and does not invoke it automatically. The reset contract should be confirmed with PGN before independent cases are isolated automatically.

The report contains `Results` and `Transcript` worksheets. `CAPTURED` means a response was captured, not that its expected behaviour was automatically judged. Failures save screenshots and diagnostics under `artifacts/debug/`.
