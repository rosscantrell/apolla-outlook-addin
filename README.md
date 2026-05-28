# apolla-outlook-addin

Outlook Add-in for capturing emails into the [Viktora Apolla](https://viktora.ai) corpus.

**Productized in WP-OCR-05** (v1.2-FINAL brief in the [AI-Light-Prototype](https://github.com/rosscantrell/AI-Light-Prototype) repo). One of three complementary Apolla capture surfaces:

| Surface | Where it runs | What it captures |
|---|---|---|
| **This Add-in** | Outlook desktop (Mac + Windows) | Open email's subject + body + recipients + attachment metadata via Office.js |
| Browser extension (WP-B-3.1) | Chromium tabs | DOM-extracted content from web Slack/Teams/Gmail |
| ocr-capture utility (WP-OCR-01/03) | Mac desktop | Vision-OCR of any app window (full-window or region-select) |

All three POST to the same `/api/ingest-document` endpoint on the Apolla server with `sourceMetadata.captureMethod` distinguishing the source.

---

## Install

Pilot users: see [PILOT-INSTALL.md](PILOT-INSTALL.md) or the friendlier HTML version at <https://rosscantrell.github.io/apolla-outlook-addin/>.

---

## Architecture

```
   Outlook ribbon                  Office.js API                 fetch() POST
[Capture button] ─→ [item.subject/body/from/to/cc/attachments] ─→ /api/ingest-document
                                                                        │
                                                          (bearer auth — D-OCR-04)
                                                                        │
                                                                        ▼
                                                              Apolla schema-browser
                                                                        │
                                                                  IngestionResult
                                                                        │
                          ┌─────────────────────────────────────────────┴──────────────┐
                          │                                                            │
                  alreadyExisted: true                                              new ingest
              ↺ Already captured: <subject>                              ✓ Captured: <subject> | terms: N
```

### Locked decisions

- **D-ADDIN-02** `sourceMetadata.captureMethod = 'outlook-addin'` literal (per WP-OCR-01 D-OCR-08 universal taxonomy)
- **D-ADDIN-03** `sourceMetadata.sourceApp = 'com.microsoft.outlook'` (Mac-style bundle ID; cross-platform identical)
- **D-ADDIN-04** Bearer token + base URL stored in `Office.context.roamingSettings` (encrypted, cloud-synced per user)
- **D-ADDIN-05** `documentId = 'EMAIL-' + SHA-256(normalize(body text)).slice(0,16)` — deterministic; identical body content → identical id → idempotent re-POST
- **D-ADDIN-06** `title = item.subject` (structured field, NOT first-line-of-body heuristic)
- **D-ADDIN-07** Plain text only (`Office.CoercionType.Text`); HTML body NOT captured (FN-ADDIN-03 future)
- **D-ADDIN-08** Attachment metadata only; file bodies NOT downloaded (FN-ADDIN-02 future)
- **D-ADDIN-09** Recipients = `to` + `cc` only; `bcc` deliberately NOT captured (privacy)
- **D-ADDIN-10** Two ribbon buttons: Capture (ExecuteFunction) + Configure (ShowTaskpane)
- **D-ADDIN-11** Capture button auto-opens Configure pane if token unset (first-run UX)
- **D-ADDIN-12** No telemetry; only outbound HTTP call is the ingest POST
- **D-ADDIN-13** v1 distribution = sideload; AppSource is FN-ADDIN-05 (post-pilot)
- **D-ADDIN-14** v1 captures `itemType === 'message'` only; meetings/appointments are FN-ADDIN-10 territory (Microsoft Graph integration)
- **D-ADDIN-15** Server-side CORS via `ALLOWED_WEB_ORIGINS` env var (lands in AI-Light-Prototype PR #174)
- **D-ADDIN-16** Public standalone repo (this one)

---

## URL-param pre-fill contract (WP-Onboarding-One-Click v1.1)

`configure.html` accepts two query-string parameters when opened from a droplet-hosted manifest URL:

| Param   | Source                                          | Behavior                                                                                                              |
|---------|-------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------|
| `tenant` | The droplet hostname's leading label (e.g., `threshold-eval` from `threshold-eval.viktora.ai`) | Reconstructs `apollaBaseUrl` as `https://<tenant>.viktora.ai` and is rendered in the green "Connected ✓" hint banner. |
| `token`  | The per-user `addinToken` minted by `POST /api/auth/issue-addin-token` on the schema-browser  | Pre-fills `apollaToken` field.                                                                                        |

When BOTH params are present AND no roaming settings exist, the page **auto-saves** the values via `Office.context.roamingSettings.saveAsync()` immediately on `Office.onReady`, surfaces a "Connected ✓" status, and (best-effort) tries `Office.context.ui.messageParent` to close the pane programmatically. The user is then guided to close the pane manually if Office.js declines the close request — subsequent Capture clicks see saved settings and skip the configure flow entirely.

When EITHER param is missing OR roaming settings already exist (the re-configure scenario), the page renders the legacy manual form (URL + token + Test + Save) so power users can re-point at a different server.

**Cross-repo coordination:** the droplet's `GET /outlook-manifest.xml?token=<addinToken>` endpoint emits a manifest whose configurePane URL is `https://rosscantrell.github.io/apolla-outlook-addin/configure.html?tenant=<pilot>&token=<apolla_…>` — this URL is the load-bearing contract this file relies on. Both ends must agree.

### Lifecycle event reporting

Both `configure.html` and `functions.html` POST best-effort events to `<apolla-base>/api/auth/addin-lifecycle-event` so operators can debug "did Trisha's add-in ever phone home?" via a single `grep` on the server-side `addin-lifecycle-log.jsonl`. Events: `configure_loaded`, `configure_autosaved`, `capture_attempted`, `capture_succeeded`, `capture_failed`. Failures are silent (no UX impact); requires the bearer token to be set (it gets sent in the `Authorization` header).

---

## Server-side prerequisite

The Apolla schema-browser server must have CORS web-origin support (PR #174 in AI-Light-Prototype). Before pilot users can POST from `https://rosscantrell.github.io`:

```bash
# On the server:
export ALLOWED_WEB_ORIGINS=https://rosscantrell.github.io
export INGESTION_API_KEY=$(openssl rand -hex 32)   # share with pilot users via Configure
cd schema-browser && npm run dev
```

---

## Layout

```
apolla-outlook-addin/
├── manifest.xml         ← Office Add-in manifest; references GitHub Pages URLs
├── functions.html       ← Capture handler (Office.js + fetch POST + bearer auth)
├── configure.html       ← Configure task pane (token + base URL + Test connection + Save)
├── taskpane.html        ← ItemRead form placeholder (manifest requires; minimally used)
├── index.html           ← GitHub Pages landing page with copy-able manifest URL
├── icon-{16,32,80}.png  ← Branding icons referenced from manifest
├── PILOT-INSTALL.md     ← Pilot user install guide
├── README.md            ← This file
└── .gitignore           ← .DS_Store, etc.
```

---

## Validating the manifest

Microsoft's official validator catches subtle schema errors:

```bash
npx -y office-addin-manifest validate manifest.xml
```

Run before any push that modifies `manifest.xml`.

---

## Forward notes (when activation makes sense)

- **FN-ADDIN-01** Move hosting from GitHub Pages to Apolla-owned infrastructure
- **FN-ADDIN-02** Attachment ingest (file bodies, separate documents)
- **FN-ADDIN-03** HTML-body capture for fidelity preservation
- **FN-ADDIN-04** Compose-mode capture (drafts)
- **FN-ADDIN-05** AppSource submission (1-2 week Microsoft validation cycle)
- **FN-ADDIN-06** Meeting / calendar / appointment item capture
- **FN-ADDIN-07** Bulk-capture ("select N emails, capture all")
- **FN-ADDIN-08** Office Add-in variants for Word / OneNote / PowerPoint
- **FN-ADDIN-10** Microsoft Graph integration (sendMail, findMeetingTimes, create events) — bidirectional
- **FN-ADDIN-11** Loom-style video walkthrough of sideload + Configure flow

---

*Apolla Capture · WP-OCR-05 v1 · part of the Viktora Apolla capture surface*
