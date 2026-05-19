# Apolla Capture — pilot install guide

This is the no-developer-tools install path for the Apolla Capture Outlook Add-in. Works on Outlook for Mac and Outlook for Windows. ≤5 minutes from start to first capture.

A friendlier HTML version of this guide with copy-buttons is at <https://rosscantrell.github.io/apolla-outlook-addin/>.

---

## 1. Sideload the Add-in (one-time, ~3 minutes)

You'll paste a URL into Outlook's Custom Add-ins UI. The URL points at a GitHub Pages-hosted manifest; the Add-in itself runs entirely in Outlook's process.

**The manifest URL:**

```
https://rosscantrell.github.io/apolla-outlook-addin/manifest.xml
```

**Steps:**

1. Open Outlook. Open any email (the install path lives inside an email-message ribbon, not in the inbox toolbar).
2. In the email's toolbar, click **Get Add-ins**.
   - On Outlook for Mac, you may need to click the "More" (⋯) overflow first.
   - On Outlook for Windows, it's typically in the Home ribbon.
3. In the Add-ins dialog: **My Add-ins** → scroll to **Custom Add-ins** → **Add a custom add-in** → **Add from URL…**
4. Paste the manifest URL → **OK** → **Install** → accept the "untrusted source" prompt.
   - "Untrusted source" is expected because the Add-in is sideloaded, not yet published to Microsoft AppSource. Tracked as future activation FN-ADDIN-05.
5. An **Apolla** group appears in the email ribbon with two buttons: **Capture** and **Configure**.

---

## 2. Configure your token (one-time, ~1 minute)

The Add-in needs to know which Apolla server to POST to, and a bearer token to authenticate.

1. Open any email. Click **Configure** in the Apolla ribbon group. A task pane opens.
2. Paste your **Apolla bearer token** — same value as `INGESTION_API_KEY` on your Apolla server.
   - Don't have one yet? On the machine running schema-browser: `openssl rand -hex 32`. Use that value for both the server's `INGESTION_API_KEY` env var AND this field.
3. **Base URL** defaults to `http://localhost:3001` (for local-dev pilots). If you're pointing at a different Apolla server, change it.
4. Click **Test connection**. Expect "Apolla server reachable…". If it fails with a CORS error, ask whoever runs the Apolla server to set `ALLOWED_WEB_ORIGINS=https://rosscantrell.github.io` (per D-ADDIN-15).
5. Click **Save**. The token persists via Office's `roamingSettings` (encrypted; syncs across your Outlook installs — Mac, Windows, mobile — for the same Microsoft account).

---

## 3. Capture an email (per-email, ~1 click)

1. Open the email you want to capture.
2. Click **Capture** in the Apolla ribbon.
3. A notification banner appears above the message:
   - `Captured: <subject> | terms: N` — first time capture; document now in your Apolla corpus.
   - `Already captured: <subject> | terms: N` — same email content was captured before; idempotent (no duplicate ingest).

That's it. The Add-in does the rest: reads subject + sender + recipients (excluding Bcc — see Privacy below) + plain-text body + attachment metadata, hashes the body for a deterministic `documentId`, and POSTs to your Apolla server's `/api/ingest-document` endpoint with bearer auth.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| "Apolla not configured" notification | Token not set | Click Configure; complete setup |
| "Apolla auth failed — re-run Configure" | Token doesn't match server's `INGESTION_API_KEY` | Re-paste the correct value in Configure |
| "Apolla server unreachable at <url>" | Server down, URL wrong, or network blocks it | Verify schema-browser is running; check the URL in Configure |
| "Apolla returned HTTP 4xx/5xx" | Server-side validation error or CORS misconfiguration | Open Outlook's "Inspect Element" → Console for the full response |
| Configure pane won't open | Manifest cache | Remove + re-add the Add-in; verify <https://rosscantrell.github.io/apolla-outlook-addin/configure.html> is reachable directly |
| Capture button missing from ribbon | Add-in didn't install, or install was for a different email type | Repeat Step 1; ensure you opened a regular email (not a meeting invite — D-ADDIN-14 limits v1 to messages only) |

---

## Privacy posture

- The Add-in only POSTs email content to the Apolla server URL **you** configure. No telemetry, no analytics, no third-party calls (D-ADDIN-12).
- **Bcc recipients are deliberately not captured** (D-ADDIN-09) — bcc is a privacy boundary by design.
- **Attachment file bodies are not captured** — only metadata (name, content-type, size). Attachment ingest is FN-ADDIN-02 if pilot evidence later shows it's needed.
- **Email body is captured as plain text, not HTML** (D-ADDIN-07) — to avoid markup polluting the corpus's LLM extraction.

---

## Removing the Add-in

In Outlook: **Get Add-ins** → **My Add-ins** → find "Apolla Capture" under Custom Add-ins → click the **…** → **Remove**. Token in `roamingSettings` is cleared automatically by Office.

---

*Apolla Capture · WP-OCR-05 v1 · part of the [Viktora Apolla](https://viktora.ai) capture surface · sibling to the [browser extension](#) (web Slack/Teams/Gmail) and [ocr-capture utility](#) (desktop OCR).*
