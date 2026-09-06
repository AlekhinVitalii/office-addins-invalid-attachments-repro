# Fix variant — on-send content stamping (no body round-trip)

This branch adds a **fix variant** next to the existing repro so you can run an
A/B on the same broken-signature message:

| Variant | Manifest | Handler | Body write | Expected on broken inline image |
| --- | --- | --- | --- | --- |
| **Repro (bug)** | `repro-manifest.xml` | `repro-attachment-fail-open.js` | `body.getAsync(Html)` + `body.setAsync(Html)` (full round-trip) | Send fails: *"one or more attachments failed to upload"* |
| **Fix** | `repro-fix-manifest.xml` | `repro-fix.js` | none — `prependOnSendAsync` (top) + `appendOnSendAsync` (bottom), armed at compose | Send succeeds; header + footer stamped |

## Why the fix avoids the bug

The failure is triggered by the **`getAsync` + `setAsync` body round-trip**: rewriting
the whole body makes Outlook re-parse and re-upload the existing broken `cid:` inline
images at send. The fix never reads or rewrites the body. It queues the classification
header and footer at **compose time** via `prependOnSendAsync` / `appendOnSendAsync`;
Outlook stamps them itself at send, without re-touching the inline attachments.

`appendOnSendAsync` / `prependOnSendAsync` **cannot** be called from the send handler
(`OnMessageSend`/`ItemSend`) — that returns `ApiCallNotSupportedByExtensionPoint`. They
must be armed from a **compose-time** event. `OnMessageCompose` fires automatically (no
button, no task pane), which keeps the flow frictionless and non-opt-out.

Requirements: **Mailbox 1.13** (`prependOnSendAsync`) and the **`AppendOnSend`**
extended permission (both in `repro-fix-manifest.xml`).

## Fully-local test (no Cloudflare Pages)

Everything is served from `https://localhost:3000`. You need Node/npm.

### 1. Trusted HTTPS on localhost

```powershell
npx office-addin-dev-certs install
```

This installs a trusted CA and emits `localhost.crt` / `localhost.key` (default:
`%USERPROFILE%\.office-addin-dev-certs\`). Outlook will not load add-in content over
plain HTTP, and the cert must be trusted — this handles both.

### 2. Serve the repo folder over HTTPS

From the repo root:

```powershell
npx http-server . -S -C "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" -K "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" -p 3000
```

Confirm `https://localhost:3000/repro-fix.js` loads in a browser without a cert warning.

> The manifest's `IconUrl`/`HighResolutionIconUrl` point at `localhost` too; there are no
> icon files in the repo, so either drop `icon-32.png`/`icon-64.png` into the folder or
> ignore the missing-icon 404 (it doesn't affect the handlers).

### 3. Sideload

**New Outlook / OWA** (real dev tools for logs):
Settings → General → **Manage add-ins** (or **Get Add-ins**) → **My add-ins** →
**Add a custom add-in → Add from file** → pick `repro-fix-manifest.xml`.

**Classic Outlook on Windows** (this is where the bug actually reproduces):
1. Registry: under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\` add a string
   value (any name) whose data is the **full path** to `repro-fix-manifest.xml`.
2. Enable handler logging: add a string value `RuntimeLogging` under the same key, set
   to a writable `.log` path (e.g. `C:\temp\addin.log`). The `[fix]` lines land there.
3. Restart Outlook. If it doesn't pick up, clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
   and restart again.

### 4. Run the A/B

1. Compose a message and insert the **broken signature image** (the case from
   office-js issue #6780).
2. **Bug:** with `repro-manifest.xml` sideloaded, send → expect the upload failure.
3. **Fix:** with `repro-fix-manifest.xml` sideloaded, send → expect success, with the
   INTERNAL header at the top and footer at the bottom in **Sent Items**.

## What this PoC is validating

1. **Does compose-time `prepend/appendOnSendAsync` avoid the broken-inline upload
   failure?** The entire recommendation rests on "no body round-trip → Outlook doesn't
   re-touch the broken `cid:` images." Confirm on **classic Outlook**, since new Outlook
   (web-based) may not reproduce the original failure at all.
2. **Draft / idempotency:** re-open a saved draft and send. Check the Sent copy for
   double-stamped header/footer, since `OnMessageCompose` fires again on draft edit.
