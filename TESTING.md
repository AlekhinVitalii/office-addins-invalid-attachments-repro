# Exact test procedure (fully local, no Cloudflare)

Goal: reproduce the send **failure** with the original body-round-trip path, then
confirm the **fix** (compose-time prepend/append-on-send) sends clean on the same
broken-inline message.

**Test on classic Outlook on Windows.** That is where issue #6780 reproduces. New
Outlook / OWA is web-based and may not fail at all, so a "pass" there proves nothing.
Use new Outlook only for quick iteration + real dev-tools logs.

Files used:
- `test-broken-inline.eml` — the trigger (a signature with a corrupt inline image + a
  dangling `cid:`).
- `repro-manifest.local.xml` — **BUG** path (ItemSend + `body.setAsync` round-trip), localhost.
- `repro-fix-manifest.xml` — **FIX** path (prepend/append-on-send), localhost.

---

## Step 0 — Serve the add-in files over trusted HTTPS

From the repo root:

```powershell
npx office-addin-dev-certs install
npx http-server . -S `
  -C "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" `
  -K "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" -p 3000
```

Sanity check in a browser (no cert warning):
- `https://localhost:3000/repro-attachment-fail-open.js`
- `https://localhost:3000/repro-fix.js`

Leave this server running for the whole test.

## Step 1 — Turn on handler logging (classic Outlook)

Event/handler `console.log` isn't visible in a window on classic Outlook. Log to a file:

```powershell
New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" `
  -Name "RuntimeLogging" -Value "C:\temp\addin.log" -PropertyType String -Force | Out-Null
New-Item -Path "C:\temp" -ItemType Directory -Force | Out-Null
```

Tail it while testing: `Get-Content C:\temp\addin.log -Wait`

---

## Step 2 — Reproduce the BUG (baseline)

1. **Sideload the bug manifest.** Add a string value under
   `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\` whose data is the full path to
   `repro-manifest.local.xml`:

   ```powershell
   New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" `
     -Name "ReproBug" -Value "C:\path\to\repro-addin\repro-manifest.local.xml" `
     -PropertyType String -Force | Out-Null
   ```
2. **Restart classic Outlook.** If the add-in doesn't load, close Outlook, delete
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, restart.
3. **Open the trigger.** Double-click `test-broken-inline.eml` → it opens as a read
   message. The company-logo image should render **broken / whited-out** (that's the
   corrupt inline part) and the tiny icon is a dangling reference.
4. **Get it into compose.** Click **Forward**. Put **your own address** in *To*. The
   broken inline images carry into the compose body.
5. **Send.**
6. **Expected (bug):** send fails with *"There was a problem sending this message because
   one or more attachments failed to upload. Please try again"* — often after a hang, and
   the item is stuck in Drafts. `C:\temp\addin.log` shows the `[repro]` lines and the
   `body.setAsync` call.

If the send actually succeeds here, the synthetic `.eml` isn't triggering it on your
build — see **Trigger didn't fire?** below before moving on.

---

## Step 3 — Confirm the FIX

1. **Swap manifests.** Remove the bug value and add the fix one:

   ```powershell
   Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name "ReproBug"
   New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" `
     -Name "ReproFix" -Value "C:\path\to\repro-addin\repro-fix-manifest.xml" `
     -PropertyType String -Force | Out-Null
   ```
2. **Restart Outlook** (clear the `Wef` cache again if needed).
3. **Repeat Steps 2.3–2.5:** open `test-broken-inline.eml` → **Forward** → your address →
   **Send**.
4. **Expected (fix):**
   - Send **completes** — no upload error.
   - Open the message in **Sent Items**: a blue **"CLASSIFICATION: INTERNAL"** banner at
     the **top** and the italic footer at the **bottom**, with the (still-broken) signature
     image untouched in between.
   - `C:\temp\addin.log` shows `[fix] prependOnSendAsync returned {status: succeeded}`,
     `[fix] appendOnSendAsync returned {status: succeeded}`, and
     `[fix] Completing send, allowEvent=true` — and **no** `body.setAsync` anywhere.

---

## Step 4 — Draft / idempotency check

1. With the **fix** manifest active, Forward `test-broken-inline.eml` to yourself, then
   **Save** as a draft and close it.
2. Reopen the draft from **Drafts** and **Send**.
3. **Check the Sent copy:** exactly **one** header and **one** footer. If you see them
   **twice**, re-arming on draft re-open is stacking — note it; the handler would then need
   an idempotency guard (e.g. a marker) or to switch `OnMessageCompose` → `OnNewMessageCompose`.

---

## Interpreting results

| Bug manifest | Fix manifest | Conclusion |
| --- | --- | --- |
| Send fails | Send succeeds | ✅ Hypothesis confirmed: the `getAsync`/`setAsync` round-trip is the trigger; on-send stamping fixes it. Ship it. |
| Send fails | Send **also fails** | ❌ The round-trip isn't the (only) trigger. Capture the fix log + error and rethink — likely a genuine platform bug to press on #6780. |
| Send succeeds | — | Trigger didn't fire on this build/client — see below. |

## Trigger didn't fire?

The `.eml` is a best-effort synthetic of a corrupted signature image; the exact server-side
state in #6780 is hard to reproduce offline. Try, in order:
1. Run it on **classic** Outlook, not new Outlook.
2. Use a **real** offending message if you have one (a forwarded external thread whose
   signature images show as broken/whited-out) instead of the synthetic `.eml`.
3. Reply-all on a long external thread rather than Forward.
4. Confirm the broken image is a **`cid:` inline** image (an attachment), not a plain
   `http(s)` external image — only inline images upload on send, so only they hit this path.

## New Outlook / OWA (quick iteration only)

Sideload via Settings → **Manage add-ins** → **My add-ins** → **Add a custom add-in →
Add from file** → pick the manifest. Logs appear in the browser dev tools (F12). Handy for
shaking out the handler, but do the authoritative pass on classic.
