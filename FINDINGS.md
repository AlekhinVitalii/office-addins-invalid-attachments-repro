# Findings — prepend/append-on-send fix for the broken-inline send failure

## The bug (office-js #6780)

An Outlook add-in that, on send, does a `body.getAsync(Html)` + `body.setAsync(Html)`
round-trip on a message containing **broken `cid:` inline images** (e.g. corrupted
signature images) makes Outlook re-parse and re-upload those inline attachments during
send commit. The broken ones fail to upload → *"one or more attachments failed to
upload"*, and the send is blocked. This is the production `office-classifier` behavior,
ported verbatim in `repro-attachment-fail-open.js`.

## Root cause

The `getAsync`/`setAsync` **whole-body round-trip** is the trigger. Rewriting the body
re-serializes the existing `cid:` inline image references, forcing Outlook to re-upload
them on send.

## The fix

Never round-trip the body. Add the classification header (top) and footer (bottom) with
**`prependOnSendAsync` + `appendOnSendAsync`**, which queue content that Outlook stamps at
send without reading/rewriting the existing body — so the broken inline images are never
re-touched.

Key API facts established during this investigation:

- `append/prependOnSendAsync` **cannot** be called from the send handler
  (`OnMessageSend`/`ItemSend`) → `ApiCallNotSupportedByExtensionPoint`. They must be armed
  from a **compose-time** trigger.
- Two valid compose-time triggers: an **event** (`OnMessageCompose`, automatic/frictionless)
  or a **ribbon button** (user-activated).
- They require the **`AppendOnSend`** permission and **Mailbox 1.13** (`prependOnSend`).
- The legacy **`ItemSend`** on-send event is now **rejected by manifest validation**
  ("ItemSend event in VersionOverrides is not allowed") and also needs a tenant on-send
  policy — so the modern path (Smart Alerts + compose events) is the supported one.

## What was proven, and where

Tested on a personal **outlook.com** account via **Outlook on the web** (the corporate
Forcepoint tenant blocks user sideloading of custom add-ins — the "Add a custom add-in"
option is removed and the classic `WEF\Developer` registry sideload silently does not
register).

Using the **button variant** (`repro-fix-button-manifest.xml`), a single compose session
produced this console trace and a correctly stamped message:

```
[fix-btn] body.getTypeAsync returned. {status: 'succeeded', type: 'html'}
[fix-btn] prependOnSendAsync returned. {status: 'succeeded'}
[fix-btn] appendOnSendAsync returned. {status: 'succeeded'}
[fix-btn] Header + footer armed for send.
```

The sent message showed the INTERNAL header at the top and the footer at the bottom with
the original body untouched between them.

✅ **Proven:** `prepend/appendOnSendAsync` add top+bottom content with no body round-trip,
on a live mailbox; APIs and permission are available.

⚠️ **Idempotency (confirmed live):** clicking the button twice produced **two** headers and
**two** footers — repeated arming **stacks**, it does not replace. Production must arm
**exactly once per send**:
  - Button model → de-bounce / disable after first arm.
  - `OnMessageCompose` → fires again on draft reopen; guard with a marker, or use
    `OnNewMessageCompose` (doesn't fire on draft edit).

✅ **Event-based auto-fire + no consent (confirmed live):** the event-based variant
(`repro-fix-manifest.xml`, no button) fired `onMessageComposeHandler` **automatically** on
compose open with **no user action and no consent window**; `prepend/appendOnSendAsync`
both succeeded and the sent message had a **single** header + footer (no doubling — event
fires once per compose). `onMessageSendHandler` also set the `X-GV-Repro` internet header
on send. This holds on a **self-installed personal outlook.com** account via Outlook on the
web. (An earlier "didn't fire" reading was a false negative from an http-server request log
that wasn't being captured — the F12 console is the reliable signal.)

Consent summary: no runtime prompt at compose or send, ever. The only place a user sees
anything is the one-time **install** warning when self-sideloading; under **admin
deployment** the admin consents once and users get it silently (no per-user consent).

❌ **Not reproduced here:** the original send *failure* — it is a classic-desktop + Exchange
behavior; personal OWA won't trigger it. The bug-vs-fix A/B still needs classic Outlook +
Exchange Online (Vitalii's env or an M365 dev tenant).

**Rendering gotcha (not a bug):** when the fix is applied to a **forwarded** message,
Outlook on the web's conversation/threaded view collapses the forwarded original as
"message history" and shows the prepended header separately — it *looks* like the header
was sent as a second mail. It isn't: Sent Items has a single message, and delivered to an
external client (e.g. Gmail) it renders correctly as one message with header at top,
forwarded content in the middle, and footer at the bottom. Turn off conversation view to
see it as one message in OWA.

## Recommended next validation (Vitalii's env or an M365 dev tenant)

1. Reproduce the failure with `repro-bug-smartalerts-manifest.xml` (Smart Alerts
   `OnMessageSend` doing the `setAsync` round-trip) on **classic Outlook + Exchange Online**,
   using `test-broken-inline.eml` (Forward → Send).
2. Swap to `repro-fix-manifest.xml` (`OnMessageCompose` + `prepend/appendOnSend`) and confirm
   the same message sends clean with header+footer stamped.
3. Verify the event-based idempotency guard (reopen a draft, confirm single stamp).

## File map

| File | Purpose |
| --- | --- |
| `repro-manifest.xml` / `repro-attachment-fail-open.js` | Original repro (legacy ItemSend; manifest now invalid) |
| `repro-manifest.local.xml` | Localhost copy of the original (also invalid — kept for reference) |
| `repro-bug-smartalerts-manifest.xml` / `repro-bug-smartalerts.js` | **Valid** bug baseline via Smart Alerts + `setAsync` round-trip |
| `repro-fix-manifest.xml` / `repro-fix.js` | **Fix** via `OnMessageCompose` + `prepend/appendOnSend` (frictionless) |
| `repro-fix-button-manifest.xml` / `repro-fix-button.js` | **Fix (button)** for envs without event activation (what was validated live) |
| `test-broken-inline.eml` | Broken-inline trigger for the A/B |
| `icon-*.png` | Button/tile icons |
| `TESTING.md` | Step-by-step A/B procedure |
| `FINDINGS.md` | This file |
