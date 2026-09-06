/**
 * FIX variant of repro-attachment-fail-open.js.
 *
 * The original repro reproduces the "one or more attachments failed to upload"
 * send failure by doing a body.getAsync(Html) + body.setAsync(Html) round-trip
 * inside a legacy ItemSend handler on a message that contains broken inline
 * (cid) signature images. The round-trip makes Outlook re-parse and re-upload
 * those broken inline attachments on send, which fails the send.
 *
 * This variant NEVER round-trips the body. Instead:
 *
 *   - onMessageComposeHandler (OnMessageCompose, event-based, fires
 *     automatically with no button) queues a classification HEADER at the top
 *     via prependOnSendAsync and a FOOTER at the bottom via appendOnSendAsync.
 *     Outlook itself stamps both into the body at send time. Because we never
 *     read+rewrite the existing body, the broken inline images are never
 *     re-touched.
 *
 *   - onMessageSendHandler (OnMessageSend, Smart Alerts) still performs the
 *     production-equivalent attachment scan and sets the X-GV internet header,
 *     but does NO body.getAsync / body.setAsync at all, then allows the send.
 *
 * Watch the [fix] console lines (real dev tools in new Outlook / OWA; the
 * RuntimeLogging log file in classic Outlook on Windows). The header and
 * footer appear in the message in Sent Items, and the send completes even when
 * a broken signature image is present.
 */

Office.onReady(() => {
  console.log('[fix] Office.js ready.');
});

// --- helpers -------------------------------------------------------------

function getBodyType(item) {
  return new Promise((resolve) => {
    item.body.getTypeAsync((res) => {
      const type = res.status === Office.AsyncResultStatus.Succeeded ? res.value : Office.CoercionType.Html;
      console.log('[fix] body.getTypeAsync returned.', { status: res.status, type, error: res.error });
      resolve(type);
    });
  });
}

function prependOnSend(item, data, coercionType) {
  return new Promise((resolve, reject) => {
    console.log(`[fix] Calling prependOnSendAsync (coercionType=${coercionType})...`);
    item.body.prependOnSendAsync(data, { coercionType }, (res) => {
      console.log('[fix] prependOnSendAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

function appendOnSend(item, data, coercionType) {
  return new Promise((resolve, reject) => {
    console.log(`[fix] Calling appendOnSendAsync (coercionType=${coercionType})...`);
    item.body.appendOnSendAsync(data, { coercionType }, (res) => {
      console.log('[fix] appendOnSendAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

function setReproHeader(item) {
  return new Promise((resolve, reject) => {
    const addedAt = new Date().toISOString();
    console.log(`[fix] Calling internetHeaders.setAsync(X-GV-Repro=${addedAt})...`);
    item.internetHeaders.setAsync({ 'X-GV-Repro': addedAt }, (res) => {
      console.log('[fix] internetHeaders.setAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

// Verbatim-style port of the production attachment scan (fail-open, no forced
// failures). Kept so the send-time work mirrors production; note it does NOT
// touch the body.
function getAttachments(item) {
  return new Promise((resolve) => {
    console.log('[fix] Calling getAttachmentsAsync...');
    item.getAttachmentsAsync((result) => {
      console.log('[fix] getAttachmentsAsync returned.', {
        status: result.status,
        count: result.value?.length,
        error: result.error,
      });

      if (result.status === Office.AsyncResultStatus.Failed || !result.value || result.value.length === 0) {
        resolve([]);
        return;
      }

      const contentPromises = [];
      for (const attachment of result.value) {
        if (attachment.isInline) {
          console.log(`[fix] Skip inline attachment: ${attachment.name}`);
          continue;
        }
        if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File) {
          contentPromises.push(new Promise((contentResolve) => {
            item.getAttachmentContentAsync(attachment.id, (ac) => {
              // Fail-open on a broken/slow download so the event never hangs.
              if (ac.status === Office.AsyncResultStatus.Failed || !ac.value) {
                console.error(`[fix] getAttachmentContent failed for "${attachment.name}".`, ac.error);
                contentResolve(null);
                return;
              }
              contentResolve({ name: attachment.name });
            });
          }));
        }
      }

      Promise.all(contentPromises).then((v) => resolve(v.filter(Boolean)));
    });
  });
}

// --- handlers ------------------------------------------------------------

// Runs automatically when a message is composed (new/reply/reply-all/forward
// or draft edit). Arms the top header + bottom footer for send. No button, no
// task pane, no body round-trip.
async function onMessageComposeHandler(event) {
  console.log(`[fix] onMessageComposeHandler started at ${new Date().toISOString()}`);
  const item = Office.context.mailbox.item;

  const bodyType = await getBodyType(item);
  const isHtml = bodyType === Office.CoercionType.Html;

  // HTML can't be inserted into a plain-text body, so match the format.
  const header = isHtml
    ? '<div style="border:2px solid #2b6cb0;background:#ebf4ff;padding:6px 10px;font-family:sans-serif;"><strong>CLASSIFICATION: INTERNAL</strong> &mdash; prepended on send</div><br>'
    : 'CLASSIFICATION: INTERNAL — prepended on send\n\n';

  const footer = isHtml
    ? '<br><div style="border-top:1px solid #a0aec0;padding-top:6px;font-family:sans-serif;color:#4a5568;"><em>This message is classified INTERNAL. Appended on send.</em></div>'
    : '\n\n---\nThis message is classified INTERNAL. Appended on send.';

  try {
    await prependOnSend(item, header, bodyType);
    await appendOnSend(item, footer, bodyType);
    console.log('[fix] Classification header + footer armed for send.');
  } catch (err) {
    console.error('[fix] Failed to arm on-send content.', err);
  }

  // IMPORTANT: re-arming on a re-opened draft may stack or replace the queued
  // content. Watch the Sent Items copy for double-stamping and note it here.
  event.completed();
}

// Smart Alerts send handler. Mirrors the production attachment scan + X-GV
// header, but performs NO body operations. Always allows the send.
async function onMessageSendHandler(event) {
  console.log(`[fix] onMessageSendHandler started at ${new Date().toISOString()}`);
  const item = Office.context.mailbox.item;

  try {
    const attachments = await getAttachments(item);
    console.log('[fix] Collected attachments:', attachments);
  } catch (err) {
    console.error('[fix] Attachment collection failed — failing open.', err);
  }

  try {
    await setReproHeader(item);
    console.log('[fix] Set X-GV-Repro internet header.');
  } catch (err) {
    console.error('[fix] Failed to set internet header.', err);
  }

  console.log('[fix] Completing send, allowEvent=true.');
  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler);
Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
