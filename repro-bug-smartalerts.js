/**
 * BUG baseline, modern activation.
 *
 * Reproduces the same failure as the original legacy-ItemSend repro, but via a
 * VALID Smart Alerts OnMessageSend handler (the legacy ItemSend event is no
 * longer allowed in a manifest). The failure mechanism is identical: read the
 * whole body with body.getAsync(Html) and rewrite it with body.setAsync(Html).
 * On a message with broken cid inline images, that round-trip makes Outlook
 * re-parse and re-upload the broken inline attachments on send.
 *
 * This is the A/B counterpart to repro-fix.js: SAME activation (OnMessageSend
 * Smart Alerts), the ONLY difference is body.setAsync round-trip here vs
 * prepend/appendOnSend armed at compose in the fix.
 */

Office.onReady(() => {
  console.log('[bug] Office.js ready.');
});

function getBodyHtml(item) {
  return new Promise((resolve, reject) => {
    console.log('[bug] Calling body.getAsync(Html)...');
    item.body.getAsync(Office.CoercionType.Html, (res) => {
      console.log('[bug] body.getAsync returned.', { status: res.status, len: res.value?.length, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error);
    });
  });
}

// The offending call: replace the WHOLE body (banner + original HTML) in one
// setAsync. This re-serializes the existing cid inline images.
function setBodyHtml(item, html) {
  return new Promise((resolve, reject) => {
    console.log(`[bug] Calling body.setAsync(Html), length=${html.length}...`);
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (res) => {
      console.log('[bug] body.setAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

function setReproHeader(item) {
  return new Promise((resolve, reject) => {
    const addedAt = new Date().toISOString();
    item.internetHeaders.setAsync({ 'X-GV-Repro': addedAt }, (res) => {
      console.log('[bug] internetHeaders.setAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

async function onMessageSendBugHandler(event) {
  console.log(`[bug] onMessageSendBugHandler started at ${new Date().toISOString()}`);
  const item = Office.context.mailbox.item;

  try {
    const currentHtml = await getBodyHtml(item);
    const banner = '<div style="border:2px solid #c53030;background:#fff5f5;padding:6px 10px;font-family:sans-serif;"><strong>CLASSIFICATION: INTERNAL</strong> &mdash; set via body.setAsync round-trip (BUG path)</div><br>';
    await setBodyHtml(item, banner + currentHtml);
    console.log('[bug] Body replaced via setAsync.');
  } catch (err) {
    console.error('[bug] Body round-trip failed.', err);
  }

  try {
    await setReproHeader(item);
    console.log('[bug] Set X-GV-Repro internet header.');
  } catch (err) {
    console.error('[bug] Failed to set internet header.', err);
  }

  console.log('[bug] Completing send, allowEvent=true.');
  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageSendBugHandler', onMessageSendBugHandler);
