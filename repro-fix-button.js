/**
 * Button-triggered variant of the fix, for environments where event-based
 * activation doesn't fire (e.g. a personal outlook.com account, which can't
 * admin-deploy). Same on-send stamping APIs as repro-fix.js
 * (prependOnSendAsync + appendOnSendAsync, NO body round-trip) - just armed by
 * a compose ribbon button instead of the OnMessageCompose event.
 *
 * Flow: open a compose window -> click "Stamp classification" in the ribbon ->
 * this arms the top header + bottom footer -> send -> open the message in Sent
 * Items to see both stamped. Watch the [fix-btn] lines in the browser dev tools.
 */

Office.onReady(() => {
  console.log('[fix-btn] Office.js ready.');
});

function getBodyType(item) {
  return new Promise((resolve) => {
    item.body.getTypeAsync((res) => {
      const type = res.status === Office.AsyncResultStatus.Succeeded ? res.value : Office.CoercionType.Html;
      console.log('[fix-btn] body.getTypeAsync returned.', { status: res.status, type, error: res.error });
      resolve(type);
    });
  });
}

function prependOnSend(item, data, coercionType) {
  return new Promise((resolve, reject) => {
    console.log(`[fix-btn] Calling prependOnSendAsync (coercionType=${coercionType})...`);
    item.body.prependOnSendAsync(data, { coercionType }, (res) => {
      console.log('[fix-btn] prependOnSendAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

function appendOnSend(item, data, coercionType) {
  return new Promise((resolve, reject) => {
    console.log(`[fix-btn] Calling appendOnSendAsync (coercionType=${coercionType})...`);
    item.body.appendOnSendAsync(data, { coercionType }, (res) => {
      console.log('[fix-btn] appendOnSendAsync returned.', { status: res.status, error: res.error });
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

function notify(message, icon) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync('fixBtnStatus', {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message.substring(0, 150),
      icon: icon,
      persistent: false,
    });
  } catch (e) { /* notifications are best-effort */ }
}

// Function-command handler wired to the ribbon button.
async function stampClassification(event) {
  console.log(`[fix-btn] stampClassification clicked at ${new Date().toISOString()}`);
  const item = Office.context.mailbox.item;

  try {
    const bodyType = await getBodyType(item);
    const isHtml = bodyType === Office.CoercionType.Html;

    const header = isHtml
      ? '<div style="border:2px solid #2b6cb0;background:#ebf4ff;padding:6px 10px;font-family:sans-serif;"><strong>CLASSIFICATION: INTERNAL</strong> &mdash; prepended on send</div><br>'
      : 'CLASSIFICATION: INTERNAL — prepended on send\n\n';

    const footer = isHtml
      ? '<br><div style="border-top:1px solid #a0aec0;padding-top:6px;font-family:sans-serif;color:#4a5568;"><em>This message is classified INTERNAL. Appended on send.</em></div>'
      : '\n\n---\nThis message is classified INTERNAL. Appended on send.';

    await prependOnSend(item, header, bodyType);
    await appendOnSend(item, footer, bodyType);
    console.log('[fix-btn] Header + footer armed for send.');
    notify('Classification armed: header + footer will be stamped on send.');
  } catch (err) {
    console.error('[fix-btn] Failed to arm on-send content.', err);
    notify('Failed to arm classification: ' + (err && err.message ? err.message : err));
  }

  event.completed();
}

Office.actions.associate('stampClassification', stampClassification);
