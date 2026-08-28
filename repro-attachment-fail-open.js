/**
 * Minimal repro for the "failed attachment download" fail-open path.
 *
 * Mirrors OutlookDocument.getAttachments() + SendMessageHandlerV2's contract:
 *   - a failed attachment content download REJECTS (does not silently resolve/hang)
 *   - the OnMessageSend handler must still complete the event (fail open)
 *
 * The content-fetch step below is forced to fail for every non-inline file
 * attachment, instead of really calling getAttachmentContentAsync, so the
 * bug is reproducible without needing a genuinely broken attachment / a
 * flaky OWA attachment service.
 *
 * To use: attach any file to a message, point a manifest's OnMessageSend
 * LaunchEvent at this file (as the runtime script), and send. Watch the
 * console (or Outlook's "System logs" if wired into the real app) for the
 * [repro] lines, and confirm the added X-GV-Repro header lands even though
 * attachment collection "failed".
 */

Office.onReady(() => {
  console.log('[repro] Office.js ready.');
});

function withTimeout(operation, promise, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`Timed out after ${timeoutMs}ms: ${operation}`)), timeoutMs);

    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      }
    );
  });
}

// Same shape as OutlookDocument.getAttachments(), except the content-fetch
// step is forced to fail for every non-inline file attachment.
function getAttachmentsForcingFailure(item) {
  return new Promise((resolve, reject) => {
    item.getAttachmentsAsync({ asyncContext: { currentItem: item } }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
        return;
      }

      const fileAttachments = result.value.filter(
        (a) => !a.isInline && a.attachmentType === Office.MailboxEnums.AttachmentType.File
      );

      if (fileAttachments.length === 0) {
        console.warn('[repro] No file attachments found — attach any file to exercise the failure path.');
        resolve([]);
        return;
      }

      // FORCED FAILURE: simulate every attachment content download failing
      // (e.g. an OWA attachment service 503) instead of calling the real
      // getAttachmentContentAsync.
      const failures = fileAttachments.map(
        (a) =>
          new Promise((_resolve, contentReject) =>
            contentReject({ message: `Simulated attachment service failure for "${a.name}"` })
          )
      );

      Promise.all(failures).then(resolve).catch(reject);
    });
  });
}

function addReproHeader(item) {
  return new Promise((resolve, reject) => {
    item.internetHeaders.setAsync({ 'X-GV-Repro': `attachment-fetch-failed-${Date.now()}` }, (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        reject(res.error);
        return;
      }
      resolve();
    });
  });
}

async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  try {
    console.log('[repro] Collecting attachments (forced failure)...');
    const attachments = await withTimeout('Collecting attachments', getAttachmentsForcingFailure(item), 15000);
    console.log('[repro] Collected attachments (only reached if there were none to fail):', attachments);
  } catch (error) {
    // This is the path the fix guarantees: getAttachments() rejects instead
    // of hanging, and the caller must treat it as fail-open.
    console.warn('[repro] Attachment collection failed as expected — proving the fail-open path.', error);
  }

  try {
    await addReproHeader(item);
    console.log('[repro] Added X-GV-Repro header — send proceeds despite the attachment failure.');
  } catch (error) {
    console.error('[repro] Failed to add header.', error);
  }

  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
