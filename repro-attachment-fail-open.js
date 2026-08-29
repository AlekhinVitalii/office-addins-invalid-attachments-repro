/**
 * Minimal repro for ItemSend attachment handling.
 *
 * getAttachments() below is a verbatim port of OutlookDocument.getAttachments()
 * as it currently stands in office-classifier — real getAttachmentsAsync /
 * getAttachmentContentAsync calls, no simulated/forced failures. Every step
 * is logged so a real failure (a slow/broken attachment, a cloud attachment,
 * an OWA attachment service hiccup, etc.) is fully visible in the console.
 *
 * The body is modified the same way office-classifier does it (see
 * OutlookDocument.setHtml): read the current HTML via body.getAsync(), then
 * replace the whole body in one body.setAsync() call. A custom internet
 * header (X-GV-Repro) is also set via internetHeaders.setAsync(), the same
 * way OutlookDocument.setHeaders() does it.
 *
 * To use: attach one or more files to a message, point a manifest's
 * ItemSend Events extension point (FunctionExecution="synchronous", matching
 * the real production manifest.outlook.xml — not the LaunchEvent/OnMessageSend
 * mechanism used by the block/soft-block variants) at this file, and send.
 * Watch the console for the [repro] lines, and open the sent message — the
 * "Header added and time when it was added" banner at the top of the body
 * (and the X-GV-Repro header, visible via the message source / EML) confirms
 * the send completed.
 */

Office.onReady(() => {
  console.log('[repro] Office.js ready.');
});

function withTimeout(operation, promise, timeoutMs = 15000) {
  console.log(`[repro][withTimeout] ${operation} started`);

  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      console.error(`[repro][withTimeout] ${operation} timed out after ${timeoutMs}ms`);
      reject(new Error(`Timed out after ${timeoutMs}ms: ${operation}`));
    }, timeoutMs);

    promise.then(
      (value) => {
        clearTimeout(timer);
        console.log(`[repro][withTimeout] ${operation} succeeded`);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        console.error(`[repro][withTimeout] ${operation} failed`, error);
        reject(error);
      }
    );
  });
}

// Verbatim port of OutlookDocument.getAttachments() (office-classifier) —
// real Office.js calls, no forced failures — with logging added at every step.
function getAttachments(item) {
  return new Promise((resolve) => {
    const options = { asyncContext: { currentItem: item } };

    console.log('[repro] Calling getAttachmentsAsync...');

    item.getAttachmentsAsync(options, (result) => {
      console.log('[repro] getAttachmentsAsync returned.', {
        status: result.status,
        count: result.value?.length,
        error: result.error,
      });

      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('[repro] getAttachmentsAsync failed.', result.error);
        resolve([]);
        return;
      }

      if (result.value.length > 0) {
        const contentPromises = [];

        for (const attachment of result.value) {
          console.log('[repro] Found attachment.', {
            id: attachment.id,
            name: attachment.name,
            isInline: attachment.isInline,
            attachmentType: attachment.attachmentType,
          });

          if (attachment.isInline) {
            console.log(`[repro] Skip processing inline attachment: ${attachment.name}`);
            continue;
          }

          if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File) {
            const contentPromise = new Promise((contentResolve) => {
              console.log(`[repro] Calling getAttachmentContentAsync for "${attachment.name}" (id=${attachment.id})...`);

              result.asyncContext.currentItem.getAttachmentContentAsync(attachment.id, (attachmentContent) => {
                console.log(`[repro] getAttachmentContentAsync returned for "${attachment.name}".`, {
                  status: attachmentContent.status,
                  format: attachmentContent.value?.format,
                  error: attachmentContent.error,
                });

                // A failed download (e.g. OWA attachment service 503) must resolve,
                // otherwise the OnMessageSend event never completes and the send hangs.
                if (attachmentContent.status === Office.AsyncResultStatus.Failed || !attachmentContent.value) {
                  console.error(`[repro] getAttachmentContent failed for "${attachment.name}".`, attachmentContent.error);
                  contentResolve(null);
                  return;
                }

                if (attachmentContent.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                  console.log(
                    `[repro] Got base64 content for "${attachment.name}", length=${attachmentContent.value.content.length}.`
                  );
                  // Allow file attachment, but not email item, .icalendar or cloud
                  contentResolve({
                    name: attachment.name,
                    base64Content: attachmentContent.value.content,
                  });
                } else {
                  console.log(`[repro] Non-base64 format for "${attachment.name}": ${attachmentContent.value.format} — skipping.`);
                  contentResolve(null);
                }
              });
            });
            contentPromises.push(contentPromise);
          }
        }

        if (contentPromises.length > 0) {
          Promise.all(contentPromises).then((values) => {
            // Remove empty values
            const filtered = values.filter((x) => x);
            console.log(`[repro] Attachment collection complete. ${filtered.length}/${contentPromises.length} succeeded.`);
            resolve(filtered);
          });
        } else {
          console.log('[repro] No file attachments to process (only inline, or none).');
          resolve([]);
        }
      } else {
        console.log('[repro] No attachments on this message.');
        resolve([]);
      }
    });
  });
}

function getBodyHtml(item) {
  return new Promise((resolve, reject) => {
    console.log('[repro] Calling body.getAsync(Html)...');

    item.body.getAsync(Office.CoercionType.Html, (res) => {
      console.log('[repro] body.getAsync returned.', { status: res.status, error: res.error });

      if (res.status === Office.AsyncResultStatus.Failed) {
        reject(res.error);
        return;
      }
      resolve(res.value);
    });
  });
}

// Visible proof-of-life, written the same way office-classifier writes the
// body: read the current HTML, then replace the whole body in one
// mailboxBody.setAsync() call (see OutlookDocument.setHtml), rather than
// body.prependAsync().
function setReproBanner(item) {
  return new Promise((resolve, reject) => {
    getBodyHtml(item)
      .then((currentHtml) => {
        const addedAt = new Date().toISOString();
        const banner = `<div style="border:2px solid #d9822b;background:#fff8e6;padding:8px 12px;margin-bottom:12px;font-family:sans-serif;">
          <strong>Header added and time when it was added:</strong> ${addedAt}
        </div>`;
        const htmlToSet = banner + currentHtml;

        console.log(`[repro] Calling body.setAsync(Html), addedAt=${addedAt}, length=${htmlToSet.length}...`);

        item.body.setAsync(htmlToSet, { coercionType: Office.CoercionType.Html }, (res) => {
          console.log('[repro] body.setAsync returned.', { status: res.status, error: res.error });

          if (res.status === Office.AsyncResultStatus.Failed) {
            reject(res.error);
            return;
          }
          resolve();
        });
      })
      .catch(reject);
  });
}

// Mirrors OutlookDocument.setHeaders() — internetHeaders.setAsync() with a
// plain header-name -> value map.
function setReproHeader(item) {
  return new Promise((resolve, reject) => {
    const addedAt = new Date().toISOString();
    const headers = { 'X-GV-Repro': addedAt };

    console.log(`[repro] Calling internetHeaders.setAsync(X-GV-Repro=${addedAt})...`);

    item.internetHeaders.setAsync(headers, (res) => {
      console.log('[repro] internetHeaders.setAsync returned.', { status: res.status, error: res.error });

      if (res.status === Office.AsyncResultStatus.Failed) {
        reject(res.error);
        return;
      }
      resolve();
    });
  });
}

async function onMessageSendHandler(event) {
  console.log(`[repro] onMessageSendHandler started at ${new Date().toISOString()}`);

  const item = Office.context.mailbox.item;

  // try {
  //   const attachments = await withTimeout('Collecting attachments', getAttachments(item), 15000);
  //   console.log('[repro] Collected attachments:', attachments);
  // } catch (error) {
  //   console.error('[repro] Attachment collection failed/timed out — failing open.', error);
  // }

  try {
    await setReproBanner(item);
    console.log('[repro] Set banner into body.');
  } catch (error) {
    console.error('[repro] Failed to set banner.', error);
  }

  // try {
  //   await setReproHeader(item);
  //   console.log('[repro] Set X-GV-Repro internet header.');
  // } catch (error) {
  //   console.error('[repro] Failed to set internet header.', error);
  // }

    // try {
    //     const attachments = await withTimeout('Check attachments 2nd time Collecting attachments', getAttachments(item), 15000);
    //     console.log('[repro] Check attachments 2nd time Collected attachments:', attachments);
    // } catch (error) {
    //     console.error('[repro] Check attachments 2nd time Attachment collection failed/timed out — failing open.', error);
    // }

  console.log('[repro] Completing event, allowEvent=true.');
  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
