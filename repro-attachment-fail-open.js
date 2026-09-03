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

// Lists inline attachments only (id/name/isInline), no content download —
// used by findBrokenInlineAttachment below.
function getInlineAttachments(item) {
  return new Promise((resolve) => {
    console.log('[repro] Calling getAttachmentsAsync (inline check)...');

    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('[repro] getAttachmentsAsync (inline check) failed.', result.error);
        resolve([]);
        return;
      }

      const inlineAttachments = result.value.filter((attachment) => attachment.isInline);
      console.log(`[repro] Found ${inlineAttachments.length} inline attachment(s).`, inlineAttachments);
      resolve(inlineAttachments);
    });
  });
}

// Confirmed-broken inline images came back status: 'succeeded' with tiny
// non-empty base64 content (~123-171 raw bytes) instead of failing outright —
// almost certainly a placeholder, not real image data. A real inline
// photo/screenshot is realistically many KB once base64-encoded, so anything
// under this floor is treated as broken too. Arbitrary threshold picked well
// above the observed placeholder sizes and well below a real image; revisit
// if it misclassifies a legitimately tiny inline image (e.g. a 1x1 spacer).
const MIN_INLINE_ATTACHMENT_CONTENT_LENGTH = 512;

// Checks each inline attachment's content via getAttachmentContentAsync,
// stopping as soon as one comes back broken (failed status, no value, or
// suspiciously small content). Returns the first broken attachment found, or
// null if all inline attachments downloaded fine.
async function findBrokenInlineAttachment(item) {
  const inlineAttachments = await getInlineAttachments(item);

  for (const attachment of inlineAttachments) {
    const result = await new Promise((resolve) => {
      console.log(`[repro] Calling getAttachmentContentAsync for inline "${attachment.name}" (id=${attachment.id})...`);
      item.getAttachmentContentAsync(attachment.id, resolve);
    });

    const contentLength = result.value?.content?.length ?? 0;
    const broken =
      result.status === Office.AsyncResultStatus.Failed ||
      !result.value ||
      contentLength < MIN_INLINE_ATTACHMENT_CONTENT_LENGTH;

    console.log(`[repro] Inline attachment content check for "${attachment.name}".`, {
      status: result.status,
      format: result.value?.format,
      contentLength,
      error: result.error,
      broken,
    });

    if (broken) {
      return attachment;
    }
  }

  return null;
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

// Same visible proof-of-life as setReproBanner, but written via
// body.prependAsync() instead of a getAsync()+setAsync() round trip.
function setReproBannerPrepend(item) {
  return new Promise((resolve, reject) => {
    const addedAt = new Date().toISOString();
    const banner = `<div style="border:2px solid #d9822b;background:#fff8e6;padding:8px 12px;margin-bottom:12px;font-family:sans-serif;">
      <strong>Prepended via body.prependAsync at:</strong> ${addedAt}
    </div>`;

    console.log(`[repro] Calling body.prependAsync(Html), addedAt=${addedAt}, length=${banner.length}...`);

    item.body.prependAsync(banner, { coercionType: Office.CoercionType.Html }, (res) => {
      console.log('[repro] body.prependAsync returned.', { status: res.status, error: res.error });

      if (res.status === Office.AsyncResultStatus.Failed) {
        reject(res.error);
        return;
      }
      resolve();
    });
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

  try {
    const attachments = await getAttachments(item);
    console.log('[repro] Collected attachments:', attachments);
  } catch (error) {
    console.error('[repro] Attachment collection failed/timed out — failing open.', error);
  }

  // let brokenInlineAttachment = null;
  // try {
  //   brokenInlineAttachment = await findBrokenInlineAttachment(item);
  // } catch (error) {
  //   console.error('[repro] Inline attachment check failed — defaulting to setBody.', error);
  // }

  // if (brokenInlineAttachment) {
  //   console.log(`[repro] Broken inline attachment detected: "${brokenInlineAttachment.name}" — using prependAsync instead of setBody.`);
    try {
      await setReproBannerPrepend(item);
      console.log('[repro] Prepended banner into body.');
    } catch (error) {
      console.error('[repro] Failed to prepend banner.', error);
    }
  // } else {
  //   try {
  //     await setReproBanner(item);
  //     console.log('[repro] Set banner into body.');
  //   } catch (error) {
  //     console.error('[repro] Failed to set banner.', error);
  //   }
  // }

  try {
    await setReproHeader(item);
    console.log('[repro] Set X-GV-Repro internet header.');
  } catch (error) {
    console.error('[repro] Failed to set internet header.', error);
  }

  console.log('[repro] Completing event, allowEvent=true.');
  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
