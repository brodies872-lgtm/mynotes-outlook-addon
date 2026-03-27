let currentEmail = null;

function $(id){ return document.getElementById(id); }

function setStatus(text){
  $("statusLine").textContent = text;
}

function htmlToText(html){
  const div = document.createElement("div");
  div.innerHTML = html || "";
  return (div.textContent || div.innerText || "").trim();
}

function inferPriority(text){
  const haystack = String(text || "").toLowerCase();
  if (/\b(urgent|asap|critical|immediately|priority)\b/.test(haystack)) return "high";
  if (/\b(review|investigate|check|follow up|follow-up)\b/.test(haystack)) return "medium";
  return "medium";
}

function senderAddressFromDisplay(display){
  const match = String(display || "").match(/<([^>]+)>/);
  return match ? match[1].trim() : String(display || "").trim();
}

function buildPayload(email){
  const summaryOverride = $("summaryInput").value.trim();
  const tags = $("tagsInput").value
    .split(/[;,]+/)
    .map(v => v.trim())
    .filter(Boolean);

  const bodyPreview = (email.bodyText || "").slice(0, 1200).trim();
  const summary = summaryOverride || bodyPreview.slice(0, 280);

  return {
    type: "mynotes-email-capture",
    version: 1,
    capturedAt: new Date().toISOString(),
    source: {
      platform: "outlook",
      messageId: email.itemId || "",
      conversationId: email.conversationId || "",
      folder: email.folder || ""
    },
    email: {
      subject: email.subject || "",
      from: email.from || "",
      to: email.to || [],
      date: email.date || "",
      bodyPreview,
      attachments: (email.attachments || []).map(name => ({ name }))
    },
    action: {
      target: $("targetInput").value || "task",
      title: email.subject || "Email capture",
      summary,
      priority: $("priorityInput").value || inferPriority(email.subject + " " + bodyPreview),
      tags
    }
  };
}

function renderPayload(){
  if (!currentEmail) return;
  const payload = buildPayload(currentEmail);
  $("payloadPreview").value = JSON.stringify(payload, null, 2);
}

function renderEmail(email){
  currentEmail = email;
  $("subjectValue").textContent = email.subject || "—";
  $("fromValue").textContent = email.from || "—";
  $("dateValue").textContent = email.date || "—";
  $("attachmentsValue").textContent = email.attachments.length ? email.attachments.join(", ") : "None";
  $("summaryInput").value = (email.bodyText || "").slice(0, 280).trim();
  $("priorityInput").value = inferPriority((email.subject || "") + " " + (email.bodyText || ""));
  renderPayload();
}

function getAttachmentsAsync(item){
  return new Promise(resolve => {
    try{
      item.getAttachmentsAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve((result.value || []).map(a => a.name));
        } else {
          resolve([]);
        }
      });
    }catch{
      resolve([]);
    }
  });
}

async function readCurrentEmail(){
  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus("No Outlook item is open.");
    return;
  }

  const bodyText = await new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Text, result => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });

  const attachments = await getAttachmentsAsync(item);

  const email = {
    itemId: item.itemId || "",
    conversationId: item.conversationId || "",
    subject: item.subject || "",
    from: senderAddressFromDisplay(item.from?.displayName || item.from?.emailAddress || ""),
    to: (item.to || []).map(r => r.emailAddress || r.displayName).filter(Boolean),
    date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString(),
    bodyText,
    attachments
  };

  renderEmail(email);
  setStatus("Email loaded.");
}

async function copyPayload(){
  if (!currentEmail) {
    setStatus("No email loaded.");
    return;
  }
  const payload = JSON.stringify(buildPayload(currentEmail), null, 2);
  try{
    await navigator.clipboard.writeText(payload);
    setStatus("MyNotes JSON copied to clipboard.");
  }catch(err){
    setStatus("Clipboard copy failed. You can still copy from the preview box.");
  }
}

Office.onReady(() => {
  $("copyJsonBtn").addEventListener("click", copyPayload);
  $("refreshBtn").addEventListener("click", readCurrentEmail);
  $("targetInput").addEventListener("change", renderPayload);
  $("priorityInput").addEventListener("change", renderPayload);
  $("tagsInput").addEventListener("input", renderPayload);
  $("summaryInput").addEventListener("input", renderPayload);
  readCurrentEmail();
});
