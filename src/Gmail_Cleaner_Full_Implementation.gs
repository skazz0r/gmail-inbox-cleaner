
// Gmail Inbox Cleaner Script - Fully Implemented Version

function cleanInboxAndLog() {
  const props = PropertiesService.getScriptProperties();
  let sheetUrl = props.getProperty("sheetUrl");
  let ss;

  if (!sheetUrl) {
    ss = SpreadsheetApp.create("Gmail Inbox Cleanup Log");
    sheetUrl = ss.getUrl();
    props.setProperty("sheetUrl", sheetUrl);
  } else {
    ss = SpreadsheetApp.openByUrl(sheetUrl);
    updateDashboard();
  }

  const logSheet = ensureSheet(ss, "Log", ["Date", "Subject", "Label Applied", "Sender", "Received Date", "Sender Domain"]);
  const trustedSheet = ensureSheet(ss, "Trusted", ["Trusted Email or Domain"]);
  const senderSheet = ensureSheet(ss, "Unique Senders", ["Sender", "Domain"]);
  const trustedList = trustedSheet.getRange(2, 1, Math.max(trustedSheet.getLastRow() - 1, 1), 1).getValues().flat().map(s => s.toLowerCase());
  const existingSenders = senderSheet.getRange(2, 1, Math.max(senderSheet.getLastRow() - 1, 1), 1).getValues().flat();

  const whitelist = ["chase.com", "bankofamerica.com", "school", "reset", "verification", "enrollment"];
  const jobKeywords = ["linkedin", "glassdoor", "indeed", "ziprecruiter", "hiring"];
  const receiptKeywords = /(receipt|invoice|order confirmation|payment confirmation|thank you for your purchase)/i;
  const importantKeywords = /(verification code|reset password|account alert|bank|chase|school district|enrollment|security code|statement)/i;

  const threads = GmailApp.search('category:promotions is:unread');
  const purchaseLabel = GmailApp.getUserLabelByName("Purchase") || GmailApp.createLabel("Purchase");
  const unsubLabel = GmailApp.getUserLabelByName("Review_Unsubscribe") || GmailApp.createLabel("Review_Unsubscribe");
  const jobLabel = GmailApp.getUserLabelByName("Job Alerts") || GmailApp.createLabel("Job Alerts");

  const labelStats = { purchase: 0, unsub: 0, job: 0 };
  const newSenders = new Set();

  for (const thread of threads.slice(0, 100)) {
    const messages = thread.getMessages();
    let handled = false;

    for (const msg of messages) {
      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      const sender = msg.getFrom().toLowerCase();
      const receivedDate = msg.getDate();
      const domain = extractDomain(sender);
      const now = new Date();

      const isTrusted = trustedList.some(t => sender.includes(t)) || threadHasLabel(thread, "TrustedSender");
      if (isTrusted) break;

      if (!existingSenders.includes(sender)) {
        senderSheet.appendRow([sender, domain]);
        newSenders.add(sender);
      }

      if (importantKeywords.test(subject + body)) break;

      if (receiptKeywords.test(subject + body)) {
        purchaseLabel.addToThread(thread);
        logRowWithColor(logSheet, [now, subject, "Purchase", sender, receivedDate, domain], "#d9ead3");
        thread.moveToArchive();
        labelStats.purchase++;
        handled = true;
        break;
      }

      if (jobKeywords.some(k => sender.includes(k) || subject.toLowerCase().includes(k))) {
        jobLabel.addToThread(thread);
        logRowWithColor(logSheet, [now, subject, "Job Alerts", sender, receivedDate, domain], "#cfe2f3");
        thread.moveToArchive();
        labelStats.job++;
        handled = true;
        break;
      }
    }

    if (!handled && thread.isUnread()) {
      const msg = thread.getMessages()[0];
      const subject = msg.getSubject();
      const sender = msg.getFrom().toLowerCase();
      const receivedDate = msg.getDate();
      const domain = extractDomain(sender);
      const now = new Date();

      unsubLabel.addToThread(thread);
      thread.markRead();
      thread.moveToArchive();
      logRowWithColor(logSheet, [now, subject, "Review_Unsubscribe", sender, receivedDate, domain], "#f4cccc");
      labelStats.unsub++;
    }
  }

  autoDeleteOldUnsub(180);
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function logRowWithColor(sheet, rowData, color) {
  sheet.appendRow(rowData);
  const row = sheet.getLastRow();
  sheet.getRange(row, 1, 1, rowData.length).setBackground(color);
}

function extractDomain(sender) {
  const match = sender.match(/@([\w.-]+)/);
  return match ? match[1] : "unknown";
}

function threadHasLabel(thread, labelName) {
  return thread.getLabels().some(label => label.getName().toLowerCase() === labelName.toLowerCase());
}

function autoDeleteOldUnsub(days) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const threads = GmailApp.search(`label:Review_Unsubscribe before:${Utilities.formatDate(cutoff, Session.getScriptTimeZone(), "yyyy/MM/dd")}`);
  threads.forEach(t => t.moveToTrash());
}
