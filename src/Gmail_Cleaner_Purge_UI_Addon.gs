
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Gmail Cleaner")
    .addItem("Run Inbox Cleanup", "cleanInboxAndLog")
    .addItem("Run Weekly Summary", "sendWeeklySummary")
    .addSeparator()
    .addItem("Mark All Inbox as Read & Archive", "confirmAndPurgeInbox")
    .addToUi();
}

function confirmAndPurgeInbox() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Are you sure you want to mark ALL inbox emails as read and archive them?",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    archiveAllInboxAndMarkRead();
    ui.alert("✅ All inbox emails marked as read and archived.");
  } else {
    ui.alert("❌ Action canceled.");
  }
}

function archiveAllInboxAndMarkRead() {
  const batchSize = 100;
  const threads = GmailApp.search('in:inbox');
  for (let i = 0; i < threads.length; i += batchSize) {
    let batch = threads.slice(i, i + batchSize);
    GmailApp.markThreadsRead(batch);
    GmailApp.moveThreadsToArchive(batch);
    Utilities.sleep(500);
  }
}
