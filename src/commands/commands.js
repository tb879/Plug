const Office = require("O")

Office.onReady(() => {
  console.log("ECCP Compliance Tools loaded.");
});

// 1. Create new ECCP document
function newDocument() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1:B8").values = [
      ["Document Title", ""],
      ["Document ID", ""],
      ["Revision Number", "Rev 1.0.0"],
      ["Date of Issue", new Date().toLocaleDateString()],
      ["Owner/Author", ""],
      ["Approver(s)", ""],
      ["Department/Team", ""],
      ["Standard", ""],
    ];
    sheet.getRange("A1:A8").format.font.bold = true;
    sheet.getRange("A1:B8").format.fill.color = "#FAF3D1";
    await context.sync();
  });
}

// 2. Save & Commit Revision
function saveAndCommit() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const revisionCell = sheet.getRange("B3");
    revisionCell.load("values");
    await context.sync();

    let rev = revisionCell.values[0][0]; // e.g. "Rev 1.0.0"
    let revMatch = rev.match(/(\d+)\.(\d+)\.(\d+)/);
    if (!revMatch) {
      console.error("Invalid revision format");
      return;
    }

    let [_, major, minor, patch] = revMatch.map(Number);
    patch++;
    const newRev = `Rev ${major}.${minor}.${patch}`;
    revisionCell.values = [[newRev]];
    await context.sync();

    console.log(`Revision updated to ${newRev}`);
  });
}

// 3. Edit Metadata (open task pane if needed)
function editMetadata() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/src/taskpane/taskpane.html",
    { height: 45, width: 30, displayInIframe: true }
  );
}

// 4. Track Changes
function trackChanges() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    // Placeholder logic for enabling track changes on critical zones
    console.log("Tracking enabled on critical zones.");
    await context.sync();
  });
}

// 5. Send for Review
function sendForReview() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("B1").comments.add("Sent for review by user at " + new Date().toLocaleString());
    await context.sync();
  });
}

// 6. Approve Document
function approveDocument() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("B1").comments.add("Approved by user at " + new Date().toLocaleString());
    await context.sync();
  });
}

// 7. Export to PDF (this needs a workaround, stubbed for now)
function exportPDF() {
  console.log("PDF export requested. This must be done via Office desktop automation or SharePoint integration.");
}

// 8. View Change Log (open task pane view)
function viewChangeLog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/src/taskpane/taskpane.html?view=log",
    { height: 50, width: 40, displayInIframe: true }
  );
}

// Export functions to global scope
window.newDocument = newDocument;
window.saveAndCommit = saveAndCommit;
window.editMetadata = editMetadata;
window.trackChanges = trackChanges;
window.sendForReview = sendForReview;
window.approveDocument = approveDocument;
window.exportPDF = exportPDF;
window.viewChangeLog = viewChangeLog;
