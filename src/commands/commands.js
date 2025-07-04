Office.onReady(() => {
  console.log("ECCP ribbon commands ready.");
});

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://tb879.github.io/Plug/src/taskpane/taskpane.html",
    { height: 50, width: 40 }
  );
}

// Export all functions used in the ribbon
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

function viewChangeLog() {
  console.log("COMMANDS.....");

  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();
    sheets.items.forEach(sheet => {
      sheet.visibility = Excel.SheetVisibility.visible;
    });
    await context.sync();
    console.log("All sheets made visible.");
  });
}

// Expose globally
window.saveAndCommit = saveAndCommit;
window.viewChangeLog = viewChangeLog;
