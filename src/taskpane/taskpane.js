Office.onReady(() => {
  console.log("ECCP Add-in loaded");
});

function setStatus(msg) {
  document.getElementById("status").innerText = msg;
}

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
      ["Standard", ""]
    ];
    await context.sync();
    setStatus("New document created");
  });
}

function insertMetadata() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const fields = ["Title", "ID", "Revision", "Date", "Author", "Approver", "Department", "Standard"];
    sheet.getRange("A1:A8").values = fields.map(x => [x]);
    sheet.getRange("B1:B8").clear();
    await context.sync();
    setStatus("Metadata inserted");
  });
}

function saveAndCommit() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange("B3");
    cell.load("values");
    await context.sync();

    const rev = cell.values[0][0];
    const match = rev.match(/(\d+)\.(\d+)\.(\d+)/);
    let [_, major, minor, patch] = match.map(Number);
    patch++;
    const newRev = `Rev ${major}.${minor}.${patch}`;
    cell.values = [[newRev]];
    await context.sync();
    setStatus(`Revision updated to ${newRev}`);
  });
}

function editMetadata() {
  Office.context.ui.displayDialogAsync("https://tb879.github.io/Plug/taskpane.html", {
    height: 45, width: 30
  });
}

function trackChanges() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(async (event) => {
      const desc = prompt(`Change detected at ${event.address}. Describe:`);
      const logSheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
      await context.sync();
      const log = logSheet.isNullObject ? context.workbook.worksheets.add("ChangeLog") : logSheet;
      const nextRow = log.getUsedRange().getLastRow().rowIndex + 2;
      log.getRange(`A${nextRow}:C${nextRow}`).values = [[new Date().toLocaleString(), event.address, desc]];
      await context.sync();
    });
    setStatus("Tracking changes enabled");
  });
}

function sendForReview() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("B1").comments.add("Sent for review @" + new Date().toLocaleString());
    await context.sync();
    setStatus("Document marked for review");
  });
}

function approveDocument() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("B1").comments.add("Approved @" + new Date().toLocaleString());
    await context.sync();
    setStatus("Document approved");
  });
}

function exportPDF() {
  console.warn("Export to PDF only works with desktop integration");
  setStatus("Export PDF is a future feature");
}

function viewChangeLog() {
  Office.context.ui.displayDialogAsync("https://tb879.github.io/Plug/taskpane.html?view=log", {
    height: 50, width: 40
  });
}