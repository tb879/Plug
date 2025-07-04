Office.onReady(() => {
  document.getElementById("insertMetadata").onclick = insertMetadata;
  document.getElementById("enableTracking").onclick = trackChanges;
  document.getElementById("commitRevision").onclick = saveAndCommit;
  document.getElementById("see").onclick = unhideAllSheets;
});

function insertMetadata() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Define metadata fields
    const metadataFields = [
      "Document Title",
      "Document ID",
      "Revision Number",
      "Date of Issue",
      "Owner/Author",
      "Approver(s)",
      "Department/Team",
      "Standard",
    ];

    // Prepare data for Column A
    const labels = metadataFields.map((field) => [field]);

    // Insert labels into Column A (A1 to A8)
    const labelRange = sheet.getRange(`A1:A${metadataFields.length}`);
    labelRange.values = labels;

    // Clear Column B (user input cells)
    const valueRange = sheet.getRange(`B1:B${metadataFields.length}`);
    valueRange.clear(); // Make them empty

    // Optional: style
    labelRange.format.font.bold = true;
    valueRange.format.fill.color = "#FFF3CD"; // soft yellow for editable fields

    // Autofit both columns
    sheet.getRange("A1:B8").format.autofitColumns();

    await context.sync();
  });
}

// function insertMetadata() {
//   Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     sheet.getRange("A1").values = [["Document Title"]];
//     sheet.getRange("A2").values = [["Revision Number"]];
//     sheet.getRange("B1").values = [["ECCP Sample"]];
//     sheet.getRange("B2").values = [["Rev 1.0.0"]];
//     await context.sync();
//     setStatus("Metadata inserted.");
//   });
// }

function trackChanges() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.onChanged.add(async (event) => {
      const cell = event.address;
      const description = prompt(`Change detected at ${cell}. Enter reason:`);

      const logSheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
      await context.sync();

      const logger = logSheet.isNullObject
        ? context.workbook.worksheets.add("ChangeLog")
        : logSheet;
      const range = logger.getRange("A1").getSurroundingRegion().load("rowCount");
      await context.sync();

      const nextRow = range.rowCount + 1;
      logger.getRange(`A${nextRow}:C${nextRow}`).values = [
        [new Date().toLocaleString(), cell, description],
      ];
      await context.sync();

      setStatus("Change logged.");
    });

    setStatus("Change tracking enabled.");
  });
}

function saveAndCommit() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Load the revision value (assumed B3)
    const revCell = sheet.getRange("B3");
    revCell.load("values");

    await context.sync();

    let currentRev = revCell.values?.[0]?.[0];

    // Normalize the value to string
    if (typeof currentRev === "number") {
      // If user typed 1, treat it as base revision
      currentRev = `Rev ${currentRev}.0.0`;
    } else if (typeof currentRev === "string") {
      if (!currentRev.startsWith("Rev")) {
        currentRev = `Rev ${currentRev}`;
      }
    } else {
      console.error("Unsupported revision format:", currentRev);
      return;
    }

    // Extract revision numbers
    const parts = currentRev.match(/\d+/g)?.map(Number);
    if (!parts || parts.length < 3) {
      console.error("Invalid revision format. Expected format: Rev X.Y.Z");
      return;
    }

    // Increment patch version
    parts[2] += 1;
    const newRevision = `Rev ${parts[0]}.${parts[1]}.${parts[2]}`;

    // Write new revision back
    revCell.values = [[newRevision]];

    await context.sync();
    setStatus(`Revision updated to ${newRevision}`);
  });
}

// function saveAndCommit() {
//   Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();

//     // Step 1: Load revision number
//     const revRange = sheet.getRange("B3").load("values"); // B3 = Revision Number
//     const metadataRange = sheet.getRange("A1:B8").load("values"); // All metadata

//     await context.sync();

//     let currentRev = revRange.values?.[0]?.[0]; // "Rev 1.0.0"
//     if (typeof currentRev !== "string") {
//       console.error("Revision cell value is not a string:", currentRev);
//       return;
//     }
//     let parts = currentRev.match(/\d+/g).map(Number); // [1, 0, 0]
//     parts[2] += 1; // Increment patch version
//     const newRevision = `Rev ${parts[0]}.${parts[1]}.${parts[2]}`;

//     // Step 2: Update revision number in main sheet
//     revRange.values = [[newRevision]];

//     // Step 3: Save old metadata in hidden sheet "Revisions"
//     let revisionsSheet = context.workbook.worksheets.getItemOrNullObject("Revisions");
//     await context.sync();

//     if (revisionsSheet.isNullObject) {
//       revisionsSheet = context.workbook.worksheets.add("Revisions");
//       revisionsSheet.visibility = Excel.SheetVisibility.hidden;
//     }

//     const revLogRange = revisionsSheet.getRange("A1:H1"); // headers
//     revLogRange.values = [
//       [
//         "Timestamp",
//         "Revision",
//         "Document Title",
//         "Owner",
//         "Approvers",
//         "Department",
//         "Standard",
//         "Description",
//       ],
//     ];

//     const now = new Date().toLocaleString();
//     const description = prompt("Enter change description:");
//     const rowData = [
//       now,
//       newRevision,
//       metadataRange.values[0][1], // Title
//       metadataRange.values[4][1], // Owner
//       metadataRange.values[5][1], // Approvers
//       metadataRange.values[6][1], // Department
//       metadataRange.values[7][1], // Standard
//       description,
//     ];

//     const lastRow = revisionsSheet.getUsedRange().getLastRow().load("rowIndex");
//     await context.sync();

//     const nextRow = lastRow.rowIndex + 2;
//     revisionsSheet.getRange(`A${nextRow}:H${nextRow}`).values = [rowData];

//     // Step 4: Save changelog
//     let changeLogSheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
//     await context.sync();

//     if (changeLogSheet.isNullObject) {
//       changeLogSheet = context.workbook.worksheets.add("ChangeLog");
//       changeLogSheet.getRange("A1:D1").values = [
//         ["Timestamp", "Initials", "Revision", "Description"],
//       ];
//     }

//     const userInitials = "TB"; // In future: dynamic from login
//     const changeLogRow = [now, userInitials, newRevision, description];

//     const lastLogRow = changeLogSheet.getUsedRange().getLastRow().load("rowIndex");
//     await context.sync();

//     const nextLogRow = lastLogRow.rowIndex + 2;
//     changeLogSheet.getRange(`A${nextLogRow}:D${nextLogRow}`).values = [changeLogRow];

//     await context.sync();
//     setStatus(`Committed ${newRevision}`);
//   });
// }


// Force unhide Revisions and ChangeLog (for debugging)
function unhideAllSheets() {
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

function setStatus(msg) {
  document.getElementById("status").innerText = msg;
}
