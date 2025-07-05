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
  // Office.context.ui.displayDialogAsync("https://tb879.github.io/Plug/taskpane.html?view=log", {
  //   height: 50, width: 40
  // });
console.log("TASK....");

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

const saveVersion = async () => {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Get the workbook as a file
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const file = result.value;
          const sliceCount = file.sliceCount;
          const slices = [];
          let sliceIndex = 0;

          const getSlice = () => {
            file.getSliceAsync(sliceIndex, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                slices.push(sliceResult.value.data);
                sliceIndex++;
                if (sliceIndex < sliceCount) {
                  getSlice();
                } else {
                  file.closeAsync();
                  const blob = new Blob(slices);
                  saveAsJSON(blob);
                }
              } else {
                console.error("Failed to get slice", sliceResult.error.message);
              }
            });
          };

          getSlice();
        } else {
          console.error("Failed to get file", result.error.message);
        }
      });
    });
  } catch (err) {
    console.error("Excel run failed", err);
  }
};

const saveAsJSON = (blob) => {
  const reader = new FileReader();

  reader.onload = () => {
    const base64Data = reader.result.split(',')[1]; // Remove prefix
    const revision = {
      filename: `excel-version-${new Date().toISOString()}.xlsx`,
      user: Office.context?.userProfile?.displayName || "unknown",
      timestamp: new Date().toISOString(),
      comment: prompt("Enter a comment for this revision:") || "",
      fileData: base64Data,
    };

    const jsonBlob = new Blob([JSON.stringify(revision, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(jsonBlob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `excel-revision-${new Date().toISOString()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    alert("Revision saved as JSON.");
  };

  reader.readAsDataURL(blob);
};
