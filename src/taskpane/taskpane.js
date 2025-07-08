/** taskpane.js */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");
    // document.getElementById("saveJsonBtn")?.addEventListener("click", saveVersionAsJSON);
    // document.getElementById("downloadXlsxBtn")?.addEventListener("click", downloadExcelFile);
    // document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    // document.getElementById("loadVersionBtn")?.addEventListener("click", loadSelectedVersion);
    window.addEventListener("load", populateVersionDropdown);
  }
});

function getNextVersion(existingVersions) {
  if (!existingVersions.length) return "1.0.0";
  const lastVersion = existingVersions[existingVersions.length - 1][0];
  let [major, minor, patch] = lastVersion.split(".").map(Number);
  patch++;
  return `${major}.${minor}.${patch}`;
}

async function saveVersionAsJSON() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const values = range.values;
      const headers = values[0];
      const rows = values.slice(1);

      const dataAsJson = rows.map((row) =>
        Object.fromEntries(row.map((cell, i) => [headers[i], cell]))
      );

      const version = "manual-save";
      const revision = {
        version,
        filename: `excel-version-${Date.now()}.json`,
        user: Office.context?.userProfile?.displayName || "unknown",
        timestamp: new Date().toISOString(),
        comment: "Manual JSON Export",
        headers,
        rows: dataAsJson,
      };

      const jsonBlob = new Blob([JSON.stringify(revision, null, 2)], {
        type: "application/json",
      });

      const url = URL.createObjectURL(jsonBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = revision.filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      console.log(`Data exported as JSON.`);
    });
  } catch (err) {
    console.log("Excel run failed");
    console.log("Failed to save JSON.");
  }
}

async function downloadExcelFile() {
  try {
    await Excel.run(async (context) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            const sliceCount = file.sliceCount;
            const slices = [];
            let sliceIndex = 0;

            const getSlice = () => {
              file.getSliceAsync(sliceIndex, (sliceResult) => {
                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                  slices.push(new Uint8Array(sliceResult.value.data));
                  sliceIndex++;
                  if (sliceIndex < sliceCount) {
                    getSlice();
                  } else {
                    file.closeAsync();
                    const totalLength = slices.reduce((sum, arr) => sum + arr.length, 0);
                    const mergedArray = new Uint8Array(totalLength);
                    let offset = 0;
                    for (const arr of slices) {
                      mergedArray.set(arr, offset);
                      offset += arr.length;
                    }
                    const blob = new Blob([mergedArray], {
                      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });

                    const url = URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = "exported-excel-file.xlsx";
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                  }
                } else {
                  console.error("Failed to get slice:", sliceResult.error.message);
                }
              });
            };

            getSlice();
          } else {
            console.error("Failed to get file:", result.error.message);
          }
        }
      );
    });
  } catch (err) {
    console.error("Excel run failed:", err);
  }
}

async function saveAndCommitVersion() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const values = range.values;
    const headers = values[0];
    const data = values.slice(1);
    const jsonData = data.map((row) => Object.fromEntries(row.map((val, i) => [headers[i], val])));

    let versionSheet;
    const sheets = context.workbook.worksheets;
    try {
      versionSheet = sheets.getItem("VersionHistory");
      versionSheet.load("visibility");
      await context.sync();
    } catch {
      versionSheet = sheets.add("VersionHistory");
    }

    versionSheet.visibility = Excel.SheetVisibility.hidden;

    const historyRange = versionSheet.getUsedRangeOrNullObject();
    historyRange.load("values, rowCount");
    await context.sync();

    const existing = historyRange.isNullObject ? [] : historyRange.values.slice(1);
    const newVersion = getNextVersion(existing);

    const newRow = [
      newVersion,
      new Date().toISOString(),
      Office.context?.userProfile?.displayName || "unknown",
      JSON.stringify(jsonData),
    ];

    const targetRange = versionSheet.getRange(`A${existing.length + 2}:D${existing.length + 2}`);
    targetRange.values = [newRow];
    versionSheet.getRange("A1:D1").values = [["Version", "Timestamp", "User", "Data"]];

    await context.sync();
    console.log(`Version ${newVersion} committed to hidden sheet.`);
    populateVersionDropdown();
  });
}

async function populateVersionDropdown() {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getItem("VersionHistory");
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();

      const values = range.values.slice(1);
      const dropdown = document.getElementById("versionDropdown");
      dropdown.innerHTML = '<option value="">Select Version</option>';
      values.forEach((row, index) => {
        const version = row[0];
        dropdown.innerHTML += `<option value="${index + 2}">${version}</option>`;
      });
    } catch {
      console.log("VersionHistory sheet not found.");
    }
  });
}

async function loadSelectedVersion() {
  const rowIndex = document.getElementById("versionDropdown").value;
  if (!rowIndex) {
    console.log("Select a version.");
    return;
  }

  await Excel.run(async (context) => {
    const versionSheet = context.workbook.worksheets.getItem("VersionHistory");
    const versionRow = versionSheet.getRange(`A${rowIndex}:D${rowIndex}`);
    versionRow.load("values");
    await context.sync();

    const json = JSON.parse(versionRow.values[0][3]);
    const headers = Object.keys(json[0]);
    const allData = [headers, ...json.map((obj) => headers.map((h) => obj[h]))];

    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const range = activeSheet.getRangeByIndexes(0, 0, allData.length, headers.length);
    range.values = allData;
    await context.sync();
    console.log(`Version ${versionRow.values[0][0]} loaded into sheet.`);
  });
}
