/** taskpane.js */

// Wait for Office.js to be ready
Office.onReady((info) => {
    console.log("Excel Add-in is ready")
});

function getNextVersion() {
  const versionKey = "excel-revision-version";
  let current = localStorage.getItem(versionKey) || "1.0.0";
  let [major, minor, patch] = current.split(".").map(Number);
  patch += 1;
  const next = `${major}.${minor}.${patch}`;
  localStorage.setItem(versionKey, next);
  return current;
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

      const version = getNextVersion();
      const revision = {
        version,
        filename: `excel-revision-v${version}.json`,
        user: Office.context?.userProfile?.displayName || "unknown",
        timestamp: new Date().toISOString(),
        comment: "Auto-saved structured data as JSON",
        headers,
        rows: dataAsJson,
      };

      const jsonBlob = new Blob([JSON.stringify(revision, null, 2)], {
        type: "application/json",
      });

      const url = URL.createObjectURL(jsonBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `excel-revision-v${version}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      alert(`Revision v${version} saved as readable JSON.`);
    });
  } catch (err) {
    console.error("Excel run failed", err);
    alert("Failed to save revision. See console for details.");
  }
}

async function downloadExcelFile() {
  try {
    await Excel.run(async (context) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
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
      });
    });
  } catch (err) {
    console.error("Excel run failed:", err);
  }
}
