/** taskpane.js - FINAL FIXED VERSION for dropdown + version load */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");

    document.getElementById("saveJsonBtn")?.addEventListener("click", saveVersionAsJSON);
    document.getElementById("downloadXlsxBtn")?.addEventListener("click", downloadExcelFile);
    document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    document.getElementById("loadVersionBtn")?.addEventListener("click", handleVersionLoad);

    // ✅ Ensure dropdown is always populated on load
    populateVersionDropdown();
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
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const values = range.values;
    const headers = values[0];
    const data = values.slice(1);
    const jsonData = data.map(row => Object.fromEntries(row.map((val, i) => [headers[i], val])));

    const revision = {
      version: "manual-save",
      timestamp: new Date().toISOString(),
      user: Office.context?.userProfile?.displayName || "unknown",
      data: jsonData,
    };

    const blob = new Blob([JSON.stringify(revision, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `excel-manual-version.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
}

async function downloadExcelFile() {
  await Excel.run(async (context) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const file = result.value;
        const slices = [];
        let index = 0;

        const getSlice = () => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              slices.push(new Uint8Array(sliceResult.value.data));
              index++;
              if (index < file.sliceCount) {
                getSlice();
              } else {
                file.closeAsync();
                const totalLength = slices.reduce((sum, s) => sum + s.length, 0);
                const merged = new Uint8Array(totalLength);
                let offset = 0;
                for (const s of slices) {
                  merged.set(s, offset);
                  offset += s.length;
                }
                const blob = new Blob([merged], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = "excel-export.xlsx";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
              }
            }
          });
        };

        getSlice();
      }
    });
  });
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
    const jsonData = data.map(row => Object.fromEntries(row.map((val, i) => [headers[i], val])));

    let versionSheet;
    const sheets = context.workbook.worksheets;
    try {
      versionSheet = sheets.getItem("VersionHistory");
      await context.sync();
    } catch {
      versionSheet = sheets.add("VersionHistory");
    }
    versionSheet.visibility = Excel.SheetVisibility.hidden;

    const rangeCheck = versionSheet.getUsedRangeOrNullObject();
    rangeCheck.load("values, rowCount");
    await context.sync();

    const existing = rangeCheck.isNullObject ? [] : rangeCheck.values.slice(1);
    const newVersion = getNextVersion(existing);

    const newRow = [newVersion, new Date().toISOString(), Office.context?.userProfile?.displayName || "unknown", JSON.stringify(jsonData)];
    const targetRange = versionSheet.getRange(`A${existing.length + 2}:D${existing.length + 2}`);
    targetRange.values = [newRow];
    versionSheet.getRange("A1:D1").values = [["Version", "Timestamp", "User", "Data"]];
    await context.sync();
    console.log(`Version ${newVersion} saved.`);
    await populateVersionDropdown(newVersion); // Pass to pre-select the saved version
  });
}

async function populateVersionDropdown(selectVersion = null) {
  await Excel.run(async (context) => {
    const dropdown = document.getElementById("versionDropdown");
    dropdown.innerHTML = '<option value="">Select Version</option>';
    try {
      const versionSheet = context.workbook.worksheets.getItem("VersionHistory");
      versionSheet.load("name");
      const range = versionSheet.getUsedRange();
      range.load("values");
      await context.sync();

      const values = range.values.slice(1);
      values.forEach((row, i) => {
        const version = row[0];
        const opt = document.createElement("option");
        opt.value = i + 2;
        opt.textContent = version;
        if (version === selectVersion) opt.selected = true;
        dropdown.appendChild(opt);
      });
    } catch (e) {
      console.warn("VersionHistory sheet not found or not accessible.", e);
    }
  });
}

function handleVersionLoad() {
  const dropdown = document.getElementById("versionDropdown");
  const value = dropdown.value;
  if (!value) {
    console.log("Please select a version.");
    return;
  }
  loadSelectedVersion(parseInt(value));
}

async function loadSelectedVersion(rowIndex) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("VersionHistory");
    const row = sheet.getRange(`A${rowIndex}:D${rowIndex}`);
    row.load("values");
    await context.sync();

    const json = JSON.parse(row.values[0][3]);
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = activeSheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();

    if (!usedRange.isNullObject) usedRange.clear();

    if (!json || json.length === 0) {
      activeSheet.getRange("A1").values = [[""]];
      await context.sync();
      console.log(`Blank version loaded.`);
      return;
    }

    const headers = Object.keys(json[0]);
    const rows = [headers, ...json.map(obj => headers.map(h => obj[h]))];
    const writeRange = activeSheet.getRangeByIndexes(0, 0, rows.length, headers.length);
    writeRange.values = rows;
    await context.sync();
    console.log(`Version ${row.values[0][0]} loaded.`);
  });
}
