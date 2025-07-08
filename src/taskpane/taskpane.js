/** taskpane.js - FINAL FIXED VERSION for dropdown + version load */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");

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

async function saveAndCommitVersion() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const values = range.values;
    const headers = values.length ? values[0] : [];
    const data = values.length > 1 ? values.slice(1) : [];
    const jsonData = headers.length ? data.map(row => Object.fromEntries(row.map((val, i) => [headers[i], val]))) : [];

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
    await populateVersionDropdown(newVersion);
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
