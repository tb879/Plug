/** taskpane.js - Updated with UI-driven version list rendering */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");

    document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);

    // Load version list on startup
    renderVersionHistory();
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
    if (values.length === 0 || values[0].length === 0) return console.log("Cannot save empty sheet");

    const headers = values[0];
    const data = values.slice(1);
    const jsonData = data.map(row => Object.fromEntries(row.map((val, i) => [headers[i], val])));

    let versionSheet;
    try {
      versionSheet = context.workbook.worksheets.getItem("VersionHistory");
    } catch {
      versionSheet = context.workbook.worksheets.add("VersionHistory");
      versionSheet.visibility = Excel.SheetVisibility.hidden;
    }

    const used = versionSheet.getUsedRangeOrNullObject();
    used.load("values, rowCount");
    await context.sync();

    const existing = used.isNullObject ? [] : used.values.slice(1);
    const newVersion = getNextVersion(existing);
    const timestamp = new Date().toISOString();
    const user = "Jay Yadav"; // Static placeholder

    const newRow = [newVersion, timestamp, user, JSON.stringify(jsonData)];
    versionSheet.getRange("A1:D1").values = [["Version", "Timestamp", "User", "Data"]];
    versionSheet.getRange(`A${existing.length + 2}:D${existing.length + 2}`).values = [newRow];
    await context.sync();

    console.log(`Version ${newVersion} saved.`);
    renderVersionHistory();
  });
}

function getRelativeTime(isoString) {
  const diff = Date.now() - new Date(isoString).getTime();
  if (diff < 60000) return "Just now";
  if (diff < 3600000) return `${Math.floor(diff / 60000)} minutes ago`;
  if (diff < 86400000) return `${Math.floor(diff / 3600000)} hours ago`;
  return new Date(isoString).toLocaleString();
}

async function renderVersionHistory() {
  await Excel.run(async (context) => {
    const container = document.getElementById("versionList");
    container.innerHTML = "Loading...";

    try {
      const sheet = context.workbook.worksheets.getItem("VersionHistory");
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();

      const values = range.values.slice(1);
      if (values.length === 0) {
        container.innerHTML = "No versions found.";
        return;
      }

      container.innerHTML = "";
      [...values].reverse().forEach((row, idx) => {
        const [version, timestamp, user, _data] = row;
        const timeLabel = idx === 0 ? "Current Version" : getRelativeTime(timestamp);

        const div = document.createElement("div");
        div.className = "version-entry";
        div.onclick = () => loadVersionByVersion(version);
        div.innerHTML = `
          <div class="version-title">${getRelativeTime(timestamp)}</div>
          <div class="version-sub">${timeLabel}</div>
          <div class="user-info"><span class="user-bullet"></span>${user}</div>
        `;
        container.appendChild(div);
      });
    } catch (e) {
      container.innerHTML = "No version history found.";
    }
  });
}

async function loadVersionByVersion(versionToLoad) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("VersionHistory");
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const values = range.values;
    const match = values.find(row => row[0] === versionToLoad);
    if (!match) return console.log("Version not found");

    const json = JSON.parse(match[3]);
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const used = activeSheet.getUsedRangeOrNullObject();
    used.load("address");
    await context.sync();

    if (!used.isNullObject) used.clear();

    if (!json || json.length === 0) {
      activeSheet.getRange("A1").values = [[""]];
      await context.sync();
      return;
    }

    const headers = Object.keys(json[0]);
    const data = [headers, ...json.map(obj => headers.map(h => obj[h]))];
    const rangeToWrite = activeSheet.getRangeByIndexes(0, 0, data.length, headers.length);
    rangeToWrite.values = data;
    await context.sync();
  });
}
