/** taskpane.js - Uses CustomProperties for version storage (Web + Desktop compatible) */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");

    document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    document.getElementById("loadVersionBtn")?.addEventListener("click", handleVersionLoad);

    populateVersionDropdown();
  }
});

function getNextVersion(versions) {
  if (!versions.length) return "1.0.0";
  const last = versions[versions.length - 1].version;
  let [major, minor, patch] = last.split(".").map(Number);
  patch++;
  return `${major}.${minor}.${patch}`;
}

async function getStoredVersions() {
  await Excel.run(async (context) => {
    const props = context.workbook.properties.custom;
    props.load("items");
    await context.sync();

    const item = props.items.find(p => p.key === "ECCP_VERSIONS");
    window._versionData = item ? JSON.parse(item.value) : [];
  });
  return window._versionData || [];
}

async function saveVersionData(versions) {
  await Excel.run(async (context) => {
    const props = context.workbook.properties.custom;
    props.add("ECCP_VERSIONS", JSON.stringify(versions));
    await context.sync();
  });
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

    const versions = await getStoredVersions();
    const newVersion = getNextVersion(versions);

    const newEntry = {
      version: newVersion,
      timestamp: new Date().toISOString(),
      user: Office.context?.userProfile?.displayName || "unknown",
      data: jsonData,
    };

    const updatedVersions = [...versions, newEntry];
    await saveVersionData(updatedVersions);
    console.log(`Version ${newVersion} saved.`);
    await populateVersionDropdown(newVersion);
  });
}

async function populateVersionDropdown(selectVersion = null) {
  const dropdown = document.getElementById("versionDropdown");
  dropdown.innerHTML = '<option value="">Select Version</option>';

  const versions = await getStoredVersions();

  versions.forEach((v, i) => {
    const opt = document.createElement("option");
    opt.value = i;
    opt.textContent = `${v.version} (${v.timestamp.split("T")[0]})`;
    if (v.version === selectVersion) opt.selected = true;
    dropdown.appendChild(opt);
  });
}

function handleVersionLoad() {
  const dropdown = document.getElementById("versionDropdown");
  const idx = dropdown.value;
  if (idx === "") {
    console.log("Please select a version.");
    return;
  }
  loadSelectedVersion(parseInt(idx));
}

async function loadSelectedVersion(index) {
  const versions = await getStoredVersions();
  const version = versions[index];
  if (!version) return;

  await Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = activeSheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();

    if (!usedRange.isNullObject) usedRange.clear();

    const json = version.data;
    if (!json || json.length === 0) {
      activeSheet.getRange("A1").values = [[""]];
      await context.sync();
      console.log("Blank version loaded.");
      return;
    }

    const headers = Object.keys(json[0]);
    const rows = [headers, ...json.map(obj => headers.map(h => obj[h]))];
    const writeRange = activeSheet.getRangeByIndexes(0, 0, rows.length, headers.length);
    writeRange.values = rows;
    await context.sync();
    console.log(`Version ${version.version} loaded.`);
  });
}
