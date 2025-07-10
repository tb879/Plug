Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    document.getElementById("viewMetadataBtn")?.addEventListener("click", showMetadataSheet);
    renderVersionHistory();

    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(handleChangeEvent);
      await context.sync();
    });
  }
});

let currentVersion = null;
const criticalZones = [
  { sheet: "Sheet1", range: "A1:C10" }
];

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
    const range = sheet.getUsedRangeOrNullObject();
    range.load("values");
    await context.sync();

    const values = range.isNullObject ? [] : range.values;
    const headers = values[0] || [];
    const dataRows = values.length > 1 ? values.slice(1) : [];

    let storedData = [];

    if (headers.length === 0 && dataRows.length === 0) {
      storedData = []; // Completely blank
    } else {
      storedData = { headers, data: dataRows };
    }

    let versionSheet = context.workbook.worksheets.getItemOrNullObject("VersionHistory");
    await context.sync();

    if (versionSheet.isNullObject) {
      versionSheet = context.workbook.worksheets.add("VersionHistory");
      versionSheet.visibility = Excel.SheetVisibility.hidden;
    }

    const used = versionSheet.getUsedRangeOrNullObject();
    used.load("values");
    await context.sync();

    const existing = used.isNullObject ? [] : used.values.slice(1);
    const newVersion = getNextVersion(existing);
    const timestamp = new Date().toISOString();
    const user = "User One";

    const newRow = [newVersion, timestamp, user, JSON.stringify(storedData)];
    versionSheet.getRange("A1:D1").values = [["Version", "Timestamp", "User", "Data"]];
    versionSheet.getRange(`A${existing.length + 2}:D${existing.length + 2}`).values = [newRow];
    await context.sync();

    await writeMetadataSheet(context, newVersion, user);
    await sheet.protection.protect(); // Lock sheet

    currentVersion = newVersion;
    renderVersionHistory();
  });
}

async function handleChangeEvent(eventArgs) {
  const inZone = criticalZones.some(zone => {
    return zone.sheet === eventArgs.worksheetId && eventArgs.address.startsWith(zone.range.split(":")[0]);
  });

  if (inZone) {
    const reason = await promptForChangeReason();
    await logChange(eventArgs, reason);
  }
}

function promptForChangeReason() {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync("https://tb879.github.io/Plug/index.html?dialog=true", { height: 40, width: 30 }, (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        resolve(arg.message);
      });
    });
  });
}

async function logChange(eventArgs, reason) {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
    await context.sync();
    if (sheet.isNullObject) {
      sheet = context.workbook.worksheets.add("ChangeLog");
      sheet.visibility = Excel.SheetVisibility.hidden;
      sheet.getRange("A1:F1").values = [["Timestamp", "Sheet", "Address", "Type", "Details", "Reason"]];
    }

    const time = new Date().toISOString();
    const logRow = [time, eventArgs.worksheetId, eventArgs.address, eventArgs.changeType, JSON.stringify(eventArgs), reason];

    const used = sheet.getUsedRangeOrNullObject();
    used.load("values");
    await context.sync();

    const rowNum = used.isNullObject ? 2 : used.values.length + 1;
    sheet.getRange(`A${rowNum}:F${rowNum}`).values = [logRow];
    await context.sync();
  });
}

async function loadVersionByVersion(versionToLoad) {
  await Excel.run(async (context) => {
    const versionSheet = context.workbook.worksheets.getItem("VersionHistory");
    const range = versionSheet.getUsedRange();
    range.load("values");
    await context.sync();

    const match = range.values.find(row => row[0] === versionToLoad);
    if (!match) return console.warn("Version not found");

    const parsed = JSON.parse(match[3]);
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();

    // ✅ Unprotect before writing
    await activeSheet.protection.unprotect();
    await context.sync();

    const used = activeSheet.getUsedRangeOrNullObject();
    used.load("address");
    await context.sync();
    if (!used.isNullObject) used.clear();

    if (Array.isArray(parsed) && parsed.length === 0) {
      activeSheet.getRange("A1").values = [[""]];
    } else if (parsed.headers && Array.isArray(parsed.headers)) {
      const rows = [parsed.headers, ...(parsed.data || [])];
      const rangeToWrite = activeSheet.getRangeByIndexes(0, 0, rows.length, parsed.headers.length);
      rangeToWrite.values = rows;
    }

    await context.sync();
    currentVersion = versionToLoad;
    renderVersionHistory();
  });
}

async function renderVersionHistory() {
  await Excel.run(async (context) => {
    const container = document.getElementById("versionList");
    container.innerHTML = "Loading...";
    const sheet = context.workbook.worksheets.getItemOrNullObject("VersionHistory");
    await context.sync();

    if (sheet.isNullObject) {
      container.innerHTML = "No version history found.";
      return;
    }

    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const versions = range.values.slice(1);
    if (!versions.length) {
      container.innerHTML = "No versions found.";
      return;
    }

    currentVersion ||= versions[versions.length - 1][0];
    container.innerHTML = "";

    [...versions].reverse().forEach(([ver, time, user]) => {
      const isCurrent = ver === currentVersion;
      const div = document.createElement("div");
      div.className = "version-entry";
      div.onclick = () => loadVersionByVersion(ver);
      div.innerHTML = `
        <div class="version-title">${getRelativeTime(time)}</div>
        <div class="version-sub">Version: ${ver}${isCurrent ? " (current)" : ""}</div>
        <div class="user-info"><span class="user-bullet"></span>${user}</div>
      `;
      container.appendChild(div);
    });
  });
}

function getRelativeTime(iso) {
  const diff = Date.now() - new Date(iso).getTime();
  if (diff < 60000) return "Just now";
  if (diff < 3600000) return `${Math.floor(diff / 60000)} min ago`;
  if (diff < 86400000) return `${Math.floor(diff / 3600000)} hours ago`;
  return new Date(iso).toLocaleString();
}

async function writeMetadataSheet(context, version, user) {
  const metaSheet = context.workbook.worksheets.getItemOrNullObject("Metadata");
  metaSheet.load("isNullObject");
  await context.sync();

  let sheet;
  if (metaSheet.isNullObject) {
    sheet = context.workbook.worksheets.add("Metadata");
    sheet.visibility = Excel.SheetVisibility.hidden;
  } else {
    sheet = metaSheet;
    const used = sheet.getUsedRangeOrNullObject();
    used.load("address");
    await context.sync();
    if (!used.isNullObject) used.clear();
  }

  const today = new Date().toISOString().split("T")[0];
  const docId = `DOC-${today.replace(/-/g, "")}-001`;
  const meta = [
    ["Document Title", "The Doc"],
    ["Document ID", docId],
    ["Revision Number", version],
    ["Date of Issue", today],
    ["Owner/Author", user],
    ["Approver(s)", "John Smith"],
    ["Department/Team", "Quality"],
    ["Standard", "ISO 9001"]
  ];

  const range = sheet.getRange(`A1:B${meta.length}`);
  range.values = meta;
  await context.sync();
  sheet.getRange("B1:B" + meta.length).format.autofitColumns();
  await context.sync();
}

async function showMetadataSheet() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("Metadata");
    sheet.load("isNullObject");
    await context.sync();

    if (sheet.isNullObject) return;
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.activate();
    await context.sync();
  });
}
