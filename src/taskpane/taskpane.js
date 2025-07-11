// A placeholder for your critical zones configuration.
// It is now specific to your sheet "Book 5" and only defines a range.
const criticalZones = {
  "Book 5": [
    { type: "range", address: "A1:C5" }
  ]
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    document.getElementById("viewMetadataBtn")?.addEventListener("click", showMetadataSheet);
    renderVersionHistory();
    // Start monitoring for changes once the add-in is ready
    startChangeMonitoring();
  }
});

let currentVersion = null;

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
    } else if (headers.length > 0 && dataRows.length === 0) {
      storedData = { headers, data: [] };
    } else if (headers.length && dataRows.length) {
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

    currentVersion = newVersion;
    renderVersionHistory();
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
    const used = activeSheet.getUsedRangeOrNullObject();
    used.load("address");
    await context.sync();
    if (!used.isNullObject) used.clear();

    if (Array.isArray(parsed) && parsed.length === 0) {
      // Blank sheet
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

// --- New Change Management Functionality ---

/**
 * Initializes the change monitoring system by adding an onChanged event listener
 * to the active worksheet.
 */
async function startChangeMonitoring() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(handleWorksheetChange);
    await context.sync();
    console.log("Change monitoring started on active sheet.");
  });
}

/**
 * Event handler for worksheet changes. It checks if the change is in a critical zone.
 * @param {Excel.WorksheetChangedEvent} event
 */
async function handleWorksheetChange(event) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.get(event.worksheetId);
    sheet.load("name");
    await context.sync();

    // Check if the change occurred in a critical zone on the current sheet
    const isCritical = await isChangeInCriticalZone(context, sheet, event.address);

    if (isCritical) {
      console.log(`Critical change detected in ${event.address}`);
      promptAndLogChange(event.worksheetId, event.address);
    }
  });
}

/**
 * Checks if a given range address intersects with any of the defined critical zones.
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {string} changedAddress The address of the changed range.
 * @returns {Promise<boolean>} True if the change is within a critical zone.
 */
async function isChangeInCriticalZone(context, sheet, changedAddress) {
  const zones = criticalZones[sheet.name] || [];
  for (const zone of zones) {
    if (zone.type === "range") {
      const criticalRange = sheet.getRange(zone.address);
      criticalRange.load("address");
      const intersection = criticalRange.getIntersectionOrNullObject(changedAddress);
      intersection.load("isNullObject");
      await context.sync();
      if (!intersection.isNullObject) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Prompts the user for a description and logs the change to the ChangeLog sheet.
 * @param {string} worksheetId The ID of the worksheet where the change occurred.
 * @param {string} changedAddress The address of the edited range.
 */
function promptAndLogChange(worksheetId, changedAddress, dialogUrl) {
  // Office.context.ui.displayDialogAsync can only be called from an HTTPS page.
  // Make sure your add-in is served over HTTPS for this to work.
  const url = dialogUrl || `${window.location.origin}/dialog.html?address=${encodeURIComponent(changedAddress)}`;

  Office.context.ui.displayDialogAsync(url, {
    height: 25,
    width: 35
  }, asyncResult => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
        const message = JSON.parse(event.message);
        if (message.action === "logChange") {
          dialog.close();
          writeChangeLog(worksheetId, changedAddress, message.description);
        }
      });
    } else {
      console.error("Failed to open dialog: " + asyncResult.error.message);
    }
  });
}

/**
 * Writes the details of a critical change to a hidden ChangeLog sheet.
 * @param {string} worksheetId The ID of the worksheet.
 * @param {string} changedAddress The address of the changed range.
 * @param {string} description The user-provided change description.
 */
async function writeChangeLog(worksheetId, changedAddress, description) {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
    await context.sync();

    if (sheet.isNullObject) {
      sheet = context.workbook.worksheets.add("ChangeLog");
      sheet.visibility = Excel.SheetVisibility.hidden;
      sheet.getRange("A1:E1").values = [["Timestamp", "User", "Sheet", "Address", "Description"]];
    }

    // Load the sheet name based on its ID
    const changedSheet = context.workbook.worksheets.get(worksheetId);
    changedSheet.load("name");
    await context.sync();

    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount");
    await context.sync();

    const nextRow = usedRange.isNullObject ? 2 : usedRange.rowCount + 1;
    const newRowData = [
      new Date().toISOString(),
      "User One", // Placeholder, you can get the actual user from a login service
      changedSheet.name,
      changedAddress,
      description
    ];
    sheet.getRange(`A${nextRow}:E${nextRow}`).values = [newRowData];
    await context.sync();
    console.log("Change logged successfully.");
  });
}