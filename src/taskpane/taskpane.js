Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("saveCommitBtn")?.addEventListener("click", saveAndCommitVersion);
    // document.getElementById("viewMetadataBtn")?.addEventListener("click", showMetadataSheet);
    document.getElementById("submitChangeReason")?.addEventListener("click", submitChangeReason);
    document.getElementById("cancelChangeReason")?.addEventListener("click", cancelChangeReason);
    renderVersionHistory();

    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(onCellChanged);
      await context.sync();
      console.log("✅ onChanged event attached to active sheet");
    });
  }
});

const CRITICAL_ZONES = [
  { address: "A1:C10", name: "Header Table" },
  { address: "E5:E15", name: "Finance Column" }
];

let currentVersion = null;
let pendingLogData = null;

function showChangePrompt(zoneName, address) {
  document.getElementById("promptMessage").textContent = `You changed a CRITICAL zone: "${zoneName}" (${address}). Please describe your change.`;
  document.getElementById("changeReason").value = "";
  document.getElementById("changePrompt").style.display = "block";
  pendingLogData = { zoneName, address };
}

function submitChangeReason() {
  const reason = document.getElementById("changeReason").value.trim();
  document.getElementById("changePrompt").style.display = "none";
  if (reason && pendingLogData) {
    logEditChange(pendingLogData.zoneName, pendingLogData.address, reason);
  }
  pendingLogData = null;
}

function cancelChangeReason() {
  document.getElementById("changePrompt").style.display = "none";
  pendingLogData = null;
}

async function onCellChanged(event) {
  console.log("🛠️ Cell changed:", event.address);
  const address = event.address.replace(/^.*?!/, "");

  const zone = CRITICAL_ZONES.find(zone => isRangeIntersecting(address, zone.address));
  if (zone) {
    showChangePrompt(zone.name, address);
  }
}

function isRangeIntersecting(edited, critical) {
  const [aStart, aEnd] = getBounds(edited);
  const [bStart, bEnd] = getBounds(critical);

  return (
    aStart.row <= bEnd.row &&
    aEnd.row >= bStart.row &&
    aStart.col <= bEnd.col &&
    aEnd.col >= bStart.col
  );
}

function getBounds(address) {
  const match = address.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
  if (!match) return [{ row: 0, col: 0 }, { row: 0, col: 0 }];
  const startCol = colToNum(match[1]), startRow = parseInt(match[2]);
  const endCol = colToNum(match[3] || match[1]), endRow = parseInt(match[4] || match[2]);
  return [
    { row: startRow, col: startCol },
    { row: endRow, col: endCol }
  ];
}

function colToNum(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

async function logEditChange(zoneName, cellAddress, reason) {
  await Excel.run(async (context) => {
    let logSheet = context.workbook.worksheets.getItemOrNullObject("ChangeLog");
    await context.sync();

    if (logSheet.isNullObject) {
      logSheet = context.workbook.worksheets.add("ChangeLog");
      logSheet.getRange("A1:E1").values = [["Timestamp", "User", "Zone", "Cell", "Change Description"]];
    }

    const now = new Date().toISOString();
    const user = "User One";

    const used = logSheet.getUsedRangeOrNullObject();
    used.load("rowCount");
    await context.sync();

    const row = used.isNullObject ? 2 : used.rowCount + 1;
    logSheet.getRange(`A${row}:E${row}`).values = [[now, user, zoneName, cellAddress, reason]];
    await context.sync();
  });
}

// Versioning and Metadata functions...

// (keep your saveAndCommitVersion, loadVersionByVersion, renderVersionHistory, getNextVersion,
//  writeMetadataSheet, getRelativeTime, and showMetadataSheet functions here, unchanged)
// -- for brevity, I’ve omitted them but they remain exactly as before.
