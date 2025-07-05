// Wait for Office.js to be ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");
    // Optional: Bind to a button
    document.getElementById("saveJsonBtn")?.addEventListener("click", saveVersionAsJSON);
  }
});

/**
 * Get the next revision version like 1.0.0 → 1.0.1
 */
function getNextVersion() {
  const versionKey = "excel-revision-version";
  let current = localStorage.getItem(versionKey) || "1.0.0";

  let [major, minor, patch] = current.split(".").map(Number);
  patch += 1;

  const next = `${major}.${minor}.${patch}`;
  localStorage.setItem(versionKey, next);
  return current; // return current for this version
}

/**
 * Save the Excel content (values) as a structured JSON revision
 */
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
