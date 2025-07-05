// Wait for Office.js to be ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready");

    // Attach click handler (optional if in HTML)
    const btn = document.getElementById("saveBtn");
    if (btn) {
      btn.addEventListener("click", saveVersion);
    }
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
  return current; // return the one used for this save
}

/**
 * Save the current Excel document as JSON revision
 */
async function saveVersion() {
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
                slices.push(sliceResult.value.data);
                sliceIndex++;
                if (sliceIndex < sliceCount) {
                  getSlice();
                } else {
                  file.closeAsync();
                  const blob = new Blob(slices);
                  saveAsJSON(blob);
                }
              } else {
                console.error("Failed to get slice", sliceResult.error.message);
              }
            });
          };

          getSlice();
        } else {
          console.error("Failed to get file", result.error.message);
        }
      });
    });
  } catch (err) {
    console.error("Excel run failed", err);
  }
}

/**
 * Save the file data and metadata as a local .json file
 */
function saveAsJSON(blob) {
  const reader = new FileReader();

  reader.onload = () => {
    const base64Data = reader.result.split(',')[1];
    const version = getNextVersion();

    const revision = {
      version,
      filename: `excel-version-v${version}.xlsx`,
      user: "unknown", // You can prompt or fetch via SSO if needed
      timestamp: new Date().toISOString(),
      comment: "COMMENT", // Replace with input if you want
      fileData: base64Data,
    };

    const jsonBlob = new Blob([JSON.stringify(revision, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(jsonBlob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `excel-revision-v${version}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    alert(`Revision v${version} saved as JSON.`);
  };

  reader.readAsDataURL(blob);
}
