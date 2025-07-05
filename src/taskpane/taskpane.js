const saveVersion = async () => {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Get the workbook as a file
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
};

const saveAsJSON = (blob) => {
  const reader = new FileReader();

  reader.onload = () => {
    const base64Data = reader.result.split(',')[1]; // Remove prefix
    const revision = {
      filename: `excel-version-${new Date().toISOString()}.xlsx`,
      user: Office.context?.userProfile?.displayName || "unknown",
      timestamp: new Date().toISOString(),
      comment: prompt("Enter a comment for this revision:") || "",
      fileData: base64Data,
    };

    const jsonBlob = new Blob([JSON.stringify(revision, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(jsonBlob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `excel-revision-${new Date().toISOString()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    alert("Revision saved as JSON.");
  };

  reader.readAsDataURL(blob);
};
