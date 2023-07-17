

async function mergeSheets() {
  const excelFileInput = document.getElementById("excelFileInput");
  const file = excelFileInput.files[0];

  if (!file) {
    alert("Please select an Excel file");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const mergedWorkbook = XLSX.utils.book_new();
    const mergedSheet = XLSX.utils.aoa_to_sheet([]);

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      XLSX.utils.sheet_add_json(mergedSheet, sheetData, {
        skipHeader: true,
        origin: -1,
      });

      processedSheets++;

      // Update progress
      const progress = Math.round((processedSheets / totalSheets) * 100);
      updateProgress(progress);
    });

    XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, "Merged Sheet");

    const outputFilePath = "merged.xlsx";

    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(
      new Blob(
        [
          s2ab(
            XLSX.write(mergedWorkbook, { bookType: "xlsx", type: "binary" })
          ),
        ],
        { type: "application/octet-stream" }
      )
    );
    downloadLink.download = outputFilePath;
    downloadLink.click();
  

  updateProgress(0);
}

reader.onerror = function (e) {
  alert("Error occurred while reading the file");
  // Reset progress on error
  updateProgress(0);
};
reader.onprogress = function (e) {
  if (e.lengthComputable) {
    const progress = Math.round((e.loaded / e.total) * 100);
    updateProgress(progress);
  }
};

reader.readAsArrayBuffer(file);
}
// Utility function to convert string to ArrayBuffer
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
  return buf;
}
function updateProgress(progress) {
  const progressBar = document.getElementById("progressBar");
  progressBar.style.width = progress + "%";
  progressBar.textContent = progress + "%";
}
