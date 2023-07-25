async function mergeSheets() {
    const excelFileInput = document.getElementById('excelFileInput');
    const file = excelFileInput.files[0];
  
    if (!file) {
      alert('Please select an Excel file');
      return;
    }
  
    const reader = new FileReader();
  
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
  
      const mergedWorkbook = XLSX.utils.book_new();
      const mergedSheet = XLSX.utils.aoa_to_sheet([]);
  
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        XLSX.utils.sheet_add_json(mergedSheet, sheetData, { skipHeader: true, origin: -1 });
      });
  
      XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, 'Merged Sheet');
  
      const outputFilePath = 'merged.xlsx';
      XLSX.writeFile(mergedWorkbook, outputFilePath);
  
      console.log(`Merged sheets saved to ${outputFilePath}`);
  
      const downloadLink = document.createElement('a');
      downloadLink.href = URL.createObjectURL(new Blob([s2ab(XLSX.write(mergedWorkbook, { bookType: 'xlsx', type: 'binary' }))], { type: 'application/octet-stream' }));
      downloadLink.download = 'merged.xlsx';
      downloadLink.click();
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
  

  
