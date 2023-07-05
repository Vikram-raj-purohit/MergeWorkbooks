

document.addEventListener('DOMContentLoaded', function() {
    const mergeBtn = document.getElementById('mergeBtn');
    mergeBtn.addEventListener('click', mergeExcelFiles);
  });
  
  function mergeExcelFiles() {
    const input = document.getElementById('excelFiles');
    const files = input.files;
  
    if (files.length < 2) {
      console.log('Please select at least two Excel files.');
      return;
    }
  
    const workbooks = [];
  
    // Read each file and store the workbook
    for (let i = 0; i < files.length; i++) {
      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        workbooks.push(workbook);
        
        if (workbooks.length === files.length) {
          // Perform the Excel merging
          const mergedWorkbook = mergeWorkbooks(workbooks);
  
          // Save the merged workbook
          const mergedWorkbookData = XLSX.write(mergedWorkbook, { bookType: 'xlsx', type: 'array' });
          saveFile(mergedWorkbookData, 'merged_file.xlsx');
        }
      };
      reader.readAsArrayBuffer(files[i]);
    }
  }
  
  function mergeWorkbooks(workbooks) {
    // Create an empty workbook
    const mergedWorkbook = XLSX.utils.book_new();
  
    // Counter for duplicate worksheet names
    let counter = 1;
  
    // Iterate over each workbook
    for (let i = 0; i < workbooks.length; i++) {
      const workbook = workbooks[i];
  
      // Iterate over each sheet in the workbook
      workbook.SheetNames.forEach(function(sheetName) {
        const worksheet = workbook.Sheets[sheetName];
  
        // Generate a unique name for the worksheet
        let uniqueSheetName = sheetName;
        while (mergedWorkbook.Sheets[uniqueSheetName] !== undefined) {
          uniqueSheetName = `${sheetName} (${counter})`;
          counter++;
        }
  
        // Convert the worksheet to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  
        // Create a new worksheet in the merged workbook
        const mergedWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
  
        // Add the worksheet to the merged workbook with the unique name
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedWorksheet, uniqueSheetName);
      });
    }
  
    return mergedWorkbook;
  }
  
  function saveFile(data, filename) {
    const blob = new Blob([data], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
  
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }