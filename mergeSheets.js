const ExcelJS = require('exceljs');

// Rest of the code...


async function mergeSheets() {
    const excelFileInput = document.getElementById('excelFileInput');
    const file = excelFileInput.files[0];
  
    if (!file) {
      alert('Please select an Excel file');
      return;
    }
  
    // Load the workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
  
    // Create a new workbook for the merged sheets
    const mergedWorkbook = new ExcelJS.Workbook();
    const mergedSheet = mergedWorkbook.addWorksheet('Merged Sheet');
  
    // Loop through each sheet in the original workbook
    workbook.eachSheet((worksheet, sheetId) => {
      // Copy the sheet data to the merged sheet
      worksheet.eachRow((row, rowNumber) => {
        const rowData = row.values;
        mergedSheet.addRow(rowData);
      });
    });
  
    // Save the merged workbook to a new file
    const outputFilePath = 'merged.xlsx';
    await mergedWorkbook.xlsx.writeFile(outputFilePath);
    console.log(`Merged sheets saved to ${outputFilePath}`);
  
    // Provide download link to the user
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(new Blob([await mergedWorkbook.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    downloadLink.download = 'merged.xlsx';
    downloadLink.click();
  }
  