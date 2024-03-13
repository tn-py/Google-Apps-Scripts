// Code.gs
function exportSheetsToDrive() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    if (sheets === undefined || sheets.length === 0) {
      return;
    }
    const now = new Date();
    
    const csvBlobs = sheets.map((sheet) => {
      const name = sheet.getName();
      const csv = convertSheetToCsv(sheet);
      Logger.log({ name, length: csv.length });
      return Utilities.newBlob(csv, MimeType.CSV, `${name}.csv`)
    });
  
    const zipName = `export_${ss.getName()}_${now.toISOString()}.zip`;
    const zip = Utilities.zip(csvBlobs, zipName);
    DriveApp.createFile(zip);
  }
  
  function convertSheetToCsv(sheet) {
    return sheet
      .getDataRange()
      .getValues()
      .map((row) =>
        row
          .map((value) => value.toString())
          .map((value) =>
            value.includes("\n") || value.includes(",")
              ? '"' + value + '"'
              : value
          )
          .join(",")
      )
      .join("\n");
  }