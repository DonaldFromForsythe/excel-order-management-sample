
function main(workbook: ExcelScript.Workbook) {
   //Select the active sheet
  const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

/*On the workbook there are 4 mains Sheets for each quarter to prevent overloading
Select the table that starts with "Q" (for Quarter) in the active sheet*/
  
  const table: ExcelScript.Table = sheet.getTables().find(t => t.getName().startsWith("Q"));
  if (!table) {
    console.log("Table not found.");
    return;
  }
   const dataRange: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
   const values: (string | number | boolean)[][] = dataRange.getValues();
   const headers: (string | number | boolean)[] = table.getHeaderRowRange().getValues()[0];
   const nameIndex: number = headers.indexOf("Name");
   const authCodeIndex: number = headers.indexOf("AUTH CODE");
   const orderTimeIndex: number = headers.indexOf("Order Time");
   
   //If one if one of these are not found stop the script and log the error
   if ([nameIndex, authCodeIndex, orderTimeIndex].includes(-1)) {
    console.log("One or more required columns are missing.");
    return;
  }

   
}
