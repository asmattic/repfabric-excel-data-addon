export const getWorksheetNames = async () => {
  return Excel.run(async context => {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    if (sheets.items.length > 1) {
      console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
    } else {
      console.log(`There is one worksheet in the workbook:`);
    }
    const sheetNames = [];
    sheets.items.forEach(sheet => {
      console.log(sheet.name);
      sheetNames.push(sheet.name);
    });
    return sheetNames;
  });
};
