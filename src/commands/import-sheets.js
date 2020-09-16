export async function getAllWorksheetData() {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem("badger-company49759_accounts_15");
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    const range = sheet.getRange(`A2:A${largeRange.rowCount}`);
    range.load("values");
    await context.sync();

    console.log("range", JSON.stringify(range.values, null, 4));
    //console.log("Fuse", Fuse);
    //const sheetCopy = context.workbook.worksheets.add("badger-company49759-copy");

    //queueCommandsToCreateTemperatureTable(sheet);
    //sheet.activate();

    await context.sync();
    //console.log("range", JSON.stringify(range.text, null, 4));
  });
}
