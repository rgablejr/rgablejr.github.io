// This tells Office what to do when the library is ready
Office.onReady((info) => {
if (info.host === Office.HostType.Excel) {
console.log("Office.js is ready in Excel.");
}
});

/**
* The function name must match the <FunctionName> in the manifest.xml
*/
async function myRibbonAction(event) {
try {
await Excel.run(async (context) => {
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1");

range.values = [["Action Triggered!"]];
range.format.font.bold = true;
range.format.fill.color = "yellow";

await context.sync();
});
} catch (error) {
console.error("Error: " + error);
} finally {
// MANDATORY: Signals to Excel that the function is finished.
// If you omit this, the button will stay grayed out or show a spinner.
event.completed();
}
}

// Map the function name in the manifest to the JavaScript function
Office.actions.associate("myRibbonAction", myRibbonAction);