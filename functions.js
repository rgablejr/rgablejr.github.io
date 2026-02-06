Office.onReady(() => {
    console.log("Office initialized");
});

/**
 * Your main ribbon function
 */
async function helloWorld(event) {
    try {
        // 1. Run your logic
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [["Command Received!"]];
            await context.sync();
        });

        // 2. Log success
        await writeLog("Success: helloWorld triggered.");

    } catch (error) {
        // 3. Log errors so you can actually see them
        await writeLog("Error: " + error.message);
    } finally {
        // 4. ALWAYS tell Office you are done
        event.completed();
    }
}

/**
 * Helper function to log actions to a sheet named 'DebugLog'
 */
async function writeLog(message) {
    await Excel.run(async (context) => {
        let sheets = context.workbook.worksheets;
        let logSheet = sheets.getItemOrNullObject("DebugLog");
        
        await context.sync();

        if (logSheet.isNullObject) {
            logSheet = sheets.add("DebugLog");
        }

        const lastRow = logSheet.getUsedRange(true).getLastRowOrNullObject();
        await context.sync();

        // Write to the next available row (or row 1 if empty)
        const targetRow = lastRow.isNullObject ? 0 : lastRow.rowIndex + 1;
        const range = logSheet.getRangeByIndexes(targetRow, 0, 1, 2);
        
        range.values = [[new Date().toLocaleTimeString(), message]];
        logSheet.activate();
        
        await context.sync();
    });
}

// Map the function name in the manifest to the JS function
Office.actions.associate("helloWorld", helloWorld);