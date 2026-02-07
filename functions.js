// Initialize Office
Office.onReady(() => {
    // Office is ready
});

function helloWorld(event) {
    Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [["Hello from the Ribbon!"]];
        range.format.fill.color = "yellow";
        
        await context.sync();
    }).catch((error) => {
        console.error(error);
    }).finally(() => {
        // IMPORTANT: You must signal that the function is done
        event.completed();
    });
}

// Map the function name in the manifest to the JS function
Office.actions.associate("helloWorld", helloWorld);