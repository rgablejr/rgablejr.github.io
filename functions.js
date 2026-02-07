// Initialize the Office Add-in.
Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

// The command function.
async function highlightSelection(event) {

    // Implement your custom code here. The following code is a simple Excel example.
    try {
          await Excel.run(async (context) => {
              const range = context.workbook.getSelectedRange();
			  const colors = ['red', 'green', 'blue', 'orange', 'yellow', 'purple'];
			  const randomIndex = Math.floor(Math.random() * colors.length);
			  const randomColor = colors[randomIndex];
			  
              range.format.fill.color = randomColor;
              await context.sync();
          });
      } catch (error) {
          // Note: In a production add-in, notify the user through your add-in's UI.
          console.error(error);
      }

    // Calling event.completed is required. The event.completed call lets the platform know that processing has completed.
    event.completed();
}

// This maps the function to the action ID specified in the manifest.
Office.actions.associate("highlightSelection", highlightSelection);