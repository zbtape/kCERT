// Wait for Office.js to load
Office.onReady(() => {
    // Commands are registered in the manifest
});

/**
 * Shows the task pane
 * @param event
 */
function showTaskPane(event: Office.AddinCommands.Event): void {
    // Task pane will be shown automatically by Office
    event.completed();
}

// Register functions for add-in commands
(global as any).showTaskPane = showTaskPane; 