/* global Office */

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  console.log("LLM add-in for Excel action triggered.");
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
