function setupTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('assignUniqueBelCodes') // Calls your main function
           .forSpreadsheet(sheet)
           .onFormSubmit()  // Trigger happens when a form is submitted
           .create();
}
