/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
const ADDON_TITLE = 'Form Notifications123';


/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
const onOpen = (e) => {
  FormApp.getUi()
    .createMenu(ADDON_TITLE)
    .addItem("Open", "showSidebar")
    .addToUi();
};

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
const onInstall = (e) => {
  onOpen(e);
};

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
const showSidebar = () => {
  const ui = HtmlService.createTemplateFromFile("src/sidebar")
    .evaluate()
    .setTitle(ADDON_TITLE);
  FormApp.getUi().showSidebar(ui);
};



/**
 * Parses the file with the given file name. This is used to inject CSS and Javascript
 * in an HTML file.
 */
// const include = (filename: string) =>
//   HtmlService.createHtmlOutputFromFile(filename).getContent();

/**
 *
 * @param filename
 * @returns
 */
const loadSettingsUI = () => {
  const htmlSev = HtmlService.createTemplateFromFile("src/settings");
  const html = htmlSev.evaluate();
  html.setWidth(650).setHeight(400);
  const ui = FormApp.getUi();
  ui.showModalDialog(html, "Settings Page");
};