/* v1.0.2-v1.0.5 registered a duplicate CSS-bearing message script. Remove
 * that persisted registration during upgrade; the manifest's guarded JS now
 * owns adaptive message-canvas styling without attempting CSS injection into
 * the native chrome:// thread-summary document. */
messenger.scripting.messageDisplay
  .getRegisteredScripts()
  .then(registered => {
    if (
      registered.some(script => script.id === "fluent-mail-light-message-view")
    ) {
      return messenger.scripting.messageDisplay.unregisterScripts({
        ids: ["fluent-mail-light-message-view"],
      });
    }
    return undefined;
  })
  .catch(error => {
    console.error(
      "Outlook Style Companion could not remove its legacy display script:",
      error
    );
  });

/* Build the direct Calendar / To Do controls and Outlook-style agenda details
 * in Thunderbird's native Today pane. The startup listener guarantees that
 * Thunderbird wakes the MV3 event page after an application restart; the
 * immediate call covers install/update in the current session. */
function installTodayPane() {
  return messenger.outlookThreadView.installTodayPane().catch(error => {
    console.error(
      "Outlook Style Companion could not enhance the Today pane:",
      error
    );
  });
}

function installReminderDialog() {
  return messenger.outlookThreadView.installReminderDialog().catch(error => {
    console.error(
      "Outlook Style Companion could not enhance the reminder dialog:",
      error
    );
  });
}

function installEditorSurfaces() {
  return messenger.outlookThreadView.installEditorSurfaces().catch(error => {
    console.error(
      "Outlook Style Companion could not style the compose and event editors:",
      error
    );
  });
}

function installCalendarChooser() {
  return messenger.outlookThreadView.installCalendarChooser().catch(error => {
    console.error(
      "Outlook Style Companion could not enhance the calendar chooser:",
      error
    );
  });
}

function installOutlookEnhancements() {
  installTodayPane();
  installReminderDialog();
  installEditorSurfaces();
  installCalendarChooser();
}

messenger.runtime.onStartup.addListener(installOutlookEnhancements);
installOutlookEnhancements();

/* Thunderbird already summarizes collapsed threads. For an expanded thread,
 * ask the companion API for its native conversation accordion so every compact
 * message can open its complete trusted reader in place. The API revalidates
 * the current selection and leaves child and standalone selections alone. */
messenger.mailTabs.onSelectedMessagesChanged.addListener(tab => {
  messenger.outlookThreadView.showParentThread(tab.id).catch(error => {
    console.error(
      "Outlook Style Companion could not display the thread conversation:",
      error
    );
  });
});
