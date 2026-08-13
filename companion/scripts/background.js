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

/* A restored or reactivated mail tab can already have a selection without
 * emitting a tree event. Re-evaluate that tab on activation and install the
 * privileged per-pane selection listener used for subsequent row changes. */
messenger.tabs.onActivated.addListener(({ tabId }) => {
  messenger.outlookThreadView.showParentThread(tabId).catch(error => {
    console.error(
      "Outlook Style Companion could not restore the thread conversation:",
      error
    );
  });
});

/* Apply the default immediately to an already-active restored tab when the
 * extension starts or updates, rather than waiting for the next click. */
messenger.tabs
  .query({ active: true })
  .then(tabs => {
    for (const tab of tabs) {
      messenger.outlookThreadView.showParentThread(tab.id).catch(error => {
        console.error(
          "Outlook Style Companion could not initialize the thread conversation:",
          error
        );
      });
    }
  })
  .catch(error => {
    console.error(
      "Outlook Style Companion could not inspect the active mail tab:",
      error
    );
  });
