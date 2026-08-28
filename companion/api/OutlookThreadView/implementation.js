const { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
const { cal } = ChromeUtils.importESModule(
  "resource:///modules/calendar/calUtils.sys.mjs"
);
const { CalMetronome } = ChromeUtils.importESModule(
  "resource:///modules/CalMetronome.sys.mjs"
);
const { CalAttendee } = ChromeUtils.importESModule(
  "resource:///modules/CalAttendee.sys.mjs"
);
const { MailServices } = ChromeUtils.importESModule(
  "resource:///modules/MailServices.sys.mjs"
);
const { MailUtils } = ChromeUtils.importESModule(
  "resource:///modules/MailUtils.sys.mjs"
);
const { openLinkExternally } = ChromeUtils.importESModule(
  "resource:///modules/LinkHelper.sys.mjs"
);
const TODAY_PANE_LISTENER_ID = "fluent-mail-outlook-inspired-today-pane";
const MAIN_PANE_SCHEME_LISTENER_ID =
  "outlook-style-companion-main-pane-scheme";
const MESSAGE_DISPLAY_SURFACE_LISTENER_ID =
  "outlook-style-companion-message-display-surface";
const MESSENGER_WINDOW_URL = "chrome://messenger/content/messenger.xhtml";
const MESSAGE_WINDOW_URL =
  "chrome://messenger/content/messageWindow.xhtml";
const REMINDER_DIALOG_LISTENER_ID =
  "outlook-style-companion-reminder-dialog";
const REMINDER_DIALOG_WINDOW_URL =
  "chrome://calendar/content/calendar-alarm-dialog.xhtml";
const EDITOR_SURFACE_LISTENER_ID =
  "outlook-style-companion-editor-surfaces";
const COMPOSE_WINDOW_URL =
  "chrome://messenger/content/messengercompose/messengercompose.xhtml";
const SPELL_CHECK_DIALOG_WINDOW_URL =
  "chrome://messenger/content/messengercompose/EdSpellCheck.xhtml";
const EVENT_DIALOG_WINDOW_URL =
  "chrome://calendar/content/calendar-event-dialog.xhtml";
const EVENT_ATTENDEES_DIALOG_WINDOW_URL =
  "chrome://calendar/content/calendar-event-dialog-attendees.xhtml";
const EVENT_SUMMARY_DIALOG_WINDOW_URL =
  "chrome://calendar/content/calendar-summary-dialog.xhtml";
const CALENDAR_CHOOSER_LISTENER_ID =
  "outlook-style-companion-calendar-chooser";
const CALENDAR_CHOOSER_WINDOW_URL =
  "chrome://calendar/content/chooseCalendarDialog.xhtml";
const HTML_NS = "http://www.w3.org/1999/xhtml";
const PANE_WIDTHS_PREF =
  "extensions.fluent-mail-outlook-inspired.pane-widths-v106";
const EVENT_WINDOW_WIDTH_PREF =
  "extensions.fluent-mail-outlook-inspired.event-window-width";
const EVENT_WINDOW_HEIGHT_PREF =
  "extensions.fluent-mail-outlook-inspired.event-window-height";
const DEFAULT_EVENT_WINDOW_WIDTH = 900;
const DEFAULT_EVENT_WINDOW_HEIGHT = 840;
const DEFAULT_ATTENDEES_WINDOW_WIDTH = 1100;
const DEFAULT_ATTENDEES_WINDOW_HEIGHT = 720;
const CONVERSATION_VIEW_PREF = "mail.thread.conversation.enabled";
const GLOBAL_INDEXER_PREF = "mailnews.database.global.indexer.enabled";
const CONVERSATION_VIEW_STYLE_ID = "outlook-style-conversation-view";
const MESSAGE_DOCUMENT_URI_PATTERN =
  /^(?:imap|mailbox|news|nntp|snews|file):/i;
const CONVERSATION_VIEW_LOAD_TIMEOUT_MS = 4000;
const THREAD_PARENT_OPEN_RETRY_DELAY_MS = 100;
const THREAD_PARENT_OPEN_MAX_ATTEMPTS = 60;
const OWNED_CONVERSATION_ATTRIBUTE = "data-outlook-style-owned-conversation";
const MSG_VIEW_FLAG_DUMMY = 0x20000000;
const enhancedTodayPaneWindows = new Set();
const todayPaneState = new WeakMap();
let todayPaneListenerRegistered = false;
const enhancedReminderDialogWindows = new Set();
const reminderDialogState = new WeakMap();
let reminderDialogListenerRegistered = false;
const enhancedEditorSurfaceWindows = new Set();
const editorSurfaceState = new WeakMap();
let editorSurfaceListenerRegistered = false;
const enhancedMainPaneSchemeWindows = new Set();
const mainPaneSchemeState = new WeakMap();
let mainPaneSchemeListenerRegistered = false;
const enhancedMessageDisplayWindows = new Set();
const messageDisplayWindowState = new WeakMap();
const messageDisplayDocumentStyles = new WeakMap();
let messageDisplaySurfaceListenerRegistered = false;
const enhancedCalendarChooserWindows = new Set();
const calendarChooserState = new WeakMap();
let calendarChooserListenerRegistered = false;
const guardedConversationViews = new Set();
const conversationGuardState = new WeakMap();
const conversationSelectionCapturePanes = new Set();
const conversationSelectionCaptureState = new WeakMap();
const unreadThreadDefaultsAppliedViews = new WeakSet();
const nativeEditorRequestKeys = new Set();
const unifiedItemDescriptionMenuState = new WeakMap();

const OUTLOOK_COLOR_SCHEME_ATTRIBUTE = "data-outlook-color-scheme";

function runEnhancementStep(description, callback) {
  try {
    return callback();
  } catch (error) {
    console.error(
      `Outlook Style Companion could not install ${description}:`,
      error
    );
    return false;
  }
}

function enhanceWindowSafely(description, window, enhance, rollback) {
  try {
    return enhance(window);
  } catch (error) {
    console.error(
      `Outlook Style Companion could not enhance ${description}:`,
      error
    );
    if (rollback) {
      try {
        rollback(window);
      } catch (cleanupError) {
        console.error(
          `Outlook Style Companion could not roll back ${description}:`,
          cleanupError
        );
      }
    }
    return false;
  }
}

function getMessageDisplayCss(scheme, preserveBodyCanvas) {
  const colors =
    scheme === "dark"
      ? {
          link: "#62abf5",
          linkVisited: "#96c6fa",
          surface: "#292929",
          text: "#ffffff",
        }
      : {
          link: "#0f6cbd",
          linkVisited: "#115ea3",
          surface: "#ffffff",
          text: "#242424",
        };
  const background = preserveBodyCanvas
    ? "transparent"
    : "var(--outlook-message-surface)";
  return `
  :root {
    --outlook-message-surface: ${colors.surface};
    --outlook-message-text: ${colors.text};
    --outlook-message-link: ${colors.link};
    --outlook-message-link-visited: ${colors.linkVisited};
    color-scheme: ${scheme};
    background-color: ${background};
    color: var(--outlook-message-text);
  }

  :where(a:link) {
    color: var(--outlook-message-link);
  }

  :where(a:visited) {
    color: var(--outlook-message-link-visited);
  }

  @media (forced-colors: active) {
    :root {
      background-color: Canvas;
      color: CanvasText;
    }

    :where(a:link),
    :where(a:visited) {
      color: LinkText;
    }
  }
`;
}

const CONVERSATION_VIEW_CSS = `
  :host {
    background: var(
      --outlook-surface,
      var(--layout-background-0, Canvas)
    );
    color: var(--outlook-text, var(--layout-color-1, CanvasText));
  }

  header {
    background: var(
      --outlook-surface-alt,
      var(--layout-background-1, Canvas)
    );
    color: var(--outlook-text, var(--layout-color-1, CanvasText));
    border-block-end-color: var(
      --outlook-divider,
      var(--color-surface-border, ButtonBorder)
    );
  }

  #mainConversation {
    background: var(
      --outlook-surface,
      var(--layout-background-0, Canvas)
    );
    color: var(--outlook-text, var(--layout-color-1, CanvasText));
  }

  #mainConversation > article[aria-expanded="false"] {
    box-sizing: border-box;
    min-block-size: 54px;
    background: var(
      --outlook-surface,
      var(--layout-background-0, Canvas)
    );
  }

  #mainConversation > article[aria-expanded="false"]:hover {
    background: var(
      --outlook-hover,
      var(--layout-background-2, ButtonFace)
    );
  }

  #mainConversation > article[aria-expanded="false"]:focus-visible {
    outline: 2px solid var(--outlook-accent, #0f6cbd);
    outline-offset: -2px;
    background: var(--outlook-accent-light, #cfe4fa);
  }

  #mainConversation > article[aria-expanded="true"]:focus-visible {
    outline: 2px solid var(--outlook-accent, #0f6cbd);
    outline-offset: -2px;
  }

  #mainConversation > article[aria-expanded="true"] {
    background: var(
      --outlook-surface,
      var(--layout-background-0, Canvas)
    );
  }

  #mainConversation > article[aria-expanded="true"] > browser {
    background: var(
      --outlook-surface,
      var(--layout-background-0, Canvas)
    );
    color-scheme: light dark;
  }

  :host([data-outlook-color-scheme="light"])
    #mainConversation > article[aria-expanded="true"] > browser {
    color-scheme: light;
  }

  :host([data-outlook-color-scheme="dark"])
    #mainConversation > article[aria-expanded="true"] > browser {
    color-scheme: dark;
  }

  #mainConversation > article + article {
    border-block-start-color: var(
      --outlook-divider,
      var(--color-surface-border, ButtonBorder)
    );
  }

  #mainConversation address {
    color: var(--outlook-text, var(--layout-color-1, CanvasText));
  }

  #mainConversation :is(time, p),
  header .details {
    color: var(
      --outlook-text-secondary,
      var(--layout-color-2, CanvasText)
    );
  }

  @media (forced-colors: active) {
    :host,
    header,
    #mainConversation,
    #mainConversation > article[aria-expanded] {
      background: Canvas;
      color: CanvasText;
      border-color: ButtonBorder;
    }

    #mainConversation > article[aria-expanded="false"]:hover {
      background: ButtonFace;
    }

    #mainConversation > article[aria-expanded="false"]:focus-visible {
      background: SelectedItem;
      color: SelectedItemText;
      outline: 2px solid SelectedItemText;
      outline-offset: -2px;
    }

    #mainConversation > article[aria-expanded="true"]:focus-visible {
      outline: 2px solid Highlight;
      outline-offset: -2px;
    }

    #mainConversation > article[aria-expanded="false"]:focus-visible
      :is(address, time, p) {
      color: SelectedItemText;
    }
  }
`;

function getOutlookColorScheme(window) {
  const nativeScheme = window.document?.documentElement?.getAttribute(
    "lwt-sidebar"
  );
  if (nativeScheme === "light" || nativeScheme === "dark") {
    return nativeScheme;
  }
  const stampedScheme = window.document?.documentElement?.getAttribute(
    OUTLOOK_COLOR_SCHEME_ATTRIBUTE
  );
  if (stampedScheme === "light" || stampedScheme === "dark") {
    return stampedScheme;
  }
  return window.matchMedia?.("(prefers-color-scheme: dark)")?.matches
    ? "dark"
    : "light";
}

function isAboutMessageWindow(window) {
  try {
    return (
      window?.document?.defaultView === window &&
      window.document.documentURI === "about:message"
    );
  } catch (error) {
    return false;
  }
}

function getDisplayedMimeDocument(aboutMessageWindow) {
  if (!isAboutMessageWindow(aboutMessageWindow)) {
    return null;
  }
  try {
    const messageBrowser =
      aboutMessageWindow.document.getElementById("messagepane");
    const document = messageBrowser?.contentDocument;
    if (
      !document?.documentElement ||
      messageBrowser.contentDocument !== document ||
      !MESSAGE_DOCUMENT_URI_PATTERN.test(document.documentURI || "")
    ) {
      return null;
    }
    return document;
  } catch (error) {
    /* The MIME browser can change remoteness while a message is replaced. */
    return null;
  }
}

function removeDisplayedMimeDocumentStyle(document) {
  const record = messageDisplayDocumentStyles.get(document);
  if (!record) {
    return;
  }
  try {
    document.adoptedStyleSheets = document.adoptedStyleSheets.filter(
      candidate => candidate !== record.styleSheet
    );
  } catch (error) {
    /* A detached message document may already have been destroyed. */
  }
  messageDisplayDocumentStyles.delete(document);
}

function ensureDisplayedMimeDocumentStyle(document, scheme) {
  const root = document?.documentElement;
  if (
    !root ||
    !MESSAGE_DOCUMENT_URI_PATTERN.test(document.documentURI || "") ||
    (scheme !== "light" && scheme !== "dark") ||
    typeof document.defaultView?.CSSStyleSheet !== "function" ||
    !("adoptedStyleSheets" in document)
  ) {
    return null;
  }

  let record = messageDisplayDocumentStyles.get(document);
  if (record && record.root !== root) {
    removeDisplayedMimeDocumentStyle(document);
    record = null;
  }
  if (!record) {
    let styleSheet;
    try {
      styleSheet = new document.defaultView.CSSStyleSheet();
      document.adoptedStyleSheets = [
        ...document.adoptedStyleSheets,
        styleSheet,
      ];
    } catch (error) {
      /* Fail closed if a MIME document rejects non-DOM style sheets. */
      return null;
    }
    record = {
      bodyChecked: false,
      preserveBodyCanvas: false,
      root,
      scheme,
      styleSheet,
    };
    messageDisplayDocumentStyles.set(document, record);
  }
  record.scheme = scheme;
  try {
    record.styleSheet.replaceSync(
      getMessageDisplayCss(record.scheme, record.preserveBodyCanvas)
    );
  } catch (error) {
    removeDisplayedMimeDocumentStyle(document);
    return null;
  }
  return record;
}

function finalizeDisplayedMimeDocumentStyle(document, record) {
  if (!record || record.bodyChecked) {
    return;
  }
  const body = document?.body;
  if (!body || record.root !== document.documentElement) {
    return;
  }
  try {
    const bodyStyle = document.defaultView.getComputedStyle(body);
    const backgroundColor = bodyStyle.backgroundColor
      .replaceAll(" ", "")
      .toLowerCase();
    const hasOpaqueBackground =
      backgroundColor !== "transparent" &&
      backgroundColor !== "rgba(0,0,0,0)";
    const hasBackgroundImage = bodyStyle.backgroundImage !== "none";
    record.preserveBodyCanvas = hasOpaqueBackground || hasBackgroundImage;
    record.styleSheet.replaceSync(
      getMessageDisplayCss(record.scheme, record.preserveBodyCanvas)
    );
    record.bodyChecked = true;
  } catch (error) {
    /* A document replaced during MsgLoaded will be handled by its own event. */
  }
}

function styleAboutMessageWindow(aboutMessageWindow, scheme, finalize) {
  const document = getDisplayedMimeDocument(aboutMessageWindow);
  if (!document) {
    return;
  }
  const record = ensureDisplayedMimeDocumentStyle(document, scheme);
  if (finalize) {
    finalizeDisplayedMimeDocumentStyle(document, record);
  }
}

function forEachAboutMessageWindow(window, callback) {
  let browsingContexts = [];
  try {
    browsingContexts =
      window.browsingContext?.getAllBrowsingContextsInSubtree() || [];
  } catch (error) {
    return;
  }
  for (const browsingContext of browsingContexts) {
    try {
      if (browsingContext.currentURI?.spec === "about:message") {
        callback(browsingContext.window);
      }
    } catch (error) {
      /* A tab or conversation browser can detach during the traversal. */
    }
  }
}

function syncDisplayedMimeDocuments(window) {
  if (window.closed) {
    return;
  }
  const scheme = getOutlookColorScheme(window);
  forEachAboutMessageWindow(window, aboutMessageWindow => {
    stampOutlookColorScheme(aboutMessageWindow.document, scheme);
    installInvitationAvailabilitySurface(aboutMessageWindow.document);
    let finalize = false;
    try {
      finalize =
        aboutMessageWindow.msgLoaded === true &&
        aboutMessageWindow.msgLoading !== true;
    } catch (error) {
      /* A reader can detach between traversal and state inspection. */
    }
    styleAboutMessageWindow(aboutMessageWindow, scheme, finalize);
  });
}

function installMessageDisplaySurfaceBridge(window) {
  if (messageDisplayWindowState.has(window)) {
    return true;
  }
  const url = window.document?.location?.href;
  if (url !== MESSENGER_WINDOW_URL && url !== MESSAGE_WINDOW_URL) {
    return false;
  }

  const state = {
    schemeFrame: 0,
    shuttingDown: false,
  };
  const onMsgLoading = event => {
    if (state.shuttingDown || !isAboutMessageWindow(event.target)) {
      return;
    }
    /* MsgLoading identifies the exact about:message reader synchronously.
     * Styling its current MIME document directly avoids resolving a tab after
     * a conversation browser has already been replaced or detached. */
    styleAboutMessageWindow(
      event.target,
      getOutlookColorScheme(window),
      false
    );
    stampOutlookColorScheme(
      event.target.document,
      getOutlookColorScheme(window)
    );
    installInvitationAvailabilitySurface(event.target.document);
  };
  const onMsgLoaded = event => {
    if (state.shuttingDown || !isAboutMessageWindow(event.target)) {
      return;
    }
    /* Capture before about:message's target listener can adapt the MIME body
     * for dark mode. Authored-canvas detection must inspect the original
     * message, not colors introduced by Thunderbird's DarkReader. */
    styleAboutMessageWindow(
      event.target,
      getOutlookColorScheme(window),
      true
    );
    stampOutlookColorScheme(
      event.target.document,
      getOutlookColorScheme(window)
    );
    installInvitationAvailabilitySurface(event.target.document);
  };
  const syncScheme = () => {
    state.schemeFrame = 0;
    if (!state.shuttingDown) {
      syncDisplayedMimeDocuments(window);
    }
  };
  const scheduleSchemeSync = () => {
    if (window.closed || state.shuttingDown || state.schemeFrame) {
      return;
    }
    state.schemeFrame = window.requestAnimationFrame(syncScheme);
  };

  const systemScheme = window.matchMedia("(prefers-color-scheme: dark)");
  const schemeObserver = new window.MutationObserver(scheduleSchemeSync);
  Object.assign(state, {
    onMsgLoaded,
    onMsgLoading,
    scheduleSchemeSync,
    schemeObserver,
    systemScheme,
  });
  messageDisplayWindowState.set(window, state);
  enhancedMessageDisplayWindows.add(window);

  schemeObserver.observe(window.document.documentElement, {
    attributes: true,
    attributeFilter: [
      OUTLOOK_COLOR_SCHEME_ATTRIBUTE,
      "class",
      "lwt-sidebar",
      "style",
    ],
  });
  window.addEventListener("MsgLoading", onMsgLoading, true);
  window.addEventListener("MsgLoaded", onMsgLoaded, true);
  window.addEventListener("windowlwthemeupdate", scheduleSchemeSync);
  systemScheme.addEventListener("change", scheduleSchemeSync);
  syncDisplayedMimeDocuments(window);
  return true;
}

function removeMessageDisplaySurfaceBridge(window) {
  const state = messageDisplayWindowState.get(window);
  if (!state) {
    return;
  }
  state.shuttingDown = true;
  window.removeEventListener("MsgLoading", state.onMsgLoading, true);
  window.removeEventListener("MsgLoaded", state.onMsgLoaded, true);
  window.removeEventListener(
    "windowlwthemeupdate",
    state.scheduleSchemeSync
  );
  state.systemScheme.removeEventListener("change", state.scheduleSchemeSync);
  state.schemeObserver.disconnect();
  if (state.schemeFrame) {
    window.cancelAnimationFrame(state.schemeFrame);
  }
  forEachAboutMessageWindow(window, aboutMessageWindow => {
    const document = getDisplayedMimeDocument(aboutMessageWindow);
    if (document) {
      removeDisplayedMimeDocumentStyle(document);
    }
  });
  messageDisplayWindowState.delete(window);
  enhancedMessageDisplayWindows.delete(window);
}

function installMessageDisplaySurfaceEnhancement() {
  if (messageDisplaySurfaceListenerRegistered) {
    return true;
  }
  const enhanceWindow = window =>
    enhanceWindowSafely(
      "a message display window",
      window,
      installMessageDisplaySurfaceBridge,
      removeMessageDisplaySurfaceBridge
    );
  ExtensionSupport.registerWindowListener(
    MESSAGE_DISPLAY_SURFACE_LISTENER_ID,
    {
      chromeURLs: [MESSENGER_WINDOW_URL, MESSAGE_WINDOW_URL],
      onLoadWindow: enhanceWindow,
      onUnloadWindow: removeMessageDisplaySurfaceBridge,
    }
  );
  messageDisplaySurfaceListenerRegistered = true;

  const windows = Services.wm.getEnumerator(null);
  while (windows.hasMoreElements()) {
    enhanceWindow(windows.getNext());
  }
  return true;
}

function sameMessage(left, right) {
  try {
    const leftHeader = left?.folderMessage || left;
    const rightHeader = right?.folderMessage || right;
    const leftKey = leftHeader?.messageKey;
    const rightKey = rightHeader?.messageKey;
    const leftFolder = String(leftHeader?.folder?.URI || "").trim();
    const rightFolder = String(rightHeader?.folder?.URI || "").trim();
    return Boolean(
      leftHeader &&
        rightHeader &&
        Number.isInteger(leftKey) &&
        Number.isInteger(rightKey) &&
        leftKey >= 0 &&
        rightKey >= 0 &&
        leftKey !== 0xffffffff &&
        rightKey !== 0xffffffff &&
        leftFolder &&
        rightFolder &&
        leftKey === rightKey &&
        leftFolder === rightFolder
    );
  } catch (error) {
    /* A Gloda result can lose its backing folder header during teardown. */
    return false;
  }
}

function getMessageId(message) {
  try {
    return String(
      message?.folderMessage?.messageId || message?.messageId || ""
    ).trim();
  } catch (error) {
    return "";
  }
}

function getNativeMessageIdentity(message) {
  try {
    const header = message?.folderMessage || message;
    const key = header?.messageKey;
    const folder = String(header?.folder?.URI || "").trim();
    if (
      !header ||
      !Number.isInteger(key) ||
      key < 0 ||
      key === 0xffffffff ||
      !folder
    ) {
      return "";
    }
    return `${folder.length}:${folder}:${key}`;
  } catch (error) {
    return "";
  }
}

function hasUniqueMessageIds(messages) {
  const ids = messages.map(getMessageId);
  return Boolean(
    ids.length && ids.every(Boolean) && new Set(ids).size === ids.length
  );
}

function sameCrossFolderCopy(left, right) {
  try {
    const leftHeader = left?.folderMessage || left;
    const rightHeader = right?.folderMessage || right;
    return Boolean(
      getMessageId(leftHeader) &&
        getMessageId(leftHeader) === getMessageId(rightHeader) &&
        Number(leftHeader?.dateInSeconds) ===
          Number(rightHeader?.dateInSeconds) &&
        String(leftHeader?.author || "") ===
          String(rightHeader?.author || "") &&
        String(leftHeader?.subject || "") ===
          String(rightHeader?.subject || "") &&
        Number(leftHeader?.messageSize) === Number(rightHeader?.messageSize)
    );
  } catch (error) {
    /* A cross-folder Gloda wrapper can lose its backing header mid-query. */
    return false;
  }
}

/**
 * Match every visible thread header to exactly one Gloda conversation result.
 * Prefer Thunderbird's native folder/message-key identity. A Message-ID is
 * used only for a unique cross-folder copy (for example Gmail All Mail), never
 * as an unqualified or duplicate identity.
 */
function matchConversationMessages(loadedMessages, expectedMessages) {
  if (
    !hasUniqueMessageIds(loadedMessages) ||
    !hasUniqueMessageIds(expectedMessages)
  ) {
    return null;
  }

  const loadedByNativeIdentity = new Map();
  const loadedByMessageId = new Map();
  for (const loaded of loadedMessages) {
    const nativeIdentity = getNativeMessageIdentity(loaded);
    if (nativeIdentity) {
      const candidates = loadedByNativeIdentity.get(nativeIdentity) || [];
      candidates.push(loaded);
      loadedByNativeIdentity.set(nativeIdentity, candidates);
    }
    const messageId = getMessageId(loaded);
    const messageIdCandidates = loadedByMessageId.get(messageId) || [];
    messageIdCandidates.push(loaded);
    loadedByMessageId.set(messageId, messageIdCandidates);
  }

  const matches = new Map();
  const usedLoadedMessages = new Set();
  for (const expected of expectedMessages) {
    const nativeCandidates = (
      loadedByNativeIdentity.get(getNativeMessageIdentity(expected)) || []
    ).filter(message => !usedLoadedMessages.has(message));
    let candidate = null;
    if (nativeCandidates.length === 1) {
      candidate = nativeCandidates[0];
    } else if (!nativeCandidates.length) {
      const crossFolderCandidates = (
        loadedByMessageId.get(getMessageId(expected)) || []
      ).filter(message => !usedLoadedMessages.has(message));
      const crossFolderCandidate =
        crossFolderCandidates.length === 1
          ? crossFolderCandidates[0]
          : null;
      if (
        crossFolderCandidate &&
        sameCrossFolderCopy(crossFolderCandidate, expected)
      ) {
        candidate = crossFolderCandidate;
      }
    }
    if (!candidate) {
      return null;
    }
    matches.set(expected, candidate);
    usedLoadedMessages.add(candidate);
  }
  return matches;
}

function getCurrentCollapsedThreadParent(about3Pane, selectedIndex) {
  try {
    const dbView = about3Pane?.gDBView;
    const threadTree = about3Pane?.threadTree;
    if (
      about3Pane?.closed ||
      !dbView ||
      !threadTree ||
      !about3Pane.gViewWrapper?.showThreaded ||
      threadTree.selectedIndex !== selectedIndex ||
      dbView.selection?.count !== 1 ||
      threadTree.selectedIndices?.length !== 1 ||
      threadTree.selectedIndices[0] !== selectedIndex ||
      dbView.numSelected <= 1 ||
      dbView.getFlagsAt(selectedIndex) & MSG_VIEW_FLAG_DUMMY ||
      !dbView.isContainer(selectedIndex) ||
      dbView.isContainerOpen(selectedIndex)
    ) {
      return null;
    }

    const thread = dbView.getThreadContainingIndex(selectedIndex);
    const selectedMessage = dbView.getMsgHdrAt(selectedIndex);
    const rootMessage = thread?.getRootHdr();
    if (
      !thread ||
      thread.numChildren < 2 ||
      !sameMessage(selectedMessage, rootMessage)
    ) {
      return null;
    }

    const expectedMessages = [];
    for (let index = 0; index < thread.numChildren; index++) {
      expectedMessages.push(thread.getChildHdrAt(index));
    }
    const selectedMessages = dbView.getSelectedMsgHdrs();
    if (
      selectedMessages.length <= 1 ||
      selectedMessages.length !== expectedMessages.length
    ) {
      return null;
    }
    const unmatchedExpected = [...expectedMessages];
    for (const selected of selectedMessages) {
      const matchIndex = unmatchedExpected.findIndex(expected =>
        sameMessage(selected, expected)
      );
      if (matchIndex < 0) {
        return null;
      }
      unmatchedExpected.splice(matchIndex, 1);
    }
    if (unmatchedExpected.length) {
      return null;
    }

    return {
      about3Pane,
      dbView,
      expectedMessages,
      rootMessage,
      selectedIndex,
      threadTree,
    };
  } catch (error) {
    /* A view can be replaced between the mouse event and this inspection. */
    return null;
  }
}

function isCurrentCollapsedThreadParentActivation(activation) {
  try {
    const topWindow = activation.about3Pane.browsingContext?.topChromeWindow;
    const tabmail = topWindow?.document?.getElementById("tabmail");
    if (
      activation.about3Pane.tabOrWindow?.selected !== true ||
      tabmail?.currentAbout3Pane !== activation.about3Pane
    ) {
      return false;
    }
  } catch (error) {
    return false;
  }
  const current = getCurrentCollapsedThreadParent(
    activation.about3Pane,
    activation.selectedIndex
  );
  if (
    !current ||
    current.dbView !== activation.dbView ||
    current.threadTree !== activation.threadTree ||
    !sameMessage(current.rootMessage, activation.rootMessage) ||
    current.expectedMessages.length !== activation.expectedMessages.length
  ) {
    return false;
  }
  const unmatchedCurrent = [...current.expectedMessages];
  for (const expected of activation.expectedMessages) {
    const matchIndex = unmatchedCurrent.findIndex(currentMessage =>
      sameMessage(expected, currentMessage)
    );
    if (matchIndex < 0) {
      return false;
    }
    unmatchedCurrent.splice(matchIndex, 1);
  }
  return !unmatchedCurrent.length;
}

/**
 * Resolve the exact complete message currently visible in the reading pane.
 * A collapsed thread can include hidden rows and cross-folder Gloda results,
 * so never guess from DB row order or trust a Message-ID by itself.
 */
function getVerifiedDisplayedThreadMessage(about3Pane, activation) {
  try {
    const { expectedMessages, rootMessage } = activation;
    const messagePane = about3Pane.document.querySelector("message-pane");
    if (
      !messagePane?.isConnected ||
      about3Pane.paneLayout?.messagePaneVisible?.isCollapsed ||
      messagePane.classList.contains("collapsed-by-splitter")
    ) {
      return null;
    }

    const conversationView = messagePane.querySelector(
      ":scope > conversation-view"
    );
    if (conversationView && !conversationView.hidden) {
      if (
        !conversationView.isConnected ||
        messagePane.conversationView !== conversationView
      ) {
        return null;
      }
    }
    if (
      conversationView?.isConnected &&
      !conversationView.hidden &&
      messagePane.conversationView === conversationView
    ) {
      if (!hasUniqueMessageIds(expectedMessages)) {
        return null;
      }
      const guard = conversationGuardState.get(conversationView);
      if (
        guard &&
        (guard.disabled ||
          !guard.ready ||
          !isExpectedConversationSelection(guard))
      ) {
        return null;
      }
      const main = conversationView.shadowRoot?.getElementById(
        "mainConversation"
      );
      const expandedArticles = [...(main?.children || [])].filter(child =>
        child.matches?.('article[aria-expanded="true"]')
      );
      if (expandedArticles.length !== 1) {
        return null;
      }
      const browsers = [...expandedArticles[0].children].filter(child =>
        child.matches?.('browser[src="about:message"]')
      );
      if (browsers.length !== 1 || browsers[0].hidden) {
        return null;
      }
      const browser = browsers[0];
      if (browser.contentWindow?.location?.href !== "about:message") {
        return null;
      }
      const displayedMessage = browser.contentWindow.gMessage;
      if (!getNativeMessageIdentity(displayedMessage)) {
        return null;
      }
      const loadedMessages = Array.isArray(conversationView.messages)
        ? conversationView.messages
        : [];
      if (!matchConversationMessages(loadedMessages, expectedMessages)) {
        return null;
      }
      const loadedMatches = loadedMessages.filter(message =>
        sameMessage(message, displayedMessage)
      );
      if (loadedMatches.length !== 1) {
        return null;
      }
      const articleMessageId = String(
        expandedArticles[0].dataset.messageId || ""
      ).trim();
      const loadedMessageId = getMessageId(loadedMatches[0]);
      const displayedMessageId = getMessageId(displayedMessage);
      if (
        !articleMessageId ||
        articleMessageId !== loadedMessageId ||
        articleMessageId !== displayedMessageId
      ) {
        return null;
      }
      return displayedMessage;
    }

    /* Gloda can fail closed to Thunderbird's ordinary full-message reader.
     * Honor that exact visible message too, but only when it is a unique
     * native member of the selected thread. */
    const browser = messagePane.messageBrowser;
    if (
      !browser?.isConnected ||
      browser.hidden ||
      browser.contentWindow?.location?.href !== "about:message"
    ) {
      return null;
    }
    const displayedMessage = browser.contentWindow.gMessage;
    if (!getNativeMessageIdentity(displayedMessage)) {
      return null;
    }
    if (sameMessage(rootMessage, displayedMessage)) {
      return displayedMessage;
    }
    if (
      !sameCrossFolderCopy(rootMessage, displayedMessage) ||
      !getMessageId(displayedMessage) ||
      getMessageId(rootMessage) !== getMessageId(displayedMessage)
    ) {
      return null;
    }
    return displayedMessage;
  } catch (error) {
    /* Readers and their XPConnect headers can disappear during navigation. */
    return null;
  }
}

function openSingleActivatedThreadMessage(about3Pane, message, shiftKey) {
  try {
    const topWindow = about3Pane?.browsingContext?.topChromeWindow;
    const folder = message?.folder;
    const messageUri = folder?.getUriForMsg(message);
    if (!topWindow || !folder || !messageUri) {
      return false;
    }

    let composeType = null;
    if (folder.isSpecialFolder(Ci.nsMsgFolderFlags.Drafts, true)) {
      composeType = Ci.nsIMsgCompType.Draft;
    } else if (
      folder.isSpecialFolder(Ci.nsMsgFolderFlags.Templates, true)
    ) {
      composeType = Ci.nsIMsgCompType.Template;
    }
    if (composeType !== null) {
      if (typeof topWindow.ComposeMessage !== "function") {
        return false;
      }
      const composeResult = topWindow.ComposeMessage(
        composeType,
        shiftKey
          ? Ci.nsIMsgCompFormat.OppositeOfDefault
          : Ci.nsIMsgCompFormat.Default,
        folder,
        [messageUri],
        null,
        undefined
      );
      Promise.resolve(composeResult).catch(error => {
        console.error(
          "Outlook Style Companion could not open the displayed draft or template:",
          error
        );
      });
      return true;
    }

    const tabmail = topWindow.document?.getElementById("tabmail");
    if (!tabmail) {
      return false;
    }
    MailUtils.displayMessages(
      [message],
      about3Pane.gViewWrapper,
      tabmail,
      false,
      shiftKey
    );
    return true;
  } catch (error) {
    console.error(
      "Outlook Style Companion could not open the displayed thread message:",
      error
    );
    return false;
  }
}

function isExpandedSelectedRoot(
  dbView,
  selectedIndex,
  selectedMessage,
  rootMessage
) {
  return (
    sameMessage(selectedMessage, rootMessage) &&
    dbView.isContainer(selectedIndex) &&
    dbView.isContainerOpen(selectedIndex)
  );
}

function isExpectedConversationSelection(state) {
  const {
    about3Pane,
    conversationView,
    dbView: expectedDbView,
    messagePane,
    rootMessage,
    selectedMessage,
    selectedRootOpen,
  } = state;
  try {
    const dbView = about3Pane.gDBView;
    const threadTree = about3Pane.threadTree;
    const selectedIndex = threadTree?.selectedIndex ?? -1;
    const currentThread =
      selectedIndex >= 0
        ? dbView?.getThreadContainingIndex(selectedIndex)
        : null;
    return (
      !about3Pane.closed &&
      dbView === expectedDbView &&
      about3Pane.gViewWrapper?.showThreaded &&
      conversationView.isConnected &&
      messagePane.querySelector(":scope > conversation-view") ===
        conversationView &&
      selectedIndex >= 0 &&
      dbView?.selection?.count === 1 &&
      threadTree.selectedIndices?.length === 1 &&
      sameMessage(dbView.getMsgHdrAt(selectedIndex), selectedMessage) &&
      sameMessage(currentThread?.getRootHdr(), rootMessage) &&
      isExpandedSelectedRoot(
        dbView,
        selectedIndex,
        dbView.getMsgHdrAt(selectedIndex),
        currentThread?.getRootHdr()
      ) === selectedRootOpen
    );
  } catch (error) {
    return false;
  }
}

function removeConversationGuard(conversationView) {
  const state = conversationGuardState.get(conversationView);
  if (state?.timeoutId) {
    state.about3Pane.clearTimeout(state.timeoutId);
  }
  if (state?.onBrowserLoad) {
    state.main.removeEventListener("load", state.onBrowserLoad, true);
  }
  state?.observer.disconnect();
  conversationGuardState.delete(conversationView);
  guardedConversationViews.delete(conversationView);
}

function retireConversationView(messagePane, conversationView) {
  if (!conversationView) {
    return;
  }
  conversationView.hidden = true;
  /* Detach the host from MessagePane before removing it. Thunderbird's own
   * selection cleanup may run immediately afterward; leaving this reference
   * live makes it call clear() on an already-disconnected about:message
   * browser whose displayMessage function is no longer available. */
  if (messagePane?.conversationView === conversationView) {
    messagePane.conversationView = null;
  }
  removeConversationGuard(conversationView);
  conversationView.remove();
}

function threadContainsUnreadMessage(thread) {
  try {
    for (let index = 0; index < thread.numChildren; index++) {
      const message = thread.getChildHdrAt(index);
      if (!(message.flags & Ci.nsMsgMessageFlags.Read)) {
        return true;
      }
    }
  } catch (error) {
    /* A live view can replace a thread while its folder is refreshing. */
  }
  return false;
}

/**
 * Establish the Outlook-style default for a newly loaded threaded view.
 * Read threads begin collapsed and threads containing any unread message begin
 * expanded. The view is handled only once so later manual toggles remain under
 * the user's control. Restoring the selection may reveal a selected read reply.
 */
function applyUnreadThreadDefaults(about3Pane) {
  const dbView = about3Pane?.gDBView;
  const viewWrapper = about3Pane?.gViewWrapper;
  const threadPane = about3Pane?.threadPane;
  const threadTree = about3Pane?.threadTree;
  if (
    about3Pane?.closed ||
    !dbView ||
    !viewWrapper?.showThreaded ||
    viewWrapper.showGroupedBySort ||
    !threadPane ||
    !threadTree ||
    about3Pane.dbViewWrapperListener?.allMessagesLoaded !== true ||
    unreadThreadDefaultsAppliedViews.has(dbView)
  ) {
    return false;
  }

  /* Capture message URIs before collapsing so Thunderbird can restore the
   * exact selected reply after the visible row indices change. */
  threadPane.saveSelection();
  dbView.doCommand(Ci.nsMsgViewCommandType.collapseAll);
  viewWrapper._threadExpandAll = false;

  /* Expanding from the bottom keeps every not-yet-visited root index stable as
   * child rows are inserted beneath later roots. */
  const unreadThreadRootIndices = [];
  for (let index = dbView.rowCount - 1; index >= 0; index--) {
    if (
      dbView.getFlagsAt(index) & MSG_VIEW_FLAG_DUMMY ||
      !dbView.isContainer(index) ||
      dbView.isContainerOpen(index)
    ) {
      continue;
    }
    const thread = dbView.getThreadContainingIndex(index);
    if (thread?.numChildren > 1 && threadContainsUnreadMessage(thread)) {
      unreadThreadRootIndices.push(index);
    }
  }
  for (const index of unreadThreadRootIndices) {
    threadTree.expandRowAtIndex(index);
  }

  threadPane.restoreSelection();
  unreadThreadDefaultsAppliedViews.add(dbView);
  return true;
}

function removeConversationSelectionCapture(about3Pane) {
  const state = conversationSelectionCaptureState.get(about3Pane);
  if (!state) {
    return;
  }
  state.threadTree.removeEventListener("select", state.onSelect, true);
  state.threadTree.removeEventListener(
    "dblclick",
    state.onDoubleClick,
    true
  );
  about3Pane.removeEventListener(
    "allMessagesLoaded",
    state.onAllMessagesLoaded
  );
  about3Pane.removeEventListener("unload", state.onUnload);
  if (state.refreshTimer) {
    about3Pane.clearTimeout(state.refreshTimer);
  }
  if (state.pendingOpenTimer) {
    about3Pane.clearTimeout(state.pendingOpenTimer);
  }
  conversationSelectionCaptureState.delete(about3Pane);
  conversationSelectionCapturePanes.delete(about3Pane);
}

function retireConversationGuardsForPane(about3Pane) {
  for (const conversationView of [...guardedConversationViews]) {
    const state = conversationGuardState.get(conversationView);
    if (state?.about3Pane === about3Pane) {
      /* The pane is unloading, so remove its observers and timers without
       * trying to recreate a native reader inside the closing document. */
      retireConversationView(state.messagePane, conversationView);
    }
  }
}

function installConversationSelectionCapture(about3Pane, refreshConversation) {
  const threadTree = about3Pane?.threadTree;
  if (!threadTree || about3Pane.closed) {
    return false;
  }
  const previous = conversationSelectionCaptureState.get(about3Pane);
  if (previous?.threadTree === threadTree) {
    if (typeof refreshConversation === "function") {
      previous.refreshConversation = refreshConversation;
    }
    return true;
  }
  removeConversationSelectionCapture(about3Pane);

  let state = null;
  const scheduleRefresh = () => {
    if (!state || about3Pane.closed) {
      return;
    }
    if (state.refreshTimer) {
      about3Pane.clearTimeout(state.refreshTimer);
    }
    state.refreshTimer = about3Pane.setTimeout(() => {
      state.refreshTimer = 0;
      Promise.resolve(state.refreshConversation?.()).catch(error => {
        console.error(
          "Outlook Style Companion could not refresh the selected thread:",
          error
        );
      });
    }, 0);
  };

  const clearPendingOpen = () => {
    if (!state) {
      return;
    }
    if (state.pendingOpenTimer) {
      about3Pane.clearTimeout(state.pendingOpenTimer);
      state.pendingOpenTimer = 0;
    }
    state.pendingOpen = null;
  };

  const tryPendingOpen = () => {
    const pending = state?.pendingOpen;
    if (!pending) {
      return;
    }
    state.pendingOpenTimer = 0;
    if (!isCurrentCollapsedThreadParentActivation(pending.activation)) {
      clearPendingOpen();
      return;
    }
    const displayedMessage = getVerifiedDisplayedThreadMessage(
      about3Pane,
      pending.activation
    );
    if (displayedMessage) {
      const { shiftKey } = pending;
      clearPendingOpen();
      openSingleActivatedThreadMessage(
        about3Pane,
        displayedMessage,
        shiftKey
      );
      return;
    }
    pending.attemptsRemaining--;
    if (pending.attemptsRemaining <= 0) {
      clearPendingOpen();
      return;
    }
    state.pendingOpenTimer = about3Pane.setTimeout(
      tryPendingOpen,
      THREAD_PARENT_OPEN_RETRY_DELAY_MS
    );
  };

  const onDoubleClick = event => {
    if (
      event.target?.closest?.("button") ||
      event.target?.closest?.("menupopup")
    ) {
      return;
    }
    const row = event.target?.closest?.('tr[is^="thread-"]');
    if (!row?.isConnected || !threadTree.contains(row)) {
      return;
    }
    const selectedIndex = row.index;
    let isCollapsedAggregate = false;
    try {
      const dbView = about3Pane.gDBView;
      isCollapsedAggregate = Boolean(
        Number.isInteger(selectedIndex) &&
          selectedIndex >= 0 &&
          selectedIndex === threadTree.selectedIndex &&
          dbView?.selection?.count === 1 &&
          threadTree.selectedIndices?.length === 1 &&
          dbView.numSelected > 1
      );
      if (!isCollapsedAggregate) {
        return;
      }
    } catch (error) {
      return;
    }

    /* Thunderbird expands a collapsed parent into every hidden child before
     * running cmd_openMessage. Claim the event before any fallible work so a
     * stale reader can open zero messages, but can never fall through and open
     * one window or tab per child. */
    event.preventDefault();
    event.stopImmediatePropagation();

    if (event.button !== 0) {
      clearPendingOpen();
      return;
    }

    try {
      if (threadTree.getRowAtIndex(selectedIndex) !== row) {
        clearPendingOpen();
        return;
      }
    } catch (error) {
      clearPendingOpen();
      return;
    }

    const activation = getCurrentCollapsedThreadParent(
      about3Pane,
      selectedIndex
    );
    if (!activation) {
      clearPendingOpen();
      return;
    }
    clearPendingOpen();
    state.pendingOpen = {
      activation,
      attemptsRemaining: THREAD_PARENT_OPEN_MAX_ATTEMPTS,
      shiftKey: Boolean(event.shiftKey),
    };
    tryPendingOpen();
  };

  /* Thunderbird clears the previous reader synchronously in its own bubble
   * listener. A newly-cloned conversation browser is briefly about:blank and
   * does not yet expose displayMessage(), so clear() can throw during rapid
   * row changes. Detach an unsafe host in capture phase, before native selection
   * handling; Thunderbird then renders the stock selection and the Companion
   * replaces it with the next guarded conversation afterward. */
  const onSelect = () => {
    clearPendingOpen();
    const messagePane = about3Pane.document.querySelector("message-pane");
    const conversationView = messagePane?.querySelector(
      ":scope > conversation-view"
    );
    if (conversationView) {
      let shouldRetire =
        conversationView.getAttribute(OWNED_CONVERSATION_ATTRIBUTE) === "true";
      if (!shouldRetire) {
        const browser = conversationView.shadowRoot?.querySelector(
          'browser[src="about:message"]'
        );
        if (!browser) {
          /* The native Gloda query has not completed. Reusing this host would
           * let its uncancelled result overwrite the next selection. */
          shouldRetire = true;
        } else {
          try {
            /* Also cover the short pre-ownership interval when the user enabled
             * Thunderbird's native conversation pref. A cloned about:blank
             * browser exists then, but its about:message API is not loaded yet. */
            const expandedArticle = browser.closest(
              'article[aria-expanded="true"]'
            );
            const expandedMessageId = expandedArticle?.dataset?.messageId || "";
            const loadedLatestMessageId =
              conversationView.messages?.at(-1)?.messageId || "";
            shouldRetire =
              typeof browser.contentWindow?.displayMessage !== "function" ||
              browser.hidden ||
              !expandedArticle ||
              !expandedMessageId ||
              (loadedLatestMessageId &&
                loadedLatestMessageId !== expandedMessageId);
          } catch (error) {
            shouldRetire = true;
          }
        }
      }
      if (shouldRetire) {
        retireConversationView(messagePane, conversationView);
      }
    }
    /* Run after Thunderbird's bubble-phase selection handler has rendered its
     * stock reader. This privileged trigger also covers expand/collapse, where
     * the selected header does not change and the public mailTabs event can
     * fail before reaching the extension. */
    scheduleRefresh();
  };
  const onAllMessagesLoaded = () => {
    try {
      if (applyUnreadThreadDefaults(about3Pane)) {
        scheduleRefresh();
      }
    } catch (error) {
      console.error(
        "Outlook Style Companion could not apply unread thread defaults:",
        error
      );
    }
  };
  const onUnload = () => {
    retireConversationGuardsForPane(about3Pane);
    removeConversationSelectionCapture(about3Pane);
  };
  threadTree.addEventListener("select", onSelect, true);
  threadTree.addEventListener("dblclick", onDoubleClick, true);
  about3Pane.addEventListener("allMessagesLoaded", onAllMessagesLoaded);
  about3Pane.addEventListener("unload", onUnload, { once: true });
  state = {
    onAllMessagesLoaded,
    onDoubleClick,
    onSelect,
    onUnload,
    pendingOpen: null,
    pendingOpenTimer: 0,
    refreshConversation,
    refreshTimer: 0,
    threadTree,
  };
  conversationSelectionCaptureState.set(about3Pane, state);
  conversationSelectionCapturePanes.add(about3Pane);
  onAllMessagesLoaded();
  return true;
}

function restoreCurrentNativeSelection(about3Pane, messagePane) {
  try {
    const dbView = about3Pane?.gDBView;
    if (about3Pane?.closed || !dbView || !messagePane?.isConnected) {
      return;
    }
    const selectedIndex = about3Pane.threadTree?.selectedIndex ?? -1;
    if (selectedIndex < 0 || dbView.selection?.count === 0) {
      messagePane.displayMessage();
      return;
    }
    if (dbView.getFlagsAt(selectedIndex) & MSG_VIEW_FLAG_DUMMY) {
      messagePane.displayMessage();
      return;
    }
    if (dbView.numSelected === 1) {
      const selected = dbView.getMsgHdrAt(selectedIndex);
      messagePane.displayMessage(selected.folder.getUriForMsg(selected));
      return;
    }
    messagePane.displayMessages(dbView.getSelectedMsgHdrs());
  } catch (error) {
    /* The view can be replaced while a folder or tab is closing. */
  }
}

function restoreNativeConversationSelection(state) {
  if (state.about3Pane.gDBView !== state.dbView) {
    return;
  }
  restoreCurrentNativeSelection(state.about3Pane, state.messagePane);
}

function displayGuardFallbackMessage(state) {
  const {
    about3Pane,
    dbView,
    fallbackMessage,
    messagePane,
    rootMessage,
    selectedMessage,
    selectedRootOpen,
  } = state;
  try {
    const threadTree = about3Pane.threadTree;
    const selectedIndex = threadTree?.selectedIndex ?? -1;
    const currentThread =
      selectedIndex >= 0
        ? dbView.getThreadContainingIndex(selectedIndex)
        : null;
    if (
      about3Pane.closed ||
      about3Pane.gDBView !== dbView ||
      !messagePane?.isConnected ||
      !about3Pane.gViewWrapper?.showThreaded ||
      selectedIndex < 0 ||
      dbView.selection?.count !== 1 ||
      threadTree.selectedIndices?.length !== 1 ||
      !sameMessage(dbView.getMsgHdrAt(selectedIndex), selectedMessage) ||
      !sameMessage(currentThread?.getRootHdr(), rootMessage) ||
      isExpandedSelectedRoot(
        dbView,
        selectedIndex,
        dbView.getMsgHdrAt(selectedIndex),
        currentThread?.getRootHdr()
      ) !== selectedRootOpen
    ) {
      return;
    }
    const message = fallbackMessage || selectedMessage;
    messagePane.displayMessage(message.folder.getUriForMsg(message));
  } catch (error) {
    /* A selection, folder, or tab can disappear during asynchronous recovery. */
  }
}

function retireGuardedConversationWithFallback(conversationView) {
  const state = conversationGuardState.get(conversationView);
  if (!state) {
    return;
  }
  const shouldRestore = isExpectedConversationSelection(state);
  const { messagePane } = state;
  retireConversationView(messagePane, conversationView);
  if (shouldRestore) {
    restoreNativeConversationSelection(state);
  }
}

function armConversationGuardTimeout(state) {
  if (state.timeoutId) {
    state.about3Pane.clearTimeout(state.timeoutId);
  }
  state.timeoutId = state.about3Pane.setTimeout(() => {
    state.timeoutId = 0;
    if (
      state.disabled ||
      state.ready
    ) {
      return;
    }
    const shouldFallback = isExpectedConversationSelection(state);
    state.disabled = true;
    retireConversationView(state.messagePane, state.conversationView);
    if (shouldFallback) {
      displayGuardFallbackMessage(state);
    }
  }, CONVERSATION_VIEW_LOAD_TIMEOUT_MS);
}

function failGuardedConversation(state, conversationView) {
  if (state.timeoutId) {
    state.about3Pane.clearTimeout(state.timeoutId);
    state.timeoutId = 0;
  }
  state.disabled = true;
  retireConversationView(state.messagePane, conversationView);
  displayGuardFallbackMessage(state);
}

function revealGuardedConversation(state, conversationView) {
  if (state.timeoutId) {
    state.about3Pane.clearTimeout(state.timeoutId);
    state.timeoutId = 0;
  }
  if (state.onBrowserLoad) {
    state.main.removeEventListener("load", state.onBrowserLoad, true);
    state.onBrowserLoad = null;
  }
  /* A collapsed thread may have rendered Thunderbird's legacy snippet
   * summary before the extension event crossed into this privileged API.
   * Remove that reader before showing the verified full conversation. */
  try {
    state.messagePane.multiMessageBrowser.contentWindow.gMessageSummary?.clear();
    state.messagePane.multiMessageBrowser.hidden = true;
  } catch (error) {
    /* The legacy browser can disappear while the mail tab is closing. */
  }
  state.ready = true;
  conversationView.style.removeProperty("visibility");
  conversationView.style.removeProperty("pointer-events");
  conversationView.removeAttribute("aria-hidden");
  conversationView.hidden = false;
}

/**
 * Thunderbird 153 initially opens the newest conversation message from a
 * target-level `load` listener on a freshly inserted about:message browser.
 * When another visible thread row is selected, intercept that listener in the
 * ancestor capture phase and use Thunderbird's own compact-row click handler
 * to create the selected message's browser instead. This prevents the hidden
 * provisional newest message from being requested (and potentially marked
 * read) while preserving the native conversation accordion.
 *
 * @param {HTMLElement} conversationView
 * @param {Event} event
 */
function interceptInitialConversationBrowserLoad(conversationView, event) {
  const state = conversationGuardState.get(conversationView);
  if (
    !state ||
    state.disabled ||
    state.ready ||
    state.activationStarted ||
    !state.targetMessage
  ) {
    return;
  }

  const browser =
    event
      .composedPath?.()
      .find(
        node =>
          node?.matches?.('browser[src="about:message"]') &&
          state.main.contains(node)
      ) || event.target;
  if (
    !browser?.matches?.('browser[src="about:message"]') ||
    !state.main.contains(browser)
  ) {
    return;
  }

  const stopInitialLoad = () => {
    /* The native listener is registered on the browser itself. Stopping at the
     * main-conversation ancestor prevents only that provisional display call;
     * the replacement browser receives its own independent load event. */
    event.stopImmediatePropagation();
  };
  const failClosed = () => {
    stopInitialLoad();
    failGuardedConversation(state, conversationView);
  };

  try {
    if (!isExpectedConversationSelection(state)) {
      failClosed();
      return;
    }

    const loadedMessages = Array.isArray(conversationView.messages)
      ? conversationView.messages
      : [];
    const matchedMessages = matchConversationMessages(
      loadedMessages,
      state.expectedMessages
    );
    const loadedLatest = loadedMessages.at(-1);
    const loadedTarget = matchedMessages
      ? [...matchedMessages].find(([expected]) =>
          sameMessage(expected, state.targetMessage)
        )?.[1]
      : null;
    if (!matchedMessages || !loadedLatest || !loadedTarget) {
      failClosed();
      return;
    }

    /* The selected row is already Thunderbird's newest-first default. Allow
     * its native target listener to load it normally. */
    if (loadedTarget === loadedLatest) {
      return;
    }

    const expandedArticle = browser.closest('article[aria-expanded="true"]');
    const latestMessageId = getMessageId(loadedLatest);
    const targetMessageId = getMessageId(loadedTarget);
    const targetArticle = [
      ...state.main.querySelectorAll('article[aria-expanded="false"]'),
    ].find(
      article =>
        String(article.dataset.messageId || "").trim() === targetMessageId
    );
    if (
      !expandedArticle ||
      !latestMessageId ||
      !targetMessageId ||
      !targetArticle
    ) {
      failClosed();
      return;
    }

    /* TB153 interpolates the Message-ID into a private selector. Exercise that
     * exact lookup before dispatching the click so malformed IDs fail safely. */
    const nativeTargetArticle = state.main.querySelector(
      `article[data-message-id="${targetMessageId}"]`
    );
    if (nativeTargetArticle !== targetArticle) {
      failClosed();
      return;
    }

    stopInitialLoad();
    /* #openMessage expects the current full article to identify the message it
     * is collapsing. Normally MsgLoaded adds this attribute, but that is
     * exactly the provisional load being skipped. The browser is made
     * non-hidden only so the companion's click-safety capture listener permits
     * the native click; the host stays hidden until the target is verified. */
    expandedArticle.dataset.messageId = latestMessageId;
    browser.hidden = false;
    state.activationStarted = true;
    state.targetMessageId = targetMessageId;
    state.loadedTarget = loadedTarget;
    armConversationGuardTimeout(state);
    targetArticle.click();
  } catch (error) {
    failClosed();
    return;
  }
  try {
    const targetArticle = state.main.querySelector(
      `article[data-message-id="${state.targetMessageId}"][aria-expanded="false"]`
    );
    if (targetArticle) {
      failClosed();
    }
  } catch (error) {
    failClosed();
  }
}

function validateGuardedConversation(conversationView) {
  const state = conversationGuardState.get(conversationView);
  if (!state) {
    return;
  }

  if (state.disabled) {
    if (state.timeoutId) {
      state.about3Pane.clearTimeout(state.timeoutId);
      state.timeoutId = 0;
    }
    retireConversationView(state.messagePane, conversationView);
    return;
  }

  if (!isExpectedConversationSelection(state)) {
    if (state.timeoutId) {
      state.about3Pane.clearTimeout(state.timeoutId);
      state.timeoutId = 0;
    }
    state.disabled = true;
    retireConversationView(state.messagePane, conversationView);
    /* Thunderbird may already have reused the host for the newly selected
     * row when the stale observer runs. Recreate that row's stock reader after
     * detaching our generation so a native-pref conversation cannot go blank. */
    restoreCurrentNativeSelection(state.about3Pane, state.messagePane);
    return;
  }

  if (!state.main.childElementCount) {
    return;
  }

  /* Once the selected left-hand row has been opened, leave subsequent manual
   * accordion changes entirely to Thunderbird. The guard remains only to
   * retire this host safely when the tree selection changes. */
  if (state.ready) {
    return;
  }

  const loadedMessages = Array.isArray(conversationView.messages)
    ? conversationView.messages
    : [];
  const expandedArticle = state.main.querySelector(
    'article[aria-expanded="true"]'
  );
  const expandedBrowser = expandedArticle?.querySelector(
    'browser[src="about:message"]'
  );
  const matchedMessages = matchConversationMessages(
    loadedMessages,
    state.expectedMessages
  );
  const containsExpectedThread = Boolean(matchedMessages);
  const loadedLatest = loadedMessages.at(-1);
  if (containsExpectedThread) {
    const latestMessageId = getMessageId(loadedLatest);
    const expandedMessageId = String(
      expandedArticle?.dataset?.messageId || ""
    ).trim();
    let displayedMessage = null;
    try {
      displayedMessage = expandedBrowser?.contentWindow?.gMessage;
    } catch (error) {
      /* A nested reader can be replaced between the DOM query and this read.
       * Its mutation or timeout will retry/fall back without exposing it. */
      return;
    }

    /* The load-capture guard must redirect an older target before Thunderbird
     * calls displayMessage for its newest-first browser. Once a message has
     * started loading it is too late to prevent IMAP or auto-read side effects,
     * so this validation path only accepts the native newest target. */
    if (!state.activationStarted) {
      const loadedTarget = state.targetMessage
        ? [...matchedMessages].find(([expected]) =>
            sameMessage(expected, state.targetMessage)
          )?.[1]
        : loadedLatest;
      if (!loadedTarget) {
        failGuardedConversation(state, conversationView);
        return;
      }

      if (loadedTarget !== loadedLatest) {
        /* Before the initial about:message load event, wait for the capture
         * listener to perform the safe native-row switch. Any evidence that
         * the provisional reader started displaying is a fail-closed state. */
        if (
          displayedMessage ||
          expandedMessageId ||
          (expandedBrowser && !expandedBrowser.hidden)
        ) {
          failGuardedConversation(state, conversationView);
        }
        return;
      }

      if (
        !loadedLatest ||
        !latestMessageId ||
        !expandedBrowser ||
        expandedBrowser.hidden ||
        expandedMessageId !== latestMessageId ||
        !sameMessage(displayedMessage, loadedLatest)
      ) {
        return;
      }
      revealGuardedConversation(state, conversationView);
      return;
    }

    /* The replacement article receives its identity only after its own
     * about:message browser fires MsgLoaded. Reveal nothing until both the DOM
     * stamp and the trusted reader agree with the selected left-hand message. */
    if (
      !expandedBrowser ||
      expandedBrowser.hidden ||
      !expandedMessageId
    ) {
      return;
    }
    if (
      expandedMessageId !== state.targetMessageId ||
      !sameMessage(displayedMessage, state.loadedTarget)
    ) {
      failGuardedConversation(state, conversationView);
      return;
    }
    revealGuardedConversation(state, conversationView);
    return;
  }

  /* A query that does not contain the selected root must never replace the
   * reading pane. The host is still hidden, so fall back before it can paint. */
  /* A failed or incomplete Gloda result must not send a collapsed thread back
   * to Thunderbird's permanently truncated legacy summary. Keep the selected
   * email complete in the ordinary trusted about:message reader instead. */
  failGuardedConversation(state, conversationView);
}

function guardConversationView(
  about3Pane,
  messagePane,
  conversationView,
  selectedMessage,
  expectedMessages,
  fallbackMessage,
  dbView,
  rootMessage,
  selectedRootOpen,
  targetMessage
) {
  const main = conversationView?.shadowRoot?.getElementById(
    "mainConversation"
  );
  if (!main) {
    return false;
  }

  let state = conversationGuardState.get(conversationView);
  if (!state || state.main !== main) {
    removeConversationGuard(conversationView);
    const onBrowserLoad = event => {
      interceptInitialConversationBrowserLoad(conversationView, event);
    };
    main.addEventListener("load", onBrowserLoad, true);
    const observer = new about3Pane.MutationObserver(() => {
      validateGuardedConversation(conversationView);
    });
    observer.observe(main, {
      attributes: true,
      attributeFilter: ["data-message-id", "hidden"],
      childList: true,
      subtree: true,
    });
    state = { main, observer, onBrowserLoad };
    conversationGuardState.set(conversationView, state);
    guardedConversationViews.add(conversationView);
  }
  Object.assign(state, {
    about3Pane,
    conversationView,
    dbView,
    disabled: false,
    activationStarted: false,
    expectedMessages,
    fallbackMessage,
    loadedTarget: null,
    messagePane,
    ready: false,
    rootMessage,
    selectedMessage,
    selectedRootOpen,
    targetMessage,
    targetMessageId: "",
  });
  armConversationGuardTimeout(state);
  /* Do not paint or expose an asynchronous result until it has been matched to
   * the live root selection. Keep the host in layout, however: Thunderbird's
   * one-shot MsgLoaded handler measures the nested body while this guard is
   * active, and display:none would make that measurement zero. */
  conversationView.hidden = false;
  conversationView.style.setProperty("visibility", "hidden");
  conversationView.style.setProperty("pointer-events", "none");
  conversationView.setAttribute("aria-hidden", "true");
  validateGuardedConversation(conversationView);
  return true;
}

async function refreshConversationForPane(about3Pane) {
  if (
    about3Pane.closed ||
    about3Pane.location.href !== "about:3pane"
  ) {
    return false;
  }

  /* Retry until an active vertical mail layout is measurable. Once it
   * succeeds, the versioned preference makes this a permanent no-op. */
  maybeApplyOutlookPaneWidths(about3Pane);

  const dbView = about3Pane.gDBView;
  const threadTree = about3Pane.threadTree;
  const messagePane = about3Pane.document.querySelector("message-pane");
  const selectedIndex = threadTree?.selectedIndex ?? -1;
  const ownedConversationView = messagePane?.querySelector(
    `:scope > conversation-view[${OWNED_CONVERSATION_ATTRIBUTE}="true"]`
  );

  /* Thunderbird's legacy multi-message branch does not clear a native
   * conversation host. Remove the host created for our previous
   * request before a non-target selection can leave it over another
   * reader. Unmarked native hosts remain outside this cleanup. */
  const retireOwnedConversation = () => {
    if (ownedConversationView) {
      retireConversationView(messagePane, ownedConversationView);
    }
  };

  /* Thunderbird represents a collapsed thread as one selected UI row
   * but counts every hidden child in numSelected. Target one visible
   * row instead, so the same conversation UI handles a collapsed
   * parent, an expanded parent, or any selected reply. */
  if (
    !dbView ||
    !messagePane ||
    selectedIndex < 0 ||
    dbView.selection?.count !== 1 ||
    threadTree.selectedIndices?.length !== 1 ||
    !about3Pane.gViewWrapper?.showThreaded
  ) {
    retireOwnedConversation();
    return false;
  }

  if (dbView.getFlagsAt(selectedIndex) & MSG_VIEW_FLAG_DUMMY) {
    retireOwnedConversation();
    return false;
  }

  const thread = dbView.getThreadContainingIndex(selectedIndex);
  if (!thread || thread.numChildren < 2) {
    retireOwnedConversation();
    return false;
  }

  const selectedMessage = dbView.getMsgHdrAt(selectedIndex);
  const rootMessage = thread.getRootHdr();
  if (!selectedMessage || !rootMessage) {
    retireOwnedConversation();
    return false;
  }
  const selectedRootOpen = isExpandedSelectedRoot(
    dbView,
    selectedIndex,
    selectedMessage,
    rootMessage
  );
  /* A collapsed parent represents the conversation as a whole, so its
   * native default remains the newest reply. Once the parent is open,
   * or a visible reply is selected, open that exact left-hand row. */
  const targetMessage =
    selectedRootOpen || !sameMessage(selectedMessage, rootMessage)
      ? selectedMessage
      : null;

  const selectionStillMatches = () => {
    try {
      const currentIndex = threadTree.selectedIndex;
      const currentThread =
        currentIndex >= 0
          ? dbView.getThreadContainingIndex(currentIndex)
          : null;
      return (
        !about3Pane.closed &&
        about3Pane.gDBView === dbView &&
        currentIndex === selectedIndex &&
        dbView.selection?.count === 1 &&
        threadTree.selectedIndices?.length === 1 &&
        sameMessage(dbView.getMsgHdrAt(currentIndex), selectedMessage) &&
        sameMessage(currentThread?.getRootHdr(), rootMessage) &&
        isExpandedSelectedRoot(
          dbView,
          currentIndex,
          dbView.getMsgHdrAt(currentIndex),
          currentThread?.getRootHdr()
        ) === selectedRootOpen
      );
    } catch (error) {
      return false;
    }
  };

  const displaySelectedFallback = () => {
    if (!selectionStillMatches() || !messagePane.isConnected) {
      return;
    }
    try {
      messagePane.displayMessage(
        selectedMessage.folder.getUriForMsg(selectedMessage)
      );
    } catch (error) {
      /* A folder or tab can disappear between revalidation and recovery. */
    }
  };

  const messages = [];
  for (let index = 0; index < thread.numChildren; index++) {
    messages.push(thread.getChildHdrAt(index));
  }

  /* Thunderbird's legacy multi-message summary permanently truncates
   * every message to a short snippet. Prefer its native conversation
   * accordion, which lazy-loads the complete message in an embedded
   * about:message reader when a compact row is activated. Gloda owns
   * that view's thread membership, so do not replace the already-full
   * parent message unless every visible member is safely indexed. */
  let canUseConversationView = false;
  try {
    if (
      Services.prefs.getBoolPref(GLOBAL_INDEXER_PREF, true) &&
      about3Pane.document.getElementById("conversationViewTemplate") &&
      about3Pane.document.getElementById(
        "conversationViewMessageBrowserTemplate"
      )
    ) {
      const { Gloda } = ChromeUtils.importESModule(
        "resource:///modules/gloda/GlodaPublic.sys.mjs"
      );
      canUseConversationView = messages.every(message =>
        Gloda.isMessageIndexed(message)
      );
    }
  } catch (error) {
    /* Unsupported folders and an index being rebuilt are safe
     * fallbacks: keep Thunderbird's full single-message reader. */
    canUseConversationView = false;
  }

  if (!canUseConversationView) {
    retireOwnedConversation();
    /* A collapsed thread would otherwise remain in Thunderbird's
     * permanently truncated legacy summary. When Gloda cannot build
     * a trustworthy conversation, keep the selected message complete
     * in the normal reader instead. */
    displaySelectedFallback();
    return false;
  }

  /* Revalidate after collecting the headers in case the user moved to
   * another message while the event crossed the extension boundary. */
  if (!selectionStillMatches()) {
    retireOwnedConversation();
    return false;
  }

  /* The native accordion is still pref-gated in Thunderbird 153.
   * Enable it only for this synchronous display call and restore the
   * user's preference immediately, so the Companion does not leave a
   * hidden global setting behind or alter unrelated selections. */
  const hadConversationUserValue = Services.prefs.prefHasUserValue(
    CONVERSATION_VIEW_PREF
  );
  const previousConversationValue = Services.prefs.getBoolPref(
    CONVERSATION_VIEW_PREF,
    false
  );
  /* Host identity is the request-generation boundary. Thunderbird 153
   * does not cancel a conversation's async Gloda, snippet, or MsgLoaded
   * callbacks, so never reuse a host from an earlier root selection. */
  retireConversationView(
    messagePane,
    messagePane.querySelector(":scope > conversation-view")
  );
  const shouldRestoreConversationPref = !previousConversationValue;
  if (!previousConversationValue) {
    Services.prefs.setBoolPref(CONVERSATION_VIEW_PREF, true);
  }
  try {
    messagePane.displayMessages(messages);
    const conversationView = messagePane.querySelector(
      ":scope > conversation-view"
    );
    conversationView?.setAttribute(
      OWNED_CONVERSATION_ATTRIBUTE,
      "true"
    );
    if (
      !guardConversationView(
        about3Pane,
        messagePane,
        conversationView,
        selectedMessage,
        messages,
        selectedMessage,
        dbView,
        rootMessage,
        selectedRootOpen,
        targetMessage
      )
    ) {
      retireConversationView(messagePane, conversationView);
      displaySelectedFallback();
      return false;
    }
  } catch (error) {
    /* Native conversation construction spans private Thunderbird DOM and
     * Gloda APIs. If any step changes within the supported branch, never leave
     * the reader blank after retiring the previous generation. */
    const failedConversation = messagePane.querySelector(
      `:scope > conversation-view[${OWNED_CONVERSATION_ATTRIBUTE}="true"], :scope > conversation-view`
    );
    retireConversationView(messagePane, failedConversation);
    displaySelectedFallback();
    console.error(
      "Outlook Style Companion could not create the selected conversation:",
      error
    );
    return false;
  } finally {
    if (shouldRestoreConversationPref) {
      if (hadConversationUserValue) {
        Services.prefs.setBoolPref(
          CONVERSATION_VIEW_PREF,
          previousConversationValue
        );
      } else {
        Services.prefs.clearUserPref(CONVERSATION_VIEW_PREF);
      }
    }
  }
  return true;
}

function getAbout3PaneWindows(window) {
  const windows = new Set();
  const tabmail = window.document?.getElementById("tabmail");
  const tabs = [
    ...(tabmail?.tabInfo || []),
    tabmail?.currentTabInfo,
  ].filter(Boolean);
  for (const tab of tabs) {
    try {
      const about3Pane = tab.chromeBrowser?.contentWindow;
      if (about3Pane?.location?.href === "about:3pane") {
        windows.add(about3Pane);
      }
    } catch (error) {
      /* A tab can disappear while Thunderbird is switching or restoring it. */
    }
  }
  return windows;
}

function getAboutMessageBrowsers(about3Pane) {
  try {
    const browsers = new Set(
      about3Pane.document.querySelectorAll(
        'browser[src="about:message"], #messageBrowser, #multiMessageBrowser'
      )
    );
    const conversationView = about3Pane.document.querySelector(
      "message-pane > conversation-view"
    );
    for (const browser of
      conversationView?.shadowRoot?.querySelectorAll(
        'browser[src="about:message"]'
      ) || []) {
      browsers.add(browser);
    }
    return browsers;
  } catch (error) {
    return new Set();
  }
}

function stampOutlookColorScheme(document, scheme) {
  const root = document?.documentElement;
  if (!root || (scheme !== "light" && scheme !== "dark")) {
    return;
  }
  if (root.getAttribute(OUTLOOK_COLOR_SCHEME_ATTRIBUTE) !== scheme) {
    root.setAttribute(OUTLOOK_COLOR_SCHEME_ATTRIBUTE, scheme);
  }
}

/* Keep the decision cues intentionally conservative. A confirmed meeting that
 * overlaps the invitation is always a conflict; a confirmed meeting ending or
 * starting within half an hour is a back-to-back warning. The local calendar
 * cache is the only data source, so a provider failure never gets reported as
 * a misleading free slot. */
const AVAILABILITY_NEARBY_MINUTES = 30;
const AVAILABILITY_LOOKAROUND_MINUTES = 240;
const AVAILABILITY_MIN_TIMELINE_MINUTES = 180;
const invitationAvailabilityDocumentState = new WeakMap();

const AVAILABILITY_SURFACE_CSS = `
  .outlook-style-availability {
    box-sizing: border-box;
    margin: 16px 0 4px;
    padding: 12px;
    border: 1px solid var(--outlook-divider-strong, #d1d1d1);
    border-radius: 8px;
    background: var(--outlook-surface-alt, #fafafa);
    color: var(--outlook-text, #242424);
    font: 13px/18px var(--fluent-font-family, "Segoe UI", sans-serif);
  }
  .outlook-style-availability-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .outlook-style-availability-title {
    margin: 0;
    color: inherit;
    font: 600 14px/20px var(--fluent-font-family, "Segoe UI", sans-serif);
  }
  .outlook-style-availability-status {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    flex: 0 0 auto;
    font-weight: 600;
    white-space: nowrap;
  }
  .outlook-style-availability-status::before {
    content: "";
    inline-size: 10px;
    block-size: 10px;
    border: 1px solid currentColor;
    border-radius: 50%;
    background: currentColor;
  }
  .outlook-style-availability[data-availability-state="green"] .outlook-style-availability-status {
    color: #107c10;
  }
  .outlook-style-availability[data-availability-state="yellow"] .outlook-style-availability-status {
    color: #8a6100;
  }
  .outlook-style-availability[data-availability-state="red"] .outlook-style-availability-status {
    color: #c50f1f;
  }
  .outlook-style-availability[data-availability-state="unknown"] .outlook-style-availability-status {
    color: var(--outlook-text-muted, #616161);
  }
  .outlook-style-availability-summary {
    margin: 4px 0 0;
    color: var(--outlook-text-secondary, #424242);
  }
  .outlook-style-availability-timeline {
    position: relative;
    block-size: 58px;
    margin-block: 12px 4px;
    border-block: 1px solid var(--outlook-divider, #e0e0e0);
    background:
      repeating-linear-gradient(
        90deg,
        transparent 0,
        transparent calc(25% - 1px),
        var(--outlook-divider, #e0e0e0) calc(25% - 1px),
        var(--outlook-divider, #e0e0e0) 25%
      );
  }
  .outlook-style-availability-block {
    position: absolute;
    box-sizing: border-box;
    overflow: hidden;
    min-inline-size: 4px;
    padding: 2px 5px;
    border-radius: 3px;
    color: #fff;
    font-size: 11px;
    line-height: 14px;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .outlook-style-availability-block[data-kind="target"] {
    inset-block-start: 7px;
    block-size: 20px;
    background: var(--outlook-accent, #0f6cbd);
  }
  .outlook-style-availability-block[data-kind="confirmed"] {
    inset-block-start: 31px;
    block-size: 20px;
    background: #5c5c5c;
  }
  .outlook-style-availability-block[data-relation="overlap"] {
    background: #c50f1f;
  }
  .outlook-style-availability-block[data-relation="nearby"] {
    background: #8a6100;
  }
  .outlook-style-availability-scale {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    color: var(--outlook-text-muted, #616161);
    font-size: 11px;
  }
  .outlook-style-availability-events {
    display: grid;
    gap: 3px;
    margin: 10px 0 0;
    padding: 0;
    list-style: none;
  }
  .outlook-style-availability-events[hidden],
  .outlook-style-availability-timeline[hidden],
  .outlook-style-availability-scale[hidden] {
    display: none;
  }
  .outlook-style-availability-event {
    overflow: hidden;
    color: var(--outlook-text-secondary, #424242);
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .outlook-style-availability-event::before {
    content: "";
    display: inline-block;
    inline-size: 7px;
    block-size: 7px;
    margin-inline-end: 6px;
    border-radius: 50%;
    background: #5c5c5c;
  }
  .outlook-style-availability-event[data-relation="overlap"]::before {
    background: #c50f1f;
  }
  .outlook-style-availability-event[data-relation="nearby"]::before {
    background: #8a6100;
  }
  @media (prefers-color-scheme: dark) {
    .outlook-style-availability[data-availability-state="green"] .outlook-style-availability-status {
      color: #6ccb5f;
    }
    .outlook-style-availability[data-availability-state="yellow"] .outlook-style-availability-status {
      color: #fce100;
    }
    .outlook-style-availability[data-availability-state="red"] .outlook-style-availability-status {
      color: #ff99a4;
    }
    .outlook-style-availability-block[data-kind="confirmed"] {
      background: #adadad;
      color: #242424;
    }
  }
  @media (forced-colors: active) {
    .outlook-style-availability {
      border-color: ButtonBorder;
      background: Canvas;
      color: CanvasText;
    }
    .outlook-style-availability-status,
    .outlook-style-availability[data-availability-state] .outlook-style-availability-status {
      color: CanvasText;
    }
    .outlook-style-availability-block,
    .outlook-style-availability-event::before {
      forced-color-adjust: none;
      border: 1px solid CanvasText;
      background: Highlight;
      color: HighlightText;
    }
  }
`;

function getAvailabilityDateTime(dateTime) {
  if (!dateTime || dateTime.isDate) {
    return null;
  }
  try {
    return dateTime.getInTimezone(cal.dtz.defaultTimezone);
  } catch (error) {
    try {
      return dateTime.clone();
    } catch (cloneError) {
      return null;
    }
  }
}

function offsetAvailabilityDateTime(dateTime, minutes) {
  const result = dateTime?.clone?.();
  if (!result || !Number.isFinite(minutes)) {
    return null;
  }
  try {
    result.addDuration(cal.createDuration(`PT${minutes}M`));
    return result;
  } catch (error) {
    return null;
  }
}

function getAvailabilityBounds(item) {
  const start = getAvailabilityDateTime(item?.startDate);
  const end = getAvailabilityDateTime(item?.endDate);
  if (!start || !end) {
    return null;
  }
  try {
    return end.compare(start) > 0 ? { start, end } : null;
  } catch (error) {
    return null;
  }
}

function getAvailabilitySecondsBetween(later, earlier) {
  try {
    return later.subtractDate(earlier).inSeconds;
  } catch (error) {
    return NaN;
  }
}

function isAvailabilityCalendarEnabled(calendar) {
  try {
    const disabled = calendar?.getProperty("disabled");
    const composite = calendar?.getProperty("calendar-main-in-composite");
    if (disabled === true || disabled === "true") {
      return false;
    }
    return composite !== false && composite !== "false";
  } catch (error) {
    return false;
  }
}

function hasAcceptedParticipation(item, calendar) {
  const status = String(
    item?.status || item?.getProperty?.("STATUS") || ""
  ).toUpperCase();
  const transparency = String(item?.getProperty?.("TRANSP") || "").toUpperCase();
  if (status === "CANCELLED" || transparency === "TRANSPARENT") {
    return false;
  }

  let attendee = null;
  try {
    attendee = cal.itip.getInvitedAttendee(item, calendar);
  } catch (error) {
    /* A non-scheduling calendar can still contain a personal meeting. */
  }
  if (attendee) {
    return String(attendee.participationStatus || "").toUpperCase() === "ACCEPTED";
  }

  const attendees = item?.getAttendees?.() || [];
  if (!attendees.length) {
    return true;
  }

  const organizerId = normalizeAttendeeId(item?.organizer?.id);
  const calendarOrganizerId = normalizeAttendeeId(
    getCalendarOrganizerId(calendar)
  );
  return Boolean(
    organizerId && calendarOrganizerId && organizerId === calendarOrganizerId
  );
}

function isSameAvailabilityItem(candidate, target) {
  if (!candidate || !target) {
    return false;
  }
  if (candidate === target) {
    return true;
  }
  return Boolean(candidate.id && target.id && candidate.id === target.id);
}

function getAvailabilityItemKey(item) {
  const bounds = getAvailabilityBounds(item);
  if (!bounds) {
    return "";
  }
  return [item?.id || "", bounds.start.icalString, bounds.end.icalString].join(
    "|"
  );
}

async function getConfirmedAvailabilityEvents(item) {
  const targetBounds = getAvailabilityBounds(item);
  const rangeStart = targetBounds
    ? offsetAvailabilityDateTime(
        targetBounds.start,
        -AVAILABILITY_LOOKAROUND_MINUTES
      )
    : null;
  const rangeEnd = targetBounds
    ? offsetAvailabilityDateTime(
        targetBounds.end,
        AVAILABILITY_LOOKAROUND_MINUTES
      )
    : null;
  if (!targetBounds || !rangeStart || !rangeEnd) {
    return { events: [], failed: true, targetBounds: null };
  }

  const events = [];
  const eventKeys = new Set();
  let failed = false;
  const filter =
    Ci.calICalendar.ITEM_FILTER_TYPE_EVENT |
    Ci.calICalendar.ITEM_FILTER_CLASS_OCCURRENCES;
  for (const calendar of cal.manager.getCalendars()) {
    if (!isAvailabilityCalendarEnabled(calendar)) {
      continue;
    }
    try {
      for await (const itemBatch of cal.iterate.streamValues(
        calendar.getItems(filter, 0, rangeStart, rangeEnd)
      )) {
        for (const candidate of itemBatch) {
          const bounds = getAvailabilityBounds(candidate);
          if (
            !bounds ||
            isSameAvailabilityItem(candidate, item) ||
            !hasAcceptedParticipation(candidate, calendar)
          ) {
            continue;
          }
          const occurrenceKey = [
            calendar.id || "",
            candidate.hashId || candidate.id || "",
            candidate.recurrenceId?.icalString || bounds.start.icalString,
          ].join("|");
          if (eventKeys.has(occurrenceKey)) {
            continue;
          }
          eventKeys.add(occurrenceKey);
          events.push({
            bounds,
            calendar,
            item: candidate,
          });
        }
      }
    } catch (error) {
      /* Do not turn a temporary provider/cache error into a green light. */
      failed = true;
    }
  }

  return { events, failed, targetBounds };
}

function getAvailabilityRelation(candidateBounds, targetBounds) {
  const startsBeforeTargetEnds =
    getAvailabilitySecondsBetween(candidateBounds.start, targetBounds.end) < 0;
  const endsAfterTargetStarts =
    getAvailabilitySecondsBetween(candidateBounds.end, targetBounds.start) > 0;
  if (startsBeforeTargetEnds && endsAfterTargetStarts) {
    return "overlap";
  }

  const beforeGap = getAvailabilitySecondsBetween(
    targetBounds.start,
    candidateBounds.end
  );
  const afterGap = getAvailabilitySecondsBetween(
    candidateBounds.start,
    targetBounds.end
  );
  if (
    (beforeGap >= 0 && beforeGap <= AVAILABILITY_NEARBY_MINUTES * 60) ||
    (afterGap >= 0 && afterGap <= AVAILABILITY_NEARBY_MINUTES * 60)
  ) {
    return "nearby";
  }
  return "other";
}

function getAvailabilityDisplayEvents(events, targetBounds) {
  const overlapping = [];
  let before = null;
  let after = null;

  for (const event of events) {
    const relation = getAvailabilityRelation(event.bounds, targetBounds);
    if (relation === "overlap") {
      overlapping.push({ ...event, relation });
      continue;
    }
    if (getAvailabilitySecondsBetween(targetBounds.start, event.bounds.end) >= 0) {
      if (
        !before ||
        getAvailabilitySecondsBetween(event.bounds.end, before.bounds.end) > 0
      ) {
        before = { ...event, relation };
      }
      continue;
    }
    if (getAvailabilitySecondsBetween(event.bounds.start, targetBounds.end) >= 0) {
      if (
        !after ||
        getAvailabilitySecondsBetween(after.bounds.start, event.bounds.start) > 0
      ) {
        after = { ...event, relation };
      }
    }
  }

  return [...overlapping, ...(before ? [before] : []), ...(after ? [after] : [])]
    .sort(
      (first, second) =>
        getAvailabilitySecondsBetween(
          first.bounds.start,
          second.bounds.start
        ) || 0
    );
}

function formatAvailabilityTime(document, dateTime) {
  try {
    return cal.dtz.formatter.formatTime(dateTime);
  } catch (error) {
    return new Intl.DateTimeFormat(document.documentElement.lang || undefined, {
      hour: "numeric",
      minute: "2-digit",
    }).format(new Date(dateTime.nativeTime / 1000));
  }
}

function formatAvailabilityInterval(document, bounds) {
  return `${formatAvailabilityTime(document, bounds.start)}–${formatAvailabilityTime(
    document,
    bounds.end
  )}`;
}

function createAvailabilitySurface(document) {
  const surface = document.createElementNS(HTML_NS, "section");
  surface.className = "outlook-style-availability";
  surface.setAttribute("role", "region");
  surface.setAttribute("aria-label", "Your schedule around this event");

  const style = document.createElementNS(HTML_NS, "style");
  style.textContent = AVAILABILITY_SURFACE_CSS;
  surface.appendChild(style);

  const header = document.createElementNS(HTML_NS, "div");
  header.className = "outlook-style-availability-header";
  const title = document.createElementNS(HTML_NS, "h2");
  title.className = "outlook-style-availability-title";
  title.textContent = "Your schedule";
  const status = document.createElementNS(HTML_NS, "span");
  status.className = "outlook-style-availability-status";
  header.append(title, status);

  const summary = document.createElementNS(HTML_NS, "p");
  summary.className = "outlook-style-availability-summary";
  const timeline = document.createElementNS(HTML_NS, "div");
  timeline.className = "outlook-style-availability-timeline";
  const scale = document.createElementNS(HTML_NS, "div");
  scale.className = "outlook-style-availability-scale";
  const start = document.createElementNS(HTML_NS, "span");
  const end = document.createElementNS(HTML_NS, "span");
  scale.append(start, end);
  const events = document.createElementNS(HTML_NS, "ul");
  events.className = "outlook-style-availability-events";
  surface.append(header, summary, timeline, scale, events);
  return { end, events, scale, start, status, summary, surface, timeline };
}

function renderAvailabilitySurface(view, item, result) {
  const { events, failed, targetBounds } = result;
  if (!targetBounds) {
    view.surface.dataset.availabilityState = "unknown";
    view.status.textContent = "Schedule unavailable";
    view.summary.textContent = "This event does not have a timed start and end.";
    view.timeline.hidden = true;
    view.scale.hidden = true;
    view.events.hidden = true;
    return;
  }

  const displayEvents = getAvailabilityDisplayEvents(events, targetBounds);
  const hasConflict = displayEvents.some(event => event.relation === "overlap");
  const hasNearby = displayEvents.some(event => event.relation === "nearby");
  const state = hasConflict ? "red" : failed ? "unknown" : hasNearby ? "yellow" : "green";
  const statusText = {
    green: "Free to accept",
    red: "Conflict",
    unknown: "Schedule unavailable",
    yellow: "Back-to-back",
  }[state];
  const summaryText = {
    green: "No confirmed meeting overlaps this time.",
    red: "A confirmed meeting overlaps this time.",
    unknown:
      "Some calendar data could not be read, so availability cannot be confirmed.",
    yellow: `A confirmed meeting is within ${AVAILABILITY_NEARBY_MINUTES} minutes of this time.`,
  }[state];
  view.surface.dataset.availabilityState = state;
  view.status.textContent = statusText;
  view.summary.textContent = summaryText;
  view.timeline.hidden = false;
  view.scale.hidden = false;
  view.events.hidden = !displayEvents.length;
  view.timeline.replaceChildren();
  view.events.replaceChildren();

  let displayStart = offsetAvailabilityDateTime(
    targetBounds.start,
    -AVAILABILITY_NEARBY_MINUTES
  );
  let displayEnd = offsetAvailabilityDateTime(
    targetBounds.end,
    AVAILABILITY_NEARBY_MINUTES
  );
  for (const event of displayEvents) {
    if (getAvailabilitySecondsBetween(displayStart, event.bounds.start) > 0) {
      displayStart = event.bounds.start.clone();
    }
    if (getAvailabilitySecondsBetween(event.bounds.end, displayEnd) > 0) {
      displayEnd = event.bounds.end.clone();
    }
  }
  const currentDuration = getAvailabilitySecondsBetween(displayEnd, displayStart);
  if (currentDuration < AVAILABILITY_MIN_TIMELINE_MINUTES * 60) {
    const missingMinutes = Math.ceil(
      (AVAILABILITY_MIN_TIMELINE_MINUTES * 60 - currentDuration) / 120
    );
    displayStart = offsetAvailabilityDateTime(displayStart, -missingMinutes);
    displayEnd = offsetAvailabilityDateTime(displayEnd, missingMinutes);
  }
  const duration = Math.max(
    1,
    getAvailabilitySecondsBetween(displayEnd, displayStart)
  );
  const renderBlock = (bounds, label, kind, relation = "") => {
    const left = Math.max(
      0,
      Math.min(
        100,
        (getAvailabilitySecondsBetween(bounds.start, displayStart) / duration) *
          100
      )
    );
    const right = Math.max(
      left,
      Math.min(
        100,
        (getAvailabilitySecondsBetween(bounds.end, displayStart) / duration) *
          100
      )
    );
    const block = view.surface.ownerDocument.createElementNS(HTML_NS, "span");
    block.className = "outlook-style-availability-block";
    block.dataset.kind = kind;
    if (relation) {
      block.dataset.relation = relation;
    }
    block.style.insetInlineStart = `${left}%`;
    block.style.inlineSize = `${Math.max(2, right - left)}%`;
    block.textContent = label;
    block.title = label;
    view.timeline.appendChild(block);
  };

  renderBlock(
    targetBounds,
    `This event · ${formatAvailabilityInterval(view.surface.ownerDocument, targetBounds)}`,
    "target"
  );
  for (const event of displayEvents) {
    const title = String(event.item.title || "Confirmed meeting").trim() || "Confirmed meeting";
    const interval = formatAvailabilityInterval(view.surface.ownerDocument, event.bounds);
    const label = `${title} · ${interval}`;
    renderBlock(event.bounds, label, "confirmed", event.relation);
    const listItem = view.surface.ownerDocument.createElementNS(HTML_NS, "li");
    listItem.className = "outlook-style-availability-event";
    if (event.relation !== "other") {
      listItem.dataset.relation = event.relation;
    }
    const relationship =
      event.relation === "overlap"
        ? "Overlaps"
        : event.relation === "nearby"
          ? "Adjacent"
          : "Nearby";
    listItem.textContent = `${relationship}: ${label}`;
    listItem.title = listItem.textContent;
    view.events.appendChild(listItem);
  }
  view.start.textContent = formatAvailabilityTime(view.surface.ownerDocument, displayStart);
  view.end.textContent = formatAvailabilityTime(view.surface.ownerDocument, displayEnd);
}

function updateAvailabilitySurface(view, item) {
  const key = getAvailabilityItemKey(item);
  if (!key || view.surface.dataset.availabilityItemKey === key) {
    return;
  }
  view.surface.dataset.availabilityItemKey = key;
  const revision = String((Number(view.surface.dataset.availabilityRevision) || 0) + 1);
  view.surface.dataset.availabilityRevision = revision;
  view.surface.dataset.availabilityState = "unknown";
  view.status.textContent = "Checking schedule";
  view.summary.textContent = "Checking confirmed meetings around this time…";
  view.timeline.hidden = true;
  view.scale.hidden = true;
  view.events.hidden = true;
  void getConfirmedAvailabilityEvents(item)
    .then(result => {
      if (
        !view.surface.isConnected ||
        view.surface.dataset.availabilityRevision !== revision
      ) {
        return;
      }
      renderAvailabilitySurface(view, item, result);
    })
    .catch(error => {
      if (
        !view.surface.isConnected ||
        view.surface.dataset.availabilityRevision !== revision
      ) {
        return;
      }
      renderAvailabilitySurface(view, item, {
        events: [],
        failed: true,
        targetBounds: getAvailabilityBounds(item),
      });
    });
}

/* A single read-only item card is shared by invitation previews, summary
 * dialogs, and the first view of existing Calendar/Tasks items. Each native
 * host retains its own response, save, recurrence, and reminder plumbing;
 * choosing Edit simply reveals that host's existing editor. */
const UNIFIED_ITEM_VIEW_CSS = `
  .outlook-style-item-view {
    box-sizing: border-box;
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(280px, 0.72fr);
    align-content: start;
    gap: 18px 24px;
    margin: 12px;
    padding: 20px;
    border: 1px solid var(--outlook-divider-strong, #d1d1d1);
    border-radius: 10px;
    background: var(--outlook-surface, #ffffff);
    color: var(--outlook-text, #242424);
    font: 14px/20px var(--fluent-font-family, "Segoe UI", sans-serif);
  }
  .outlook-style-item-view-header {
    grid-column: 1 / -1;
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 16px;
    padding-block-end: 14px;
    border-block-end: 1px solid var(--outlook-divider, #e0e0e0);
  }
  .outlook-style-item-view-kind {
    margin: 0 0 4px;
    color: var(--outlook-text-secondary, #424242);
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.02em;
    text-transform: uppercase;
  }
  .outlook-style-item-view-title {
    margin: 0;
    color: inherit;
    font: 600 22px/28px var(--fluent-font-family, "Segoe UI", sans-serif);
  }
  .outlook-style-item-view-edit {
    flex: 0 0 auto;
    min-block-size: 32px;
    padding: 5px 14px;
    border: 1px solid var(--outlook-primary-fill, #0f6cbd);
    border-radius: 4px;
    background: var(--outlook-primary-fill, #0f6cbd);
    color: var(--outlook-on-primary, #ffffff);
    font: inherit;
    font-weight: 600;
  }
  .outlook-style-item-view-edit:hover {
    background: var(--outlook-primary-fill-hover, #115ea3);
  }
  .outlook-style-item-view-details,
  .outlook-style-item-view-sidebar {
    min-inline-size: 0;
  }
  .outlook-style-item-view-fields {
    display: grid;
    grid-template-columns: max-content minmax(0, 1fr);
    gap: 7px 14px;
    margin: 0;
  }
  .outlook-style-item-view-fields dt {
    color: var(--outlook-text-secondary, #424242);
    font-weight: 600;
  }
  .outlook-style-item-view-fields dd {
    min-inline-size: 0;
    margin: 0;
    overflow-wrap: anywhere;
  }
  .outlook-style-item-view-section {
    margin-block-start: 18px;
  }
  .outlook-style-item-view-section h3 {
    margin: 0 0 8px;
    color: inherit;
    font: 600 15px/20px var(--fluent-font-family, "Segoe UI", sans-serif);
  }
  .outlook-style-item-view-description {
    max-block-size: 340px;
    margin: 0;
    padding: 12px;
    overflow: auto;
    border: 1px solid var(--outlook-divider-strong, #d1d1d1);
    border-radius: 4px;
    background: var(--outlook-surface-alt, #fafafa);
    color: inherit;
    font: inherit;
    white-space: pre-wrap;
  }
  .outlook-style-item-view-description a {
    color: var(--outlook-accent, #0f6cbd);
    cursor: pointer;
    overflow-wrap: anywhere;
    text-decoration: underline;
  }
  .outlook-style-item-view-attendees {
    display: grid;
    gap: 5px;
    max-block-size: 310px;
    margin: 0;
    padding: 10px;
    overflow: auto;
    border: 1px solid var(--outlook-divider-strong, #d1d1d1);
    border-radius: 4px;
    background: var(--outlook-surface-alt, #fafafa);
    list-style: none;
  }
  .outlook-style-item-view-attendee {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .outlook-style-item-view-attendee::before {
    content: "●";
    margin-inline-end: 7px;
    color: var(--outlook-accent, #0f6cbd);
  }
  .outlook-style-item-view .outlook-style-availability {
    margin: 0;
  }
  .outlook-style-item-view-single {
    grid-template-columns: minmax(0, 1fr);
  }
  .outlook-style-item-view-single .outlook-style-item-view-sidebar {
    display: none;
  }
  @media (max-width: 780px) {
    .outlook-style-item-view {
      grid-template-columns: minmax(0, 1fr);
      margin: 8px;
      padding: 14px;
    }
    .outlook-style-item-view-sidebar {
      grid-row: auto;
    }
    .outlook-style-item-view-header {
      align-items: stretch;
      flex-direction: column;
    }
    .outlook-style-item-view-edit {
      align-self: flex-start;
    }
  }
  @media (forced-colors: active) {
    .outlook-style-item-view,
    .outlook-style-item-view-description,
    .outlook-style-item-view-attendees {
      border-color: ButtonBorder;
      background: Canvas;
      color: CanvasText;
    }
    .outlook-style-item-view-description a {
      color: LinkText;
    }
    .outlook-style-item-view-edit {
      border-color: Highlight;
      background: Highlight;
      color: HighlightText;
    }
  }
`;

function getUnifiedItemDateRange(item) {
  const isTask = item?.isTodo?.();
  const start = isTask ? item.entryDate : item?.startDate;
  let end = isTask ? item.dueDate : item?.endDate;
  if (!isTask && end?.isDate) {
    end = end.clone();
    end.day--;
  }
  const format = value => {
    if (!value) {
      return "Not set";
    }
    try {
      if (value.isDate) {
        return cal.dtz.formatter.formatDateLong(value);
      }
      return cal.dtz.formatter.formatDateTime(
        value.getInTimezone(cal.dtz.defaultTimezone)
      );
    } catch (error) {
      return value.icalString || "Not set";
    }
  };
  if (!start && !end) {
    return "Not scheduled";
  }
  if (!start) {
    return `Due ${format(end)}`;
  }
  if (!end) {
    return `Starts ${format(start)}`;
  }
  if (start?.isDate && start.compare(end) === 0) {
    return `${format(start)} (all day)`;
  }
  return `${format(start)} – ${format(end)}`;
}

function getUnifiedItemStatus(item) {
  return String(item?.status || item?.getProperty?.("STATUS") || "").trim();
}

function getNativeEditorRequestKey(item) {
  return [
    item?.calendar?.id || "",
    item?.id || "",
    item?.recurrenceId?.icalString || "",
  ].join("|");
}

const GO_SHORTCUT_PATTERN =
  /\bgo\/[a-z0-9](?:[a-z0-9._~/-]*[a-z0-9_~/-])?/giu;

function linkifyGoShortcuts(document, fragment) {
  const textNodes = [];
  const walker = document.createTreeWalker(
    fragment,
    document.defaultView.NodeFilter.SHOW_TEXT
  );
  while (walker.nextNode()) {
    if (!walker.currentNode.parentElement?.closest?.("a")) {
      textNodes.push(walker.currentNode);
    }
  }

  for (const textNode of textNodes) {
    const text = textNode.data;
    const matches = [...text.matchAll(GO_SHORTCUT_PATTERN)];
    if (!matches.length) {
      continue;
    }
    const replacement = document.createDocumentFragment();
    let offset = 0;
    for (const match of matches) {
      replacement.append(text.slice(offset, match.index));
      const anchor = document.createElementNS(HTML_NS, "a");
      anchor.href = `http:${"/".repeat(2)}${match[0]}`;
      anchor.textContent = match[0];
      replacement.appendChild(anchor);
      offset = match.index + match[0].length;
    }
    replacement.append(text.slice(offset));
    textNode.replaceWith(replacement);
  }
}

function isEmailLink(href) {
  return /^mailto:/i.test(href);
}

function getEmailAddressFromLink(href) {
  const addressStart = "mailto:".length;
  const queryStart = href.indexOf("?");
  let address = href.slice(
    addressStart,
    queryStart > addressStart ? queryStart : undefined
  );
  try {
    address = Services.textToSubURI.unEscapeURIForUI(address);
  } catch (error) {
    /* Keep the encoded address if it cannot be decoded for display. */
  }
  return address;
}

function openUnifiedItemLink(href, event) {
  if (isEmailLink(href)) {
    MailServices.compose.OpenComposeWindowWithURI(
      null,
      Services.io.newURI(href),
      null,
      event.shiftKey
        ? Ci.nsIMsgCompFormat.OppositeOfDefault
        : Ci.nsIMsgCompFormat.Default
    );
    return;
  }
  openLinkExternally(href, { addToHistory: false });
}

function copyUnifiedItemText(value) {
  if (!value) {
    return;
  }
  Cc["@mozilla.org/widget/clipboardhelper;1"]
    .getService(Ci.nsIClipboardHelper)
    .copyString(value);
}

function getUnifiedItemCommandHost(document) {
  try {
    const opener = document.defaultView?.opener?.top;
    if (opener && !opener.closed) {
      return opener;
    }
  } catch (error) {
    /* A closing or remote opener may no longer be accessible. */
  }
  return (
    Services.wm.getMostRecentWindow("mail:3pane") ||
    Services.wm.getMostRecentWindow("mail:messageWindow")
  );
}

function getUnifiedItemDescriptionMenu(document) {
  const current = unifiedItemDescriptionMenuState.get(document);
  if (current?.popup.isConnected) {
    return current;
  }
  if (!document.createXULElement || !document.body) {
    return null;
  }

  const popup = document.createXULElement("menupopup");
  popup.id = "outlook-style-item-description-context";
  const createItem = (id, label) => {
    const item = document.createXULElement("menuitem");
    item.id = `outlook-style-item-description-${id}`;
    item.setAttribute("label", label);
    return item;
  };
  const open = createItem("open", "Open Link In Browser");
  const addAddress = createItem("add-address", "Add to Address Book…");
  const compose = createItem("compose", "Compose Message To");
  const copyTarget = createItem("copy-target", "Copy Link Location");
  const save = createItem("save", "Save Link As…");
  const targetSeparator = document.createXULElement("menuseparator");
  const copySelection = createItem("copy", "Copy");
  const selectAll = createItem("select-all", "Select All");
  popup.append(
    open,
    addAddress,
    compose,
    copyTarget,
    save,
    targetSeparator,
    copySelection,
    selectAll
  );
  document.body.appendChild(popup);

  const state = {
    addAddress,
    anchor: null,
    compose,
    copySelection,
    copyTarget,
    description: null,
    open,
    popup,
    save,
    selectAll,
    targetSeparator,
  };
  const onCommand = event => {
    const anchor = state.anchor;
    const href = anchor?.href || "";
    try {
      if (event.target === open && href) {
        openUnifiedItemLink(href, event);
      } else if (event.target === addAddress && href) {
        getUnifiedItemCommandHost(document)?.toAddressBook?.([
          "cmd_newCard",
          getEmailAddressFromLink(href),
        ]);
      } else if (event.target === compose && href) {
        openUnifiedItemLink(href, event);
      } else if (event.target === copyTarget && href) {
        copyUnifiedItemText(
          isEmailLink(href) ? getEmailAddressFromLink(href) : href
        );
      } else if (event.target === save && href) {
        getUnifiedItemCommandHost(document)?.saveURL?.(
          href,
          null,
          anchor.textContent,
          null,
          true,
          false,
          null,
          null,
          document,
          null,
          Services.scriptSecurityManager.getSystemPrincipal()
        );
      } else if (event.target === copySelection) {
        copyUnifiedItemText(document.defaultView?.getSelection?.().toString());
      } else if (event.target === selectAll && state.description) {
        const selection = document.defaultView?.getSelection?.();
        const range = document.createRange();
        range.selectNodeContents(state.description);
        selection?.removeAllRanges();
        selection?.addRange(range);
      }
    } catch (error) {
      console.error(
        "Outlook Style Companion could not run a calendar description command:",
        error
      );
    }
  };
  const onPopupHidden = () => {
    state.anchor = null;
    state.description = null;
  };
  const cleanup = () => {
    popup.removeEventListener("command", onCommand);
    popup.removeEventListener("popuphidden", onPopupHidden);
    popup.remove();
    unifiedItemDescriptionMenuState.delete(document);
  };
  popup.addEventListener("command", onCommand);
  popup.addEventListener("popuphidden", onPopupHidden);
  document.defaultView?.addEventListener("unload", cleanup, { once: true });
  unifiedItemDescriptionMenuState.set(document, state);
  return state;
}

function openUnifiedItemDescriptionMenu(document, description, event) {
  const state = getUnifiedItemDescriptionMenu(document);
  if (!state || !event.isTrusted) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  const anchor = event.target?.closest?.("a[href]") || null;
  const email = Boolean(anchor && isEmailLink(anchor.href));
  const webLink = Boolean(anchor && !email);
  const commandHost = getUnifiedItemCommandHost(document);
  const selectionText = document.defaultView?.getSelection?.().toString() || "";
  state.anchor = anchor;
  state.description = description;
  state.open.hidden = !webLink;
  state.addAddress.hidden = !email || !commandHost?.toAddressBook;
  state.compose.hidden = !email;
  state.copyTarget.hidden = !anchor;
  state.save.hidden = !webLink || !commandHost?.saveURL;
  state.targetSeparator.hidden = !anchor;
  state.copySelection.hidden = !selectionText;
  state.copyTarget.setAttribute(
    "label",
    email ? "Copy Email Address" : "Copy Link Location"
  );
  if (event.screenX || event.screenY) {
    state.popup.openPopupAtScreen(
      event.screenX,
      event.screenY,
      true,
      event
    );
  } else {
    state.popup.openPopup(anchor, {
      position: "after_start",
      triggerEvent: event,
    });
  }
}

function createUnifiedItemDescription(document, item, description) {
  const fragment = cal.view.textToHtmlDocumentFragment(
    description,
    document,
    item?.descriptionHTML
  );
  linkifyGoShortcuts(document, fragment);

  for (const anchor of fragment.querySelectorAll("a[href]")) {
    anchor.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      if (!event.isTrusted || event.button !== 0) {
        return;
      }
      try {
        openUnifiedItemLink(anchor.href, event);
      } catch (error) {
        console.error(
          "Outlook Style Companion could not open a calendar description link:",
          error
        );
      }
    });
  }
  return fragment;
}

function createUnifiedItemView(document, item, options = {}) {
  const root = document.createElementNS(HTML_NS, "article");
  root.className = "outlook-style-item-view";
  root.setAttribute("role", "region");
  root.setAttribute("aria-label", "Event details");
  root.dataset.outlookItemKey = getAvailabilityItemKey(item);

  const style = document.createElementNS(HTML_NS, "style");
  style.textContent = UNIFIED_ITEM_VIEW_CSS;
  root.appendChild(style);

  const header = document.createElementNS(HTML_NS, "header");
  header.className = "outlook-style-item-view-header";
  const heading = document.createElementNS(HTML_NS, "div");
  const kind = document.createElementNS(HTML_NS, "p");
  kind.className = "outlook-style-item-view-kind";
  kind.textContent = item?.isTodo?.() ? "Task" : "Event";
  const title = document.createElementNS(HTML_NS, "h1");
  title.className = "outlook-style-item-view-title";
  title.textContent = String(item?.title || "Untitled").trim() || "Untitled";
  heading.append(kind, title);
  header.appendChild(heading);
  if (typeof options.onEdit === "function") {
    const edit = document.createElementNS(HTML_NS, "button");
    edit.type = "button";
    edit.className = "outlook-style-item-view-edit";
    edit.textContent = options.editLabel || "Edit";
    edit.addEventListener("click", options.onEdit);
    header.appendChild(edit);
  }

  const details = document.createElementNS(HTML_NS, "div");
  details.className = "outlook-style-item-view-details";
  const fields = document.createElementNS(HTML_NS, "dl");
  fields.className = "outlook-style-item-view-fields";
  const appendField = (label, value) => {
    if (!value) {
      return;
    }
    const term = document.createElementNS(HTML_NS, "dt");
    term.textContent = label;
    const definition = document.createElementNS(HTML_NS, "dd");
    definition.textContent = value;
    fields.append(term, definition);
  };
  appendField("When", getUnifiedItemDateRange(item));
  appendField("Location", item?.getProperty?.("LOCATION")?.trim());
  const organizer = item?.organizer;
  appendField(
    "Organizer",
    organizer
      ? organizer.commonName || cal.email.removeMailTo(organizer.id || "")
      : ""
  );
  appendField("Status", getUnifiedItemStatus(item));
  details.appendChild(fields);

  const description = String(item?.descriptionText || "").trim();
  if (description) {
    const descriptionSection = document.createElementNS(HTML_NS, "section");
    descriptionSection.className = "outlook-style-item-view-section";
    const heading = document.createElementNS(HTML_NS, "h3");
    heading.textContent = "Description";
    const content = document.createElementNS(HTML_NS, "div");
    content.className = "outlook-style-item-view-description";
    content.appendChild(
      createUnifiedItemDescription(document, item, description)
    );
    content.addEventListener("contextmenu", event => {
      openUnifiedItemDescriptionMenu(document, content, event);
    });
    descriptionSection.append(heading, content);
    details.appendChild(descriptionSection);
  }

  const sidebar = document.createElementNS(HTML_NS, "aside");
  sidebar.className = "outlook-style-item-view-sidebar";
  if (item?.isEvent?.() && getAvailabilityBounds(item)) {
    const availability = createAvailabilitySurface(document);
    sidebar.appendChild(availability.surface);
    updateAvailabilitySurface(availability, item);
  }
  const attendees = item?.getAttendees?.() || [];
  if (attendees.length) {
    const attendeesSection = document.createElementNS(HTML_NS, "section");
    attendeesSection.className = "outlook-style-item-view-section";
    const heading = document.createElementNS(HTML_NS, "h3");
    heading.textContent = "Attendees";
    const list = document.createElementNS(HTML_NS, "ul");
    list.className = "outlook-style-item-view-attendees";
    for (const attendee of attendees) {
      const entry = document.createElementNS(HTML_NS, "li");
      entry.className = "outlook-style-item-view-attendee";
      const email = cal.email.removeMailTo(attendee.id || "");
      entry.textContent = attendee.commonName
        ? `${attendee.commonName} <${email}>`
        : email;
      entry.title = entry.textContent;
      list.appendChild(entry);
    }
    attendeesSection.append(heading, list);
    sidebar.appendChild(attendeesSection);
  }

  if (!sidebar.childElementCount) {
    root.classList.add("outlook-style-item-view-single");
  }

  root.append(header, details, sidebar);
  return root;
}

function installInvitationAvailabilitySurface(document) {
  if (invitationAvailabilityDocumentState.has(document)) {
    invitationAvailabilityDocumentState.get(document).schedule();
    return;
  }
  const display = document.getElementById("calendarInvitationDisplay");
  const legacyAnchor = document.getElementById("mail-notification-top");
  const legacyBar = document.getElementById("imip-bar");
  if (!display && !legacyAnchor) {
    return;
  }

  const getLegacyInvitationItem = () => {
    if (legacyBar?.collapsed) {
      return null;
    }
    try {
      const itipItem =
        document.defaultView?.calImipBar?.itipItem ||
        document.defaultView?.calImipBar?.loadingItipItem;
      return itipItem?.getItemList?.()[0] || null;
    } catch (error) {
      return null;
    }
  };

  const removeLegacySurface = () => {
    document
      .getElementById("outlook-style-legacy-item-view")
      ?.remove();
  };

  let legacyDetails = null;
  let legacyDetailsWasHidden = false;
  let invitationWrapper = null;
  let invitationWrapperWasHidden = false;
  const restoreLegacyDetails = () => {
    if (legacyDetails) {
      legacyDetails.toggleAttribute("hidden", legacyDetailsWasHidden);
    }
    legacyDetails = null;
  };
  const hideLegacyDetails = () => {
    const nextDetails = messagePane?.contentDocument?.getElementById(
      "imipHTMLDetails"
    );
    if (nextDetails === legacyDetails) {
      return;
    }
    restoreLegacyDetails();
    if (nextDetails) {
      legacyDetails = nextDetails;
      legacyDetailsWasHidden = nextDetails.hasAttribute("hidden");
      nextDetails.setAttribute("hidden", "hidden");
    }
  };
  const restoreInvitationWrapper = () => {
    if (invitationWrapper) {
      invitationWrapper.toggleAttribute("hidden", invitationWrapperWasHidden);
    }
    invitationWrapper = null;
  };

  let frame = 0;
  const refresh = () => {
    frame = 0;
    const invitation = display?.querySelector("calendar-invitation-panel");
    const shadowRoot = invitation?.shadowRoot;
    const item = invitation?.item;
    if (invitation && shadowRoot && item) {
      removeLegacySurface();
      restoreLegacyDetails();
      let view = shadowRoot.querySelector(":scope > .outlook-style-item-view");
      if (!view) {
        view = createUnifiedItemView(document, item, {
          onEdit: () => document.defaultView?.calImipBar?.executeAction?.("X-SHOWDETAILS"),
          editLabel: "Details",
        });
        const footer = shadowRoot.getElementById("footer");
        if (footer) {
          footer.before(view);
        } else {
          shadowRoot.appendChild(view);
        }
      }
      const nextWrapper = shadowRoot.querySelector(
        ".calendar-invitation-panel-wrapper"
      );
      if (nextWrapper && nextWrapper !== invitationWrapper) {
        restoreInvitationWrapper();
        invitationWrapper = nextWrapper;
        invitationWrapperWasHidden = nextWrapper.hasAttribute("hidden");
        nextWrapper.setAttribute("hidden", "hidden");
      }
      return;
    }

    restoreInvitationWrapper();

    const legacyItem = getLegacyInvitationItem();
    if (!legacyAnchor || !legacyItem) {
      removeLegacySurface();
      restoreLegacyDetails();
      return;
    }
    let view = document.getElementById("outlook-style-legacy-item-view");
    if (!view) {
      view = createUnifiedItemView(document, legacyItem, {
        onEdit: () => document.defaultView?.calImipBar?.executeAction?.("X-SHOWDETAILS"),
        editLabel: "Details",
      });
      view.id = "outlook-style-legacy-item-view";
      legacyAnchor.after(view);
    }
    hideLegacyDetails();
  };
  const schedule = () => {
    const documentWindow = document.defaultView;
    if (frame || !documentWindow || documentWindow.closed) {
      return;
    }
    frame = documentWindow.requestAnimationFrame(refresh);
  };
  const observer = new document.defaultView.MutationObserver(schedule);
  if (display) {
    observer.observe(display, { childList: true, subtree: true });
  }
  if (legacyBar) {
    observer.observe(legacyBar, {
      attributes: true,
      attributeFilter: ["collapsed", "hidden"],
      childList: true,
      subtree: true,
    });
  }
  const messagePane = document.getElementById("messagepane");
  messagePane?.addEventListener("load", schedule, true);
  const cleanup = () => {
    observer.disconnect();
    messagePane?.removeEventListener("load", schedule, true);
    restoreLegacyDetails();
    restoreInvitationWrapper();
    if (frame) {
      document.defaultView?.cancelAnimationFrame(frame);
      frame = 0;
    }
    invitationAvailabilityDocumentState.delete(document);
  };
  invitationAvailabilityDocumentState.set(document, { schedule });
  document.defaultView.addEventListener("unload", cleanup, { once: true });
  schedule();
}

function installEventSummaryAvailabilitySurface(window, cleanups) {
  const document = window.document;
  let retryTimer = 0;
  let disposed = false;
  const install = () => {
    retryTimer = 0;
    if (disposed || window.closed) {
      return;
    }
    const item = window.calendarItem;
    const summary = document.getElementById("calendar-item-summary");
    if (!item || !summary?.isConnected) {
      retryTimer = window.setTimeout(install, 100);
      return;
    }
    let view = document.getElementById("outlook-style-summary-item-view");
    if (!view) {
      const canEdit = !window.isInvitation && !summary.readOnly;
      view = createUnifiedItemView(document, item, {
        onEdit: canEdit
          ? () => {
              nativeEditorRequestKeys.add(
                getNativeEditorRequestKey(window.calendarItem)
              );
              window.useEditDialog?.(window.calendarItem);
            }
          : null,
      });
      view.id = "outlook-style-summary-item-view";
      summary.before(view);
      const wasHidden = summary.hasAttribute("hidden");
      const footer = document.getElementById(
        "calendar-summary-dialog-custom-button-footer"
      );
      const footerWasHidden = footer?.hasAttribute("hidden") || false;
      summary.setAttribute("hidden", "hidden");
      footer?.setAttribute("hidden", "hidden");
      cleanups.push(() => {
        summary.toggleAttribute("hidden", wasHidden);
        footer?.toggleAttribute("hidden", footerWasHidden);
        view.remove();
      });
    }
  };
  cleanups.push(() => {
    disposed = true;
    if (retryTimer) {
      window.clearTimeout(retryTimer);
    }
  });
  install();
}

function installCalendarCreationEnhancements(window) {
  const document = window.document;
  const controller = window.calendarViewController;
  const originalCreateNewEvent = controller?.createNewEvent;
  let toolbarObserver = null;
  let popup = null;
  let popupShowing = false;
  const dropdownButtons = new Set();

  for (const staleButton of document.querySelectorAll(
    ".outlook-style-new-item-dropdown"
  )) {
    staleButton.remove();
  }
  for (const staleContainer of document.querySelectorAll(
    ".outlook-style-new-item-split"
  )) {
    staleContainer.classList.remove("outlook-style-new-item-split");
  }

  /* Thunderbird normally commits a timed event immediately when its range is
   * swept out in a day or week view. Keep the selected range, but send that
   * one path through the native New Event dialog so the item is not created
   * until the user has had a chance to add its details. */
  let createNewEventWrapper = null;
  if (
    controller &&
    typeof originalCreateNewEvent === "function" &&
    typeof window.createEventWithDialog === "function"
  ) {
    createNewEventWrapper = function (
      calendar,
      startTime,
      endTime,
      forceAllDay
    ) {
      if (
        startTime &&
        endTime &&
        !startTime.isDate &&
        !endTime.isDate
      ) {
        return window.createEventWithDialog(
          calendar,
          startTime,
          endTime,
          null,
          null,
          forceAllDay
        );
      }
      return originalCreateNewEvent.apply(this, arguments);
    };
    controller.createNewEvent = createNewEventWrapper;
  }

  const popupSet =
    document.getElementById("mainPopupSet") ||
    document.getElementById("calendar-popupset");
  if (popupSet) {
    document.getElementById("outlook-style-new-item-popup")?.remove();
    popup = document.createXULElement("menupopup");
    popup.id = "outlook-style-new-item-popup";
    popup.className = "outlook-style-new-item-popup";

    const newMessage = document.createXULElement("menuitem");
    newMessage.id = "outlook-style-new-message-menuitem";
    newMessage.setAttribute("label", "New Message");
    newMessage.setAttribute("command", "cmd_newMessage");

    const newEvent = document.createXULElement("menuitem");
    newEvent.id = "outlook-style-new-event-menuitem";
    newEvent.setAttribute("label", "New Event");
    newEvent.setAttribute("command", "calendar_new_event_command");

    popup.append(newMessage, newEvent);
    popupSet.appendChild(popup);
  }

  const syncPopupState = () => {
    popupShowing = true;
    for (const button of dropdownButtons) {
      button.setAttribute("aria-expanded", "true");
    }
  };
  const clearPopupState = () => {
    popupShowing = false;
    for (const button of dropdownButtons) {
      button.setAttribute("aria-expanded", "false");
    }
  };
  popup?.addEventListener("popupshown", syncPopupState);
  popup?.addEventListener("popuphidden", clearPopupState);

  const decorateWriteButtons = () => {
    if (!popup) {
      return;
    }
    for (const liveContent of document.querySelectorAll(
      '#unifiedToolbarContent li[item-id="write-message"] > .live-content'
    )) {
      liveContent.classList.add("outlook-style-new-item-split");
      if (
        liveContent.querySelector(
          ":scope > .outlook-style-new-item-dropdown"
        )
      ) {
        continue;
      }
      const button = document.createElementNS(HTML_NS, "button");
      button.type = "button";
      button.className = "outlook-style-new-item-dropdown";
      button.title = "Create a new message or event";
      button.setAttribute("aria-label", "More new item options");
      button.setAttribute("aria-haspopup", "menu");
      button.setAttribute("aria-expanded", "false");
      button.addEventListener("click", event => {
        event.preventDefault();
        event.stopPropagation();
        if (popupShowing || popup.state === "open") {
          popup.hidePopup();
        } else {
          popup.openPopup(button, {
            position: "after_end",
            triggerEvent: event,
          });
        }
      });
      liveContent.appendChild(button);
      dropdownButtons.add(button);
    }
    for (const button of [...dropdownButtons]) {
      if (!button.isConnected) {
        dropdownButtons.delete(button);
      }
    }
  };

  const unifiedToolbar = document.querySelector("unified-toolbar");
  if (unifiedToolbar) {
    toolbarObserver = new window.MutationObserver(decorateWriteButtons);
    toolbarObserver.observe(unifiedToolbar, {
      childList: true,
      subtree: true,
    });
    decorateWriteButtons();
  }

  return () => {
    toolbarObserver?.disconnect();
    popup?.removeEventListener("popupshown", syncPopupState);
    popup?.removeEventListener("popuphidden", clearPopupState);
    popup?.hidePopup?.();
    popup?.remove();
    for (const button of dropdownButtons) {
      button.remove();
    }
    for (const liveContent of document.querySelectorAll(
      ".outlook-style-new-item-split"
    )) {
      liveContent.classList.remove("outlook-style-new-item-split");
    }
    dropdownButtons.clear();
    if (
      createNewEventWrapper &&
      controller.createNewEvent === createNewEventWrapper
    ) {
      controller.createNewEvent = originalCreateNewEvent;
    }
  };
}

function installFolderPaneNewItemButton(about3Pane) {
  const document = about3Pane.document;
  const writeButton = document.getElementById("folderPaneWriteMessage");
  if (!writeButton?.parentElement) {
    return null;
  }
  document.getElementById("outlook-style-folder-new-item-dropdown")?.remove();
  document.getElementById("outlook-style-folder-new-item-popup")?.remove();

  const dropdown = document.createElementNS(HTML_NS, "button");
  dropdown.id = "outlook-style-folder-new-item-dropdown";
  dropdown.type = "button";
  dropdown.title = "Create a new message or event";
  dropdown.setAttribute("aria-label", "More new item options");
  dropdown.setAttribute("aria-haspopup", "menu");
  dropdown.setAttribute("aria-expanded", "false");

  const popup = document.createXULElement("menupopup");
  popup.id = "outlook-style-folder-new-item-popup";
  popup.className = "outlook-style-new-item-popup";
  const newMessage = document.createXULElement("menuitem");
  newMessage.setAttribute("label", "New Message");
  const newEvent = document.createXULElement("menuitem");
  newEvent.setAttribute("label", "New Event");
  popup.append(newMessage, newEvent);

  const onDropdown = event => {
    event.preventDefault();
    event.stopPropagation();
    if (popup.state === "open" || popup.state === "showing") {
      popup.hidePopup();
    } else {
      popup.openPopup(dropdown, {
        position: "after_start",
        triggerEvent: event,
      });
    }
  };
  const onPopupShown = () =>
    dropdown.setAttribute("aria-expanded", "true");
  const onPopupHidden = () =>
    dropdown.setAttribute("aria-expanded", "false");
  const onNewMessage = event => about3Pane.top.MsgNewMessage(event);
  const onNewEvent = () =>
    about3Pane.top.goDoCommand("calendar_new_event_command");

  dropdown.addEventListener("click", onDropdown);
  popup.addEventListener("popupshown", onPopupShown);
  popup.addEventListener("popuphidden", onPopupHidden);
  newMessage.addEventListener("command", onNewMessage);
  newEvent.addEventListener("command", onNewEvent);
  writeButton.setAttribute("data-outlook-style-split-button", "true");
  writeButton.before(dropdown);
  document.body.appendChild(popup);

  return () => {
    dropdown.removeEventListener("click", onDropdown);
    popup.removeEventListener("popupshown", onPopupShown);
    popup.removeEventListener("popuphidden", onPopupHidden);
    newMessage.removeEventListener("command", onNewMessage);
    newEvent.removeEventListener("command", onNewEvent);
    popup.hidePopup?.();
    popup.remove();
    dropdown.remove();
    writeButton.removeAttribute("data-outlook-style-split-button");
  };
}

function installMainPaneSchemeBridge(window) {
  if (mainPaneSchemeState.has(window)) {
    return true;
  }
  const tabmail = window.document?.getElementById("tabmail");
  if (!tabmail) {
    return false;
  }

  const state = {
    schemeFrame: 0,
    retryTimer: 0,
    shuttingDown: false,
    trackedDocuments: new Set(),
    browserListeners: new Map(),
    paneListeners: new Map(),
    conversationListeners: new Map(),
    browserDocuments: new Map(),
    folderPaneCreationCleanups: new Map(),
    removeCalendarCreationEnhancements: null,
  };

  const calendarCreationCleanup = runEnhancementStep(
    "calendar creation controls",
    () => installCalendarCreationEnhancements(window)
  );
  state.removeCalendarCreationEnhancements =
    typeof calendarCreationCleanup === "function"
      ? calendarCreationCleanup
      : null;

  const stampDocument = (document, scheme) => {
    try {
      stampOutlookColorScheme(document, scheme);
      installInvitationAvailabilitySurface(document);
      if (document?.documentElement) {
        state.trackedDocuments.add(document);
      }
    } catch (error) {
      /* Ignore a document that is being replaced during navigation. */
    }
  };

  const untrackBrowserDocument = browser => {
    const document = state.browserDocuments.get(browser);
    if (!document) {
      return;
    }
    try {
      document.documentElement?.removeAttribute(
        OUTLOOK_COLOR_SCHEME_ATTRIBUTE
      );
    } catch (error) {
      /* A replaced about:message document may already be destroyed. */
    }
    state.trackedDocuments.delete(document);
    state.browserDocuments.delete(browser);
  };

  /* Thunderbird sizes a conversation article once when MsgLoaded fires by
   * adding the message document height to #singleMessage. The Outlook card
   * adds wrapper padding and borders outside those two measurements, and late
   * header/body layout can make the one-shot result stale. On a short message
   * with a tall recipient header, those missing pixels can consume the entire
   * MIME browser and make a correctly loaded body appear blank. */
  const elementIsHidden = element => {
    if (!element?.isConnected) {
      return true;
    }
    const hidden = element.getAttribute?.("hidden");
    const collapsed = element.getAttribute?.("collapsed");
    return Boolean(
      element.hidden ||
        (hidden != null && hidden !== "false") ||
        (collapsed != null && collapsed !== "false")
    );
  };

  const getContentScrollHeight = element => {
    if (!element?.isConnected) {
      return 0;
    }
    try {
      if (elementIsHidden(element)) {
        return 0;
      }
      return Number(element.scrollHeight) || 0;
    } catch (error) {
      return 0;
    }
  };

  const getVisibleBlockSize = element => {
    if (!element?.isConnected || elementIsHidden(element)) {
      return 0;
    }
    try {
      return Number(element.getBoundingClientRect().height) || 0;
    } catch (error) {
      return 0;
    }
  };

  const disconnectConversationSizingObservers = record => {
    for (const observer of record.resizeObservers) {
      observer.disconnect();
    }
    record.resizeObservers.length = 0;
    record.readerDocument = null;
    record.mimeDocument = null;
  };

  const removeConversationBrowserSizing = record => {
    record.browser.removeEventListener("load", record.onLoad, true);
    record.browser.removeEventListener("MsgLoaded", record.onMsgLoaded, true);
    if (record.resizeFrame) {
      window.cancelAnimationFrame(record.resizeFrame);
      record.resizeFrame = 0;
    }
    for (const timer of record.settleTimers) {
      window.clearTimeout(timer);
    }
    record.settleTimers.clear();
    if (record.repairBudgetTimer) {
      window.clearTimeout(record.repairBudgetTimer);
      record.repairBudgetTimer = 0;
    }
    disconnectConversationSizingObservers(record);
  };

  const measureConversationBrowser = record => {
    const { browser } = record;
    const article = browser.parentElement;
    if (
      !browser.isConnected ||
      !article?.matches?.('article[aria-expanded="true"]')
    ) {
      return;
    }

    let readerDocument;
    let mimeDocument;
    let messageBrowser;
    try {
      readerDocument = browser.contentDocument;
      if (readerDocument?.documentURI !== "about:message") {
        return;
      }
      messageBrowser = readerDocument.getElementById("messagepane");
      mimeDocument = messageBrowser?.contentDocument;
      if (
        !mimeDocument?.documentElement ||
        !MESSAGE_DOCUMENT_URI_PATTERN.test(mimeDocument.documentURI || "")
      ) {
        return;
      }
    } catch (error) {
      /* A nested reader can change remoteness while the message is loading. */
      return;
    }

    const header = readerDocument.getElementById("singleMessage");
    const notification = readerDocument.getElementById(
      "mail-notification-top"
    );
    const invitation = readerDocument.getElementById(
      "calendarInvitationDisplayContainer"
    );
    const attachmentView = readerDocument.getElementById("attachmentView");
    const findToolbar = readerDocument.getElementById("findToolbar");

    const contentHeight = Math.max(
      getContentScrollHeight(mimeDocument.documentElement),
      getContentScrollHeight(mimeDocument.body),
      getContentScrollHeight(invitation)
    );
    const desiredBodyHeight = Math.max(160, contentHeight);
    const visibleBodyHeight = Math.max(
      getVisibleBlockSize(messageBrowser),
      getVisibleBlockSize(invitation)
    );
    const missingBodyHeight = desiredBodyHeight - visibleBodyHeight;
    if (
      Number.isFinite(missingBodyHeight) &&
      missingBodyHeight > 1 &&
      record.repairBudget > 0 &&
      record.totalGrowth < 100000
    ) {
      const renderedArticleHeight = getVisibleBlockSize(article);
      const currentMinimum = Number.parseFloat(article.style.minHeight) || 0;
      const growth = Math.min(
        Math.ceil(missingBodyHeight),
        100000 - record.totalGrowth
      );
      const requiredMinimum = Math.ceil(
        Math.max(currentMinimum, renderedArticleHeight) + growth
      );
      if (requiredMinimum > currentMinimum) {
        article.style.minHeight = `${requiredMinimum}px`;
        record.totalGrowth += growth;
        record.repairBudget--;
        if (record.repairBudgetTimer) {
          window.clearTimeout(record.repairBudgetTimer);
        }
        record.repairBudgetTimer = window.setTimeout(() => {
          record.repairBudgetTimer = 0;
          record.repairBudget = 2;
        }, 1000);
      }
    }

    if (
      record.readerDocument === readerDocument &&
      record.mimeDocument === mimeDocument
    ) {
      return;
    }
    disconnectConversationSizingObservers(record);
    record.readerDocument = readerDocument;
    record.mimeDocument = mimeDocument;
    try {
      const readerObserver = new readerDocument.defaultView.ResizeObserver(
        record.schedule
      );
      for (const element of [
        header,
        notification,
        invitation,
        attachmentView,
        findToolbar,
        messageBrowser,
      ]) {
        if (element?.isConnected) {
          readerObserver.observe(element);
        }
      }
      record.resizeObservers.push(readerObserver);
    } catch (error) {
      /* Delayed settle passes still cover readers without ResizeObserver. */
    }
    try {
      const mimeObserver = new mimeDocument.defaultView.ResizeObserver(
        record.schedule
      );
      mimeObserver.observe(mimeDocument.documentElement);
      if (mimeDocument.body) {
        mimeObserver.observe(mimeDocument.body);
      }
      record.resizeObservers.push(mimeObserver);
    } catch (error) {
      /* A MIME document can be replaced while observers are being attached. */
    }
  };

  const bindConversationBrowserSizing = (browser, listener) => {
    for (const [trackedBrowser, record] of [...listener.sizingRecords]) {
      if (!trackedBrowser.isConnected) {
        removeConversationBrowserSizing(record);
        listener.sizingRecords.delete(trackedBrowser);
      }
    }
    let record = listener.sizingRecords.get(browser);
    if (record) {
      record.schedule();
      return;
    }
    record = {
      browser,
      mimeDocument: null,
      onLoad: null,
      onMsgLoaded: null,
      queueSettle: null,
      readerDocument: null,
      repairBudget: 2,
      repairBudgetTimer: 0,
      resizeFrame: 0,
      resizeObservers: [],
      schedule: null,
      settleTimers: new Set(),
      totalGrowth: 0,
    };
    record.schedule = () => {
      if (
        state.shuttingDown ||
        !record.browser.isConnected ||
        record.resizeFrame
      ) {
        return;
      }
      record.resizeFrame = window.requestAnimationFrame(() => {
        record.resizeFrame = window.requestAnimationFrame(() => {
          record.resizeFrame = 0;
          measureConversationBrowser(record);
        });
      });
    };
    record.queueSettle = () => {
      record.repairBudget = 2;
      if (record.repairBudgetTimer) {
        window.clearTimeout(record.repairBudgetTimer);
        record.repairBudgetTimer = 0;
      }
      for (const timer of record.settleTimers) {
        window.clearTimeout(timer);
      }
      record.settleTimers.clear();
      record.schedule();
      for (const delay of [100, 500, 1500]) {
        const timer = window.setTimeout(() => {
          record.settleTimers.delete(timer);
          record.schedule();
        }, delay);
        record.settleTimers.add(timer);
      }
    };
    record.onLoad = record.queueSettle;
    record.onMsgLoaded = record.queueSettle;
    browser.addEventListener("load", record.onLoad, true);
    browser.addEventListener("MsgLoaded", record.onMsgLoaded, true);
    listener.sizingRecords.set(browser, record);
    record.queueSettle();
  };

  const stampBrowserDocument = (browser, document, scheme) => {
    const previousDocument = state.browserDocuments.get(browser);
    if (previousDocument && previousDocument !== document) {
      untrackBrowserDocument(browser);
    }
    stampDocument(document, scheme);
    if (document?.documentElement) {
      state.browserDocuments.set(browser, document);
    }
  };

  const bindBrowser = browser => {
    if (!browser || state.browserListeners.has(browser)) {
      return;
    }
    const onLoad = () => {
      untrackBrowserDocument(browser);
      scheduleSync();
    };
    browser.addEventListener("load", onLoad, true);
    state.browserListeners.set(browser, onLoad);
  };

  const removeConversationBinding = conversationView => {
    const listener = state.conversationListeners.get(conversationView);
    if (!listener) {
      removeConversationGuard(conversationView);
      return;
    }
    listener.observer.disconnect();
    listener.main.removeEventListener("keydown", listener.onKeyDown);
    listener.main.removeEventListener("click", listener.onClick, true);
    for (const record of listener.sizingRecords.values()) {
      removeConversationBrowserSizing(record);
    }
    listener.sizingRecords.clear();
    for (const browser of listener.main.querySelectorAll(
      'browser[src="about:message"]'
    )) {
      const onLoad = state.browserListeners.get(browser);
      if (onLoad) {
        browser.removeEventListener("load", onLoad, true);
        state.browserListeners.delete(browser);
      }
      untrackBrowserDocument(browser);
    }
    for (const article of listener.main.querySelectorAll(
      'article[data-outlook-style-openable="true"]'
    )) {
        article.removeAttribute("data-outlook-style-openable");
        if (article.getAttribute("data-outlook-style-tabindex") === "true") {
          article.removeAttribute("tabindex");
          article.removeAttribute("data-outlook-style-tabindex");
        }
    }
    for (const article of listener.main.querySelectorAll(
      'article[data-outlook-style-focus-target="true"]'
    )) {
      article.removeAttribute("tabindex");
      article.removeAttribute("data-outlook-style-focus-target");
    }
    listener.style.remove();
    conversationView.removeAttribute("data-outlook-style-conversation");
    conversationView.removeAttribute(OUTLOOK_COLOR_SCHEME_ATTRIBUTE);
    state.conversationListeners.delete(conversationView);
    removeConversationGuard(conversationView);
  };

  const bindConversationView = (about3Pane, scheme) => {
    const conversationView = about3Pane.document.querySelector(
      "message-pane > conversation-view"
    );
    const shadowRoot = conversationView?.shadowRoot;
    const main = shadowRoot?.getElementById("mainConversation");
    if (!conversationView || !shadowRoot || !main) {
      return;
    }
    conversationView.setAttribute(OUTLOOK_COLOR_SCHEME_ATTRIBUTE, scheme);

    let listener = state.conversationListeners.get(conversationView);
    const decorate = () => {
      for (const article of main.querySelectorAll(
        'article[aria-expanded="false"]'
      )) {
        article.setAttribute("data-outlook-style-openable", "true");
        if (!article.hasAttribute("tabindex")) {
          article.tabIndex = 0;
          article.setAttribute("data-outlook-style-tabindex", "true");
        }
      }
      for (const browser of shadowRoot.querySelectorAll(
        'browser[src="about:message"]'
      )) {
        bindBrowser(browser);
        if (listener) {
          bindConversationBrowserSizing(browser, listener);
        }
        try {
          if (browser.contentWindow?.location?.href === "about:message") {
            stampBrowserDocument(
              browser,
              browser.contentWindow.document,
              conversationView.getAttribute(
                OUTLOOK_COLOR_SCHEME_ATTRIBUTE
              ) || scheme
            );
          }
        } catch (error) {
          /* The embedded reader may still be navigating to about:message. */
        }
      }
    };

    if (!listener || listener.main !== main) {
      if (listener) {
        removeConversationBinding(conversationView);
      }
      const style = about3Pane.document.createElement("style");
      style.id = CONVERSATION_VIEW_STYLE_ID;
      style.textContent = CONVERSATION_VIEW_CSS;
      shadowRoot.appendChild(style);

      const onKeyDown = event => {
        if (
          (event.key !== "Enter" && event.key !== " ") ||
          event.repeat ||
          event.altKey ||
          event.ctrlKey ||
          event.metaKey
        ) {
          return;
        }
        const article = event.target?.closest?.(
          'article[aria-expanded="false"]'
        );
        if (!article || !main.contains(article)) {
          return;
        }
        const openBrowser = main.querySelector(
          'article[aria-expanded="true"] > browser[src="about:message"]'
        );
        if (
          openBrowser?.hidden ||
          !openBrowser?.parentElement?.dataset?.messageId
        ) {
          event.preventDefault();
          event.stopPropagation();
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        article.click();
        about3Pane.setTimeout(() => {
          const expandedArticle = main.querySelector(
            'article[aria-expanded="true"]'
          );
          const browser = expandedArticle?.querySelector(
            'browser[src="about:message"]'
          );
          if (expandedArticle) {
            expandedArticle.tabIndex = -1;
            expandedArticle.setAttribute(
              "data-outlook-style-focus-target",
              "true"
            );
            expandedArticle.focus();
          }
          if (browser) {
            const focusBrowser = () => browser.focus();
            if (browser.hidden) {
              browser.addEventListener("MsgLoaded", focusBrowser, {
                capture: true,
                once: true,
              });
            } else {
              focusBrowser();
            }
          }
        }, 0);
      };
      const onClick = event => {
        const article = event.target?.closest?.(
          'article[aria-expanded="false"]'
        );
        if (!article || !main.contains(article)) {
          return;
        }
        const openBrowser = main.querySelector(
          'article[aria-expanded="true"] > browser[src="about:message"]'
        );
        if (
          openBrowser?.hidden ||
          !openBrowser?.parentElement?.dataset?.messageId
        ) {
          event.preventDefault();
          event.stopPropagation();
        }
      };
      main.addEventListener("keydown", onKeyDown);
      main.addEventListener("click", onClick, true);
      const observer = new about3Pane.MutationObserver(() => {
        decorate();
        scheduleSync();
      });
      observer.observe(main, { childList: true, subtree: true });
      listener = {
        main,
        observer,
        onClick,
        onKeyDown,
        sizingRecords: new Map(),
        style,
      };
      state.conversationListeners.set(conversationView, listener);
      conversationView.setAttribute(
        "data-outlook-style-conversation",
        "true"
      );
    }
    decorate();
  };

  const bindAbout3Pane = (about3Pane, scheme) => {
    const paneDocument = about3Pane.document;
    const previousCreation = state.folderPaneCreationCleanups.get(about3Pane);
    if (previousCreation?.document !== paneDocument) {
      previousCreation?.cleanup?.();
      state.folderPaneCreationCleanups.delete(about3Pane);
    }
    if (!state.folderPaneCreationCleanups.has(about3Pane)) {
      const cleanup = installFolderPaneNewItemButton(about3Pane);
      if (cleanup) {
        state.folderPaneCreationCleanups.set(about3Pane, {
          cleanup,
          document: paneDocument,
        });
      }
    }
    installConversationSelectionCapture(about3Pane, () =>
      refreshConversationForPane(about3Pane)
    );
    stampDocument(paneDocument, scheme);
    const previousListener = state.paneListeners.get(about3Pane);
    if (previousListener?.paneDocument !== paneDocument) {
      previousListener?.paneDocument.removeEventListener(
        "load",
        previousListener.onLoad,
        true
      );
      previousListener?.observer?.disconnect();
      state.paneListeners.delete(about3Pane);
    }
    if (!state.paneListeners.has(about3Pane)) {
      const onLoad = () => scheduleSync();
      paneDocument.addEventListener("load", onLoad, true);
      const observer = new about3Pane.MutationObserver(scheduleSync);
      state.paneListeners.set(about3Pane, {
        paneDocument,
        onLoad,
        observer,
        observedMessagePane: null,
      });
    }
    const paneListener = state.paneListeners.get(about3Pane);
    const messagePane = paneDocument.querySelector("message-pane");
    if (
      messagePane &&
      paneListener.observedMessagePane !== messagePane
    ) {
      paneListener.observer.disconnect();
      paneListener.observer.observe(messagePane, { childList: true });
      paneListener.observedMessagePane = messagePane;
    }
    bindConversationView(about3Pane, scheme);
    for (const browser of getAboutMessageBrowsers(about3Pane)) {
      bindBrowser(browser);
      try {
        const contentWindow = browser.contentWindow;
        const href = contentWindow?.location?.href || "";
        if (
          href === "about:message" ||
          href.startsWith(
            "chrome://messenger/content/multimessageview.xhtml"
          )
        ) {
          stampBrowserDocument(browser, contentWindow.document, scheme);
        }
      } catch (error) {
        /* The browser may still be navigating to its chrome document. */
      }
    }
  };

  const syncScheme = () => {
    state.schemeFrame = 0;
    const scheme = getOutlookColorScheme(window);
    for (const about3Pane of getAbout3PaneWindows(window)) {
      bindAbout3Pane(about3Pane, scheme);
    }
    for (const tab of tabmail.tabInfo || []) {
      bindBrowser(tab.chromeBrowser);
    }
    if (tabmail.currentTabInfo?.chromeBrowser) {
      bindBrowser(tabmail.currentTabInfo.chromeBrowser);
    }

    /* Release listeners held only by tabs or message browsers that closed. */
    for (const [browser, onLoad] of [...state.browserListeners]) {
      if (!browser.isConnected) {
        browser.removeEventListener("load", onLoad, true);
        untrackBrowserDocument(browser);
        state.browserListeners.delete(browser);
      }
    }
    for (const [about3Pane, listener] of [...state.paneListeners]) {
      if (about3Pane.closed || !listener.paneDocument.defaultView) {
        listener.paneDocument.removeEventListener(
          "load",
          listener.onLoad,
          true
        );
        listener.observer?.disconnect();
        try {
          listener.paneDocument.documentElement?.removeAttribute(
            OUTLOOK_COLOR_SCHEME_ATTRIBUTE
          );
        } catch (error) {
          /* The closed pane document may already be unavailable. */
        }
        state.trackedDocuments.delete(listener.paneDocument);
        state.folderPaneCreationCleanups.get(about3Pane)?.cleanup?.();
        state.folderPaneCreationCleanups.delete(about3Pane);
        removeConversationSelectionCapture(about3Pane);
        state.paneListeners.delete(about3Pane);
      }
    }
    for (const conversationView of [...state.conversationListeners.keys()]) {
      if (!conversationView.isConnected) {
        removeConversationBinding(conversationView);
      }
    }
  };
  const scheduleSync = () => {
    if (window.closed || state.shuttingDown || state.schemeFrame) {
      return;
    }
    state.schemeFrame = window.requestAnimationFrame(syncScheme);
  };

  const systemScheme = window.matchMedia("(prefers-color-scheme: dark)");
  const tabEventTarget =
    tabmail.tabContainer ||
    window.document.getElementById("tabmail-tabs") ||
    window.document;
  Object.assign(state, {
    scheduleSync,
    removeConversationBinding,
    systemScheme,
    tabmail,
    tabEventTarget,
    untrackBrowserDocument,
  });
  mainPaneSchemeState.set(window, state);
  enhancedMainPaneSchemeWindows.add(window);

  window.addEventListener("windowlwthemeupdate", scheduleSync);
  systemScheme.addEventListener("change", scheduleSync);
  tabEventTarget.addEventListener("TabSelect", scheduleSync);
  tabEventTarget.addEventListener("TabOpen", scheduleSync);
  tabEventTarget.addEventListener("TabClose", scheduleSync);
  scheduleSync();
  /* about:3pane can finish after the already-open messenger window is handed
   * to the extension. One bounded retry covers that startup race. */
  state.retryTimer = window.setTimeout(() => {
    state.retryTimer = 0;
    scheduleSync();
  }, 1000);
  return true;
}

function removeMainPaneSchemeBridge(window) {
  const state = mainPaneSchemeState.get(window);
  if (!state) {
    return;
  }
  state.shuttingDown = true;
  state.removeCalendarCreationEnhancements?.();
  window.removeEventListener("windowlwthemeupdate", state.scheduleSync);
  state.systemScheme.removeEventListener("change", state.scheduleSync);
  state.tabEventTarget.removeEventListener("TabSelect", state.scheduleSync);
  state.tabEventTarget.removeEventListener("TabOpen", state.scheduleSync);
  state.tabEventTarget.removeEventListener("TabClose", state.scheduleSync);
  if (state.schemeFrame) {
    window.cancelAnimationFrame(state.schemeFrame);
  }
  if (state.retryTimer) {
    window.clearTimeout(state.retryTimer);
  }
  for (const [browser, onLoad] of state.browserListeners) {
    browser.removeEventListener("load", onLoad, true);
    state.untrackBrowserDocument(browser);
  }
  for (const listener of state.paneListeners.values()) {
    listener.paneDocument.removeEventListener("load", listener.onLoad, true);
    listener.observer?.disconnect();
  }
  for (const record of state.folderPaneCreationCleanups.values()) {
    record.cleanup();
  }
  state.folderPaneCreationCleanups.clear();
  const ownedConversations = [];
  for (const about3Pane of getAbout3PaneWindows(window)) {
    const conversationView = about3Pane.document.querySelector(
      `message-pane > conversation-view[${OWNED_CONVERSATION_ATTRIBUTE}="true"]`
    );
    if (conversationView) {
      ownedConversations.push(conversationView);
    }
  }
  for (const conversationView of ownedConversations) {
    retireGuardedConversationWithFallback(conversationView);
  }
  for (const conversationView of [...state.conversationListeners.keys()]) {
    state.removeConversationBinding(conversationView);
  }
  for (const document of state.trackedDocuments) {
    try {
      document.documentElement?.removeAttribute(
        OUTLOOK_COLOR_SCHEME_ATTRIBUTE
      );
    } catch (error) {
      /* A tracked document may already have been destroyed. */
    }
  }
  mainPaneSchemeState.delete(window);
  enhancedMainPaneSchemeWindows.delete(window);
}

const EDITOR_SURFACE_CSS = `
  :root {
    color-scheme: light dark !important;
    --outlook-editor-background: #ffffff;
    --outlook-editor-color: #242424;
    --outlook-editor-link: #0f6cbd;
    --outlook-editor-visited-link: #5c2e91;
    --outlook-editor-selection-background: #cfe4fa;
    --outlook-editor-selection-color: #242424;
    background-color: var(--outlook-editor-background) !important;
  }

  __OUTLOOK_EDITOR_BACKGROUND_SELECTOR__ {
    background-color: var(--outlook-editor-background) !important;
  }

  __OUTLOOK_EDITOR_COLOR_SELECTOR__ {
    color: var(--outlook-editor-color) !important;
  }

  a:link {
    color: var(--outlook-editor-link);
  }

  a:visited {
    color: var(--outlook-editor-visited-link);
  }

  ::selection {
    color: var(--outlook-editor-selection-color);
    background-color: var(--outlook-editor-selection-background);
  }

  @media (prefers-color-scheme: dark) {
    :root {
      --outlook-editor-background: #1f1f1f;
      --outlook-editor-color: #f5f5f5;
      --outlook-editor-link: #75b9f2;
      --outlook-editor-visited-link: #d8b4fe;
      --outlook-editor-selection-background: #115ea3;
      --outlook-editor-selection-color: #ffffff;
    }
  }

  @media (forced-colors: active) {
    :root {
      color-scheme: light dark !important;
      background-color: Canvas !important;
    }

    __OUTLOOK_EDITOR_BACKGROUND_SELECTOR__ {
      background-color: Canvas !important;
    }

    __OUTLOOK_EDITOR_COLOR_SELECTOR__ {
      color: CanvasText !important;
    }

    a:link,
    a:visited {
      color: LinkText;
    }

    ::selection {
      color: HighlightText;
      background-color: Highlight;
    }
  }
`;

function maybeApplyOutlookPaneWidths(about3Pane) {
  if (
    Services.prefs.getBoolPref(PANE_WIDTHS_PREF, false) ||
    about3Pane?.location?.href !== "about:3pane"
  ) {
    return false;
  }

  const layout = about3Pane.paneLayout;
  const folderSplitter = layout?.folderPaneSplitter;
  const messageSplitter = layout?.messagePaneSplitter;
  const threadPane = about3Pane.document.getElementById("threadPane");
  const messagePane = about3Pane.document.getElementById("messagePane");
  if (
    !layout?.classList.contains("layout-vertical") ||
    layout.getBoundingClientRect().width <= 0 ||
    !folderSplitter ||
    !messageSplitter ||
    folderSplitter.isCollapsed ||
    messageSplitter.isCollapsed ||
    !threadPane ||
    !messagePane
  ) {
    return false;
  }

  const initialCombinedWidth =
    threadPane.getBoundingClientRect().width +
    messagePane.getBoundingClientRect().width;
  if (!Number.isFinite(initialCombinedWidth) || initialCombinedWidth < 800) {
    return false;
  }

  /* PaneSplitter.width is Thunderbird's own persisted sizing API. Applying
   * this once corrects layouts inherited from the very wide pre-1.0.6 list;
   * the native splitter-resized event saves it, and every later user drag
   * remains authoritative. */
  folderSplitter.width = 230;
  const threadWidth = threadPane.getBoundingClientRect().width;
  const messageWidth = messagePane.getBoundingClientRect().width;
  const combinedWidth = threadWidth + messageWidth;

  /* Outlook's list is about 390 px on a roomy window. Preserve at least
   * 640 px for reading when possible and gracefully tighten on smaller ones. */
  const targetThreadWidth = Math.min(
    390,
    Math.max(320, Math.round(combinedWidth - 640))
  );
  messageSplitter.width = Math.max(
    480,
    Math.round(combinedWidth - targetThreadWidth)
  );

  for (const splitter of [folderSplitter, messageSplitter]) {
    splitter.dispatchEvent(
      new about3Pane.CustomEvent("splitter-resized", { bubbles: true })
    );
  }
  Services.prefs.setBoolPref(PANE_WIDTHS_PREF, true);
  return true;
}

function scheduleOutlookPaneWidths(window) {
  const about3Pane = window.document.getElementById("tabmail")?.currentTabInfo
    ?.chromeBrowser?.contentWindow;
  const ready = about3Pane?.hasDOMContentLoaded?.promise;
  if (!ready) {
    maybeApplyOutlookPaneWidths(about3Pane);
    return;
  }
  ready.then(() => maybeApplyOutlookPaneWidths(about3Pane)).catch(error => {
    console.error(
      "Outlook Style Companion could not set the initial pane widths:",
      error
    );
  });
}

function setTodayPaneMode(window, mode) {
  const agendaPanel = window.document.getElementById("agenda-panel");
  const todoPanel = window.document.getElementById("todo-tab-panel");
  if (!agendaPanel || !todoPanel || !window.TodayPane?.isLoaded) {
    return;
  }

  const showCalendar = mode === "calendar";
  agendaPanel.setVisible(showCalendar);
  todoPanel.setVisible(!showCalendar);
  if (showCalendar) {
    window.TodayPane.displayMiniSection?.("minimonth");
  }
  window.TodayPane.updateDisplay();
}

function getTodayPaneSectionVisibility(window) {
  const document = window.document;
  return {
    calendar: Boolean(document.getElementById("agenda-panel")?.isVisible()),
    tasks: Boolean(document.getElementById("todo-tab-panel")?.isVisible()),
  };
}

function setTodayPaneSectionVisibility(window, showCalendar, showTasks) {
  const document = window.document;
  const agendaPanel = document.getElementById("agenda-panel");
  const todoPanel = document.getElementById("todo-tab-panel");
  if (!agendaPanel || !todoPanel || !window.TodayPane?.isLoaded) {
    return false;
  }

  /* Thunderbird's TodayPane also protects against this state, but prevent it
   * here so the Outlook-style menu never briefly leaves an empty sidebar. */
  if (!showCalendar && !showTasks) {
    return false;
  }

  agendaPanel.setVisible(showCalendar);
  todoPanel.setVisible(showTasks);
  if (showCalendar) {
    window.TodayPane.displayMiniSection?.("minimonth");
  }
  window.TodayPane.updateDisplay();
  syncTodayPaneTabs(window);
  return true;
}

function setMenuItemChecked(menuItem, checked) {
  if (checked) {
    menuItem.setAttribute("checked", "true");
  } else {
    menuItem.removeAttribute("checked");
  }
}

function getDisplayedMiniMonthDate(minimonth) {
  const displayedYear = Number(minimonth?.getAttribute("year"));
  const displayedMonth = Number(minimonth?.getAttribute("month"));
  if (
    Number.isInteger(displayedYear) &&
    displayedYear > 0 &&
    Number.isInteger(displayedMonth) &&
    displayedMonth >= 0 &&
    displayedMonth <= 11
  ) {
    return new Date(displayedYear, displayedMonth, 1, 12);
  }
  const date = minimonth?.value || new Date();
  return new Date(date.getFullYear(), date.getMonth(), 1, 12);
}

function syncTodayPaneTabs(window) {
  const document = window.document;
  const nativeHeader = document.getElementById("today-pane-header");
  const calendarTab = document.getElementById("fluent-myday-calendar-tab");
  const todoTab = document.getElementById("fluent-myday-todo-tab");
  if (!nativeHeader || !calendarTab || !todoTab) {
    return;
  }

  /* Thunderbird's index 0 is its combined Events and Tasks view. Treat that
   * as Calendar so an existing preference remains usable until either direct
   * tab is selected. */
  const todoSelected = nativeHeader.getAttribute("index") === "1";
  calendarTab.setAttribute("aria-selected", String(!todoSelected));
  todoTab.setAttribute("aria-selected", String(todoSelected));
}

function parseAgendaDate(value) {
  const match = /^(\d{4})(\d{2})(\d{2})/.exec(String(value || ""));
  if (!match) {
    return null;
  }
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function dayDifference(left, right) {
  const leftUtc = Date.UTC(left.getFullYear(), left.getMonth(), left.getDate());
  const rightUtc = Date.UTC(right.getFullYear(), right.getMonth(), right.getDate());
  return Math.round((leftUtc - rightUtc) / 86400000);
}

function formatAgendaHeading(window, date) {
  const formatted = new window.Intl.DateTimeFormat(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
  }).format(date);
  const difference = dayDifference(date, new Date());
  if (difference === 0) {
    return `Today \u2022 ${formatted}`;
  }
  if (difference === 1) {
    return `Tomorrow \u2022 ${formatted}`;
  }
  return formatted;
}

function formatEventDuration(seconds) {
  const totalMinutes = Math.max(0, Math.round(seconds / 60));
  if (!totalMinutes) {
    return "";
  }
  if (totalMinutes < 60) {
    return `${totalMinutes} min`;
  }

  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return minutes ? `${hours} hr ${minutes} min` : `${hours} hr`;
}

function formatAgendaCountdown(minutesUntil) {
  return minutesUntil === 1 ? "in 1 min" : `in ${minutesUntil} mins`;
}

function updateAgendaTimeMarker(document, listItem, markerDetails) {
  const details = listItem.querySelector(":scope > .agenda-listitem-details");
  let marker = details?.querySelector(":scope > .fluent-time-marker");
  if (!details || !markerDetails) {
    marker?.remove();
    return;
  }

  if (!marker) {
    marker = document.createElementNS(HTML_NS, "span");
    marker.className = "fluent-time-marker";
    marker.setAttribute("role", "note");

    const pill = document.createElementNS(HTML_NS, "span");
    pill.className = "fluent-time-pill";

    const icon = document.createElementNS(HTML_NS, "span");
    icon.className = "fluent-time-pill-icon";
    icon.setAttribute("aria-hidden", "true");
    icon.textContent = "\u25f7";

    const label = document.createElementNS(HTML_NS, "span");
    label.className = "fluent-time-pill-label";
    pill.append(icon, label);
    marker.append(pill);
    details.append(marker);
  }

  const label = marker.querySelector(":scope .fluent-time-pill-label");
  if (label && label.textContent !== markerDetails.label) {
    label.textContent = markerDetails.label;
  }
  if (marker.getAttribute("aria-label") !== markerDetails.description) {
    marker.setAttribute("aria-label", markerDetails.description);
  }
}

function decorateAgenda(window) {
  const document = window.document;
  const agenda = document.getElementById("agenda");
  const heading = document.getElementById("fluent-today-date-heading");
  if (!agenda) {
    return;
  }

  const paneStart = window.TodayPane?.start;
  let paneDate = null;
  let paneDateKey = "";
  if (paneStart) {
    paneDate = new Date(paneStart.year, paneStart.month, paneStart.day);
    paneDateKey = `${String(paneStart.year).padStart(4, "0")}${String(
      paneStart.month + 1
    ).padStart(2, "0")}${String(paneStart.day).padStart(2, "0")}`;
  }
  if (heading && paneDate) {
    heading.setAttribute("value", formatAgendaHeading(window, paneDate));
  }
  const quickTask = document.querySelector(
    "#fluent-today-intro .fluent-quick-task"
  );
  const quickTaskCommand = document.getElementById(
    "calendar_new_todo_todaypane_command"
  );
  if (quickTask) {
    quickTask.disabled = quickTaskCommand?.getAttribute("disabled") === "true";
  }

  const listItems = [...agenda.querySelectorAll(".agenda-listitem")];
  for (const listItem of listItems) {
    listItem.removeAttribute("data-fluent-next-event");
    listItem.removeAttribute("data-fluent-time-state");
    listItem.removeAttribute("data-fluent-time-attention");
  }

  /* Outlook displays a single Up Next treatment for accepted invitations and
   * personally-created meetings. An accepted meeting within five minutes
   * away takes priority over an older meeting still in progress. Otherwise,
   * follow the active meeting that started most recently, or the earliest
   * upcoming meeting. The Companion owns the exact countdown text so "now" is
   * never shown before the selected meeting has actually started. */
  let nextEvent = null;
  if (window.TodayPane?.showsToday) {
    const now = cal.dtz.now();
    const eligibleEvents = [];
    for (const [index, listItem] of listItems.entries()) {
      const item = listItem.item;
      if (
        listItem.hidden ||
        listItem.classList.contains("agenda-listitem-end") ||
        listItem.classList.contains("agenda-listitem-past") ||
        !item?.startDate ||
        item.startDate.isDate ||
        (typeof item.isEvent === "function" && !item.isEvent())
      ) {
        continue;
      }

      const itemStatus = String(
        item.status || item.getProperty?.("STATUS") || ""
      ).toUpperCase();
      const participationStatus = String(
        listItem.getAttribute("status") || ""
      ).toUpperCase();
      if (
        itemStatus === "CANCELLED" ||
        (participationStatus && participationStatus !== "ACCEPTED")
      ) {
        continue;
      }

      const startDate = listItem._localStartDate || item.startDate;
      try {
        const endDate = listItem._localEndDate || item.endDate;
        if (endDate && endDate.compare(now) <= 0) {
          continue;
        }
      } catch (error) {
        /* If a provider supplies an unusual date object, retain the future
         * row rather than hiding the only useful Up Next indicator. */
      }

      let isActive = listItem.classList.contains("agenda-listitem-now");
      let secondsUntil = null;
      try {
        secondsUntil = startDate.subtractDate(now).inSeconds;
        isActive = secondsUntil <= 0;
      } catch (error) {
        try {
          isActive = startDate.compare(now) <= 0;
        } catch (compareError) {
          /* Thunderbird's native now class remains a safe fallback. */
        }
      }
      eligibleEvents.push({
        index,
        isActive,
        listItem,
        secondsUntil,
        startDate,
      });
    }

    const pickEarlier = (earliest, candidate) => {
      if (!earliest) {
        return candidate;
      }
      try {
        return candidate.startDate.compare(earliest.startDate) < 0
          ? candidate
          : earliest;
      } catch (error) {
        return candidate.index < earliest.index ? candidate : earliest;
      }
    };
    const activeEvents = eligibleEvents.filter(event => event.isActive);
    const activeEvent = activeEvents.reduce((latest, candidate) => {
      if (!latest) {
        return candidate;
      }
      try {
        return candidate.startDate.compare(latest.startDate) > 0
          ? candidate
          : latest;
      } catch (error) {
        return candidate.index > latest.index ? candidate : latest;
      }
    }, null);
    const futureEvents = eligibleEvents.filter(event => !event.isActive);
    const imminentEvent = futureEvents
      .filter(
        event =>
          Number.isFinite(event.secondsUntil) &&
          event.secondsUntil > 0 &&
          event.secondsUntil <= 5 * 60
      )
      .reduce(pickEarlier, null);
    const futureEvent = futureEvents.reduce(pickEarlier, null);
    nextEvent = imminentEvent || activeEvent || futureEvent;

    if (nextEvent) {
      let label = "";
      let description = "";
      let state = "upcoming";
      let shouldShow = false;
      let needsAttention = false;
      if (nextEvent.isActive) {
        label = "now";
        description = "Meeting in progress";
        state = "active";
        shouldShow = true;
        needsAttention =
          Number.isFinite(nextEvent.secondsUntil) &&
          nextEvent.secondsUntil > -3 * 60;
      } else if (Number.isFinite(nextEvent.secondsUntil)) {
        const minutesUntil = Math.max(
          1,
          Math.ceil(nextEvent.secondsUntil / 60)
        );
        label = formatAgendaCountdown(minutesUntil);
        description =
          minutesUntil === 1
            ? "Meeting starts in 1 minute"
            : `Meeting starts in ${minutesUntil} minutes`;
        const isSameDay =
          nextEvent.startDate.year === now.year &&
          nextEvent.startDate.month === now.month &&
          nextEvent.startDate.day === now.day;
        shouldShow = isSameDay || nextEvent.secondsUntil < 12 * 60 * 60;
        if (nextEvent.secondsUntil <= 5 * 60) {
          state = "imminent";
          needsAttention = true;
        }
      }

      if (shouldShow && label) {
        nextEvent.markerDetails = { description, label };
        nextEvent.listItem.setAttribute("data-fluent-next-event", "true");
        nextEvent.listItem.setAttribute("data-fluent-time-state", state);
        if (needsAttention) {
          nextEvent.listItem.setAttribute(
            "data-fluent-time-attention",
            "true"
          );
        }
      }
    }
  }

  for (const listItem of listItems) {
    updateAgendaTimeMarker(
      document,
      listItem,
      listItem === nextEvent?.listItem ? nextEvent.markerDetails : null
    );
    const dateHeader = listItem.querySelector(":scope > .agenda-date-header");
    const date = parseAgendaDate(listItem.dateString);
    listItem.removeAttribute("data-fluent-hide-date-header");
    if (
      dateHeader &&
      !dateHeader.hidden &&
      listItem.dateString?.slice(0, 8) === paneDateKey
    ) {
      listItem.setAttribute("data-fluent-hide-date-header", "true");
    } else if (dateHeader && date) {
      const heading = formatAgendaHeading(window, date);
      dateHeader.removeAttribute("data-l10n-id");
      if (dateHeader.textContent !== heading) {
        dateHeader.textContent = heading;
      }
    }

    const detailsInner = listItem.querySelector(".agenda-listitem-details-inner");
    if (!detailsInner) {
      continue;
    }

    let duration = detailsInner.querySelector(".fluent-event-duration");
    const item = listItem.item;
    let durationText = "";
    if (
      item?.startDate &&
      item?.endDate &&
      !item.startDate.isDate &&
      !listItem.classList.contains("agenda-listitem-end")
    ) {
      try {
        durationText = formatEventDuration(
          item.endDate.subtractDate(item.startDate).inSeconds
        );
      } catch (error) {
        console.error(
          "Outlook Style Companion could not calculate an event duration:",
          error
        );
      }
    }

    if (durationText) {
      if (!duration) {
        duration = document.createElementNS(HTML_NS, "span");
        duration.className = "fluent-event-duration";
        detailsInner.append(duration);
      }
      if (duration.textContent !== durationText) {
        duration.textContent = durationText;
      }
    } else {
      duration?.remove();
    }
  }
}

function removeStaleTodayPaneMarkup(window) {
  const document = window.document;
  for (const orphan of document.querySelectorAll(
    [
      "#fluent-myday-tabs",
      "#fluent-myday-popout",
      "#fluent-today-intro",
      "#fluent-today-month-button",
      "#fluent-today-month-panel",
      "#fluent-today-more",
      "#fluent-today-more-popup",
    ].join(",")
  )) {
    orphan.hidePopup?.();
    orphan.remove();
  }
  document
    .querySelector("#today-pane-panel > .sidebar-header")
    ?.classList.remove("fluent-myday-enhanced");
  document
    .getElementById("today-pane-panel")
    ?.removeAttribute("data-fluent-myday");
  const nativeMonthInput = document
    .getElementById("today-minimonth")
    ?.querySelector(".minimonth-header .minimonth-month-name");
  nativeMonthInput?.removeAttribute("hidden");
  nativeMonthInput?.removeAttribute("aria-hidden");
}

function enhanceTodayPane(window) {
  const document = window.document;
  const panel = document.getElementById("today-pane-panel");
  const header = panel?.querySelector(":scope > .sidebar-header");
  const nativeHeader = document.getElementById("today-pane-header");
  const closeButton = document.getElementById("today-closer");
  const agenda = document.getElementById("agenda");
  const agendaContainer = document.getElementById("agenda-container");
  const minimonth = document.getElementById("today-minimonth");
  const minimonthHeader = minimonth?.querySelector(".minimonth-header");
  const nativeMonthInput = minimonthHeader?.querySelector(
    ".minimonth-month-name"
  );
  const popupSet =
    document.getElementById("calendar-popupset") ||
    document.getElementById("mainPopupSet");
  if (
    !panel ||
    !header ||
    !nativeHeader ||
    !closeButton ||
    !agenda ||
    !agendaContainer ||
    !minimonth ||
    !minimonthHeader ||
    !nativeMonthInput ||
    !popupSet
  ) {
    return false;
  }
  if (todayPaneState.has(window)) {
    return true;
  }

  /* Thunderbird can session-restore chrome DOM while an add-on is updated.
   * If the old module did not receive its shutdown hook, discard its inert
   * injected controls before creating this version's live set. querySelectorAll
   * is intentional because duplicate IDs are exactly the state being repaired. */
  removeStaleTodayPaneMarkup(window);

  const tabs = document.createXULElement("hbox");
  tabs.id = "fluent-myday-tabs";
  tabs.setAttribute("role", "tablist");

  const calendarTab = document.createXULElement("toolbarbutton");
  calendarTab.id = "fluent-myday-calendar-tab";
  calendarTab.className = "fluent-myday-tab";
  calendarTab.setAttribute("label", "Calendar");
  calendarTab.setAttribute("role", "tab");
  calendarTab.setAttribute("tooltiptext", "Show calendar and events");

  const todoTab = document.createXULElement("toolbarbutton");
  todoTab.id = "fluent-myday-todo-tab";
  todoTab.className = "fluent-myday-tab";
  todoTab.setAttribute("label", "To Do");
  todoTab.setAttribute("role", "tab");
  todoTab.setAttribute("tooltiptext", "Show tasks");

  tabs.append(calendarTab, todoTab);
  header.insertBefore(tabs, header.firstElementChild);

  const popout = document.createXULElement("toolbarbutton");
  popout.id = "fluent-myday-popout";
  popout.setAttribute("tooltiptext", "Open Calendar");
  popout.setAttribute("aria-label", "Open Calendar");
  header.insertBefore(popout, closeButton);
  header.classList.add("fluent-myday-enhanced");
  panel.setAttribute("data-fluent-myday", "true");

  /* The native month name is a disabled input. Keep it in the DOM because
   * calendar-minimonth updates it internally, but replace it visually and for
   * accessibility with a real button backed by a functional month picker. */
  const nativeMonthInputWasHidden = nativeMonthInput.hasAttribute("hidden");
  const nativeMonthInputAriaHidden = nativeMonthInput.getAttribute("aria-hidden");
  nativeMonthInput.setAttribute("hidden", "true");
  nativeMonthInput.setAttribute("aria-hidden", "true");

  const monthButton = document.createElementNS(HTML_NS, "button");
  monthButton.id = "fluent-today-month-button";
  monthButton.className = "fluent-today-month-button";
  monthButton.type = "button";
  monthButton.setAttribute("aria-haspopup", "dialog");
  monthButton.setAttribute("aria-expanded", "false");
  monthButton.setAttribute("aria-controls", "fluent-today-month-panel");
  const monthLabel = document.createElementNS(HTML_NS, "span");
  monthLabel.className = "fluent-today-month-label";
  const monthChevron = document.createElementNS(HTML_NS, "span");
  monthChevron.className = "fluent-today-month-chevron";
  monthChevron.setAttribute("aria-hidden", "true");
  monthButton.append(monthLabel, monthChevron);
  nativeMonthInput.insertAdjacentElement("afterend", monthButton);

  const monthPanel = document.createXULElement("panel");
  monthPanel.id = "fluent-today-month-panel";
  monthPanel.className = "fluent-today-month-panel";
  monthPanel.setAttribute("type", "arrow");
  monthPanel.setAttribute("role", "dialog");
  monthPanel.setAttribute("aria-label", "Choose month");
  monthPanel.setAttribute("consumeoutsideclicks", "true");

  const monthPicker = document.createElementNS(HTML_NS, "div");
  monthPicker.className = "fluent-month-picker";
  const monthPickerHeader = document.createElementNS(HTML_NS, "div");
  monthPickerHeader.className = "fluent-month-picker-header";
  const yearPrevious = document.createElementNS(HTML_NS, "button");
  yearPrevious.id = "fluent-month-year-prev";
  yearPrevious.className = "fluent-month-year-button";
  yearPrevious.type = "button";
  yearPrevious.setAttribute("aria-label", "Previous year");
  yearPrevious.textContent = "\u2039";
  const yearLabel = document.createElementNS(HTML_NS, "span");
  yearLabel.id = "fluent-month-year-label";
  yearLabel.setAttribute("role", "heading");
  yearLabel.setAttribute("aria-level", "2");
  const yearNext = document.createElementNS(HTML_NS, "button");
  yearNext.id = "fluent-month-year-next";
  yearNext.className = "fluent-month-year-button";
  yearNext.type = "button";
  yearNext.setAttribute("aria-label", "Next year");
  yearNext.textContent = "\u203a";
  monthPickerHeader.append(yearPrevious, yearLabel, yearNext);

  const monthGrid = document.createElementNS(HTML_NS, "div");
  monthGrid.id = "fluent-month-grid";
  monthGrid.className = "fluent-month-grid";
  monthGrid.setAttribute("role", "group");
  monthGrid.setAttribute("aria-labelledby", yearLabel.id);
  const monthChoices = [];
  for (let month = 0; month < 12; month++) {
    const choice = document.createElementNS(HTML_NS, "button");
    choice.className = "fluent-month-choice";
    choice.type = "button";
    choice.dataset.month = String(month);
    choice.setAttribute("aria-pressed", "false");
    monthGrid.append(choice);
    monthChoices.push(choice);
  }

  const monthToday = document.createElementNS(HTML_NS, "button");
  monthToday.id = "fluent-month-today";
  monthToday.className = "fluent-month-today";
  monthToday.type = "button";
  monthToday.textContent = "Today";
  monthPicker.append(monthPickerHeader, monthGrid, monthToday);
  monthPanel.append(monthPicker);
  popupSet.append(monthPanel);

  const moreButton = document.createXULElement("toolbarbutton");
  moreButton.id = "fluent-today-more";
  moreButton.setAttribute("label", "\u2022\u2022\u2022");
  moreButton.setAttribute("tooltiptext", "Today pane options");
  moreButton.setAttribute("aria-label", "Today pane options");
  moreButton.setAttribute("aria-haspopup", "menu");
  moreButton.setAttribute("aria-expanded", "false");
  moreButton.setAttribute("aria-controls", "fluent-today-more-popup");
  minimonthHeader.append(moreButton);

  const morePopup = document.createXULElement("menupopup");
  morePopup.id = "fluent-today-more-popup";
  morePopup.className = "fluent-today-more-popup";

  const switchHeading = document.createXULElement("menuitem");
  switchHeading.id = "fluent-today-more-switch-heading";
  switchHeading.className = "fluent-today-more-heading";
  switchHeading.setAttribute("label", "Switch views");
  switchHeading.setAttribute("disabled", "true");

  const agendaItem = document.createXULElement("menuitem");
  agendaItem.id = "fluent-today-more-agenda";
  agendaItem.setAttribute("label", "Agenda");
  agendaItem.setAttribute("type", "radio");
  agendaItem.setAttribute("name", "fluent-today-view");
  agendaItem.setAttribute("checked", "true");
  agendaItem.setAttribute("autocheck", "false");

  const dayItem = document.createXULElement("menuitem");
  dayItem.id = "fluent-today-more-day";
  dayItem.setAttribute("label", "Day");
  dayItem.setAttribute("type", "radio");
  dayItem.setAttribute("name", "fluent-today-view");
  dayItem.setAttribute("autocheck", "false");

  const moreSeparator = document.createXULElement("menuseparator");
  moreSeparator.className = "fluent-today-more-separator";

  const calendarsHeading = document.createXULElement("menuitem");
  calendarsHeading.id = "fluent-today-more-calendars-heading";
  calendarsHeading.className = "fluent-today-more-heading";
  calendarsHeading.setAttribute("label", "Calendars");
  calendarsHeading.setAttribute("disabled", "true");

  const calendarItem = document.createXULElement("menuitem");
  calendarItem.id = "fluent-today-more-calendar";
  calendarItem.setAttribute("label", "Calendar");
  calendarItem.setAttribute("type", "checkbox");
  calendarItem.setAttribute("autocheck", "false");

  const tasksItem = document.createXULElement("menuitem");
  tasksItem.id = "fluent-today-more-tasks";
  tasksItem.setAttribute("label", "Tasks");
  tasksItem.setAttribute("type", "checkbox");
  tasksItem.setAttribute("autocheck", "false");

  const showAllItem = document.createXULElement("menuitem");
  showAllItem.id = "fluent-today-show-all";
  showAllItem.setAttribute("label", "Show all");

  morePopup.append(
    switchHeading,
    agendaItem,
    dayItem,
    moreSeparator,
    calendarsHeading,
    calendarItem,
    tasksItem,
    showAllItem
  );
  popupSet.append(morePopup);

  const intro = document.createXULElement("vbox");
  intro.id = "fluent-today-intro";
  const dateHeading = document.createXULElement("label");
  dateHeading.id = "fluent-today-date-heading";
  dateHeading.setAttribute("role", "heading");
  dateHeading.setAttribute("aria-level", "2");
  const quickTask = document.createElementNS(HTML_NS, "button");
  quickTask.className = "fluent-quick-task";
  quickTask.type = "button";
  quickTask.textContent = "Add a task due today";
  intro.append(dateHeading, quickTask);
  agendaContainer.insertBefore(intro, agenda);

  const monthNameFormatter = new window.Intl.DateTimeFormat(undefined, {
    month: "long",
  });
  const monthYearFormatter = new window.Intl.DateTimeFormat(undefined, {
    month: "long",
    year: "numeric",
  });
  const yearFormatter = new window.Intl.DateTimeFormat(undefined, {
    year: "numeric",
  });
  let pickerYear = getDisplayedMiniMonthDate(minimonth).getFullYear();

  const renderMonthPicker = () => {
    const displayed = getDisplayedMiniMonthDate(minimonth);
    const yearDate = new Date(pickerYear, 0, 1, 12);
    yearLabel.textContent = yearFormatter.format(yearDate);
    yearPrevious.setAttribute(
      "aria-label",
      `Show ${yearFormatter.format(new Date(pickerYear - 1, 0, 1, 12))}`
    );
    yearNext.setAttribute(
      "aria-label",
      `Show ${yearFormatter.format(new Date(pickerYear + 1, 0, 1, 12))}`
    );

    for (const [month, choice] of monthChoices.entries()) {
      const choiceDate = new Date(pickerYear, month, 1, 12);
      const selected =
        pickerYear === displayed.getFullYear() &&
        month === displayed.getMonth();
      choice.textContent = monthNameFormatter.format(choiceDate);
      choice.setAttribute("aria-label", monthYearFormatter.format(choiceDate));
      choice.setAttribute("aria-pressed", String(selected));
      choice.classList.toggle("selected", selected);
    }
  };

  const syncMonthButton = () => {
    const displayed = getDisplayedMiniMonthDate(minimonth);
    pickerYear = displayed.getFullYear();
    monthLabel.textContent = monthNameFormatter.format(displayed);
    monthButton.setAttribute(
      "aria-label",
      `Choose month, currently ${monthYearFormatter.format(displayed)}`
    );
    renderMonthPicker();
  };

  const onMonthButton = () => {
    if (monthPanel.state === "open" || monthPanel.state === "showing") {
      monthPanel.hidePopup();
      return;
    }
    syncMonthButton();
    /* Right-align the 296 px picker to its anchor so it remains inside the
     * minimum-width (320 px) Today pane at the edge of the app window. */
    monthPanel.openPopup(monthButton, "after_end", 0, 0, false, false);
  };
  const onMonthPopupShown = () => {
    monthButton.setAttribute("aria-expanded", "true");
    const selected = monthChoices.find(
      choice => choice.getAttribute("aria-pressed") === "true"
    );
    (selected || monthChoices[0])?.focus();
  };
  const onMonthPopupHidden = () => {
    monthButton.setAttribute("aria-expanded", "false");
  };
  const onPreviousYear = () => {
    pickerYear--;
    renderMonthPicker();
  };
  const onNextYear = () => {
    pickerYear++;
    renderMonthPicker();
  };
  const onMonthChoice = event => {
    const choice = event.target.closest?.(".fluent-month-choice");
    if (!choice || !monthGrid.contains(choice)) {
      return;
    }
    const month = Number(choice.dataset.month);
    if (!Number.isInteger(month) || month < 0 || month > 11) {
      return;
    }

    /* showMonth is the same operation as Thunderbird's native side arrows:
     * it navigates without changing the selected day or the agenda date. */
    minimonth.showMonth(new Date(pickerYear, month, 1, 12));
    monthPanel.hidePopup();
    monthButton.focus();
  };
  const onMonthPickerKeyDown = event => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      monthPanel.hidePopup();
      monthButton.focus();
      return;
    }
    const choice = event.target.closest?.(".fluent-month-choice");
    if (!choice || !monthGrid.contains(choice)) {
      return;
    }
    const index = monthChoices.indexOf(choice);
    let nextIndex = index;
    if (event.key === "ArrowLeft") {
      nextIndex = (index + 11) % 12;
    } else if (event.key === "ArrowRight") {
      nextIndex = (index + 1) % 12;
    } else if (event.key === "ArrowUp") {
      nextIndex = (index + 8) % 12;
    } else if (event.key === "ArrowDown") {
      nextIndex = (index + 4) % 12;
    } else if (event.key === "Home") {
      nextIndex = 0;
    } else if (event.key === "End") {
      nextIndex = 11;
    } else {
      return;
    }
    event.preventDefault();
    monthChoices[nextIndex].focus();
  };
  const onMonthToday = () => {
    /* Today intentionally selects today, matching Thunderbird's native Today
     * button, while ordinary month choices remain navigation-only. */
    minimonth.value = new Date();
    monthPanel.hidePopup();
    monthButton.focus();
  };

  const syncMorePopup = () => {
    const visibility = getTodayPaneSectionVisibility(window);
    setMenuItemChecked(agendaItem, true);
    setMenuItemChecked(dayItem, false);
    setMenuItemChecked(calendarItem, visibility.calendar);
    setMenuItemChecked(tasksItem, visibility.tasks);
    calendarItem.disabled = visibility.calendar && !visibility.tasks;
    tasksItem.disabled = visibility.tasks && !visibility.calendar;
  };

  const onMore = () => {
    if (morePopup.state === "open" || morePopup.state === "showing") {
      morePopup.hidePopup();
      return;
    }
    morePopup.openPopup(moreButton, "after_end", 0, 0, false, false);
  };
  const onMorePopupShowing = syncMorePopup;
  const onMorePopupShown = () => {
    moreButton.setAttribute("aria-expanded", "true");
  };
  const onMorePopupHidden = () => {
    moreButton.setAttribute("aria-expanded", "false");
  };
  const onAgendaView = () => {
    setMenuItemChecked(agendaItem, true);
  };
  const onDayView = () => {
    try {
      if (typeof window.switchCalendarView === "function") {
        window.switchCalendarView("day", true);
      }
      document.getElementById("tabmail")?.openTab("calendar");
    } catch (error) {
      console.error(
        "Outlook Style Companion could not open Calendar Day view:",
        error
      );
      document.getElementById("tabmail")?.openTab("calendar");
    }
  };
  const onToggleCalendar = () => {
    const visibility = getTodayPaneSectionVisibility(window);
    setTodayPaneSectionVisibility(
      window,
      !visibility.calendar,
      visibility.tasks
    );
    syncMorePopup();
  };
  const onToggleTasks = () => {
    const visibility = getTodayPaneSectionVisibility(window);
    setTodayPaneSectionVisibility(
      window,
      visibility.calendar,
      !visibility.tasks
    );
    syncMorePopup();
  };
  const onShowAll = () => {
    setTodayPaneSectionVisibility(window, true, true);
    syncMorePopup();
  };

  const onCalendar = () => setTodayPaneMode(window, "calendar");
  const onTodo = () => setTodayPaneMode(window, "todo");
  const onPopout = () => document.getElementById("tabmail")?.openTab("calendar");
  const onQuickTask = () =>
    window.goDoCommand("calendar_new_todo_todaypane_command");
  let decorationFrame = 0;
  const scheduleDecoration = () => {
    if (decorationFrame) {
      return;
    }
    decorationFrame = window.requestAnimationFrame(() => {
      decorationFrame = 0;
      decorateAgenda(window);
    });
  };
  const tabObserver = new window.MutationObserver(() =>
    syncTodayPaneTabs(window)
  );
  const agendaObserver = new window.MutationObserver(scheduleDecoration);
  todayPaneState.set(window, {
    agendaObserver,
    tabObserver,
    calendarTab,
    todoTab,
    popout,
    intro,
    minimonth,
    nativeMonthInput,
    nativeMonthInputWasHidden,
    nativeMonthInputAriaHidden,
    monthButton,
    monthPanel,
    yearPrevious,
    yearNext,
    monthGrid,
    monthPicker,
    monthToday,
    moreButton,
    morePopup,
    agendaItem,
    dayItem,
    calendarItem,
    tasksItem,
    showAllItem,
    quickTask,
    onCalendar,
    onTodo,
    onPopout,
    onQuickTask,
    onMonthButton,
    onMonthPopupShown,
    onMonthPopupHidden,
    onPreviousYear,
    onNextYear,
    onMonthChoice,
    onMonthPickerKeyDown,
    onMonthToday,
    syncMonthButton,
    onMore,
    onMorePopupShowing,
    onMorePopupShown,
    onMorePopupHidden,
    onAgendaView,
    onDayView,
    onToggleCalendar,
    onToggleTasks,
    onShowAll,
    scheduleDecoration,
    metronomeCallback: scheduleDecoration,
    get decorationFrame() {
      return decorationFrame;
    },
  });
  enhancedTodayPaneWindows.add(window);

  calendarTab.addEventListener("command", onCalendar);
  todoTab.addEventListener("command", onTodo);
  popout.addEventListener("command", onPopout);
  quickTask.addEventListener("click", onQuickTask);
  monthButton.addEventListener("click", onMonthButton);
  monthPanel.addEventListener("popupshown", onMonthPopupShown);
  monthPanel.addEventListener("popuphidden", onMonthPopupHidden);
  yearPrevious.addEventListener("click", onPreviousYear);
  yearNext.addEventListener("click", onNextYear);
  monthGrid.addEventListener("click", onMonthChoice);
  monthPicker.addEventListener("keydown", onMonthPickerKeyDown);
  monthToday.addEventListener("click", onMonthToday);
  moreButton.addEventListener("command", onMore);
  morePopup.addEventListener("popupshowing", onMorePopupShowing);
  morePopup.addEventListener("popupshown", onMorePopupShown);
  morePopup.addEventListener("popuphidden", onMorePopupHidden);
  agendaItem.addEventListener("command", onAgendaView);
  dayItem.addEventListener("command", onDayView);
  calendarItem.addEventListener("command", onToggleCalendar);
  tasksItem.addEventListener("command", onToggleTasks);
  showAllItem.addEventListener("command", onShowAll);
  tabObserver.observe(nativeHeader, {
    attributes: true,
    attributeFilter: ["index"],
  });
  agendaObserver.observe(agenda, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: true,
    attributeFilter: ["hidden", "class", "status"],
  });
  CalMetronome.on("minute", scheduleDecoration);
  minimonth.addEventListener("change", scheduleDecoration);
  minimonth.addEventListener("monthchange", syncMonthButton);
  syncTodayPaneTabs(window);
  syncMonthButton();
  scheduleDecoration();
  return true;
}

function removeTodayPaneEnhancement(window) {
  const state = todayPaneState.get(window);
  if (!state) {
    return;
  }

  state.agendaObserver.disconnect();
  state.tabObserver.disconnect();
  CalMetronome.off("minute", state.metronomeCallback);
  if (state.decorationFrame) {
    window.cancelAnimationFrame(state.decorationFrame);
  }
  state.calendarTab.removeEventListener("command", state.onCalendar);
  state.todoTab.removeEventListener("command", state.onTodo);
  state.popout.removeEventListener("command", state.onPopout);
  state.quickTask.removeEventListener("click", state.onQuickTask);
  state.minimonth.removeEventListener("change", state.scheduleDecoration);
  state.minimonth.removeEventListener("monthchange", state.syncMonthButton);
  state.monthButton.removeEventListener("click", state.onMonthButton);
  state.monthPanel.removeEventListener("popupshown", state.onMonthPopupShown);
  state.monthPanel.removeEventListener("popuphidden", state.onMonthPopupHidden);
  state.yearPrevious.removeEventListener("click", state.onPreviousYear);
  state.yearNext.removeEventListener("click", state.onNextYear);
  state.monthGrid.removeEventListener("click", state.onMonthChoice);
  state.monthPicker.removeEventListener(
    "keydown",
    state.onMonthPickerKeyDown
  );
  state.monthToday.removeEventListener("click", state.onMonthToday);
  state.moreButton.removeEventListener("command", state.onMore);
  state.morePopup.removeEventListener(
    "popupshowing",
    state.onMorePopupShowing
  );
  state.morePopup.removeEventListener("popupshown", state.onMorePopupShown);
  state.morePopup.removeEventListener("popuphidden", state.onMorePopupHidden);
  state.agendaItem.removeEventListener("command", state.onAgendaView);
  state.dayItem.removeEventListener("command", state.onDayView);
  state.calendarItem.removeEventListener("command", state.onToggleCalendar);
  state.tasksItem.removeEventListener("command", state.onToggleTasks);
  state.showAllItem.removeEventListener("command", state.onShowAll);
  state.monthPanel.hidePopup?.();
  state.morePopup.hidePopup?.();
  state.calendarTab.parentElement?.remove();
  state.popout.remove();
  state.intro.remove();
  state.monthButton.remove();
  state.monthPanel.remove();
  state.moreButton.remove();
  state.morePopup.remove();
  if (state.nativeMonthInputWasHidden) {
    state.nativeMonthInput.setAttribute("hidden", "true");
  } else {
    state.nativeMonthInput.removeAttribute("hidden");
  }
  if (state.nativeMonthInputAriaHidden === null) {
    state.nativeMonthInput.removeAttribute("aria-hidden");
  } else {
    state.nativeMonthInput.setAttribute(
      "aria-hidden",
      state.nativeMonthInputAriaHidden
    );
  }

  const document = window.document;
  document
    .querySelector("#today-pane-panel > .sidebar-header")
    ?.classList.remove("fluent-myday-enhanced");
  document.getElementById("today-pane-panel")?.removeAttribute("data-fluent-myday");
  for (const element of document.querySelectorAll(
    "#agenda .fluent-event-duration, #agenda .fluent-time-marker"
  )) {
    element.remove();
  }
  for (const listItem of document.querySelectorAll("#agenda .agenda-listitem")) {
    listItem.removeAttribute("data-fluent-hide-date-header");
    listItem.removeAttribute("data-fluent-next-event");
    listItem.removeAttribute("data-fluent-time-state");
    listItem.removeAttribute("data-fluent-time-attention");
    if (listItem.dateString) {
      listItem.dateString = listItem.dateString;
    }
  }

  todayPaneState.delete(window);
  enhancedTodayPaneWindows.delete(window);
}

function rollbackTodayPaneEnhancement(window) {
  if (todayPaneState.has(window)) {
    removeTodayPaneEnhancement(window);
  } else {
    removeStaleTodayPaneMarkup(window);
  }
}

function installTodayPaneEnhancement() {
  if (todayPaneListenerRegistered) {
    return true;
  }
  const enhanceWindow = window =>
    enhanceWindowSafely(
      "a Today pane window",
      window,
      enhanceTodayPane,
      rollbackTodayPaneEnhancement
    );
  ExtensionSupport.registerWindowListener(TODAY_PANE_LISTENER_ID, {
    chromeURLs: [MESSENGER_WINDOW_URL],
    onLoadWindow: enhanceWindow,
    onUnloadWindow: removeTodayPaneEnhancement,
  });
  todayPaneListenerRegistered = true;
  /* The extension background starts inside an already-open messenger window.
   * ExtensionSupport does not consistently replay that window to a newly
   * registered listener in all Thunderbird startup paths, so cover it here. */
  const windows = Services.wm.getEnumerator("mail:3pane");
  while (windows.hasMoreElements()) {
    const window = windows.getNext();
    enhanceWindowSafely("an existing mail window", window, currentWindow => {
      if (currentWindow.document?.location?.href !== MESSENGER_WINDOW_URL) {
        return false;
      }
      enhanceWindow(currentWindow);
      return enhanceWindowSafely(
        "pane widths for a mail window",
        currentWindow,
        scheduleOutlookPaneWidths
      );
    });
  }
  return todayPaneListenerRegistered;
}

function installMainPaneSchemeEnhancement() {
  if (mainPaneSchemeListenerRegistered) {
    return true;
  }
  const enhanceWindow = window =>
    enhanceWindowSafely(
      "a main-pane scheme window",
      window,
      installMainPaneSchemeBridge,
      removeMainPaneSchemeBridge
    );
  ExtensionSupport.registerWindowListener(MAIN_PANE_SCHEME_LISTENER_ID, {
    chromeURLs: [MESSENGER_WINDOW_URL],
    onLoadWindow: enhanceWindow,
    onUnloadWindow: removeMainPaneSchemeBridge,
  });
  mainPaneSchemeListenerRegistered = true;
  const windows = Services.wm.getEnumerator("mail:3pane");
  while (windows.hasMoreElements()) {
    const window = windows.getNext();
    enhanceWindowSafely("an existing mail window", window, currentWindow => {
      if (currentWindow.document?.location?.href !== MESSENGER_WINDOW_URL) {
        return false;
      }
      return enhanceWindow(currentWindow);
    });
  }
  return mainPaneSchemeListenerRegistered;
}

function getMeetingStartSnooze(window, widget) {
  try {
    const item = widget?.item;
    if (
      !item ||
      !widget.alarm ||
      typeof item.isEvent !== "function" ||
      !item.isEvent() ||
      !item.startDate ||
      item.startDate.isDate ||
      !cal.acl.isCalendarWritable(item.calendar) ||
      !cal.acl.userCanModifyItem(item) ||
      widget.querySelector(".alarm-snooze-button")?.disabled
    ) {
      return null;
    }

    /* A floating calendar time represents wall-clock time in Thunderbird's
     * current calendar timezone. Resolve that meaning before comparing it
     * with an absolute instant so DST and provider timezones remain correct. */
    let start = item.startDate;
    if (start.timezone?.isFloating) {
      start = start.getInTimezone(cal.dtz.defaultTimezone);
    }
    const displayStart = start.getInTimezone(cal.dtz.defaultTimezone);
    const startUtc = start.getInTimezone(cal.dtz.UTC);
    const nowUtc = cal.dtz.now().getInTimezone(cal.dtz.UTC);
    const duration = startUtc.subtractDate(nowUtc);
    if (!Number.isFinite(duration.inSeconds) || duration.inSeconds <= 0) {
      return null;
    }

    return {
      item,
      alarm: widget.alarm,
      duration,
      label: `Until meeting starts \u2014 ${cal.dtz.formatter.formatTime(
        displayStart
      )}`,
    };
  } catch (error) {
    console.error(
      "Outlook Style Companion could not inspect a meeting start time:",
      error
    );
    return null;
  }
}

function removeReminderWidgetDecoration(entry) {
  try {
    entry.popup.removeEventListener("popupshowing", entry.onPopupShowing);
  } catch (error) {
    /* A closing reminder can destroy its popup before owned cleanup. */
  }
  try {
    entry.menuItem.removeEventListener("command", entry.onCommand);
  } catch (error) {
    /* Continue removing the remaining owned nodes. */
  }
  try {
    entry.menuItem.remove();
  } catch (error) {
    /* The menu item may already have been removed by native teardown. */
  }
  try {
    entry.separator.remove();
  } catch (error) {
    /* The separator may already have been removed by native teardown. */
  }
}

function decorateReminderWidgets(window) {
  const state = reminderDialogState.get(window);
  if (!state) {
    return;
  }

  for (const [widget, entry] of state.widgets) {
    if (!widget.isConnected) {
      removeReminderWidgetDecoration(entry);
      state.widgets.delete(widget);
    }
  }

  for (const widget of state.richlist.querySelectorAll(
    'richlistitem[is="calendar-alarm-widget-richlistitem"]'
  )) {
    if (state.widgets.has(widget)) {
      continue;
    }

    const popup = widget.querySelector(
      '.alarm-snooze-button > menupopup[is="calendar-snooze-popup"]'
    );
    const firstPreset = popup?.querySelector(":scope > menuitem[value]");
    if (!popup || !firstPreset) {
      continue;
    }

    const menuItem = window.document.createXULElement("menuitem");
    menuItem.className = "outlook-snooze-until-start";
    menuItem.setAttribute("closemenu", "single");
    menuItem.setAttribute("label", "Until meeting starts");

    const separator = window.document.createXULElement("menuseparator");
    separator.className = "outlook-snooze-until-start-separator";

    const updateMenuItem = () => {
      const details = getMeetingStartSnooze(window, widget);
      menuItem.hidden = !details;
      separator.hidden = !details;
      if (details) {
        menuItem.disabled = false;
        menuItem.setAttribute("label", details.label);
      }
    };

    const onCommand = event => {
      event.preventDefault();
      event.stopPropagation();
      const details = getMeetingStartSnooze(window, widget);
      if (!details) {
        updateMenuItem();
        popup.hidePopup();
        return;
      }

      /* Use the same safety limit and native alarm service as Thunderbird.
       * Passing a calIDuration retains seconds, unlike the built-in menu's
       * minute-only values, so the snooze lands on the meeting start. */
      if (
        typeof window.aboveSnoozeLimit === "function" &&
        window.aboveSnoozeLimit(details.duration)
      ) {
        return;
      }

      popup.hidePopup();
      try {
        const alarmService = Cc[
          "@mozilla.org/calendar/alarm-service;1"
        ].getService(Ci.calIAlarmService);
        const operation = alarmService.snoozeAlarm(
          details.item,
          details.alarm,
          details.duration
        );
        Promise.resolve(operation).catch(error => {
          console.error(
            "Outlook Style Companion could not snooze until the meeting starts:",
            error
          );
        });
      } catch (error) {
        console.error(
          "Outlook Style Companion could not snooze until the meeting starts:",
          error
        );
      }
    };

    const entry = {
      popup,
      menuItem,
      separator,
      onPopupShowing: updateMenuItem,
      onCommand,
    };
    state.widgets.set(widget, entry);
    try {
      popup.insertBefore(menuItem, firstPreset);
      popup.insertBefore(separator, firstPreset);
      popup.addEventListener("popupshowing", updateMenuItem);
      menuItem.addEventListener("command", onCommand);
      updateMenuItem();
    } catch (error) {
      state.widgets.delete(widget);
      removeReminderWidgetDecoration(entry);
      console.error(
        "Outlook Style Companion could not enhance a reminder widget:",
        error
      );
    }
  }
}

function bodyHasExplicitBackground(body) {
  return Boolean(
    body.hasAttribute("bgcolor") ||
      body.hasAttribute("background") ||
      body.style.background ||
      body.style.backgroundColor ||
      body.style.backgroundImage
  );
}

function bodyHasExplicitColor(body) {
  return Boolean(body.hasAttribute("text") || body.style.color);
}

/**
 * Make an editor content document follow the system light/dark scheme. These
 * documents are about:blank browsing contexts, so a theme-experiment
 * stylesheet cannot reach them. Compose documents retain an explicitly
 * authored body color or background (for example stationery or a draft); the
 * calendar description editor always receives Thunderbird's neutral adaptive
 * editing canvas. CSS media queries update an open editor immediately when the
 * system scheme changes, without rebuilding its document.
 *
 * @param {Document} document - The editor's content document.
 * @param {boolean} preserveExplicitBodyColors - Whether authored body colors
 *   should take precedence over the Outlook defaults.
 * @returns {?Function} A cleanup function, or null if the document is not ready.
 */
function applyEditorDocumentSurface(document, preserveExplicitBodyColors) {
  const root = document?.documentElement;
  const body = document?.body;
  if (!root || !body) {
    return null;
  }

  const CSSStyleSheet = document.defaultView?.CSSStyleSheet;
  if (
    typeof CSSStyleSheet !== "function" ||
    !("adoptedStyleSheets" in document)
  ) {
    return null;
  }
  let styleSheet;
  try {
    styleSheet = new CSSStyleSheet();
    document.adoptedStyleSheets = [
      ...document.adoptedStyleSheets,
      styleSheet,
    ];
  } catch (error) {
    /* Fail closed if this editor does not support non-serializing sheets. */
    return null;
  }

  const updateBodyDefaults = () => {
    const useDefaultBackground =
      !preserveExplicitBodyColors || !bodyHasExplicitBackground(body);
    const useDefaultColor =
      !preserveExplicitBodyColors || !bodyHasExplicitColor(body);
    const backgroundSelector = useDefaultBackground ? "body" : "body:not(body)";
    const colorSelector = useDefaultColor ? "body" : "body:not(body)";
    styleSheet.replaceSync(
      EDITOR_SURFACE_CSS.replaceAll(
        "__OUTLOOK_EDITOR_BACKGROUND_SELECTOR__",
        backgroundSelector
      ).replaceAll("__OUTLOOK_EDITOR_COLOR_SELECTOR__", colorSelector)
    );
  };
  let observer = null;
  const cleanup = () => {
    observer?.disconnect();
    try {
      document.adoptedStyleSheets = document.adoptedStyleSheets.filter(
        candidate => candidate !== styleSheet
      );
    } catch (error) {
      /* A closing editor can destroy its document before window cleanup. */
    }
  };
  try {
    updateBodyDefaults();
    if (preserveExplicitBodyColors) {
      observer = new document.defaultView.MutationObserver(updateBodyDefaults);
      observer.observe(body, {
        attributes: true,
        attributeFilter: ["style", "bgcolor", "background", "text"],
      });
    }
  } catch (error) {
    cleanup();
    throw error;
  }

  return cleanup;
}

/**
 * Track a XUL editor across its about:blank document loads.
 *
 * @param {Element} editor - Thunderbird's editor element.
 * @param {boolean} preserveExplicitBodyColors - Preserve authored body colors.
 * @returns {Function} Cleanup function.
 */
function trackEditorSurface(editor, preserveExplicitBodyColors) {
  let removeDocumentSurface = null;

  const applyCurrentDocument = () => {
    removeDocumentSurface?.();
    removeDocumentSurface = applyEditorDocumentSurface(
      editor.contentDocument,
      preserveExplicitBodyColors
    );
  };
  const onLoad = event => {
    const currentDocument = editor.contentDocument;
    if (
      event.target === editor ||
      event.target === currentDocument ||
      event.originalTarget === currentDocument
    ) {
      applyCurrentDocument();
    }
  };

  const cleanup = () => {
    editor.removeEventListener("load", onLoad, true);
    removeDocumentSurface?.();
    removeDocumentSurface = null;
  };
  try {
    editor.addEventListener("load", onLoad, true);
    applyCurrentDocument();
  } catch (error) {
    cleanup();
    throw error;
  }

  return cleanup;
}

function enhanceComposeEditorSurface(window, cleanups) {
  const document = window.document;
  let editor = null;
  let removeEditorTracking = null;
  let removeMentions = null;

  const bindEditor = () => {
    const nextEditor = document.getElementById("messageEditor");
    if (nextEditor === editor) {
      return;
    }
    removeEditorTracking?.();
    removeMentions?.();
    editor = nextEditor;
    removeEditorTracking = editor
      ? trackEditorSurface(editor, true)
      : null;
    removeMentions = editor ? installComposeMentionAutocomplete(editor) : null;
  };

  const observer = new window.MutationObserver(bindEditor);
  cleanups.push(() => {
    observer.disconnect();
    removeEditorTracking?.();
    removeEditorTracking = null;
    removeMentions?.();
    removeMentions = null;
    editor = null;
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  bindEditor();
}

/**
 * Return local address-book entries matching a compose @-mention query. This
 * deliberately uses Thunderbird's address-book service, so it neither reads
 * messages nor sends contact data anywhere.
 *
 * @param {string} query - The text after the @ sign.
 * @returns {{name: string, email: string}[]} Matching contacts.
 */
function getComposeMentionContacts(query) {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const contacts = new Map();
  for (const directory of Array.from(MailServices.ab.directories || [])) {
    if (directory.isMailList) {
      continue;
    }
    let cards;
    try {
      cards = Array.from(directory.childCards || []);
    } catch (error) {
      /* An unavailable directory must not prevent other address books from
       * supplying a compose suggestion. */
      continue;
    }
    for (const card of cards) {
      if (card.isMailList) {
        continue;
      }
      const name = String(
        card.displayName || card.getProperty?.("DisplayName", "") || ""
      ).trim();
      const emails = [
        card.primaryEmail,
        card.getProperty?.("SecondEmail", ""),
      ];
      for (const candidate of emails) {
        const email = String(candidate || "").trim();
        if (!email) {
          continue;
        }
        const label = name || email;
        if (
          normalizedQuery &&
          !label.toLocaleLowerCase().includes(normalizedQuery) &&
          !email.toLocaleLowerCase().includes(normalizedQuery)
        ) {
          continue;
        }
        contacts.set(email.toLocaleLowerCase(), { name: label, email });
      }
    }
  }
  return [...contacts.values()]
    .sort((first, second) =>
      first.name.localeCompare(second.name, undefined, { sensitivity: "base" })
    )
    .slice(0, 8);
}

/**
 * Add an Outlook-style contact picker to the compose editor. Thunderbird's
 * normal address autocomplete only operates in recipient fields; this makes
 * @ work in the message body for new messages and replies as well.
 *
 * @param {Element} editor - Thunderbird's compose editor element.
 * @returns {Function} A cleanup function.
 */
function installComposeMentionAutocomplete(editor) {
  /* XUL's compose editor in Thunderbird 154 does not always expose
   * `ownerGlobal`, even though it belongs to the compose chrome document.
   * Resolve the owning window through that document as well so the mention
   * feature can attach on both supported Thunderbird versions. */
  const window = editor.ownerGlobal || editor.ownerDocument?.defaultView;
  if (!window) {
    return () => {};
  }
  const chromeDocument = window.document;
  const popupSet = chromeDocument.getElementById("mainPopupSet");
  if (!popupSet) {
    return () => {};
  }

  let editorDocument = null;
  let popup = null;
  let list = null;
  let mentionRange = null;
  let contacts = [];
  let activeIndex = 0;
  let hideTimer = 0;
  let handledKeyDown = "";
  const handledPickerKeyEvents = new WeakSet();
  const systemKeyListenerOptions = {
    capture: true,
    mozSystemGroup: true,
  };
  let suppressSelectionUpdate = false;
  let suppressSelectionTimer = 0;

  const isOpen = () => popup?.state === "open" || popup?.state === "showing";
  const hide = () => {
    if (hideTimer) {
      window.clearTimeout(hideTimer);
      hideTimer = 0;
    }
    if (suppressSelectionTimer) {
      window.clearTimeout(suppressSelectionTimer);
      suppressSelectionTimer = 0;
    }
    suppressSelectionUpdate = false;
    handledKeyDown = "";
    mentionRange = null;
    contacts = [];
    activeIndex = 0;
    try {
      popup?.hidePopup();
    } catch (error) {
      /* A closing compose window can tear down its popup first. */
    }
  };

  const getMentionContext = () => {
    const selection = editor.contentWindow?.getSelection?.();
    const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
    const Node = editorDocument?.defaultView?.Node;
    if (
      !range?.collapsed ||
      !Node ||
      range.startContainer.nodeType !== Node.TEXT_NODE
    ) {
      return null;
    }
    /* A mention's visible name is intentionally editable. Never reopen the
     * contact picker while the caret is inside an existing mention link; its
     * mailto target and stored email remain unchanged as the label is shortened. */
    if (
      range.startContainer.parentElement?.closest?.(
        "a.outlook-compose-mention[data-outlook-mention-email]"
      )
    ) {
      return null;
    }
    const beforeCaret = range.startContainer.data.slice(0, range.startOffset);
    const match = beforeCaret.match(/(?:^|\s|[([{"'.,;:!?])@([^\s@]*)$/u);
    if (!match) {
      return null;
    }
    const tokenOffset = match.index + match[0].lastIndexOf("@");
    const tokenRange = editorDocument.createRange();
    tokenRange.setStart(range.startContainer, tokenOffset);
    tokenRange.setEnd(range.startContainer, range.startOffset);
    return { query: match[1], range: tokenRange };
  };

  const render = () => {
    if (!list) {
      return;
    }
    list.replaceChildren();
    if (!contacts.length) {
      const empty = chromeDocument.createElementNS(HTML_NS, "div");
      empty.className = "outlook-compose-mention-empty";
      empty.setAttribute("role", "status");
      empty.textContent = "No matching contacts";
      list.appendChild(empty);
      return;
    }
    contacts.forEach((contact, index) => {
      const item = chromeDocument.createElementNS(HTML_NS, "button");
      const name = chromeDocument.createElementNS(HTML_NS, "span");
      const email = chromeDocument.createElementNS(HTML_NS, "span");
      item.className = "outlook-compose-mention-item";
      item.type = "button";
      item.setAttribute("role", "option");
      item.setAttribute("tooltiptext", contact.email);
      item.setAttribute("aria-label", `Mention ${contact.name}, ${contact.email}`);
      item.setAttribute("aria-selected", String(index === activeIndex));
      if (index === activeIndex) {
        item.setAttribute("data-active", "true");
      }
      name.className = "outlook-compose-mention-name";
      name.textContent = contact.name;
      email.className = "outlook-compose-mention-email";
      email.textContent = contact.email;
      item.addEventListener("mousedown", event => event.preventDefault());
      item.addEventListener("click", () => choose(index));
      item.append(name, email);
      list.appendChild(item);
    });
    list
      .querySelector('[data-active="true"]')
      ?.scrollIntoView({ block: "nearest" });
  };

  const getMentionLink = selection => {
    const Node = editorDocument?.defaultView?.Node;
    const node = selection?.anchorNode;
    if (!Node || !node) {
      return null;
    }
    const element =
      node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
    return element?.closest?.("a[href]") || null;
  };

  const choose = index => {
    const contact = contacts[index];
    if (!contact || !mentionRange || !editorDocument) {
      hide();
      return;
    }
    const selection = editor.contentWindow?.getSelection?.();
    try {
      selection.removeAllRanges();
      selection.addRange(mentionRange);
      const mention = `@${contact.name}`;
      let insertedText = null;
      if (editorDocument.execCommand("insertText", false, mention)) {
        const insertedRange = selection.rangeCount
          ? selection.getRangeAt(0)
          : null;
        const Node = editorDocument.defaultView?.Node;
        if (
          Node &&
          insertedRange?.collapsed &&
          insertedRange.startContainer.nodeType === Node.TEXT_NODE &&
          insertedRange.startOffset >= mention.length
        ) {
          insertedText = editorDocument.createRange();
          insertedText.setStart(
            insertedRange.startContainer,
            insertedRange.startOffset - mention.length
          );
          insertedText.setEnd(
            insertedRange.startContainer,
            insertedRange.startOffset
          );
        }
      } else {
        mentionRange.deleteContents();
        const text = editorDocument.createTextNode(mention);
        mentionRange.insertNode(text);
        insertedText = editorDocument.createRange();
        insertedText.selectNodeContents(text);
        selection.removeAllRanges();
        selection.addRange(insertedText);
      }
      if (insertedText) {
        selection.removeAllRanges();
        selection.addRange(insertedText);
        editorDocument.execCommand(
          "createLink",
          false,
          `mailto:${encodeURIComponent(contact.email)}`
        );
      }
      const link = getMentionLink(selection);
      if (link) {
        link.classList.add("outlook-compose-mention");
        link.setAttribute("title", contact.email);
        link.setAttribute("data-outlook-mention-email", contact.email);
        Object.assign(link.style, {
          backgroundColor: "#e6f2ff",
          borderRadius: "3px",
          color: "#0f6cbd",
          fontWeight: "600",
          padding: "1px 3px",
          textDecoration: "none",
        });
        const afterLink = editorDocument.createRange();
        afterLink.setStartAfter(link);
        afterLink.collapse(true);
        selection.removeAllRanges();
        selection.addRange(afterLink);
        if (!editorDocument.execCommand("insertText", false, " ")) {
          const space = editorDocument.createTextNode(" ");
          afterLink.insertNode(space);
          afterLink.setStartAfter(space);
          afterLink.collapse(true);
          selection.removeAllRanges();
          selection.addRange(afterLink);
        }
      }
      editor.contentWindow?.focus();
    } finally {
      hide();
    }
  };

  const update = () => {
    const context = getMentionContext();
    if (!context) {
      hide();
      return;
    }
    mentionRange = context.range;
    contacts = getComposeMentionContacts(context.query);
    activeIndex = 0;
    render();
    if (!isOpen()) {
      try {
        /* `after_start` anchors below the entire editor element. In a compose
         * window that editor fills the remaining height, so the suggestion
         * panel would normally open beyond the bottom edge of the window.
         * `overlap` starts at the editor's top-left instead; the range rect is
         * in the editor document's viewport, which gives us a stable position
         * immediately beneath the @ token. */
        const tokenRect = context.range.getBoundingClientRect();
        popup.openPopup(
          editor,
          "overlap",
          Math.max(0, Math.round(tokenRect.right)),
          Math.max(0, Math.round(tokenRect.bottom)),
          false,
          false
        );
      } catch (error) {
        hide();
      }
    }
  };

  const onInput = () => update();
  const onSelectionChange = () => {
    if (isOpen() && !suppressSelectionUpdate) {
      update();
    }
  };
  const suppressSelectionSync = () => {
    suppressSelectionUpdate = true;
    if (suppressSelectionTimer) {
      window.clearTimeout(suppressSelectionTimer);
    }
    suppressSelectionTimer = window.setTimeout(() => {
      suppressSelectionTimer = 0;
      suppressSelectionUpdate = false;
    }, 0);
  };
  const getMentionAtBackspaceCaret = selection => {
    const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
    const Node = editorDocument?.defaultView?.Node;
    if (!range?.collapsed || !Node) {
      return null;
    }
    const container = range.startContainer;
    const parentElement =
      container.nodeType === Node.ELEMENT_NODE
        ? container
        : container.parentElement;
    const containingMention = parentElement?.closest?.(
      "a.outlook-compose-mention[data-outlook-mention-email]"
    );
    if (containingMention) {
      return containingMention;
    }

    let candidate = null;
    if (container.nodeType === Node.TEXT_NODE) {
      if (container.data.slice(0, range.startOffset).trim()) {
        return null;
      }
      candidate = container.previousSibling;
    } else if (range.startOffset > 0) {
      candidate = container.childNodes[range.startOffset - 1];
    }
    while (candidate?.nodeType === Node.TEXT_NODE && !candidate.data.trim()) {
      candidate = candidate.previousSibling;
    }
    return candidate?.matches?.(
      "a.outlook-compose-mention[data-outlook-mention-email]"
    )
      ? candidate
      : null;
  };
  const handleMentionBackspace = event => {
    if (
      event.key !== "Backspace" ||
      event.altKey ||
      event.ctrlKey ||
      event.metaKey ||
      event.shiftKey ||
      event.isComposing ||
      !editorDocument
    ) {
      return false;
    }
    const selection = editor.contentWindow?.getSelection?.();
    const link = getMentionAtBackspaceCaret(selection);
    const textNode = link?.firstChild;
    const Node = editorDocument.defaultView?.Node;
    if (!link || !Node || textNode?.nodeType !== Node.TEXT_NODE) {
      return false;
    }
    const trailingWord = textNode.data.match(/\s+\S+$/u);
    if (!trailingWord) {
      return false;
    }

    event.preventDefault();
    event.stopPropagation();
    const deletion = editorDocument.createRange();
    deletion.setStart(textNode, trailingWord.index);
    deletion.setEnd(textNode, textNode.data.length);
    selection.removeAllRanges();
    selection.addRange(deletion);
    if (!editorDocument.execCommand("delete", false)) {
      textNode.deleteData(trailingWord.index, trailingWord[0].length);
    }
    const afterMention = editorDocument.createRange();
    const trailingText = link.nextSibling;
    if (
      trailingText?.nodeType === Node.TEXT_NODE &&
      trailingText.data.startsWith(" ")
    ) {
      afterMention.setStart(trailingText, 1);
    } else {
      afterMention.setStartAfter(link);
    }
    afterMention.collapse(true);
    selection.removeAllRanges();
    selection.addRange(afterMention);
    editor.contentWindow?.focus();
    return true;
  };
  const handlePickerKey = event => {
    if (!isOpen() || !contacts.length || event.isComposing) {
      return false;
    }
    if (
      ![
        "ArrowDown",
        "ArrowUp",
        "Enter",
        "Tab",
        "Escape",
      ].includes(event.key)
    ) {
      return false;
    }
    event.preventDefault();
    event.stopPropagation();
    suppressSelectionSync();
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      activeIndex =
        (activeIndex + (event.key === "ArrowDown" ? 1 : -1) + contacts.length) %
        contacts.length;
      render();
    } else if (event.key === "Enter" || event.key === "Tab") {
      choose(activeIndex);
    } else {
      hide();
    }
    return true;
  };
  const onKeyDown = event => {
    /* Thunderbird's editor can mark navigation keys handled before add-on
     * listeners receive them. Deliberately handle those events anyway, then
     * suppress its matching selectionchange so the highlighted result is not
     * reset to the first row. This uses standard DOM key values and therefore
     * works on Windows, macOS, and Linux. */
    if (handledPickerKeyEvents.has(event)) {
      return;
    }
    if (handleMentionBackspace(event)) {
      handledPickerKeyEvents.add(event);
      handledKeyDown = event.key;
      return;
    }
    if (handlePickerKey(event)) {
      handledPickerKeyEvents.add(event);
      handledKeyDown = event.key;
    }
  };
  const onKeyUp = event => {
    if (handledPickerKeyEvents.has(event)) {
      return;
    }
    handledPickerKeyEvents.add(event);
    if (event.key === handledKeyDown) {
      handledKeyDown = "";
      return;
    }
    handlePickerKey(event);
  };
  const onBlur = () => {
    hideTimer = window.setTimeout(hide, 150);
  };
  const onFocus = () => {
    if (hideTimer) {
      window.clearTimeout(hideTimer);
      hideTimer = 0;
    }
  };
  const onPopupHidden = () => {
    mentionRange = null;
  };
  const removeDocumentListeners = () => {
    editorDocument?.removeEventListener("input", onInput, true);
    editorDocument?.removeEventListener("keydown", onKeyDown, true);
    editorDocument?.removeEventListener("keyup", onKeyUp, true);
    editorDocument?.defaultView?.removeEventListener(
      "keydown",
      onKeyDown,
      systemKeyListenerOptions
    );
    editorDocument?.defaultView?.removeEventListener(
      "keyup",
      onKeyUp,
      systemKeyListenerOptions
    );
    editorDocument?.removeEventListener("selectionchange", onSelectionChange);
    editorDocument?.defaultView?.removeEventListener("blur", onBlur);
    editorDocument?.defaultView?.removeEventListener("focus", onFocus);
    editorDocument = null;
  };
  const bindEditorDocument = () => {
    removeDocumentListeners();
    const nextDocument = editor.contentDocument;
    if (!nextDocument?.body) {
      return;
    }
    editorDocument = nextDocument;
    editorDocument.addEventListener("input", onInput, true);
    editorDocument.addEventListener("keydown", onKeyDown, true);
    editorDocument.addEventListener("keyup", onKeyUp, true);
    /* Gecko's HTML editor handles physical arrow keys in its privileged system
     * event group. Listen at the editor window's capture phase in that same
     * group so real keyboards reach the picker on macOS, Windows, and Linux. */
    editorDocument.defaultView?.addEventListener(
      "keydown",
      onKeyDown,
      systemKeyListenerOptions
    );
    editorDocument.defaultView?.addEventListener(
      "keyup",
      onKeyUp,
      systemKeyListenerOptions
    );
    editorDocument.addEventListener("selectionchange", onSelectionChange);
    editorDocument.defaultView?.addEventListener("blur", onBlur);
    editorDocument.defaultView?.addEventListener("focus", onFocus);
  };
  const onEditorLoad = () => bindEditorDocument();

  popup = chromeDocument.createXULElement("panel");
  popup.id = "outlook-compose-mention-popup";
  popup.className = "outlook-compose-mention-popup";
  popup.setAttribute("type", "arrow");
  popup.setAttribute("role", "listbox");
  popup.setAttribute("aria-label", "Mention a contact");
  popup.setAttribute("noautofocus", "true");
  popup.setAttribute("consumeoutsideclicks", "false");
  list = chromeDocument.createElementNS(HTML_NS, "div");
  list.className = "outlook-compose-mention-list";
  popup.appendChild(list);
  popupSet.appendChild(popup);
  popup.addEventListener("popuphidden", onPopupHidden);
  editor.addEventListener("keydown", onKeyDown, true);
  editor.addEventListener("keyup", onKeyUp, true);
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("keyup", onKeyUp, true);
  editor.addEventListener("load", onEditorLoad, true);
  bindEditorDocument();

  return () => {
    hide();
    removeDocumentListeners();
    editor.removeEventListener("keydown", onKeyDown, true);
    editor.removeEventListener("keyup", onKeyUp, true);
    window.removeEventListener("keydown", onKeyDown, true);
    window.removeEventListener("keyup", onKeyUp, true);
    editor.removeEventListener("load", onEditorLoad, true);
    popup?.removeEventListener("popuphidden", onPopupHidden);
    popup?.remove();
    popup = null;
    list = null;
  };
}

/**
 * Apply the user's remembered calendar-dialog dimensions after Thunderbird's
 * own content sizing has finished.
 *
 * @param {Window} window - Calendar dialog window.
 * @param {Function[]} cleanups - Window cleanup callbacks.
 * @param {object} [options] - Geometry behavior.
 * @param {boolean} [options.rememberUserResize=true] - Whether later manual
 *   resizes should update the shared event-window preferences.
 * @returns {{schedule: Function}} Geometry scheduling controller.
 */
function installRememberedEventWindowGeometry(
  window,
  cleanups,
  { rememberUserResize = true } = {}
) {
  let geometryApplied = false;
  let applyingGeometry = false;
  let geometryTrackingReady = false;
  let userResizedGeometry = false;
  let geometryTimer = 0;
  let geometryReleaseTimer = 0;
  let geometrySettleTimer = 0;
  let resizeTimer = 0;

  const isNormalWindow = () =>
    typeof window.STATE_NORMAL !== "number" ||
    window.windowState === window.STATE_NORMAL;

  const saveWindowGeometry = () => {
    if (
      !rememberUserResize ||
      !geometryApplied ||
      applyingGeometry ||
      !geometryTrackingReady ||
      !userResizedGeometry ||
      !isNormalWindow()
    ) {
      return;
    }
    const width = Math.round(window.outerWidth);
    const height = Math.round(window.outerHeight);
    if (Number.isFinite(width) && width >= 400) {
      Services.prefs.setIntPref(EVENT_WINDOW_WIDTH_PREF, width);
    }
    if (Number.isFinite(height) && height >= 500) {
      Services.prefs.setIntPref(EVENT_WINDOW_HEIGHT_PREF, height);
    }
  };

  const onWindowResize = () => {
    if (
      !rememberUserResize ||
      !geometryApplied ||
      applyingGeometry ||
      !geometryTrackingReady
    ) {
      return;
    }
    userResizedGeometry = true;
    if (resizeTimer) {
      window.clearTimeout(resizeTimer);
    }
    resizeTimer = window.setTimeout(() => {
      resizeTimer = 0;
      saveWindowGeometry();
    }, 250);
  };

  const applyWindowGeometry = force => {
    geometryTimer = 0;
    if (
      (!force && geometryApplied) ||
      window.closed ||
      !isNormalWindow()
    ) {
      return;
    }

    /* DOM Screen values are CSS pixels and refer to the monitor containing
     * this window, including negative coordinates on left-hand displays. */
    const screen = window.screen;
    const availableWidth = Math.max(480, Number(screen?.availWidth) || 900);
    const availableHeight = Math.max(560, Number(screen?.availHeight) || 840);
    const availableLeft = Number(screen?.availLeft) || 0;
    const availableTop = Number(screen?.availTop) || 0;
    const maxWidth = Math.max(480, availableWidth - 48);
    const maxHeight = Math.max(560, availableHeight - 48);
    const hasStoredWidth = Services.prefs.prefHasUserValue(
      EVENT_WINDOW_WIDTH_PREF
    );
    const hasStoredHeight = Services.prefs.prefHasUserValue(
      EVENT_WINDOW_HEIGHT_PREF
    );
    const hasStoredGeometry = hasStoredWidth && hasStoredHeight;
    const storedWidth = hasStoredWidth
      ? Services.prefs.getIntPref(
          EVENT_WINDOW_WIDTH_PREF,
          DEFAULT_EVENT_WINDOW_WIDTH
        )
      : DEFAULT_EVENT_WINDOW_WIDTH;
    const storedHeight = hasStoredHeight
      ? Services.prefs.getIntPref(
          EVENT_WINDOW_HEIGHT_PREF,
          DEFAULT_EVENT_WINDOW_HEIGHT
        )
      : DEFAULT_EVENT_WINDOW_HEIGHT;
    const targetWidth = Math.min(
      maxWidth,
      Math.max(
        hasStoredGeometry ? 480 : Math.round(window.outerWidth),
        storedWidth
      )
    );
    const targetHeight = Math.min(
      maxHeight,
      Math.max(
        hasStoredGeometry ? 560 : Math.round(window.outerHeight),
        storedHeight
      )
    );

    let targetLeft;
    let targetTop;
    if (hasStoredGeometry) {
      targetLeft = Math.min(
        Math.max(window.screenX, availableLeft),
        availableLeft + availableWidth - targetWidth
      );
      targetTop = Math.min(
        Math.max(window.screenY, availableTop),
        availableTop + availableHeight - targetHeight
      );
    } else {
      targetLeft = availableLeft + Math.round((availableWidth - targetWidth) / 2);
      targetTop = availableTop + Math.round((availableHeight - targetHeight) / 2);
    }

    applyingGeometry = true;
    try {
      window.resizeTo(targetWidth, targetHeight);
      window.moveTo(targetLeft, targetTop);
      geometryApplied = true;
    } catch (error) {
      console.error(
        "Outlook Style Companion could not size the event window:",
        error
      );
    }
    if (geometryReleaseTimer) {
      window.clearTimeout(geometryReleaseTimer);
    }
    geometryReleaseTimer = window.setTimeout(() => {
      geometryReleaseTimer = 0;
      applyingGeometry = false;
    }, 0);
    /* Ignore the resize event generated by resizeTo() and any final native
     * content-fit adjustment. Only a later user resize should update the
     * remembered normal dimensions. This also preserves a larger preference
     * when the window is temporarily clamped to a smaller monitor. */
    if (rememberUserResize && !geometryTrackingReady) {
      if (geometrySettleTimer) {
        window.clearTimeout(geometrySettleTimer);
      }
      geometrySettleTimer = window.setTimeout(() => {
        geometrySettleTimer = 0;
        geometryTrackingReady = true;
      }, 750);
    }
  };

  const scheduleWindowGeometry = (delay = 100, force = false) => {
    if ((!force && geometryApplied) || window.closed) {
      return;
    }
    if (geometryTimer) {
      if (!force) {
        return;
      }
      window.clearTimeout(geometryTimer);
    }
    geometryTimer = window.setTimeout(
      () => applyWindowGeometry(force),
      delay
    );
  };

  cleanups.push(() => {
    try {
      saveWindowGeometry();
    } finally {
      if (rememberUserResize) {
        window.removeEventListener("resize", onWindowResize);
      }
      if (geometryTimer) {
        window.clearTimeout(geometryTimer);
      }
      if (geometryReleaseTimer) {
        window.clearTimeout(geometryReleaseTimer);
        geometryReleaseTimer = 0;
      }
      if (geometrySettleTimer) {
        window.clearTimeout(geometrySettleTimer);
      }
      if (resizeTimer) {
        window.clearTimeout(resizeTimer);
      }
    }
  });
  if (rememberUserResize) {
    window.addEventListener("resize", onWindowResize);
  }

  return { schedule: scheduleWindowGeometry };
}

function normalizeAttendeeId(id) {
  return cal.email.removeMailTo(id || "").trim().toLowerCase();
}

function getCurrentEventCalendar(frameWindow) {
  try {
    return (
      frameWindow.getCurrentCalendar?.() ||
      frameWindow.calendarItem?.calendar ||
      null
    );
  } catch (error) {
    return frameWindow.calendarItem?.calendar || null;
  }
}

function resolveInlineAttendeeMailboxes(value) {
  const addresses = MailServices.headerParser.makeFromDisplayAddress(value);
  const mailboxes = new Map();
  const visitedLists = new Set();

  const resolveAddress = address => {
    let list = null;
    if (address.name) {
      try {
        list = MailUtils.findListInAddressBooks(address.name);
      } catch (error) {
        /* A plain display name does not necessarily name an address-book list. */
      }
    }
    if (list) {
      const listKey = list.UID || list.URI || address.name.toLowerCase();
      if (visitedLists.has(listKey)) {
        return;
      }
      visitedLists.add(listKey);
      for (const childCard of Array.from(list.childCards || [])) {
        const card = childCard.QueryInterface(Ci.nsIAbCard);
        resolveAddress(
          MailServices.headerParser.makeMailboxObject(
            card.displayName,
            card.primaryEmail
          )
        );
      }
      return;
    }

    const email = address.email?.trim();
    if (!email) {
      return;
    }
    mailboxes.set(email.toLowerCase(), {
      email,
      name: address.name?.trim() || "",
    });
  };

  for (const address of addresses) {
    resolveAddress(address);
  }
  return [...mailboxes.values()];
}

function getCalendarOrganizerId(calendar) {
  try {
    return calendar?.getProperty("organizerId") || "";
  } catch (error) {
    return "";
  }
}

function getCalendarOrganizerName(calendar) {
  try {
    return calendar?.getProperty("organizerCN") || "";
  } catch (error) {
    return "";
  }
}

function addDefaultOrganizerAttendee(frameWindow, calendar, attendees) {
  if (attendees.length) {
    return;
  }
  if (frameWindow.organizer?.id) {
    attendees.push(frameWindow.organizer.clone());
    return;
  }
  const organizerId = getCalendarOrganizerId(calendar);
  if (!organizerId) {
    return;
  }
  const organizerAsAttendee = new CalAttendee();
  organizerAsAttendee.id = cal.email.removeMailTo(organizerId);
  organizerAsAttendee.commonName = getCalendarOrganizerName(calendar);
  organizerAsAttendee.role = "REQ-PARTICIPANT";
  organizerAsAttendee.participationStatus = "ACCEPTED";
  attendees.push(organizerAsAttendee);
}

function reconcileInlineOrganizer(frameWindow, calendar, attendees) {
  const existingOrganizer = frameWindow.organizer;
  const organizerId = existingOrganizer?.id || getCalendarOrganizerId(calendar);
  let organizer;

  if (organizerId) {
    const organizerKey = normalizeAttendeeId(organizerId);
    const nonOrganizerAttendees = attendees.filter(
      attendee => normalizeAttendeeId(attendee.id) !== organizerKey
    );
    if (nonOrganizerAttendees.length) {
      if (existingOrganizer) {
        organizer = existingOrganizer;
      } else {
        organizer = new CalAttendee();
        organizer.id = cal.email.removeMailTo(organizerId);
        organizer.commonName = getCalendarOrganizerName(calendar);
        organizer.isOrganizer = true;
      }
    } else {
      attendees.length = 0;
    }
  }

  return { attendees, organizer };
}

function applyInlineAttendeeState(frameWindow, attendees) {
  const calendar = getCurrentEventCalendar(frameWindow);
  const nextState = reconcileInlineOrganizer(
    frameWindow,
    calendar,
    attendees
  );
  frameWindow.attendees = nextState.attendees;
  frameWindow.organizer = nextState.organizer;
  frameWindow.updateAttendeeInterface();
}

function formatInlineAttendee(attendee) {
  const email = cal.email.removeMailTo(attendee.id || "");
  return {
    email,
    label: attendee.commonName || email,
  };
}

/**
 * Add a small Outlook-style Guests editor to Thunderbird's event iframe while
 * keeping Thunderbird's cloned attendee/organizer state and native Save/Send
 * path authoritative.
 *
 * @param {Element} itemFrame - The native calendar item iframe.
 * @returns {Function} Cleanup function.
 */
function enhanceInlineGuestsEditor(itemFrame) {
  const hostWindow = itemFrame.ownerDocument.defaultView;
  const installDeadline = hostWindow.performance.now() + 15000;
  let disposed = false;
  let retryTimer = 0;
  let attendeeObserver = null;
  let commandObserver = null;
  let frameWindow = null;
  let frameDocument = null;
  let row = null;
  let attendeeContainer = null;
  let chips = null;
  let input = null;
  let autocompletePopup = null;
  let openSchedulingButton = null;
  let feedback = null;
  let calendarPicker = null;

  const scheduleInstallRetry = install => {
    if (disposed || hostWindow.performance.now() >= installDeadline) {
      return;
    }
    retryTimer = hostWindow.setTimeout(install, 100);
  };

  const setFeedback = message => {
    if (feedback) {
      feedback.textContent = message;
    }
  };
  const attendeeEditingDisabled = () => {
    const command = frameDocument?.getElementById("cmd_attendees");
    return Boolean(
      frameWindow?.gIsReadOnly ||
        attendeeContainer?.disabled ||
        attendeeContainer?.hasAttribute("disabled") ||
        command?.disabled ||
        command?.hasAttribute("disabled")
    );
  };

  const renderChips = () => {
    if (!chips || !frameWindow || disposed) {
      return;
    }
    chips.replaceChildren();
    const attendees = Array.isArray(frameWindow.attendees)
      ? frameWindow.attendees
      : [];
    const organizerKey = normalizeAttendeeId(
      frameWindow.organizer?.id ||
        getCalendarOrganizerId(getCurrentEventCalendar(frameWindow))
    );
    const editingDisabled = attendeeEditingDisabled();
    if (input) {
      input.disabled = editingDisabled;
    }
    if (openSchedulingButton) {
      openSchedulingButton.disabled = editingDisabled;
    }
    attendees.forEach((attendee, index) => {
      if (
        !attendee?.id ||
        (organizerKey && normalizeAttendeeId(attendee.id) === organizerKey)
      ) {
        return;
      }
      const display = formatInlineAttendee(attendee);
      const chip = frameDocument.createElement("span");
      chip.className = "outlook-style-guest-chip";
      chip.setAttribute("role", "listitem");
      chip.title = display.email;

      const label = frameDocument.createElement("span");
      label.className = "outlook-style-guest-chip-label";
      label.textContent = display.label;
      chip.appendChild(label);

      const remove = frameDocument.createElement("button");
      remove.type = "button";
      remove.className = "outlook-style-guest-chip-remove";
      remove.textContent = "×";
      remove.setAttribute("aria-label", `Remove ${display.label}`);
      remove.disabled = editingDisabled;
      remove.addEventListener("click", () => {
        if (attendeeEditingDisabled()) {
          return;
        }
        const nextAttendees = [...frameWindow.attendees];
        nextAttendees.splice(index, 1);
        applyInlineAttendeeState(frameWindow, nextAttendees);
        renderChips();
        setFeedback(`${display.label} removed.`);
        input?.focus();
      });
      chip.appendChild(remove);
      chips.appendChild(chip);
    });
  };

  const commitInput = () => {
    const value = input?.value.trim();
    if (!value || !frameWindow || disposed || attendeeEditingDisabled()) {
      return;
    }

    let mailboxes;
    try {
      mailboxes = resolveInlineAttendeeMailboxes(value);
    } catch (error) {
      setFeedback("Enter a valid attendee name or email address.");
      return;
    }
    if (!mailboxes.length) {
      setFeedback("Enter a valid attendee name or email address.");
      return;
    }

    const calendar = getCurrentEventCalendar(frameWindow);
    const nextAttendees = Array.isArray(frameWindow.attendees)
      ? [...frameWindow.attendees]
      : [];
    addDefaultOrganizerAttendee(frameWindow, calendar, nextAttendees);
    const existingIds = new Set(
      nextAttendees.map(attendee => normalizeAttendeeId(attendee.id))
    );
    let added = 0;
    for (const mailbox of mailboxes) {
      const attendeeId = cal.email.prependMailTo(mailbox.email);
      const attendeeKey = normalizeAttendeeId(attendeeId);
      if (!attendeeKey || existingIds.has(attendeeKey)) {
        continue;
      }
      const attendee = new CalAttendee();
      attendee.id = attendeeId;
      attendee.role = "REQ-PARTICIPANT";
      attendee.userType = "INDIVIDUAL";
      if (mailbox.name && mailbox.name !== mailbox.email) {
        attendee.commonName = mailbox.name;
      }
      nextAttendees.push(attendee);
      existingIds.add(attendeeKey);
      added++;
    }

    input.value = "";
    applyInlineAttendeeState(frameWindow, nextAttendees);
    renderChips();
    setFeedback(
      added
        ? `${added} attendee${added === 1 ? "" : "s"} added.`
        : "That attendee is already added."
    );
  };

  const onInputChange = () => commitInput();
  const onInputKeyDown = event => {
    if (event.isComposing || attendeeEditingDisabled()) {
      return;
    }
    if (event.key === "Enter") {
      if (input.popupOpen) {
        frameWindow.setTimeout(commitInput, 0);
      } else {
        event.preventDefault();
        commitInput();
      }
    } else if (event.key === ";") {
      event.preventDefault();
      commitInput();
    } else if (
      event.key === "Backspace" &&
      !input.value &&
      frameWindow.attendees?.length
    ) {
      const nextAttendees = [...frameWindow.attendees];
      const organizerKey = normalizeAttendeeId(
        frameWindow.organizer?.id ||
          getCalendarOrganizerId(getCurrentEventCalendar(frameWindow))
      );
      const removeIndex = nextAttendees.findLastIndex(
        attendee =>
          attendee?.id &&
          (!organizerKey || normalizeAttendeeId(attendee.id) !== organizerKey)
      );
      if (removeIndex < 0) {
        return;
      }
      event.preventDefault();
      const [removedAttendee] = nextAttendees.splice(removeIndex, 1);
      const removed = formatInlineAttendee(removedAttendee);
      applyInlineAttendeeState(frameWindow, nextAttendees);
      renderChips();
      setFeedback(`${removed.label} removed.`);
    } else if (event.key === "Escape" && input.value) {
      event.preventDefault();
      input.value = "";
      setFeedback("Attendee entry cleared.");
    }
  };
  const onOpenScheduling = () => {
    if (attendeeEditingDisabled()) {
      return;
    }
    try {
      frameWindow.editAttendees();
    } finally {
      renderChips();
    }
  };
  const onCalendarChange = () => renderChips();

  const install = () => {
    retryTimer = 0;
    if (disposed) {
      return;
    }
    frameWindow = itemFrame.contentWindow;
    frameDocument = itemFrame.contentDocument;
    if (
      frameDocument?.documentURI !==
        "chrome://calendar/content/calendar-item-iframe.xhtml" ||
      !frameWindow?.calendarItem ||
      !Array.isArray(frameWindow.attendees) ||
      typeof frameWindow.updateAttendeeInterface !== "function" ||
      typeof frameWindow.editAttendees !== "function"
    ) {
      scheduleInstallRetry(install);
      return;
    }
    if (!frameWindow.calendarItem.isEvent()) {
      return;
    }

    const titleRow = frameDocument.getElementById("event-grid-title-row");
    if (!titleRow) {
      scheduleInstallRetry(install);
      return;
    }
    frameDocument
      .getElementById("outlook-style-guests-row")
      ?.remove();
    frameDocument
      .getElementById("outlook-style-guests-autocomplete-popup")
      ?.remove();
    frameDocument.documentElement.setAttribute(
      "data-outlook-inline-guests",
      "true"
    );

    row = frameDocument.createElementNS(HTML_NS, "tr");
    row.id = "outlook-style-guests-row";

    const heading = frameDocument.createElementNS(HTML_NS, "th");
    const guestsLabel = frameDocument.createElementNS(HTML_NS, "label");
    guestsLabel.id = "outlook-style-guests-label";
    guestsLabel.htmlFor = "outlook-style-guests-input";
    guestsLabel.textContent = "Guests:";
    heading.appendChild(guestsLabel);
    row.appendChild(heading);

    const cell = frameDocument.createElementNS(HTML_NS, "td");
    cell.className = "event-input-td";
    const stack = frameDocument.createElementNS(HTML_NS, "div");
    stack.className = "outlook-style-guests-stack";
    const field = frameDocument.createElementNS(HTML_NS, "div");
    field.className = "outlook-style-guests-field";
    field.setAttribute("role", "group");
    field.setAttribute("aria-labelledby", guestsLabel.id);

    chips = frameDocument.createElementNS(HTML_NS, "div");
    chips.className = "outlook-style-guest-chips";
    chips.setAttribute("role", "list");
    field.appendChild(chips);

    input = frameDocument.createElement("input", {
      is: "autocomplete-input",
    });
    input.id = "outlook-style-guests-input";
    input.className = "outlook-style-guests-input";
    input.placeholder = "Add required attendees";
    input.setAttribute("aria-labelledby", guestsLabel.id);
    input.setAttribute("aria-describedby", "outlook-style-guests-help");
    input.setAttribute("autocompletesearch", "addrbook ldap");
    input.setAttribute("autocompletesearchparam", "{}");
    input.setAttribute("forcecomplete", "true");
    input.setAttribute("timeout", "200");
    input.setAttribute("completedefaultindex", "true");
    input.setAttribute("completeselectedindex", "true");
    input.setAttribute("minresultsforpopup", "1");
    const popupSet = frameDocument.getElementById("event-dialog-popupset");
    if (popupSet) {
      autocompletePopup = frameDocument.createXULElement("panel", {
        is: "autocomplete-richlistbox-popup",
      });
      autocompletePopup.id = "outlook-style-guests-autocomplete-popup";
      autocompletePopup.setAttribute("type", "autocomplete-richlistbox");
      autocompletePopup.setAttribute("role", "group");
      autocompletePopup.setAttribute("noautofocus", "noautofocus");
      popupSet.appendChild(autocompletePopup);
      input.setAttribute("autocompletepopup", autocompletePopup.id);
    }
    input.addEventListener("change", onInputChange);
    input.addEventListener("keydown", onInputKeyDown);
    field.appendChild(input);

    openSchedulingButton = frameDocument.createElementNS(
      HTML_NS,
      "button"
    );
    openSchedulingButton.id = "outlook-style-open-scheduling";
    openSchedulingButton.type = "button";
    openSchedulingButton.textContent = "Scheduling assistant";
    openSchedulingButton.title =
      "Add optional attendees, rooms, resources, and check provider availability";
    openSchedulingButton.addEventListener("click", onOpenScheduling);
    field.appendChild(openSchedulingButton);
    stack.appendChild(field);

    const help = frameDocument.createElementNS(HTML_NS, "div");
    help.id = "outlook-style-guests-help";
    help.className = "outlook-style-guests-help";
    help.textContent =
      "Optional attendees, rooms, resources, roles, and availability.";
    stack.appendChild(help);

    feedback = frameDocument.createElementNS(HTML_NS, "div");
    feedback.id = "outlook-style-guests-feedback";
    feedback.className = "outlook-style-guests-feedback";
    feedback.setAttribute("role", "status");
    feedback.setAttribute("aria-live", "polite");
    stack.appendChild(feedback);
    cell.appendChild(stack);
    row.appendChild(cell);
    titleRow.after(row);

    attendeeContainer = frameDocument.querySelector(
      ".item-attendees-list-container"
    );
    if (attendeeContainer) {
      attendeeObserver = new frameWindow.MutationObserver(renderChips);
      attendeeObserver.observe(attendeeContainer, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["disabled"],
      });
    }
    calendarPicker = frameDocument.getElementById("item-calendar");
    calendarPicker?.addEventListener("command", onCalendarChange);
    const attendeesCommand = frameDocument.getElementById("cmd_attendees");
    if (attendeesCommand) {
      commandObserver = new frameWindow.MutationObserver(renderChips);
      commandObserver.observe(attendeesCommand, {
        attributes: true,
        attributeFilter: ["disabled"],
      });
    }
    renderChips();
  };

  const cleanup = () => {
    disposed = true;
    if (retryTimer) {
      hostWindow.clearTimeout(retryTimer);
      retryTimer = 0;
    }
    attendeeObserver?.disconnect();
    attendeeObserver = null;
    commandObserver?.disconnect();
    commandObserver = null;
    input?.removeEventListener("change", onInputChange);
    input?.removeEventListener("keydown", onInputKeyDown);
    try {
      autocompletePopup?.hidePopup?.();
    } catch (error) {
      /* A closing event window may already have disconnected the popup. */
    }
    try {
      input?.detachController?.();
    } catch (error) {
      /* The autocomplete controller may already be detached. */
    }
    try {
      if (autocompletePopup?._appendResultTimeout != null) {
        frameWindow?.clearTimeout(autocompletePopup._appendResultTimeout);
        autocompletePopup._appendResultTimeout = null;
      }
      if (autocompletePopup?._adjustHeightRAFToken != null) {
        frameWindow?.cancelAnimationFrame(
          autocompletePopup._adjustHeightRAFToken
        );
        autocompletePopup._adjustHeightRAFToken = null;
      }
    } catch (error) {
      /* Private autocomplete work can disappear during window teardown. */
    }
    try {
      if (autocompletePopup) {
        autocompletePopup.mInput = null;
      }
    } catch (error) {
      /* A disconnected popup no longer needs its input reference cleared. */
    }
    autocompletePopup?.remove();
    openSchedulingButton?.removeEventListener("click", onOpenScheduling);
    calendarPicker?.removeEventListener("command", onCalendarChange);
    row?.remove();
    frameDocument?.documentElement?.removeAttribute(
      "data-outlook-inline-guests"
    );
    row = null;
    attendeeContainer = null;
    chips = null;
    input = null;
    autocompletePopup = null;
    openSchedulingButton = null;
    feedback = null;
    calendarPicker = null;
    frameWindow = null;
    frameDocument = null;
  };
  try {
    install();
  } catch (error) {
    cleanup();
    throw error;
  }
  return cleanup;
}

function getChildDialogThemeSource(window) {
  try {
    if (window.opener?.top && !window.opener.top.closed) {
      return window.opener.top;
    }
    if (window.opener && !window.opener.closed) {
      return window.opener;
    }
  } catch (error) {
    /* Fall back to the child dialog if its opener is being destroyed. */
  }
  return window;
}

function installChildDialogThemeBridge(window, cleanups) {
  const sourceWindow = getChildDialogThemeSource(window);
  const sourceRoot = sourceWindow.document?.documentElement;
  const getSourceScheme = () => {
    const nativeScheme = sourceRoot?.getAttribute("lwt-sidebar");
    if (nativeScheme === "light" || nativeScheme === "dark") {
      return nativeScheme;
    }
    if (sourceWindow !== window) {
      const stampedScheme = sourceRoot?.getAttribute(
        OUTLOOK_COLOR_SCHEME_ATTRIBUTE
      );
      if (stampedScheme === "light" || stampedScheme === "dark") {
        return stampedScheme;
      }
    }
    return sourceWindow.matchMedia("(prefers-color-scheme: dark)").matches
      ? "dark"
      : "light";
  };
  const syncScheme = () => {
    if (!window.closed) {
      stampOutlookColorScheme(window.document, getSourceScheme());
    }
  };
  const sourceObserver = sourceRoot
    ? new sourceWindow.MutationObserver(syncScheme)
    : null;
  const sourceSystemScheme = sourceWindow.matchMedia(
    "(prefers-color-scheme: dark)"
  );
  cleanups.push(() => {
    sourceObserver?.disconnect();
    sourceWindow.removeEventListener("windowlwthemeupdate", syncScheme);
    sourceSystemScheme.removeEventListener("change", syncScheme);
    window.document.documentElement?.removeAttribute(
      OUTLOOK_COLOR_SCHEME_ATTRIBUTE
    );
  });
  sourceObserver?.observe(sourceRoot, {
    attributes: true,
    attributeFilter: [
      OUTLOOK_COLOR_SCHEME_ATTRIBUTE,
      "lwt-sidebar",
      "class",
      "style",
    ],
  });
  sourceWindow.addEventListener("windowlwthemeupdate", syncScheme);
  sourceSystemScheme.addEventListener("change", syncScheme);
  syncScheme();
}

function installAttendeeDialogGeometry(window, cleanups) {
  let geometryTimer = 0;
  cleanups.push(() => {
    if (geometryTimer) {
      window.clearTimeout(geometryTimer);
      geometryTimer = 0;
    }
  });
  geometryTimer = window.setTimeout(() => {
    geometryTimer = 0;
    if (
      window.closed ||
      (typeof window.STATE_NORMAL === "number" &&
        window.windowState !== window.STATE_NORMAL)
    ) {
      return;
    }
    let opener = null;
    let screen = window.screen;
    try {
      opener = getChildDialogThemeSource(window);
      if (opener !== window && !opener.closed) {
        screen = opener.screen || screen;
      } else {
        opener = null;
      }
    } catch (error) {
      /* The child window's screen remains a safe fallback. */
      opener = null;
    }
    const availableLeft = Number(screen?.availLeft) || 0;
    const availableTop = Number(screen?.availTop) || 0;
    const availableWidth = Math.max(640, Number(screen?.availWidth) || 1100);
    const availableHeight = Math.max(560, Number(screen?.availHeight) || 720);
    const targetWidth = Math.min(
      availableWidth - 48,
      Math.max(Number(window.outerWidth) || 0, DEFAULT_ATTENDEES_WINDOW_WIDTH)
    );
    const targetHeight = Math.min(
      availableHeight - 48,
      Math.max(
        Number(window.outerHeight) || 0,
        DEFAULT_ATTENDEES_WINDOW_HEIGHT
      )
    );

    let centerX = availableLeft + availableWidth / 2;
    let centerY = availableTop + availableHeight / 2;
    try {
      if (opener) {
        centerX = opener.screenX + opener.outerWidth / 2;
        centerY = opener.screenY + opener.outerHeight / 2;
      }
    } catch (error) {
      /* The active screen center remains a safe fallback. */
    }
    const targetLeft = Math.min(
      Math.max(Math.round(centerX - targetWidth / 2), availableLeft),
      availableLeft + availableWidth - targetWidth
    );
    const targetTop = Math.min(
      Math.max(Math.round(centerY - targetHeight / 2), availableTop),
      availableTop + availableHeight - targetHeight
    );

    try {
      if (
        Math.round(window.outerWidth) !== Math.round(targetWidth) ||
        Math.round(window.outerHeight) !== Math.round(targetHeight)
      ) {
        window.resizeTo(targetWidth, targetHeight);
      }
      window.moveTo(targetLeft, targetTop);
    } catch (error) {
      console.error(
        "Outlook Style Companion could not size the attendee window:",
        error
      );
    }
  }, 100);

}

function enhanceAttendeeDialogSurface(window, cleanups) {
  const document = window.document;
  let notice = null;
  let findTimeButton = null;
  let findTimeStatus = null;
  let findTimeGeneration = 0;
  let allDayToggle = null;
  const zoomButtonSnapshots = new Map();
  cleanups.push(() => {
    findTimeGeneration++;
    allDayToggle?.removeEventListener("command", syncFindTimeButton);
    findTimeButton?.removeEventListener("click", onFindTime);
    findTimeButton?.remove();
    findTimeStatus?.remove();
    notice?.remove();
    for (const [button, snapshot] of zoomButtonSnapshots) {
      for (const [name, value] of snapshot) {
        if (value === null) {
          button.removeAttribute(name);
        } else {
          button.setAttribute(name, value);
        }
      }
    }
    zoomButtonSnapshots.clear();
    document.documentElement.removeAttribute(
      "data-outlook-attendee-dialog"
    );
  });
  document.documentElement.setAttribute(
    "data-outlook-attendee-dialog",
    "true"
  );
  installChildDialogThemeBridge(window, cleanups);
  installAttendeeDialogGeometry(window, cleanups);

  const labelZoomButton = (id, label, description) => {
    const button = document.getElementById(id);
    if (!button) {
      return;
    }
    const attributes = [
      "aria-label",
      "data-outlook-explicit-zoom",
      "label",
      "tooltiptext",
    ];
    zoomButtonSnapshots.set(
      button,
      new Map(attributes.map(name => [name, button.getAttribute(name)]))
    );
    button.setAttribute("data-outlook-explicit-zoom", "true");
    button.setAttribute("label", label);
    button.setAttribute("aria-label", description);
    button.setAttribute("tooltiptext", description);
  };
  labelZoomButton("zoom-out-button", "\u2212", "Zoom out");
  labelZoomButton("zoom-in-button", "+", "Zoom in");

  document.getElementById("outlook-style-availability-notice")?.remove();
  notice = document.createElementNS(HTML_NS, "div");
  notice.id = "outlook-style-availability-notice";
  notice.setAttribute("role", "note");
  notice.textContent =
    "Availability comes from your calendar provider. A blank attendee row can mean the person is free, or that no availability data was returned. Google calendars connected through CalDAV do not provide participant free/busy lookup.";
  const outer = document.getElementById("outer");
  if (outer?.parentNode) {
    outer.parentNode.insertBefore(notice, outer);
  } else {
    (document.querySelector("dialog") || document.body).prepend(notice);
  }

  const topBar = document.querySelector("dialog > hbox:first-child");
  const topBarSpacer = topBar?.querySelector(":scope > spacer");
  if (!topBar || !topBarSpacer) {
    return;
  }

  findTimeButton = document.createElementNS(HTML_NS, "button");
  findTimeButton.id = "outlook-style-find-time";
  findTimeButton.type = "button";
  findTimeButton.textContent = "Find time";
  findTimeButton.title =
    "Select the first work-time slot when everyone is free or tentative";

  findTimeStatus = document.createElementNS(HTML_NS, "span");
  findTimeStatus.id = "outlook-style-find-time-status";
  findTimeStatus.setAttribute("role", "status");
  findTimeStatus.setAttribute("aria-live", "polite");
  topBar.insertBefore(findTimeButton, topBarSpacer);
  topBar.insertBefore(findTimeStatus, topBarSpacer);

  const getAttendeeIds = () => {
    const ids = new Map();
    for (const row of document.querySelectorAll(
      "#attendee-list event-attendee"
    )) {
      let id = "";
      try {
        id = row.attendee?.id || "";
      } catch (error) {
        /* A row can be rebuilding while autocomplete commits its value. */
      }
      const email = cal.email.removeMailTo(id).trim();
      if (email) {
        ids.set(email.toLowerCase(), cal.email.prependMailTo(email));
      }
    }
    return [...ids.values()];
  };

  const getBusyIntervals = (attendeeId, from, to, generation) =>
    new Promise(resolve => {
      const intervals = [];
      let finished = false;
      let timer = 0;
      const finish = complete => {
        if (finished) {
          return;
        }
        finished = true;
        if (timer) {
          window.clearTimeout(timer);
        }
        resolve({ complete, intervals });
      };
      timer = window.setTimeout(() => finish(false), 15000);
      try {
        cal.freeBusyService.getFreeBusyIntervals(
          attendeeId,
          from,
          to,
          Ci.calIFreeBusyInterval.BUSY_ALL,
          {
            onResult(operation, results) {
              if (generation !== findTimeGeneration) {
                finish(false);
                return;
              }
              for (const result of results || []) {
                const type = Number(result.freeBusyType);
                if (
                  type !== Ci.calIFreeBusyInterval.FREE &&
                  type !== Ci.calIFreeBusyInterval.BUSY_TENTATIVE
                ) {
                  intervals.push(result.interval);
                }
              }
              if (!operation.isPending) {
                finish(true);
              }
            },
          }
        );
      } catch (error) {
        finish(false);
      }
    });

  const roundUpToQuarterHour = dateTime => {
    const rounded = dateTime.clone();
    rounded.isDate = false;
    const extraMinutes = (15 - (rounded.minute % 15)) % 15;
    if (rounded.second || extraMinutes) {
      rounded.minute += extraMinutes || 15;
    }
    rounded.second = 0;
    return rounded;
  };

  const dayOffPreferences = [
    "calendar.week.d0sundaysoff",
    "calendar.week.d1mondaysoff",
    "calendar.week.d2tuesdaysoff",
    "calendar.week.d3wednesdaysoff",
    "calendar.week.d4thursdaysoff",
    "calendar.week.d5fridaysoff",
    "calendar.week.d6saturdaysoff",
  ];
  const isWorkingDay = dateTime =>
    !Services.prefs.getBoolPref(
      dayOffPreferences[dateTime.weekday],
      false
    );
  const overlaps = (candidateStart, candidateEnd, intervals) =>
    intervals.some(
      interval =>
        interval?.start?.compare(candidateEnd) < 0 &&
        interval?.end?.compare(candidateStart) > 0
    );

  const applyFoundTime = (candidateStart, duration) => {
    const picker = window.dateTimePickerUI;
    const originalZone = picker.startValue.timezone;
    const nextStart = candidateStart.getInTimezone(originalZone);
    const nextEnd = nextStart.clone();
    nextEnd.addDuration(duration);
    window.updateByFunction = true;
    try {
      picker.startValue = nextStart;
      picker.endValue = nextEnd;
      picker.saveOldValues();
    } finally {
      window.updateByFunction = false;
    }
    window.updateChange();
    window.updateRange();
    window.eventBar.update(true);
    try {
      const left = window.getOffsetLeft(
        nextStart.getInTimezone(cal.dtz.defaultTimezone)
      );
      const grid = document.getElementById("freebusy-grid");
      grid.scrollLeft = Math.max(0, left - grid.clientWidth / 3);
      document.getElementById("day-header-outer").scrollLeft =
        grid.scrollLeft;
    } catch (error) {
      /* The native range update already makes the selected day visible. */
    }
  };

  async function onFindTime() {
    const picker = window.dateTimePickerUI;
    if (
      !picker?.startValue ||
      !picker?.endValue ||
      picker.allDay.checked ||
      window.readOnly
    ) {
      findTimeStatus.textContent =
        "Find time is available for editable, timed events.";
      return;
    }
    const attendeeIds = getAttendeeIds();
    if (!attendeeIds.length) {
      findTimeStatus.textContent = "Add at least one attendee first.";
      return;
    }

    const generation = ++findTimeGeneration;
    findTimeButton.disabled = true;
    findTimeButton.setAttribute("aria-busy", "true");
    findTimeStatus.textContent = "Checking availability…";
    try {
      const originalStart = picker.startValue;
      const duration = picker.endValue.subtractDate(originalStart);
      if (duration.inSeconds <= 0) {
        findTimeStatus.textContent = "Choose a valid event duration first.";
        return;
      }
      const workStartHour = Services.prefs.getIntPref(
        "calendar.view.daystarthour",
        8
      );
      const workEndHour = Services.prefs.getIntPref(
        "calendar.view.dayendhour",
        17
      );
      if (
        workStartHour < 0 ||
        workEndHour > 24 ||
        workEndHour <= workStartHour ||
        duration.inSeconds > (workEndHour - workStartHour) * 60 * 60
      ) {
        findTimeStatus.textContent =
          "The event does not fit inside the configured work day.";
        return;
      }

      let searchStart = originalStart.getInTimezone(
        cal.dtz.defaultTimezone
      );
      const now = cal.dtz.now().getInTimezone(cal.dtz.defaultTimezone);
      if (searchStart.compare(now) < 0) {
        searchStart = now;
      }
      searchStart = roundUpToQuarterHour(searchStart);

      const queryStart = searchStart.clone();
      queryStart.hour = workStartHour;
      queryStart.minute = 0;
      queryStart.second = 0;
      const queryEnd = queryStart.clone();
      queryEnd.day += 30;
      const attendeeResults = await Promise.all(
        attendeeIds.map(id =>
          getBusyIntervals(id, queryStart, queryEnd, generation)
        )
      );
      if (generation !== findTimeGeneration) {
        return;
      }
      if (attendeeResults.some(result => !result.complete)) {
        findTimeStatus.textContent =
          "Availability could not be confirmed for every attendee.";
        return;
      }
      const blockedIntervals = attendeeResults.flatMap(
        result => result.intervals
      );

      let found = null;
      const searchDate = queryStart.clone();
      for (let dayOffset = 0; dayOffset < 30 && !found; dayOffset++) {
        if (dayOffset) {
          searchDate.day++;
        }
        if (!isWorkingDay(searchDate)) {
          continue;
        }
        const workStart = searchDate.clone();
        workStart.hour = workStartHour;
        workStart.minute = 0;
        workStart.second = 0;
        const workEnd = searchDate.clone();
        workEnd.hour = workEndHour;
        workEnd.minute = 0;
        workEnd.second = 0;
        let candidate =
          dayOffset === 0 && searchStart.compare(workStart) > 0
            ? searchStart.clone()
            : workStart;
        if (candidate.compare(workEnd) >= 0) {
          continue;
        }
        while (candidate.compare(workEnd) < 0) {
          const candidateEnd = candidate.clone();
          candidateEnd.addDuration(duration);
          if (candidateEnd.compare(workEnd) > 0) {
            break;
          }
          if (!overlaps(candidate, candidateEnd, blockedIntervals)) {
            found = candidate;
            break;
          }
          candidate = candidate.clone();
          candidate.minute += 15;
        }
      }

      if (!found) {
        findTimeStatus.textContent =
          "No common work-time slot was found in the next 30 days.";
        return;
      }
      applyFoundTime(found, duration);
      findTimeStatus.textContent = `Selected ${cal.dtz.formatter.formatDateShort(
        found
      )} at ${cal.dtz.formatter.formatTime(found)}.`;
    } catch (error) {
      console.error(
        "Outlook Style Companion could not find a meeting time:",
        error
      );
      findTimeStatus.textContent = "A meeting time could not be selected.";
    } finally {
      if (generation === findTimeGeneration) {
        findTimeButton.removeAttribute("aria-busy");
        syncFindTimeButton();
      }
    }
  }

  function syncFindTimeButton() {
    findTimeButton.disabled = Boolean(
      window.readOnly ||
        window.dateTimePickerUI?.allDay?.checked ||
        findTimeButton.hasAttribute("aria-busy")
    );
  }

  allDayToggle = document.getElementById("all-day");
  allDayToggle?.addEventListener("command", syncFindTimeButton);
  findTimeButton.addEventListener("click", onFindTime);
  syncFindTimeButton();

}

function enhanceSpellCheckDialogSurface(window, cleanups) {
  const document = window.document;
  const root = document.documentElement;
  const suggestedList = document.getElementById("SuggestedList");
  const dictionaryList = document.getElementById("dictionary-list");
  let disposed = false;
  let firstFrame = 0;
  let secondFrame = 0;
  let fallbackTimer = 0;
  let settleTimer = 0;
  cleanups.push(() => {
    root.removeAttribute("data-outlook-style-spell-dialog");
    root.style.removeProperty("--outlook-spell-dialog-min-width");
    root.style.removeProperty("--outlook-spell-dialog-min-height");
  });

  let targetScreen = window.screen;
  try {
    if (window.opener && !window.opener.closed) {
      targetScreen = window.opener.screen || targetScreen;
    }
  } catch (error) {
    /* Use the spell window's current screen if its opener is closing. */
  }
  const availableWidth = Math.max(
    1,
    Number(targetScreen?.availWidth) || 520
  );
  const availableHeight = Math.max(
    1,
    Number(targetScreen?.availHeight) || 360
  );
  root.style.setProperty(
    "--outlook-spell-dialog-min-width",
    `${Math.min(520, Math.max(1, availableWidth - 48))}px`
  );
  root.style.setProperty(
    "--outlook-spell-dialog-min-height",
    `${Math.min(360, Math.max(1, availableHeight - 48))}px`
  );
  root.setAttribute("data-outlook-style-spell-dialog", "true");
  installChildDialogThemeBridge(window, cleanups);

  const centerAfterPaint = () => {
    if (disposed || window.closed || firstFrame || secondFrame) {
      return;
    }
    firstFrame = window.requestAnimationFrame(() => {
      firstFrame = 0;
      secondFrame = window.requestAnimationFrame(() => {
        secondFrame = 0;
        centerWindowOverOpener(window);
      });
    });
  };
  let readyObserver = null;
  const onInitialResize = () => centerAfterPaint();
  cleanups.push(() => {
    disposed = true;
    readyObserver?.disconnect();
    window.removeEventListener("resize", onInitialResize);
    if (firstFrame) {
      window.cancelAnimationFrame(firstFrame);
    }
    if (secondFrame) {
      window.cancelAnimationFrame(secondFrame);
    }
    if (fallbackTimer) {
      window.clearTimeout(fallbackTimer);
    }
    if (settleTimer) {
      window.clearTimeout(settleTimer);
    }
  });

  /* Put the dialog somewhere visible immediately. Thunderbird finishes its
   * dictionary query asynchronously and calls sizeToContent() afterward, so
   * repeat the placement after that final native sizing pass. */
  centerWindowOverOpener(window);
  centerAfterPaint();

  readyObserver = new window.MutationObserver(() => {
    if (
      suggestedList?.childElementCount ||
      dictionaryList?.childElementCount
    ) {
      readyObserver.disconnect();
      centerAfterPaint();
    }
  });
  if (suggestedList?.childElementCount || dictionaryList?.childElementCount) {
    centerAfterPaint();
  } else {
    if (suggestedList) {
      readyObserver.observe(suggestedList, { childList: true, subtree: true });
    }
    if (dictionaryList) {
      readyObserver.observe(dictionaryList, { childList: true, subtree: true });
    }
  }

  window.addEventListener("resize", onInitialResize);
  fallbackTimer = window.setTimeout(() => {
    fallbackTimer = 0;
    centerAfterPaint();
  }, 500);
  settleTimer = window.setTimeout(() => {
    settleTimer = 0;
    window.removeEventListener("resize", onInitialResize);
  }, 1000);

}

function enhanceEventEditorSurface(window, cleanups) {
  const document = window.document;
  let itemFrame = null;
  let removeEditorTracking = null;
  let removeInlineGuests = null;
  let unifiedViewer = null;
  let viewerFrame = null;
  let viewerFrameWasHidden = false;
  let viewerToolbox = null;
  let viewerToolboxWasHidden = false;
  let viewerTimer = 0;
  let nativeEditorRequested = false;
  const geometry = installRememberedEventWindowGeometry(window, cleanups);

  const scheduleWindowGeometry = () => {
    /* Thunderbird performs its own iframe content sizing from the same load
     * event. Run afterward so the native compact minimum cannot overwrite the
     * Outlook-sized window. */
    geometry.schedule(100);
  };

  const bindDescriptionEditor = () => {
    removeEditorTracking?.();
    removeEditorTracking = null;
    const editor = itemFrame?.contentDocument?.getElementById(
      "item-description"
    );
    if (editor) {
      removeEditorTracking = trackEditorSurface(editor, false);
    }
  };
  const bindInlineGuests = () => {
    removeInlineGuests?.();
    removeInlineGuests = itemFrame
      ? enhanceInlineGuestsEditor(itemFrame)
      : null;
  };
  const revealNativeEditor = () => {
    if (viewerTimer) {
      window.clearTimeout(viewerTimer);
      viewerTimer = 0;
    }
    unifiedViewer?.remove();
    unifiedViewer = null;
    if (viewerFrame) {
      viewerFrame.toggleAttribute("hidden", viewerFrameWasHidden);
    }
    if (viewerToolbox) {
      viewerToolbox.toggleAttribute("hidden", viewerToolboxWasHidden);
    }
    viewerFrame = null;
    viewerToolbox = null;
    scheduleWindowGeometry();
  };
  const showUnifiedViewer = () => {
    viewerTimer = 0;
    if (
      unifiedViewer ||
      nativeEditorRequested ||
      !itemFrame?.isConnected ||
      window.closed
    ) {
      return;
    }
    const frameWindow = itemFrame.contentWindow;
    const item = frameWindow?.calendarItem;
    /* New Calendar/Tasks items must open directly in their native editable
     * form. Existing items use the common read-only card until Edit is chosen. */
    if (!item || frameWindow.mode !== "modify" || !item.id) {
      return;
    }
    const requestKey = getNativeEditorRequestKey(item);
    if (nativeEditorRequestKeys.delete(requestKey)) {
      nativeEditorRequested = true;
      return;
    }
    viewerFrame = itemFrame;
    viewerFrameWasHidden = itemFrame.hasAttribute("hidden");
    viewerToolbox = document.getElementById("event-toolbox");
    viewerToolboxWasHidden = viewerToolbox?.hasAttribute("hidden") || false;
    unifiedViewer = createUnifiedItemView(document, item, {
      onEdit: revealNativeEditor,
    });
    unifiedViewer.id = "outlook-style-editor-item-view";
    unifiedViewer.style.flex = "1 1 auto";
    unifiedViewer.style.maxBlockSize = "none";
    unifiedViewer.style.overflow = "auto";
    itemFrame.before(unifiedViewer);
    itemFrame.setAttribute("hidden", "hidden");
    viewerToolbox?.setAttribute("hidden", "hidden");
    scheduleWindowGeometry();
  };
  const scheduleUnifiedViewer = () => {
    if (viewerTimer || unifiedViewer || !itemFrame?.isConnected) {
      return;
    }
    viewerTimer = window.setTimeout(showUnifiedViewer, 100);
  };
  const onItemFrameLoad = event => {
    if (
      event.target === itemFrame ||
      event.target === itemFrame?.contentDocument ||
      event.originalTarget === itemFrame?.contentDocument
    ) {
      bindDescriptionEditor();
      bindInlineGuests();
      scheduleUnifiedViewer();
      scheduleWindowGeometry();
    }
  };
  const bindItemFrame = () => {
    const nextFrame = document.getElementById("calendar-item-panel-iframe");
    if (nextFrame === itemFrame) {
      return;
    }
    removeEditorTracking?.();
    removeEditorTracking = null;
    removeInlineGuests?.();
    removeInlineGuests = null;
    revealNativeEditor();
    itemFrame?.removeEventListener("load", onItemFrameLoad, true);
    itemFrame = nextFrame;
    if (itemFrame) {
      itemFrame.addEventListener("load", onItemFrameLoad, true);
      bindDescriptionEditor();
      bindInlineGuests();
      scheduleUnifiedViewer();
      if (
        itemFrame.contentDocument?.documentURI ===
          "chrome://calendar/content/calendar-item-iframe.xhtml" &&
        itemFrame.contentDocument.readyState === "complete"
      ) {
        scheduleWindowGeometry();
      }
    }
  };

  const observer = new window.MutationObserver(bindItemFrame);
  cleanups.push(() => {
    observer.disconnect();
    revealNativeEditor();
    try {
      removeInlineGuests?.();
    } catch (error) {
      console.error(
        "Outlook Style Companion could not clean the inline Guests editor:",
        error
      );
    }
    removeInlineGuests = null;
    try {
      removeEditorTracking?.();
    } catch (error) {
      console.error(
        "Outlook Style Companion could not clean the event description editor:",
        error
      );
    }
    removeEditorTracking = null;
    try {
      itemFrame?.removeEventListener("load", onItemFrameLoad, true);
    } catch (error) {
      /* The dynamically created item frame may already be disconnected. */
    }
    itemFrame = null;
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  bindItemFrame();
}

function enhanceEventSummarySurface(window, cleanups) {
  const document = window.document;
  /* The read-only summary uses the same dimensions as the event editor, but
   * Thunderbird calls sizeToContent() while populating it. Do not treat those
   * automatic resizes as a user preference change. */
  const geometry = installRememberedEventWindowGeometry(window, cleanups, {
    rememberUserResize: false,
  });
  installEventSummaryAvailabilitySurface(window, cleanups);
  const dialog = document.querySelector("dialog");
  const statusNotifications = document.getElementById(
    "status-notifications"
  );
  let fallbackTimer = 0;

  const scheduleAfterNativeLoad = () => {
    if (dialog?.classList.contains("resize")) {
      return;
    }
    /* Native onLoad removes .resize only after the item details and Fluent
     * strings have finished loading. Apply after that final layout pass. */
    geometry.schedule(100, true);
  };

  const dialogObserver = new window.MutationObserver(scheduleAfterNativeLoad);
  const onStatusTransitionEnd = event => {
    if (!statusNotifications?.contains(event.target)) {
      return;
    }
    /* Thunderbird's transitionend handler calls sizeToContent() on this same
     * event. A window-level bubbling listener runs afterward, so queue one
     * final forced application rather than racing the native handler. */
    window.removeEventListener("transitionend", onStatusTransitionEnd);
    if (fallbackTimer) {
      window.clearTimeout(fallbackTimer);
      fallbackTimer = 0;
    }
    geometry.schedule(0, true);
  };
  cleanups.push(() => {
    dialogObserver.disconnect();
    window.removeEventListener("transitionend", onStatusTransitionEnd);
    if (fallbackTimer) {
      window.clearTimeout(fallbackTimer);
    }
  });
  if (dialog) {
    dialogObserver.observe(dialog, {
      attributes: true,
      attributeFilter: ["class"],
    });
  }
  window.addEventListener("transitionend", onStatusTransitionEnd);

  scheduleAfterNativeLoad();
  /* Read-only items do not create a status notification, and an already
   * populated summary may not produce a class mutation visible to us. */
  fallbackTimer = window.setTimeout(() => {
    fallbackTimer = 0;
    geometry.schedule(0, true);
  }, 750);

}

function enhanceEditorSurfaceWindow(window) {
  const url = window.document?.location?.href;
  if (
    editorSurfaceState.has(window) ||
    (url !== COMPOSE_WINDOW_URL &&
      url !== SPELL_CHECK_DIALOG_WINDOW_URL &&
      url !== EVENT_DIALOG_WINDOW_URL &&
      url !== EVENT_ATTENDEES_DIALOG_WINDOW_URL &&
      url !== EVENT_SUMMARY_DIALOG_WINDOW_URL)
  ) {
    return editorSurfaceState.has(window);
  }

  const state = { cleanups: [] };
  editorSurfaceState.set(window, state);
  enhancedEditorSurfaceWindows.add(window);
  const { cleanups } = state;
  if (url === COMPOSE_WINDOW_URL) {
    enhanceComposeEditorSurface(window, cleanups);
  } else if (url === SPELL_CHECK_DIALOG_WINDOW_URL) {
    enhanceSpellCheckDialogSurface(window, cleanups);
  } else if (url === EVENT_DIALOG_WINDOW_URL) {
    enhanceEventEditorSurface(window, cleanups);
  } else if (url === EVENT_ATTENDEES_DIALOG_WINDOW_URL) {
    enhanceAttendeeDialogSurface(window, cleanups);
  } else {
    enhanceEventSummarySurface(window, cleanups);
  }

  return true;
}

function removeEditorSurfaceWindow(window) {
  const state = editorSurfaceState.get(window);
  if (!state) {
    return;
  }
  for (const cleanup of state.cleanups.reverse()) {
    try {
      cleanup();
    } catch (error) {
      /* One disconnected nested editor must not prevent the remaining owned
       * UI from being removed during disable, update, or window teardown. */
      console.error(
        "Outlook Style Companion could not completely clean an editor surface:",
        error
      );
    }
  }
  state.cleanups.length = 0;
  editorSurfaceState.delete(window);
  enhancedEditorSurfaceWindows.delete(window);
}

function installEditorSurfaceEnhancement() {
  if (editorSurfaceListenerRegistered) {
    return true;
  }

  const enhanceWindow = window =>
    enhanceWindowSafely(
      "an editor surface window",
      window,
      enhanceEditorSurfaceWindow,
      removeEditorSurfaceWindow
    );
  ExtensionSupport.registerWindowListener(EDITOR_SURFACE_LISTENER_ID, {
    chromeURLs: [
      COMPOSE_WINDOW_URL,
      SPELL_CHECK_DIALOG_WINDOW_URL,
      EVENT_DIALOG_WINDOW_URL,
      EVENT_ATTENDEES_DIALOG_WINDOW_URL,
      EVENT_SUMMARY_DIALOG_WINDOW_URL,
    ],
    onLoadWindow: enhanceWindow,
    onUnloadWindow: removeEditorSurfaceWindow,
  });
  editorSurfaceListenerRegistered = true;

  const windows = Services.wm.getEnumerator(null);
  while (windows.hasMoreElements()) {
    enhanceWindow(windows.getNext());
  }
  return true;
}

function enhanceReminderDialog(window) {
  if (
    window.document?.location?.href !== REMINDER_DIALOG_WINDOW_URL ||
    reminderDialogState.has(window)
  ) {
    return reminderDialogState.has(window);
  }

  const root = window.document.documentElement;
  const richlist = window.document.getElementById("alarm-richlist");
  if (!root || !richlist) {
    return false;
  }

  let decorationFrame = 0;
  const scheduleDecoration = () => {
    if (decorationFrame || window.closed) {
      return;
    }
    decorationFrame = window.requestAnimationFrame(() => {
      decorationFrame = 0;
      decorateReminderWidgets(window);
    });
  };
  const observer = new window.MutationObserver(scheduleDecoration);

  /* A physical double-click emits two click events before dblclick. The
   * native Details label handles each click, so suppress only its repeated
   * click; the first click still opens exactly the same window as before. */
  const onRepeatedDetailsClick = event => {
    if (event.button !== 0 || event.detail < 2) {
      return;
    }
    const details = event.composedPath?.().find(
      node => node?.matches?.(".alarm-details-label")
    );
    if (!details || !richlist.contains(details)) {
      return;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
  };

  const onDoubleClick = event => {
    if (event.button !== 0 || event.defaultPrevented) {
      return;
    }

    const path = event.composedPath?.() || [];
    const widget = path.find(
      node =>
        node?.matches?.(
          'richlistitem[is="calendar-alarm-widget-richlistitem"]'
        )
    );
    if (!widget || !richlist.contains(widget) || !widget.item) {
      return;
    }

    /* Leave links and controls alone. This prevents a double-click on the
     * existing Details link or a Snooze/Dismiss control from opening the
     * summary twice or swallowing the control's own action. */
    const interactiveSelector = [
      "button",
      "toolbarbutton",
      "menulist",
      "menupopup",
      "menuitem",
      "input",
      "textarea",
      "select",
      "a[href]",
      ".text-link",
      ".alarm-action-buttons",
      ".snooze-options-box",
      "[onclick]",
      "[oncommand]",
    ].join(",");
    for (const node of path) {
      if (node === widget) {
        break;
      }
      if (node?.matches?.(interactiveSelector)) {
        return;
      }
    }

    /* This is Thunderbird's native contract used by the Details link. The
     * alarm dialog's own listener opens the correct read-only summary or
     * editable event window for the calendar and occurrence. */
    widget.dispatchEvent(
      new window.Event("itemdetails", {
        bubbles: true,
        cancelable: false,
      })
    );
    event.preventDefault();
    event.stopPropagation();
  };
  reminderDialogState.set(window, {
    richlist,
    observer,
    widgets: new Map(),
    scheduleDecoration,
    onRepeatedDetailsClick,
    onDoubleClick,
    get decorationFrame() {
      return decorationFrame;
    },
  });
  enhancedReminderDialogWindows.add(window);

  observer.observe(richlist, { childList: true, subtree: true });
  richlist.addEventListener("click", onRepeatedDetailsClick, true);
  richlist.addEventListener("dblclick", onDoubleClick);
  root.setAttribute("data-outlook-style-reminder", "true");
  scheduleDecoration();
  return true;
}

function removeReminderDialogEnhancement(window) {
  const state = reminderDialogState.get(window);
  if (!state) {
    return;
  }

  state.observer.disconnect();
  state.richlist.removeEventListener(
    "click",
    state.onRepeatedDetailsClick,
    true
  );
  state.richlist.removeEventListener("dblclick", state.onDoubleClick);
  if (state.decorationFrame) {
    window.cancelAnimationFrame(state.decorationFrame);
  }
  for (const entry of state.widgets.values()) {
    removeReminderWidgetDecoration(entry);
  }
  state.widgets.clear();
  window.document.documentElement?.removeAttribute(
    "data-outlook-style-reminder"
  );
  reminderDialogState.delete(window);
  enhancedReminderDialogWindows.delete(window);
}

function installReminderDialogEnhancement() {
  if (reminderDialogListenerRegistered) {
    return true;
  }

  const enhanceWindow = window =>
    enhanceWindowSafely(
      "a reminder dialog",
      window,
      enhanceReminderDialog,
      removeReminderDialogEnhancement
    );
  ExtensionSupport.registerWindowListener(REMINDER_DIALOG_LISTENER_ID, {
    chromeURLs: [REMINDER_DIALOG_WINDOW_URL],
    onLoadWindow: enhanceWindow,
    onUnloadWindow: removeReminderDialogEnhancement,
  });
  reminderDialogListenerRegistered = true;

  const windows = Services.wm.getEnumerator("Calendar:AlarmWindow");
  while (windows.hasMoreElements()) {
    enhanceWindow(windows.getNext());
  }
  return true;
}

function centerWindowOverOpener(window) {
  if (window.closed) {
    return;
  }

  let opener = null;
  try {
    if (window.opener && !window.opener.closed) {
      opener = window.opener;
    }
  } catch (error) {
    /* A closing opener is equivalent to having no parent for placement. */
  }

  let targetScreen;
  try {
    targetScreen = opener?.screen || window.screen;
  } catch (error) {
    targetScreen = window.screen;
  }

  const finiteOr = (value, fallback) =>
    Number.isFinite(Number(value)) ? Number(value) : fallback;
  const availableLeft = finiteOr(
    targetScreen?.availLeft,
    finiteOr(targetScreen?.left, 0)
  );
  const availableTop = finiteOr(
    targetScreen?.availTop,
    finiteOr(targetScreen?.top, 0)
  );
  const availableWidth = Math.max(
    1,
    finiteOr(targetScreen?.availWidth, finiteOr(targetScreen?.width, 1))
  );
  const availableHeight = Math.max(
    1,
    finiteOr(targetScreen?.availHeight, finiteOr(targetScreen?.height, 1))
  );
  const chooserWidth = Math.max(1, finiteOr(window.outerWidth, 1));
  const chooserHeight = Math.max(1, finiteOr(window.outerHeight, 1));

  let targetLeft = availableLeft + (availableWidth - chooserWidth) / 2;
  let targetTop = availableTop + (availableHeight - chooserHeight) / 2;
  if (opener) {
    try {
      const openerLeft = finiteOr(opener.screenX, NaN);
      const openerTop = finiteOr(opener.screenY, NaN);
      const openerWidth = finiteOr(opener.outerWidth, NaN);
      const openerHeight = finiteOr(opener.outerHeight, NaN);
      if (
        Number.isFinite(openerLeft) &&
        Number.isFinite(openerTop) &&
        Number.isFinite(openerWidth) &&
        openerWidth > 0 &&
        Number.isFinite(openerHeight) &&
        openerHeight > 0
      ) {
        targetLeft = openerLeft + (openerWidth - chooserWidth) / 2;
        targetTop = openerTop + (openerHeight - chooserHeight) / 2;
      }
    } catch (error) {
      /* Fall back to the center of the opener's screen. */
    }
  }

  /* Do not clamp against zero: macOS and Windows both use negative screen
   * coordinates for monitors positioned left of or above the primary one. */
  const maximumLeft =
    availableLeft + Math.max(0, availableWidth - chooserWidth);
  const maximumTop =
    availableTop + Math.max(0, availableHeight - chooserHeight);
  const clampedLeft = Math.min(
    Math.max(targetLeft, availableLeft),
    maximumLeft
  );
  const clampedTop = Math.min(
    Math.max(targetTop, availableTop),
    maximumTop
  );
  try {
    window.moveTo(Math.round(clampedLeft), Math.round(clampedTop));
  } catch (error) {
    /* Keep the native dialog usable if a window manager rejects placement. */
  }
}

async function maybeAcceptDefaultInvitationCalendar(window, state) {
  if (
    state.disposed ||
    state.autoAcceptInProgress ||
    state.autoAcceptFinished ||
    window.closed ||
    window.document?.location?.href !== CALENDAR_CHOOSER_WINDOW_URL
  ) {
    return;
  }

  const args = window.arguments?.[0];
  const listbox = window.document.getElementById("calendar-list");
  const dialog = window.document.querySelector("dialog");
  if (!args || !listbox || !dialog) {
    return;
  }

  state.autoAcceptInProgress = true;
  try {
    /* This chrome dialog is shared by invitation import, file export, paste,
     * and publish. Compare the exact localized native invitation prompt so
     * only the iTIP import path may be accepted without another click. */
    const importPrompt = await window.document.l10n.formatValue(
      "import-prompt"
    );
    if (
      state.disposed ||
      window.closed ||
      args.promptText !== importPrompt ||
      !args.calendars
    ) {
      return;
    }

    let eligibleCalendars;
    try {
      eligibleCalendars = Array.from(args.calendars);
    } catch (error) {
      return;
    }

    let defaultCalendar = null;
    try {
      defaultCalendar = cal.view.getCompositeCalendar(window.opener)
        ?.defaultCalendar;
    } catch (error) {
      return;
    }
    if (!defaultCalendar?.id) {
      return;
    }

    const eligibleDefault = eligibleCalendars.find(
      calendar => calendar?.id === defaultCalendar.id
    );
    if (!eligibleDefault) {
      return;
    }
    try {
      if (
        !cal.manager.getCalendarById(eligibleDefault.id) ||
        eligibleDefault.getProperty("disabled") ||
        eligibleDefault.readOnly
      ) {
        return;
      }
    } catch (error) {
      return;
    }

    const matchingRow = [...listbox.children].find(
      row => row.calendar?.id === eligibleDefault.id
    );
    if (!matchingRow) {
      return;
    }

    listbox.selectedItem = matchingRow;
    listbox.ensureElementIsVisible?.(matchingRow);
    if (
      listbox.selectedItem !== matchingRow ||
      dialog.getButton("accept")?.disabled
    ) {
      return;
    }

    /* Use the dialog's native accept path so its dialogaccept listener runs
     * the caller-supplied onOk callback and assigns the iTIP target calendar. */
    state.autoAcceptFinished = true;
    dialog.acceptDialog();
  } catch (error) {
    console.error(
      "Outlook Style Companion could not use the default invitation calendar:",
      error
    );
  } finally {
    state.autoAcceptInProgress = false;
  }
}

function enhanceCalendarChooser(window) {
  if (
    window.document?.location?.href !== CALENDAR_CHOOSER_WINDOW_URL ||
    calendarChooserState.has(window)
  ) {
    return calendarChooserState.has(window);
  }

  const state = {
    disposed: false,
    autoAcceptInProgress: false,
    autoAcceptFinished: false,
    startTimer: 0,
    firstFrame: 0,
    secondFrame: 0,
    settleTimer: 0,
    schedulePlacement: null,
  };

  const clearScheduledPlacement = () => {
    if (state.startTimer) {
      window.clearTimeout(state.startTimer);
      state.startTimer = 0;
    }
    if (state.firstFrame) {
      window.cancelAnimationFrame(state.firstFrame);
      state.firstFrame = 0;
    }
    if (state.secondFrame) {
      window.cancelAnimationFrame(state.secondFrame);
      state.secondFrame = 0;
    }
    if (state.settleTimer) {
      window.clearTimeout(state.settleTimer);
      state.settleTimer = 0;
    }
  };

  state.schedulePlacement = () => {
    if (state.disposed || window.closed) {
      return;
    }
    clearScheduledPlacement();
    /* The native chooser adds its rows and performs a one-pixel resize in a
     * requestAnimationFrame. Two frames plus one short settling pass place
     * the final-sized dialog, including when the native resize fires later. */
    state.startTimer = window.setTimeout(() => {
      state.startTimer = 0;
      state.firstFrame = window.requestAnimationFrame(() => {
        state.firstFrame = 0;
        state.secondFrame = window.requestAnimationFrame(() => {
          state.secondFrame = 0;
          centerWindowOverOpener(window);
          void maybeAcceptDefaultInvitationCalendar(window, state);
          state.settleTimer = window.setTimeout(() => {
            state.settleTimer = 0;
            centerWindowOverOpener(window);
            void maybeAcceptDefaultInvitationCalendar(window, state);
          }, 150);
        });
      });
    }, 0);
  };

  calendarChooserState.set(window, state);
  enhancedCalendarChooserWindows.add(window);

  window.document.documentElement?.setAttribute(
    "data-outlook-style-calendar-chooser",
    "true"
  );
  window.addEventListener("resize", state.schedulePlacement);
  state.schedulePlacement();
  return true;
}

function removeCalendarChooserEnhancement(window) {
  const state = calendarChooserState.get(window);
  if (!state) {
    return;
  }
  state.disposed = true;
  window.removeEventListener("resize", state.schedulePlacement);
  if (state.startTimer) {
    window.clearTimeout(state.startTimer);
  }
  if (state.firstFrame) {
    window.cancelAnimationFrame(state.firstFrame);
  }
  if (state.secondFrame) {
    window.cancelAnimationFrame(state.secondFrame);
  }
  if (state.settleTimer) {
    window.clearTimeout(state.settleTimer);
  }
  window.document.documentElement?.removeAttribute(
    "data-outlook-style-calendar-chooser"
  );
  calendarChooserState.delete(window);
  enhancedCalendarChooserWindows.delete(window);
}

function installCalendarChooserEnhancement() {
  if (calendarChooserListenerRegistered) {
    return true;
  }

  const enhanceWindow = window =>
    enhanceWindowSafely(
      "a calendar chooser",
      window,
      enhanceCalendarChooser,
      removeCalendarChooserEnhancement
    );
  ExtensionSupport.registerWindowListener(CALENDAR_CHOOSER_LISTENER_ID, {
    chromeURLs: [CALENDAR_CHOOSER_WINDOW_URL],
    onLoadWindow: enhanceWindow,
    onUnloadWindow: removeCalendarChooserEnhancement,
  });
  calendarChooserListenerRegistered = true;

  const windows = Services.wm.getEnumerator(null);
  while (windows.hasMoreElements()) {
    enhanceWindow(windows.getNext());
  }
  return true;
}

var outlookThreadView = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    const cleanup = (description, callback) => {
      try {
        callback();
      } catch (error) {
        console.error(
          `Outlook Style Companion could not clean up ${description}:`,
          error
        );
      }
    };

    try {
      for (const window of [...enhancedTodayPaneWindows]) {
        cleanup("a Today pane window", () =>
          removeTodayPaneEnhancement(window)
        );
      }
      if (todayPaneListenerRegistered) {
        cleanup("the Today pane listener", () =>
          ExtensionSupport.unregisterWindowListener(TODAY_PANE_LISTENER_ID)
        );
        todayPaneListenerRegistered = false;
      }

      for (const window of [...enhancedMessageDisplayWindows]) {
        cleanup("a message display window", () =>
          removeMessageDisplaySurfaceBridge(window)
        );
      }
      if (messageDisplaySurfaceListenerRegistered) {
        cleanup("the message display listener", () =>
          ExtensionSupport.unregisterWindowListener(
            MESSAGE_DISPLAY_SURFACE_LISTENER_ID
          )
        );
        messageDisplaySurfaceListenerRegistered = false;
      }

      for (const window of [...enhancedMainPaneSchemeWindows]) {
        cleanup("a main-pane scheme window", () =>
          removeMainPaneSchemeBridge(window)
        );
      }
      if (mainPaneSchemeListenerRegistered) {
        cleanup("the main-pane scheme listener", () =>
          ExtensionSupport.unregisterWindowListener(
            MAIN_PANE_SCHEME_LISTENER_ID
          )
        );
        mainPaneSchemeListenerRegistered = false;
      }

      for (const conversationView of [...guardedConversationViews]) {
        cleanup("a guarded conversation", () =>
          retireGuardedConversationWithFallback(conversationView)
        );
      }
      for (const about3Pane of [...conversationSelectionCapturePanes]) {
        cleanup("a conversation selection listener", () =>
          removeConversationSelectionCapture(about3Pane)
        );
      }

      for (const window of [...enhancedReminderDialogWindows]) {
        cleanup("a reminder dialog", () =>
          removeReminderDialogEnhancement(window)
        );
      }
      if (reminderDialogListenerRegistered) {
        cleanup("the reminder dialog listener", () =>
          ExtensionSupport.unregisterWindowListener(
            REMINDER_DIALOG_LISTENER_ID
          )
        );
        reminderDialogListenerRegistered = false;
      }

      for (const window of [...enhancedEditorSurfaceWindows]) {
        cleanup("an editor surface window", () =>
          removeEditorSurfaceWindow(window)
        );
      }
      if (editorSurfaceListenerRegistered) {
        cleanup("the editor surface listener", () =>
          ExtensionSupport.unregisterWindowListener(EDITOR_SURFACE_LISTENER_ID)
        );
        editorSurfaceListenerRegistered = false;
      }

      for (const window of [...enhancedCalendarChooserWindows]) {
        cleanup("a calendar chooser window", () =>
          removeCalendarChooserEnhancement(window)
        );
      }
      if (calendarChooserListenerRegistered) {
        cleanup("the calendar chooser listener", () =>
          ExtensionSupport.unregisterWindowListener(
            CALENDAR_CHOOSER_LISTENER_ID
          )
        );
        calendarChooserListenerRegistered = false;
      }
    } finally {
      /* Experiment implementation scripts are cached outside the WebExtension
       * lifecycle. Invalidate that cache after a disable, uninstall, or update
       * even if a concurrently destroyed window could not be cleaned fully. */
      cleanup("the privileged implementation cache", () =>
        Services.obs.notifyObservers(null, "startupcache-invalidate")
      );
    }
  }

  getAPI(context) {
    return {
      outlookThreadView: {
        async installTodayPane() {
          const schemeBridgeInstalled = runEnhancementStep(
            "the main-pane scheme listener",
            installMainPaneSchemeEnhancement
          );
          const messageDisplayInstalled =
            runEnhancementStep(
              "the message display listener",
              installMessageDisplaySurfaceEnhancement
            );
          const todayPaneInstalled = runEnhancementStep(
            "the Today pane listener",
            installTodayPaneEnhancement
          );
          return (
            schemeBridgeInstalled &&
            messageDisplayInstalled &&
            todayPaneInstalled
          );
        },

        async installReminderDialog() {
          return installReminderDialogEnhancement();
        },

        async installEditorSurfaces() {
          return installEditorSurfaceEnhancement();
        },

        async installCalendarChooser() {
          return installCalendarChooserEnhancement();
        },

        async showParentThread(tabId) {
          const tab = context.extension.tabManager.get(tabId);
          if (!tab || tab.type !== "mail") {
            return false;
          }

          const about3Pane = tab.nativeTab?.chromeBrowser?.contentWindow;
          if (!about3Pane || about3Pane.location.href !== "about:3pane") {
            return false;
          }

          await about3Pane.hasDOMContentLoaded?.promise;
          const refreshConversation = () =>
            refreshConversationForPane(about3Pane);
          installConversationSelectionCapture(about3Pane, refreshConversation);
          return refreshConversation();
        },
      },
    };
  }
};
