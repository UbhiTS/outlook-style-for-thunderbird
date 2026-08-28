import net from "node:net";
import fs from "node:fs";

class MarionetteClient {
  constructor(port = 2828) {
    this.port = port;
    this.socket = null;
    this.buffer = Buffer.alloc(0);
    this.messages = [];
    this.waiters = [];
    this.nextId = 1;
  }

  async connect() {
    this.socket = net.createConnection({ host: "127.0.0.1", port: this.port });
    this.socket.on("data", chunk => this.#onData(chunk));
    await new Promise((resolve, reject) => {
      this.socket.once("connect", resolve);
      this.socket.once("error", reject);
    });
    return this.#nextMessage();
  }

  #onData(chunk) {
    this.buffer = Buffer.concat([this.buffer, chunk]);
    while (true) {
      const separator = this.buffer.indexOf(58);
      if (separator < 0) {
        return;
      }
      const length = Number(this.buffer.subarray(0, separator).toString("ascii"));
      if (!Number.isInteger(length) || length < 0) {
        throw new Error("Invalid Marionette frame length.");
      }
      const frameEnd = separator + 1 + length;
      if (this.buffer.length < frameEnd) {
        return;
      }
      const payload = this.buffer.subarray(separator + 1, frameEnd).toString("utf8");
      this.buffer = this.buffer.subarray(frameEnd);
      const message = JSON.parse(payload);
      const waiter = this.waiters.shift();
      if (waiter) {
        waiter.resolve(message);
      } else {
        this.messages.push(message);
      }
    }
  }

  #nextMessage() {
    if (this.messages.length) {
      return Promise.resolve(this.messages.shift());
    }
    return new Promise((resolve, reject) => this.waiters.push({ resolve, reject }));
  }

  async command(name, params = {}) {
    const id = this.nextId++;
    const payload = JSON.stringify([0, id, name, params]);
    this.socket.write(`${Buffer.byteLength(payload, "utf8")}:${payload}`);
    const response = await this.#nextMessage();
    if (!Array.isArray(response) || response[0] !== 1 || response[1] !== id) {
      throw new Error(`Unexpected Marionette response: ${JSON.stringify(response)}`);
    }
    if (response[2]) {
      throw new Error(`${name} failed: ${JSON.stringify(response[2])}`);
    }
    return response[3];
  }

  close() {
    this.socket?.end();
  }
}

const client = new MarionetteClient();
try {
  const greeting = await client.connect();
  await client.command("WebDriver:NewSession", {
    capabilities: { alwaysMatch: {} },
  });
  await client.command("Marionette:SetContext", { value: "chrome" });
  const execute = script =>
    client.command("WebDriver:ExecuteScript", {
      script,
      args: [],
      newSandbox: true,
      sandbox: "default",
      line: 1,
      filename: "live-thunderbird-check.mjs",
    });
  const action = process.argv[2] || "audit";
  let result;
  if (action === "audit") {
    result = await execute(`
      const results = [];
      for (const host of Services.wm.getEnumerator(null)) {
        const contexts = host.browsingContext?.getAllBrowsingContextsInSubtree() || [];
        for (const context of contexts) {
          let child;
          try {
            child = context.window;
          } catch (error) {
            continue;
          }
          const document = child.document;
          const card = document.querySelector?.(".outlook-style-item-view");
          if (!card && document.documentURI !== "about:message") {
            continue;
          }
          const style = card ? child.getComputedStyle(card) : null;
          results.push({
            card: Boolean(card),
            cardDisplay: style?.display || "",
            cardGridColumns: style?.gridTemplateColumns || "",
            cardWidth: card?.getBoundingClientRect?.().width || 0,
            documentURI: document.documentURI,
            invitationBarVisible: !document.getElementById("imip-bar")?.collapsed,
            title: document.title,
          });
        }
      }
      return results;
    `);
  } else if (action === "select-invitation") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const tabmail = host.document.getElementById("tabmail");
      const pane = tabmail.currentTabInfo.chromeBrowser.contentWindow;
      let target = -1;
      for (let index = 0; index < pane.gDBView?.rowCount; index++) {
        const header = pane.gDBView.getMsgHdrAt(index);
        if (/^(updated\s+)?invitation:|meeting invitation|accepted:/i.test(header.mime2DecodedSubject || header.subject || "")) {
          target = index;
          break;
        }
      }
      if (target >= 0) {
        pane.threadTree.selectedIndex = target;
        pane.threadTree?.scrollToIndex?.(target);
      }
      const selectedHeader = target >= 0 ? pane.gDBView.getMsgHdrAt(target) : null;
      return {
        rowCount: pane.gDBView?.rowCount || 0,
        selected: target,
        subject: selectedHeader?.mime2DecodedSubject || selectedHeader?.subject || "",
      };
    `);
  } else if (action === "open-standalone") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      host.MsgOpenNewWindowForMessage?.();
      return true;
    `);
  } else if (action === "open-eml-tab") {
    const emlPath = process.argv[3];
    if (!emlPath) {
      throw new Error("open-eml-tab requires an EML path");
    }
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
      file.initWithPath(${JSON.stringify(emlPath)});
      const url = Services.io.newFileURI(file)
        .mutate()
        .setQuery("type=application/x-message-display")
        .finalize();
      host.document.getElementById("tabmail").openTab("mailMessageTab", {
        messageURI: url.spec,
      });
      return { path: file.path, url: url.spec };
    `);
  } else if (action === "open-create-event") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      host.createEventWithDialog(null, null, null, null, null, false, []);
      return true;
    `);
  } else if (action === "install-temporary-xpi") {
    const xpiPath = process.argv[3];
    if (!xpiPath) {
      throw new Error("install-temporary-xpi requires an XPI path");
    }
    result = await execute(`
      return (async () => {
        const { AddonManager } = ChromeUtils.importESModule(
          "resource://gre/modules/AddonManager.sys.mjs"
        );
        const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
        file.initWithPath(${JSON.stringify(xpiPath)});
        const addon = await AddonManager.installTemporaryAddon(file);
        return { id: addon.id, name: addon.name, version: addon.version };
      })();
    `);
  } else if (action === "inspect-calendar-creation") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const dropdown = host.document.querySelector(".outlook-style-new-item-dropdown");
      const tabmail = host.document.getElementById("tabmail");
      const pane = [...(tabmail?.tabInfo || []), tabmail?.currentTabInfo]
        .map(tab => tab?.chromeBrowser?.contentWindow)
        .find(candidate => candidate?.location?.href === "about:3pane");
      const folderDropdown = pane?.document.getElementById(
        "outlook-style-folder-new-item-dropdown"
      );
      return {
        controllerWrapped: /createEventWithDialog/.test(
          String(host.calendarViewController?.createNewEvent)
        ),
        dropdownCount: host.document.querySelectorAll(
          ".outlook-style-new-item-dropdown"
        ).length,
        dropdownLabel: dropdown?.getAttribute("aria-label") || "",
        folderDropdown: Boolean(folderDropdown),
        folderDropdownDisplay: folderDropdown
          ? pane.getComputedStyle(folderDropdown).display
          : "",
        folderDropdownStyle: folderDropdown
          ? {
              background: pane.getComputedStyle(folderDropdown).backgroundColor,
              height: folderDropdown.getBoundingClientRect().height,
              width: folderDropdown.getBoundingClientRect().width,
            }
          : null,
        writeButtonStyle: pane
          ? (() => {
              const button = pane.document.getElementById("folderPaneWriteMessage");
              return {
                height: button?.getBoundingClientRect().height || 0,
                width: button?.getBoundingClientRect().width || 0,
              };
            })()
          : null,
        folderPopupItems: Array.from(
          pane?.document.querySelectorAll(
            "#outlook-style-folder-new-item-popup > menuitem"
          ) || []
        ).map(item => item.getAttribute("label")),
        toolbarPresent: Boolean(host.document.querySelector("unified-toolbar")),
        toolbarItems: Array.from(
          host.document.querySelectorAll("#unifiedToolbarContent li[item-id]")
        ).map(item => ({
          hidden: item.hidden,
          id: item.getAttribute("item-id"),
          liveChildren: item.querySelector(".live-content")?.childElementCount || 0,
        })),
        popupItems: Array.from(
          host.document.querySelectorAll("#outlook-style-new-item-popup > menuitem")
        ).map(item => item.getAttribute("label")),
      };
    `);
  } else if (action === "open-attendees") {
    result = await execute(`
      const editor = Services.wm.getMostRecentWindow("Calendar:EventDialog");
      const frame = editor?.document.getElementById("calendar-item-panel-iframe");
      const canOpen = Boolean(frame?.contentWindow?.editAttendees);
      editor?.setTimeout(() => frame?.contentWindow?.editAttendees?.(), 0);
      return canOpen;
    `);
  } else if (action === "inspect-attendees") {
    result = await execute(`
      const attendeeWindow = Services.wm.getMostRecentWindow(
        "Calendar:EventDialog:Attendees"
      );
      const document = attendeeWindow?.document;
      const openerDocument = attendeeWindow?.opener?.document;
      const editorFrame = openerDocument?.getElementById(
        "calendar-item-panel-iframe"
      );
      const editorDocument = editorFrame?.contentDocument || openerDocument;
      const guestsHeading = editorDocument?.querySelector(
        "#outlook-style-guests-row > th"
      );
      const titleHeading = editorDocument?.querySelector(
        "#event-grid-title-row > th"
      );
      const styleOf = selector => {
        const element = document?.querySelector(selector);
        const style = element ? attendeeWindow.getComputedStyle(element) : null;
        return {
          backgroundColor: style?.backgroundColor || "",
          backgroundImage: style?.backgroundImage || "",
          border: style?.border || "",
          childCount: element?.childElementCount || 0,
          children: Array.from(element?.children || []).map(child => ({
            className: child.className || "",
            display: attendeeWindow.getComputedStyle(child).display,
            height: child.getBoundingClientRect().height,
            tag: child.localName,
            value: child.getAttribute("value") || "",
            width: child.getBoundingClientRect().width,
          })),
          height: element?.getBoundingClientRect().height || 0,
          label: element?.getAttribute("label") || "",
          listStyleImage: style?.listStyleImage || "",
          shadowText: element?.shadowRoot?.textContent?.trim() || "",
          text: element?.textContent?.trim() || "",
          width: element?.getBoundingClientRect().width || 0,
        };
      };
      return {
        findTime: Boolean(document?.getElementById("outlook-style-find-time")),
        findTimeLabel:
          document?.getElementById("outlook-style-find-time")?.textContent || "",
        attendeeCount: Array.from(
          document?.querySelectorAll("#attendee-list event-attendee") || []
        ).filter(row => row.attendee?.id).length,
        guestHeadingInlineStart: guestsHeading?.getBoundingClientRect().x || 0,
        titleHeadingInlineStart: titleHeading?.getBoundingClientRect().x || 0,
        tentative: styleOf('.legend[status="BUSY_TENTATIVE"]'),
        busy: styleOf('.legend[status="BUSY"]'),
        zoomIn: styleOf("#zoom-in-button"),
        zoomOut: styleOf("#zoom-out-button"),
      };
    `);
  } else if (action === "exercise-find-time") {
    result = await execute(`
      return (async () => {
        const attendeeWindow = Services.wm.getMostRecentWindow(
          "Calendar:EventDialog:Attendees"
        );
        const document = attendeeWindow?.document;
        const button = document?.getElementById("outlook-style-find-time");
        const status = document?.getElementById("outlook-style-find-time-status");
        button?.click();
        const deadline = Date.now() + 5000;
        while (button?.hasAttribute("aria-busy") && Date.now() < deadline) {
          await new Promise(resolve => attendeeWindow.setTimeout(resolve, 50));
        }
        return {
          busy: button?.hasAttribute("aria-busy") || false,
          disabled: button?.disabled || false,
          status: status?.textContent || "",
        };
      })();
    `);
  } else if (action === "exercise-zoom-controls") {
    result = await execute(`
      return (async () => {
        const attendeeWindow = Services.wm.getMostRecentWindow(
          "Calendar:EventDialog:Attendees"
        );
        const document = attendeeWindow?.document;
        const zoomIn = document?.getElementById("zoom-in-button");
        const zoomOut = document?.getElementById("zoom-out-button");
        const grid = document?.getElementById("freebusy-grid-inner");
        const before = grid?.scrollWidth || 0;
        zoomIn?.click();
        await new Promise(resolve => attendeeWindow.setTimeout(resolve, 0));
        const afterZoomIn = grid?.scrollWidth || 0;
        zoomOut?.click();
        await new Promise(resolve => attendeeWindow.setTimeout(resolve, 0));
        return {
          afterZoomIn,
          before,
          restored: grid?.scrollWidth || 0,
          zoomInLabel: zoomIn?.getAttribute("label") || "",
          zoomOutLabel: zoomOut?.getAttribute("label") || "",
        };
      })();
    `);
  } else if (action === "close-attendees") {
    result = await execute(`
      const attendeeWindow = Services.wm.getMostRecentWindow(
        "Calendar:EventDialog:Attendees"
      );
      attendeeWindow?.close();
      return Boolean(attendeeWindow);
    `);
  } else if (action === "close-new-event") {
    result = await execute(`
      const editors = [...Services.wm.getEnumerator("Calendar:EventDialog")];
      const editor = editors.find(candidate =>
        candidate.document.getElementById("calendar-item-panel-iframe")
          ?.contentWindow?.mode === "new"
      );
      editor?.close();
      return Boolean(editor);
    `);
  } else if (action === "open-existing-event" || action === "open-existing-task") {
    const wantTask = action === "open-existing-task";
    result = await execute(`
      return (async () => {
        const host = Services.wm.getMostRecentWindow("mail:3pane");
        const now = host.cal.dtz.now();
        const start = now.clone();
        start.addDuration(host.cal.createDuration("-P90D"));
        const end = now.clone();
        end.addDuration(host.cal.createDuration("P180D"));
        const filter = ${wantTask ? "Ci.calICalendar.ITEM_FILTER_TYPE_TODO" : "Ci.calICalendar.ITEM_FILTER_TYPE_EVENT"} |
          Ci.calICalendar.ITEM_FILTER_CLASS_OCCURRENCES;
        for (const calendar of host.cal.manager.getCalendars()) {
          if (calendar.getProperty("disabled")) {
            continue;
          }
          let items;
          try {
            items = await calendar.getItemsAsArray(filter, 0, start, end);
          } catch (error) {
            continue;
          }
          const item = items.find(candidate => candidate?.id);
          if (!item) {
            continue;
          }
          host.modifyEventWithDialog(item, false);
          return { calendar: calendar.name, id: item.id, title: item.title || "" };
        }
        return null;
      })();
    `);
  } else if (action === "open-synthetic-task") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const { CalTodo } = ChromeUtils.importESModule(
        "resource:///modules/CalTodo.sys.mjs"
      );
      const calendar = host.cal.manager.getCalendars().find(candidate =>
        !candidate.readOnly && !candidate.getProperty("disabled")
      );
      const task = new CalTodo();
      task.id = "outlook-style-live-smoke-task";
      task.calendar = calendar;
      task.title = "Outlook Style live task check";
      task.status = "IN-PROCESS";
      task.entryDate = host.cal.dtz.now();
      task.dueDate = task.entryDate.clone();
      task.dueDate.addDuration(host.cal.createDuration("P1D"));
      task.descriptionText = "Temporary unsaved task used to verify the shared view and editor.";
      host.modifyEventWithDialog(task, false);
      return { calendar: calendar?.name || "", id: task.id, title: task.title };
    `);
  } else if (action === "open-synthetic-summary") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const { CalEvent } = ChromeUtils.importESModule(
        "resource:///modules/CalEvent.sys.mjs"
      );
      const { CalAttendee } = ChromeUtils.importESModule(
        "resource:///modules/CalAttendee.sys.mjs"
      );
      const calendar = host.cal.manager.getCalendars().find(candidate =>
        !candidate.readOnly && !candidate.getProperty("disabled")
      );
      const event = new CalEvent();
      event.id = "outlook-style-live-smoke-summary";
      event.calendar = calendar;
      event.title = "Outlook Style live summary check";
      event.status = "CONFIRMED";
      event.startDate = host.cal.dtz.now();
      event.endDate = event.startDate.clone();
      event.endDate.addDuration(host.cal.createDuration("PT1H"));
      event.setProperty("LOCATION", "Test Room");
      event.descriptionText = "Temporary unsaved event used to verify the details dialog.";
      const attendee = new CalAttendee();
      attendee.id = "mailto:attendee@example.com";
      attendee.commonName = "Example Attendee";
      attendee.participationStatus = "ACCEPTED";
      event.addAttendee(attendee);
      host.openDialog(
        "chrome://calendar/content/calendar-summary-dialog.xhtml",
        "_blank",
        "chrome,dialog,resizable",
        { calendarEvent: event, isInvitation: false, onOk() {} }
      );
      return { calendar: calendar?.name || "", id: event.id, title: event.title };
    `);
  } else if (action === "inspect-dialogs") {
    result = await execute(`
      return [...Services.wm.getEnumerator(null)].map(window => {
        const document = window.document;
        const frame = document.getElementById("calendar-item-panel-iframe");
        const card = document.querySelector(".outlook-style-item-view");
        const cardStyle = card ? window.getComputedStyle(card) : null;
        return {
          card: Boolean(card),
          cardDisplay: cardStyle?.display || "",
          cardGridColumns: cardStyle?.gridTemplateColumns || "",
          cardWidth: card?.getBoundingClientRect?.().width || 0,
          frameHidden: frame?.hidden ?? null,
          mode: frame?.contentWindow?.mode || "",
          title: document.title,
          toolboxHidden: document.getElementById("event-toolbox")?.hidden ?? null,
          url: document.location.href,
          windowType: document.documentElement.getAttribute("windowtype") || "",
        };
      });
    `);
  } else if (action === "focus-dialog") {
    result = await execute(`
      const dialogs = [...Services.wm.getEnumerator(null)].filter(window =>
        /calendar-(?:event|summary)-dialog\.xhtml$/.test(window.document.location.href)
      );
      const dialog = dialogs.at(-1);
      dialog?.focus();
      return { focused: Boolean(dialog), title: dialog?.document.title || "" };
    `);
  } else if (action === "focus-attendees") {
    result = await execute(`
      const dialog = Services.wm.getMostRecentWindow(
        "Calendar:EventDialog:Attendees"
      );
      dialog?.focus();
      return { focused: Boolean(dialog), title: dialog?.document.title || "" };
    `);
  } else if (action === "focus-message-window") {
    result = await execute(`
      const messageWindow = Services.wm.getMostRecentWindow("mail:messageWindow");
      messageWindow?.focus();
      return {
        focused: Boolean(messageWindow),
        title: messageWindow?.document.title || "",
      };
    `);
  } else if (action === "focus-main") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      host?.focus();
      return { focused: Boolean(host), title: host?.document.title || "" };
    `);
  } else if (action === "focus-3pane") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const tabmail = host.document.getElementById("tabmail");
      const tab = (tabmail?.tabInfo || []).find(
        candidate => candidate.chromeBrowser?.contentWindow?.location?.href === "about:3pane"
      );
      if (tab) {
        tabmail.switchToTab(tab);
      }
      host.focus();
      return { focused: Boolean(tab), title: host.document.title };
    `);
  } else if (action === "reload-3pane") {
    result = await execute(`
      const host = Services.wm.getMostRecentWindow("mail:3pane");
      const tabmail = host.document.getElementById("tabmail");
      const tab = (tabmail?.tabInfo || []).find(
        candidate => candidate.chromeBrowser?.contentWindow?.location?.href === "about:3pane"
      );
      tab?.chromeBrowser?.contentWindow?.location.reload();
      return Boolean(tab);
    `);
  } else if (action === "click-edit") {
    result = await execute(`
      const dialogs = [...Services.wm.getEnumerator(null)].filter(window =>
        /calendar-(?:event|summary)-dialog\.xhtml$/.test(window.document.location.href)
      );
      const dialog = dialogs.at(-1) || Services.wm.getMostRecentWindow("mail:messageWindow");
      let edit = dialog?.document.querySelector(".outlook-style-item-view-edit");
      if (!edit) {
        for (const context of dialog?.browsingContext?.getAllBrowsingContextsInSubtree() || []) {
          edit = context.window.document.querySelector?.(".outlook-style-item-view-edit");
          if (edit) {
            break;
          }
        }
      }
      edit?.click();
      return { clicked: Boolean(edit), title: dialog?.document.title || "" };
    `);
  } else if (action === "close-dialog") {
    result = await execute(`
      const dialogs = [...Services.wm.getEnumerator(null)].filter(window =>
        /calendar-(?:event|summary)-dialog\.xhtml$/.test(window.document.location.href)
      );
      for (const dialog of dialogs) {
        dialog.close();
      }
      return { closed: dialogs.length };
    `);
  } else if (action === "close-message-window") {
    result = await execute(`
      const windows = [...Services.wm.getEnumerator("mail:messageWindow")];
      for (const window of windows) {
        window.close();
      }
      return { closed: windows.length };
    `);
  } else if (action === "exercise-unread-thread-defaults") {
    result = await execute(`
      return (async () => {
        const { MailServices } = ChromeUtils.importESModule(
          "resource:///modules/MailServices.sys.mjs"
        );
        if (!MailServices.accounts.localFoldersServer) {
          MailServices.accounts.createLocalMailAccount();
        }
        const root = MailServices.accounts.localFoldersServer.rootFolder;
        const folderName = "Unread Thread Defaults " + Date.now();
        root.createSubfolder(folderName, null);
        const folder = root.getChildNamed(folderName);
        const localFolder = folder.QueryInterface(Ci.nsIMsgLocalMailFolder);
        const message = ({ id, inReplyTo = "", read, subject }) => [
          "From: Test Sender <sender@example.invalid>",
          "To: Test Recipient <recipient@example.invalid>",
          "Date: Thu, 27 Aug 2026 12:00:00 -0700",
          "Message-ID: <" + id + ">",
          ...(inReplyTo
            ? [
                "In-Reply-To: <" + inReplyTo + ">",
                "References: <" + inReplyTo + ">",
              ]
            : []),
          "Subject: " + subject,
          "X-Mozilla-Status: " + (read ? "0001" : "0000"),
          "X-Mozilla-Status2: 00000000",
          "Content-Type: text/plain; charset=UTF-8",
          "",
          "Thread default regression fixture.",
          "",
        ].join("\\r\\n");

        localFolder.addMessage(message({
          id: "read-root@example.invalid",
          read: true,
          subject: "Fully read thread",
        }));
        localFolder.addMessage(message({
          id: "read-reply@example.invalid",
          inReplyTo: "read-root@example.invalid",
          read: true,
          subject: "Re: Fully read thread",
        }));
        localFolder.addMessage(message({
          id: "unread-root@example.invalid",
          read: true,
          subject: "Thread with unread reply",
        }));
        localFolder.addMessage(message({
          id: "unread-reply@example.invalid",
          inReplyTo: "unread-root@example.invalid",
          read: false,
          subject: "Re: Thread with unread reply",
        }));

        Services.prefs.setIntPref(
          "mailnews.default_view_flags",
          Ci.nsMsgViewFlagsType.kThreadedDisplay
        );
        const host = Services.wm.getMostRecentWindow("mail:3pane");
        const tabmail = host.document.getElementById("tabmail");
        let pane = tabmail.currentAbout3Pane;
        if (!pane) {
          tabmail.openTab("mail3PaneTab", { folderURI: folder.URI });
          const paneDeadline = Date.now() + 10000;
          while (
            (!pane ||
              pane.location.href !== "about:3pane" ||
              typeof pane.displayFolder !== "function") &&
            Date.now() < paneDeadline
          ) {
            await new Promise(resolve => host.setTimeout(resolve, 50));
            pane = tabmail.currentAbout3Pane;
          }
        }
        if (!pane || typeof pane.displayFolder !== "function") {
          throw new Error("Timed out opening the mail three-pane");
        }
        const loaded = new Promise((resolve, reject) => {
          const timeout = pane.setTimeout(
            () => reject(new Error("Timed out loading thread fixture")),
            10000
          );
          pane.addEventListener(
            "allMessagesLoaded",
            () => {
              pane.clearTimeout(timeout);
              pane.setTimeout(resolve, 100);
            },
            { once: true }
          );
        });
        pane.displayFolder(folder.URI);
        await loaded;

        const rows = [];
        for (let index = 0; index < pane.gDBView.rowCount; index++) {
          if (
            pane.gDBView.getLevel(index) !== 0 ||
            !pane.gDBView.isContainer(index)
          ) {
            continue;
          }
          const thread = pane.gDBView.getThreadContainingIndex(index);
          let unread = 0;
          for (let child = 0; child < thread.numChildren; child++) {
            if (!(thread.getChildHdrAt(child).flags & Ci.nsMsgMessageFlags.Read)) {
              unread++;
            }
          }
          rows.push({
            open: pane.gDBView.isContainerOpen(index),
            subject: thread.getRootHdr().mime2DecodedSubject,
            unread,
          });
        }
        const readThread = rows.find(row => row.subject === "Fully read thread");
        const unreadThread = rows.find(
          row => row.subject === "Thread with unread reply"
        );
        return {
          folder: folder.URI,
          passed: Boolean(
            readThread &&
            !readThread.open &&
            unreadThread &&
            unreadThread.open
          ),
          rows,
        };
      })();
    `);
  } else if (action === "screenshot") {
    const outputPath = process.argv[3];
    if (!outputPath) {
      throw new Error("screenshot requires an output path");
    }
    const base64 = await client.command("WebDriver:TakeScreenshot", { full: false });
    fs.writeFileSync(outputPath, Buffer.from(base64.value || base64, "base64"));
    result = { outputPath };
  } else {
    throw new Error(`Unknown action: ${action}`);
  }
  process.stdout.write(`${JSON.stringify({ action, greeting, result }, null, 2)}\n`);
  await client.command("WebDriver:DeleteSession", {});
} finally {
  client.close();
}
