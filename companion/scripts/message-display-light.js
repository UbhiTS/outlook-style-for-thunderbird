/* Thunderbird also fires message-display hooks for its privileged
 * chrome:// multi-message summary. Leave that document to the theme itself:
 * inserting a WebExtension sheet there is both unnecessary and rejected by
 * Gecko. Only actual MIME message documents receive these adaptive defaults.
 *
 * Deliberately style the document canvas rather than `body`. That supplies a
 * readable Outlook fallback for ordinary messages while leaving foregrounds
 * and backgrounds intentionally provided by a message author untouched. */
(function applyOutlookAdaptiveMessageView() {
  try {
    const messageDocumentPattern =
      /^(?:imap|mailbox|news|nntp|snews|file):/i;
    if (!messageDocumentPattern.test(document.documentURI || "")) {
      return;
    }

    const applyStyle = () => {
      const root = document.documentElement;
      if (!root) {
        return;
      }
      root.classList.add("fluent-mail-light-message-view");

      if (!document.getElementById("fluent-mail-light-message-style")) {
        const style = document.createElement("style");
        style.id = "fluent-mail-light-message-style";
        style.textContent = `
          :root,
          html {
            --outlook-message-surface: #ffffff;
            --outlook-message-text: #242424;
            --outlook-message-link: #0f6cbd;
            --outlook-message-link-visited: #115ea3;
            color-scheme: light dark;
          }
          html.fluent-mail-light-message-view {
            background-color: var(--outlook-message-surface);
            color: var(--outlook-message-text);
          }
          html.fluent-mail-light-message-view[data-outlook-preserve-body-canvas] {
            background-color: transparent;
          }
          :where(a:link) {
            color: var(--outlook-message-link);
          }
          :where(a:visited) {
            color: var(--outlook-message-link-visited);
          }
          @media (prefers-color-scheme: dark) {
            :root,
            html {
              --outlook-message-surface: #292929;
              --outlook-message-text: #ffffff;
              --outlook-message-link: #62abf5;
              --outlook-message-link-visited: #96c6fa;
            }
          }
        `;
        root.prepend(style);
      }

      const preserveAuthoredBodyCanvas = () => {
        const body = document.body;
        if (!body) {
          return;
        }

        const bodyStyle = getComputedStyle(body);
        const hasOpaqueBackground =
          bodyStyle.backgroundColor !== "transparent" &&
          bodyStyle.backgroundColor !== "rgba(0, 0, 0, 0)";
        const hasBackgroundImage = bodyStyle.backgroundImage !== "none";
        root.toggleAttribute(
          "data-outlook-preserve-body-canvas",
          hasOpaqueBackground || hasBackgroundImage
        );
      };

      if (document.readyState === "loading") {
        document.addEventListener(
          "DOMContentLoaded",
          preserveAuthoredBodyCanvas,
          { once: true }
        );
      } else {
        preserveAuthoredBodyCanvas();
      }
    };

    if (document.documentElement) {
      applyStyle();
    } else {
      document.addEventListener("DOMContentLoaded", applyStyle, { once: true });
    }
  } catch (error) {
    console.error(
      "Outlook Style Companion could not apply its adaptive message style:",
      error
    );
  }
})();
