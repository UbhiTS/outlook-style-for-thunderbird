# Outlook Style for Thunderbird

Current package version: **1.0.52**.

An unofficial adaptive Outlook-style theme for Mozilla Thunderbird **153–154**
(minimum 153.0). It follows the operating system's light or dark appearance
automatically and uses Outlook-inspired Fluent colors in both modes.

The package uses two add-ons:

- **Outlook Style for Thunderbird** controls the app's colors, typography, spacing, mail list, reading pane, Today pane, and reminder-window appearance.
- **Outlook Style Companion** keeps displayed messages aligned with the active system appearance and adds the behaviors that a static theme cannot provide: parent-thread summaries, functional My Day controls, inline event guests, per-meeting snooze-until-start, and calendar-detail enhancements.

Together they provide:

- A Fluent-blue app/search bar and neutral Outlook-style tab strip
- Coordinated light and dark reading, compose, settings, calendar, and reminder surfaces
- Compact folder and message-list panes with subtle separators
- A grid-edge blue rail for unread messages, regular read-message styling, and an outlined Fluent selection
- Native indented thread guides and an Outlook-like full-message conversation accordion
- Segoe UI when available, with system-font fallbacks
- An Outlook-style Spaces application rail
- An Outlook My Day-inspired Calendar and To Do pane
- An Outlook-style Guests field and a polished Scheduling Assistant with one-click common-time selection
- A New Message/New Event split button and full event details after a timed calendar drag
- A readable adaptive reminder window with an exact **Until meeting starts — time** snooze action
- A readable adaptive event-details window, opened by either **Details…** or a reminder-row double-click

## Install or update

Install both files. Existing 1.0.x installations update in place because the internal add-on IDs have not changed.

1. Open Thunderbird.
2. Open **Menu (≡) → Add-ons and Themes**.
3. Click the gear button in Add-ons Manager.
4. Choose **Install Add-on From File…**.
5. Select `dist/outlook-style-for-thunderbird-1.0.52.xpi` and confirm.
6. Repeat **Install Add-on From File…** and select `dist/outlook-style-companion-1.0.52.xpi`.
7. Restart Thunderbird so already-open Settings, Add-ons, mail, and reminder windows reload with the new scheme.

The companion integrates with Thunderbird's native message reader, Today pane, and reminder dialog, so Thunderbird displays an elevated-access warning. It operates locally, contains no telemetry or network requests, and neither extracts nor transmits message or calendar data.

### Recovering from Companion 1.0.14

Companion 1.0.14 can crash Thunderbird during startup. If it is installed, start Thunderbird in Troubleshoot Mode (hold **Option** while launching on macOS), remove **Outlook Style Companion 1.0.14**, and restart normally before installing the current Companion. Do not update directly over an enabled 1.0.14 copy because its shutdown path contains the same unsafe operation. The visual theme can remain installed.

### Automatic light and dark modes

The two packages form one adaptive theme: there is no separate dark-theme file to install. Thunderbird follows the operating system preference through its native `system` color-scheme setting, while Outlook Style supplies matching Fluent light and dark palettes. Changes to the system appearance are reflected without switching add-ons.

For displayed HTML mail, the privileged Companion follows each live native reader into its final MIME document and colors only the document canvas and default inherited text. It does not request WebExtension mail-reading or scripting permissions, and it does not overwrite a sender's explicit body background or foreground. Deliberately authored mail can therefore retain its own appearance instead of being forcibly recolored.

To remove the visual theme, open **Add-ons and Themes → Themes**, find **Outlook Style for Thunderbird**, and choose **Disable** or **Remove**. Remove **Outlook Style Companion** from **Extensions** as well if you no longer want its message, thread, Today-pane, and reminder enhancements.

## Recommended Thunderbird layout

For the closest Outlook arrangement:

1. Choose **Cards View** in the message-list display menu.
2. Choose **Vertical View** so folders, the message list, and the reading pane form three columns.
3. Use **View → Density → Default**.
4. Keep the **Spaces Toolbar** visible.
5. Show the **Today Pane** for the Outlook My Day-style sidebar.

The companion applies Outlook-like folder/list/reader proportions once when upgrading from older layouts. Later splitter adjustments remain under your control and are saved normally by Thunderbird.

## Snooze until a meeting starts

Open an individual reminder's **Snooze for** menu and choose **Until meeting starts — 4:00 PM** (the time reflects that meeting). The option is shown only for a future timed event that Thunderbird can modify. It is intentionally omitted for tasks, all-day events, meetings that already started, and read-only calendars.

**Snooze All** remains duration-based because the reminders in the window may have different start times. If an event has multiple display alarms, Thunderbird's native alarm service applies the snooze to that event's alarms using its normal provider and recurrence handling.

## Build

On Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

The package builder supports Windows PowerShell 5.1 or newer. The complete
release gate also requires `git` and `node` for source-integrity and JavaScript
syntax checks.

The script creates deterministic, byte-identical `.xpi` and `.zip` files in
`dist/`, with `manifest.json` at each archive root, and writes
`dist/SHA256SUMS.txt`. Package entry ordering and timestamps are normalized so
the same source and toolchain produce the same bytes.

After each successful direct build, the XPI files, ZIP files, and checksum are
also copied to `\\ubhinas\Shared\thunderbird`. The checksum is copied last, so
it identifies a complete remote build set.

Run the complete local production gate before publishing:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\release-gate.ps1
```

The gate validates metadata, least privilege, local-only assets, licenses,
JavaScript syntax, archive contents, checksums, and two-pass reproducibility.

### GitHub Actions builds

The **Build add-ons** workflow runs on every push, for pull requests targeting
`main`, and by manual dispatch. It validates both manifests and JavaScript files,
builds the theme and Companion packages twice, requires byte-for-byte
reproducibility, verifies every archive and checksum against its source tree,
and publishes these workflow artifacts:

- `outlook-style-for-thunderbird-VERSION.xpi`
- `outlook-style-for-thunderbird-VERSION.zip`
- `outlook-style-companion-VERSION.xpi`
- `outlook-style-companion-VERSION.zip`
- `SHA256SUMS.txt`

The workflow only uploads Actions artifacts; it never creates or updates GitHub
Releases or tags. Publish a release manually whenever a build is ready.

See the [production release checklist](docs/RELEASE_CHECKLIST.md) for the clean
build, testing, integrity, publication, and rollback gates. Security reports and
data-handling details are documented in [SECURITY.md](SECURITY.md) and
[PRIVACY.md](PRIVACY.md). Development and review expectations are in
[CONTRIBUTING.md](CONTRIBUTING.md).

## Compatibility and limits

- Minimum version: Thunderbird 153.0.
- Maximum version: Thunderbird 154.x. Add-ons using Thunderbird Theme or
  Extension Experiments are required by ATN validation to declare an upper
  bound. After each new Thunderbird major passes clean-install, upgrade, and
  regression testing, this limit can be advanced through the add-on's ATN
  administration page.
- Detailed styling and companion behavior use internal Thunderbird interfaces.
  A future Thunderbird redesign may therefore require a theme update before its
  compatibility limit is advanced.
- Modern Thunderbird supports lightweight themes, not legacy complete themes. This project can closely reproduce Outlook's colors, density, surfaces, selection, hierarchy, and selected workflows, but it cannot replace Thunderbird with Outlook's Ribbon.
- The theme follows the operating system's light/dark preference for both application chrome and content. Message authors can still specify deliberate colors in their own HTML.
- Participant availability still depends on the selected calendar provider. Thunderbird's Google connection uses CalDAV, and Google's CalDAV service does not provide participant free/busy lookup; an empty Scheduling Assistant row can therefore mean either free or unavailable data. A true Google availability pane would require separate Google Calendar API authorization and network access, which this local companion intentionally does not request.

## Default calendar for invitations

Thunderbird already stores a native default calendar: open **Calendar** and
select the calendar you want to use in the calendar list. Outlook Style
Companion now uses that default automatically when it is one of Thunderbird's
eligible calendars for an invitation response. It never bypasses read-only,
disabled, identity, or scheduling checks. If the default is not eligible, the
normal **Select Calendar** dialog remains available and is centered over the
window that opened it instead of appearing in a screen corner.

## Event guests and Scheduling Assistant

New and editable events now include a **Guests** field directly below the title.
Enter an address (or an address-book list) and press Enter; Thunderbird's own
attendee and organizer objects are updated, so its normal **Send**, invitation,
notification, recurrence, and calendar-provider workflows remain authoritative.
Guests appear as removable chips, and read-only events keep these controls
disabled.

Use **Scheduling assistant** for optional attendees, rooms,
resources, roles, and provider-supplied free/busy details. The native dialog is
now larger, centered, adaptive in light and dark modes, and uses the same solid,
hatched, and unavailable status treatments as the calendar. **Find time** checks
15-minute increments from the selected start onward, preserves the event
duration, honors Thunderbird's configured work hours and days, and selects the
first slot within 30 days where every returned participant interval is free or
tentative. A warning remains because a blank row is ambiguous rather than proof
that a provider returned availability data.

## Version 1.0.52 calendar creation and scheduling

- Opens the native New Event editor immediately after a timed range is dragged
  in the day or week calendar instead of committing an unnamed event directly.
- Adds a real split button beside **New Message** with **New Message** and
  **New Event** commands, including the equivalent enhancement when Write is
  present in the unified toolbar.
- Adds **Find time** to Invite Attendees. It preserves the event duration,
  searches 15-minute increments during configured work hours and workdays, and
  selects the first slot within 30 days where returned attendee intervals are
  free or tentative.
- Uses solid blue for busy availability, blue diagonal hatching for tentative,
  and distinct unavailable and unknown treatments throughout Invite Attendees.
- Replaces Thunderbird's inherited circular font-size zoom artwork with
  explicit minus and plus controls, and aligns **Guests:** with the other event
  editor headings.
- Must be installed as both 1.0.52 packages; the earlier local 1.0.51 build used
  the same version number as its predecessor and could therefore remain cached
  during an in-place update.

## Version 1.0.51 full-window event canvas

- Uses the complete height of full event/task windows, expanding Description
  and Attendees into the available space with independent scrolling while
  keeping compact mail previews unchanged.

## Version 1.0.50 direct Edit handoff

- Makes **Edit** from the shared Details card open Thunderbird's editable form
  directly instead of passing through a second read-only card.

## Version 1.0.49 fully styled summary schedule

- Packages the nested availability timeline styles for CSP-restricted event
  summaries and removes the redundant native footer Edit button when the
  shared card supplies its own Edit action.

## Version 1.0.48 cross-surface event regression fix

- Moves the shared card rules into the packaged theme so Thunderbird's
  CSP-restricted summary dialog receives the same layout as mail and editor
  surfaces instead of displaying unstyled markup.
- Uses concise local calendar times without repeating the source timezone and
  labels invitation navigation as **Details** rather than **Edit**.
- Caps existing-item cards to their useful content height while retaining
  scrolling for descriptions or long attendee lists.

## Version 1.0.47 live-tested event-card layout

- Keeps shared event/task card rows top-aligned instead of stretching sparse
  items across a full-height window, and collapses an empty sidebar cleanly.
- Displays single-day all-day events as one day rather than exposing the
  exclusive following-day end date used internally by calendar providers.
- Adds a repeatable live Thunderbird smoke-check harness for create, view,
  edit, task, invitation, and standalone-window regression checks.

## Version 1.0.46 unified card in full details

- Applies the shared event/task card stylesheet at the native event-details
  document level, giving the full details/edit view the same card layout as
  mail preview and standalone message windows.

## Version 1.0.45 Outlook-style invitation commands

- Restyles the native invitation commands as a larger, flat Outlook command
  bar in both preview and standalone message windows.
- Keeps Thunderbird's context-sensitive actions intact: new invitations offer
  Accept, Tentative, and Decline, while already-processed updates correctly
  offer Details instead.

## Version 1.0.44 standalone invitation consistency

- Applies the shared event card to standalone message windows opened by a
  double-click, closing the last gap between message preview and full-window
  invitation views.

## Version 1.0.43 unified event and task view

- Uses one shared read-only event/task card for mail invitations, full event
  details, and existing Calendar or Tasks items. The card consistently shows
  item details, availability, attendees, and description where applicable.
- Keeps Thunderbird's native RSVP, reminder, recurrence, save, and task
  behavior authoritative. Selecting **Edit** reveals the native editor only
  when a change is needed; new items still open directly in that editor.

## Version 1.0.42 build artifact publishing

- Copies every successful direct build's XPI files, ZIP files, and checksum to
  `\\ubhinas\Shared\thunderbird`, verifying each copied artifact's SHA-256
  hash before publishing the checksum as the final completion marker.

## Version 1.0.41 event availability layout

- Shows the schedule timeline and traffic-light availability signal on both of
  Thunderbird's invitation layouts, including the classic inline iTIP message
  display used by existing profiles.
- Gives the full event-details view a dedicated schedule-and-attendees sidebar,
  keeping the event information wide while making every attendee visible in a
  larger scrollable list below the availability signal.

## Version 1.0.40 word-aware mention editing

- Makes an unmodified Backspace remove the mention's complete trailing name
  segment while preserving its linked email address. Once only the first name
  remains, Thunderbird's normal character deletion resumes.

## Version 1.0.39 visible keyboard selection

- Correctly stamps the active contact row as `data-active="true"`, making
  Up/Down movement visible on macOS as well as Windows and Linux.
- Scrolls the keyboard-selected contact into view when navigating a long list.

## Version 1.0.38 physical keyboard navigation

- Captures real Up/Down keyboard events from Gecko's privileged editor system
  event group, while retaining the ordinary listener path used by assistive and
  synthetic input across macOS, Windows, and Linux.

## Version 1.0.37 editable compose mention labels

- Lets the sender shorten a selected mention to a first name or a first and
  middle name while preserving the contact's underlying `mailto:` address.
- Keeps autocomplete closed while editing an existing mention label.

## Version 1.0.36 compose mention keyboard focus

- Gives the keyboard-selected contact a high-contrast focus ring that remains
  obvious in either system appearance.

## Version 1.0.35 compose mention keyboard navigation

- Handles Thunderbird's pre-consumed navigation events and protects the picker
  selection from the resulting editor selection update, so Up/Down selection is
  reliable across Windows, macOS, and Linux.

## Version 1.0.34 compose mention interaction

- Gives each contact a compact, two-line picker row and preserves long email
  addresses with ellipsis rather than clipping or enlarging the popup.
- Captures navigation keys from the compose editor itself, so Up/Down, Enter,
  Tab, and Escape reliably control the picker.
- Inserts a styled `mailto:` hyperlink for each selected mention. This is
  portable HTML email markup, rather than the plain text used previously.
- Adds a compact local-calendar timeline directly to meeting invitations and
  event-detail windows, so the surrounding confirmed meetings are visible
  before responding.
- Shows **green** when the time is clear, **red** for an overlapping accepted
  or personally-created meeting, and **yellow** when a confirmed meeting is
  within 30 minutes before or after. A neutral state is used when a provider
  cannot return enough local calendar data to make a reliable call.
- Uses only Thunderbird's already-synced calendar data; it does not request
  network access or send calendar details anywhere.

## Version 1.0.33 compose mention attachment

- Supports Thunderbird 154's compose editor, which lacks the `ownerGlobal`
  property used by the initial mention implementation.

## Version 1.0.32 compose mention picker position

- Opens the compose `@` contact picker immediately below the typed mention,
  instead of anchoring it below the full-height message editor where it could
  be off-screen.

## Version 1.0.31 reliable conversation expansion

- Recalculates an expanded conversation card after the nested reader and MIME
  body finish layout, instead of relying on Thunderbird's one-time height
  measurement.
- Includes the Outlook card's wrapper padding and borders and reserves a usable
  body viewport, preventing short messages with tall recipient headers from
  appearing header-only.
- Tracks later header, notification, attachment, and message-body size changes
  while the card remains open, with full cleanup when switching messages.

## Version 1.0.30 Thunderbird 154 compatibility

- Keeps the Thunderbird 153.0 minimum and raises the tested maximum to 154.x.
- Retains the `strict_max_version` required by ATN for Theme and Extension
  Experiments, while allowing future compatibility bumps through ATN after
  testing each Thunderbird major.

## Version 1.0.29 compose mentions and Spaces order

- Typing **@** in the body of a new message or reply now opens an address-book
  contact picker. Keep typing to narrow the list; use the arrow keys and
  Enter or Tab to insert a contact mention.
- Orders the left Spaces rail as **Mail**, **Calendar**, **Contacts**,
  **Tasks**, and **Chat**.

## Version 1.0.28 ATN compatibility

- Uses the Thunderbird Add-ons catalog's recognized 153.0 minimum version so
  both packages can pass catalog validation.
- Retains the strict 153.x maximum and makes no product-behavior or add-on-ID
  changes.

## Version 1.0.27 production hardening

- Shows **in 5 mins** through **in 1 min** before the next accepted meeting,
  and uses **now** only after that meeting has actually started.
- Gives an accepted meeting within five minutes priority over an older meeting
  still in progress, then advances automatically on the minute.
- Adds a full-width current-time line behind the left-aligned pill and gently
  pulses it during the five-minute approach and first three meeting minutes,
  with reduced-motion and forced-colors alternatives.
- Double-clicking a collapsed thread parent opens exactly the complete message
  currently shown in the reading pane, instead of opening one window or tab
  for every hidden reply.

## Version 1.0.26 consistent thread conversations

- Uses the full native conversation accordion for every single selected email
  thread, whether its row is expanded or collapsed and whether the selected
  row is the parent or a reply.
- Opens the newest message when a conversation is collapsed. In an expanded
  thread, selecting the parent or any reply opens that exact email in full.
- Keeps every other message available for in-place expansion without changing
  the selected row in the message list.
- Avoids Thunderbird's truncated legacy thread summary for indexed threads.

## Version 1.0.25 hollow unanswered meetings

- Shows unanswered invitations with a hollow calendar-colour rail instead of
  the segmented treatment.
- Keeps tentative meetings diagonal, accepted or personal meetings solid, and
  declined or delegated meetings hollow gray.

## Version 1.0.24 compose contacts and spelling

- Keeps **Add to To**, **Add to Cc**, and **Add to Bcc** on one responsive row
  when the Contacts sidebar is wide enough, then wraps them naturally at
  narrower widths.
- Centers Thunderbird's full spell-check window over the message being
  composed, including after Thunderbird finishes its delayed content sizing.
- Gives the spell checker coordinated Fluent fields, suggestions, buttons,
  disabled states, and light, dark, and forced-colors appearances.

## Version 1.0.23 distinct tentative and unanswered meetings

- Introduced separate patterned rails for tentative and unanswered meetings;
  version 1.0.25 later simplified unanswered meetings to a hollow blue rail.
- Preserves solid accepted rails and hollow declined/delegated rails across
  light, dark, and forced-colors modes.

## Version 1.0.22 accurate Up Next meeting

- Keeps the single **Up Next** pill on accepted invitations and meetings you
  created, instead of unanswered or tentative invitations.
- An accepted meeting within five minutes takes priority for its exact
  countdown; otherwise overlapping meetings follow the one that started most
  recently. The pill advances automatically on the minute.

## Version 1.0.21 full thread messages

- Replaces Thunderbird's permanently truncated parent-thread snippets with its
  native full-message conversation accordion when the thread index is ready.
- Opens any compact parent or reply in place with the complete trusted message
  reader; HTML, attachments, invitations, remote-content protection, and
  signed or encrypted messages remain handled by Thunderbird.
- Adds keyboard activation and focus treatment, preserves live light/dark
  switching, and safely keeps the complete parent reader when indexing is not
  available instead of showing a clipped or empty summary.

## Version 1.0.20 event guests and scheduling

- Adds a discoverable Outlook-style Guests field to the main event editor with
  native address-book autocomplete, mailing-list expansion, removable chips,
  and read-only protection.
- Restyles, enlarges, and centers Thunderbird's native Invite Attendees window,
  preserving its optional attendee, resource, room, role, time, and invitation
  behavior.
- Makes the free/busy grid coherent in light, dark, RTL, and forced-colors
  modes, and explains when a provider may not have returned availability data.
- Keeps all event saving and invitation delivery on Thunderbird's native path;
  the Companion adds no calendar-network or OAuth permissions.

## Version 1.0.19 message states and invitation calendar

- Reserves the grid-edge blue rail for unread messages and keeps their sender,
  subject, and time emphasized; read messages use regular text without a rail.
- Gives the selected message an Outlook-like light-blue fill and blue outline
  without confusing selection with unread state, including indented threads,
  Grouped by Sort, Table View, RTL, dark mode, and forced colors.
- Automatically uses Thunderbird's eligible native default calendar for
  invitation responses, while centering the calendar chooser whenever manual
  selection is still required.

## Version 1.0.18 meeting-response indicators

- Keeps accepted invitations and personally-created meetings on the existing solid calendar-colour rail.
- Uses a segmented rail for invitations awaiting a response and Outlook's
  diagonal rail for tentatively accepted meetings.
- Keeps declined and delegated meetings visible but gives them a muted hollow rail; declined titles are struck through.
- Adapts the response indicators to light, dark, and forced-colors modes without relying on colour alone.

## Version 1.0.17 event-summary window sizing

- Opens invitations and read-only event summaries from a double-click at the same practical size as the New Event editor.
- Reapplies the preferred size after Thunderbird finishes loading event details and its invitation notification, preventing a late native content-fit resize from shrinking the window.
- Leaves Accept, Tentative, Decline, recurrence, reminder, and read-only event behavior unchanged.

## Version 1.0.16 event-window and Up Next refinements

- Opens New Event and New Task in a practical 900 x 840 default window instead of Thunderbird's compact content minimum.
- Remembers later user resizing and restores the last normal event-window dimensions within the current display.
- Shows exactly one blue countdown pill above the relevant accepted or
  personally-created timed meeting, excluding ended, all-day, cancelled, and
  unaccepted invitations.

## Version 1.0.15 crash-safe macOS hotfix

- Removes the unsafe nested browsing-context override introduced in Companion 1.0.14, which could terminate Thunderbird during startup or extension shutdown.
- Keeps macOS live appearance synchronization through DOM theme-state markers and browser-host color-scheme propagation only.
- Preserves the Inbox, message-list, reading-pane, and thread-summary light/dark synchronization without writing protected Gecko state.

## Version 1.0.14 macOS live-appearance fix

- Keeps the already-open Inbox, folder pane, message list, reading pane, and thread summary synchronized when macOS switches between Light and Dark appearances.
- Bridges Thunderbird's live theme state across its nested `about:3pane` and `about:message` documents instead of depending on a stale child media query.
- Uses the Companion's privileged reader lifecycle to propagate that scheme into the final displayed MIME document while preserving sender-authored message colors and forced-colors accessibility behavior.

## Version 1.0.13 adaptive theme and editor-window fixes

- Adds coordinated Outlook-inspired Fluent light and dark palettes that follow the operating system automatically.
- Rebuilds New Message and New Event as consistent adaptive Fluent surfaces instead of mixed light, dark, and blue canvases.
- Keeps compose recipients, subject, formatting controls, Contacts, event fields, tabs, and editor toolbars readable in either appearance.
- Adapts displayed-message, message-composition, and event-description canvases through the Companion while preserving intentionally authored message colors.

## Version 1.0.12 message-list spacing

- Gives Threaded roots and Unthreaded cards the same 32 px Outlook-style leading gutter as Grouped by Sort.
- Preserves a second hierarchy inset for replies and moves Thunderbird's thread guide and nodes with the indented cards.
- Leaves grouped headers, right-side spacing, virtualized row heights, and Table View unchanged.

## Version 1.0.11 Today-pane controls

- Replaces the decorative month chevron with a functional Outlook-style month picker, including year navigation and a **Today** shortcut.
- Removes Thunderbird's side month arrows while the Companion is active; the native arrows return automatically if the Companion is disabled.
- Replaces the inactive ellipsis with a functional Outlook-style menu for Agenda/Day switching and visible-calendar selection.

## Version 1.0.10 polish

- Restores the visible Quick Filter and Table/Cards View icons in the white message-list header.
- Adds an Outlook-style inset around the message list while preserving Thunderbird's saved splitters, thread indentation, and virtualized row sizing.
- Refines the Today pane with consistent 16 px content insets, a shorter calendar and footer, full-height event color rails, Outlook-sized event cards, and clearer separation between times and titles.

## Version 1.0.9 fixes

- Restyles the separate event-summary window opened from a reminder so its notification, event fields, attendees, controls, and meeting description remain readable in the light Outlook palette.
- Opens Thunderbird's native event-summary or editor window when a reminder row is double-clicked, while leaving Snooze, Dismiss, links, and form controls untouched.

## Version 1.0.8 improvements

- Renames the packages to **Outlook Style for Thunderbird** and **Outlook Style Companion**, while preserving their existing internal IDs for in-place upgrades.
- Rebuilds Thunderbird's reminder window on a neutral light surface with readable primary, secondary, and disabled buttons.
- Styles the snooze menu and custom-duration controls to match the rest of the Outlook-style interface.
- Adds a functional per-event **Until meeting starts — time** command that snoozes to the occurrence's exact start using Thunderbird's native calendar alarm service.

## Version 1.0.7 fix

- Forces Thunderbird's isolated Account Central page to resolve text, icons, buttons, and backgrounds from the light palette, eliminating white-on-white content when the operating system prefers dark mode.

## Version 1.0.6 improvements

- Rebalances wide vertical layouts toward Outlook's proportions: a compact folder pane and message list with a larger reading pane.
- Uses one blue app/search bar followed by a neutral tab strip with an Outlook-blue active underline and a neutral status bar.
- Gives unread messages blue subjects, dates, and leading rails; active messages use neutral-gray selection while retaining the blue rail.
- Presents the active message on a lightly inset white reading card with compact actions and a smaller subject hierarchy.

## Version 1.0.5 improvements

- Rebuilds the Today pane around Outlook's My Day layout, with direct Calendar and To Do tabs, calendar pop-out, roomier month grid, circular date selection, event rails, durations, and richer day headings.
- Adds a functional **Add a task due today** action and moves Thunderbird's real **New event** action into a bottom command bar.

## Version 1.0.4 fixes

- Draws the Outlook-blue leading rail for standalone messages, thread parents, and replies.
- Introduced parent-thread summaries; version 1.0.26 later replaced the
  truncated summary with the full native conversation accordion.

## Earlier fixes

- Brings Settings, Add-ons, Account Central, single-message, and multi-message surfaces into the shared Outlook-style palette.
- Restores 14 px inbox sender and subject text without breaking Thunderbird's virtualized row sizing.
- Preserves native thread indentation, hierarchy guides, and reply nodes with Outlook-style coloring.

## Branding

This project is independent and is not affiliated with, endorsed by, or sponsored by Microsoft or Mozilla. “Microsoft Outlook,” “Microsoft,” “Mozilla,” and “Thunderbird” are trademarks of their respective owners. No Microsoft product logos or font files are included.

The palette and styling are inspired by the public Microsoft Fluent 2 design system. Source code in this repository is released under the MIT License.
