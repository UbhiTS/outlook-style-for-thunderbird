# Outlook Style for Thunderbird

Current package version: **1.0.20**.

An unofficial adaptive Outlook-style theme for Mozilla Thunderbird **153.0.3 and later**. It follows the operating system's light or dark appearance automatically and uses Outlook-inspired Fluent colors in both modes.

The package uses two add-ons:

- **Outlook Style for Thunderbird** controls the app's colors, typography, spacing, mail list, reading pane, Today pane, and reminder-window appearance.
- **Outlook Style Companion** keeps displayed messages aligned with the active system appearance and adds the behaviors that a static theme cannot provide: parent-thread summaries, functional My Day controls, inline event guests, per-meeting snooze-until-start, and calendar-detail enhancements.

Together they provide:

- A Fluent-blue app/search bar and neutral Outlook-style tab strip
- Coordinated light and dark reading, compose, settings, calendar, and reminder surfaces
- Compact folder and message-list panes with subtle separators
- A grid-edge blue rail for unread messages, regular read messages, and an outlined Fluent selection
- Native indented thread guides and Outlook-like thread summaries
- Segoe UI when available, with system-font fallbacks
- An Outlook-style Spaces application rail
- An Outlook My Day-inspired Calendar and To Do pane
- An Outlook-style Guests field and a polished, one-click Scheduling Assistant for events
- A readable adaptive reminder window with an exact **Until meeting starts — time** snooze action
- A readable adaptive event-details window, opened by either **Details…** or a reminder-row double-click

## Install or update

Install both files. Existing 1.0.x installations update in place because the internal add-on IDs have not changed.

1. Open Thunderbird.
2. Open **Menu (≡) → Add-ons and Themes**.
3. Click the gear button in Add-ons Manager.
4. Choose **Install Add-on From File…**.
5. Select `dist/outlook-style-for-thunderbird-1.0.20.xpi` and confirm.
6. Repeat **Install Add-on From File…** and select `dist/outlook-style-companion-1.0.20.xpi`.
7. Restart Thunderbird so already-open Settings, Add-ons, mail, and reminder windows reload with the new scheme.

The companion integrates with Thunderbird's native message reader, Today pane, and reminder dialog, so Thunderbird displays an elevated-access warning. It operates locally, contains no telemetry or network requests, and neither extracts nor transmits message or calendar data.

### Recovering from Companion 1.0.14

Companion 1.0.14 can crash Thunderbird during startup. If it is installed, start Thunderbird in Troubleshoot Mode (hold **Option** while launching on macOS), remove **Outlook Style Companion 1.0.14**, and restart normally before installing the current Companion. Do not update directly over an enabled 1.0.14 copy because its shutdown path contains the same unsafe operation. The visual theme can remain installed.

### Automatic light and dark modes

The two packages form one adaptive theme: there is no separate dark-theme file to install. Thunderbird follows the operating system preference through its native `system` color-scheme setting, while Outlook Style supplies matching Fluent light and dark palettes. Changes to the system appearance are reflected without switching add-ons.

For displayed HTML mail, the Companion colors the document canvas and default inherited text, but does not overwrite a sender's explicit body background or foreground. Deliberately authored mail can therefore retain its own appearance instead of being forcibly recolored.

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

The script creates matching `.xpi` and `.zip` files in `dist/`, with `manifest.json` at each archive root.

### GitHub Actions builds

The **Build add-ons** workflow runs for pull requests, pushes to `main`, manual
dispatches, and version tags. It validates both manifests and JavaScript files,
builds the theme and Companion packages, verifies every archive against its
source tree, and publishes these workflow artifacts:

- `outlook-style-for-thunderbird-VERSION.xpi`
- `outlook-style-for-thunderbird-VERSION.zip`
- `outlook-style-companion-VERSION.xpi`
- `outlook-style-companion-VERSION.zip`
- `SHA256SUMS.txt`

After a successful `main` build, the workflow automatically creates a tag and
GitHub Release named `vVERSION` when that manifest version has not been
published before, and attaches the same five verified files. Rebuilding an
existing version replaces its published packages and checksum and moves that
version tag to the successful build commit. Manually dispatching the workflow
on `main` performs the same verified rebuild-and-update process.

## Compatibility and limits

- Minimum version: Thunderbird 153.0.3.
- No maximum version is declared, so later Thunderbird versions can install it.
- Detailed styling and companion behavior use internal Thunderbird interfaces. A future Thunderbird redesign may require an update.
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
now larger, centered, adaptive in light and dark modes, and includes a warning
that a blank row is ambiguous rather than proof that someone is available.

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
- Uses an Outlook-style diagonal rail for invitations awaiting a response and tentatively accepted meetings.
- Keeps declined and delegated meetings visible but gives them a muted hollow rail; declined titles are struck through.
- Adapts the response indicators to light, dark, and forced-colors modes without relying on colour alone.

## Version 1.0.17 event-summary window sizing

- Opens invitations and read-only event summaries from a double-click at the same practical size as the New Event editor.
- Reapplies the preferred size after Thunderbird finishes loading event details and its invitation notification, preventing a late native content-fit resize from shrinking the window.
- Leaves Accept, Tentative, Decline, recurrence, reminder, and read-only event behavior unchanged.

## Version 1.0.16 event-window and Up Next refinements

- Opens New Event and New Task in a practical 900 x 840 default window instead of Thunderbird's compact content minimum.
- Remembers later user resizing and restores the last normal event-window dimensions within the current display.
- Shows exactly one blue countdown pill above the first active or upcoming timed meeting, excluding ended, all-day, cancelled, declined, and delegated entries.

## Version 1.0.15 crash-safe macOS hotfix

- Removes the unsafe nested browsing-context override introduced in Companion 1.0.14, which could terminate Thunderbird during startup or extension shutdown.
- Keeps macOS live appearance synchronization through DOM theme-state markers and browser-host color-scheme propagation only.
- Preserves the Inbox, message-list, reading-pane, and thread-summary light/dark synchronization without writing protected Gecko state.

## Version 1.0.14 macOS live-appearance fix

- Keeps the already-open Inbox, folder pane, message list, reading pane, and thread summary synchronized when macOS switches between Light and Dark appearances.
- Bridges Thunderbird's live theme state across its nested `about:3pane` and `about:message` documents instead of depending on a stale child media query.
- Propagates the selected scheme into the displayed-message browser while preserving sender-authored message colors and forced-colors accessibility behavior.

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
- Selecting an expanded thread parent shows the complete thread in Thunderbird's native multi-message reader while keeping only the parent row selected.

## Earlier fixes

- Brings Settings, Add-ons, Account Central, single-message, and multi-message surfaces into the shared Outlook-style palette.
- Restores 14 px inbox sender and subject text without breaking Thunderbird's virtualized row sizing.
- Preserves native thread indentation, hierarchy guides, and reply nodes with Outlook-style coloring.

## Branding

This project is independent and is not affiliated with, endorsed by, or sponsored by Microsoft or Mozilla. “Microsoft Outlook,” “Microsoft,” “Mozilla,” and “Thunderbird” are trademarks of their respective owners. No Microsoft product logos or font files are included.

The palette and styling are inspired by the public Microsoft Fluent 2 design system. Source code in this repository is released under the MIT License.
