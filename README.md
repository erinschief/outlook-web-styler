# Outlook Web Styler and QOL Improvements

Two independent Tampermonkey userscripts for Outlook on the web:

- `outlook-web-styler.user.js` restyles Outlook and adds the existing quality-of-life improvements.
- `outlook-inbox-widgets.user.js` adds the collapsible Unread, Schedule, Focus, and To-do widget above the personal inbox. It does not run in shared mailboxes and does not require the styler.

## Install

1. Install Tampermonkey in your browser.
2. Open the raw [`outlook-web-styler.user.js`](https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-web-styler.user.js) file.
3. Tampermonkey will recognize the `.user.js` file and show its installation screen.
4. Choose **Install**.

Repeat those steps once on each browser or device where you use Outlook.

### Install the optional inbox widgets

1. Open the raw [`outlook-inbox-widgets.user.js`](https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-inbox-widgets.user.js) file.
2. Choose **Install** on Tampermonkey's installation screen.
3. Reload the personal Outlook inbox.

The widget sweeps Outlook's virtualized personal-inbox list and restores the original scroll position. It follows a growing list until its bottom settles, then replaces its message cache. Inbox changes do not trigger automatic rescans while you triage; use **Refresh inbox data** when you want to rebuild the full widget cache. Scans stop safely if you navigate away, collapse the widget, or hide the tab. An interrupted scan keeps the previous cache and is never labelled complete. While collapsed, the summary reads Outlook's lightweight Inbox-folder totals instead of scanning the message list. A scan covers the currently displayed inbox view, including any native Outlook filter.

It follows the visible message-list container in both single-pane and multi-pane layouts. Open Outlook's native **My Day** pane when you want to refresh Schedule. Calendar parsing supports English all-day labels, 12/24-hour clocks, time ranges, and compound hour/minute durations. If no end time or duration is exposed, the event tooltip explicitly labels the 30-minute estimate. Schedule's time indicators update while visible; the script does not write to mail categories or the Outlook calendar.

The focus timer, display preferences, Self ordering, title edits, and local planning blocks are stored by Tampermonkey. Storage is separated by Outlook host and the account email exposed by Outlook's account control. Until an email is available, storage uses a separate tab-session key (survives reloads in that tab). Legacy settings migrate once to the first active scope. Expired local blocks are removed; title/order customizations absent from the inbox are retained for 90 days, so a temporary inbox filter does not erase them. The focus timer uses elapsed wall-clock time across reloads and sleep, stops at each phase boundary, and waits for you to start the next phase.

Background changes defer widget replacement while you edit a title, enter a time-block name, or drag. Scroll positions are retained on other refreshes. The timer updates its clock and ring in place instead of rebuilding the entire widget.

### Apps menu keyboard controls

Focus the header apps button and press **Space**, **Arrow Down**, or **Arrow Right** to open the flyout. Use arrow keys, **Home**, and **End** to select an app; **Enter** activates it. **Escape** closes the flyout and returns focus to the header button. Enter on the header button keeps its existing Back to inbox action. Closed flyout buttons are excluded from keyboard navigation.

## Configure a shared mailbox

The repository does not contain a mailbox address. On each device:

1. Open the Tampermonkey menu while Outlook is open.
2. Open this script's **Configure shared mailbox** command.
3. Enter the shared mailbox email address and the label to show in the apps flyout.

The setting is stored locally by Tampermonkey and is preserved when the script updates.

## Automatic updates

The userscript metadata points `@updateURL` and `@downloadURL` at the raw file in this repository. Tampermonkey periodically checks that URL and compares the published `@version` with the installed version.

For every release:

1. Edit `outlook-web-styler.user.js`.
2. Increase `@version` (for example, `0.67` to `0.68`).
3. Commit and push the change to the `main` branch.

Tampermonkey will offer or install the newer version according to each browser's update settings. You can also open Tampermonkey's dashboard and run **Check for userscript updates** to fetch it immediately.

Do not rename or move `outlook-web-styler.user.js` without also updating the two metadata URLs and reinstalling the script on each device.

The widget script updates independently from its own raw URL. Increase its `@version` whenever `outlook-inbox-widgets.user.js` changes.

## Supported Outlook hosts

- `outlook.office.com`
- `outlook.office365.com`
- `outlook.live.com`
- `outlook.cloud.microsoft`

## Validation

Run the dependency-free automated checks with Node.js 20 or newer:

```sh
node --check outlook-web-styler.user.js
node --check outlook-inbox-widgets.user.js
node --test tests/regression.test.cjs
```

GitHub Actions runs syntax and regression checks on pushes and pull requests. The tests exercise both full-script startup paths with a mocked loading shell, plus scanning, cancellation, cache reconciliation, message routing, calendar boundaries, time parsing, storage isolation, editing protection, pointer cancellation, timer updates, observer filtering, and flyout accessibility state.

For manual Chromium smoke testing against a synthetic Outlook DOM:

```sh
node tests/build-browser-fixture.cjs
```

Open the printed local file URL in Chromium; it displays its own pass/fail results. The fixture uses the actual script sources with test-only exports and a simulated Outlook URL and Tampermonkey storage. It does not access a real mailbox. Live Outlook selectors, visual layout, and browser-specific behavior still need verification in Outlook; isolated checks do not establish live compatibility.

## Releases

- **Widgets 0.3.5:** preserves the inbox snapshot across search, folder navigation, visibility changes, and collapse/reopen; full virtual-list traversal now occurs only for an empty first load or an explicit refresh.
- **Widgets 0.3.4:** changes the oldest-unread connector to a straight vertical drop, one compact rounded right-angle turn, and a straight horizontal run to the label.
- **Widgets 0.3.3:** restores the mockup's curved dashed connector from the center of the oldest-unread timeline dot to its endpoint label, with the connector layered behind all timeline text.
- **Styler 0.72:** explicitly clips the outer Inbox paint surface to its top-corner radii, preventing Outlook's translucent header layer from visually squaring off the left corner.
- **Styler 0.71:** mirrors Outlook's native right-side message-list inset on the left in list-only mode, giving both top corners equal room to render their rounding.
- **Styler 0.70:** keeps the multi-pane alignment shift but disables it in list-only mode, preventing `#MainModule` from moving outside the viewport and clipping the message-list and widget left corners.
- **Styler 0.69:** targeted observers, batched and rate-limited sign-in handling, idempotent promotion hiding, stable app IDs, keyboard flyout controls, and cleanup of listeners when Outlook replaces the header.
- **Widgets 0.3.2:** uses one theme-derived color token for both the oldest-unread dot and its matching endpoint label.
- **Widgets 0.3.1:** preserves the 0.3.0 stability work while making full inbox scans manual during triage, using live folder-tree counts while collapsed, simplifying the refresh glyph, and softening the oldest-unread marker.
- **Widgets 0.3.0:** safe cached-message opening, transactional growing-list scans, calendar-correct buckets, incremental timer updates and persistence, batched scraping, edit/drag protection, account-scoped storage, and bounded customization retention.

## License

[MIT](LICENSE)

