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

The header observes Outlook's lightweight Inbox-folder unread count whether expanded or collapsed, batching sidebar changes into count-only updates. Collapsed widgets do not scrape messages, walk the conversation cache, or render hidden panels. The collapsed toolbar shows only the unread count; if the sidebar is unavailable, it retains the last confirmed value. The expanded toolbar counts unique conversations in the cached Inbox view, never the sidebar’s individual-message total. Message arrivals, read/unread changes, category edits, and scrolling automatically update the visible conversation cache in short batches without moving the inbox. A five-second fallback checks visible data while expanded when Outlook misses a mutation. Background updates pause in hidden tabs; returning to the tab refreshes visible data. Counts update in place even while you edit or drag.

The Unread timeline and To-do cards are conversation snapshots, not whole-mailbox message totals. Outlook virtualizes off-screen rows, so archived/deleted conversations and off-screen changes can remain in the snapshot until a full refresh. Native filters and Focused/Other views also limit what rows are available. The header uses the concise label **conversations** for the local snapshot total; its tooltip explains the separate unread-count source. A completed full refresh reconciles the conversation total with the current Inbox view. Folder-title parsing currently supports English Inbox labels.

On the first expanded personal-Inbox load in a page session, the widget performs one full refresh of the current Inbox view without opening **My Day**. It follows a growing list until its bottom settles, replaces the cache only after completion, and restores the original scroll position. If the widget was saved collapsed, the startup refresh waits until you open it.

After startup, message arrivals, read/unread changes, categories, and scrolling use the lightweight live updater without moving the inbox. Use **Full refresh of the current inbox view** whenever you want to reconcile the whole displayed list again; that action does not open My Day. Opening the **Schedule** tab refreshes its cache from My Day, then closes the pane if the widget opened it. If My Day was already open, the widget leaves it open. Rendering changes trigger adaptive settling rather than a fixed wait at every step. Scans stop safely if you navigate away, collapse the widget, or hide the tab. An interrupted scan keeps the previous cache and retries the startup pass when the Inbox is ready again.

It follows the visible message-list container in both single-pane and multi-pane layouts. Calendar parsing supports English all-day labels, 12/24-hour clocks, time ranges, and compound hour/minute durations. If no end time or duration is exposed, the event tooltip explicitly labels the 30-minute estimate. Schedule's time indicators update while visible; the script does not write to mail categories or the Outlook calendar.

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
node --test tests/*.test.cjs
```

GitHub Actions runs syntax and regression checks on pushes and pull requests. The tests exercise both full-script startup paths with a mocked loading shell, plus live folder counts, incremental row changes, adaptive scan settling, scanning, cancellation, cache reconciliation, message routing, calendar boundaries, time parsing, storage isolation, editing protection, pointer cancellation, timer updates, observer filtering, and flyout accessibility state.

For manual Chromium smoke testing against a synthetic Outlook DOM:

```sh
node tests/build-browser-fixture.cjs
```

Open the printed local file URL in Chromium; it displays its own pass/fail results. The fixture uses the actual script sources with test-only exports and a simulated Outlook URL and Tampermonkey storage. It does not access a real mailbox. Live Outlook selectors, visual layout, and browser-specific behavior still need verification in Outlook; isolated checks do not establish live compatibility.

## Releases

- **Styler 0.73:** finds the Promotions subject tag without relying on Outlook's generated label class and updates it immediately when virtualized inbox rows are recycled.
- **Widgets 0.4.17:** counts unique cached conversations instead of the sidebar’s message total; collapsed mode displays only the live unread count.
- **Widgets 0.4.16:** keeps collapsed toolbar counts current from sidebar mutations without refreshing widget content; hides the collapsed conversation total when sidebar totals are unavailable.
- **Widgets 0.4.15:** restores the concise empty-state message, “No unread emails, woo!”
- **Widgets 0.4.14:** recognizes Outlook's zero-unread Inbox label when it omits the `(0 unread)` suffix, so the live header reaches zero without requiring a full refresh.
- **Widgets 0.4.13:** gives the oldest-conversation date the same ten-pixel clearance above the timeline's vertical endpoint cap as the angled relative-date labels at every width.
- **Widgets 0.4.12:** anchors the oldest-conversation date to the timeline's responsive axis so its transformed bottom edge stays four pixels above the line and its right edge remains flush at every width.
- **Widgets 0.4.11:** omits Outlook's Pinned group from the unread timeline and places pinned categorized conversations first in the matching Self and Waiting lists.
- **Widgets 0.4.10:** preserves both the names and displayed order of Outlook's own inbox group headers while the virtualized list scrolls, rather than deriving either from message dates or recycled-row arrival order.
- **Widgets 0.4.9:** uses Outlook's own inbox group labels and preserves the active group while the virtualized list scrolls, rather than recalculating relative-date sections independently.
- **Widgets 0.4.8:** places the rotated oldest-conversation date beside the timeline endpoint instead of leaving it detached above the axis.
- **Widgets 0.4.7:** keeps Inbox refreshes independent of My Day and refreshes the cached schedule only when the Schedule tab is opened.
- **Widgets 0.4.6:** omits empty relative-date buckets and their dividers, and rotates the oldest-date label as normal horizontal Segoe text so its glyph weight matches the other timeline labels.
- **Widgets 0.4.5:** keeps every rendered timeline bucket wide enough to own its marks, softens the vertical oldest-date label, closes My Day after script-initiated startup refreshes, and shortens the header label to conversations.
- **Widgets 0.4.4:** replaces read-message dots with low-emphasis micro ticks while retaining equal chronological spacing for every message inside each timeline bucket.
- **Widgets 0.4.3:** locks the vertical oldest-conversation date to the same theme-derived text color as the Today, Yesterday, and other timeline bucket labels.
- **Widgets 0.4.2:** preserves the last confirmed Inbox totals while My Day hides Outlook's folder tree and keeps adjusting unread changes even when the widget is collapsed.
- **Widgets 0.4.1:** restores one full current-view Inbox sweep on first expanded load, opens native My Day before that sweep, then returns to lightweight live updates for ongoing triage.
- **Widgets 0.4.0:** live Inbox-folder counts expanded and collapsed; batched background updates of visible conversations and My Day; immediate visible-row startup without a full sweep; adaptive full-refresh settling; explicit conversation-snapshot coverage and cached-count fallback labels.
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


