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

The widget automatically sweeps Outlook's virtualized personal-inbox list, caches the complete set of message rows, and restores the original list position. Use the widget's refresh button to rescan after large mailbox changes. It follows the visible message-list container in both single-pane and multi-pane layouts. Open Outlook's native **My Day** pane once when you want to refresh the Schedule page. The focus timer, display preferences, Self ordering, and planning blocks are stored locally by Tampermonkey. The script does not write to mail categories or the Outlook calendar.

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

## License

[MIT](LICENSE)
