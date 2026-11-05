# Outlook Web Styler and QOL Improvements

A Tampermonkey userscript that restyles Outlook on the web and adds several quality-of-life improvements.

## Install

1. Install Tampermonkey in your browser.
2. Open the raw [`outlook-web-styler.user.js`](https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-web-styler.user.js) file.
3. Tampermonkey will recognize the `.user.js` file and show its installation screen.
4. Choose **Install**.

Repeat those steps once on each browser or device where you use Outlook.

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

## Supported Outlook hosts

- `outlook.office.com`
- `outlook.office365.com`
- `outlook.live.com`
- `outlook.cloud.microsoft`

## License

[MIT](LICENSE)
