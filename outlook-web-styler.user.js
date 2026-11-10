// ==UserScript==
// @name         Outlook Web Styler and QOL Improvements
// @version      0.72
// @description  Outlook web: hides left rail, app launcher and Teams button; adds a collapsible apps menu (Mail/Calendar/To Do) in the top bar; floating full-width layout.
// @author       Erin
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.live.com/*
// @match        https://outlook.cloud.microsoft/*
// @updateURL    https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-web-styler.user.js
// @downloadURL  https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-web-styler.user.js
// @icon         data:image/png;base64,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
// @run-at       document-start
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// ==/UserScript==

(async () => {
    'use strict';

    // Positive main-shell gate: inject nothing unless Outlook renders its full title bar.
    // PWA compose/reply windows may be about:blank and may have no opener, so URL and
    // window.opener checks are not reliable enough to distinguish them.
    async function waitForPrimaryOutlookShell(timeoutMs = 10000) {
        const isMainShell = () =>
        !!document.getElementById('OwaTitleBar') ||
              !!document.querySelector('[data-app-section="TitleBar"]');

        if (isMainShell()) return true;

        return new Promise(resolve => {
            let settled = false;

            const finish = result => {
                if (settled) return;
                settled = true;
                observer.disconnect();
                clearTimeout(timer);
                resolve(result);
            };

            const observer = new MutationObserver(() => {
                if (isMainShell()) finish(true);
            });

            observer.observe(document.documentElement, {
                childList: true,
                subtree: true
            });

            const timer = setTimeout(() => finish(false), timeoutMs);
        });
    }

    if (!await waitForPrimaryOutlookShell()) return;
    document.documentElement.dataset.erinWindow = 'main';

    // Pop-out compose/reply windows fail the positive main-shell gate above, so no CSS or features run there.

    GM_addStyle(`
    :root { --erin-gap: 8px; --erin-radius: 8px; } /* 8px = --borderRadiusXLarge, the secondary toolbar's radius */

    /* Unify existing rounded corners. Fluent tokens are re-declared on each FluentProvider,
       so override them there (and on <html>) rather than only on :root. */
    html, .fui-FluentProvider {
      --borderRadiusMedium: var(--erin-radius) !important;
      --borderRadiusLarge: var(--erin-radius) !important;
    }
    #searchBoxColumnContainerId,
    #searchBoxColumnContainerId :is(form, div, input, button) { border-radius: var(--erin-radius) !important; }
    [data-app-section="MessageList"] [role="complementary"] {
      border-radius: var(--erin-radius) var(--erin-radius) 0 0 !important;
      /* Outlook's translucent header paint can remain square at the left edge even though this
         surface computes the correct radius. Clip the painted layer itself, not only its box. */
      clip-path: inset(0 round var(--erin-radius) var(--erin-radius) 0 0);
    }
    #LeftRail { display: none !important; }

    /* Applied only after the script verifies the MessageBar's complete visible label. */
    [data-erin-never-delete-retention] { display: none !important; }

    /* Floating full-width layout. Pin the app area to the viewport (below the 48px title bar)
       so no ancestor padding/scrollbar gutter can leave a bigger gap on one side.
       Equal left/right gap, no bottom gap; content is rounded at the top like Arc/Dia. */
    /* Outlook PWA pop-outs can clone this stylesheet and even the data-erin-window attribute.
       Require #OwaTitleBar in the same document so these layout rules cannot match a compose window. */
    html:has(#OwaTitleBar) #mainApp {
      position: fixed !important;
      top: 48px !important; bottom: 0 !important;
      left: var(--erin-gap) !important; right: var(--erin-gap) !important;
      width: auto !important; height: auto !important;
      padding: 0 !important; margin: 0 !important; box-sizing: border-box !important;
    }
    /* Only the mail column (the child that holds the Ribbon) fills the row. Other children, notably the
       empty spacer (.dbRkh) that reserves room for the right-hand side pane (My Day, To Do, ...), keep
       Outlook's own sizing; stretching them too gave the side pane half the window. */
    html:has(#OwaTitleBar) #mainApp > :has(> [data-app-section="Ribbon"]) {
      flex: 1 1 0 !important; width: auto !important; min-width: 0 !important;
      max-width: none !important; margin: 0 !important;
    }
    /* Right-side toolbar pane: reserve width only on the direct #mainApp child that actually owns the
       DrawerPane. The former selector could size multiple structural children and create a large unused gap. */
    html:has(#OwaTitleBar) #mainApp > [data-erin-drawer-host] {
      flex: 0 0 var(--erin-drawer-width, var(--flexPaneWidth, 320px)) !important;
      width: var(--erin-drawer-width, var(--flexPaneWidth, 320px)) !important;
      min-width: 0 !important; max-width: none !important;
      margin-left: var(--erin-gap) !important;
    }
    html:has(#OwaTitleBar)[data-erin-drawer="open"] #mainApp >
      :not(#LeftRail):not([data-erin-drawer-host]):not(:has(> [data-app-section="Ribbon"])) {
      flex-basis: 0 !important; width: 0 !important; min-width: 0 !important; max-width: 0 !important;
      margin: 0 !important; overflow: visible !important;
    }
    html:has(#OwaTitleBar) #MainModule { border-radius: var(--erin-radius) var(--erin-radius) 0 0 !important; overflow: hidden !important; }
    html:has(#OwaTitleBar) [data-app-section="NotificationPane"] { padding-right: var(--erin-gap) !important; }
    #owaAppLauncherBtn_container { display: none !important; }
    #owaTeamsBtn_container { display: none !important; }

    /* Right-hand button group. It is a Fluent Overflow container; its Teams item (hidden by us) still
       reserved a slot in Outlook's width math, so My Day and Settings were pushed into the "..." menu
       while ~50px sat empty. Take overflow out of the equation: keep My Day, bell and Settings always
       displayed (ID selectors beat Outlook's overflow hiding), drop the "..." menu button, and size the
       group to its three buttons. The page's own min-width is 418px, where all three buttons still fit; below 450px the
       "Outlook" word is hidden to make room. */
    div:has(> #owaHeaderButtonGroupRight_container) {
      flex: 0 0 auto !important; width: auto !important;
      min-width: 0 !important; max-width: none !important; margin-right: 0 !important;
    }
    #owaHeaderButtonGroupRight_container { width: auto !important; }
    #owaHeaderButtonGroupRight_container > [data-overflow-item]:not(#owaTeamsBtn_container) { display: flex !important; }
    #owaOverflowMenuBtn_container { display: none !important; }
    /* Outlook hides the "Outlook" label whenever search is active, even in wide windows where there is
       room. Show it again unless the compact (full-bar) search is open; hide it below 450px for space. */
    html:not([data-erin-search="open"]) #owaBranding_container { display: flex !important; }
    @media (max-width: 450px) { html:not([data-erin-search="open"]) #owaBranding_container { display: none !important; } }
    [data-erin-hide] { display: none !important; }
    [data-erin-drop-row] > :not([data-erin-hide]) { width: 100% !important; max-width: none !important; flex: 1 1 100% !important; }

    /* Thin scrollbars: Outlook's custom ones are 18px wide and eat into the panes */
    html:has(#OwaTitleBar) #mainApp ::-webkit-scrollbar { width: 8px !important; height: 8px !important; }
    html:has(#OwaTitleBar) #mainApp ::-webkit-scrollbar-track { background: transparent !important; }
    html:has(#OwaTitleBar) #mainApp ::-webkit-scrollbar-thumb { border-radius: 4px !important; }

    /* Reading pane scrolling. The scroller is [data-app-section="ConversationContainer"]
       (overflow:auto + scrollbar-gutter:stable + 18px custom scrollbar), so the track sat outside
       the cards. scrollbar-width overrides the ::-webkit-scrollbar styling: on macOS this gives
       the standard overlay scrollbar (floats over the card, no gutter); with a mouse it is a thin
       classic bar. */
    html:has(#OwaTitleBar) #mainApp :is(.customScrollBar, [data-is-scrollable="true"], [data-app-section="ConversationContainer"]) {
      scrollbar-width: thin !important;
      scrollbar-gutter: auto !important;
      scrollbar-color: rgba(128,128,128,.55) transparent !important;
    }

    /* Subject bar (X / title / arrows): inline margin is 8px 19px 8px 2px, where the 19px
       right margin pre-dates the gutter. Outlook now sometimes inserts another child before
       this row, so also follow the JS-tagged arrow controls to the actual direct child. */
    #ConversationReadingPaneContainer > :first-child,
    #ConversationReadingPaneContainer > :has([data-erin="prev"], [data-erin="next"]) {
      margin-right: 2px !important; border-radius: var(--erin-radius) !important;
    }
    #ConversationReadingPaneContainer > :first-child > div > div { border-radius: var(--erin-radius) !important; }

    /* Subject-bar controls, tagged by JS (data-erin) because the narrow-layout DOM has no stable hooks.
       Prev/next chevrons rotate 90deg clockwise: < becomes up, > becomes down. */
    [data-erin="close"] {
      translate: calc(-1 * var(--erin-close-shift, 4px)) 0;
      border-radius: var(--erin-radius) !important;
    }
    /* Prev/next: hide Outlook's chevron and draw a centered up/down chevron instead. Rotating the
       glyph left it optically off-center (the chevron is drawn off-center in its box), so the two
       arrows sat ~1px apart vertically. Both are drawn with the same box and path centre here. */
    :is([data-erin="prev"], [data-erin="next"]) {
      position: relative !important; font-size: 0 !important;
    }
    :is([data-erin="prev"], [data-erin="next"]) > * { visibility: hidden !important; }
    :is([data-erin="prev"], [data-erin="next"])::after {
      content: ""; position: absolute; left: 50%; top: 50%; width: 22px; height: 22px;
      margin: -11px 0 0 -11px; pointer-events: none;
      background: currentColor;
      -webkit-mask: var(--erin-chev) center / contain no-repeat;
      mask: var(--erin-chev) center / contain no-repeat;
    }
    [data-erin="prev"] { --erin-chev: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'><path d='M3.5 10.25L8 5.75l4.5 4.5' fill='none' stroke='black' stroke-width='1.3' stroke-linecap='round' stroke-linejoin='round'/></svg>"); }
    [data-erin="next"] { --erin-chev: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'><path d='M3.5 5.75L8 10.25l4.5-4.5' fill='none' stroke='black' stroke-width='1.3' stroke-linecap='round' stroke-linejoin='round'/></svg>"); }
    :is([data-erin="prev"], [data-erin="next"]):is([disabled], [aria-disabled="true"])::after { opacity: .35; }

    /* Pull the panes left to meet the ribbon's left edge (they sit ~11px further in). */
    html:has(#OwaTitleBar) #MainModule { margin-left: calc(-1 * var(--erin-pane-shift, 11px)) !important; }

    /* Single-column (list only) layout: Outlook sizes the list to a fixed pixel width, which left an
       empty strip on the right once the panes were shifted/widened. When the reading pane has no
       width (tagged by JS as data-erin-layout=list), let the list fill the row. */
    html[data-erin-layout="list"] [data-app-section="MessageList"],
    html[data-erin-layout="list"] [data-app-section="MessageList"] > [role="complementary"] {
      width: auto !important; max-width: none !important; flex: 1 1 auto !important;
    }
    /* In list-only mode the mail column already begins at the app's equal 8px viewport inset. The
       multi-pane -11px alignment shift moved MainModule to x=-3px, where its own overflow:hidden
       clipped the rounded left corners of both the message list and the widget inserted above it. */
    html[data-erin-layout="list"]:has(#OwaTitleBar) #MainModule {
      margin-left: 0 !important;
    }
    /* Outlook leaves a native 2px reveal only on the message surface's right side. Mirror that
       inset in list-only mode so the clipped surface exposes both rounded top corners equally. */
    html[data-erin-layout="list"] [data-app-section="MessageList"] > [role="complementary"] {
      margin-inline: 2px !important;
    }

    /* Message list header: Filter stays visible; only the sort-direction toggle is hidden (below). */
    /* Sort-direction toggle ("Sorted: Newest on top"). The menu button keeps the same aria-label prefix
       in its icon-only state, so exclude it by id / aria-haspopup. */
    [data-app-section="MessageList"] button[aria-label^="Sorted:"]:not(#mailListSortMenu):not([aria-haspopup]) {
      display: none !important;
    }

    /* Archive hover action (the strip at the row's right edge). Outlook only swaps the icon to a
       green fill on hover, which vanishes against the highlighted row. Make the whole strip a solid
       square-cornered target, tinted from the theme brand color (the New mail button's color) mixed
       toward white so it is lighter than Fluent's neon success green (#107c10) and follows theme
       changes. Icon turns white, is drawn at 20px instead of 24px, and after a short pause plays a
       quick hop with no scaling. Outlook's own lid-open animation still runs underneath. */
    [data-app-section="MessageList"] div[title="Archive"] {
      --erin-archive-mix: 82%; /* % of brand color; the rest is white. Lower = lighter. */
      transition: background-color .15s ease;
    }
    [data-app-section="MessageList"] div[title="Archive"] svg { width: 20px !important; height: 20px !important; }
    [data-app-section="MessageList"] div[title="Archive"]:hover {
      background-color: color-mix(in srgb,
        var(--erin-brand, var(--colorBrandBackground, #157956)) var(--erin-archive-mix), white) !important;
      --iconFill: var(--colorNeutralForegroundOnBrand, #fff);
    }
    [data-app-section="MessageList"] div[title="Archive"]:hover svg {
      animation: erin-archive-hop .4s cubic-bezier(.3, .7, .4, 1) .75s 1;
      transform-origin: 50% 60%;
    }
    @keyframes erin-archive-hop {
      0%   { transform: translateY(0) rotate(0); }
      25%  { transform: translateY(-3px) rotate(-12deg); }
      50%  { transform: translateY(0) rotate(10deg); }
      75%  { transform: translateY(-1.5px) rotate(-5deg); }
      100% { transform: translateY(0) rotate(0); }
    }
    @media (prefers-reduced-motion: reduce) {
      [data-app-section="MessageList"] div[title="Archive"]:hover svg { animation: none; }
    }

    /* Action toasts (Archived / Undo): Fluent positions the toaster with inline left/bottom and a
       fixed 292px width, tuned for the old left rail. JS publishes the message list's geometry so
       the toast is centered on the list and as wide as it, minus padding. */
    .fui-Toaster[data-toaster-position="bottom-start"] {
      left: var(--erin-toast-left, 15px) !important;
      width: var(--erin-toast-width, 292px) !important;
      bottom: var(--erin-toast-bottom, 16px) !important;
      transition: bottom .2s ease;
    }

    /* Folder-pane flyout (hamburger). It is a Fabric Panel whose root Outlook pins at left:48px (the old
       left-rail width). Fix it in CSS rather than measuring, so it does not fight the slide-in animation
       (the animated element is .ms-Panel-main; the root is static). Left corners squared to sit flush.
       The panel's ::before is a click-catcher/highlight over the hamburger button; it was also placed
       for the old layout (it landed on "File"), so JS publishes the hamburger's real rect. */
    .ms-Panel:has(#folderPaneDroppableContainer) { left: 0 !important; }
    .ms-Panel:has(#folderPaneDroppableContainer) .ms-Panel-main {
      border-top-left-radius: 0 !important; border-bottom-left-radius: 0 !important;
    }
    .ms-Panel:has(#folderPaneDroppableContainer)::before {
      left: var(--erin-ham-left, 8px) !important; top: var(--erin-ham-top, 8px) !important;
      width: var(--erin-ham-w, 40px) !important; height: var(--erin-ham-h, 34px) !important;
    }

    /* Person hover card (Live Persona Card): hidden while you are typing, so brushing a name does not pop it up.
       Its content sits in a slotted <div slot="background">; the host is that div's parent. Pure CSS off a flag
       that JS sets on keystrokes in an editable field and clears 2s after the last one, so it comes back on its own. */
    html:is([data-erin-typing], [data-erin-menu]) *:has(> [slot="background"] [data-testid="lpc-card-content"]),
    html:is([data-erin-typing], [data-erin-menu]) [data-testid="lpc-card-content"] { display: none !important; }

    /* Sort button ("By Date"): reuse an exact clone of Outlook's native Fluent icon whenever Outlook
       changes this control to its text variant. This preserves the original icon weight and rendering. */
    #mailListSortMenu[data-erin-sort-text] {
      position: relative !important; width: 32px !important; min-width: 32px !important;
      max-width: 32px !important; padding: 0 !important; overflow: hidden !important;
      white-space: nowrap !important; color: transparent !important; text-indent: -9999px !important;
    }
    /* The narrow list can render the text variant before Outlook has ever created its native icon.
       Show a small Fluent-style fallback immediately; the exact native clone replaces it once available. */
    #mailListSortMenu[data-erin-sort-text]:not(:has(> [data-erin-sort-icon]))::before {
      content: "" !important;
      position: absolute !important; left: 8px !important; top: 8px !important;
      width: 16px !important; height: 16px !important;
      background-color: var(--colorNeutralForeground2, currentColor) !important;
      -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath fill='black' d='M2.5 3.5a.5.5 0 0 1 .5-.5h10a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Zm0 4a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Zm0 4a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Z'/%3E%3C/svg%3E") center / contain no-repeat !important;
      mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath fill='black' d='M2.5 3.5a.5.5 0 0 1 .5-.5h10a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Zm0 4a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Zm0 4a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 0 1H3a.5.5 0 0 1-.5-.5Z'/%3E%3C/svg%3E") center / contain no-repeat !important;
      text-indent: 0 !important; pointer-events: none !important;
    }
    #mailListSortMenu[data-erin-sort-text] > [data-erin-sort-icon] {
      position: absolute !important; inset: 0 !important; display: flex !important;
      align-items: center !important; justify-content: center !important;
      width: 32px !important; height: 32px !important; margin: 0 !important;
      color: var(--colorNeutralForeground2, currentColor) !important;
      text-indent: 0 !important; pointer-events: none !important;
    }
    #mailListSortMenu[data-erin-sort-text] > [data-erin-sort-icon] svg { display: block !important; }

    /* Message cards */
    [data-app-section="ConversationContainer"] [aria-label="Email message"],
    [data-app-section="ConversationContainer"] [data-itempart-id] > div,
    [data-app-section="ConversationContainer"] .wide-content-host {
      border-radius: var(--erin-radius) !important;
    }

    /* Ribbon region ships with margin:0 21px 0 0 (sized for the old 18px scrollbar gutter).
       That is why the toolbar stopped ~20px short of the message view and collapsed items into
       the overflow chevrons early. Use 2px so its right edge lines up with the subject bar and
       message cards, which also sit 2px in. */
    [data-app-section="Ribbon"] { margin-right: 2px !important; }

    /* Header buttons (My Day, bell, gear, "...", search) ship with border-radius:0 and a full-height
       hover fill. Round them with the shared radius and inset the fill vertically like the apps button. */
    #owaHeaderButtonGroupRight_container > button {
      border-radius: var(--erin-radius) !important; height: 40px !important; align-self: center !important;
    }

    /* Search: Outlook swaps the field for an icon below ~718px of header width. Keep a field-looking
       control instead (the same collapsed button, restyled), centered in the bar; clicking still runs
       Outlook's own expand. JS supplies --erin-search-shift to center it on the whole title bar. */
    div:has(> div > #owaBranding_container):has(#searchBoxColumnContainerId > button[role="menuitem"])
      > div:has(> #owaBranding_container) { flex: 0 0 auto !important; min-width: 0 !important; }
    div:has(> #searchBoxColumnContainerId > button[role="menuitem"]) {
      flex: 1 1 0 !important; min-width: 0 !important; justify-content: center !important;
    }
    #searchBoxColumnContainerId:has(> button[role="menuitem"]) {
      height: auto !important; max-width: 380px; justify-content: center !important; overflow: visible !important;
      transform: translateX(var(--erin-search-shift, 0px));
    }
    /* Same metrics as Outlook's own field (32px, white, inset hairline shadow, icon at x=7, text at x=31),
       so the pill and the native field look identical as the window grows and Outlook swaps them. Max
       width is the native 380px. */
    #searchBoxColumnContainerId > button[role="menuitem"] {
      width: 100% !important; height: 32px !important; box-sizing: border-box !important;
      justify-content: flex-start !important; gap: 0; padding: 0 0 0 7px !important; overflow: hidden;
      border: 0 !important; border-radius: var(--erin-radius) !important;
      background: var(--white, var(--colorNeutralBackground1, #fff)) !important;
      box-shadow: var(--search-box-rest-shadow, inset 0 -1px var(--neutralSecondary), inset 0 0 0 1px var(--neutralTertiary)) !important;
      color: var(--headerTextIcons, #424242) !important;
      transition: box-shadow .15s ease;
    }
    #searchBoxColumnContainerId > button[role="menuitem"] > i { margin: 0 2px; }
    #searchBoxColumnContainerId > button[role="menuitem"]::after {
      content: "Search"; font: 14px/1 var(--fontFamilyBase, "Segoe UI", sans-serif); white-space: nowrap; margin-left: 2px;
    }
    #searchBoxColumnContainerId > button[role="menuitem"]:hover {
      box-shadow: inset 0 -1px var(--themeDark, #146348), inset 0 0 0 1px var(--themeDark, #146348) !important;
    }
    /* Opening animation: a clip-path reveal from the collapsed pill's box to the full bar. clip-path does
       not change layout and nothing is transformed, so it cannot fight Outlook's own sizing (the old
       scaleX version did). --erin-sx-il/ir are the pill's insets, captured on click (JS). */
    @keyframes erin-search-reveal {
      from { clip-path: inset(6px var(--erin-sx-ir, 0px) 6px var(--erin-sx-il, 0px) round var(--erin-radius)); }
      to   { clip-path: inset(0 0 0 0 round 0); }
    }
    html[data-erin-sx] div:has(> #searchBoxColumnContainerId [role="search"][aria-expanded="true"]),
    html[data-erin-sx] .x8k1J.kDnI3.FJmRM, html[data-erin-sx] [data-erin-compact] {
      animation: erin-search-reveal .2s cubic-bezier(.2, .8, .2, 1);
    }

    /* Apps toggle */
    #erin-apps-btn {
      all: unset; box-sizing: border-box; cursor: pointer;
      width: 40px; height: 40px; margin-left: 4px; border-radius: var(--erin-radius);
      display: grid; place-items: center; color: var(--headerTextIcons, var(--neutralPrimary, #242424));
    }
    #erin-apps-btn:hover { background: var(--themeDark, #146348); color: var(--copilotThemeWhite, #fff); }
    #erin-apps-btn[aria-expanded="true"] {
      background: var(--erin-brand, var(--colorBrandBackground, var(--themePrimary))); color: var(--erin-brand-fg, var(--colorNeutralForegroundOnBrand, #fff));
    }

    /* Compact search reuses the actual apps button rather than drawing a second copy. JS captures the
       button's normal viewport position before moving it to <body>, so the icon does not shift at all. */
    html[data-erin-search="open"] #erin-apps-btn {
      position: fixed !important; z-index: 2147483002 !important;
      left: var(--erin-apps-left, 4px) !important;
      top: var(--erin-apps-top, 4px) !important;
      margin: 0 !important; visibility: visible !important; display: grid !important;
    }
    /* Preserve Outlook's native compact-search width, max-width, right edge and right padding. The only
       layout change is moving its left edge just far enough right to uncover the existing apps button. */
    .x8k1J.kDnI3.FJmRM, [data-erin-compact] {
      left: var(--erin-search-left, 52px) !important;
    }
    /* Dropdown + field + glass read as one cell: square the inner corners, merge the shared hairline. */
    /* Same joint, CSS-only so it cannot lag behind the DOM (JS tagging above is the fallback). The scope
       dropdown only exists in the wide expanded state, so :has(#searchScopeButtonId) scopes it. The wrappers
       around the field (#searchBoxColumnContainerId, [role=search]) carry their own 8px radius and clip the
       field, so they are squared on the left too. The field is pulled 1px left so the dropdown's right
       border and the field's left hairline overlap into one line instead of two. */
    #OwaTitleBar #searchScopeButtonId { border-radius: var(--erin-radius) 0 0 var(--erin-radius) !important; }
    #OwaTitleBar:has(#searchScopeButtonId) :is(#searchBoxColumnContainerId, #searchBoxColumnContainerId > [role="search"], #searchBoxId-Mail) {
      border-top-left-radius: 0 !important; border-bottom-left-radius: 0 !important;
    }
    #OwaTitleBar:has(#searchScopeButtonId) #searchBoxColumnContainerId { margin-left: -1px !important; }
    #OwaTitleBar #searchBoxId-Mail:has(+ button[aria-label="Search"]) {
      border-top-right-radius: 0 !important; border-bottom-right-radius: 0 !important;
    }
    #OwaTitleBar #searchBoxId-Mail + button[aria-label="Search"] {
      border-radius: 0 var(--erin-radius) var(--erin-radius) 0 !important; margin-left: -1px !important;
    }
    #OwaTitleBar [data-erin-seg="first"] { border-radius: var(--erin-radius) 0 0 var(--erin-radius) !important; }
    #OwaTitleBar [data-erin-seg="mid"] { border-radius: 0 !important; }
    #OwaTitleBar [data-erin-seg="last"] { border-radius: 0 var(--erin-radius) var(--erin-radius) 0 !important; margin-left: -1px !important; }

    /* Expanding pill: fixed so ancestor overflow can't clip it; grows over the "Outlook" text */
    #erin-apps-panel {
      position: fixed; z-index: 2147483000; height: 40px; box-sizing: border-box;
      display: flex; align-items: center; gap: 2px; padding: 0 4px;
      border-radius: 20px; overflow: hidden;
      background: var(--erin-brand, var(--colorBrandBackground, var(--themePrimary))); color: var(--erin-brand-fg, var(--colorNeutralForegroundOnBrand, #fff));
      box-shadow: 0 2px 8px rgba(0,0,0,.25);
      width: 0; opacity: 0; pointer-events: none;
      transition: width .18s ease, opacity .12s ease, padding .18s ease;
    }
    #erin-apps-panel.open { width: max-content; max-width: calc(100vw - 56px); opacity: 1; pointer-events: auto; }
    #erin-apps-btn:focus-visible, #erin-apps-panel button:focus-visible { outline: 2px solid var(--erin-brand, #0f6cbd); outline-offset: 2px; }
    #erin-apps-panel button:focus-visible { outline-color: currentColor; outline-offset: -3px; }
    @media (prefers-reduced-motion: reduce) { #erin-apps-panel, #erin-apps-btn svg { transition: none !important; animation: none !important; } }
    #erin-apps-panel button {
      all: unset; box-sizing: border-box; cursor: pointer; flex: 0 0 auto;
      width: 40px; height: 32px; border-radius: 16px;
      display: grid; place-items: center; color: inherit;
    }
    #erin-apps-panel button:hover { background: rgba(255,255,255,.22); }
    #erin-apps-panel button[aria-pressed="true"] { background: rgba(0,0,0,.3); }
    #erin-apps-btn[aria-expanded="true"] svg.grid { transform: rotate(90deg); }
    /* Matches Outlook's own tooltip: white card, 4px radius, soft shadow, arrow on top */
    #erin-tip {
      position: fixed; z-index: 2147483001; pointer-events: none; box-sizing: border-box;
      padding: 4px 11px 6px; border-radius: var(--erin-radius); max-width: 240px;
      border: 1px solid var(--colorTransparentStroke, transparent);
      background: var(--colorNeutralBackground1, #fff);
      color: var(--colorNeutralForeground1, #242424);
      font-family: var(--fontFamilyBase, "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif);
      font-size: var(--fontSizeBase200, 12px); line-height: var(--lineHeightBase200, 16px);
      white-space: nowrap;
      filter: drop-shadow(0 0 2px rgba(0,0,0,.12)) drop-shadow(0 4px 8px rgba(0,0,0,.14));
      opacity: 0; transition: opacity .12s ease;
    }
    #erin-tip::before {
      content: ""; position: absolute; top: -4px; left: calc(var(--erin-arrow-x, 50%) - 4px);
      width: 8px; height: 8px; transform: rotate(45deg);
      background: inherit; border-radius: 1px;
    }
    #erin-tip.show { opacity: 1; }
    #erin-apps-btn svg, #erin-apps-panel svg {
      width: 20px; height: 20px; fill: none; stroke: currentColor;
      stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round;
    }
  `);

    const ICONS = {
        grid: '<svg viewBox="0 0 20 20"><rect x="2.5" y="2.5" width="15" height="15" rx="4"/><rect x="6" y="6" width="3" height="3" rx=".8"/><rect x="11" y="6" width="3" height="3" rx=".8"/><rect x="6" y="11" width="3" height="3" rx=".8"/><rect x="11" y="11" width="3" height="3" rx=".8"/></svg>',
        Mail: '<svg viewBox="0 0 20 20"><rect x="2.5" y="4.5" width="15" height="11" rx="2"/><path d="M3 6.5l7 5 7-5"/></svg>',
        Calendar: '<svg viewBox="0 0 20 20"><rect x="3" y="4.5" width="14" height="12" rx="2"/><path d="M3 8.5h14M7 2.5v3M13 2.5v3"/></svg>',
        'To Do': '<svg viewBox="0 0 20 20"><rect x="3" y="3" width="14" height="14" rx="3.5"/><path d="M7 10l2.2 2.2L13.2 8"/></svg>',
    };
    const getSharedMailbox = () => {
        const saved = GM_getValue('sharedMailbox', {});
        return {
            address: typeof saved.address === 'string' ? saved.address.trim().toLowerCase() : '',
            label: typeof saved.label === 'string' && saved.label.trim() ? saved.label.trim() : 'Shared Mailbox'
        };
    };
    let sharedMailbox = getSharedMailbox();

    GM_registerMenuCommand('Configure shared mailbox', () => {
        const address = prompt('Shared mailbox email address (leave blank to remove the button):', sharedMailbox.address);
        if (address === null) return;

        const normalizedAddress = address.trim().toLowerCase();
        let label = sharedMailbox.label;
        if (normalizedAddress) {
            const enteredLabel = prompt('Button label:', label || 'Shared Mailbox');
            if (enteredLabel === null) return;
            label = enteredLabel.trim() || 'Shared Mailbox';
        }

        GM_setValue('sharedMailbox', { address: normalizedAddress, label });
        location.reload();
    });

    const APPS = [
        { id: 'Mail', label: 'Mail' }, { id: 'Calendar', label: 'Calendar' }, { id: 'To Do', label: 'To Do' },
        ...(sharedMailbox.address ? [{ id: 'shared-mailbox', label: sharedMailbox.label, shared: true }] : [])
    ];

async function openSharedMailboxInbox() {
    const waitFor = async (get, timeout = 4000) => {
        const deadline = Date.now() + timeout;
        let value;
        while (!(value = get()) && Date.now() < deadline) {
            await new Promise(resolve => setTimeout(resolve, 50));
        }
        return value;
    };

    const rootId = `sharedFolderRoot_${sharedMailbox.address}`;
    let root = document.getElementById(rootId);
    if (!root) {
        [...document.querySelectorAll('button')]
            .find(button => button.getAttribute('aria-label') === 'Show navigation pane')
            ?.click();
        root = await waitFor(() => document.getElementById(rootId));
    }
    if (!root) return;

    if (root.getAttribute('aria-expanded') !== 'true') root.click();

    const inbox = await waitFor(() =>
        root.parentElement?.nextElementSibling?.querySelector('[data-folder-name="inbox"]')
    );
    inbox?.click();
}

    // Track the color of the ribbon's "New mail" button so the flyout follows the theme.
    const isClear = (c) => !c || c === 'transparent' || /rgba\(.*,\s*0\)$/.test(c);
    let storedBrand = null;
    function syncBrand() {
        const root = document.documentElement;
        let bg, fg;
        const nm = document.querySelector('button[label="New mail"]');
        if (nm && !nm.matches(':hover, :active')) {
            const cs = getComputedStyle(nm);
            if (!isClear(cs.backgroundColor)) { bg = cs.backgroundColor; fg = cs.color; }
        }
        if (!bg) {
            // New mail only renders on the Home tab; fall back to the Fluent brand token in scope.
            const prov = document.querySelector('[data-app-section="Ribbon"] .fui-FluentProvider') ||
                  document.querySelector('.fui-FluentProvider');
            if (prov) {
                const cs = getComputedStyle(prov);
                bg = cs.getPropertyValue('--colorBrandBackground').trim() || undefined;
                fg = cs.getPropertyValue('--colorNeutralForegroundOnBrand').trim() || undefined;
            }
        }
        if (bg) {
            if (root.style.getPropertyValue('--erin-brand') !== bg) root.style.setProperty('--erin-brand', bg);
            if (storedBrand !== bg) {
                try { localStorage.setItem('erin-brand', bg); storedBrand = bg; } catch (e) {}
            }
        }
        if (fg && root.style.getPropertyValue('--erin-brand-fg') !== fg) root.style.setProperty('--erin-brand-fg', fg);
    }
    // Paint last-known color immediately on load to avoid a green flash.
    try {
        const c = localStorage.getItem('erin-brand');
        storedBrand = c;
        if (c) document.documentElement.style.setProperty('--erin-brand', c);
    } catch (e) {}

    const railBtn = (label) =>
    [...document.querySelectorAll('#LeftRail button[aria-label]')].find(button => button.getAttribute('aria-label') === label);

    let tip;
    let tipAnchor = null;
    let tipHoldUntil = 0;
    let tipHideTimer = 0;
    function showTip(el, holdFor = 0) {
        if (!el?.dataset.tip) return;
        clearTimeout(tipHideTimer);
        tipAnchor = el;
        if (holdFor > 0) tipHoldUntil = Math.max(tipHoldUntil, performance.now() + holdFor);
        if (!tip) { tip = document.createElement('div'); tip.id = 'erin-tip'; document.body.appendChild(tip); }
        tip.textContent = el.dataset.tip;
        tip.classList.remove('show');
        const r = el.getBoundingClientRect();
        if (!r.width || !r.height) return;
        const panel = document.getElementById('erin-apps-panel');
        const panelRect = el.closest('#erin-apps-panel') && panel ? panel.getBoundingClientRect() : null;
        const bottom = panelRect?.bottom || r.bottom;
        const cx = r.left + r.width / 2;
        tip.style.left = '0px'; tip.style.top = '0px'; tip.style.visibility = 'hidden';
        tip.classList.add('show');
        const w = tip.offsetWidth;
        const left = Math.max(8, Math.min(cx - w / 2, innerWidth - w - 8));
        tip.style.left = left + 'px'; tip.style.top = bottom + 8 + 'px';
        tip.style.setProperty('--erin-arrow-x', Math.max(10, Math.min(cx - left, w - 10)) + 'px');
        tip.style.visibility = '';
    }
    function hideTip(force = false) {
        clearTimeout(tipHideTimer);
        if (!force) {
            const held = performance.now() < tipHoldUntil;
            const anchorHovered = !!tipAnchor?.isConnected && tipAnchor.matches(':hover');
            const menuBridge = tipAnchor?.id === 'erin-apps-btn' && document.getElementById('erin-apps-panel')?.classList.contains('open');
            if (held || anchorHovered || menuBridge) {
                const delay = held ? Math.max(40, tipHoldUntil - performance.now() + 20) : 120;
                tipHideTimer = setTimeout(() => hideTip(false), delay);
                return;
            }
        }
        tip?.classList.remove('show'); tipAnchor = null; tipHoldUntil = 0;
    }
    function setOpen(open) {
        const btn = document.getElementById('erin-apps-btn');
        const panel = document.getElementById('erin-apps-panel');
        if (!btn || !panel) return;
        // Compact search keeps the existing button available as Back to inbox, but suppresses its
        // hover apps flyout so the panel cannot cover or distract from the active search interface.
        if (open && document.documentElement.dataset.erinSearch === 'open') open = false;
        if (open) {
            const r = btn.getBoundingClientRect();
            panel.style.left = r.right + 4 + 'px';
            panel.style.top = r.top + (r.height - 40) / 2 + 'px';
            showTip(btn);
        }
        if (panel.classList.contains('open') !== open) panel.classList.toggle('open', open);
        panel.inert = !open;
        panel.setAttribute('aria-hidden', String(!open));
        if (btn.getAttribute('aria-expanded') !== String(open)) btn.setAttribute('aria-expanded', String(open));
        if (!open && !btn.matches(':hover') && !panel.matches(':hover')) hideTip();
    }

    // Subject bar (narrow layout): leftmost control = close; last two right of the subject = prev / next.
    // Controls are found from their icons, so disabled or non-<button> variants are still caught.
    function tagBar() {
        const root = document.querySelector('#ConversationReadingPaneContainer');
        if (!root) return;

        const wanted = new Map();
        for (const control of root.querySelectorAll('button, [role="button"]')) {
            const label = [
                control.getAttribute('aria-label'),
                control.getAttribute('title')
            ].filter(Boolean).join(' ').toLowerCase();

            if (/open the previous item|previous item/.test(label)) wanted.set(control, 'prev');
            else if (/open the next item|next item/.test(label)) wanted.set(control, 'next');
            else if (/^close(?:\s|$)/.test(label)) wanted.set(control, 'close');
        }

        for (const el of root.querySelectorAll('[data-erin]')) {
            if (!wanted.has(el)) el.removeAttribute('data-erin');
        }
        for (const [el, role] of wanted) el.setAttribute('data-erin', role);
    }

    function syncToast() {
        const root = document.documentElement;
        const set = (k, v) => { if (root.style.getPropertyValue(k) !== v) root.style.setProperty(k, v); };
        const card = document.querySelector('[data-app-section="MessageList"] [role="complementary"]');
        const r = card?.getBoundingClientRect();
        const pad = 20; // px of inset from the list's left/right edges
        if (r && r.width > 160) {
            set('--erin-toast-left', Math.round(r.left + pad) + 'px');
            set('--erin-toast-width', Math.round(r.width - 2 * pad) + 'px');
        } else {
            root.style.removeProperty('--erin-toast-left');
            root.style.removeProperty('--erin-toast-width');
        }
    }

    // Outlook supplies its own Fluent tooltips for these right-hand header buttons. Older versions of
    // this script also added data-tip, which produced two labels. Remove only those legacy tags; the
    // custom tooltips used by the left-hand apps menu are intentionally left alone.
    const NATIVE_HEADER_TIP_IDS = [
        'owaTimePanelBtn_container',
        'owaActivityFeedButton_container',
        'owaSettingsBtn_container'
    ];
    function tagHeaderTips() {
        for (const id of NATIVE_HEADER_TIP_IDS) {
            document.getElementById(id)?.removeAttribute('data-tip');
        }
    }

    function syncHamburger() {
        const btn = document.querySelector('button[aria-label$="navigation pane"]');
        const r = btn?.getBoundingClientRect();
        if (!r || !r.width) return;
        const root = document.documentElement;
        const set = (k, v) => { if (root.style.getPropertyValue(k) !== v) root.style.setProperty(k, v); };
        set('--erin-ham-left', Math.round(r.left) + 'px');
        set('--erin-ham-top', Math.round(r.top) + 'px');
        set('--erin-ham-w', Math.round(r.width) + 'px');
        set('--erin-ham-h', Math.round(r.height) + 'px');
    }

    function syncSearchCenter() {
        const root = document.documentElement;
        const box = document.getElementById('searchBoxColumnContainerId');
        const slot = box?.querySelector(':scope > button[role="menuitem"]') ? box.parentElement : null;
        const bar = document.getElementById('OwaTitleBar');
        if (!slot || !bar) { root.style.removeProperty('--erin-search-shift'); return; }
        const sr = slot.getBoundingClientRect(), br = bar.getBoundingClientRect();
        const slack = Math.max(0, (sr.width - Math.min(380, sr.width)) / 2);
        const want = br.left + br.width / 2 - (sr.left + sr.width / 2);
        const v = Math.round(Math.max(-slack, Math.min(slack, want))) + 'px';
        if (root.style.getPropertyValue('--erin-search-shift') !== v) root.style.setProperty('--erin-search-shift', v);
    }

    // Stack separate toasters so a newer toast never covers an older one (e.g. "Message sent" + "Archived").
    // Each toaster is fixed at the same spot; give each a --erin-toast-bottom offset, newest at the bottom.
    const toastSeq = new WeakMap();
    let toastN = 0, stackTimer;
    function stackToasts() {
        const live = [...document.querySelectorAll('.fui-Toaster[data-toaster-position="bottom-start"]')]
        .filter((t) => t.querySelector('.fui-ToastContainer'));
        const key = new Map();
        for (const t of live) {
            let k = 0;
            for (const c of t.querySelectorAll('.fui-ToastContainer')) {
                if (!toastSeq.has(c)) toastSeq.set(c, ++toastN);
                k = Math.max(k, toastSeq.get(c));
            }
            key.set(t, k);
        }
        live.sort((a, b) => key.get(b) - key.get(a)); // newest first = lowest on screen
        let offset = 16;
        for (const t of live) {
            const v = offset + 'px';
            if (t.style.getPropertyValue('--erin-toast-bottom') !== v) t.style.setProperty('--erin-toast-bottom', v);
            offset += Math.round(t.getBoundingClientRect().height) + 8;
        }
        // Heights change while toasts animate in; settle once more.
        if (live.length > 1) { clearTimeout(stackTimer); stackTimer = setTimeout(stackToasts, 320); }
    }

    // Remember where the collapsed search pill was, for the opening reveal.
    document.addEventListener('click', (e) => {
        const pill = e.target.closest?.('#searchBoxColumnContainerId > button[role="menuitem"]');
        if (!pill) return;
        const r = pill.getBoundingClientRect(), root = document.documentElement;
        root.style.setProperty('--erin-sx-il', Math.max(0, Math.round(r.left - 52)) + 'px');
        const av = document.getElementById('owa-me-control-container');
        const endM = av ? innerWidth - av.getBoundingClientRect().left : 48;
        root.style.setProperty('--erin-sx-ir', Math.max(0, Math.round(innerWidth - r.right - endM)) + 'px');
        root.dataset.erinSx = '1';
        setTimeout(() => { delete root.dataset.erinSx; }, 700);
    }, true);

    // Esc after navigating to Mail: drops focus/selection from the open email. Blur first so the key
    // is not attributed to our own button.
    function pressEscape() {
        // Clicking Mail moves focus to the open email asynchronously. Wait for focus to land inside the app
        // (not on our buttons), then send Esc *to that element* so the reading pane's handlers see it.
        let tries = 0;
        const go = () => {
            const el = document.activeElement;
            const inApp = !!el && el !== document.body && !!el.closest('#mainApp') &&
                  !el.closest('#erin-apps-btn, #erin-apps-panel');
            if (inApp || ++tries > 8) {
                const target = inApp ? el : document.body;
                for (const type of ['keydown', 'keyup']) {
                    target.dispatchEvent(new KeyboardEvent(type, {
                        key: 'Escape', code: 'Escape', keyCode: 27, which: 27,
                        bubbles: true, cancelable: true, composed: true }));
                }
            } else setTimeout(go, 80);
        };
        setTimeout(go, 100);
    }

    let appsBtnHomeMarker = null;

    // Search open at narrow width: keep a way back to the inbox (the apps button gets hidden/covered).
    function syncSearchOpen() {
        const root = document.documentElement;
        const box = document.getElementById('searchBoxColumnContainerId');
        const appsBtn = document.getElementById('erin-apps-btn');
        // Compact (small-width) search expands to a fixed overlay over the whole header, which covers the
        // apps button. Detect that by the wrapper being position:fixed (not by aria-expanded, which is
        // false when the bar is open but unfocused). Wide-layout search is not fixed.
        const wrap = box?.parentElement;
        const open = !!wrap && getComputedStyle(wrap).position === 'fixed';

        // Move the real button out of Outlook's hidden left-charms container while compact search is open.
        // A comment marker preserves its exact original insertion point for restoration afterward.
        if (open && appsBtn && appsBtn.parentElement !== document.body) {
            // Capture the button before removing it from Outlook's header layout. This preserves the exact
            // normal x/y alignment and lets the search overlay begin four pixels after its right edge.
            const r = appsBtn.getBoundingClientRect();
            if (r.width && r.height) {
                root.style.setProperty('--erin-apps-left', Math.round(r.left) + 'px');
                root.style.setProperty('--erin-apps-top', Math.round(r.top) + 'px');
                root.style.setProperty('--erin-search-left', Math.round(r.right + 4) + 'px');
            }
            appsBtnHomeMarker = document.createComment('erin-apps-btn-home');
            appsBtn.parentNode.insertBefore(appsBtnHomeMarker, appsBtn);
            document.body.appendChild(appsBtn);
        } else if (!open && appsBtn && appsBtnHomeMarker?.isConnected) {
            appsBtnHomeMarker.parentNode.insertBefore(appsBtn, appsBtnHomeMarker);
            appsBtnHomeMarker.remove();
            appsBtnHomeMarker = null;
            root.style.removeProperty('--erin-apps-left');
            root.style.removeProperty('--erin-apps-top');
            root.style.removeProperty('--erin-search-left');
        }
        // Segments of the expanded search (scope dropdown, field, glass button), left to right.
        // Expanded = the "Exit search" arrow exists (aria-expanded stays "false" in some states).
        const row = box?.querySelector('button[aria-label="Exit search"]') ? box.querySelector('[role="search"]') : null;
        const nextSegs = new Map();
        if (row) {
            const scope = document.getElementById('searchScopeButtonId'); // "All folders" lives in the left charms, not the field
            const segs = [...(scope ? [scope] : []), ...[...wrap.children].filter((c) => c !== box), ...[...box.children].filter((c) => c !== row), ...row.children]
            .map((el) => ({ el, r: el.getBoundingClientRect() }))
            .filter((x) => x.r.width > 4 && x.r.height > 4)
            .sort((x, y) => x.r.left - y.r.left);
            const firstNonScope = segs.find((x) => x.el !== scope);
            const visibleSegs = segs.filter((x) => x.el !== scope || (firstNonScope && Math.abs(firstNonScope.r.left - x.r.right) < 16));
            if (visibleSegs.length >= 2) visibleSegs.forEach((x, i) => nextSegs.set(x.el, i === 0 ? 'first' : i === visibleSegs.length - 1 ? 'last' : 'mid'));
        }
        for (const old of document.querySelectorAll('#OwaTitleBar [data-erin-seg]')) {
            const want = nextSegs.get(old);
            if (!want) old.removeAttribute('data-erin-seg');
        }
        for (const [el, v] of nextSegs) {
            if (el.getAttribute('data-erin-seg') !== v) el.setAttribute('data-erin-seg', v);
        }
        if (wrap) {
            if (open && !wrap.hasAttribute('data-erin-compact')) wrap.setAttribute('data-erin-compact', '');
            if (!open && wrap.hasAttribute('data-erin-compact')) wrap.removeAttribute('data-erin-compact');
        }
        if (open && root.dataset.erinSearch !== 'open') {
            root.dataset.erinSearch = 'open';
            setOpen(false);
        }
        if (!open && root.dataset.erinSearch) delete root.dataset.erinSearch;
    }

    // Drag-and-drop onto a compose window. Outlook creates this overlay only after a file drag starts. Its
    // row is already position:absolute with top/bottom set; changing that row to position:relative collapses
    // its hit area to zero height. Preserve Outlook's row geometry, hide the OneDrive column, and stretch the
    // native Attach files column across the existing overlay.
    GM_addStyle(`
        [data-erin-drop-row] { isolation: isolate !important; }
        html[data-erin-file-drag] [data-erin-drop-row] {
            z-index: 2147483645 !important; visibility: visible !important; opacity: 1 !important;
        }
        [data-erin-drop-row] > [data-erin-hide] { display: none !important; pointer-events: none !important; }
        [data-erin-drop-row] > [data-erin-attach] {
            position: absolute !important; inset: 0 !important; z-index: 2147483646 !important;
            width: auto !important; max-width: none !important; height: auto !important;
            flex: none !important; pointer-events: auto !important;
        }
        html[data-erin-file-drag] [data-erin-drop-row] > [data-erin-attach] {
            visibility: visible !important; opacity: 1 !important;
        }
    `);

    function isFileDrag(event) {
        const dt = event?.dataTransfer;
        if (!dt) return false;
        return Array.from(dt.types || []).some(type => String(type).toLowerCase() === 'files') ||
            Array.from(dt.items || []).some(item => item.kind === 'file');
    }

    function hideOneDriveDrop(root = document) {
        const scope = root?.querySelectorAll ? root : document;
        const containingDock = scope.closest?.('[id^="docking_DockingTriggerPart"]');
        const docks = scope.matches?.('[id^="docking_DockingTriggerPart"]') ? [scope]
            : containingDock ? [containingDock]
            : scope.querySelectorAll('[id^="docking_DockingTriggerPart"]');

        for (const dock of docks) {
            for (const row of dock.children) {
                if (row.tagName !== 'SPAN') continue;
                const columns = Array.from(row.children);
                const oneDrive = columns.find(col => /onedrive/i.test(col.textContent || ''));
                if (!oneDrive) continue;

                const attach = columns.find(col => col !== oneDrive && /attach|email/i.test(col.textContent || ''))
                    || columns.find(col => col !== oneDrive);
                if (!attach) continue;

                row.setAttribute('data-erin-drop-row', '');
                oneDrive.setAttribute('data-erin-hide', '');
                oneDrive.setAttribute('aria-hidden', 'true');
                attach.setAttribute('data-erin-attach', '');
            }
        }
    }

    // Keep repairing for a short window. Real drags emit dragover repeatedly, extending the window until
    // drop. This catches overlays inserted between events without leaving a permanent animation loop.
    let dropRepairFrame = 0;
    let dropRepairUntil = 0;
    function keepFileDropOverlayRepaired(durationMs = 500) {
        dropRepairUntil = Math.max(dropRepairUntil, performance.now() + durationMs);
        document.documentElement.dataset.erinFileDrag = '1';
        if (dropRepairFrame) return;
        const tick = () => {
            hideOneDriveDrop();
            if (performance.now() < dropRepairUntil) {
                dropRepairFrame = requestAnimationFrame(tick);
            } else {
                dropRepairFrame = 0;
                delete document.documentElement.dataset.erinFileDrag;
            }
        };
        tick();
    }

    // Chromium may withhold Files/dataTransfer.items while a native Finder drag is only hovering.
    // A drag already inside a compose dock is safe to repair: non-file drags never create the
    // Attach/OneDrive overlay, so hideOneDriveDrop simply has nothing to change.
    function isComposeDrag(event) {
        const target = event.target;
        return isFileDrag(event) ||
            (target instanceof Element && Boolean(target.closest('[id^="docking_DockingTriggerPart"]')));
    }

    document.addEventListener('dragenter', event => {
        if (isComposeDrag(event)) keepFileDropOverlayRepaired();
    }, true);
    document.addEventListener('dragover', event => {
        if (isComposeDrag(event)) keepFileDropOverlayRepaired();
    }, true);

    document.addEventListener('drop', event => {
        if (!isFileDrag(event)) return;
        hideOneDriveDrop(event.target?.closest?.('[id^="docking_DockingTriggerPart"]') || document);
        dropRepairUntil = 0;
        delete document.documentElement.dataset.erinFileDrag;
    }, true);

    function positionedRoot(menu) {
        for (let e = menu; e && e !== document.body; e = e.parentElement) {
            const p = getComputedStyle(e).position;
            if (p === 'absolute' || p === 'fixed') return e;
        }
        return menu;
    }
    // Some "Move to" menus mount and stay at the viewport's top-left corner. Re-measure a few times (popup
    // libraries place menus a beat after mounting); if one is still parked at (0,0) while the control that
    // opened it is elsewhere, shift it next to that control. The shift is recomputed each pass from the
    // un-shifted position, so it never stacks on top of a correct placement.
    function checkMenu(menu) {
        if (!menu.isConnected) return;
        const root = positionedRoot(menu);
        root.style.removeProperty('translate');
        const r = root.getBoundingClientRect();
        const parked = r.width > 0 && r.left < 12 && r.top < 12;
        const lc = lastClick && Date.now() - lastClick.t < 4000 ? lastClick : null;
        if (parked && lc && lc.r.width > 0 && (lc.r.left > 60 || lc.r.top > 60)) {
            const x = Math.min(Math.max(8, lc.r.left), innerWidth - r.width - 8);
            let y = lc.r.bottom + 4;
            if (y + r.height > innerHeight - 8) y = Math.max(8, lc.r.top - r.height - 4);
            root.style.setProperty('translate', Math.round(x - r.left) + 'px ' + Math.round(y - r.top) + 'px', 'important');
        }
    }

    function within(root, selector) {
        const found = [];
        if (root?.matches?.(selector)) found.push(root);
        if (root?.querySelectorAll) found.push(...root.querySelectorAll(selector));
        return found;
    }

    function checkAddedMenus(root) {
        for (const mn of within(root, '[role="menu"]')) {
            for (const d of [250, 700, 1500]) setTimeout(() => checkMenu(mn), d);
        }
    }

    function checkAddedDropTargets(root) {
        const dock = root.matches?.('[id^="docking_DockingTriggerPart"]') ? root
            : root.closest?.('[id^="docking_DockingTriggerPart"]')
            || root.querySelector?.('[id^="docking_DockingTriggerPart"]');
        if (dock) hideOneDriveDrop(dock);
    }

    // Hover lock: while any floating menu (Categories, Move to, Filter, ...) is open, other JS-driven hover
    // behaviour must not fire, because opening a person card / tooltip elsewhere takes focus and light-dismisses
    // the menu. Hover events are swallowed in the capture phase unless they are inside a menu (or our own UI).
    // CSS :hover is unaffected, so highlights still work. data-erin-menu also hides an already-open person card.
    let menuAt = 0, menuVal = false;
    function anyMenuOpen() {
        const now = performance.now();
        if (now - menuAt > 100) {
            menuAt = now;
            const v = [...document.querySelectorAll('[role="menu"]')].some((m) => { const r = m.getBoundingClientRect(); return r.width > 0 && r.height > 0; });
            if (v !== menuVal) {
                menuVal = v;
                const root = document.documentElement;
                if (v) root.dataset.erinMenu = '1'; else delete root.dataset.erinMenu;
            }
        }
        return menuVal;
    }
    for (const type of ['mouseover', 'mouseout', 'mousemove']) {
        document.addEventListener(type, (e) => {
            if (!anyMenuOpen()) return;
            if (e.target.closest?.('[role="menu"], #erin-apps-panel, #erin-apps-btn')) return;
            e.stopPropagation();
        }, true);
    }

    // ---- Search menus (Categories, Move to) -------------------------------------------------------------
    // Both are Fluent menus (.fui-MenuPopover > [role=menu] > SearchBox). Two protections:
    //  1) Focus keeper: if the field loses focus while the menu is still open and you did not mean it to
    //     (no click, Tab, Enter or Esc just before), focus goes straight back, so typing keeps landing in the
    //     field instead of firing Outlook shortcuts (compose, archive, ...).
    //  2) Keystroke guard: if the menu vanishes while you were typing and nothing you did explains it
    //     (no Enter/Esc/click/Tab), the next ~2.5s of plain keystrokes are swallowed instead of becoming
    //     shortcuts. A burst of typing keeps extending the guard; clicking anywhere ends it.
    let intentUntil = 0, guardUntil = 0, lastTab = 0, watch = null;
    const typingKey = (e) => e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
    const editable = (t) => !!t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName));
    window.addEventListener('keydown', (e) => {
        const now = Date.now();
        if (now < guardUntil && typingKey(e) && !editable(e.target)) {
            e.preventDefault(); e.stopImmediatePropagation();
            guardUntil = Math.max(guardUntil, now + 900);
            return;
        }
        if (e.key === 'Tab') lastTab = now;
        if (e.key === 'Enter' || e.key === 'Escape') intentUntil = now + 700;
        const t = e.target;
        if (t instanceof HTMLInputElement && t.closest('[role="menu"]')) watch = { menu: t.closest('[role="menu"]'), lastTyped: now };
    }, true);
    document.addEventListener('mousedown', () => { intentUntil = Date.now() + 700; guardUntil = 0; }, true);
    document.addEventListener('focusout', (e) => {
        const inp = e.target;
        if (!(inp instanceof HTMLInputElement) || !inp.closest('[role="menu"]')) return;
        const menu = inp.closest('[role="menu"]');
        setTimeout(() => {
            if (!menu.isConnected || !inp.isConnected) return;
            const a = document.activeElement;
            if (a && menu.contains(a)) return;                                   // arrow keys / clicking an item
            if (Date.now() < intentUntil || Date.now() - lastTab < 700) return;  // you meant to leave
            inp.focus({ preventScroll: true });
        }, 0);
    }, true);
    function checkSearchMenu() {
        if (!watch || watch.menu.isConnected) return;
        const now = Date.now();
        if (now - watch.lastTyped < 2000 && !(now < intentUntil || now - lastTab < 700)) {
            guardUntil = now + 2500;
        }
        watch = null;
    }

    // Typing flag for the person hover card (CSS above does the hiding).
    let typingT;
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey || e.metaKey || e.altKey) return;
        const t = e.target;
        if (!t || !(t.isContentEditable || /^(INPUT|TEXTAREA)$/.test(t.tagName))) return;
        if (e.key.length !== 1 && e.key !== 'Backspace' && e.key !== 'Enter') return;
        const root = document.documentElement;
        root.dataset.erinTyping = '1';
        clearTimeout(typingT);
        typingT = setTimeout(() => { delete root.dataset.erinTyping; }, 2000);
    }, true);

    // Reuse Outlook's native sort icon when the responsive layout switches the control to text.
    // CSS supplies an immediate fallback until Outlook has rendered the native icon at least once.
    let nativeSortIconTemplate = null;
    function syncSortIcon() {
        const button = document.getElementById('mailListSortMenu');
        if (!button) return;
        const nativeIcon = button.querySelector('.fui-Button__icon:not([data-erin-sort-icon])') ||
              button.querySelector('svg:not([data-erin-sort-icon] svg)')?.closest('.fui-Button__icon, span, i');
        if (nativeIcon && !nativeIcon.closest('[data-erin-sort-icon]')) {
            nativeSortIconTemplate = nativeIcon.cloneNode(true);
            button.removeAttribute('data-erin-sort-text');
            for (const old of button.querySelectorAll(':scope > [data-erin-sort-icon]')) old.remove();
            return;
        }
        const visibleText = (button.textContent || '').replace(/\s+/g, ' ').trim();
        const textVariant = /(?:by\s+date|sorted\s*:)/i.test(visibleText);
        if (!textVariant) {
            button.removeAttribute('data-erin-sort-text');
            for (const old of button.querySelectorAll(':scope > [data-erin-sort-icon]')) old.remove();
            return;
        }
        button.setAttribute('data-erin-sort-text', '');
        let clone = button.querySelector(':scope > [data-erin-sort-icon]');
        if (!clone && nativeSortIconTemplate) {
            clone = nativeSortIconTemplate.cloneNode(true);
            clone.setAttribute('data-erin-sort-icon', '');
            clone.removeAttribute('aria-label'); clone.setAttribute('aria-hidden', 'true');
            button.appendChild(clone);
        }
    }

    // Tag only the direct #mainApp child that actually contains the open right-side DrawerPane.
    function syncDrawerPane() {
        const root = document.documentElement;
        const main = document.getElementById('mainApp');
        if (!main) { delete root.dataset.erinDrawer; root.style.removeProperty('--erin-drawer-width'); return; }
        const drawer = main.querySelector('[data-app-section="DrawerPane"]');
        const oldHost = main.querySelector(':scope > [data-erin-drawer-host]');
        if (!drawer) {
            oldHost?.removeAttribute('data-erin-drawer-host');
            delete root.dataset.erinDrawer; root.style.removeProperty('--erin-drawer-width'); return;
        }
        let host = drawer;
        while (host.parentElement && host.parentElement !== main) host = host.parentElement;
        if (host.parentElement !== main) {
            oldHost?.removeAttribute('data-erin-drawer-host');
            delete root.dataset.erinDrawer; root.style.removeProperty('--erin-drawer-width'); return;
        }
        if (oldHost && oldHost !== host) oldHost.removeAttribute('data-erin-drawer-host');
        host.setAttribute('data-erin-drawer-host', ''); root.dataset.erinDrawer = 'open';
        const drawerRect = drawer.getBoundingClientRect(), hostRect = host.getBoundingClientRect();
        let width = drawerRect.width || hostRect.width;
        if (!width || width < 180) {
            const token = getComputedStyle(root).getPropertyValue('--flexPaneWidth').trim();
            width = parseFloat(token) || 320;
        }
        width = Math.max(240, Math.min(Math.round(width), Math.max(240, innerWidth - 320)));
        const value = width + 'px';
        if (root.style.getPropertyValue('--erin-drawer-width') !== value) root.style.setProperty('--erin-drawer-width', value);
    }

    function tagLayout() {
        if (!document.querySelector('[data-app-section="MessageList"]')) return;
        const pane = document.querySelector('[data-app-section="MailReadCompose"]');
        const v = (pane ? pane.getBoundingClientRect().width : 0) < 2 ? 'list' : 'multi';
        const root = document.documentElement;
        if (root.dataset.erinLayout !== v) root.dataset.erinLayout = v;
    }

    let appsEvents = null;
    function build() {
        syncBrand();
        tagBar();
        tagLayout();
        syncSortIcon();
        syncDrawerPane();
        syncToast();
        syncHamburger();
        anyMenuOpen();
        checkSearchMenu();
        tagHeaderTips();
        syncSearchCenter();
        syncSearchOpen();
        stackToasts();
        hideSenderCategories();
        const waffle = document.getElementById('owaAppLauncherBtn_container');
        if (!waffle || !document.getElementById('LeftRail')) return;

        if (!document.getElementById('erin-apps-btn') || !document.getElementById('erin-apps-panel')) {
            appsEvents?.abort();
            appsEvents = new AbortController();
            const eventOptions = { signal: appsEvents.signal };
            document.getElementById('erin-apps-btn')?.remove();
            document.getElementById('erin-apps-panel')?.remove();
            const btn = document.createElement('button');
            btn.id = 'erin-apps-btn';
            btn.type = 'button';
            btn.dataset.tip = 'Back to inbox';
            btn.setAttribute('aria-label', 'Back to inbox');
            btn.setAttribute('aria-expanded', 'false');
            btn.setAttribute('aria-controls', 'erin-apps-panel');
            btn.setAttribute('aria-haspopup', 'true');
            // Two icons stacked in one grid cell (inline styles, so this never depends on a selector matching):
            // the boxes, and a mail icon that fades in over them on click.
            btn.innerHTML = ICONS.grid + ICONS.Mail;
            const [gridSvg, mailSvg] = btn.querySelectorAll('svg');
            gridSvg.classList.add('grid'); mailSvg.classList.add('mail');
            for (const sv of [gridSvg, mailSvg]) sv.style.cssText += ';grid-area:1/1;transition:transform .25s ease,opacity .15s ease';
            mailSvg.style.opacity = '0'; mailSvg.style.pointerEvents = 'none';
            btn._flash = () => {
                gridSvg.style.opacity = '0'; mailSvg.style.opacity = '1';
                mailSvg.animate([{ transform: 'scale(1)' }, { transform: 'scale(1.18)' }, { transform: 'scale(1)' }], { duration: 650, easing: 'ease-in-out' });
                clearTimeout(btn._flashT);
                btn._flashT = setTimeout(() => { gridSvg.style.opacity = '1'; mailSvg.style.opacity = '0'; }, 900);
            };
            // Click = refresh the mail list (same as clicking Mail in the old left rail); hover opens the flyout.
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                railBtn('Mail')?.click();
                pressEscape();
                btn._flash();
                requestAnimationFrame(() => showTip(btn, 950));
            });
            waffle.insertAdjacentElement('afterend', btn);

            const panel = document.createElement('div');
            panel.id = 'erin-apps-panel';
            panel.inert = true; panel.setAttribute('aria-hidden', 'true');
            panel.setAttribute('role', 'group'); panel.setAttribute('aria-label', 'Outlook apps');
            for (const app of APPS) {
                const { id, label } = app;
                const b = document.createElement('button');
                const isSharedMailbox = !!app.shared;
                const appUrl = ({
                    Mail: location.origin + '/mail/',
                    Calendar: location.origin + '/calendar/',
                    'To Do': 'https://to-do.office.com/tasks/'
                })[id] || (isSharedMailbox ? `${location.origin}/mail/${sharedMailbox.address}/` : location.origin);
                b.type = 'button';
                b.dataset.for = id;
                b.dataset.tip = label;
                b.setAttribute('aria-label', label);
                b.innerHTML = isSharedMailbox ? ICONS.Mail : ICONS[id];

                const openSeparate = () => window.open(appUrl, '_blank', 'noopener');
                b.addEventListener('auxclick', (event) => {
                    if (event.button !== 1) return;
                    event.preventDefault();
                    openSeparate();
                    setOpen(false);
                });
                b.addEventListener('click', (event) => {
                    if (event.metaKey || event.ctrlKey) {
                        event.preventDefault();
                        openSeparate();
                        setOpen(false);
                        return;
                    }
                    if (isSharedMailbox) {
                        void openSharedMailboxInbox();
                    } else {
                        railBtn(id)?.click();
                    }
                    setOpen(false);
                });
                panel.appendChild(b);
            }
            document.body.appendChild(panel);
            btn.addEventListener('keydown', event => {
                if (!['ArrowDown', 'ArrowRight', ' '].includes(event.key)) return;
                event.preventDefault(); event.stopPropagation(); setOpen(true);
                if (!panel.inert) panel.querySelector('button')?.focus();
            });
            panel.addEventListener('keydown', event => {
                const buttons = [...panel.querySelectorAll('button')], index = buttons.indexOf(document.activeElement);
                let next = index;
                if (event.key === 'ArrowRight' || event.key === 'ArrowDown') next = (index + 1) % buttons.length;
                else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') next = (index - 1 + buttons.length) % buttons.length;
                else if (event.key === 'Home') next = 0;
                else if (event.key === 'End') next = buttons.length - 1;
                else if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); setOpen(false); btn.focus(); return; }
                else return;
                event.preventDefault(); event.stopPropagation(); buttons[next]?.focus();
            });
            panel.addEventListener('focusout', () => requestAnimationFrame(() => {
                if (!panel.contains(document.activeElement) && document.activeElement !== btn && !panel.matches(':hover')) setOpen(false);
            }));

            for (const host of [btn, panel]) {
                host.addEventListener('mouseover', (e) => {
                    const t = e.target.closest('[data-tip]');
                    if (t) showTip(t);
                });
                host.addEventListener(host === btn ? 'mouseleave' : 'mouseout', () => hideTip());
                host.addEventListener('focusin', (e) => { const t = e.target.closest('[data-tip]'); if (t) showTip(t); });
                host.addEventListener('focusout', () => hideTip());
            }

            // Hover opens (short delay avoids flicker when passing over). It stays open while the pointer is on
            // the icon or the pill, and closes as soon as it moves off both.
            let openT;
            btn.addEventListener('pointerenter', () => {
                clearTimeout(openT);
                showTip(btn);
                if (document.documentElement.dataset.erinSearch !== 'open') {
                    openT = setTimeout(() => { setOpen(true); showTip(btn); }, 60);
                }
            });
            btn.addEventListener('pointerleave', () => { clearTimeout(openT); hideTip(); });
            document.addEventListener('pointermove', (e) => {
                if (!panel.classList.contains('open') || panel.contains(document.activeElement)) return;
                // Extra room on the right of the pill so overshooting toward the search box does not close it.
                const near = (r, right = 6) => e.clientX >= r.left - 6 && e.clientX <= r.right + right &&
                      e.clientY >= r.top - 6 && e.clientY <= r.bottom + 6;
                if (!near(btn.getBoundingClientRect()) && !near(panel.getBoundingClientRect(), 56)) setOpen(false);
            }, { passive: true, signal: appsEvents.signal });
            document.documentElement.addEventListener('mouseleave', () => { if (!panel.contains(document.activeElement)) setOpen(false); }, eventOptions);

            document.addEventListener('click', (e) => {
                if (!panel.contains(e.target)) setOpen(false);
            }, eventOptions);
            document.addEventListener('keydown', (e) => { if (e.key === 'Escape') setOpen(false); }, eventOptions);
            window.addEventListener('resize', () => setOpen(false), eventOptions);
        }

        // Mirror active app
        for (const app of APPS) {
            const mine = [...document.querySelectorAll('#erin-apps-panel [data-for]')].find(button => button.dataset.for === app.id);
            const v = app.shared ? 'false' : railBtn(app.id)?.getAttribute('aria-pressed') || 'false';
            if (mine && mine.getAttribute('aria-pressed') !== v) mine.setAttribute('aria-pressed', v);
        }
    }

    window.addEventListener('resize', () => requestAnimationFrame(() => {
        tagBar(); tagLayout(); syncSortIcon(); syncDrawerPane(); syncToast(); syncSearchCenter();
    }));

    for (const ev of ['focusin', 'focusout']) document.addEventListener(ev, () => requestAnimationFrame(syncSearchOpen), true);

    let queued = false;
    const relevantAttrTarget = (target) => target.nodeType === Node.ELEMENT_NODE && (
        target.id === 'mainApp' ||
        !!target.closest?.(
            '#OwaTitleBar, #LeftRail, #searchBoxColumnContainerId, #mailListSortMenu, ' +
            '#owaAppLauncherBtn_container, #owaHeaderButtonGroupRight_container, #ConversationReadingPaneContainer, ' +
            '.fui-Toaster, [role="menu"], [data-app-section="DrawerPane"], ' +
            '[data-app-section="MailReadCompose"], [data-app-section="Ribbon"]'
        )
    );

    const scheduleBuild = () => {
        if (queued) return;
        queued = true;
        requestAnimationFrame(() => {
            queued = false;
            build();
        });
    };

    const signInClicks = new WeakMap();
    function clickSessionExpiredSignIn(scope = document) {
        const groups = within(scope, '[role="group"]');
        for (const banner of groups) {
            if (!banner.textContent.includes('You need to sign in.')) continue;
            const btn = [...banner.querySelectorAll('button')].find(button => /sign in/i.test(button.getAttribute('aria-label') || button.textContent));
            if (!btn || Date.now() - (signInClicks.get(btn) || 0) < 5000) continue;
            signInClicks.set(btn, Date.now()); btn.click(); return true;
        }
        return false;
    }

    function hideSenderCategories() {
        // Avoid writing the same style attribute from inside an observer that watches styles.
        for (const label of document.querySelectorAll('span.t1Kw7')) {
            const host = label.closest('.BNwHL');
            if (host) host.toggleAttribute('data-erin-promotions', label.textContent?.trim() === 'Promotions');
        }
    }
    GM_addStyle('[data-erin-promotions] { display: none !important; }');

    const NEVER_DELETE_RETENTION_TEXT = 'Retention: Default Never Delete';

    function hideNeverDeleteRetention(root = document) {
        if (root.nodeType !== Node.ELEMENT_NODE && root !== document) return;

        const bodies = [];
        if (root.matches?.('.fui-MessageBarBody')) bodies.push(root);
        bodies.push(...root.querySelectorAll('.fui-MessageBarBody'));

        for (const body of bodies) {
            if (body.textContent?.trim() !== NEVER_DELETE_RETENTION_TEXT) continue;
            const group = body.closest('.fui-MessageBarGroup');
            if (group && !group.hasAttribute('data-erin-never-delete-retention')) group.setAttribute('data-erin-never-delete-retention', '');
        }
    }

    const ownedUI = node => node.nodeType === Node.ELEMENT_NODE && !!node.closest?.('#erin-inbox-widgets, #erin-apps-panel, #erin-apps-btn, #erin-tip');
    const shellSelector = '#OwaTitleBar, #LeftRail, #mainApp, #ConversationReadingPaneContainer, .fui-Toaster, [role="menu"], [data-app-section="DrawerPane"], [data-app-section="Ribbon"]';
    const mainObserver = new MutationObserver((mutations) => {
        let needsBuild = false;
        const signInScopes = new Set();
        for (const mutation of mutations) {
            if (ownedUI(mutation.target)) continue;
            if (mutation.type === 'childList') {
                const added = [...mutation.addedNodes].filter(node => node.nodeType === Node.ELEMENT_NODE && !ownedUI(node));
                const removed = [...mutation.removedNodes].filter(node => node.nodeType === Node.ELEMENT_NODE && !ownedUI(node));
                if (!added.length && !removed.length) continue;
                // A widget refresh must not rebuild the entire Outlook shell.
                const significant = node => !node.matches?.('#erin-inbox-widgets') &&
                    (node.matches?.(shellSelector) || node.querySelector?.(shellSelector));
                if (relevantAttrTarget(mutation.target) || added.some(significant) || removed.some(significant)) needsBuild = true;
                for (const node of added) {
                    checkAddedMenus(node); checkAddedDropTargets(node);
                    removeDeleteMenuItems(node); hideNeverDeleteRetention(node);
                    signInScopes.add(node.closest('[role="group"]') || node);
                }
            } else if (relevantAttrTarget(mutation.target)) needsBuild = true;
        }
        // Search each added scope once, rather than rescanning the whole document per added node.
        for (const scope of signInScopes) clickSessionExpiredSignIn(scope);
        if (needsBuild) scheduleBuild();
    });

    mainObserver.observe(document.documentElement, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: [
            'aria-pressed',
            'aria-expanded',
            'class',
            'style',
            'hidden'
        ],
    });

    // Process the DOM that already exists when the observer is attached.
    requestAnimationFrame(() => {
        build();
        clickSessionExpiredSignIn();
        hideSenderCategories();
        hideNeverDeleteRetention();
    });

    // Copilot's function to remove "Delete" menu items
    function removeDeleteMenuItems(root = document) {
        within(root, '.fui-MenuItem__content').forEach(label => {
            if (label.textContent.trim() === 'Delete') {
                const menuItem = label.closest('div[role="menuitem"]');
                if (menuItem) menuItem.remove();
            }
        });
    }
    // Initial scan; later menu additions are handled by the main observer above.
    removeDeleteMenuItems();
})();

