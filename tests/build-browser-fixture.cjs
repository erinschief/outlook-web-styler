// Generate a standalone browser regression page from the actual userscript sources.
// Open the printed file URL in Chromium. No Outlook account or network access is used.
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const widgets = fs.readFileSync(path.join(root, 'outlook-inbox-widgets.user.js'), 'utf8')
  .replace("'use strict';", "'use strict'; const location = window.testLocation;")
  .replace(/\}\)\(\);\s*$/, `window.widgets = {cache, get state(){return state}, get focus(){return focus}, get scanning(){return scanning}, get complete(){return scanComplete}, get root(){return root}, bucketFor, rowDate, parseClock, parseTimeRange, parseDuration, scrapeMessages, render, mount, scanAllMessages, syncAccount, flushSave, advanceFocus, startFocus, pruneState, observer, cancelScan(){scanGeneration++}, stopAutoScan(){clearTimeout(scanTimer)}, get pendingBlock(){return pendingBlock}, set pendingBlock(v){pendingBlock=v} };})();`);
const styler = fs.readFileSync(path.join(root, 'outlook-web-styler.user.js'), 'utf8')
  .replace("'use strict';", "'use strict'; const location = window.testLocation;")
  .replace('function build() {', 'function build() { window.buildCount++;')
  .replace(/\}\)\(\);\s*$/, 'window.styler = {build, setOpen, mainObserver};})();');
const tests = fs.readFileSync(path.join(__dirname, 'browser-tests.js'), 'utf8');
const prelude = `window.errors=[];addEventListener('error',e=>errors.push(e.message));addEventListener('unhandledrejection',e=>errors.push(String(e.reason)));window.buildCount=0;window.testLocation={host:'outlook.office.com',origin:'https://outlook.office.com',pathname:'/mail/inbox',assign(url){window.assignedURL=url}};window.stored=new Map();window.GM_getValue=(key,fallback)=>structuredClone(stored.has(key)?stored.get(key):fallback);window.GM_setValue=(key,value)=>stored.set(key,structuredClone(value));stored.set('sharedMailbox',{address:'team@example.com',label:'Mail \\"Ops\\"'});window.GM_registerMenuCommand=()=>{};window.GM_addStyle=css=>{const el=document.createElement('style');el.textContent=css;document.head.appendChild(el)};`;
const markup = `<!doctype html><meta charset="utf-8"><title>Outlook userscript regression results</title><style>body{font:14px system-ui;margin:20px}#results{white-space:pre-wrap;margin-top:20px}#fixture{width:800px}#OwaTitleBar{height:48px;display:flex}#LeftRail{display:none}[data-app-section=MessageList]{width:600px}[role=complementary]{width:600px}[data-testid=virtuoso-scroller]{height:240px;overflow:auto}.test-row{height:60px}.test-pass{color:#087535}.test-fail{color:#a40000}</style><h1>Outlook userscript regression results</h1><pre id="results">Running…</pre><div id="fixture"><div id="OwaTitleBar"><div id="owaAppLauncherBtn_container"></div><div id="owa-me-control-container" title="erin@example.com">Account</div></div><div id="LeftRail"><button aria-label="Mail">Mail</button><button aria-label="Calendar">Calendar</button><button aria-label="To Do">To Do</button></div><div data-app-section="Ribbon"><button label="New mail" aria-label="New mail" style="background:#157956;color:white">New mail</button></div><div id="ConversationReadingPaneContainer"><div class="BNwHL"><span class="t1Kw7">Promotions</span></div></div><div data-app-section="MessageList"><div role="complementary"><div data-testid="virtuoso-scroller"><div data-testid="virtuoso-item-list"></div></div></div></div></div>`;
const script = value => `<script>${value.replaceAll('</script', '<\\/script')}</script>`;
const output = process.argv[2] || '/tmp/outlook-browser-regression.html';
fs.writeFileSync(output, markup + script(prelude) + script(widgets) + script(styler) + script(tests));
console.log('file://' + path.resolve(output));
