// ==UserScript==
// @name         Outlook Inbox Widgets
// @namespace    https://github.com/erinschief/outlook-web-styler
// @version      0.2.0
// @description  Collapsible personal-inbox widgets for unread mail, today's schedule, focus sessions, and Self/Waiting categories.
// @author       Erin
// @icon         data:image/png;base64,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
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.live.com/*
// @match        https://outlook.cloud.microsoft/*
// @updateURL    https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-inbox-widgets.user.js
// @downloadURL  https://raw.githubusercontent.com/erinschief/outlook-web-styler/main/outlook-inbox-widgets.user.js
// @run-at       document-idle
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// ==/UserScript==

(() => {
  'use strict';

  const ROOT_ID = 'erin-inbox-widgets';
  const STORE = 'erinInboxWidgets.v1';
  const cache = new Map();
  const state = Object.assign({
    collapsed: false,
    tab: 'timeline',
    todo: 'self',
    todoExpanded: false,
    selfOrder: [],
    titleEdits: {},
    allDayExpanded: false,
    scheduleDate: '',
    schedule: [],
    localBlocks: []
  }, GM_getValue(STORE, {}));
  let root = null;
  let renderTimer = 0;
  let observer = null;
  let focusTick = 0;
  let scanning = false;
  let scanComplete = false;
  let scanProgress = 0;
  let scannedScroller = null;
  let pendingBlock = null;
  let flowCleanup = null;
  let focus = { part: 0, elapsed: 0, running: false, finished: false, last: 0 };
  const PARTS = [
    ['focus', 25, 'Focus 1 of 2'],
    ['break', 5, 'Rest'],
    ['focus', 25, 'Focus 2 of 2'],
    ['break', 20, 'Break']
  ];

  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[c]);
  const decodeEntities = value => value.replace(/&(amp|lt|gt|quot|#39|#x[0-9a-f]+|#\d+);/gi, (_, entity) => {
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', '#39': "'" };
    if (named[entity.toLowerCase()] !== undefined) return named[entity.toLowerCase()];
    const hex = entity[0] === '#' && entity[1]?.toLowerCase() === 'x';
    return String.fromCodePoint(parseInt(entity.slice(hex ? 2 : 1), hex ? 16 : 10));
  });

  // Outlook intercepts both innerHTML and DOMParser inside extension userscripts. Tokenize only our
  // controlled templates and construct the live tree exclusively through DOM creation methods.
  function setMarkup(host, html) {
    const fragment = document.createDocumentFragment();
    const stack = [fragment];
    const voidElements = new Set(['br', 'hr', 'img', 'input', 'meta', 'link']);
    const tokens = html.matchAll(/<\/(?:[a-z][\w-]*)\s*>|<[a-z][\w-]*(?:\s+[^>]*)?\/?>|[^<]+/gi);
    for (const token of tokens) {
      const source = token[0];
      if (source.startsWith('</')) {
        if (stack.length > 1) stack.pop();
        continue;
      }
      if (!source.startsWith('<')) {
        stack[stack.length - 1].appendChild(document.createTextNode(decodeEntities(source)));
        continue;
      }
      const opening = source.match(/^<([a-z][\w-]*)([\s\S]*?)\/?\s*>$/i);
      if (!opening) continue;
      const tag = opening[1].toLowerCase();
      const parent = stack[stack.length - 1];
      const svg = tag === 'svg' || parent.namespaceURI === 'http://www.w3.org/2000/svg';
      const element = svg ? document.createElementNS('http://www.w3.org/2000/svg', tag) : document.createElement(tag);
      for (const attribute of opening[2].matchAll(/([:@\w-]+)(?:\s*=\s*"([^"]*)")?/g)) {
        element.setAttribute(attribute[1], decodeEntities(attribute[2] ?? ''));
      }
      parent.appendChild(element);
      if (!voidElements.has(tag) && !/\/\s*>$/.test(source)) stack.push(element);
    }
    host.replaceChildren(fragment);
  }
  const save = () => GM_setValue(STORE, state);
  const todayKey = () => new Date().toLocaleDateString('en-CA');
  const personalInbox = () => /^\/mail(?:\/?$|\/inbox\/?$)/i.test(location.pathname);
  const searchActive = () => document.documentElement.dataset.erinSearch === 'open' ||
    !!document.querySelector('button[aria-label="Exit search"]') || /\/search(?:\/|$)/i.test(location.pathname);
  const widgetContext = () => personalInbox() && !searchActive();
  const visible = el => !!el && el.getClientRects().length > 0;
  const messageList = () => {
    const lists = [...document.querySelectorAll('[data-app-section="MessageList"]')];
    return lists.find(list => visible(list) && list.querySelector('[role="complementary"], [role="listbox"]')) ||
      lists.find(visible) || lists[0] || null;
  };

  GM_addStyle(`
    #${ROOT_ID}{--bg:var(--colorNeutralBackground1,#fff);--ink:var(--colorNeutralForeground1,#242424);--mute:var(--colorNeutralForeground3,#616161);--line:var(--colorNeutralStroke2,#e0e0e0);--brand:var(--colorBrandBackground,#0f6cbd);--red:#d13438;--orange:#e8820c;--purple:#8764b8;--blue:#4aa3f0;--green:#2e8b6e;--meet:#d6e8f8;--focus:#d5efe4;--admin:#f6e8c8;--rest:#e8e9ec;flex:0 0 auto;width:100%;box-sizing:border-box;padding:0 8px 8px;color:var(--ink);font:13px/1.35 "Segoe UI Variable","Segoe UI",sans-serif;container-type:inline-size;z-index:3}
    @media(prefers-color-scheme:dark){#${ROOT_ID}{--bg:var(--colorNeutralBackground1,#1f2226);--ink:var(--colorNeutralForeground1,#e8eaed);--mute:var(--colorNeutralForeground3,#9aa1ab);--line:var(--colorNeutralStroke2,#3a3d42);--meet:#1c3a57;--focus:#173d31;--admin:#47381a;--rest:#2b2e34}}
    #${ROOT_ID} *{box-sizing:border-box}#${ROOT_ID} button,#${ROOT_ID} input{font:inherit;color:inherit}
    #${ROOT_ID} button{cursor:pointer}#${ROOT_ID} button:focus-visible,#${ROOT_ID} [tabindex]:focus-visible{outline:2px solid var(--brand);outline-offset:1px}
    .eiw-shell{background:color-mix(in srgb,var(--bg) 88%,transparent);border:1px solid var(--line);border-radius:10px;box-shadow:0 3px 16px rgba(0,0,0,.10);overflow:hidden;backdrop-filter:blur(18px)}
    .eiw-top{height:38px;display:flex;align-items:center;gap:6px;padding:4px 6px 4px 10px}.eiw-title{font-weight:650;flex:1;white-space:nowrap}.eiw-summary{color:var(--mute);font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .eiw-iconbtn,.eiw-tab,.eiw-btn{border:1px solid transparent;background:transparent;border-radius:6px;min-height:28px}.eiw-iconbtn{width:30px;padding:0;font-size:17px}.eiw-iconbtn:hover,.eiw-btn:hover{background:color-mix(in srgb,var(--brand) 9%,transparent)}
    .eiw-tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:3px;padding:0 6px 6px}.eiw-tab{padding:4px 5px;color:var(--mute);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.eiw-tab[aria-selected="true"]{background:var(--brand);color:#fff}.eiw-panel{border-top:1px solid var(--line);padding:10px;max-height:min(480px,55vh);overflow:auto}.eiw-empty{color:var(--mute);padding:18px 8px;text-align:center}.eiw-note{font-size:11px;color:var(--mute);margin:0 0 8px}.eiw-hd{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:7px}.eiw-hd b{font-size:13px}.eiw-btn{border-color:var(--line);padding:4px 9px;background:var(--bg)}
    .eiw-timeline{position:relative;height:106px;padding-top:20px}.eiw-line{position:relative;display:flex;height:58px}.eiw-line:before{content:"";position:absolute;left:0;right:0;top:31px;border-top:2px solid #9aa1ab}.eiw-bucket{position:relative;min-width:0}.eiw-bucket+.eiw-bucket:before{content:"";position:absolute;left:0;top:22px;height:20px;border-left:2px solid #9aa1ab}.eiw-bname{position:absolute;left:6px;top:0;font-size:9px;color:var(--mute);white-space:nowrap;transform:rotate(-28deg);transform-origin:left bottom}.eiw-bcount{position:absolute;left:5px;top:42px;font-size:9px;color:var(--mute)}.eiw-dot{position:absolute;top:28px;width:7px;height:7px;margin:-1px 0 0 -3px;border-radius:50%;background:var(--brand)}.eiw-dot.newest{width:11px;height:11px;margin:-3px 0 0 -5px;box-shadow:0 0 0 3px color-mix(in srgb,var(--brand) 28%,transparent)}.eiw-dot.oldest{width:11px;height:11px;margin:-3px 0 0 -5px;background:var(--red)}.eiw-total{position:absolute;right:0;top:0;color:var(--mute);font-size:10px}.eiw-total b{font-size:14px;color:var(--brand)}.eiw-ends{display:flex;justify-content:space-between;font-size:10px;line-height:1.2}.eiw-ends span:first-child{color:var(--brand)}.eiw-ends span:last-child{text-align:right;color:var(--red)}
    .eiw-day{display:flex;gap:6px;align-items:center;overflow-x:auto;padding:5px 0}.eiw-pill{flex:none;border-left:4px solid var(--blue);background:color-mix(in srgb,var(--blue) 12%,var(--bg));border-radius:6px;padding:5px 8px;max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.eiw-strip{position:relative;height:56px;border-top:1px solid var(--line);margin:18px 0 8px}.eiw-hour{position:absolute;top:-16px;bottom:0;border-left:1px solid var(--line);font-size:9px;color:var(--mute);padding-left:2px}.eiw-event{position:absolute;top:7px;height:40px;min-width:2px;border-left:4px solid var(--brand);border-radius:5px;background:var(--meet);padding:4px 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;font-size:10px}.eiw-event.focus{background:var(--focus);border-color:var(--green)}.eiw-event.admin{background:var(--admin);border-color:#b07d12}.eiw-event.break{background:var(--rest);border-color:#7b8190}.eiw-now{position:absolute;top:-5px;bottom:0;border-left:2px solid var(--red)}.eiw-now:before{content:"Now";position:absolute;top:-13px;left:-10px;background:var(--red);color:#fff;font-size:8px;padding:0 3px;border-radius:2px}.eiw-next{font-size:11px;font-weight:650;color:var(--red);min-height:15px}.eiw-local{font-size:10px;color:var(--mute);margin-top:8px}.eiw-local-list{display:flex;gap:5px;overflow-x:auto}.eiw-local-list button{flex:none;border:1px dashed var(--brand);background:transparent;border-radius:5px;padding:4px 7px}
    .eiw-focus{position:relative;display:flex;gap:14px;align-items:center}.eiw-ring{position:relative;width:118px;height:118px;flex:none}.eiw-ring svg{width:100%;height:100%;transform:rotate(-90deg)}.eiw-ring circle{fill:none;stroke-width:9}.eiw-ring .back{stroke:var(--line)}.eiw-ring .fill{stroke:var(--green);stroke-linecap:round;stroke-dasharray:314;transition:stroke-dashoffset .2s linear}.eiw-clock{position:absolute;inset:0;display:grid;place-content:center;text-align:center}.eiw-clock b{font-size:25px;font-variant-numeric:tabular-nums}.eiw-clock span{font-size:10px;color:var(--mute)}.eiw-focus-side{flex:1;min-width:0}.eiw-parts{display:flex;gap:3px;height:42px;margin-bottom:10px}.eiw-part{display:grid;place-items:center;min-width:25px;border-radius:5px;background:var(--focus);color:var(--green);font-size:9px;opacity:.65}.eiw-part.break{background:var(--rest);color:var(--mute)}.eiw-part.current{outline:2px solid var(--brand);opacity:1}.eiw-controls{display:flex;gap:5px;flex-wrap:wrap}.eiw-finish{position:absolute;inset:-2px;background:var(--bg);display:grid;place-content:center;text-align:center;border-radius:8px;z-index:2}.eiw-finish b{font-size:18px}.eiw-finish p{color:var(--mute);margin:4px 0 8px}
    .eiw-switch{display:flex;gap:3px}.eiw-switch button{border:1px solid var(--line);background:var(--bg);padding:3px 9px;border-radius:999px}.eiw-switch button[aria-pressed="true"]{background:var(--brand);border-color:var(--brand);color:#fff}.eiw-cards{display:flex;flex-direction:column;gap:7px}.eiw-card{position:relative;height:76px;min-height:76px;display:flex;gap:7px;align-items:flex-start;border:1px solid var(--line);border-left:4px solid var(--orange);border-radius:8px;background:var(--bg);padding:8px 8px 8px 10px;overflow:hidden}.eiw-card.waiting{border-left-color:var(--purple)}.eiw-card.overdue{border-left-color:var(--red);background:color-mix(in srgb,var(--red) 5%,var(--bg))}.eiw-card.dragging{opacity:.55}.eiw-card-main{flex:1;min-width:0}.eiw-card-title,.eiw-people,.eiw-tags{height:18px;line-height:18px;white-space:nowrap;overflow-x:auto;overflow-y:hidden;scrollbar-width:none;text-overflow:ellipsis}.eiw-card-title{font-size:13px;font-weight:650}.eiw-people{font-size:11px;color:var(--mute)}.eiw-tags{display:flex;gap:5px}.eiw-tag{flex:none;border-radius:999px;padding:0 7px;background:color-mix(in srgb,var(--orange) 18%,var(--bg));color:color-mix(in srgb,var(--orange) 75%,var(--ink));font-size:10px;font-weight:600}.eiw-card.waiting .eiw-tag.main{background:color-mix(in srgb,var(--purple) 18%,var(--bg));color:color-mix(in srgb,var(--purple) 75%,var(--ink))}.eiw-tag.extra{background:color-mix(in srgb,var(--blue) 14%,var(--bg));color:color-mix(in srgb,var(--blue) 75%,var(--ink))}.eiw-age{width:42px;flex:none;text-align:center}.eiw-age b{display:block;font-size:15px}.eiw-age span{font-size:9px;color:var(--mute)}.eiw-open{align-self:center;width:28px;height:28px;border:1px solid transparent;background:transparent;border-radius:6px}.eiw-open:hover{border-color:var(--line);color:var(--brand)}.eiw-more{width:100%;margin-top:7px}
    #${ROOT_ID}[data-collapsed="true"] .eiw-tabs,#${ROOT_ID}[data-collapsed="true"] .eiw-panel{display:none}
    [data-testid="virtuoso-scroller"][data-erin-inbox-scan]{visibility:hidden!important;pointer-events:none!important}
    @container(max-width:519px){.eiw-bname,.eiw-bcount{display:none}.eiw-line{height:48px}.eiw-line:before{top:23px}.eiw-dot{top:20px}.eiw-ends{margin-top:-6px}.eiw-focus{align-items:flex-start}.eiw-ring{width:96px;height:96px}.eiw-clock b{font-size:20px}.eiw-tab{font-size:11px}.eiw-event{padding:4px 2px}.eiw-summary{max-width:90px}}
    @media(prefers-reduced-motion:reduce){#${ROOT_ID} *{scroll-behavior:auto!important;transition:none!important}}
  `);

  GM_addStyle(`
    #${ROOT_ID}{padding:0 0 8px;--on-brand:#fff;min-width:0}
    #${ROOT_ID} .eiw-shell{width:100%;min-width:0}
    #${ROOT_ID} .eiw-top{height:44px;gap:8px;padding:5px 7px 5px 12px}
    #${ROOT_ID} .eiw-title{flex:0 0 auto;font-size:14px}
    #${ROOT_ID} .eiw-summary{flex:1;min-width:0;overflow:visible;text-overflow:clip;font-size:11px;line-height:1.1}
    #${ROOT_ID} .eiw-summary b{font-size:15px;color:var(--brand);font-weight:750}
    #${ROOT_ID} .eiw-tab[aria-selected="true"],#${ROOT_ID} .eiw-switch button[aria-pressed="true"]{color:var(--on-brand)!important}
    #${ROOT_ID} .eiw-collapse{position:relative;font-size:0}
    #${ROOT_ID} .eiw-collapse:after{content:"";display:block;width:18px;height:18px;background:currentColor;mask:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath d='M3 5.5 8 10.5l5-5' fill='none' stroke='black' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E") center/contain no-repeat;transition:transform .18s ease}
    #${ROOT_ID}[data-collapsed="true"] .eiw-collapse:after{transform:rotate(180deg)}
    #${ROOT_ID} .eiw-panel{overflow:hidden;padding:10px 12px 12px}
    #${ROOT_ID} .eiw-timeline{height:98px;padding-top:12px}.eiw-total,.eiw-bcount{display:none}.eiw-line{height:54px}
    #${ROOT_ID} .eiw-all-day{display:flex;align-items:flex-start;gap:6px;margin:4px 0 9px}
    #${ROOT_ID} .eiw-day{flex:1;display:flex;flex-wrap:wrap;align-content:flex-start;gap:5px;padding:0;overflow:hidden;max-height:30px}
    #${ROOT_ID} .eiw-day.expanded{max-height:none}.eiw-pill{max-width:min(240px,100%)}
    #${ROOT_ID} .eiw-day-toggle{display:none;flex:none}.eiw-day-toggle.visible{display:block}
    #${ROOT_ID} .eiw-strip{height:70px;margin:19px 0 10px;touch-action:none;user-select:none;cursor:crosshair}
    #${ROOT_ID} .eiw-selection{position:absolute;top:6px;height:54px;border:2px dashed var(--brand);background:color-mix(in srgb,var(--brand) 12%,transparent);border-radius:5px;pointer-events:none}
    #${ROOT_ID} .eiw-composer{display:grid;grid-template-columns:minmax(110px,1fr) auto;gap:6px;align-items:center;padding:8px;border:1px solid var(--brand);border-radius:8px;background:var(--bg);margin:8px 0}
    #${ROOT_ID} .eiw-composer input{min-width:0;height:30px;border:1px solid var(--line);border-radius:5px;padding:0 8px;background:var(--bg)}
    #${ROOT_ID} .eiw-kinds{display:flex;gap:4px;flex-wrap:wrap}.eiw-kinds button{padding:3px 7px}
    #${ROOT_ID} .eiw-cf{position:relative;height:min(355px,42vh);overflow:auto;overscroll-behavior:contain;padding:4px 3px 74px;scrollbar-width:thin}
    #${ROOT_ID} .eiw-cfst{position:relative}.eiw-card{position:absolute;left:0;right:0;transform-origin:50% 0;transition:transform .16s ease,opacity .16s ease,box-shadow .16s ease;height:92px;min-height:92px;padding:9px 40px 8px 11px;display:block}
    #${ROOT_ID} .eiw-card.expanded{position:relative!important;transform:none!important;margin-bottom:7px;opacity:1!important}
    #${ROOT_ID} .eiw-cf.expanded{height:min(480px,55vh)}#${ROOT_ID} .eiw-cf.expanded .eiw-cfst{height:auto!important}
    #${ROOT_ID} .eiw-card-title{height:auto;min-height:19px;line-height:19px;white-space:nowrap;padding-right:95px;cursor:default}
    #${ROOT_ID} .eiw-reply{position:absolute;right:38px;top:8px;max-width:120px;text-align:right;color:var(--mute);font-size:9px;line-height:1.2;white-space:nowrap}.eiw-reply b{font-size:11px;color:var(--ink)}
    #${ROOT_ID} .eiw-open{position:absolute;right:7px;top:31px}.eiw-grip{position:absolute;right:8px;bottom:7px;border:0;background:transparent;color:var(--mute);cursor:grab;font-size:16px;touch-action:none}
    #${ROOT_ID} .eiw-people{display:flex;gap:0;height:18px;overflow:hidden}.eiw-person+.eiw-person:before{content:"·";padding:0 5px;color:var(--mute)}
    #${ROOT_ID} .eiw-flow-foot{display:flex;gap:6px;justify-content:flex-end;align-items:center;margin-top:7px}.eiw-jump{opacity:0;pointer-events:none;transition:opacity .15s}.eiw-jump.visible{opacity:1;pointer-events:auto}
    #${ROOT_ID} .eiw-tags{margin-top:2px}.eiw-tag{background:var(--tag-bg,color-mix(in srgb,var(--orange) 18%,var(--bg)));color:var(--tag-fg,color-mix(in srgb,var(--orange) 75%,var(--ink)));border:1px solid var(--tag-border,transparent)}
    @container(max-width:519px){#${ROOT_ID} .eiw-summary{font-size:10px}#${ROOT_ID} .eiw-summary b{font-size:14px}#${ROOT_ID} .eiw-panel{padding-inline:8px}#${ROOT_ID} .eiw-composer{grid-template-columns:1fr}.eiw-card{height:96px;min-height:96px}.eiw-reply{max-width:96px}}
  `);

  function classify(title) {
    const t = title.toLowerCase();
    if (/focus|write|draft|review/.test(t)) return 'focus';
    if (/lunch|break|birthday|out of office/.test(t)) return 'break';
    if (/admin|timesheet|pay|deadline|due/.test(t)) return 'admin';
    return 'meeting';
  }

  function parseDateTitle(title) {
    if (!title) return null;
    const cleaned = title.replace(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+/i, '');
    const date = new Date(cleaned);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function rowDate(row, group) {
    const tokens = [...row.querySelectorAll('span,time')].map(n => n.textContent.trim()).filter(Boolean);
    const clock = tokens.find(t => /^\d{1,2}:\d{2}\s*[AP]M$/i.test(t));
    if (clock && /today|pinned/i.test(group || '')) return new Date(`${todayKey()} ${clock}`);
    const dated = tokens.find(t => /^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}/i.test(t));
    return dated ? parseDateTitle(dated) : null;
  }

  function scrapeMessages() {
    const list = messageList();
    if (!list) return false;
    let changed = false;
    let currentGroup = '';
    const stream = list.querySelector('[data-testid="virtuoso-item-list"]') || list;
    [...stream.querySelectorAll('[id^="groupHeader"], [role="option"][data-convid]')].forEach(el => {
      if (el.id.startsWith('groupHeader')) {
        currentGroup = el.getAttribute('aria-label') || el.textContent.trim();
        return;
      }
      const id = el.dataset.convid;
      if (!id) return;
      const mark = el.querySelector('button[title="Mark as read"],button[aria-label="Mark as read"]');
      const unread = !!mark || /^Unread\b/i.test(el.getAttribute('aria-label') || '');
      const categoryNodes = [...el.querySelectorAll('[aria-label^="Search for all messages with the category"], .fui-Tag__primaryText')];
      const categoryStyles = categoryNodes.map(n => {
        const name = n.textContent.trim();
        const tag = n.closest('[class*="Tag"],button,span') || n;
        const style = getComputedStyle(tag);
        return { name, bg: style.backgroundColor, color: style.color, border: style.borderColor };
      }).filter(x => x.name);
      const categories = [...new Set(categoryStyles.map(x => x.name))];
      const sender = el.querySelector('span[title*="@"]')?.textContent.trim() ||
        el.querySelector('.ESO13 span')?.textContent.trim() ||
        [...el.querySelectorAll('span')].find(n => visible(n) && getComputedStyle(n).fontSize === '14px' && !n.className.includes('Avatar'))?.textContent.trim() || 'Unknown sender';
      const subject = el.querySelector('.TtcXM')?.textContent.trim() ||
        [...el.querySelectorAll('span')].find(n => visible(n) && getComputedStyle(n).fontSize === '13px' && n.textContent.trim().length > 4)?.textContent.trim() || '(no subject)';
      const dated = [...el.querySelectorAll('[title]')].find(n => /\b20\d{2}\b/.test(n.title) && /\d{1,2}:\d{2}/.test(n.title));
      const date = parseDateTitle(dated?.title) || rowDate(el, currentGroup);
      const old = cache.get(id) || {};
      const next = { id, unread, categories, categoryStyles, sender, subject, date: date || old.date || null, group: currentGroup || old.group || '', row: el };
      if (!old.id || old.unread !== next.unread || old.subject !== next.subject || old.sender !== next.sender || old.group !== next.group || String(old.date) !== String(next.date) || old.categories?.join('\n') !== next.categories.join('\n')) changed = true;
      cache.set(id, next);
    });
    return changed;
  }

  function scrapeMyDay() {
    const drawer = document.querySelector('[data-app-section="DrawerPane"]');
    if (!drawer || !drawer.querySelector('[role="tab"][aria-selected="true"][name="Calendar"],#time-panel-pivot-Calendar[aria-selected="true"]')) return false;
    const lines = drawer.innerText.split('\n').map(x => x.trim()).filter(Boolean);
    const start = lines.indexOf('Today');
    if (start < 0) return false;
    let end = lines.indexOf('Tomorrow', start + 1);
    if (end < 0) end = lines.length;
    const slice = lines.slice(start + 1, end).filter(x => !/^(Fri|Sat|Sun|Mon|Tue|Wed|Thu),\s/.test(x));
    const events = [];
    for (let i = 0; i < slice.length; i++) {
      if (slice[i] === 'All day' && slice[i + 1]) {
        events.push({ allDay: true, title: slice[++i], kind: classify(slice[i]) });
      } else if (/^\d{1,2}:\d{2}\s[AP]M$/.test(slice[i]) && slice[i + 1]) {
        const date = new Date(`${todayKey()} ${slice[i]}`);
        const title = slice[++i];
        const duration = /^\d+(?:\.\d+)?\s*(?:min|mins|hr|hrs)$/.test(slice[i + 1] || '') ? slice[++i] : '30 min';
        const amount = parseFloat(duration) || 30;
        const minutes = /hr/.test(duration) ? amount * 60 : amount;
        events.push({ allDay: false, title, start: date.getHours() * 60 + date.getMinutes(), end: date.getHours() * 60 + date.getMinutes() + minutes, kind: classify(title) });
      }
    }
    const changed = state.scheduleDate !== todayKey() || JSON.stringify(state.schedule) !== JSON.stringify(events);
    if (changed) {
      state.scheduleDate = todayKey();
      state.schedule = events;
      save();
    }
    return changed;
  }

  function bucketFor(item) {
    if (!item.date) {
      const g = item.group.toLowerCase();
      if (g.includes('today')) return 0;
      if (g.includes('yesterday')) return 1;
      return 5;
    }
    const days = Math.max(0, (Date.now() - item.date.getTime()) / 864e5);
    return days < 1 ? 0 : days < 2 ? 1 : days < 7 ? 2 : days < 31 ? 3 : days < 62 ? 4 : 5;
  }

  function formatStamp(item, short = false) {
    if (!item?.date) return item?.group || 'Date unavailable';
    return item.date.toLocaleString([], short ? { month: 'short', day: 'numeric' } : { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
  }

  function timelineHTML(items) {
    if (!items.length) return `<div class="eiw-empty">${scanning ? 'Scanning the inbox…' : 'No inbox messages loaded yet.'}</div>`;
    const names = ['Today', 'Yesterday', 'This week', 'This month', 'Last month', 'Earlier'];
    const buckets = names.map((name, index) => ({ name, items: items.filter(x => bucketFor(x) === index) }));
    const unread = items.filter(x => x.unread).sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0));
    const newest = unread[0], oldest = unread[unread.length - 1];
    const groups = buckets.map(bucket => {
      const count = bucket.items.length || 0.25;
      const dots = bucket.items.map((item, i) => item.unread ? `<i class="eiw-dot${item === newest ? ' newest' : ''}${item === oldest ? ' oldest' : ''}" style="left:${((i + .5) / bucket.items.length) * 100}%" title="${esc(item.subject)} — ${esc(formatStamp(item))}"></i>` : '').join('');
      return `<div class="eiw-bucket" style="flex:${count}"><span class="eiw-bname">${bucket.name}</span><span class="eiw-bcount"><b>${bucket.items.filter(x => x.unread).length}</b> / ${bucket.items.length}</span>${dots}</div>`;
    }).join('');
    return `<div class="eiw-timeline"><div class="eiw-line">${groups}</div><div class="eiw-ends"><span><b>Newest unread</b><br>${esc(formatStamp(newest))}</span><span><b>Oldest unread</b><br>${esc(formatStamp(oldest, true))}</span></div></div>`;
  }

  function scheduleHTML() {
    const events = state.scheduleDate === todayKey() ? state.schedule : [];
    const allDay = events.filter(x => x.allDay);
    const timed = events.filter(x => !x.allDay).concat(state.localBlocks.filter(x => x.date === todayKey() && Number.isFinite(x.start) && Number.isFinite(x.end)));
    const dayStart = 8 * 60, dayEnd = 17 * 60, nowDate = new Date(), now = nowDate.getHours() * 60 + nowDate.getMinutes();
    const pos = m => Math.max(0, Math.min(100, (m - dayStart) / (dayEnd - dayStart) * 100));
    const hours = Array.from({length: 10}, (_, i) => `<i class="eiw-hour" style="left:${i / 9 * 100}%">${(i + 8) > 12 ? i - 4 : i + 8}</i>`).join('');
    const blocks = timed.filter(x => x.end > dayStart && x.start < dayEnd).map(x => `<div class="eiw-event ${x.kind}" title="${esc(x.title)}" style="left:${pos(x.start)}%;width:${Math.max(1.5, pos(x.end) - pos(x.start))}%">${esc(x.title)}</div>`).join('');
    const upcoming = timed.filter(x => x.end > now).sort((a,b) => a.start-b.start)[0];
    const mins = upcoming ? Math.max(0, upcoming.start - now) : 0;
    const composer = pendingBlock ? `<div class="eiw-composer"><input data-block-title aria-label="Time block title" placeholder="What are you blocking time for?" value="${esc(pendingBlock.title || '')}"><div class="eiw-kinds">${[['meeting','Meeting'],['focus','Focus'],['admin','Admin'],['break','Break']].map(([k,l]) => `<button class="eiw-btn" data-save-kind="${k}">${l}</button>`).join('')}<button class="eiw-btn" data-action="cancel-block">Cancel</button></div></div>` : '';
    return `<div class="eiw-hd"><b>Today · ${new Date().toLocaleDateString([], {month:'short',day:'numeric'})}</b><span class="eiw-note">Drag the timeline to block local time</span></div>
      <div class="eiw-all-day"><div class="eiw-day ${state.allDayExpanded ? 'expanded' : ''}">${allDay.length ? allDay.map(x => `<span class="eiw-pill">${esc(x.title)}</span>`).join('') : '<span class="eiw-note">No all-day events</span>'}</div><button class="eiw-btn eiw-day-toggle" data-action="all-day-toggle">${state.allDayExpanded ? 'Collapse' : 'All'}</button></div>
      <div class="eiw-strip" data-schedule-strip data-start="${dayStart}" data-end="${dayEnd}">${hours}${blocks}${now >= dayStart && now <= dayEnd ? `<i class="eiw-now" style="left:${pos(now)}%"></i>` : ''}<i class="eiw-selection" hidden></i></div>${composer}
      <div class="eiw-next">${upcoming ? `Next: ${esc(upcoming.title)} · ${mins ? `in ${mins}m` : 'now'}` : 'No upcoming timed event loaded'}</div>
      <div class="eiw-local-list">${state.localBlocks.filter(x => x.date === todayKey()).map(x => `<button data-remove-block="${esc(x.id)}" title="Remove local block">${esc(minutesLabel(x.start))} · ${esc(x.title)}</button>`).join('')}</div>`;
  }

  function minutesLabel(minutes) {
    const hour = Math.floor(minutes / 60), minute = minutes % 60;
    return new Date(2000, 0, 1, hour, minute).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  }

  function focusHTML() {
    const part = PARTS[focus.part], total = part[1] * 60, remaining = Math.max(0, Math.ceil(total - focus.elapsed));
    const mm = String(Math.floor(remaining / 60)).padStart(2, '0'), ss = String(remaining % 60).padStart(2, '0');
    const offset = 314 * (1 - Math.min(1, focus.elapsed / total));
    const labels = ['Focus finished', 'Rest finished', 'Focus finished', 'Cycle complete'];
    const next = ['Start rest', 'Start focus 2', 'Start break', 'Start new cycle'];
    return `<div class="eiw-focus"><div class="eiw-ring"><svg viewBox="0 0 120 120"><circle class="back" cx="60" cy="60" r="50"/><circle class="fill" cx="60" cy="60" r="50" style="stroke-dashoffset:${offset}"/></svg><div class="eiw-clock"><b>${mm}:${ss}</b><span>${part[2]}</span></div></div>
      <div class="eiw-focus-side"><div class="eiw-parts">${PARTS.map((x,i) => `<div class="eiw-part ${x[0] === 'break' ? 'break' : ''} ${i === focus.part ? 'current' : ''}" style="flex:${x[1]}">${x[1]}m</div>`).join('')}</div><div class="eiw-controls"><button class="eiw-btn" data-action="focus-toggle">${focus.running ? 'Pause' : 'Start'}</button><button class="eiw-btn" data-action="focus-skip">Skip</button><button class="eiw-btn" data-action="focus-reset">Reset</button></div></div>
      ${focus.finished ? `<div class="eiw-finish"><b>${labels[focus.part]}</b><p>The timer is paused and waiting for you.</p><button class="eiw-btn" data-action="focus-next">${next[focus.part]}</button></div>` : ''}</div>`;
  }

  function todoItems(kind) {
    const prefix = kind === 'self' ? /^Self\s*:\s*/i : /^Waiting(?:\s+on)?\s*:\s*/i;
    return [...cache.values()].filter(x => x.categories.some(c => prefix.test(c))).map(x => {
      const main = x.categories.find(c => prefix.test(c)) || (kind === 'self' ? 'Self' : 'Waiting');
      return Object.assign({}, x, { action: main.replace(prefix, '').trim() || (kind === 'self' ? 'Follow up' : 'Reply'), extras: x.categories.filter(c => c !== main), days: x.date ? Math.max(0, Math.floor((Date.now() - x.date) / 864e5)) : 0 });
    });
  }

  function todoHTML() {
    let items = todoItems(state.todo);
    if (state.todo === 'waiting') items.sort((a,b) => b.days-a.days);
    else if (state.selfOrder.length) items.sort((a,b) => (state.selfOrder.indexOf(a.id) + 1 || 9999) - (state.selfOrder.indexOf(b.id) + 1 || 9999));
    const limit = matchMedia('(max-width:700px)').matches ? 3 : 5;
    const shown = state.todoExpanded ? items : items;
    const cards = shown.map(x => {
      const title = state.titleEdits[x.id] || x.subject;
      const people = x.sender.split(/\s*;\s*/).filter(Boolean).map(p => `<span class="eiw-person">${esc(p)}</span>`).join('');
      const styleFor = name => x.categoryStyles?.find(s => s.name === name) || {};
      const tag = (name, cls = '') => { const s = styleFor(name); return `<span class="eiw-tag ${cls}" style="--tag-bg:${esc(s.bg || '')};--tag-fg:${esc(s.color || '')};--tag-border:${esc(s.border || '')}">${esc(name)}</span>`; };
      const reply = x.date ? `<span class="eiw-reply"><i>Last reply:</i> ${x.date.toLocaleDateString([], {month:'short',day:'numeric'})} · <b>${x.days}</b> days ago</span>` : '';
      return `<article class="eiw-card ${state.todoExpanded ? 'expanded' : ''} ${state.todo === 'waiting' ? 'waiting' : ''} ${x.days > 30 ? 'overdue' : ''}" data-message-id="${esc(x.id)}"><div class="eiw-card-title" data-edit-title="${esc(x.id)}">${esc(title)}</div>${reply}<div class="eiw-people">${people}</div><div class="eiw-tags">${tag(x.categories.find(c => /^Self|^Waiting/i.test(c)) || x.action, 'main')}${x.extras.map(t => tag(t, 'extra')).join('')}</div><button class="eiw-open" data-open-message="${esc(x.id)}" title="View message" aria-label="View message: ${esc(title)}">→</button>${state.todo === 'self' && state.todoExpanded ? '<button class="eiw-grip" data-grip title="Drag to reorder" aria-label="Drag to reorder">⋮⋮</button>' : ''}</article>`;
    }).join('');
    return `<div class="eiw-hd"><div class="eiw-switch"><button data-todo="self" aria-pressed="${state.todo === 'self'}">Self ${todoItems('self').length}</button><button data-todo="waiting" aria-pressed="${state.todo === 'waiting'}">Waiting ${todoItems('waiting').length}</button></div></div><div class="eiw-cf ${state.todoExpanded ? 'expanded' : ''}" data-flow data-limit="${limit}"><div class="eiw-cfst">${cards || `<div class="eiw-empty">No inbox messages have this category.</div>`}</div></div><div class="eiw-flow-foot"><button class="eiw-btn eiw-jump" data-action="todo-top">↑ Top</button>${items.length > limit ? `<button class="eiw-btn" data-action="todo-expand">${state.todoExpanded ? `Collapse to top ${limit}` : `Show all ${items.length}`}</button>` : ''}</div>`;
  }

  function panelHTML() {
    const items = [...cache.values()];
    if (state.tab === 'timeline') return timelineHTML(items);
    if (state.tab === 'schedule') return scheduleHTML();
    if (state.tab === 'focus') return focusHTML();
    return todoHTML();
  }

  function render() {
    if (!root) return;
    const unread = [...cache.values()].filter(x => x.unread).length;
    root.dataset.collapsed = String(state.collapsed);
    setMarkup(root, `<div class="eiw-shell"><div class="eiw-top"><span class="eiw-title">Widgets</span><span class="eiw-summary"><b>${unread}</b> unread / ${cache.size} in inbox</span><button class="eiw-iconbtn" data-action="scan" title="Refresh inbox data" aria-label="Refresh inbox data">↻</button><button class="eiw-iconbtn eiw-collapse" data-action="collapse" aria-expanded="${!state.collapsed}" aria-label="${state.collapsed ? 'Open' : 'Close'} Widgets"></button></div><div class="eiw-tabs" role="tablist">${[['timeline','Unread'],['schedule','Schedule'],['focus','Focus'],['todo','To-do']].map(([id,label]) => `<button class="eiw-tab" role="tab" data-tab="${id}" aria-selected="${state.tab === id}">${label}</button>`).join('')}</div><div class="eiw-panel" role="tabpanel">${panelHTML()}</div></div>`);
    requestAnimationFrame(() => { syncGeometryAndTheme(); initAllDay(); initFlow(); });
  }

  function syncGeometryAndTheme() {
    if (!root?.isConnected) return;
    const region = messageList()?.querySelector('[role="complementary"]');
    if (region) {
      const rr = region.getBoundingClientRect(), hr = root.parentElement.getBoundingClientRect();
      root.style.width = `${rr.width}px`;
      root.style.marginLeft = `${rr.left - hr.left}px`;
    }
    const newMail = [...document.querySelectorAll('button')].find(b => (b.getAttribute('aria-label') || b.textContent.trim()) === 'New mail');
    if (newMail) { const s = getComputedStyle(newMail); root.style.setProperty('--brand', s.backgroundColor); root.style.setProperty('--on-brand', s.color); }
  }

  function initAllDay() {
    const day = root?.querySelector('.eiw-day'), toggle = root?.querySelector('.eiw-day-toggle');
    if (day && toggle && (day.scrollHeight > day.clientHeight + 2 || state.allDayExpanded)) toggle.classList.add('visible');
  }

  function initFlow() {
    flowCleanup?.(); flowCleanup = null;
    const flow = root?.querySelector('[data-flow]');
    if (!flow || state.todoExpanded) return;
    const cards = [...flow.querySelectorAll('.eiw-card')], stage = flow.querySelector('.eiw-cfst'), jump = root.querySelector('.eiw-jump');
    if (!cards.length) return;
    const gap = 13, base = 92, limit = Number(flow.dataset.limit) || 5;
    stage.style.height = `${Math.max(flow.clientHeight, base + (cards.length - 1) * 45)}px`;
    const update = () => {
      const y = flow.scrollTop;
      cards.forEach((card, i) => {
        const natural = i * 45, distance = natural - y, rank = Math.max(0, Math.min(limit - 1, distance / 45));
        const scale = .78 + .22 * (1 - rank / Math.max(1, limit - 1));
        card.style.transform = `translateY(${natural}px) scale(${scale})`;
        card.style.opacity = String(.48 + .52 * (1 - rank / Math.max(1, limit - 1)));
        card.style.zIndex = String(cards.length - i);
      });
      jump?.classList.toggle('visible', y > 50);
    };
    flow.addEventListener('scroll', update, {passive:true}); update();
    flowCleanup = () => flow.removeEventListener('scroll', update);
  }

  function scheduleRender() {
    if (scanning || !widgetContext()) return;
    clearTimeout(renderTimer);
    renderTimer = setTimeout(() => {
      const messagesChanged = scrapeMessages();
      const scheduleChanged = scrapeMyDay();
      if (messagesChanged || scheduleChanged) render();
    }, 120);
  }

  function mount() {
    if (!widgetContext()) {
      document.getElementById(ROOT_ID)?.remove();
      root = null;
      scannedScroller = null;
      scanComplete = false;
      cache.clear();
      return;
    }
    const list = messageList();
    const region = list ? [...list.querySelectorAll('[role="complementary"]')].find(visible) ||
      list.querySelector('[role="complementary"]') : null;
    if (!list || !region) return;
    const host = region.parentElement;
    if (!host) return;
    let created = false;
    if (!root || !root.isConnected) {
      root = document.createElement('section');
      root.id = ROOT_ID;
      root.setAttribute('aria-label', 'Widgets');
      host.insertBefore(root, region);
      bind();
      created = true;
    } else if (root.parentElement !== host || root.nextElementSibling !== region) host.insertBefore(root, region);
    const messagesChanged = scrapeMessages();
    const scheduleChanged = scrapeMyDay();
    if (created || messagesChanged || scheduleChanged) render(); else syncGeometryAndTheme();
    const scroller = list.querySelector('[data-testid="virtuoso-scroller"]');
    if (scroller && scroller !== scannedScroller && !scanning) setTimeout(() => scanAllMessages(false), 350);
  }

  const settleVirtualList = () => new Promise(resolve => {
    requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 70)));
  });

  async function scanAllMessages(force) {
    if (scanning || !widgetContext()) return;
    const scroller = messageList()?.querySelector('[data-testid="virtuoso-scroller"]');
    if (!scroller || (!force && scroller === scannedScroller)) return;
    scannedScroller = scroller;
    scanning = true;
    scanComplete = false;
    scanProgress = 0;
    cache.clear();
    render();
    const originalTop = scroller.scrollTop;
    scroller.setAttribute('data-erin-inbox-scan', '');
    try {
      const max = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
      const step = Math.max(220, Math.floor(scroller.clientHeight * .82));
      const positions = [];
      for (let top = 0; top < max; top += step) positions.push(top);
      positions.push(max);
      for (let index = 0; index < positions.length; index++) {
        if (!widgetContext()) break;
        scroller.scrollTop = positions[index];
        await settleVirtualList();
        scrapeMessages();
        scanProgress = Math.round((index + 1) / positions.length * 100);
        if (index % 4 === 0 || index === positions.length - 1) render();
      }
    } finally {
      scroller.scrollTop = originalTop;
      await settleVirtualList();
      scrapeMessages();
      scroller.removeAttribute('data-erin-inbox-scan');
      scanning = false;
      scanComplete = widgetContext();
      scanProgress = 100;
      if (widgetContext()) render(); else mount();
    }
  }

  function startFocus() {
    focus.running = true;
    focus.finished = false;
    focus.last = Date.now();
    clearInterval(focusTick);
    focusTick = setInterval(() => {
      if (!focus.running) return;
      const now = Date.now();
      focus.elapsed += (now - focus.last) / 1000;
      focus.last = now;
      if (focus.elapsed >= PARTS[focus.part][1] * 60) {
        focus.elapsed = PARTS[focus.part][1] * 60;
        focus.running = false;
        focus.finished = true;
        clearInterval(focusTick);
      }
      if (state.tab === 'focus') render();
    }, 250);
  }

  function bind() {
    root.addEventListener('click', event => {
      const target = event.target.closest('button');
      if (!target) return;
      if (target.dataset.tab) { state.tab = target.dataset.tab; save(); render(); return; }
      if (target.dataset.todo) { state.todo = target.dataset.todo; state.todoExpanded = false; save(); render(); return; }
      if (target.dataset.openMessage) {
        const item = cache.get(target.dataset.openMessage);
        if (item?.row?.isConnected) item.row.click();
        else location.assign(`/mail/inbox/id/${encodeURIComponent(target.dataset.openMessage)}`);
        return;
      }
      if (target.dataset.removeBlock) { state.localBlocks = state.localBlocks.filter(x => x.id !== target.dataset.removeBlock); save(); render(); return; }
      if (target.dataset.saveKind && pendingBlock) {
        const title = root.querySelector('[data-block-title]')?.value.trim() || 'Time block';
        state.localBlocks.push({ id: `${Date.now()}`, date: todayKey(), title, start: pendingBlock.start, end: pendingBlock.end, kind: target.dataset.saveKind });
        pendingBlock = null; save(); render(); return;
      }
      switch (target.dataset.action) {
        case 'collapse': state.collapsed = !state.collapsed; save(); render(); break;
        case 'scan': scanAllMessages(true); break;
        case 'todo-expand': state.todoExpanded = !state.todoExpanded; save(); render(); break;
        case 'todo-top': root.querySelector('[data-flow]')?.scrollTo({top:0,behavior:'smooth'}); break;
        case 'all-day-toggle': state.allDayExpanded = !state.allDayExpanded; save(); render(); break;
        case 'cancel-block': pendingBlock = null; render(); break;
        case 'focus-toggle': focus.running ? (focus.running = false, clearInterval(focusTick), render()) : startFocus(); break;
        case 'focus-skip': focus.elapsed = PARTS[focus.part][1] * 60; focus.running = false; focus.finished = true; clearInterval(focusTick); render(); break;
        case 'focus-reset': focus = { part: 0, elapsed: 0, running: false, finished: false, last: 0 }; clearInterval(focusTick); render(); break;
        case 'focus-next': focus = focus.part === PARTS.length - 1 ? { part: 0, elapsed: 0, running: false, finished: false, last: 0 } : { part: focus.part + 1, elapsed: 0, running: false, finished: false, last: 0 }; render(); break;
      }
    });
    root.addEventListener('dblclick', event => {
      const title = event.target.closest('[data-edit-title]');
      if (!title) return;
      title.contentEditable = 'true'; title.focus(); document.getSelection()?.selectAllChildren(title);
      const finish = () => { title.contentEditable = 'false'; state.titleEdits[title.dataset.editTitle] = title.textContent.trim() || cache.get(title.dataset.editTitle)?.subject; save(); title.removeEventListener('blur', finish); };
      title.addEventListener('blur', finish);
      title.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); title.blur(); } });
    });
    root.addEventListener('pointerdown', event => {
      const strip = event.target.closest('[data-schedule-strip]');
      if (strip && event.button === 0) {
        const rect = strip.getBoundingClientRect(), startDay = +strip.dataset.start, endDay = +strip.dataset.end;
        const snap = x => Math.round((startDay + Math.max(0, Math.min(1, (x - rect.left) / rect.width)) * (endDay - startDay)) / 15) * 15;
        const start = snap(event.clientX), selection = strip.querySelector('.eiw-selection');
        strip.setPointerCapture(event.pointerId); selection.hidden = false;
        const draw = x => { const end = snap(x), a = Math.min(start,end), b = Math.max(start,end) + (end === start ? 30 : 0); selection.style.left = `${(a-startDay)/(endDay-startDay)*100}%`; selection.style.width = `${(b-a)/(endDay-startDay)*100}%`; return [a,b]; };
        let range = draw(event.clientX);
        const move = e => { range = draw(e.clientX); };
        const up = e => { strip.removeEventListener('pointermove',move); strip.removeEventListener('pointerup',up); range = draw(e.clientX); pendingBlock = {start:range[0],end:Math.max(range[0]+15,range[1])}; render(); };
        strip.addEventListener('pointermove',move); strip.addEventListener('pointerup',up); return;
      }
      const grip = event.target.closest('[data-grip]');
      if (!grip || event.button !== 0) return;
      const card = grip.closest('.eiw-card'), id = card.dataset.messageId;
      grip.setPointerCapture(event.pointerId); card.classList.add('dragging');
      const move = e => {
        const over = document.elementFromPoint(e.clientX,e.clientY)?.closest('.eiw-card');
        if (!over || over === card) return;
        const ids = todoItems('self').map(x=>x.id).sort((a,b)=>(state.selfOrder.indexOf(a)+1||9999)-(state.selfOrder.indexOf(b)+1||9999));
        const from=ids.indexOf(id), to=ids.indexOf(over.dataset.messageId); if(from<0||to<0)return;
        ids.splice(to,0,ids.splice(from,1)[0]); state.selfOrder=ids; save();
        over.parentElement.insertBefore(card, to>from ? over.nextSibling : over);
      };
      const up = () => { grip.removeEventListener('pointermove',move); grip.removeEventListener('pointerup',up); card.classList.remove('dragging'); render(); };
      grip.addEventListener('pointermove',move); grip.addEventListener('pointerup',up);
    });
  }

  observer = new MutationObserver(mutations => {
    if (root && mutations.every(mutation => root.contains(mutation.target))) return;
    if (scanning) return;
    mount();
    scheduleRender();
  });
  observer.observe(document.documentElement, { childList: true, subtree: true, attributes: true, attributeFilter: ['aria-label', 'title'] });
  window.addEventListener('popstate', mount);
  setInterval(mount, 1500);
  mount();
})();
