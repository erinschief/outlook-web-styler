const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const base = path.resolve(__dirname, '..');
const widgetSource = fs.readFileSync(path.join(base, 'outlook-inbox-widgets.user.js'), 'utf8');
const stylerSource = fs.readFileSync(path.join(base, 'outlook-web-styler.user.js'), 'utf8');
// Exercise the actual functions, rather than copies of their implementations.
function section(source, from, to) {
  const start = source.indexOf(from), end = source.indexOf(to, start);
  assert(start >= 0 && end > start, `Missing source boundary: ${from}`);
  return source.slice(start, end);
}
function context(values = {}) { return vm.createContext({ Date, Map, Set, console, ...values }); }
function load(ctx, from, to, source = widgetSource) { vm.runInContext(section(source, from, to), ctx); return ctx; }

test('Both full userscripts initialize against a loading Outlook shell', async () => {
  for (const source of [widgetSource, stylerSource]) {
    const frames = [], callbacks = [], styles = [];
    const html = { dataset: {}, style: { getPropertyValue:()=>'', setProperty(){},removeProperty(){} }, addEventListener(){} };
    const doc = { documentElement:html, body:{}, hidden:false, activeElement:null,
      getElementById:id=>id==='OwaTitleBar'?{}:null, querySelector:()=>null, querySelectorAll:()=>[], addEventListener(){} };
    const ctx = context({ document:doc, window:{addEventListener(){}}, Element:class{}, Node:{ELEMENT_NODE:1,TEXT_NODE:3},
      location:{host:'outlook.office.com',pathname:'/mail/inbox'}, MutationObserver:class{observe(){}disconnect(){}},ResizeObserver:class{observe(){}disconnect(){}},
      GM_addStyle:value=>styles.push(value), GM_getValue:(_,fallback)=>fallback, GM_setValue(){},GM_registerMenuCommand(){},
      requestAnimationFrame:fn=>{frames.push(fn);return frames.length}, setTimeout:fn=>{callbacks.push(fn);return callbacks.length},clearTimeout(){},
      setInterval:()=>1,clearInterval(){},localStorage:{getItem:()=>null},performance:{now:()=>1000},getComputedStyle:()=>({}),innerWidth:1200 });
    await vm.runInContext(source,ctx);await Promise.resolve();
    for(let i=0;i<frames.length && i<10;i++) frames[i]();
    assert(styles.length>0,'No stylesheet installed');
  }
});

test('Calendar boundaries use local days, Monday weeks, months, and years', () => {
  const c=load(context(),'  function bucketFor(item,','  function formatStamp');
  const now=new Date(2026,10,8,9);
  for(const [date,expected] of [[new Date(2026,10,8),0],[new Date(2026,10,7,23,59),1],[new Date(2026,10,3),2],[new Date(2026,10,1),3],[new Date(2026,9,31),4],[new Date(2026,8,30),5]]) assert.equal(c.bucketFor({date},now),expected);
  assert.equal(c.bucketFor({date:new Date(2025,11,31)},new Date(2026,0,10)),4);
  // DST ends on Nov 1: calendar grouping must not rely on a 24-hour subtraction.
  assert.equal(c.bucketFor({date:new Date(2026,9,31,23,59)},new Date(2026,10,1,9)),1);
});

test('12/24-hour clocks, ranges, compound durations, and invalid input',()=>{
  const c=load(context(),'  function parseClock(text)','  function scrapeMyDay()');
  for(const [input,minutes] of [['12:00 AM',0],['12:00 PM',720],['9 AM',540],['17:30',1050],['25:00',null],['9:99 PM',null],['0 AM',null],['a meeting',null]]) assert.equal(c.parseClock(input),minutes);
  assert.equal(c.parseTimeRange('9:00 – 10:30 AM').end,630);
  assert.equal(c.parseTimeRange('13:00—14:00').end,840);
  assert.equal(c.parseTimeRange('17:00 - 15:00'),null);
  assert.equal(c.parseDuration('1 hr 30 min'),90);assert.equal(c.parseDuration('1.5 hours'),90);assert.equal(c.parseDuration('unknown'),null);
});

test('Pinned clock is undated; explicit datetime takes priority',()=>{
  const c=load(context({todayKey:()=> '2026-11-08'}),'  function parseDateTitle(title)','  function scrapeMessages(');
  const row={querySelector:()=>null,querySelectorAll:()=>[{textContent:'9:30 AM'}]};
  assert.equal(c.rowDate(row,'Pinned'),null);
  row.querySelector=()=>({getAttribute:()=> '2024-01-01T10:00:00Z'});
  assert.equal(c.rowDate(row,'Pinned').getFullYear(),2024);
});

test('State validation removes expired blocks and rejects malformed saved collections',()=>{
  const c=context({defaultState:()=>({tab:'timeline',todo:'self',selfSort:'custom',selfOrder:[],schedule:[],localBlocks:[],titleEdits:{}}),todayKey:()=> '2026-11-08'});
  load(c,'  function normalizeState(value)','  function readAccount()');
  const state=c.normalizeState({tab:'bad',selfOrder:null,titleEdits:[],localBlocks:[{date:'2026-11-01',start:1,end:2},{date:'2026-11-09',start:2,end:3},null]});
  assert.equal(state.tab,'timeline');assert.equal(state.selfOrder.length,0);assert.equal(state.localBlocks.length,1);assert.equal(state.localBlocks[0].date,'2026-11-09');
});

test('Cache accumulates virtualized rows without treating scrolling as deletion',()=>{
  let rows=[];const cache=new Map();
  const c=context({cache,scanTarget:null,messageList:()=>({querySelector:()=>({querySelectorAll:()=>rows})}),getComputedStyle:()=>({}),visible:()=>true,parseDateTitle:()=>null,rowDate:()=>null});
  load(c,'  function scrapeMessages(','  function parseClock(');
  const row={id:'',dataset:{convid:'A'},getAttribute:()=>'',querySelector:()=>null,querySelectorAll:()=>[]};rows=[row];c.scrapeMessages();rows=[];c.scrapeMessages();assert(cache.has('A'));
});

function scanContext({grow=false,abort=false,fail=false}={}) {
  let max=500,top=19,furthest=0,passes=0,rendered=0,pruned=0;
  const attributes=new Set();
  const scroller={isConnected:true,clientHeight:500,get scrollHeight(){return max+500},get scrollTop(){return top},set scrollTop(value){top=Math.min(max,value);furthest=Math.max(furthest,top)},setAttribute:name=>attributes.add(name),removeAttribute:name=>attributes.delete(name)};
  const cache=new Map([['stale',{id:'stale'}]]);
  const c=context({scanning:false,dataLoadingAllowed:()=>true,document:{hidden:false},interactionActive:()=>false,scannedScroller:null,scanGeneration:0,location:{pathname:'/mail/inbox'},storageKey:'account-A',messageList:()=>({querySelector:()=>scroller}),scanComplete:false,scanProgress:0,scanTarget:null,cache,
    updateSummary(){},render:()=>rendered++,pruneState:()=>pruned++,widgetContext:()=>true,mount(){},queueScan(){},delay:async()=>{},
    settleVirtualList:async()=>{passes++;if(grow && passes===2)max=1500;if(abort && passes===2)c.scanGeneration++;if(fail)throw Error('Synthetic failure');return true;},
    scrapeMessages:()=>{if(c.scanTarget)c.scanTarget.set(`at-${top}`,{id:`at-${top}`});},console:{warn(){}}});
  load(c,'  async function scanAllMessages(force)','  function updateSummary(');
  return {c,cache,attributes,get furthest(){return furthest},get top(){return top},get pruned(){return pruned},get rendered(){return rendered}};
}

test('Growing inbox is scanned to its new bottom and stale cache is reconciled',async()=>{
  const fixture=scanContext({grow:true});await fixture.c.scanAllMessages(true);
  assert.equal(fixture.furthest,1500);assert.equal(fixture.c.scanComplete,true);assert.equal(fixture.cache.has('stale'),false);assert.equal(fixture.top,19);assert.equal(fixture.pruned,1);assert.equal(fixture.attributes.size,0);
});
test('Cancelled scan preserves old cache and never prunes saved IDs',async()=>{
  const fixture=scanContext({abort:true});await fixture.c.scanAllMessages(true);
  assert.equal(fixture.c.scanComplete,false);assert(fixture.cache.has('stale'));assert.equal(fixture.pruned,0);assert.equal(fixture.c.scanning,false);assert.equal(fixture.attributes.size,0);
});
test('Scan failure cleans up overlay and scanning state',async()=>{
  const fixture=scanContext({fail:true});await fixture.c.scanAllMessages(true);
  assert.equal(fixture.c.scanning,false);assert.equal(fixture.attributes.size,0);assert(fixture.cache.has('stale'));
});

test('Recycled rows use safe navigation; matching rows use native click',()=>{
  let handler,clicked=false,assigned='';const element={dataset:{convid:'B'},isConnected:true,click:()=>clicked=true};
  const c=context({root:{addEventListener:(name,fn)=>{if(name==='click')handler=fn}},cache:new Map([['A',{row:element}]]),location:{assign:url=>assigned=url}});
  load(c,'  function bind()','  const resizeObserver');c.bind();
  const target={dataset:{openMessage:'A'}};handler({target:{closest:()=>target}});
  assert.equal(clicked,false);assert.equal(assigned,'/mail/inbox/id/A');element.dataset.convid='A';handler({target:{closest:()=>target}});assert.equal(clicked,true);
});

test('Focus timer accounts for reload/sleep elapsed time and stops at completion',()=>{
  const focus={part:0,elapsed:60,running:true,finished:false,last:10000};
  const c=context({focus,PARTS:[['focus',25]],focusTick:1,clearInterval(){}});
  load(c,'  function advanceFocus(','  function updateFocusDisplay()');
  assert.equal(c.advanceFocus(40000),false);assert.equal(focus.elapsed,90);
  assert.equal(c.advanceFocus(9999999),true);assert.equal(focus.elapsed,1500);assert.equal(focus.running,false);assert.equal(focus.finished,true);
});

test('Focus display updates nodes without replacing widget controls',()=>{
  const clock={textContent:'25:00'},writes=[];
  const c=context({root:{querySelector:selector=>selector.includes(' b')?clock:{style:{setProperty:(...args)=>writes.push(args)}}},state:{collapsed:false,tab:'focus'},document:{hidden:false},PARTS:[['focus',25]],focus:{part:0,elapsed:2}});
  load(c,'  function updateFocusDisplay()','  function startFocus(');c.updateFocusDisplay();assert.equal(clock.textContent,'24:58');assert.equal(writes.length,1);
});

test('Account change switches storage without carrying private state',()=>{
  const stored=new Map(),focus={part:0,elapsed:5,running:false};let account='second@example.com';
  const c=context({readAccount:()=>account,flushSave(){},accountId:'first@example.com',storageKey:'first',state:{localBlocks:[{title:'Private'}]},focus,location:{host:'outlook.office.com'},STORE:'erinInboxWidgets.v1',GM_getValue:(k,d)=>stored.get(k)??d,GM_setValue:(k,v)=>stored.set(k,v),normalizeState:v=>v||{localBlocks:[]},PARTS:[1,2,3,4],activeDrag:null,pendingBlock:{title:'private'},renderPending:false,resizeObserver:null,scanGeneration:0,cache:new Map(),root:null,scannedScroller:null,scanComplete:false,focusTick:0,startFocus(){},clearInterval(){},save(){}});
  load(c,'  function syncAccount()','  function flushSave()');c.syncAccount();assert.equal(c.state.localBlocks.length,0);assert.equal(c.pendingBlock,null);assert(c.storageKey.endsWith(':second@example.com'));
});

test('Closed flyout is inert; opening synchronizes accessibility state',()=>{
  let open=false;const attrs={};const panel={classList:{contains:()=>open,toggle:(_,v)=>open=v},style:{},setAttribute:(k,v)=>attrs[k]=v,matches:()=>false};
  const btn={getBoundingClientRect:()=>({right:40,top:4,height:40}),getAttribute:()=>'',setAttribute:(k,v)=>attrs[k]=v,matches:()=>false};
  const c=context({document:{documentElement:{dataset:{}},getElementById:id=>id==='erin-apps-btn'?btn:panel},showTip(){},hideTip(){}});
  load(c,'    function setOpen(open)','    // Subject bar',stylerSource);c.setOpen(true);assert.equal(panel.inert,false);assert.equal(attrs['aria-expanded'],'true');c.setOpen(false);assert.equal(panel.inert,true);assert.equal(attrs['aria-hidden'],'true');
});

test('App lookup compares labels as data, including quotes and brackets',()=>{
  const label='Mail "Ops"]',button={getAttribute:()=>label};
  const c=context({document:{querySelectorAll:()=>[button]}});
  vm.runInContext(section(stylerSource,'    const railBtn =','    let tip;')+'\nthis.lookup=railBtn;',c);
  assert.equal(c.lookup(label),button);
});

test('Debounce queues one scrape for a burst of mutations',()=>{
  let scheduled=0,mounted=0,callback;
  const c=context({renderTimer:0,scanning:false,document:{hidden:false},setTimeout:fn=>{scheduled++;callback=fn;return 1},mount:()=>mounted++});
  load(c,'  function scheduleRender()','  function queueScan(');for(let i=0;i<100;i++)c.scheduleRender();assert.equal(scheduled,1);callback();assert.equal(mounted,1);
});

test('Background rendering defers during editing and preserves scroll after editing',()=>{
  let rendered=0,interacting=true;const panel={scrollTop:75},flow={scrollTop:120};
  const c=context({root:{dataset:{},contains:()=>false,querySelector:selector=>selector==='.eiw-panel'?panel:flow},interactionActive:()=>interacting,renderPending:false,document:{activeElement:null},state:{collapsed:false,tab:'todo'},cache:new Map(),setMarkup:()=>{rendered++;panel.scrollTop=0;flow.scrollTop=0},panelHTML:()=>'',updateSummary(){},queueGeometry(){}});
  load(c,'  function render(force','  function queueGeometry()');c.render();assert.equal(rendered,0);assert.equal(c.renderPending,true);interacting=false;c.render();assert.equal(rendered,1);assert.equal(panel.scrollTop,75);assert.equal(flow.scrollTop,120);
});

test('Schedule drag cancels cleanly and clamps blocks to the displayed day',()=>{
  const events=new Map(),listeners=new Map(),selection={hidden:true,style:{}};
  const strip={dataset:{start:'480',end:'1020'},getBoundingClientRect:()=>({left:0,width:540}),querySelector:()=>selection,setPointerCapture(){},addEventListener:(k,v)=>listeners.set(k,v),removeEventListener:k=>listeners.delete(k)};
  const c=context({root:{addEventListener:(k,v)=>events.set(k,v)},activeDrag:null,renderPending:false,pendingBlock:null,render(){}});
  load(c,'  function bind()','  const resizeObserver');c.bind();
  const down=x=>events.get('pointerdown')({target:{closest:selector=>selector==='[data-schedule-strip]'?strip:null},button:0,clientX:x,pointerId:1});
  down(100);assert.equal(selection.hidden,false);listeners.get('pointercancel')();assert.equal(c.activeDrag,null);assert.equal(selection.hidden,true);assert.equal(c.pendingBlock,null);assert.equal(listeners.size,0);
  down(540);listeners.get('pointerup')({clientX:540});assert.equal(c.pendingBlock.start,1005);assert.equal(c.pendingBlock.end,1020);assert.equal(c.activeDrag,null);
});

test('Storage pruning retains filtered-out edits until their retention period expires',()=>{
  const state={localBlocks:[],selfOrder:['recent','old'],titleEdits:{recent:'Keep',old:'Remove'},idLastSeen:{recent:'2026-11-01',old:'2020-01-01'}};
  const TestDate = class extends Date { constructor(...args) { super(...(args.length ? args : [2026,10,8])); } };
  const c=context({Date:TestDate,state,cache:new Map(),todayKey:()=> '2026-11-08',dateKey:d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,save(){}});
  load(c,'  function pruneState()','  const interactionActive');c.pruneState();assert.equal(state.titleEdits.recent,'Keep');assert.equal(state.titleEdits.old,undefined);assert.deepEqual(state.selfOrder,['recent']);
});

test('Styler ignores widget mutations and batches sign-in inspection by scope',()=>{
  let observerCallback,builds=0,checks=0;
  const c=context({Node:{ELEMENT_NODE:1},MutationObserver:class{constructor(fn){observerCallback=fn}},relevantAttrTarget:()=>false,scheduleBuild:()=>builds++,clickSessionExpiredSignIn:()=>checks++,checkAddedMenus(){},checkAddedDropTargets(){},removeDeleteMenuItems(){},hideNeverDeleteRetention(){}});
  load(c,'    const ownedUI =','    mainObserver.observe',stylerSource);
  const owned={nodeType:1,closest:()=>({})};observerCallback([{type:'childList',target:owned}]);assert.equal(builds,0);assert.equal(checks,0);
  const scope={},node={nodeType:1,closest:()=>scope,matches:()=>false,querySelector:()=>null},target={nodeType:1,closest:()=>null};
  // Owned check sees closest() too: return a group only for the exact group selector.
  node.closest=selector=>selector==='[role="group"]'?scope:null;
  observerCallback([{type:'childList',target,addedNodes:[node,node],removedNodes:[]}]);assert.equal(checks,1);assert.equal(builds,0);
});

test('Sign-in banner auto-click is limited and selects the sign-in button',()=>{
  let clicks=0;const close={getAttribute:()=> 'Dismiss'},signIn={getAttribute:()=> 'Sign in',click:()=>clicks++};const banner={textContent:'You need to sign in.',querySelectorAll:()=>[close,signIn]};
  const c=context({document:{},within:()=>[banner]});load(c,'    const signInClicks =','    function hideSenderCategories()',stylerSource);
  c.clickSessionExpiredSignIn();c.clickSessionExpiredSignIn();assert.equal(clicks,1);
});
