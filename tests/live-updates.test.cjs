const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '..', 'outlook-inbox-widgets.user.js'), 'utf8');
function load(values, from, to) {
  const start = source.indexOf(from), end = source.indexOf(to, start);
  assert(start >= 0 && end > start);
  const c = vm.createContext({Date, Map, Set, ...values});
  vm.runInContext(source.slice(start, end), c);
  return c;
}

test('Header follows folder counts in both states, falls back honestly, and avoids idle DOM writes', () => {
  let folder = {unread:7,total:20}, writes=0;
  const node = () => ({value:'',get textContent(){return this.value},set textContent(v){writes++;this.value=v}});
  const values=[node(),node()], label=node();
  const summary={querySelectorAll:()=>values,querySelector:()=>label};
  const c=load({root:{querySelector:()=>summary},state:{collapsed:false},cache:new Map([['a',{unread:true}]]),inboxFolderCounts:()=>folder,scanning:false},'  function updateSummary(','  function advanceFocus(');
  c.updateSummary();assert.deepEqual(values.map(x=>x.value),['7','20']);assert.equal(label.value,' in inbox');
  const before=writes;c.updateSummary();assert.equal(writes,before);
  c.state.collapsed=true;c.updateSummary();assert.equal(values[0].value,'7');
  folder={unread:0,total:19};c.updateSummary();assert.equal(values[0].value,'0');
  c.state.collapsed=false;folder=null;c.updateSummary();assert.equal(values[0].value,'1');
  assert.equal(label.value,' cached conversations');assert.match(summary.title,/unavailable/);
});

test('Inbox selection prevents a shared mailbox or Favorites duplicate from choosing wrong counts', () => {
  let folders=[];
  const folder=(title,selected=false)=>({getAttribute:key=>key==='title'?title:selected?'true':'false'});
  const c=load({document:{querySelectorAll:()=>folders}},'  function inboxFolderCounts()','  const visible =');
  folders=[folder('Inbox - 4 items (2 unread)'),folder('Inbox – 1,234 items (50 unread)',true)];
  assert.equal(c.inboxFolderCounts().total,1234);
  folders=[folder('Inbox - 4 items (2 unread)'),folder('Inbox - 5 items (3 unread)')];assert.equal(c.inboxFolderCounts(),null);
  folders=[folder('Inbox - 0 items (0 unread)'),folder('Inbox - 0 items (0 unread)')];assert.equal(c.inboxFolderCounts().unread,0);
  folders=[folder('Sent - 3 items (0 unread)')];assert.equal(c.inboxFolderCounts(),null);
});

test('Mutation batches request a real scrape and do not queue work in a hidden tab or during scans', () => {
  let callback, calls=0, mode='unset';
  const c=load({renderTimer:0,scanning:false,document:{hidden:false},setTimeout:fn=>{calls++;callback=fn;return 1},mount:(scrape=true)=>{mode=scrape}},'  function scheduleRender()','  function queueScan(');
  for(let i=0;i<100;i++)c.scheduleRender();assert.equal(calls,1);callback();assert.equal(mode,true);
  c.document.hidden=true;c.scheduleRender();assert.equal(calls,1);
  c.document.hidden=false;c.scanning=true;c.scheduleRender();assert.equal(calls,1);
});

test('Visible unread and category changes update cache; recycled and removed rows retain off-screen records', () => {
  let rows=[], unread=true, category='Self: Review';
  const title={textContent:'Hello'}, sender={textContent:'Sender'};
  const tag={textContent:category,closest(){return this}};
  const row={id:'',dataset:{convid:'A'},getAttribute:()=>unread?'Unread':'Read',
    querySelector:selector=>selector==='.TtcXM'?title:selector==='span[title*="@"]'?sender:null,
    querySelectorAll:selector=>selector.includes('category')?[tag]:[]};
  const cache=new Map();
  const c=load({cache,scanTarget:null,messageList:()=>({querySelector:()=>({querySelectorAll:()=>rows})}),getComputedStyle:()=>({}),visible:()=>true,parseDateTitle:()=>null,rowDate:()=>null},'  function scrapeMessages(','  function parseClock(');
  rows=[row];assert.equal(c.scrapeMessages(),true);assert.equal(cache.get('A').unread,true);
  assert.equal(c.scrapeMessages(),false);
  unread=false;tag.textContent='Waiting: Reply';assert.equal(c.scrapeMessages(),true);
  assert.equal(cache.get('A').unread,false);assert.equal(cache.get('A').categories[0],'Waiting: Reply');
  row.dataset.convid='B';unread=true;c.scrapeMessages();assert.equal(cache.size,2);
  rows=[];c.scrapeMessages();assert.equal(cache.size,2);
});

function settleFixture() {
  let now=0, nextId=0, signature='old', busy=false, valid=true, disconnected=0, notify;
  const tasks=new Map();
  const c=load({Date:{now:()=>now},listSignature:()=>signature,
    setTimeout:(fn,ms)=>{const id=++nextId;tasks.set(id,{fn,at:now+ms});return id},clearTimeout:id=>tasks.delete(id),
    MutationObserver:class{constructor(fn){notify=fn}observe(){}disconnect(){disconnected++}}},
    '  const scanSignature =','  async function scanAllMessages(');
  const pending=c.settleVirtualList({scrollHeight:500,querySelector:()=>busy?{}:null},()=>valid,'500:old');
  const advance=until=>{for(;;){const entry=[...tasks].sort((a,b)=>a[1].at-b[1].at)[0];if(!entry||entry[1].at>until)break;tasks.delete(entry[0]);now=entry[1].at;entry[1].fn()}now=until};
  return {pending,advance,change(value){signature=value;notify()},busy(value){busy=value;notify()},cancel(){valid=false},get disconnected(){return disconnected},get timers(){return tasks.size}};
}

test('Adaptive settling waits for changed rows, then finishes before the old 180ms minimum', async () => {
  const f=settleFixture();f.advance(24);assert.equal(f.disconnected,0);
  f.change('new');f.advance(71);assert.equal(f.disconnected,0);f.advance(72);
  assert.equal(await f.pending,true);assert.equal(f.disconnected,1);assert.equal(f.timers,0);
});

test('Unchanged rows get a conservative wait; busy or cancelled lists never commit', async () => {
  const unchanged=settleFixture();unchanged.advance(144);assert.equal(unchanged.disconnected,0);unchanged.advance(168);assert.equal(await unchanged.pending,true);
  const busy=settleFixture();busy.busy(true);busy.advance(1512);assert.equal(await busy.pending,false);assert.equal(busy.timers,0);
  const cancelled=settleFixture();cancelled.cancel();cancelled.advance(24);assert.equal(await cancelled.pending,false);assert.equal(cancelled.disconnected,1);
});
