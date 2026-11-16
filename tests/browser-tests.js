/* Executed only by the generated local browser fixture. */
(async () => {
  const output = document.getElementById('results'), results = [];
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (condition, message) => { if (!condition) throw new Error(message); };
  const test = async (name, fn) => {
    try { await fn(); results.push(`PASS ${name}`); }
    catch (error) { results.push(`FAIL ${name}: ${error.message}`); }
    output.textContent = results.join('\n');
  };
  const w = window.widgets;
  const stream = document.querySelector('[data-testid="virtuoso-item-list"]');
  const row = (id, subject = id) => {
    const el = document.createElement('div'); el.className = 'test-row'; el.setAttribute('role', 'option'); el.dataset.convid = id; el.setAttribute('aria-label', 'Unread');
    const title = document.createElement('span'); title.className = 'TtcXM'; title.textContent = subject; el.appendChild(title);
    const tag = document.createElement('span'); tag.className = 'fui-Tag__primaryText'; tag.textContent = 'Self: Review'; el.appendChild(tag);
    stream.appendChild(el); return el;
  };
  await wait(1800); w.stopAutoScan();
  await test('Both full scripts start without exceptions', () => assert(w.root && window.styler && !errors.length, errors.join('; ') || 'Missing UI'));
  await test('Initial load completes one full Inbox sweep without using My Day', () => {
    assert(!document.querySelector('[data-app-section="DrawerPane"]'),'My Day opened during the startup Inbox sweep');
    assert(w.complete,'Startup sweep did not complete');
    assert(!w.state.schedule.length,'Schedule data was collected before opening Schedule');
  });
  await test('Opening Schedule refreshes its cache and closes script-opened My Day', async () => {
    w.root.querySelector('[data-tab="schedule"]').click();await wait(700);
    assert(w.state.schedule[0]?.title==='Initial meeting','Schedule tab did not refresh My Day data');
    assert(!document.querySelector('[data-app-section="DrawerPane"]'),'My Day stayed open after the Schedule refresh');
  });
  await test('Unread and To-do panels do not show refresh filler text', () => {
    w.state.tab='timeline';w.render();
    assert(!w.root.querySelector('.eiw-panel > .eiw-note'),'Unread filler text is still visible');
    w.state.tab='todo';w.render();
    assert(!w.root.querySelector('.eiw-panel > .eiw-note'),'To-do filler text is still visible');
  });
  await test('Oldest timeline date uses the bucket-label typography', () => {
    w.cache.set('label-color',{id:'label-color',subject:'Color check',date:new Date(),unread:true,categories:[]});
    w.state.tab='timeline';w.render();
    const oldest=w.root.querySelector('.eiw-oldest-bound'),bucket=w.root.querySelector('.eiw-bname');
    assert(oldest && bucket,'Timeline labels missing');
    assert(getComputedStyle(oldest).color===getComputedStyle(bucket).color,'Oldest date color drifted from bucket labels');
    assert(getComputedStyle(oldest).fontFamily===getComputedStyle(bucket).fontFamily,'Oldest date font family drifted from bucket labels');
    assert(getComputedStyle(oldest).fontWeight===getComputedStyle(bucket).fontWeight,'Oldest date font weight drifted from bucket labels');
    assert(getComputedStyle(oldest).opacity==='1' && getComputedStyle(oldest).writingMode==='horizontal-tb','Oldest date is still synthesized as translucent vertical text');
  });
  await test('Read micro ticks share equal slots with unread dots', () => {
    const saved=[...w.cache],now=new Date();w.cache.clear();
    for(let i=0;i<4;i++)w.cache.set(`slot-${i}`,{id:`slot-${i}`,subject:`Slot ${i}`,date:new Date(now.getTime()-i*60000),unread:i%2===0,categories:[]});
    w.state.tab='timeline';w.render();
    const marks=[...w.root.querySelector('.eiw-bucket').querySelectorAll('.eiw-dot')];
    assert(marks.map(mark=>parseFloat(mark.style.left)).join(',')==='12.5,37.5,62.5,87.5','Timeline marks are not evenly spaced');
    const ticks=marks.filter(mark=>mark.classList.contains('read'));
    assert(ticks.length===2 && ticks.every(mark=>getComputedStyle(mark).width==='1px' && getComputedStyle(mark).height==='9px'),'Read messages did not render as micro ticks');
    assert([...w.root.querySelectorAll('.eiw-bucket')].every(bucket=>bucket.getBoundingClientRect().width>=36),'Timeline bucket collapsed under its label');
    assert([...w.root.querySelectorAll('.eiw-bname')].map(label=>label.textContent).join(',')==='Today','Empty relative-date buckets were rendered');
    w.cache.clear();saved.forEach(([key,value])=>w.cache.set(key,value));w.render();
  });
  await test('Promotions tag stays hidden when Outlook recycles a message row', async () => {
    const label=document.querySelector('[data-promotion-label]'),host=label.parentElement;
    assert(host.hasAttribute('data-erin-promotions'),'Initial Promotions tag was not hidden');
    label.textContent='Other';await wait(50);
    assert(!host.hasAttribute('data-erin-promotions'),'Recycled non-Promotions tag stayed hidden');
    label.textContent='Promotions';await wait(50);
    assert(host.hasAttribute('data-erin-promotions'),'Recycled Promotions tag was not hidden immediately');
  });
  await test('Visible inbox changes update in the background without scrolling or a full scan', async () => {
    const scroller=stream.parentElement, before=scroller.scrollTop;
    const el=row('live','Live arrival');await wait(200);
    assert(w.cache.get('live')?.unread,'New arrival not collected automatically');
    el.setAttribute('aria-label','Read');el.querySelector('.fui-Tag__primaryText').textContent='Waiting: Reply';await wait(200);
    assert(!w.cache.get('live').unread && w.cache.get('live').categories.includes('Waiting: Reply'),'Row change not reflected');
    assert(scroller.scrollTop===before && !w.scanning,'Background update moved the inbox or started a scan');
  });
  await test('Folder counts remain live expanded, while editing, and collapsed', async () => {
    const folder=document.createElement('div');folder.setAttribute('role','treeitem');folder.title='Inbox - 20 items (7 unread)';folder.setAttribute('aria-selected','true');document.getElementById('fixture').appendChild(folder);
    await wait(50);
    const count=()=>w.root.querySelector('.eiw-summary b').textContent;
    assert(count()==='7','Expanded header used cached conversation count');
    w.state.tab='todo';w.state.todo='waiting';w.render();
    const title=w.root.querySelector('[data-edit-title]');title.focus();title.textContent='Unsaved live edit';
    folder.title='Inbox - 20 items (6 unread)';await wait(200);
    assert(count()==='6' && title.isConnected && title.textContent==='Unsaved live edit','Count update interrupted editing');
    title.blur();await wait(20);
    w.root.querySelector('[data-action="collapse"]').click();folder.title='Inbox - 20 items (5 unread)';await wait(50);
    assert(count()==='5','Collapsed count stopped updating');
    w.root.querySelector('[data-action="collapse"]').click();assert(count()==='5','Expanding changed count source');
    w.state.todo='self';w.state.tab='timeline';w.render();folder.remove();await wait(50);
    assert(count()==='5' && w.root.querySelector('.eiw-summary').title.includes('Last confirmed'),'Lost the Inbox baseline when My Day hid folder totals');
    w.root.querySelector('[data-action="collapse"]').click();
    const live=document.querySelector('[data-convid="live"]');live.setAttribute('aria-label','Unread');await wait(200);
    assert(count()==='6','Collapsed tracked count missed a visible unread change');
    live.setAttribute('aria-label','Read');await wait(200);assert(count()==='5','Tracked count did not restore after marking read');
    w.root.querySelector('[data-action="collapse"]').click();
  });
  await test('My Day schedule changes update without a full inbox refresh', async () => {
    const drawer=document.querySelector('[data-app-section="DrawerPane"]'),content=drawer.querySelector('[data-test-schedule]');
    await wait(200);assert(w.state.schedule[0]?.title==='Initial meeting','Calendar not collected automatically');
    content.innerText='Today\n10:00 AM\nUpdated meeting\n30 min\nTomorrow';await wait(200);
    assert(w.state.schedule[0]?.title==='Updated meeting' && w.state.schedule[0].start===600,'Calendar change remained stale');drawer.remove();
  });
  await test('Calendar buckets respect midnight, month and year', () => {
    const now = new Date(2026, 10, 8, 9);
    assert(w.bucketFor({date:new Date(2026,10,7,23,59)},now) === 1, 'Yesterday is Today');
    assert(w.bucketFor({date:new Date(2026,9,20)},now) === 4, 'Last month incorrect');
    assert(w.bucketFor({date:new Date(2025,11,20)},new Date(2026,0,10)) === 4, 'Year boundary incorrect');
  });
  await test('Clock and duration parsers support 12h, 24h and ranges', () => {
    assert(w.parseClock('12:00 AM')===0 && w.parseClock('12:00 PM')===720 && w.parseClock('17:30')===1050,'Clock parsing');
    assert(w.parseClock('25:00')===null && w.parseClock('4:99 PM')===null,'Invalid clock accepted');
    assert(w.parseTimeRange('9:00 – 10:30 AM').end===630,'Range parsing');
    assert(w.parseDuration('1 hr 30 min')===90 && w.parseDuration('1.5 hours')===90,'Duration parsing');
  });
  await test('Pinned clock does not invent a date', () => {
    const el=document.createElement('div');el.innerHTML='<span>9:30 AM</span>';
    assert(w.rowDate(el,'Pinned')===null,'Pinned date guessed');
  });
  await test('Recycled rows cannot open another message', () => {
    const el=row('A','Alpha');w.scrapeMessages();el.dataset.convid='B';w.scrapeMessages();w.state.tab='todo';w.render();
    let clicked=false;el.onclick=()=>clicked=true;
    const open=[...w.root.querySelectorAll('[data-open-message]')].find(b=>b.dataset.openMessage==='A');assert(open,'Missing A card');open.click();
    assert(!clicked && window.assignedURL.endsWith('/A'),'Recycled B was clicked');
  });
  await test('Background renders preserve active title edits', () => {
    const title=w.root.querySelector('[data-edit-title]');title.focus();title.textContent='Unsaved typing';w.render();
    assert(title.isConnected && title.textContent==='Unsaved typing','Editor replaced');title.blur();
  });
  await wait(20);
  await test('Focus tick updates in place and state persists', async () => {
    w.state.tab='focus';w.render();w.startFocus();const button=w.root.querySelector('[data-action="focus-toggle"]');
    await wait(1150);assert(button.isConnected,'Timer replaced controls');w.flushSave();
    const saved=stored.get('erinInboxWidgets.v1:outlook.office.com:erin@example.com');assert(saved.focus.running && saved.focus.elapsed>=1,'Focus not stored');
    button.click();
  });
  await test('Completed sweep reconciles removals and growing scroll height', async () => {
    w.stopAutoScan();w.cache.set('removed',{id:'removed',categories:[],unread:true});
    stream.replaceChildren();for(let i=0;i<8;i++)row(`scan-${i}`);
    const scroller=stream.parentElement;let grown=false;
    const onScroll=()=>{if(!grown && scroller.scrollTop>100){grown=true;for(let i=8;i<16;i++)row(`scan-${i}`)}};
    scroller.addEventListener('scroll',onScroll);scroller.scrollTop=17;
    await w.scanAllMessages(true);scroller.removeEventListener('scroll',onScroll);
    assert(grown && w.complete && w.cache.size===16 && !w.cache.has('removed'),'Incomplete reconciliation');
    assert(Math.abs(scroller.scrollTop-17)<2,'Scroll position not restored');
  });
  await test('Full sweep collects asynchronously recycled virtual rows', async () => {
    const scroller=stream.parentElement;stream.replaceChildren();stream.style.height='2400px';stream.style.position='relative';
    let timer=0;
    const paint=()=>{
      stream.replaceChildren();const start=Math.max(0,Math.floor(scroller.scrollTop/60)-1);
      for(let i=start;i<Math.min(40,start+7);i++){const el=row(`virtual-${i}`);el.style.position='absolute';el.style.top=`${i*60}px`;}
    };
    const onScroll=()=>{clearTimeout(timer);timer=setTimeout(paint,20)};
    scroller.scrollTop=0;paint();scroller.addEventListener('scroll',onScroll);
    try {
      await w.scanAllMessages(true);
      assert(w.complete && w.cache.size===40 && w.cache.has('virtual-39'),'Virtual rows skipped or stale records retained');
    } finally {
      scroller.removeEventListener('scroll',onScroll);clearTimeout(timer);stream.style.height='';stream.style.position='';
      for(const el of stream.children){el.style.position='';el.style.top='';}
    }
  });
  await test('Navigation aborts a scan without committing partial data', async () => {
    w.stopAutoScan();const before=[...w.cache.keys()].join('|');const pending=w.scanAllMessages(true);await wait(100);window.testLocation.pathname='/mail/sentitems';await pending;
    assert(!w.complete && !w.scanning,'Aborted scan reports complete');window.testLocation.pathname='/mail/inbox';
    assert([...w.cache.keys()].join('|')===before,'Navigation discarded the previous inbox snapshot');w.mount();
  });
  await test('Apps menu handles quotes in shared labels and keyboard opening', () => {
    const panel=document.getElementById('erin-apps-panel'),btn=document.getElementById('erin-apps-btn');
    assert(panel.inert,'Closed panel is keyboard accessible');
    btn.focus();btn.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}));
    assert(!panel.inert && panel.contains(document.activeElement),'Keyboard did not open menu');
    assert(panel.querySelectorAll('[data-for="Mail"]').length===1 && panel.querySelector('[data-for="shared-mailbox"]'),'App identity collision');
    panel.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));assert(panel.inert && document.activeElement===btn,'Escape did not return focus');
  });
  await test('Widget timer mutations do not rebuild styler', async () => {
    w.stopAutoScan();w.state.tab='focus';w.render();await wait(300);const before=buildCount;w.startFocus();await wait(1200);
    assert(buildCount-before<=1,`Unexpected shell builds: ${buildCount-before}`);w.root.querySelector('[data-action="focus-toggle"]').click();
  });
  await test('Idle styler settles instead of observing its own style writes', async () => {
    await wait(200);const before=buildCount;await wait(400);assert(buildCount-before<=1,`Idle builds: ${buildCount-before}`);
  });
  await test('Account change isolates saved calendar and edits', () => {
    w.state.localBlocks=[{id:'private',date:new Date().toLocaleDateString('en-CA'),start:500,end:600,title:'Private'}];w.flushSave();
    document.getElementById('owa-me-control-container').title='another@example.com';w.syncAccount();
    assert(w.state.localBlocks.length===0 && Object.keys(w.state.titleEdits).length===0,'Account data leaked');w.mount();
  });
  await test('No uncaught browser errors', () => assert(!errors.length,errors.join('; ')));
  w.stopAutoScan();window.testResults={passed:results.filter(x=>x.startsWith('PASS')).length,failed:results.filter(x=>x.startsWith('FAIL')).length,results};
  output.textContent=results.join('\n')+`\n\n${testResults.passed} passed, ${testResults.failed} failed`;document.title=`${testResults.failed?'FAIL':'PASS'} Outlook regressions`;output.dataset.complete='true';
})();

