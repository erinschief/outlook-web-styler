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
  await wait(250); w.stopAutoScan();
  await test('Both full scripts start without exceptions', () => assert(w.root && window.styler && !errors.length, errors.join('; ') || 'Missing UI'));
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
  await test('Navigation aborts a scan without committing partial data', async () => {
    w.stopAutoScan();const pending=w.scanAllMessages(true);await wait(100);window.testLocation.pathname='/mail/sentitems';await pending;
    assert(!w.complete && !w.scanning,'Aborted scan reports complete');window.testLocation.pathname='/mail/inbox';
    assert(w.cache.size===0,'Old inbox data retained outside inbox');
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
