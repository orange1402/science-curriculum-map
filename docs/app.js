'use strict';
(() => {
  const publishedData = window.CURRICULUM_DATA;
  const storageKey = 'science-curriculum-excel-v1';
  let data = publishedData, importedName = '', pendingImport = null, importTicket = 0;
  try { const saved=JSON.parse(localStorage.getItem(storageKey)||'null');if(saved?.data?.chains?.length){data=saved.data;importedName=saved.name;} } catch { /* Published data remains usable when local storage is unavailable. */ }
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const normal = s => String(s ?? '').normalize('NFKC').toLocaleLowerCase().replace(/\s+/g,'');
  let nodeMap,linkedCodes,expsByCode,searchIndex;
  function rebuildIndexes(){
    nodeMap=new Map(data.nodes.map(n=>[n.code,n]));linkedCodes=new Set(data.links.flatMap(l=>[l.a,l.b]));
    expsByCode=new Map(data.nodes.map(n=>[n.code,[]]));data.explanations.forEach(e=>e.codes.forEach(c=>expsByCode.get(c)?.push(e)));
    searchIndex=new Map(data.nodes.map(n=>[n.code,normal([n.code,n.originalCode,n.text,n.topic,n.subtopic,n.subject,...expsByCode.get(n.code).map(e=>e.text)].join(' '))]));
  }
  rebuildIndexes();
  const icons = {
    search:'<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 4.5 4.5"/>',
    grid:'<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    route:'<circle cx="5" cy="6" r="2.5"/><circle cx="19" cy="18" r="2.5"/><path d="M7.5 6H16a4 4 0 0 1 0 8H8a4 4 0 0 0-4 4h12.5"/>',
    reset:'<path d="M4 9a8 8 0 1 1 1 9M4 3v6h6"/>',
    info:'<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7v1"/>',
    close:'<path d="m6 6 12 12M18 6 6 18"/>',
    arrow:'<path d="M4 12h15m-5-5 5 5-5 5"/>',
    book:'<path d="M12 5C8 2 4 3 3 4v16c3-2 6-1 9 1 3-2 6-3 9-1V4c-3-2-6-1-9 1Zm0 0v16"/>',
    link:'<path d="m9 15 6-6M8 9l2-2a5 5 0 0 1 7 7l-2 2M16 15l-2 2a5 5 0 0 1-7-7l2-2"/>',
    chevron:'<path d="m9 5 7 7-7 7"/>'
  };
  const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] || icons.book}</svg>`;
  function hydrateIcons(root = document) { $$('[data-icon]', root).forEach(el => el.innerHTML = icon(el.dataset.icon)); }
  const stages = [
    {code:'Ⅱ',name:'國小中年級',grade:'三至四年級',num:'Ⅱ'},
    {code:'Ⅲ',name:'國小高年級',grade:'五至六年級',num:'Ⅲ'},
    {code:'Ⅳ',name:'國中',grade:'七至九年級',num:'Ⅳ'},
    {code:'Ⅴc',name:'高中必修',grade:'第五學習階段 · 必修',num:'Ⅴc'},
    {code:'Ⅴa',name:'高中選修',grade:'第五學習階段 · 加深加廣',num:'Ⅴa'}
  ];
  const subjectOptions = [
    ['全部科目','#75846b'],['生物','#71945b'],['物理','#6f97a8'],['化學','#b49b6c'],['地球科學','#9a88a6'],['跨科','#719c92']
  ];
  const state = {chain:'U01',subject:'全部科目',cross:true,showEmpty:true,query:'',librarySubject:'',libraryStage:'',connectedOnly:false,page:1};
  let lastFocus = null, drawerFocus = null, toastTimer, resizeObserver;
  const currentChain = () => data.chains.find(c => c.id === state.chain);
  const activeLinks = () => data.links.filter(l => l.chain === state.chain && (state.cross || !['跨科整合','同階段概念整合'].includes(l.type)));
  const clip = (s, n = 88) => s.length > n ? s.slice(0, n) + '…' : s;
  function setHash(node) {
    const hash = '#' + state.chain + (node ? '/node=' + encodeURIComponent(node) : '');
    history.replaceState(null,'',hash);
  }
  function renderNavigation() {
    $('#subjects').innerHTML = subjectOptions.map(([name,color],i) => `<button class="subject-button ${state.subject === name ? 'active' : ''}" data-subject="${esc(name)}" aria-pressed="${state.subject === name}">${i ? `<span class="subject-dot" style="--subject:${color}"></span>` : icon('grid')}<span>${name}</span></button>`).join('');
    const term=normal($('#unit-search').value);
    const visible = data.chains.filter(c => (state.subject === '全部科目' || (state.subject === '跨科' ? c.subject.startsWith('跨科') : c.subject === state.subject)) && normal(c.id+' '+c.name).includes(term));
    $('#chain-count').textContent = String(visible.length).padStart(2,'0');
    $('#chains').innerHTML = visible.map(c => `<button class="chain-button ${c.id === state.chain ? 'active' : ''}" data-chain="${c.id}" aria-current="${c.id === state.chain ? 'true' : 'false'}">${esc(c.name)}<small>${c.basis === '使用者指定' ? 'DIRECT CONNECTION · 直接關係' : 'ANALYSIS · 分析建議'}</small></button>`).join('');
  }
  function nodeCard(n, chain, i) {
    const explanations = expsByCode.get(n.code).filter(e => e.kind === '學習內容說明');
    const focus = chain.focus[n.code];
    const focused = focus ? explanations.find(e => e.seq === focus) : null;
    const shared=expsByCode.get(n.code).filter(e=>e.kind==='未編號說明（區塊共用）');
    const pages=[...new Set([n.pdf,...expsByCode.get(n.code).map(e=>e.pdf)].flatMap(s=>String(s).split('、')))];
    return `<article class="node-card ${chain.basis === '使用者指定' ? 'confirmed' : ''}" data-code="${esc(n.code)}" aria-label="${esc(n.code)} 學習內容與說明"><h3 class="node-code">${esc(n.code)}</h3><section class="card-section"><h4>學習內容</h4><p class="card-content">${esc(n.text)}</p></section><section class="card-section card-explanations"><h4>學習內容說明 <small>${explanations.length} 則</small></h4>${explanations.map(e=>`<div class="inline-explanation ${e.seq===focus?'focused':''}"><div class="inline-exp-label"><span>${esc(e.seq)}</span>${e.seq===focus?'<small>本連結聚焦</small>':''}</div><p>${esc(e.text)}</p></div>`).join('')}${shared.map(e=>`<div class="inline-explanation shared"><div class="inline-exp-label"><span>區塊共用說明</span></div><p>${esc(e.text)}</p></div>`).join('')}${!explanations.length&&!shared.length?'<p class="card-missing">未匹配獨立說明，需回查來源。</p>':''}</section><div class="node-card-footer"><div class="card-source-links">${pages.filter(p=>/^\d+$/.test(p)).map(p=>`<a href="assets/curriculum.pdf#page=${p}" target="_blank" rel="noopener">PDF ${p} 頁 ↗</a>`).join('')}</div><button class="text-button" data-node="${esc(n.code)}" aria-label="${esc(n.code)} 來源與相關連結">來源與連結 ${icon('chevron')}</button></div></article>`;
  }
  function orderConnectedCodes(codes, links) {
    const neighbors = new Map(codes.map(code => [code, new Set()]));
    for (const {a,b} of links) {
      if (!neighbors.has(a) || !neighbors.has(b)) continue;
      neighbors.get(a).add(b); neighbors.get(b).add(a);
    }
    // Keep a connected concept group together; preserve Excel order within groups.
    const groups = new Map();
    let group = 0;
    for (const code of codes) {
      if (!neighbors.get(code).size || groups.has(code)) continue;
      const queue = [code]; groups.set(code, group);
      for (let i = 0; i < queue.length; i++) {
        for (const next of neighbors.get(queue[i])) {
          if (!groups.has(next)) { groups.set(next, group); queue.push(next); }
        }
      }
      group++;
    }
    return codes.map((code,index) => ({code,index,group:groups.get(code) ?? group}))
      .sort((a,b) => a.group-b.group || a.index-b.index).map(item => item.code);
  }
  function renderGraph() {
    const chain = currentChain(), links = activeLinks();
    const orderedCodes = orderConnectedCodes(chain.codes.filter(c => nodeMap.has(c)), links);
    const visibleStages = state.showEmpty ? stages : stages.filter(s=>orderedCodes.some(c=>nodeMap.get(c).stage===s.code));
    $('#lanes').style.gridTemplateColumns=`repeat(${visibleStages.length || 1}, minmax(290px,1fr))`;
    $('#map-canvas').style.minWidth=`${Math.max(1,visibleStages.length)*311+38}px`;
    $('#lanes').innerHTML = visibleStages.map(s => {
      const codes = orderedCodes.filter(c => nodeMap.get(c)?.stage === s.code);
      return `<section class="lane" aria-label="${s.name}"><header class="stage-header"><span class="stage-num">${s.num}</span><div><strong>${s.name}</strong><small>${s.grade}</small></div></header><div class="lane-nodes">${codes.length ? codes.map(c => nodeCard(nodeMap.get(c), chain, orderedCodes.indexOf(c))).join('') : '<div class="empty-stage"><span aria-hidden="true">—</span><p>空白<br>尚未建立直接連結</p></div>'}</div></section>`;
    }).join('');
    $('#chain-title').textContent = chain.name;
    $('#chain-description').textContent = chain.description;
    $('#breadcrumb-subject').textContent = chain.subject.startsWith('跨科') ? '跨科整合' : chain.subject;
    $('#chain-number').textContent = String(data.chains.indexOf(chain)+1).padStart(2,'0')+' / '+String(data.chains.length).padStart(2,'0');
    $('#basis-badge').textContent = chain.basis === '使用者指定' ? '使用者指定・直接關係' : '分析建議・待教師審閱';
    $('#basis-badge').classList.toggle('suggested', chain.basis !== '使用者指定');
    $('#graph-stats').textContent = `${orderedCodes.length} 個條目 / ${links.length} 組關係`;
    $('#scope-note').textContent = chain.basis === '使用者指定' ? '條目可沒有前後連結；未建立直接關係的學段留白' : '分析草案待教師審閱；同單元條目不自動連線';
    $('#cross-toggle').disabled = !data.links.some(l => l.chain === chain.id && ['跨科整合','同階段概念整合'].includes(l.type));
    requestAnimationFrame(() => { equalizeCards(); drawConnections(); });
    resizeObserver?.disconnect();
    resizeObserver = new ResizeObserver(() => requestAnimationFrame(drawConnections));
    resizeObserver.observe($('#lanes'));
  }
  function equalizeCards() {
    const rows = [];
    $$('.lane').forEach(lane => $$('.node-card',lane).forEach((el,i) => { el.style.minHeight=''; (rows[i] ||= []).push(el); }));
    rows.forEach(cards => { const max = Math.max(...cards.map(c => c.getBoundingClientRect().height)); cards.forEach(c => c.style.minHeight=max+'px'); });
    const firstHeight = rows[0]?.[0]?.getBoundingClientRect().height;
    if(firstHeight) $$('.empty-stage').forEach(el => el.style.minHeight=firstHeight+'px');
  }
  function drawConnections() {
    const canvas = $('#map-canvas'), svg = $('#connections');
    if(!canvas || !svg) return;
    const rect = canvas.getBoundingClientRect();
    svg.setAttribute('viewBox',`0 0 ${rect.width} ${rect.height}`);
    svg.innerHTML = '<defs><marker id="arrow-direct" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="m1 1 6 3-6 3" fill="none" stroke="#789575" stroke-width="1.4"/></marker><marker id="arrow-cross" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="5" markerHeight="5" orient="auto"><path d="m1 1 6 3-6 3" fill="none" stroke="#b99f6c" stroke-width="1.4"/></marker></defs>';
    activeLinks().forEach(l => {
      const a = $$('[data-code]').find(el => el.dataset.code === l.a);
      const b = $$('[data-code]').find(el => el.dataset.code === l.b);
      if(!a || !b) return;
      const ar=a.getBoundingClientRect(),br=b.getBoundingClientRect();
      const same = nodeMap.get(l.a).stage === nodeMap.get(l.b).stage;
      const cross = ['跨科整合','同階段概念整合'].includes(l.type);
      let path,hitArea='';
      if(same){
        const x=ar.right-rect.left, y1=ar.top+ar.height*.72-rect.top, y2=br.top+br.height*.28-rect.top;
        path=`M ${x} ${y1} C ${x+14} ${y1}, ${x+14} ${y2}, ${br.right-rect.left} ${y2}`;
      } else {
        const x1=ar.right-rect.left,y1=ar.top+ar.height*.5-rect.top,x2=br.left-rect.left,y2=br.top+br.height*.5-rect.top;
        const dx=Math.max(12,Math.abs(x2-x1)*.4);
        path=Math.abs(y1-y2)<1 ? `M ${x1} ${y1} L ${x2} ${y2}` : `M ${x1} ${y1} C ${x1+dx} ${y1}, ${x2-dx} ${y2}, ${x2} ${y2}`;
        if(Math.abs(y1-y2)<1)hitArea=`<rect x="${Math.min(x1,x2)}" y="${y1-10}" width="${Math.abs(x2-x1)}" height="20" fill="transparent" style="pointer-events:all" data-link="${esc(l.id)}"/>`;
      }
      const cls=cross?'cross':currentChain().basis==='使用者指定'?'':'suggested';
      const desc=`${l.a} 至 ${l.b}，${l.type}。點選查看連結理由`;
      svg.insertAdjacentHTML('beforeend',`<g role="button" tabindex="0" aria-label="${esc(desc)}" data-link="${esc(l.id)}">${hitArea}<path class="connection-hit" d="${path}" data-link="${esc(l.id)}" aria-hidden="true"/><path class="connection-path ${cls}" d="${path}" marker-end="url(#arrow-${cross?'cross':'direct'})" data-link="${esc(l.id)}"><title>${esc(desc)}</title></path></g>`);
    });
  }
  function renderRelations() {
    if(!$('#relation-list'))return;
    const links=activeLinks();
    $('#relation-count').textContent=`${links.length} 組已列出關係`;
    $('#relation-list').innerHTML=links.map(l=>`<button class="relation-item" data-link="${esc(l.id)}"><div class="relation-pair"><span>${esc(l.a)}</span>${icon('arrow')}<span>${esc(l.b)}</span></div><p>${esc(l.why)}</p><div class="relation-bottom"><strong>${esc(l.type)}</strong><span>${l.review==='使用者指定'?'使用者指定':'分析初稿'} <span aria-hidden="true">↗</span></span></div></button>`).join('')||'<div class="relation-empty">目前篩選下沒有連結。開啟「跨科整合」可顯示其他關係。</div>';
  }
  function selectChain(id) {
    if(!data.chains.some(c=>c.id===id))return;
    state.chain=id;
    closeDrawer(false);
    renderNavigation();renderGraph();setHash();
    $('#map-scroller').scrollLeft=0;
  }
  function showDrawer(heading,html) {
    drawerFocus=document.activeElement;
    $('#drawer-heading').textContent=heading;
    $('#drawer-content').innerHTML=html;
    $('#drawer').hidden=false;$('#drawer-overlay').hidden=false;
    $('#drawer-content').scrollTop=0;
    document.body.classList.add('modal-open');
    $('#close-drawer').focus();
  }
  function closeDrawer(restore=true) {
    if($('#drawer').hidden)return;
    $('#drawer').hidden=true;$('#drawer-overlay').hidden=true;
    $$('.node-card.selected').forEach(el=>el.classList.remove('selected'));
    if($('#modal-overlay').hidden)document.body.classList.remove('modal-open');
    setHash();
    if(restore && drawerFocus?.isConnected)drawerFocus.focus();
  }
  function sourceLinks(pdfs) {
    return String(pdfs).split('、').filter(p=>/^\d+$/.test(p)).map(p=>`<a href="assets/curriculum.pdf#page=${p}" target="_blank" rel="noopener">PDF 第 ${p} 頁 · 內文 ${Number(p)-4} 頁 ↗</a>`).join('');
  }
  function showNode(code) {
    const n=nodeMap.get(code);if(!n)return;
    const chain=currentChain(),all=expsByCode.get(code)||[];
    const numbered=all.filter(e=>e.kind==='學習內容說明');
    const context=all.filter(e=>e.kind!=='學習內容說明');
    const focus=chain.focus[code];
    const related=data.links.filter(l=>l.a===code||l.b===code);
    const summary=chain.summaries?.[n.stage];
    const constraints=n.limits;
    showDrawer('學習內容與說明',`<div class="detail-meta"><span>${esc(stages.find(s=>s.code===n.stage)?.name)}</span><span>／</span><span>${esc(n.subject)}</span></div><div class="detail-code">${esc(n.code)}</div><h2 class="detail-title">${esc(n.subtopic||n.topic)}</h2><section class="detail-section"><h3>學習內容原文 <small>課綱文字</small></h3><blockquote class="original">${esc(n.text)}</blockquote></section>${focus?`<p class="detail-note">本條連結聚焦下方說明 ${esc(focus)}；同一課綱代碼的其他說明仍完整保留。</p>`:''}<section class="detail-section"><h3>學習內容說明 <small>${numbered.length} 則編號說明</small></h3>${numbered.map(e=>`<article class="explanation-row ${e.seq===focus?'focused':''}"><div class="exp-top"><span class="exp-seq">${esc(e.seq)}</span>${e.seq===focus?'<span class="exp-focus">本連結聚焦</span>':''}</div><p>${esc(e.text)}</p></article>`).join('')||'<p class="detail-note">未匹配獨立的編號說明，請查看共用說明與來源頁面。</p>'}</section>${constraints?`<section class="detail-section"><h3>範圍限制摘錄 <small>依原文限制語句擷取</small></h3><div class="limit-box">${esc(constraints)}</div></section>`:''}${summary?.codes.includes(code)?`<section class="detail-section"><h3>本主線分析摘要 <small>分析建議，非課綱原文</small></h3><div class="context-text">${esc(summary.text)}<br><br>新增深度：${esc(summary.new)}<br>判讀界線：${esc(summary.limit)}</div></section>`:''}${context.length?`<details class="context-details"><summary>查看共用說明、備註與參考節數（${context.length} 筆）</summary><p class="detail-note">以下依來源表格區塊保留；多條目共用範圍尚待逐條複核。</p>${context.map(e=>`<div class="context-text"><strong>${esc(e.kind)}</strong><br>${esc(e.text)}<br><small>區塊 ${esc(e.block)}</small></div>`).join('')}</details>`:''}${n.alternates?`<details class="context-details"><summary>同碼其他出處原文</summary><div class="context-text">${esc(n.alternates)}</div></details>`:''}<section class="detail-section"><h3>相關連結 <small>${related.length} 組</small></h3>${related.map(l=>`<button class="detail-link" data-link="${esc(l.id)}">${esc(l.a)} → ${esc(l.b)}<span>${esc(l.type)} · ${l.review==='使用者指定'?'使用者指定':'分析建議'}</span></button>`).join('')||'<p class="detail-note">此條目尚未建立連結，不代表沒有先備或後續內容。</p>'}</section><section class="detail-section"><h3>來源頁面</h3><div class="source-links">${sourceLinks(n.pdf)}</div><p class="detail-note">${esc(data.meta.source)}（${esc(data.meta.edition)}）。原文已移除排版換行；字形、公式及對應關係仍可回查 PDF。<br>擷取狀態：${esc(n.status)}</p></section><div class="detail-actions"><button class="button" data-copy="${esc(code)}">${icon('link')}複製條目連結</button><button class="button" data-action="library">${icon('search')}查閱其他課綱</button></div>`);
    $$('.node-card').forEach(el=>el.classList.toggle('selected',el.dataset.code===code));
    setHash(code);
  }
  function showLink(id) {
    const l=data.links.find(x=>x.id===id);if(!l)return;
    const a=nodeMap.get(l.a),b=nodeMap.get(l.b);
    const mini=n=>`<button class="mini-node" data-node="${esc(n.code)}"><code>${esc(n.code)}</code><p>${esc(n.text)}</p></button>`;
    showDrawer('連結理由',`<div class="detail-meta"><span>${esc(l.type)}</span><span>／</span><span>${l.review==='使用者指定'?'使用者指定':'分析建議・待教師審閱'}</span></div><h2 class="detail-title">從哪個概念，連到哪個條目？</h2>${mini(a)}<div class="relation-detail-arrow" aria-hidden="true">↓</div>${mini(b)}<section class="detail-section"><h3>連結理由 <small>${l.review==='使用者指定'?'依使用者指定整理':'分析文字，非課綱原文'}</small></h3><div class="original">${esc(l.why)}</div></section><section class="detail-section"><h3>關係中的深化重點</h3><p class="context-text">${esc(l.depth)}</p></section><section class="detail-section"><h3>範圍界線</h3><div class="limit-box">${esc(l.limit)}</div></section><p class="detail-note">連線只表示這兩個具體條目的已列出關係，不代表所有相同科目或相鄰學段的條目都能相連。</p>`);
  }
  function showAbout(){
    showDrawer('資料來源與判讀方式',`<div class="eyebrow">ABOUT THIS EXPLORER</div><h2 class="detail-title">以條目為節點，以關係連線。</h2><p class="context-text">這個頁面用來檢視自然科課綱的縱向關係。先看一條實際連結，再回到原文與學習內容說明確認範圍。</p><div class="about-stats"><div><strong>${data.nodes.length}</strong><small>課綱代碼</small></div><div><strong>${data.meta.numberedExplanations.toLocaleString()}</strong><small>編號說明</small></div><div><strong>${data.chains.length}</strong><small>概念單元</small></div></div><ul class="about-list"><li><strong>一張卡片，一個課綱條目。</strong>原文保留在卡片與詳細資料中。連結可從國小Ⅱ直接跨到國中Ⅳ，不補造中間條目。</li><li><strong>依直接關係連線。</strong>網站版本的「人類活動與環境影響」及「生態層次與族群特徵」各保留一組使用者指定關係；其他配對是全單元分析建議，尚待教師審閱。匯入版本依 Excel 的審閱狀態與啟用設定呈現。</li><li><strong>空白學段只是未列入。</strong>不要求每個條目都有前面的先備或後面的延伸，不能為了填滿學段而連線。空白也不能解讀為該學段沒有相關課綱。</li><li><strong>內容與說明一起判讀。</strong>同一代碼可有多條說明，連結可能只聚焦其中一條，例如高中族群條目的2-2。</li><li><strong>來源差異保留。</strong>本文與附錄代碼已比對。個別同碼多處文字、共用說明及擷取差異仍需回查來源；原文尚未全部逐條人工校對。</li><li><strong>資料範圍。</strong>本版涵蓋國小Ⅱ、Ⅲ、國中Ⅳ、高中必修Ⅴc與加深加廣選修Ⅴa的學習內容與說明，未另建學習表現與核心素養對照。</li></ul><section class="detail-section"><h3>原始資料</h3><p class="context-text">${esc(data.meta.source)}<br>版本：${esc(data.meta.edition)}<br>整理日期：${esc(data.meta.updated)}</p><div class="source-links"><a href="assets/curriculum.pdf" target="_blank" rel="noopener">開啟完整課綱 PDF ↗</a></div></section>`);
  }
  function openLibrary(){
    lastFocus=document.activeElement;
    closeDrawer(false);
    $('#modal-overlay').hidden=false;document.body.classList.add('modal-open');
    renderLibrary();$('#library-search').focus();
  }
  function closeLibrary(){
    $('#modal-overlay').hidden=true;
    if($('#drawer').hidden)document.body.classList.remove('modal-open');
    if(lastFocus?.isConnected)lastFocus.focus();
  }
  function renderLibrary(){
    const terms=state.query.trim().split(/\s+/).map(normal).filter(Boolean);
    const filtered=data.nodes.filter(n=>(!state.librarySubject||n.subject.includes(state.librarySubject))&&(!state.libraryStage||n.stage===state.libraryStage)&&(!state.connectedOnly||linkedCodes.has(n.code))&&terms.every(t=>searchIndex.get(n.code).includes(t)));
    const pageSize=16,total=Math.max(1,Math.ceil(filtered.length/pageSize));
    state.page=Math.max(1,Math.min(state.page,total));
    $('#library-count').textContent=`找到 ${filtered.length} 個條目`;
    $('#page-label').textContent=`${state.page} / ${total}`;
    $('#prev-page').disabled=state.page===1;$('#next-page').disabled=state.page===total;
    $('#library-results').innerHTML=filtered.slice((state.page-1)*pageSize,state.page*pageSize).map(n=>{
      const explanations=expsByCode.get(n.code).filter(e=>e.kind==='學習內容說明'||e.kind==='未編號說明（區塊共用）');
      return `<article class="library-row"><code>${esc(n.code)}</code><span class="result-text">${esc(n.text)}<small class="result-meta">${esc(stages.find(s=>s.code===n.stage)?.name)} · ${esc(n.subject)} · ${esc(n.subtopic||n.topic)}</small></span><span class="result-state ${linkedCodes.has(n.code)?'linked':''}">${linkedCodes.has(n.code)?'已有連結':'尚未建立連結'}</span><section class="result-inline-explanations"><h4>學習內容說明</h4>${explanations.map(e=>`<p><code>${esc(e.seq||'區塊共用')}</code>${esc(e.text)}</p>`).join('')||'<p>未匹配獨立說明，需回查來源。</p>'}</section><button class="text-button" data-node="${esc(n.code)}">來源與相關連結 ↗</button></article>`;
    }).join('')||'<div class="library-empty"><strong>沒有符合的條目</strong>試著縮短關鍵字，或放寬科目與學段篩選。</div>';
    $('#library-results').scrollTop=0;
  }
  function notify(text){$('#toast').textContent=text;$('#toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('#toast').hidden=true,2400);}
  async function copyNode(code){
    const url=new URL(location.href);url.hash=state.chain+'/node='+encodeURIComponent(code);
    try{await navigator.clipboard.writeText(url.href);notify('已複製條目連結');}
    catch{notify('目前環境無法複製，請複製瀏覽器網址列');}
  }
  document.addEventListener('click',e=>{
    const el=e.target.closest('[data-subject],[data-chain],[data-node],[data-link],[data-action],[data-copy]');
    if(!el)return;
    if(el.dataset.subject){
      state.subject=el.dataset.subject;
      $('#unit-search').value='';
      const cs=data.chains.filter(c=>state.subject==='全部科目'||(state.subject==='跨科'?c.subject.startsWith('跨科'):c.subject===state.subject));
      if(cs.length&&!cs.some(c=>c.id===state.chain))selectChain(cs[0].id);else renderNavigation();
    } else if(el.dataset.chain)selectChain(el.dataset.chain);
    else if(el.dataset.node)showNode(el.dataset.node);
    else if(el.dataset.link)showLink(el.dataset.link);
    else if(el.dataset.action==='library')openLibrary();
    else if(el.dataset.copy)copyNode(el.dataset.copy);
  });
  $('#connections').addEventListener('keydown',e=>{if(['Enter',' '].includes(e.key)&&e.target.dataset.link){e.preventDefault();showLink(e.target.dataset.link);}});
  $('#cross-toggle').addEventListener('change',e=>{state.cross=e.target.checked;renderGraph();});
  $('#empty-toggle').addEventListener('change',e=>{state.showEmpty=e.target.checked;renderGraph();});
  $('#reset-view').addEventListener('click',()=>$('#map-scroller').scrollTo({left:0,top:0,behavior:'smooth'}));
  $('#close-drawer').addEventListener('click',()=>closeDrawer());$('#drawer-overlay').addEventListener('click',()=>closeDrawer());
  $('#about-button').addEventListener('click',showAbout);
  $('#close-library').addEventListener('click',closeLibrary);
  $('#modal-overlay').addEventListener('click',e=>{if(e.target===$('#modal-overlay'))closeLibrary();});
  $('#library-search').addEventListener('input',e=>{state.query=e.target.value;state.page=1;renderLibrary();});
  $('#library-subject').addEventListener('change',e=>{state.librarySubject=e.target.value;state.page=1;renderLibrary();});
  $('#library-stage').addEventListener('change',e=>{state.libraryStage=e.target.value;state.page=1;renderLibrary();});
  $('#library-connected').addEventListener('change',e=>{state.connectedOnly=e.target.checked;state.page=1;renderLibrary();});
  $('#prev-page').addEventListener('click',()=>{state.page--;renderLibrary();});$('#next-page').addEventListener('click',()=>{state.page++;renderLibrary();});
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'){if(!$('#drawer').hidden)closeDrawer();else if(!$('#modal-overlay').hidden)closeLibrary();}
    if(e.key==='Tab'){
      const modal=!$('#drawer').hidden?$('#drawer'):!$('#modal-overlay').hidden?$('.library'):null;if(!modal)return;
      const focusable=$$('button:not(:disabled),a[href],input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',modal).filter(el=>el.getClientRects().length);
      if(!focusable.length)return;
      const first=focusable[0],last=focusable[focusable.length-1];
      if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}
      else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
    }
  });
  window.addEventListener('resize',()=>{equalizeCards();drawConnections();});
  window.addEventListener('hashchange',()=>{
    const [id,part]=location.hash.slice(1).split('/');
    if(data.chains.some(c=>c.id===id)&&id!==state.chain){state.chain=id;state.subject='全部科目';renderNavigation();renderGraph();}
    if(part?.startsWith('node='))showNode(decodeURIComponent(part.slice(5)));
  });
  function updateDatasetView(){
    rebuildIndexes();state.subject='全部科目';state.page=1;state.chain=data.chains.some(c=>c.id===state.chain)?state.chain:data.chains[0].id;
    $('#unit-search').value='';closeDrawer(false);closeLibrary();renderNavigation();renderGraph();setHash();
    $('#data-status').textContent=importedName?'本機版本：'+importedName:'網站資料';
    $('.library-header p').textContent=`${data.nodes.length} 個課綱代碼，保留學習內容與說明。`;
    $('.key').textContent=data.nodes.length;
  }
  $('#unit-search').addEventListener('input',renderNavigation);
  $('#excel-file').addEventListener('change',async e=>{
    const file=e.target.files[0];if(!file)return;const ticket=++importTicket;
    pendingImport=null;$('#apply-import').hidden=true;$('#import-preview').textContent='讀取及檢查 Excel…';
    try{
      const result=CurriculumExcel.convert(await CurriculumExcel.read(file),publishedData);if(ticket!==importTicket)return;
      if(result.errors.length){$('#import-preview').textContent=`發現 ${result.errors.length} 個問題，尚未套用：\n`+result.errors.slice(0,40).join('\n')+(result.errors.length>40?'\n其餘問題請修正後重新檢查。':'');return;}
      pendingImport={data:result.data,name:file.name};
      $('#import-preview').textContent=`檢查通過：${result.data.chains.length} 個單元、${result.data.nodes.length} 條內容、${result.data.explanations.length} 則說明、${result.data.links.length} 條連線。\n${result.unassigned?`${result.unassigned} 條內容未放入啟用單元，仍可從全部課綱查閱。`:'全部內容均已放入單元。'}按「套用資料」更新。`;
      $('#apply-import').hidden=false;
    }catch(err){if(ticket===importTicket)$('#import-preview').textContent='無法匯入，原資料保留：'+err.message;}
    finally{e.target.value='';}
  });
  $('#apply-import').addEventListener('click',()=>{
    if(!pendingImport)return;
    try{localStorage.setItem(storageKey,JSON.stringify(pendingImport));}
    catch{$('#import-preview').textContent='瀏覽器無法儲存資料（空間不足或儲存遭停用），尚未套用。請確認瀏覽器允許此網站儲存資料。';return;}
    data=pendingImport.data;importedName=pendingImport.name;pendingImport=null;updateDatasetView();$('#apply-import').hidden=true;$('#import-preview').textContent='已套用。重新整理仍會保留本機版本。';notify('Excel 已更新至網頁');
  });
  $('#restore-data').addEventListener('click',()=>{
    try{localStorage.removeItem(storageKey);}catch{$('#import-preview').textContent='無法移除本機資料，請確認瀏覽器儲存設定。';return;}
    importTicket++;pendingImport=null;data=publishedData;importedName='';updateDatasetView();$('#apply-import').hidden=true;$('#import-preview').textContent='已恢復網站原始資料；你的 Excel 檔案不受影響。';
  });
  $('#data-status').textContent=importedName?'本機版本：'+importedName:'網站資料';
  const [initialChain,initialNode]=location.hash.slice(1).split('/');
  if(!data.chains.some(c=>c.id===state.chain))state.chain=data.chains[0].id;
  if(data.chains.some(c=>c.id===initialChain))state.chain=initialChain;
  hydrateIcons();renderNavigation();renderGraph();
  if(initialNode?.startsWith('node='))showNode(decodeURIComponent(initialNode.slice(5)));else setHash();
  document.fonts.ready.then(()=>{equalizeCards();drawConnections();});
})();
