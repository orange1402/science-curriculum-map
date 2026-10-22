'use strict';
// Local XLSX reader. No uploads, macros, external links or formula evaluation.
window.CurriculumExcel = (() => {
  const decoder = new TextDecoder();
  function xml(s) { const doc=new DOMParser().parseFromString(s,'application/xml');if(doc.querySelector('parsererror'))throw Error('Excel XML 格式無法讀取');return doc; }
  const elements=(el,name)=>[...el.getElementsByTagNameNS('*',name)];
  async function unzip(buffer){
    const v=new DataView(buffer),u=new Uint8Array(buffer),files=new Map();let end=-1;
    for(let i=u.length-22;i>=Math.max(0,u.length-65557);i--)if(v.getUint32(i,true)===0x06054b50){end=i;break;}
    if(end<0)throw Error('請選擇有效的 .xlsx 檔案（不支援 .xls 或加密檔）');
    const count=v.getUint16(end+10,true);let p=v.getUint32(end+16,true),total=0;
    if(count>5000)throw Error('檔案內項目過多');
    for(let i=0;i<count;i++){
      if(v.getUint32(p,true)!==0x02014b50)throw Error('Excel 壓縮目錄損壞');
      const flag=v.getUint16(p+8,true),method=v.getUint16(p+10,true),size=v.getUint32(p+20,true),rawSize=v.getUint32(p+24,true),nl=v.getUint16(p+28,true),xl=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),offset=v.getUint32(p+42,true);
      const name=decoder.decode(u.subarray(p+46,p+46+nl));p+=46+nl+xl+cl;
      if(flag&1)throw Error('不支援加密檔案');
      total+=rawSize;if(total>60000000)throw Error('解壓後檔案超過60MB，請使用下載的範本');
      if(!name.endsWith('.xml')&&!name.endsWith('.rels'))continue;
      if(v.getUint32(offset,true)!==0x04034b50)throw Error('Excel 壓縮內容損壞');
      const start=offset+30+v.getUint16(offset+26,true)+v.getUint16(offset+28,true);
      const bytes=u.subarray(start,start+size);let raw;
      if(method===0)raw=bytes;
      else if(method===8)raw=new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
      else throw Error('不支援此壓縮格式');
      if(raw.length!==rawSize)throw Error('Excel 解壓長度不符');
      files.set(name,decoder.decode(raw));
    }
    return files;
  }
  async function read(file){
    if(file.size>15000000)throw Error('檔案超過15MB，請移除多餘圖片再匯入');
    const z=await unzip(await file.arrayBuffer());
    if(!z.has('xl/workbook.xml')||!z.has('xl/_rels/workbook.xml.rels'))throw Error('找不到 Excel 活頁簿');
    const shared=z.has('xl/sharedStrings.xml')?elements(xml(z.get('xl/sharedStrings.xml')),'si').map(s=>elements(s,'t').map(t=>t.textContent).join('')):[];
    const rels=new Map(elements(xml(z.get('xl/_rels/workbook.xml.rels')),'Relationship').map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));
    const sheets={};
    for(const s of elements(xml(z.get('xl/workbook.xml')),'sheet')){
      const rel=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
      let target=rels.get(rel);if(!target)continue;
      target=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
      if(!z.has(target))throw Error('找不到工作表：'+s.getAttribute('name'));
      sheets[s.getAttribute('name')]=elements(xml(z.get(target)),'row').map(row=>{
        const cells=[];
        for(const c of elements(row,'c')){
          const letters=(c.getAttribute('r')||'').match(/^[A-Z]+/)?.[0];if(!letters)continue;
          let col=0;for(const l of letters)col=col*26+l.charCodeAt(0)-64;
          const t=c.getAttribute('t'),value=elements(c,'v')[0]?.textContent??'';
          cells[col-1]=t==='s'?shared[Number(value)]??'':t==='inlineStr'?elements(c,'t').map(t=>t.textContent).join(''):value;
        }
        return {row:Number(row.getAttribute('r')),cells};
      });
    }
    return sheets;
  }
  function convert(sheets,base){
    const errors=[];
    const definitions={
      '單元主線':['主線ID','科目','單元名稱','單元說明','審閱狀態','啟用'],
      '單元條目':['主線ID','課綱代碼','聚焦說明序號','顯示順序','分析結果','分析註記','啟用'],
      '縱向連結':['連結ID','主線ID','起點代碼','終點代碼','連結類型','連結理由','新增深度','範圍界線','審閱狀態','啟用'],
      '課綱原文':['課綱代碼','學段','科目','主題','次主題','學習內容','限制摘錄','PDF頁次','內文頁碼','原始代碼','來源區塊','擷取狀態','其他出處'],
      '內容說明':['說明ID','適用代碼','學段','科目','說明序號','資料類型','說明內容','PDF頁次','內文頁碼','來源區塊']
    };
    const error=(r,msg)=>errors.push(`${r._sheet} 第${r._row}列：${msg}`);
    function rows(name){
      const rs=sheets[name];if(!rs){errors.push('缺少工作表：'+name);return[];}
      const header=rs.find(r=>r.row<=10&&definitions[name].every(h=>r.cells.includes(h)));
      if(!header){errors.push(name+'：前10列找不到完整欄名（請保留範本欄名）');return[];}
      return rs.filter(r=>r.row>header.row&&r.cells.some(x=>String(x??'').trim())).map(r=>Object.assign({_sheet:name,_row:r.row},Object.fromEntries(definitions[name].map(h=>[h,String(r.cells[header.cells.indexOf(h)]??'').trim()]))));
    }
    const source=Object.fromEntries(Object.keys(definitions).map(n=>[n,rows(n)]));
    function unique(rs,key){const ids=new Set();for(const r of rs){if(!r[key])error(r,key+'不可空白');else if(ids.has(r[key]))error(r,key+'重複：'+r[key]);ids.add(r[key]);}}
    function enabled(r){if(!['是','否'].includes(r.啟用)){error(r,'啟用限填「是」或「否」');return false;}return r.啟用==='是';}
    unique(source.課綱原文,'課綱代碼');unique(source.內容說明,'說明ID');unique(source.單元主線,'主線ID');unique(source.縱向連結,'連結ID');
    const stages=['Ⅱ','Ⅲ','Ⅳ','Ⅴc','Ⅴa'],subjects=['生物','物理','化學','地球科學','跨科'];
    const nodes=source.課綱原文.map(r=>{
      if(!/^[A-Za-z]+-[ⅡⅢⅣⅤ]+[ac]?-\d+$/.test(r.課綱代碼))error(r,'課綱代碼格式不符：'+r.課綱代碼);
      if(!stages.includes(r.學段))error(r,'學段限用Ⅱ、Ⅲ、Ⅳ、Ⅴc、Ⅴa');
      if(!r.學習內容)error(r,'學習內容不可空白');
      return {code:r.課綱代碼,stage:r.學段,subject:r.科目,topic:r.主題,subtopic:r.次主題,text:r.學習內容,limits:r.限制摘錄,pdf:r.PDF頁次,page:r.內文頁碼,originalCode:r.原始代碼,block:r.來源區塊,status:r.擷取狀態,alternates:r.其他出處};
    });
    const nodeMap=new Map(nodes.map(n=>[n.code,n]));
    const explanations=source.內容說明.map(r=>{
      const codes=r.適用代碼.split(/[；;]/).map(c=>c.trim()).filter(Boolean);
      if(!codes.length)error(r,'適用代碼不可空白');
      codes.forEach(c=>{if(!nodeMap.has(c))error(r,'適用代碼不存在：'+c);});
      if(!r.說明內容)error(r,'說明內容不可空白');
      return {id:r.說明ID,codes,stage:r.學段,subject:r.科目,seq:r.說明序號,kind:r.資料類型,text:r.說明內容,pdf:r.PDF頁次,page:r.內文頁碼,block:r.來源區塊};
    });
    const chains=source.單元主線.filter(r=>{
      if(!/^[A-Za-z][A-Za-z0-9_-]{0,49}$/.test(r.主線ID))error(r,'主線ID限英文字開頭的英數字、底線及連字號');
      if(!subjects.includes(r.科目))error(r,'科目限生物、物理、化學、地球科學、跨科');
      if(!r.單元名稱)error(r,'單元名稱不可空白');
      return enabled(r);
    }).map(r=>({id:r.主線ID,name:r.單元名稱,subject:r.科目,description:r.單元說明,basis:r.審閱狀態,focus:{},codes:[]}));
    const allUnits=new Set(source.單元主線.map(r=>r.主線ID)),chainMap=new Map(chains.map(c=>[c.id,c]));
    const members=[],seen=new Set();
    for(const r of source.單元條目){
      const on=enabled(r);if(!allUnits.has(r.主線ID))error(r,'主線ID不存在：'+r.主線ID);
      if(!on)continue;
      if(!nodeMap.has(r.課綱代碼))error(r,'課綱代碼不存在：'+r.課綱代碼);
      const key=r.主線ID+'|'+r.課綱代碼;if(seen.has(key))error(r,'同一單元重複條目：'+r.課綱代碼);seen.add(key);
      if(!r.顯示順序||!Number.isFinite(Number(r.顯示順序)))error(r,'顯示順序須填數字');
      if(r.聚焦說明序號&&!explanations.some(e=>e.codes.includes(r.課綱代碼)&&e.seq===r.聚焦說明序號))error(r,'此代碼找不到聚焦說明序號：'+r.聚焦說明序號);
      members.push({chain:r.主線ID,code:r.課綱代碼,focus:r.聚焦說明序號,order:Number(r.顯示順序),status:r.分析結果,note:r.分析註記});
    }
    for(const m of [...members].sort((a,b)=>a.order-b.order)){const c=chainMap.get(m.chain);if(c){c.codes.push(m.code);if(m.focus)c.focus[m.code]=m.focus;}}
    const links=[],pairs=new Set();
    for(const r of source.縱向連結){
      const on=enabled(r);if(!allUnits.has(r.主線ID))error(r,'主線ID不存在：'+r.主線ID);if(!on)continue;
      for(const key of ['起點代碼','終點代碼'])if(!nodeMap.has(r[key]))error(r,key+'不存在：'+r[key]);
      if(r.起點代碼===r.終點代碼)error(r,'起點與終點不可相同');
      const a=nodeMap.get(r.起點代碼),b=nodeMap.get(r.終點代碼);
      if(a&&b&&stages.indexOf(a.stage)>stages.indexOf(b.stage))error(r,'起點學段晚於終點，請調換代碼');
      const pair=r.主線ID+'|'+r.起點代碼+'|'+r.終點代碼;if(pairs.has(pair))error(r,'同單元重複配對');pairs.add(pair);
      if(!r.連結理由)error(r,'連結理由不可空白');
      const c=chainMap.get(r.主線ID);if(!c)continue;
      for(const code of [r.起點代碼,r.終點代碼])if(!c.codes.includes(code))c.codes.push(code);
      links.push({id:r.連結ID,chain:r.主線ID,a:r.起點代碼,b:r.終點代碼,type:r.連結類型,why:r.連結理由,depth:r.新增深度,limit:r.範圍界線,review:r.審閱狀態});
    }
    if(!chains.length)errors.push('至少需要一個啟用的單元');if(!nodes.length)errors.push('課綱原文不可空白');
    const data={nodes,explanations,chains,links,members,meta:{...base.meta,codes:nodes.length,numberedExplanations:explanations.filter(e=>e.kind==='學習內容說明').length,analysisLinks:links.filter(l=>l.review!=='使用者指定').length,confirmedLinks:links.filter(l=>l.review==='使用者指定').length}};
    const assigned=new Set(chains.flatMap(c=>c.codes));
    return {errors,data,unassigned:nodes.filter(n=>!assigned.has(n.code)).length};
  }
  return {read,convert};
})();
