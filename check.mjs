import fs from 'node:fs/promises';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const source=await fs.readFile(new URL('./docs/data.js',import.meta.url),'utf8');
const context={window:{}};vm.runInNewContext(source,context);
const d=context.window.CURRICULUM_DATA;
assert.equal(d.nodes.length,779);
assert.equal(new Set(d.nodes.map(n=>n.code)).size,779);
const codes=new Set(d.nodes.map(n=>n.code));
for(const l of d.links){assert(codes.has(l.a));assert(codes.has(l.b));}
const userChain=d.chains.find(c=>c.id==='U01');
assert.equal(userChain.codes.join('|'),'INf-Ⅱ-5|Lb-Ⅳ-2');
assert.equal(d.links.filter(l=>l.chain==='U01').length,1);
assert(!d.links.some(l=>l.a==='Lb-Ⅳ-2'&&l.b==='BLb-Ⅴa-2'));
const populationChain=d.chains.find(c=>c.id==='U02');
assert.equal(populationChain.codes.join('|'),'Fc-Ⅳ-1|BLb-Ⅴa-2');
assert(d.links.some(l=>l.chain==='U02'&&l.a==='Fc-Ⅳ-1'&&l.b==='BLb-Ⅴa-2'));
assert(!userChain.codes.some(c=>d.nodes.find(n=>n.code===c).stage.startsWith('Ⅴ')));
assert(!populationChain.codes.some(c=>['Ⅱ','Ⅲ'].includes(d.nodes.find(n=>n.code===c).stage)));
assert.equal(d.explanations.filter(e=>e.kind==='學習內容說明'&&e.codes.includes('BLb-Ⅴa-2')).length,4);
assert(d.explanations.some(e=>e.codes.includes('BLb-Ⅴa-2')&&e.seq==='2-2'&&e.text.includes('不涉及數學模式')));
for(const c of d.chains)for(const code of c.codes)assert(codes.has(code));
assert.equal(d.chains.length,84);
assert.equal(new Set(d.chains.flatMap(c=>c.codes)).size,779,'Every source code must belong to a reviewed unit');
assert.equal(d.links.length,244);
assert.equal(d.links.filter(l=>l.review==='使用者指定').length,2);
assert(d.explanations.every(e=>e.text.trim().length),'No empty explanation rows');
const pairs=new Set();
for(const l of d.links){
 const chain=d.chains.find(c=>c.id===l.chain);assert(chain.codes.includes(l.a)&&chain.codes.includes(l.b));
 const key=[l.chain,l.a,l.b].join('|');assert(!pairs.has(key));pairs.add(key);
 assert(l.why&&l.depth&&l.limit);
}
assert.equal(d.explanations.find(e=>e.codes.includes('Fb-Ⅳ-4')&&e.seq==='4-2').text,'不涉及月亮升落時間和方位問題。');
for(const file of ['index.html','style.css','app.js','data.js','favicon.svg','assets/curriculum.pdf'])assert((await fs.stat(new URL('./docs/'+file,import.meta.url))).size>0);
const html=await fs.readFile(new URL('./docs/index.html',import.meta.url),'utf8');
assert(!html.includes('class="sidebar"'));
assert(html.includes('class="app-header"'));
const script=await fs.readFile(new URL('./docs/app.js',import.meta.url),'utf8');
new vm.Script(script);
new vm.Script(await fs.readFile(new URL('./docs/excel-import.js',import.meta.url),'utf8'));
assert(script.includes('card-explanations'));
assert(script.includes('result-inline-explanations'));
const orderFunction=script.slice(script.indexOf('  function orderConnectedCodes('),script.indexOf('  function renderGraph()'));
const orderConnectedCodes=vm.runInNewContext('('+orderFunction.trim()+')');
const nodeByCode=new Map(d.nodes.map(n=>[n.code,n]));
for(const chain of d.chains){
 const links=d.links.filter(l=>l.chain===chain.id),connected=new Set(links.flatMap(l=>[l.a,l.b]));
 const sorted=Array.from(orderConnectedCodes(chain.codes,links));
 assert.deepEqual([...sorted].sort(),[...chain.codes].sort(),'Sorting must preserve every node');
 for(const stage of ['Ⅱ','Ⅲ','Ⅳ','Ⅴc','Ⅴa']){
  let reachedUnlinked=false;
  for(const code of sorted.filter(code=>nodeByCode.get(code).stage===stage)){
   if(!connected.has(code))reachedUnlinked=true;
   else assert(!reachedUnlinked,`${chain.id}: linked node below unlinked node`);
  }
 }
 assert.deepEqual(Array.from(orderConnectedCodes(chain.codes,[])),Array.from(chain.codes),'No links preserves Excel order');
}
const optics=d.chains.find(c=>c.id==='P11');
const opticsOrder=Array.from(orderConnectedCodes(optics.codes,d.links.filter(l=>l.chain==='P11')));
assert.deepEqual(opticsOrder.filter(c=>nodeByCode.get(c).stage==='Ⅳ').slice(0,2),['Ka-Ⅳ-8','Ka-Ⅳ-9']);
assert.equal(opticsOrder.filter(c=>nodeByCode.get(c).stage==='Ⅴc')[0],'PKa-Ⅴc-5');
assert.deepEqual(opticsOrder.filter(c=>nodeByCode.get(c).stage==='Ⅴa').slice(0,3),['PKa-Ⅴa-11','PKa-Ⅴa-12','PKa-Ⅴa-13']);
assert.deepEqual(Array.from(orderConnectedCodes(['z','b','x','a','y'],[{a:'a',b:'b'},{a:'a',b:'y'},{a:'x',b:'z'}])),['z','x','b','a','y'],'Transitive groups stay together without numeric sorting');
console.log('PASS: all 84 units preserve nodes and put linked concepts first; P11 concept order verified');
console.log(JSON.stringify({codes:d.nodes.length,numberedExplanations:d.explanations.filter(e=>e.kind==='學習內容說明').length,chains:d.chains.length,links:d.links.length,userConfirmedLinks:2,environmentStopsAt:'Lb-Ⅳ-2',populationStartsAt:'Fc-Ⅳ-1',navigation:'top',explanations:'inline'}));
