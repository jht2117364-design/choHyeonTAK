'use strict';

const $ = id => document.getElementById(id);
const fmt = number => Number(number).toLocaleString('ko-KR',{maximumFractionDigits:2});
const fmtRaw = number => Number(number).toLocaleString('ko-KR',{maximumFractionDigits:8});
const cfg = {
  large: {file:'io-2023-large.xlsx', count:33, label:'통합대분류'},
  medium:{file:'io-2023-medium.xlsx',count:83,label:'통합중분류'},
  small: {file:'io-2023-small.xlsx', count:165,label:'통합소분류'},
  basic: {file:'io-2023-basic.xlsx', count:380,label:'기본부문'}
};
let worker = null;
let requestCounter = 0;
const pending = new Map();
let activeLevel = null;
let sectors = [];
let rowChoices = [];
let cells = [];
let basket = [];
let result = null;
let currentUnit = '—';
let explorationSerial = 0;
let busy = false;

function element(tag,content,className=''){
  const e = document.createElement(tag);
  if(content !== undefined) e.textContent = String(content);
  if(className) e.className = className;
  return e;
}
function renderOptions(select, list){
  const fragment = document.createDocumentFragment();
  for(const x of list){
    const option = document.createElement('option');
    option.value = x.code;
    option.textContent = x.code+' · '+x.name;
    fragment.appendChild(option);
  }
  select.replaceChildren(fragment);
}
function setStatus(message,type=''){
  $('status').textContent = message;
  $('status').className = 'status'+(type?' '+type:'');
}
function setReady(ready){
  for(const id of ['sheet','row','row-search','col-search','sector','sector-search','amount','add','run']) $(id).disabled=!ready;
  $('ready-badge').textContent = ready ? '데이터 로딩 완료' : '불러오기 전';
  $('ready-badge').classList.toggle('ready',ready);
}
function bootWorker(){
  if(worker) return;
  worker = new Worker('./worker.js');
  worker.onmessage = ({data:d}) => {
    if(d.type === 'status'){
      setStatus(d.text);
      return;
    }
    const promise = pending.get(d.id);
    if(!promise) return;
    pending.delete(d.id);
    d.ok ? promise.resolve(d.result) : promise.reject(Error(d.error));
  };
  worker.onerror = event => {
    for(const {reject} of pending.values()) reject(Error(event.message||'Worker 처리 오류'));
    pending.clear();
    setStatus('브라우저 실행 엔진 오류가 발생했습니다. 새로고침 후 다시 시도하세요.','bad');
  };
}
function call(type,payload){
  bootWorker();
  return new Promise((resolve,reject)=>{
    const id=++requestCounter;
    pending.set(id,{resolve,reject});
    worker.postMessage({id,type,payload});
  });
}
function resetOutputs(){
  result = null;
  $('results').hidden = true;
  basket = [];
  renderBasket();
  cells = [];
  rowChoices = [];
  $('row').replaceChildren();
  $('sheet').replaceChildren();
  $('cells').replaceChildren(element('tr'));
  const tr=$('cells').firstElementChild;
  const td=element('td','표를 불러오면 표시됩니다.','empty');
  td.colSpan=3;
  tr.appendChild(td);
  $('explore-count').textContent='분류를 먼저 불러오세요';
  $('explore-caption').textContent='조회 결과';
  $('unit').textContent='단위: —';
  $('validation').textContent='불러온 뒤 나타납니다.';
}

async function loadWorkbook(){
  const current = $('level').value;
  const config = cfg[current];
  if(busy) return;
  busy=true;
  setReady(false);
  resetOutputs();
  $('load').disabled=true;
  $('level').disabled=true;
  $('load').textContent='불러오는 중…';
  try{
    setStatus(`${config.label} Excel 파일을 내려받고 있습니다…`);
    const response=await fetch('./data/'+config.file);
    if(!response.ok) throw Error('Excel 다운로드 실패 (HTTP '+response.status+')');
    const buffer=await response.arrayBuffer();
    const info=await call('init',{buffer,count:config.count});
    if(info.count!==config.count) throw Error('산업 수가 예상과 다릅니다.');
    activeLevel=current;
    sectors=info.sectors;
    renderOptions($('sector'),sectors);
    $('sector-search').value='';
    $('row-search').value='';
    $('col-search').value='';
    $('sheet').replaceChildren(...info.sheets.map(name=>{
      const opt=element('option',name);opt.value=name;return opt;
    }));
    $('validation').textContent=JSON.stringify(info.validation,null,2);
    setReady(true);
    await explore();
    setStatus(`${config.label} ${info.count}개 산업, ${info.sheets.length}개 시트를 불러왔으며 계수 행렬 검증을 통과했습니다.`,'good');
    $('explore-count').textContent=`산업 ${info.count}개 · 시트 ${info.sheets.length}개`;
  }catch(error){
    activeLevel=null;
    sectors=[];
    setReady(false);
    setStatus('불러오기 실패: '+error.message+'\n인터넷 연결 및 GitHub의 data/ 파일 경로를 확인하세요.','bad');
  }finally{
    busy=false;
    $('load').disabled=false;
    $('level').disabled=false;
    $('load').textContent='선택한 Excel 불러오기';
  }
}
function filterRowOptions(selected){
  const query=$('row-search').value.trim().toLowerCase();
  const filtered=rowChoices.filter(x=>(x.code+' '+x.name).toLowerCase().includes(query));
  const preferred=selected || $('row').value;
  renderOptions($('row'),filtered);
  if(filtered.some(r=>r.code===preferred))$('row').value=preferred;
  return filtered;
}
async function explore(){
  if(!activeLevel) return;
  const serial=++explorationSerial;
  try{
    const currentSheet=$('sheet').value;
    const response=await call('explore',{sheet:currentSheet,row:$('row').value});
    if(serial!==explorationSerial || currentSheet!==$('sheet').value) return;
    rowChoices=response.rows;
    const filtered=filterRowOptions(response.selected);
    if(!filtered.some(r=>r.code===response.selected)){
      $('cells').replaceChildren();
      $('explore-caption').textContent='검색에 일치하는 행이 없습니다.';
      return;
    }
    cells=response.cells;
    currentUnit=response.unit;
    $('explore-caption').textContent=`${currentSheet} / ${response.selected} ${response.name}`;
    $('unit').textContent='단위: '+currentUnit;
    renderCells();
  }catch(error){
    if(serial===explorationSerial)setStatus('표 조회 오류: '+error.message,'bad');
  }
}
function renderCells(){
  const query=$('col-search').value.trim().toLowerCase();
  const filtered=cells.filter(x=>(x.code+' '+x.name).toLowerCase().includes(query));
  const fragment=document.createDocumentFragment();
  for(const value of filtered){
    const tr=document.createElement('tr');
    tr.append(element('td',value.code),element('td',value.name),
      element('td',value.value===null?'—':typeof value.value==='number'?fmtRaw(value.value):value.value,'numeric'));
    fragment.appendChild(tr);
  }
  if(!filtered.length){const tr=document.createElement('tr');const td=element('td','검색 결과 없음','empty');td.colSpan=3;tr.append(td);fragment.append(tr);}
  $('cells').replaceChildren(fragment);
}
function filterSectors(){
  const query=$('sector-search').value.trim().toLowerCase();
  const before=$('sector').value;
  const filtered=sectors.filter(x=>(x.code+' '+x.name).toLowerCase().includes(query));
  renderOptions($('sector'),filtered);
  if(filtered.some(x=>x.code===before))$('sector').value=before;
}
function renderBasket(){
  const list=$('basket');
  if(!basket.length){list.replaceChildren(element('li','추가된 산업이 없습니다.','empty'));$('sim-status').textContent='먼저 산업을 추가하세요.';return;}
  const fragment=document.createDocumentFragment();
  basket.forEach((item,index)=>{
    const li=document.createElement('li');
    li.append(element('span',`${item.code} · ${item.name} / ${fmt(item.amount/100)}억원`));
    const button=element('button','삭제');
    button.type='button';
    button.onclick=()=>{basket.splice(index,1);result=null;$('results').hidden=true;renderBasket();};
    li.append(button);fragment.appendChild(li);
  });
  list.replaceChildren(fragment);
  $('sim-status').textContent=`${basket.length}개 시나리오 준비 · 계산 버튼을 누르세요.`;
}
function displayResults(){
  if(!result)return;
  const values=[['최종수요 증가액',result.demand],['생산유발액',result.production],['부가가치유발액',result.valueAdded],['수입유발액',result.imports]];
  const cards=document.createDocumentFragment();
  for(const [name,number] of values){
    const card=element('div',undefined,'stat');
    card.append(element('p',name,'stat-title'),element('strong',fmt(number/100)),element('span','억원','stat-unit'));
    cards.append(card);
  }
  $('cards').replaceChildren(cards);
  $('multiplier').textContent=`생산유발배율: ${fmt(result.productionMultiplier)}배 · 최종수요 1억원 증가에 대해 총생산이 약 ${fmt(result.productionMultiplier)}억원 증가한다는 의미입니다(입력한 산업 구성에 대한 가중평균).`;
  const ranking=result.rows.slice(0,12), max=Math.max(...ranking.map(x=>x.production),0.0000001);
  const bars=document.createDocumentFragment();
  ranking.forEach(item=>{
    const row=element('div',undefined,'bar-row');
    const label=element('span',item.name,'bar-label');label.title=item.code+' '+item.name;
    const track=element('div',undefined,'bar-track');
    const fill=element('div',undefined,'bar-fill');fill.style.width=Math.max(0,item.production/max*100)+'%';
    track.append(fill);
    row.append(label,track,element('span',fmt(item.production/100),'bar-value'));
    bars.append(row);
  });
  $('bars').replaceChildren(bars);
  const rows=document.createDocumentFragment();
  result.rows.forEach(item=>{
    const tr=document.createElement('tr');
    for(const text of [item.code,item.name]) tr.append(element('td',text));
    for(const value of [item.production,item.valueAdded,item.imports])tr.append(element('td',fmt(value/100),'numeric'));
    rows.append(tr);
  });
  $('result-rows').replaceChildren(rows);
  $('results').hidden=false;
  $('results').scrollIntoView({behavior:'smooth',block:'nearest'});
}
function csvValue(value){
  let text=String(value);
  // Excel에서 문자형 값이 수식으로 실행되지 않도록 이스케이프.
  if(typeof value==='string' && /^[=+@\-\t\r]/.test(text))text="'"+text;
  return '"'+text.replaceAll('"','""')+'"';
}
function downloadResults(){
  if(!result)return;
  const content=[['분류','산업코드','산업명','최종수요증가액_억원','생산유발액_억원','부가가치유발액_억원','수입유발액_억원'],
    ...result.rows.map(x=>[cfg[activeLevel].label,x.code,x.name,x.finalDemand/100,x.production/100,x.valueAdded/100,x.imports/100])];
  const csv='\ufeff'+content.map(row=>row.map(csvValue).join(',')).join('\r\n');
  const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));
  const a=document.createElement('a');a.href=url;a.download=`io-2023-${activeLevel}-results.csv`;document.body.append(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),2000);
}

$('level').onchange=()=>{
  activeLevel=null;sectors=[];setReady(false);resetOutputs();
  setStatus('산업분류가 변경되었습니다. 새로운 Excel 파일을 불러오세요.');
};
$('load').onclick=loadWorkbook;
$('sheet').onchange=()=>{$('row').replaceChildren();$('row-search').value='';$('col-search').value='';explore();};
$('row').onchange=explore;
$('row-search').oninput=()=>{
  const previous=$('row').value;
  const filtered=filterRowOptions();
  if(!filtered.length){$('cells').replaceChildren();$('explore-caption').textContent='검색에 일치하는 행이 없습니다.';return;}
  if($('row').value!==previous) explore();
};
$('col-search').oninput=renderCells;
$('sector-search').oninput=filterSectors;
$('scenario').onsubmit=event=>{
  event.preventDefault();
  const code=$('sector').value;
  const sector=sectors.find(x=>x.code===code);
  const amountInHundredMillion=Number($('amount').value);
  if(!activeLevel||!sector||!Number.isFinite(amountInHundredMillion)||amountInHundredMillion<=0){
    $('sim-status').textContent='산업과 양의 최종수요 증가액(억원)을 확인하세요.';return;
  }
  basket.push({...sector, amount:amountInHundredMillion*100});
  result=null;$('results').hidden=true;renderBasket();
};
$('clear').onclick=()=>{basket=[];result=null;$('results').hidden=true;renderBasket();};
$('run').onclick=async()=>{
  if(!activeLevel||!basket.length){$('sim-status').textContent='분석할 산업을 먼저 추가하세요.';return;}
  const snapshot=basket.map(({code,amount})=>({code,amount}));
  const check=JSON.stringify(snapshot);
  $('run').disabled=true;$('sim-status').textContent='Python으로 생산·부가가치·수입 유발액 계산 중…';
  try{
    const response=await call('simulate',snapshot);
    if(check!==JSON.stringify(basket.map(({code,amount})=>({code,amount})))){
      $('sim-status').textContent='계산 중 입력이 변경되었습니다. 다시 계산하세요.';return;
    }
    result=response;
    displayResults();
    $('sim-status').textContent='계산 완료 · 아래 결과는 현재 시나리오의 분석값입니다.';
  }catch(error){
    result=null;$('results').hidden=true;
    $('sim-status').textContent='계산 실패: '+error.message;
  }finally{$('run').disabled=!activeLevel;}
};
$('download').onclick=downloadResults;
setReady(false);
