const FILES = {
  '2G': 'https://docs.google.com/spreadsheets/d/1zwa8F_WrJS9LXArcNJqmemn7FK96Ycnu/export?format=xlsx',
  '3G': 'https://docs.google.com/spreadsheets/d/1HCI7IuWjMle50E-TRAz1cbo-yZLaUPF-/export?format=xlsx',
  '4G': 'https://docs.google.com/spreadsheets/d/13q7CDdLC0Hy4lmgyu9PY-EA-hBgDGFpj/export?format=xlsx',
  '5G': 'https://docs.google.com/spreadsheets/d/1Ff7NnCsDQl0YdbDxvEEa82rrFtYzojVB/export?format=xlsx'
};
const TECHS = ['2G', '3G', '4G', '5G'];
let rows = [];
let lastData = null;

const $ = id => document.getElementById(id);
const clean = value => value == null || String(value).toLowerCase() === 'nan' ? '' : String(value).trim();
const col = (row, name) => clean(row[name] ?? row[name.trim()]);

async function loadBases() {
  $('db-status').textContent = 'Baixando 4 bases...';
  const loaded = [];
  for (const tech of TECHS) {
    const response = await fetch(FILES[tech], {cache: 'no-store'});
    if (!response.ok) throw new Error(`${tech}: download HTTP ${response.status}`);
    const workbook = XLSX.read(await response.arrayBuffer(), {type: 'array', cellDates: false});
    for (const sheet of workbook.SheetNames) {
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {defval: ''});
      data.forEach(item => rows.push({...item, __tech: sheet === tech ? tech : sheet}));
    }
    loaded.push(tech);
    $('db-status').textContent = `Carregadas: ${loaded.join(', ')}`;
  }
  rows = rows.filter(row => TECHS.includes(row.__tech) && col(row, '[P]SITE') && col(row, '[P]UF'));
  $('db-status').textContent = `${rows.length.toLocaleString('pt-BR')} registros · ${loaded.join(', ')}`;
  $('message').textContent = 'Digite o Site e o UF para iniciar a consulta.';
}

function selectedTechs() { return [...document.querySelectorAll('.techs input:checked')].map(input => input.value); }
function valueFor(row, field, tech) {
  if (field === 'earfcn') return tech === '3G' ? col(row, '[P]DL_UARFCN') : tech === '4G' || tech === '5G' ? col(row, '[P]DL_EARFCN') : '';
  return col(row, field);
}
function unique(values) { return [...new Set(values.filter(Boolean))]; }
function consolidate(found) {
  const result = {cidade:'',bairro:'',endereco:'',latitude:'',longitude:'',tech:{}};
  for (const tech of TECHS) {
    const techRows = found.filter(row => row.__tech === tech);
    result.tech[tech] = {bands:[], azimuth:[], codes:[], rows:techRows};
    techRows.forEach(row => {
      result.cidade ||= col(row,'[P]CIDADE'); result.bairro ||= col(row,'[P]BAIRRO');
      result.latitude ||= col(row,'[P]LATITUDE'); result.longitude ||= col(row,'[P]LONGITUDE'); result.endereco ||= col(row,'[P]ENDERECO');
      const banda=valueFor(row,'[P]BANDA_OPERACAO',tech), earfcn=valueFor(row,'earfcn',tech);
      if (banda && !/iot/i.test(banda) && !result.tech[tech].bands.some(x=>x.banda===banda&&x.earfcn===earfcn)) result.tech[tech].bands.push({banda,earfcn});
      const az=valueFor(row,'[P]AZIMUTH',tech); if(az&&!result.tech[tech].azimuth.includes(az)) result.tech[tech].azimuth.push(az);
      const code=tech==='2G'?col(row,'[P]BCCH'):tech==='3G'?col(row,'[P]PSC'):col(row,'[P]PCI'); if(code&&!result.tech[tech].codes.includes(code)) result.tech[tech].codes.push(code);
    });
  }
  return result;
}
function item(label, value) { return `<div class="item"><small>${label}</small><strong>${value || '—'}</strong></div>`; }
function render(data, site, uf) {
  lastData=data; $('copy').disabled=false; let html=`<div class="card"><h3>IDENTIFICAÇÃO DO SITE</h3><div class="grid">${item('SITE',site)}${item('UF',uf)}${item('CIDADE',data.cidade)}${item('BAIRRO',data.bairro)}${item('ENDEREÇO',data.endereco)}${item('LATITUDE',data.latitude)}${item('LONGITUDE',data.longitude)}</div></div>`;
  for(const tech of selectedTechs()){const t=data.tech[tech]; if(!t.rows.length)continue; const bands=t.bands.map(x=>`${x.banda}${x.earfcn?` / ${x.earfcn}`:''}`).join('<br>'); html+=`<div class="card tech-section"><h3>${tech}</h3><div class="grid">${item('BANDA / EARFCN',bands)}${item('AZIMUTH',t.azimuth.join(', '))}${item(tech==='2G'?'BCCH':tech==='3G'?'PSC':'PCI',t.codes.join(', '))}</div></div>`;}
  $('results').innerHTML=html; $('message').textContent=`${site} · ${uf} · consulta concluída`;
}
function historySave(site,uf){const key='base-vivo-history', list=JSON.parse(localStorage.getItem(key)||'[]').filter(x=>x.site!==site||x.uf!==uf);list.unshift({site,uf,ts:new Date().toLocaleString('pt-BR')});localStorage.setItem(key,JSON.stringify(list.slice(0,20)));renderHistory();}
function renderHistory(){const list=JSON.parse(localStorage.getItem('base-vivo-history')||'[]');$('history').innerHTML=list.length?list.map(x=>`<button class="history-item" data-site="${x.site}" data-uf="${x.uf}">${x.site} · ${x.uf}<small>${x.ts}</small></button>`).join(''):'<span class="empty">Nenhuma consulta.</span>';document.querySelectorAll('.history-item').forEach(b=>b.onclick=()=>{$('site').value=b.dataset.site;$('uf').value=b.dataset.uf;$('search-form').requestSubmit()});}
$('search-form').onsubmit=e=>{e.preventDefault();const site=$('site').value.trim().toUpperCase(),uf=$('uf').value.trim().toUpperCase(),techs=selectedTechs();const found=rows.filter(r=>col(r,'[P]SITE').toUpperCase()===site&&col(r,'[P]UF').toUpperCase()===uf&&techs.includes(r.__tech));if(!found.length){$('results').innerHTML='';$('message').className='message error';$('message').textContent='Nenhum registro encontrado para os filtros informados.';return;}$('message').className='message';render(consolidate(found),site,uf);historySave(site,uf);};
$('reload').onclick=async()=>{rows=[];$('results').innerHTML='';try{await loadBases()}catch(e){$('db-status').textContent='Falha ao carregar bases';$('message').className='message error';$('message').textContent=`Não foi possível carregar as planilhas: ${e.message}`;}};
$('copy').onclick=async()=>{if(!lastData)return;try{await navigator.clipboard.writeText($('results').innerText);$('copy').textContent='COPIADO';setTimeout(()=>$('copy').textContent='COPIAR',1500)}catch{alert('Não foi possível copiar neste navegador.')}};
renderHistory();loadBases().catch(e=>{$('db-status').textContent='Falha ao carregar bases';$('message').className='message error';$('message').textContent=`Não foi possível carregar as planilhas: ${e.message}`});
