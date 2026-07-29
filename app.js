const FILES = {
  '2G': 'https://docs.google.com/spreadsheets/d/1zwa8F_WrJS9LXArcNJqmemn7FK96Ycnu/export?format=xlsx',
  '3G': 'https://docs.google.com/spreadsheets/d/1HCI7IuWjMle50E-TRAz1cbo-yZLaUPF-/export?format=xlsx',
  '4G': 'https://docs.google.com/spreadsheets/d/13q7CDdLC0Hy4lmgyu9PY-EA-hBgDGFpj/export?format=xlsx',
  '5G': 'https://docs.google.com/spreadsheets/d/1Ff7NnCsDQl0YdbDxvEEa82rrFtYzojVB/export?format=xlsx'
};
const SHEETS = {
  '2G': '1zwa8F_WrJS9LXArcNJqmemn7FK96Ycnu',
  '3G': '1HCI7IuWjMle50E-TRAz1cbo-yZLaUPF-',
  '4G': '13q7CDdLC0Hy4lmgyu9PY-EA-hBgDGFpj',
  '5G': '1Ff7NnCsDQl0YdbDxvEEa82rrFtYzojVB'
};
const TECHS = ['2G', '3G', '4G', '5G'];
let rows = [];
let lastData = null;
let localDb = null;
const CACHE_DB = 'base-vivo-cache';
const CACHE_STORE = 'datasets';

const $ = id => document.getElementById(id);
const clean = value => value == null || String(value).toLowerCase() === 'nan' ? '' : String(value).trim();
const col = (row, name) => clean(row[name] ?? row[name.trim()]);
const TARGETS_4G = {
  '700': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '1800': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '2100': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '2600': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '2300': {'tdd':['DL > 20 Mbps','UL >  2 Mbps']}
};
const TARGETS_5G = {'3500':['DL > 200 Mbps','UL > 30 Mbps'],'2300':['DL >  80 Mbps','UL > 10 Mbps'],'2100':['DL > 200 Mbps','UL > 30 Mbps']};
function target4g(banda,bw,mimo){for(const freq of Object.keys(TARGETS_4G)){if(!banda.includes(freq))continue;const table=TARGETS_4G[freq];if(freq==='2300')return table.tdd;const mimoKey=(Number(mimo)||0)>=3?4:2;return table[`${Math.round(Number(bw)||0)}/${mimoKey}`]||null;}return null;}
function target5g(banda){for(const freq of Object.keys(TARGETS_5G))if(banda.includes(freq))return TARGETS_5G[freq];return null;}

function openCache() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CACHE_DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(CACHE_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
async function readCache() {
  try {
    const db = await openCache();
    return await new Promise((resolve, reject) => {
      const request = db.transaction(CACHE_STORE).objectStore(CACHE_STORE).get('rows');
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  } catch { return []; }
}
async function writeCache(data) {
  try {
    const db = await openCache();
    await new Promise((resolve, reject) => {
      const request = db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(data, 'rows');
      request.onsuccess = resolve;
      request.onerror = () => reject(request.error);
    });
  } catch { /* cache indisponível não impede a consulta */ }
}
async function loadBases(force = false) {
  if (!force) {
    const cached = await readCache();
    if (cached.length) {
      rows = cached;
      $('db-status').textContent = `${rows.length.toLocaleString('pt-BR')} registros · cache local`;
      $('message').textContent = 'Digite o Site e o UF para iniciar a consulta.';
      return;
    }
  }
  $('db-status').textContent = 'Baixando 4 bases...';
  const loaded = [];
  for (const tech of TECHS) {
    const response = await fetch(FILES[tech], {cache: 'no-store'});
    if (!response.ok) throw new Error(`${tech}: download HTTP ${response.status}`);
    const workbook = XLSX.read(await response.arrayBuffer(), {type: 'array', cellDates: false});
    for (const sheet of workbook.SheetNames) {
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {defval: ''});
      data.forEach(item => rows.push({...item, __tech: tech}));
    }
    loaded.push(tech);
    $('db-status').textContent = `Carregadas: ${loaded.join(', ')}`;
  }
  rows = rows.filter(row => TECHS.includes(row.__tech) && col(row, '[P]SITE'));
  await writeCache(rows);
  $('db-status').textContent = `${rows.length.toLocaleString('pt-BR')} registros · ${loaded.join(', ')}`;
  $('message').textContent = 'Digite o Site e o UF para iniciar a consulta.';
}
function parseGviz(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('resposta inválida do Google Sheets');
  return JSON.parse(text.slice(start, end + 1));
}
async function querySheet(tech, site, uf) {
  const query = encodeURIComponent(`select * where A = '${site.replace(/'/g, "''")}'`);
  const url = `https://docs.google.com/spreadsheets/d/${SHEETS[tech]}/gviz/tq?tqx=out:json&sheet=Export&tq=${query}`;
  const response = await fetch(url, {cache: 'no-store'});
  if (!response.ok) throw new Error(`${tech}: consulta HTTP ${response.status}`);
  const payload = parseGviz(await response.text());
  const columns = payload.table.cols.map(column => column.label || '');
  return payload.table.rows.map(item => {
    const row = {__tech: tech};
    columns.forEach((name, index) => { if (name) row[name] = item.c[index]?.v ?? ''; });
    return row;
  }).filter(row => col(row, '[P]UF').toUpperCase() === uf);
}
async function loadLocalDatabase() {
  try {
    const db = await openCache();
    let bytes = await new Promise((resolve, reject) => {
      const request = db.transaction(CACHE_STORE).objectStore(CACHE_STORE).get('database-v2');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    if (!bytes) {
      const response = await fetch('rf_cache.db', {cache: 'no-store'});
      if (!response.ok) throw new Error('rf_cache.db não publicado');
      bytes = await response.arrayBuffer();
      db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(bytes, 'database-v2');
    }
    const SQL = await initSqlJs({locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${file}`});
    localDb = new SQL.Database(new Uint8Array(bytes));
    $('db-status').textContent = 'Base local pronta · consulta rápida';
  } catch {
    $('db-status').textContent = 'Consulta online · cache por site/UF';
  }
}
function queryLocal(site, uf, techs) {
  if (!localDb) return null;
  const placeholders = techs.map(() => '?').join(',');
  const result = localDb.exec(`SELECT tech,banda,azimuth,bcch,psc,pci,bandwidth,cidade,bairro,endereco,earfcn,latitude,longitude,mimo FROM rf WHERE site = ? AND uf = ? AND tech IN (${placeholders})`, [site, uf, ...techs]);
  if (!result.length) return [];
  const columns = result[0].columns;
  return result[0].values.map(values => Object.fromEntries(columns.map((name, index) => [name, values[index] ?? '']))).map(row => ({
    ...row,
    __tech: row.tech,
    '[P]SITE': row.site,
    '[P]UF': row.uf,
    '[P]BANDA_OPERACAO': row.banda,
    '[P]AZIMUTH': row.azimuth,
    '[P]BCCH': row.bcch,
    '[P]PSC': row.psc,
    '[P]PCI': row.pci,
    '[P]BANDWIDTH': row.bandwidth,
    '[P]CIDADE': row.cidade,
    '[P]BAIRRO': row.bairro,
    '[P]ENDERECO': row.endereco,
    '[P]LATITUDE': row.latitude,
    '[P]LONGITUDE': row.longitude,
    '[P]MIMO': row.mimo,
    '[P]DL_UARFCN': row.tech === '3G' ? row.earfcn : '',
    '[P]DL_EARFCN': ['4G', '5G'].includes(row.tech) ? row.earfcn : ''
  }));
}
async function queryRemote(site, uf, techs) {
  const local = queryLocal(site, uf, techs);
  if (local) return local;
  const key = `query:${site}:${uf}:${techs.join(',')}`;
  try {
    const db = await openCache();
    const cached = await new Promise((resolve, reject) => {
      const request = db.transaction(CACHE_STORE).objectStore(CACHE_STORE).get(key);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    if (cached?.length) return cached;
  } catch { /* consulta remota continua normalmente */ }
  const result = (await Promise.all(techs.map(tech => querySheet(tech, site, uf)))).flat();
  try {
    const db = await openCache();
    db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(result, key);
  } catch { /* cache indisponível não impede a consulta */ }
  return result;
}

function selectedTechs() { return [...document.querySelectorAll('.techs input:checked')].map(input => input.value); }
function valueFor(row, field, tech) {
  if (field === 'earfcn') return tech === '3G' ? col(row, '[P]DL_UARFCN') : tech === '4G' || tech === '5G' ? col(row, '[P]DL_EARFCN') : '';
  return col(row, field);
}
function unique(values) { return [...new Set(values.filter(Boolean))]; }
function consolidate(found) {
  const result = {cidade:'',bairro:'',endereco:'',latitude:'',longitude:'',tech:{},targets:[]};
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
      if(tech==='4G' && banda && !/iot/i.test(banda)){const bw=col(row,'[P]BANDWIDTH'), mimo=col(row,'[P]MIMO'), target=target4g(banda,bw,mimo);if(target&&!result.targets.some(x=>x.tech===tech&&x.label===`${banda} / ${earfcn}`))result.targets.push({tech,label:`${banda}${earfcn?` / ${earfcn}`:''}`,bw:bw||'—',mimo:mimo?`${Number(mimo)>=3?4:2}x${Number(mimo)>=3?4:2}`:'—',dl:target[0],ul:target[1]});}
      if(tech==='5G' && banda){const target=target5g(banda);if(target&&!result.targets.some(x=>x.tech===tech&&x.label===`${banda} / ${earfcn}`))result.targets.push({tech,label:`${banda}${earfcn?` / ${earfcn}`:''}`,bw:'—',mimo:'—',dl:target[0],ul:target[1]});}
    });
  }
  return result;
}
function item(label, value) { return `<div class="item"><small>${label}</small><strong>${value || '—'}</strong></div>`; }
function mapsItem(latitude, longitude) {
  if (!latitude || !longitude) return item('COORDENADAS', '—');
  const query = `${latitude},${longitude}`;
  return `<div class="item"><small>COORDENADAS</small><strong>${query}</strong><a class="maps-link" href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}" target="_blank" rel="noopener">VER NO GOOGLE MAPS</a></div>`;
}
function render(data, site, uf) {
  lastData=data; $('copy').disabled=false; let html=`<div class="card"><h3>SITE: ${site} - ${uf}</h3><div class="grid identity">${item('CIDADE',data.cidade)}${item('BAIRRO',data.bairro)}${item('ENDEREÇO',data.endereco)}${mapsItem(data.latitude,data.longitude)}</div></div>`;
  html+='<div class="card bands"><h3>BANDA DE OPERAÇÃO / EARFCN</h3><div class="band-grid">';for(const tech of selectedTechs()){const t=data.tech[tech];if(t.rows.length)html+=`<div class="band-column"><b>${tech}</b>${t.bands.map(x=>`<span>${x.banda}${x.earfcn?` / <em>${x.earfcn}</em>`:''}</span>`).join('')}</div>`;}html+='</div></div>';
  const active=selectedTechs().filter(tech=>data.tech[tech].rows.length), maxCodes=Math.max(...active.map(tech=>data.tech[tech].codes.length),0), maxAz=Math.max(...active.map(tech=>data.tech[tech].azimuth.length),0);html+='<div class="section-title">CÓDIGOS DE CÉLULA / AZIMUTH</div><div class="boxes"><div class="card table-box"><h3>CÓDIGOS DE CÉLULA</h3><table><thead><tr>'+active.map(t=>`<th>${t==='2G'?'BCCH':t==='3G'?'PSCs':'PCIs'}<br>${t}</th>`).join('')+'</tr></thead><tbody>'+Array.from({length:maxCodes},(_,i)=>`<tr>${active.map(t=>`<td>${data.tech[t].codes[i]||''}</td>`).join('')}</tr>`).join('')+'</tbody></table></div><div class="card table-box"><h3>AZIMUTH</h3><table><thead><tr>'+active.map(t=>`<th>${t}</th>`).join('')+'</tr></thead><tbody>'+Array.from({length:Math.max(maxAz,maxCodes)},(_,i)=>`<tr>${active.map(t=>`<td>${data.tech[t].azimuth[i]||''}</td>`).join('')}</tr>`).join('')+'</tbody></table></div></div>';
  if(data.targets.length)html+='<div class="card target-card"><h3>BANDWIDTH / TARGET VIVO</h3><div class="target-scroll"><table><thead><tr><th>TECH</th><th>BANDA</th><th>BW</th><th>MIMO</th><th>DL TARGET</th><th>UL TARGET</th></tr></thead><tbody>'+data.targets.map(x=>`<tr><td>${x.tech}</td><td>${x.label}</td><td>${x.bw}</td><td>${x.mimo}</td><td class="dl">${x.dl}</td><td class="ul">${x.ul}</td></tr>`).join('')+'</tbody></table></div></div>';
  $('results').innerHTML=html; $('message').textContent=`${site} · ${uf} · consulta concluída`;
}
function historySave(site,uf){const key='base-vivo-history', list=JSON.parse(localStorage.getItem(key)||'[]').filter(x=>x.site!==site||x.uf!==uf);list.unshift({site,uf,ts:new Date().toLocaleString('pt-BR')});localStorage.setItem(key,JSON.stringify(list.slice(0,20)));renderHistory();}
function renderHistory(){const list=JSON.parse(localStorage.getItem('base-vivo-history')||'[]');$('history').innerHTML=list.length?list.map(x=>`<button class="history-item" data-site="${x.site}" data-uf="${x.uf}">${x.site} · ${x.uf}<small>${x.ts}</small></button>`).join(''):'<span class="empty">Nenhuma consulta.</span>';document.querySelectorAll('.history-item').forEach(b=>b.onclick=()=>{$('site').value=b.dataset.site;$('uf').value=b.dataset.uf;$('search-form').requestSubmit()});}
$('search-form').onsubmit=async e=>{e.preventDefault();const site=$('site').value.trim().toUpperCase(),uf=$('uf').value.trim().toUpperCase(),techs=selectedTechs();if(!site||!uf||!techs.length)return;$('message').className='message';$('message').textContent='Consultando as bases...';try{const found=await queryRemote(site,uf,techs);if(!found.length){$('results').innerHTML='';$('message').className='message error';$('message').textContent='Nenhum registro encontrado para os filtros informados.';return;}render(consolidate(found),site,uf);historySave(site,uf);}catch(error){$('message').className='message error';$('message').textContent=`Não foi possível consultar as bases: ${error.message}`;}};
$('reload').onclick=async()=>{try{const db=await openCache();db.transaction(CACHE_STORE,'readwrite').objectStore(CACHE_STORE).clear();}catch{}localDb=null;location.reload();};
$('copy').onclick=async()=>{if(!lastData)return;try{await navigator.clipboard.writeText($('results').innerText);$('copy').textContent='COPIADO';setTimeout(()=>$('copy').textContent='COPIAR',1500)}catch{alert('Não foi possível copiar neste navegador.')}};
renderHistory();rows=[];$('db-status').textContent='Carregando base local...';$('message').textContent='Digite o Site e o UF para iniciar a consulta.';loadLocalDatabase();
