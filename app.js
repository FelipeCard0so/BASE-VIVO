/* ============================================================================
   BASE VIVO — app.js
   Projeto: Consulta de sites Vivo (2G/3G/4G/5G)
   Autor:   Felipe Cardoso

   Visão geral do funcionamento:
     1. Ao carregar, tenta abrir uma base local SQLite (rf_cache.db) via
        sql.js/WASM — é o caminho "rápido" de consulta (ver seção 6).
     2. Se a base local não existir, cai para consulta remota via Google
        Sheets (GVIZ), cacheada em IndexedDB por combinação site+UF+tecnologias
        (ver seções 5 e 7).
     3. Existe também um caminho alternativo de carregar TODAS as planilhas
        via export XLSX (loadBases) — não é chamado no fluxo atual de
        inicialização, mas fica disponível para uso/depuração.
     4. O resultado é "consolidado" (consolidate) e transformado em HTML
        (render) dentro de #results.

   Convenção adotada:
     - Nomes de função em camelCase, verbo + substantivo (ex: loadBases,
       queryRemote, renderHistory).
     - Prefixo "[P]" nos nomes de coluna vem das planilhas de origem —
       não alterar esses literais, eles são o contrato com a base de dados.
     - Comentários em blocos numerados, iguais aos usados em styles.css e
       index.html, para manter os três arquivos com a mesma "linguagem".
   ============================================================================ */


/* ============================================================================
   1. CONFIGURAÇÃO / CONSTANTES
   ----------------------------------------------------------------------------
   Fontes de dados e estado global da aplicação.
   ============================================================================ */

// URLs de export direto (.xlsx) de cada planilha por tecnologia.
// Usadas apenas pelo fluxo alternativo loadBases() (ver seção 5).
const FILES = {
  '2G': 'https://docs.google.com/spreadsheets/d/1zwa8F_WrJS9LXArcNJqmemn7FK96Ycnu/export?format=xlsx',
  '3G': 'https://docs.google.com/spreadsheets/d/1HCI7IuWjMle50E-TRAz1cbo-yZLaUPF-/export?format=xlsx',
  '4G': 'https://docs.google.com/spreadsheets/d/13q7CDdLC0Hy4lmgyu9PY-EA-hBgDGFpj/export?format=xlsx',
  '5G': 'https://docs.google.com/spreadsheets/d/1Ff7NnCsDQl0YdbDxvEEa82rrFtYzojVB/export?format=xlsx'
};

// IDs das mesmas planilhas, usados nas consultas GVIZ (querySheet),
// que trazem só as linhas do site pesquisado em vez do arquivo inteiro.
const SHEETS = {
  '2G': '1zwa8F_WrJS9LXArcNJqmemn7FK96Ycnu',
  '3G': '1HCI7IuWjMle50E-TRAz1cbo-yZLaUPF-',
  '4G': '13q7CDdLC0Hy4lmgyu9PY-EA-hBgDGFpj',
  '5G': '1Ff7NnCsDQl0YdbDxvEEa82rrFtYzojVB'
};

// Lista de tecnologias suportadas, na ordem em que devem ser exibidas.
const TECHS = ['2G', '3G', '4G', '5G'];

// --- Estado global da aplicação -------------------------------------------
let rows = [];       // linhas carregadas via loadBases() (fluxo XLSX completo)
let lastData = null;  // último resultado consolidado (usado pelo botão COPIAR)
let localDb = null;   // instância do banco SQLite local (sql.js), quando disponível

// Nome do banco/objeto usados no IndexedDB para cache local
// (linhas do loadBases, base .db local e resultados de consultas remotas).
const CACHE_DB = 'base-vivo-cache';
const CACHE_STORE = 'datasets';


/* ============================================================================
   2. HELPERS GERAIS
   ----------------------------------------------------------------------------
   Funções pequenas e reaproveitadas em várias partes do arquivo.
   ============================================================================ */

// Atalho para document.getElementById — evita repetir o nome inteiro toda hora.
const $ = id => document.getElementById(id);

/**
 * Normaliza um valor vindo da planilha/base:
 * - null/undefined ou a string "nan" (comum em export de Excel/pandas) viram ''
 * - qualquer outro valor é convertido para string e tem espaços nas pontas removidos
 */
const clean = value => value == null || String(value).toLowerCase() === 'nan' ? '' : String(value).trim();

/**
 * Lê uma coluna de uma linha (row) já normalizada pelo clean().
 * Tenta primeiro o nome exato e, se não encontrar, o nome com espaços
 * extras removidos — cobre inconsistências de cabeçalho nas planilhas.
 */
const col = (row, name) => clean(row[name] ?? row[name.trim()]);


/* ============================================================================
   3. TABELAS DE TARGET (BANDWIDTH / DL-UL ESPERADO)
   ----------------------------------------------------------------------------
   Valores de referência de qualidade (DL/UL target) por banda/frequência,
   usados no card "BANDWIDTH / TARGET VIVO" do resultado.
   ============================================================================ */

// Target 4G por frequência (700/1800/2100/2600/2300), combinando
// largura de banda (BW, em MHz) e configuração de MIMO ("BW/MIMO").
// Ex.: TARGETS_4G['700']['20/4'] = [DL, UL] para banda de 700MHz,
// 20MHz de BW e MIMO 4x4.
const TARGETS_4G = {
  '700': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '1800': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '2100': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  '2600': {'20/4':['DL > 40 Mbps','UL > 15 Mbps'],'20/2':['DL > 30 Mbps','UL > 15 Mbps'],'15/4':['DL > 30 Mbps','UL > 10 Mbps'],'15/2':['DL > 25 Mbps','UL > 10 Mbps'],'10/4':['DL > 25 Mbps','UL >  5 Mbps'],'10/2':['DL > 20 Mbps','UL >  5 Mbps']},
  // 2300 é banda TDD (não tem combinação BW/MIMO como as demais — chave única "tdd")
  '2300': {'tdd':['DL > 20 Mbps','UL >  2 Mbps']}
};

// Target 5G por frequência — mais simples: uma faixa DL/UL fixa por banda.
const TARGETS_5G = {'3500':['DL > 200 Mbps','UL > 30 Mbps'],'2300':['DL >  80 Mbps','UL > 10 Mbps'],'2100':['DL > 200 Mbps','UL > 30 Mbps']};

/**
 * Calcula o target 4G para uma banda/BW/MIMO informados.
 * @param {string} banda - texto da banda de operação (ex: "Z-700")
 * @param {string|number} bw   - largura de banda em MHz
 * @param {string|number} mimo - camadas de MIMO (2 ou 4)
 * @returns {[string,string]|null} par [DL target, UL target] ou null se não achar
 */
function target4g(banda,bw,mimo){
  for(const freq of Object.keys(TARGETS_4G)){
    if(!banda.includes(freq)) continue; // banda não é dessa frequência, pula
    const table = TARGETS_4G[freq];
    if(freq === '2300') return table.tdd; // 2300 é caso especial (TDD)
    // MIMO >= 3 camadas é tratado como 4x4; caso contrário, 2x2
    const mimoKey = (Number(mimo)||0) >= 3 ? 4 : 2;
    return table[`${Math.round(Number(bw)||0)}/${mimoKey}`] || null;
  }
  return null;
}

/**
 * Calcula o target 5G para uma banda informada.
 * @param {string} banda - texto da banda de operação
 * @returns {[string,string]|null} par [DL target, UL target] ou null se não achar
 */
function target5g(banda){
  for(const freq of Object.keys(TARGETS_5G))
    if(banda.includes(freq)) return TARGETS_5G[freq];
  return null;
}


/* ============================================================================
   4. CACHE LOCAL (INDEXEDDB)
   ----------------------------------------------------------------------------
   Camada genérica de cache usada tanto pelo fluxo XLSX (loadBases) quanto
   pelo banco SQLite local (loadLocalDatabase) e pelas consultas remotas
   (queryRemote). Um único object store ("datasets") guarda tudo, indexado
   por chave (ex: 'rows', 'database-v2', 'query:SITE:UF:TECHS').
   ============================================================================ */

/** Abre (ou cria, na primeira vez) o banco IndexedDB de cache. */
function openCache() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CACHE_DB, 1);
    // onupgradeneeded roda apenas na criação/mudança de versão do banco
    request.onupgradeneeded = () => request.result.createObjectStore(CACHE_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/** Lê a lista de linhas cacheadas pelo fluxo XLSX (chave fixa 'rows'). */
async function readCache() {
  try {
    const db = await openCache();
    return await new Promise((resolve, reject) => {
      const request = db.transaction(CACHE_STORE).objectStore(CACHE_STORE).get('rows');
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  } catch { return []; } // se IndexedDB falhar, segue sem cache (não trava a app)
}

/** Grava a lista de linhas do fluxo XLSX em cache (chave fixa 'rows'). */
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


/* ============================================================================
   5. CARREGAMENTO COMPLETO VIA XLSX (FLUXO ALTERNATIVO)
   ----------------------------------------------------------------------------
   Baixa as 4 planilhas inteiras (.xlsx) e monta o array `rows` em memória.
   Mais pesado que o fluxo local/GVIZ — não é chamado na inicialização
   atual (ver seção 9), mas continua disponível caso seja necessário
   recarregar a base inteira de uma vez.
   ============================================================================ */
async function loadBases(force = false) {
  // Se não for forçado, tenta usar o que já está em cache antes de baixar tudo de novo
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

  // Baixa e processa cada planilha (2G, 3G, 4G, 5G), sequencialmente,
  // atualizando o status a cada tecnologia concluída.
  for (const tech of TECHS) {
    const response = await fetch(FILES[tech], {cache: 'no-store'});
    if (!response.ok) throw new Error(`${tech}: download HTTP ${response.status}`);

    const workbook = XLSX.read(await response.arrayBuffer(), {type: 'array', cellDates: false});
    for (const sheet of workbook.SheetNames) {
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {defval: ''});
      // Marca cada linha com a tecnologia de origem (__tech), necessário
      // porque as planilhas não trazem essa informação por si só.
      data.forEach(item => rows.push({...item, __tech: tech}));
    }

    loaded.push(tech);
    $('db-status').textContent = `Carregadas: ${loaded.join(', ')}`;
  }

  // Mantém só linhas de tecnologias conhecidas e que tenham um Site preenchido
  rows = rows.filter(row => TECHS.includes(row.__tech) && col(row, '[P]SITE'));
  await writeCache(rows);

  $('db-status').textContent = `${rows.length.toLocaleString('pt-BR')} registros · ${loaded.join(', ')}`;
  $('message').textContent = 'Digite o Site e o UF para iniciar a consulta.';
}


/* ============================================================================
   6. CONSULTA REMOTA VIA GOOGLE SHEETS (GVIZ)
   ----------------------------------------------------------------------------
   Em vez de baixar a planilha inteira, usa a API "gviz" do Google Sheets
   para trazer só as linhas do site pesquisado — bem mais leve que o
   fluxo XLSX completo da seção 5.
   ============================================================================ */

/**
 * A resposta da API gviz não é JSON puro — vem envolvida em um comentário
 * JS (ex: "google.visualization.Query.setResponse({...})"). Esta função
 * extrai só o trecho entre a primeira "{" e a última "}" e faz o parse.
 */
function parseGviz(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('resposta inválida do Google Sheets');
  return JSON.parse(text.slice(start, end + 1));
}

/**
 * Consulta a planilha de uma tecnologia filtrando pelo Site (coluna A),
 * via query no estilo SQL suportada pela API gviz. O filtro de UF é
 * aplicado depois, no próprio JS, pois a planilha "Export" não indexa por UF.
 */
async function querySheet(tech, site, uf) {
  // Escapa aspas simples do site para não quebrar a query (SQL-like) da API
  const query = encodeURIComponent(`select * where A = '${site.replace(/'/g, "''")}'`);
  const url = `https://docs.google.com/spreadsheets/d/${SHEETS[tech]}/gviz/tq?tqx=out:json&sheet=Export&tq=${query}`;

  const response = await fetch(url, {cache: 'no-store'});
  if (!response.ok) throw new Error(`${tech}: consulta HTTP ${response.status}`);

  const payload = parseGviz(await response.text());
  const columns = payload.table.cols.map(column => column.label || '');

  // Reconstrói cada linha como objeto {nomeDaColuna: valor}, marcando a
  // tecnologia de origem, e filtra pela UF pedida.
  return payload.table.rows.map(item => {
    const row = {__tech: tech};
    columns.forEach((name, index) => { if (name) row[name] = item.c[index]?.v ?? ''; });
    return row;
  }).filter(row => col(row, '[P]UF').toUpperCase() === uf);
}


/* ============================================================================
   7. BASE LOCAL SQLite (sql.js / WASM) — CAMINHO PRINCIPAL DE CONSULTA
   ----------------------------------------------------------------------------
   Carrega um arquivo .db (rf_cache.db) publicado junto do site, roda
   inteiramente no navegador via WebAssembly (sql.js) e permite consultas
   instantâneas sem round-trip ao Google Sheets. É o que dá o status
   "Base local pronta · consulta rápida" na sidebar.
   ============================================================================ */

/**
 * Carrega o banco SQLite local:
 * 1. Tenta pegar os bytes do .db já cacheados no IndexedDB (chave 'database-v2').
 * 2. Se não tiver, baixa rf_cache.db do próprio site e salva no cache.
 * 3. Inicializa o motor sql.js (WASM, via CDN) e abre o banco em memória.
 * Se qualquer etapa falhar, cai silenciosamente para o modo "consulta
 * online" (ver queryRemote / querySheet).
 */
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
      // Salva em cache para não precisar baixar de novo na próxima visita
      db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(bytes, 'database-v2');
    }

    // initSqlJs vem do script sql-wasm.js carregado no index.html
    const SQL = await initSqlJs({locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${file}`});
    localDb = new SQL.Database(new Uint8Array(bytes));
    $('db-status').textContent = 'Base local pronta · consulta rápida';
  } catch {
    // Sem banco local disponível: aplicação continua funcionando via GVIZ
    $('db-status').textContent = 'Consulta online · cache por site/UF';
  }
}

/**
 * Consulta o banco SQLite local pela tabela `rf`, filtrando por site,
 * UF e lista de tecnologias selecionadas.
 * @returns {Array|null} linhas encontradas (já no "formato planilha", com
 *   os nomes de coluna "[P]..." usados pelo resto do app), ou `null` se
 *   não houver banco local carregado (nesse caso quem chamou deve cair
 *   para a consulta remota).
 */
function queryLocal(site, uf, techs) {
  if (!localDb) return null;

  const placeholders = techs.map(() => '?').join(',');
  const result = localDb.exec(
    `SELECT tech,banda,azimuth,bcch,psc,pci,bandwidth,cidade,bairro,endereco,earfcn,latitude,longitude,mimo
     FROM rf WHERE site = ? AND uf = ? AND tech IN (${placeholders})`,
    [site, uf, ...techs]
  );
  if (!result.length) return [];

  const columns = result[0].columns;
  return result[0].values
    // Transforma cada linha (array de valores) em objeto {coluna: valor}
    .map(values => Object.fromEntries(columns.map((name, index) => [name, values[index] ?? ''])))
    // Traduz as colunas "cruas" do SQLite para o mesmo formato "[P]..."
    // usado pelas linhas vindas do Google Sheets, para que o resto do
    // código (consolidate, render etc.) não precise saber a origem dos dados.
    .map(row => ({
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
      // 3G usa UARFCN, 4G/5G usam EARFCN — o valor bruto "earfcn" do
      // banco é redirecionado para a coluna correta conforme a tecnologia.
      '[P]DL_UARFCN': row.tech === '3G' ? row.earfcn : '',
      '[P]DL_EARFCN': ['4G', '5G'].includes(row.tech) ? row.earfcn : ''
    }));
}

/**
 * Ponto único de consulta usado pelo formulário de busca:
 * 1. Tenta a base local (queryLocal) — caminho rápido.
 * 2. Se não houver base local, verifica cache de consultas remotas
 *    anteriores (chave `query:SITE:UF:TECHS`).
 * 3. Se também não houver cache, consulta o Google Sheets (querySheet)
 *    para cada tecnologia em paralelo e salva o resultado em cache.
 */
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

  // Consulta as tecnologias selecionadas em paralelo e junta tudo num só array
  const result = (await Promise.all(techs.map(tech => querySheet(tech, site, uf)))).flat();

  try {
    const db = await openCache();
    db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(result, key);
  } catch { /* cache indisponível não impede a consulta */ }

  return result;
}


/* ============================================================================
   8. CONSOLIDAÇÃO DOS DADOS (linhas cruas -> estrutura pronta para exibir)
   ----------------------------------------------------------------------------
   Recebe o array de linhas encontradas (de queryLocal ou querySheet) e
   agrupa por tecnologia, removendo duplicidades e calculando os targets
   de bandwidth (4G/5G).
   ============================================================================ */

/** Retorna quais checkboxes de tecnologia estão marcados no momento. */
function selectedTechs() {
  return [...document.querySelectorAll('.techs input:checked')].map(input => input.value);
}

/**
 * Lê o valor de EARFCN/UARFCN de uma linha, considerando que o nome da
 * coluna muda conforme a tecnologia (3G usa UARFCN; 4G/5G usam EARFCN).
 */
function valueFor(row, field, tech) {
  if (field === 'earfcn') {
    return tech === '3G' ? col(row, '[P]DL_UARFCN')
         : tech === '4G' || tech === '5G' ? col(row, '[P]DL_EARFCN')
         : '';
  }
  return col(row, field);
}

/** Remove valores vazios/duplicados de um array, preservando a ordem. */
function unique(values) { return [...new Set(values.filter(Boolean))]; }

/**
 * Agrupa as linhas encontradas por tecnologia, extraindo:
 *  - dados de identidade do site (cidade, bairro, endereço, coordenadas)
 *  - bandas de operação por tecnologia (sem duplicar banda+earfcn)
 *  - azimutes por tecnologia (sem duplicar)
 *  - códigos de célula por tecnologia (BCCH no 2G, PSC no 3G, PCI no 4G/5G)
 *  - targets de bandwidth (4G/5G), evitando duplicar por banda/earfcn
 *
 * Bandas contendo "iot" são ignoradas (fora do escopo de bandwidth/target).
 */
function consolidate(found) {
  const result = {cidade:'', bairro:'', endereco:'', latitude:'', longitude:'', tech:{}, targets:[]};

  for (const tech of TECHS) {
    const techRows = found.filter(row => row.__tech === tech);
    result.tech[tech] = {bands: [], azimuth: [], codes: [], rows: techRows};

    techRows.forEach(row => {
      // Preenche os dados de identidade do site só uma vez (primeiro valor encontrado)
      result.cidade ||= col(row, '[P]CIDADE');
      result.bairro ||= col(row, '[P]BAIRRO');
      result.latitude ||= col(row, '[P]LATITUDE');
      result.longitude ||= col(row, '[P]LONGITUDE');
      result.endereco ||= col(row, '[P]ENDERECO');

      // --- Banda de operação (ignora "IoT" e não duplica banda+earfcn) ---
      const banda = valueFor(row, '[P]BANDA_OPERACAO', tech);
      const earfcn = valueFor(row, 'earfcn', tech);
      if (banda && !/iot/i.test(banda) && !result.tech[tech].bands.some(x => x.banda === banda && x.earfcn === earfcn)) {
        result.tech[tech].bands.push({banda, earfcn});
      }

      // --- Azimuth (sem duplicar valor) ---
      const az = valueFor(row, '[P]AZIMUTH', tech);
      if (az && !result.tech[tech].azimuth.includes(az)) result.tech[tech].azimuth.push(az);

      // --- Código de célula: BCCH (2G) / PSC (3G) / PCI (4G/5G) ---
      const code = tech === '2G' ? col(row, '[P]BCCH')
                 : tech === '3G' ? col(row, '[P]PSC')
                 : col(row, '[P]PCI');
      if (code && !result.tech[tech].codes.includes(code)) result.tech[tech].codes.push(code);

      // --- Target de bandwidth para 4G ---
      if (tech === '4G' && banda && !/iot/i.test(banda)) {
        const bw = col(row, '[P]BANDWIDTH');
        const mimo = col(row, '[P]MIMO');
        const target = target4g(banda, bw, mimo);
        if (target && !result.targets.some(x => x.tech === tech && x.label === `${banda} / ${earfcn}`)) {
          result.targets.push({
            tech,
            label: `${banda}${earfcn ? ` / ${earfcn}` : ''}`,
            bw: bw || '—',
            // Exibe "2x2" ou "4x4" conforme a mesma regra usada em target4g()
            mimo: mimo ? `${Number(mimo) >= 3 ? 4 : 2}x${Number(mimo) >= 3 ? 4 : 2}` : '—',
            dl: target[0],
            ul: target[1]
          });
        }
      }

      // --- Target de bandwidth para 5G ---
      if (tech === '5G' && banda) {
        const target = target5g(banda);
        if (target && !result.targets.some(x => x.tech === tech && x.label === `${banda} / ${earfcn}`)) {
          result.targets.push({
            tech,
            label: `${banda}${earfcn ? ` / ${earfcn}` : ''}`,
            bw: '—',
            mimo: '—',
            dl: target[0],
            ul: target[1]
          });
        }
      }
    });
  }

  return result;
}


/* ============================================================================
   9. RENDERIZAÇÃO (HTML DOS CARDS DE RESULTADO)
   ----------------------------------------------------------------------------
   Transforma o objeto consolidado (seção 8) em HTML, injetado em #results.
   ============================================================================ */

/** Monta um "item" de dado simples (rótulo + valor), com fallback "—". */
function item(label, value) {
  return `<div class="item"><small>${label}</small><strong>${value || '—'}</strong></div>`;
}

/**
 * Monta o item de coordenadas, incluindo o botão "VER NO GOOGLE MAPS"
 * (classe .maps-link, estilizada em styles.css seção 16) quando há
 * latitude/longitude disponíveis.
 */
function mapsItem(latitude, longitude) {
  if (!latitude || !longitude) return item('COORDENADAS', '—');
  const query = `${latitude},${longitude}`;
  return `<div class="item"><small>COORDENADAS</small><strong>${query}</strong><a class="maps-link" href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}" target="_blank" rel="noopener">VER NO GOOGLE MAPS</a></div>`;
}

/**
 * Renderiza o resultado completo da consulta dentro de #results:
 *  - card "SITE" (identidade + coordenadas/maps)
 *  - card "BANDA DE OPERAÇÃO / EARFCN" por tecnologia
 *  - tabelas de "CÓDIGOS DE CÉLULA" e "AZIMUTH" lado a lado
 *  - card "BANDWIDTH / TARGET VIVO" (só aparece se houver targets 4G/5G)
 * Também habilita o botão COPIAR e guarda `data` em lastData para ele usar.
 */
function render(data, site, uf) {
  lastData = data;
  $('copy').disabled = false;

  // --- Card de identidade do site ---
  let html = `<div class="card"><h3>SITE: ${site} - ${uf}</h3><div class="grid identity">${item('CIDADE',data.cidade)}${item('BAIRRO',data.bairro)}${item('ENDEREÇO',data.endereco)}${mapsItem(data.latitude,data.longitude)}</div></div>`;

  // --- Card de bandas de operação, uma coluna por tecnologia ativa ---
  html += '<div class="card bands"><h3>BANDA DE OPERAÇÃO / EARFCN</h3><div class="band-grid">';
  for (const tech of selectedTechs()) {
    const t = data.tech[tech];
    if (t.rows.length) {
      html += `<div class="band-column"><b>${tech}</b>${t.bands.map(x => `<span>${x.banda}${x.earfcn ? ` / <em>${x.earfcn}</em>` : ''}</span>`).join('')}</div>`;
    }
  }
  html += '</div></div>';

  // --- Tabelas de "Códigos de célula" e "Azimuth", lado a lado ---
  // `active` = tecnologias marcadas que realmente têm dados encontrados.
  const active = selectedTechs().filter(tech => data.tech[tech].rows.length);
  // Quantidade máxima de linhas de cada tabela, para alinhar as duas em conjunto.
  const maxCodes = Math.max(...active.map(tech => data.tech[tech].codes.length), 0);
  const maxAz = Math.max(...active.map(tech => data.tech[tech].azimuth.length), 0);

  html += '<div class="section-title">CÓDIGOS DE CÉLULA / AZIMUTH</div><div class="boxes">'
    + '<div class="card table-box"><h3>CÓDIGOS DE CÉLULA</h3><table><thead><tr>'
    + active.map(t => `<th>${t === '2G' ? 'BCCH' : t === '3G' ? 'PSCs' : 'PCIs'}<br>${t}</th>`).join('')
    + '</tr></thead><tbody>'
    + Array.from({length: maxCodes}, (_, i) => `<tr>${active.map(t => `<td>${data.tech[t].codes[i] || ''}</td>`).join('')}</tr>`).join('')
    + '</tbody></table></div>'
    + '<div class="card table-box"><h3>AZIMUTH</h3><table><thead><tr>'
    + active.map(t => `<th>${t}</th>`).join('')
    + '</tr></thead><tbody>'
    // Usa o maior entre maxAz e maxCodes para manter as duas tabelas com
    // a mesma quantidade de linhas visualmente (mesmo comportamento original)
    + Array.from({length: Math.max(maxAz, maxCodes)}, (_, i) => `<tr>${active.map(t => `<td>${data.tech[t].azimuth[i] || ''}</td>`).join('')}</tr>`).join('')
    + '</tbody></table></div></div>';

  // --- Card de bandwidth/target, só quando há algum target calculado ---
  if (data.targets.length) {
    html += '<div class="card target-card"><h3>BANDWIDTH / TARGET VIVO</h3><div class="target-scroll"><table><thead><tr><th>TECH</th><th>BANDA</th><th>BW</th><th>MIMO</th><th>DL TARGET</th><th>UL TARGET</th></tr></thead><tbody>'
      + data.targets.map(x => `<tr><td>${x.tech}</td><td>${x.label}</td><td>${x.bw}</td><td>${x.mimo}</td><td class="dl">${x.dl}</td><td class="ul">${x.ul}</td></tr>`).join('')
      + '</tbody></table></div></div>';
  }

  $('results').innerHTML = html;
  $('message').textContent = `${site} · ${uf} · consulta concluída`;
}


/* ============================================================================
   10. HISTÓRICO DE CONSULTAS (localStorage)
   ----------------------------------------------------------------------------
   Guarda as últimas 20 consultas (site+UF) no localStorage do navegador,
   exibidas na sidebar e clicáveis para repetir a consulta rapidamente.
   ============================================================================ */

/**
 * Salva uma consulta no histórico:
 * - remove entrada duplicada do mesmo site+UF (evita repetição)
 * - insere a nova consulta no topo da lista
 * - mantém só as 20 mais recentes
 */
function historySave(site, uf) {
  const key = 'base-vivo-history';
  const list = JSON.parse(localStorage.getItem(key) || '[]').filter(x => x.site !== site || x.uf !== uf);
  list.unshift({site, uf, ts: new Date().toLocaleString('pt-BR')});
  localStorage.setItem(key, JSON.stringify(list.slice(0, 20)));
  renderHistory();
}

/**
 * Redesenha a lista de histórico na sidebar (#history) e liga o clique
 * de cada item para preencher o formulário e disparar a consulta de novo.
 */
function renderHistory() {
  const list = JSON.parse(localStorage.getItem('base-vivo-history') || '[]');
  $('history').innerHTML = list.length
    ? list.map(x => `<button class="history-item" data-site="${x.site}" data-uf="${x.uf}">${x.site} · ${x.uf}<small>${x.ts}</small></button>`).join('')
    : '<span class="empty">Nenhuma consulta.</span>';

  // Reatribui o evento de clique toda vez que a lista é redesenhada
  // (os botões são recriados via innerHTML, então perdem listeners antigos)
  document.querySelectorAll('.history-item').forEach(b => b.onclick = () => {
    $('site').value = b.dataset.site;
    $('uf').value = b.dataset.uf;
    $('search-form').requestSubmit(); // dispara o mesmo fluxo do botão CONSULTAR
  });
}


/* ============================================================================
   11. EVENTOS DA INTERFACE
   ----------------------------------------------------------------------------
   Liga o formulário de busca e os botões da sidebar (Limpar cache / Copiar)
   à lógica das seções anteriores.
   ============================================================================ */

// --- Submit do formulário de consulta (botão CONSULTAR) ---
$('search-form').onsubmit = async e => {
  e.preventDefault();
  const site = $('site').value.trim().toUpperCase();
  const uf = $('uf').value.trim().toUpperCase();
  const techs = selectedTechs();
  if (!site || !uf || !techs.length) return; // campos obrigatórios não preenchidos

  $('message').className = 'message';
  $('message').textContent = 'Consultando as bases...';

  try {
    const found = await queryRemote(site, uf, techs);
    if (!found.length) {
      $('results').innerHTML = '';
      $('message').className = 'message error';
      $('message').textContent = 'Nenhum registro encontrado para os filtros informados.';
      return;
    }
    render(consolidate(found), site, uf);
    historySave(site, uf);
  } catch (error) {
    $('message').className = 'message error';
    $('message').textContent = `Não foi possível consultar as bases: ${error.message}`;
  }
};

// --- Botão "Limpar cache" (sidebar) ---
// Limpa todo o IndexedDB de cache (base local, resultados remotos etc.)
// e recarrega a página do zero para reconstruir tudo.
$('reload').onclick = async () => {
  try {
    const db = await openCache();
    db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).clear();
  } catch { /* se falhar, ainda assim recarrega a página abaixo */ }
  localDb = null;
  location.reload();
};

// --- Botão "COPIAR" (copia o texto do resultado atual para a área de transferência) ---
$('copy').onclick = async () => {
  if (!lastData) return;
  try {
    await navigator.clipboard.writeText($('results').innerText);
    $('copy').textContent = 'COPIADO';
    setTimeout(() => $('copy').textContent = 'COPIAR', 1500); // volta ao texto normal após 1.5s
  } catch {
    alert('Não foi possível copiar neste navegador.');
  }
};


/* ============================================================================
   12. INICIALIZAÇÃO
   ----------------------------------------------------------------------------
   Executado assim que o script é carregado (ver ordem de <script> no
   index.html: sql-wasm.js primeiro, depois este arquivo).
   ============================================================================ */
renderHistory();                                        // mostra histórico salvo (se houver)
rows = [];                                               // garante estado limpo do fluxo XLSX (seção 5)
$('db-status').textContent = 'Carregando base local...';
$('message').textContent = 'Digite o Site e o UF para iniciar a consulta.';
loadLocalDatabase();                                     // tenta abrir a base SQLite local (caminho principal)
