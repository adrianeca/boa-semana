// ============================================================
//  Boa Semana — Painel do Diretor
//  Code.gs
// ============================================================

const HUB_URL             = 'https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec';
const HUB_SHEET_ID        = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const MATRICULAS_SHEET_ID = '12QoeDXR86wISP-rk4pu6HybLsGnKhh3vvz7YBit37G0';
const CANCELADOS_SHEET_ID = '1PxXw7wowahFSpi380JxM2A8kCN_fwBrRWc8wWctdrUc';
const INADIMPLENCIA_SHEET_ID = '1ubH_01diSh8r2djLk0bNbS1adcPQYYEN_UJBF3EF5c4';

const NO_COBRAFIX = new Set(['BF','CH','DT','IP','IT','MR','NL','TQ','LJ','PC','PN']);
const MEU_ACESSO  = 'boa semana'; // identificador na col H (ACESSOS) da aba SESSOES
const ALL_UNITS   = ['BF','BG','CG','CH','CP','CX','DT','FG','IG','IP','IT','LJ','MR','NI','NL','NT','PC','PN','PO','RC','TJ','TQ','VP','VQ'];

const _norm = s => String(s||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');

// ── Entry point ───────────────────────────────────────────────
function doGet(e) {
  const p     = (e && e.parameter) ? e.parameter : {};
  const token = p.s || p.session || '';

  if (!token) return _redirectToHub();

  const auth = _authAndAccess(token);
  if (!auth) return _redirectToHub();

  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.sessionToken = token;
  tmpl.apelido      = auth.access.apelido;
  tmpl.unidades     = JSON.stringify(auth.access.unidades);
  // Para admin (unidades=null), passa a lista completa para o seletor
  tmpl.allUnits     = JSON.stringify(auth.access.unidades === null ? ALL_UNITS : auth.access.unidades);

  return tmpl.evaluate()
    .setTitle('Boa Semana – BRASAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Chamadas separadas do frontend (paralelas) ────────────────
function getKpiOnly(token, unitFilter) {
  try {
    const auth = _authAndAccess(token);
    if (!auth) return JSON.stringify({ ok: false, error: 'Sessão inválida ou sem acesso.' });
    const units = _resolveUnits(auth.access, unitFilter);
    if (units === false) return JSON.stringify({ ok: false, error: 'Acesso não autorizado.' });

    const cache    = CacheService.getScriptCache();
    const cacheKey = 'kpi_' + (units === null ? 'ALL' : units.slice().sort().join(','));
    const cached   = cache.get(cacheKey);
    if (cached) return cached;

    const result = JSON.stringify({ ok: true, kpi: _computeKpiData(units) });
    try { cache.put(cacheKey, result, 3600); } catch(e) {}
    return result;
  } catch(err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

function getInadOnly(token, unitFilter) {
  try {
    const auth = _authAndAccess(token);
    if (!auth) return JSON.stringify({ ok: false, error: 'Sessão inválida ou sem acesso.' });
    const units = _resolveUnits(auth.access, unitFilter);
    if (units === false) return JSON.stringify({ ok: false, error: 'Acesso não autorizado.' });

    const cache    = CacheService.getScriptCache();
    const cacheKey = 'inad_' + (units === null ? 'ALL' : units.slice().sort().join(','));
    const cached   = cache.get(cacheKey);
    if (cached) return cached;

    const result = JSON.stringify({ ok: true, inad: _computeInadData(units) });
    try { cache.put(cacheKey, result, 3600); } catch(e) {}
    return result;
  } catch(err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

// ── Auth: lê SESSOES uma vez, checa col H (ACESSOS) e col E (UNIDADE) ──
function _authAndAccess(token) {
  if (!token) return null;
  const cache   = CacheService.getScriptCache();
  const sessKey = 'sess_' + token;

  const cached = cache.get(sessKey);
  if (cached === 'invalid') return null;
  if (cached) {
    try {
      const r = JSON.parse(cached);
      if (r && new Date(r._expira) > new Date()) return r;
    } catch(e) {}
  }

  try {
    const ss    = SpreadsheetApp.openById(HUB_SHEET_ID);
    const now   = new Date();
    const sheet = ss.getSheetByName('SESSOES');
    if (!sheet) return null;

    // [TOKEN, EMAIL, NOME, ROLE, UNIDADE, CRIADO_EM, EXPIRA_EM, ACESSOS]
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(token)) continue;

      const expira = data[i][6] ? new Date(data[i][6]) : null;
      if (!expira || expira < now) { cache.put(sessKey, 'invalid', 300); return null; }

      const email = String(data[i][1]||'').trim().toLowerCase();
      if (!email) { cache.put(sessKey, 'invalid', 300); return null; }

      // Col D (índice 3) = ROLE — admins passam sem verificar col H
      const role = String(data[i][3]||'').trim().toLowerCase();
      const isAdmin = role === 'admin' || role === 'super_admin' || role === 'superadmin';

      // Col H (índice 7) = ACESSOS — verifica se tem acesso a este painel
      // Admins passam sempre; col H vazia também passa (hub ainda não configurado)
      if (!isAdmin) {
        const acessosStr = String(data[i][7]||'').toLowerCase().trim();
        if (acessosStr) {
          const acessos = acessosStr.split(',').map(a => a.trim()).filter(Boolean);
          if (!acessos.includes(MEU_ACESSO)) { cache.put(sessKey, 'invalid', 60); return null; }
        }
      }

      // Col E (índice 4) = UNIDADE, separado por pipe; vazio = acesso a todas
      const unidadeRaw  = String(data[i][4]||'').trim();
      const unidades    = unidadeRaw
        ? unidadeRaw.split('|').map(u => u.trim().toUpperCase()).filter(Boolean)
        : null; // null = acesso a todas as unidades

      const result = {
        access: {
          apelido:  String(data[i][2]||'').trim().split(' ')[0] || email.split('@')[0],
          unidades, // null = todas | string[] = apenas estas
        },
        _expira: expira.toISOString(),
      };

      const ttl = Math.min(1800, Math.max(60, Math.floor((expira - now) / 1000)));
      try { cache.put(sessKey, JSON.stringify(result), ttl); } catch(e) {}
      return result;
    }

    cache.put(sessKey, 'invalid', 300);
    return null;
  } catch(err) {
    Logger.log('_authAndAccess error: ' + err.message);
    return null;
  }
}

// false = sem permissão | null = todas as unidades | string[] = unidades específicas
function _resolveUnits(access, unitFilter) {
  if (!unitFilter) return access.unidades; // null = todas, array = específicas
  const f = String(unitFilter).trim().toUpperCase();
  if (access.unidades === null) return [f]; // acesso total → qualquer unidade ok
  return access.unidades.includes(f) ? [f] : false;
}

// ── KPI: Matrículas + Cancelados + Saldo ─────────────────────
function _computeKpiData(units) {
  const now      = new Date();
  const curYear  = now.getFullYear();
  const curMonth = now.getMonth() + 1;
  const curQ     = Math.ceil(curMonth / 3);
  const prevYear = curYear - 1;
  const qStart   = (curQ - 1) * 3 + 1;
  const qEnd     = qStart + 2;

  const matData = SpreadsheetApp.openById(MATRICULAS_SHEET_ID)
    .getSheetByName('db_all_since2020').getDataRange().getValues();
  const canData = SpreadsheetApp.openById(CANCELADOS_SHEET_ID)
    .getSheetByName('db_Cancelados_All').getDataRange().getValues();

  const matH = matData[0].map(_norm);
  const canH = canData[0].map(_norm);

  const mI = {
    unidade: matH.indexOf('unidade'),
    mes:     matH.indexOf('mes'),
    ano:     matH.indexOf('ano'),
    mat:     matH.findIndex(h => h === 'matriculas'),
  };
  const cI = {
    unidade: canH.findIndex(h => h === 'unidade'),
    mes:     canH.findIndex(h => h === 'mes'),
    ano:     canH.findIndex(h => h === 'ano'),
    can:     canH.findIndex(h => h === 'cancelados'),
  };

  // Pré-indexar apenas os anos que interessam (atual e anterior)
  const unitSet = units ? new Set(units) : null; // null = sem filtro (todas)
  const matIdx = {}, canIdx = {};
  for (let i = 1; i < matData.length; i++) {
    const r = matData[i];
    const u = String(r[mI.unidade]||'').trim().toUpperCase();
    if (unitSet && !unitSet.has(u)) continue;
    const y = +r[mI.ano], m = +r[mI.mes];
    if (y !== curYear && y !== prevYear) continue;
    const key = `${y}_${m}`;
    matIdx[key] = (matIdx[key] || 0) + (parseFloat(r[mI.mat]) || 0);
  }
  for (let i = 1; i < canData.length; i++) {
    const r = canData[i];
    const u = String(r[cI.unidade]||'').trim().toUpperCase();
    if (unitSet && !unitSet.has(u)) continue;
    const y = +r[cI.ano], m = +r[cI.mes];
    if (y !== curYear && y !== prevYear) continue;
    const key = `${y}_${m}`;
    canIdx[key] = (canIdx[key] || 0) + (parseFloat(r[cI.can]) || 0);
  }

  function sumMat(year, m1, m2) { let t = 0; for (let m = m1; m <= m2; m++) t += matIdx[`${year}_${m}`] || 0; return t; }
  function sumCan(year, m1, m2) { let t = 0; for (let m = m1; m <= m2; m++) t += canIdx[`${year}_${m}`] || 0; return t; }
  function pct(cur, prev) { return prev ? +((cur - prev) / prev * 100).toFixed(1) : null; }

  const mMes = sumMat(curYear, curMonth, curMonth);
  const mTri = sumMat(curYear, qStart, qEnd);
  const mAno = sumMat(curYear, 1, curMonth);
  const cMes = sumCan(curYear, curMonth, curMonth);
  const cTri = sumCan(curYear, qStart, qEnd);
  const cAno = sumCan(curYear, 1, curMonth);

  const mMesP = sumMat(prevYear, curMonth, curMonth);
  const mTriP = sumMat(prevYear, qStart, qEnd);
  const mAnoP = sumMat(prevYear, 1, curMonth);
  const cMesP = sumCan(prevYear, curMonth, curMonth);
  const cTriP = sumCan(prevYear, qStart, qEnd);
  const cAnoP = sumCan(prevYear, 1, curMonth);

  const sparkline = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(curYear, curMonth - 1 - i, 1);
    const y = d.getFullYear(), m = d.getMonth() + 1;
    sparkline.push({ y, m, mat: sumMat(y, m, m), can: sumCan(y, m, m) });
  }

  return {
    matriculas: {
      mes: { val: mMes, pct: pct(mMes, mMesP) },
      tri: { val: mTri, pct: pct(mTri, mTriP) },
      ano: { val: mAno, pct: pct(mAno, mAnoP) },
    },
    cancelados: {
      mes: { val: cMes, pct: pct(cMes, cMesP) },
      tri: { val: cTri, pct: pct(cTri, cTriP) },
      ano: { val: cAno, pct: pct(cAno, cAnoP) },
    },
    saldo: {
      mes: { val: mMes - cMes },
      tri: { val: mTri - cTri },
      ano: { val: mAno - cAno },
    },
    sparkline,
    periodo: { ano: curYear, mes: curMonth, trimestre: curQ },
  };
}

// ── Inadimplência ─────────────────────────────────────────────
function _computeInadData(units) {
  const unitSet    = units ? new Set(units) : null;
  // cobraUnits: unidades que usam cobrafix; null=todas exceto NO_COBRAFIX; []=nenhuma (só NO_COBRAFIX solicitadas)
  const cobraUnits = units ? units.filter(u => !NO_COBRAFIX.has(u)) : null;
  const ss = SpreadsheetApp.openById(INADIMPLENCIA_SHEET_ID);

  // 1. Totais gerais
  const totalData = ss.getSheetByName('inad_total+cobrafix').getDataRange().getValues();
  const totalH    = totalData[0].map(_norm);
  const tI = {
    unidade:   totalH.findIndex(h => h === 'unidade'),
    titulos:   totalH.findIndex(h => h === 'titulos'),
    valorPrev: totalH.findIndex(h => h.includes('valor') && h.includes('previsto')),
    dataRel:   totalH.findIndex(h => h.includes('data') && h.includes('relat')),
  };

  const latestDate = {};
  for (let i = 1; i < totalData.length; i++) {
    const unit = String(totalData[i][tI.unidade]||'').trim().toUpperCase();
    if (unitSet && !unitSet.has(unit)) continue;
    const d = totalData[i][tI.dataRel] instanceof Date
      ? totalData[i][tI.dataRel] : new Date(totalData[i][tI.dataRel]);
    if (isNaN(d.getTime())) continue;
    if (!latestDate[unit] || d > latestDate[unit]) latestDate[unit] = d;
  }

  let totalTitulos = 0, totalValor = 0;
  for (let i = 1; i < totalData.length; i++) {
    const unit = String(totalData[i][tI.unidade]||'').trim().toUpperCase();
    if ((unitSet && !unitSet.has(unit)) || !latestDate[unit]) continue;
    const d = totalData[i][tI.dataRel] instanceof Date
      ? totalData[i][tI.dataRel] : new Date(totalData[i][tI.dataRel]);
    if (isNaN(d.getTime()) || d.getTime() !== latestDate[unit].getTime()) continue;
    totalTitulos += parseFloat(totalData[i][tI.titulos])   || 0;
    totalValor   += parseFloat(totalData[i][tI.valorPrev]) || 0;
  }

  // 2. Cobrafix por ano
  const cobrafixByYear = {};
  const cobraSet = cobraUnits ? new Set(cobraUnits) : null;
  if (cobraUnits === null || cobraUnits.length > 0) {
    const baixData = ss.getSheetByName('db_baixados_max').getDataRange().getValues();
    const baixH    = baixData[0].map(_norm);
    const bI = {
      unidadeAj: baixH.findIndex(h => h.includes('unidade') && h.includes('ajust')),
      previsto:  baixH.findIndex(h => h === 'previsto'),
      dataVcto:  baixH.findIndex(h => h === 'data_vcto'),
    };
    for (let i = 1; i < baixData.length; i++) {
      const unit = String(baixData[i][bI.unidadeAj]||'').trim().toUpperCase();
      if (cobraSet ? !cobraSet.has(unit) : NO_COBRAFIX.has(unit)) continue;
      const dv = baixData[i][bI.dataVcto];
      const d  = dv instanceof Date ? dv : new Date(dv);
      const yr = isNaN(d.getTime()) ? null : d.getFullYear();
      if (!yr) continue;
      if (!cobrafixByYear[yr]) cobrafixByYear[yr] = { titulos: 0, valor: 0 };
      cobrafixByYear[yr].titulos += 1;
      cobrafixByYear[yr].valor   += parseFloat(baixData[i][bI.previsto]) || 0;
    }
  }

  // 3. Inadpf por ano
  const inadpfData = ss.getSheetByName('db_inadpf_max').getDataRange().getValues();
  const inadH      = inadpfData[0].map(_norm);
  const iI = {
    unidadeAj:  inadH.findIndex(h => h.includes('unidade') && h.includes('ajust')),
    qtReceber:  inadH.findIndex(h => h.includes('qt') && h.includes('receber')),
    valReceber: inadH.findIndex(h => h.includes('valor') && h.includes('receber')),
    valPrevPF:  inadH.findIndex(h => h.includes('valor') && h.includes('previsto') && h.includes('pf')),
    ano:        inadH.findIndex(h => h === 'ano'),
  };

  const inadpfByYear = {};
  for (let i = 1; i < inadpfData.length; i++) {
    const unit = String(inadpfData[i][iI.unidadeAj]||'').trim().toUpperCase();
    const ano  = parseInt(inadpfData[i][iI.ano]);
    if (!ano || (unitSet && !unitSet.has(unit))) continue;
    if (!NO_COBRAFIX.has(unit) && ano < 2025) continue;
    if (!inadpfByYear[ano]) inadpfByYear[ano] = { valReceber: 0, qtReceber: 0, valPrevPF: 0 };
    inadpfByYear[ano].valReceber += parseFloat(inadpfData[i][iI.valReceber]) || 0;
    inadpfByYear[ano].qtReceber  += parseFloat(inadpfData[i][iI.qtReceber])  || 0;
    inadpfByYear[ano].valPrevPF  += parseFloat(inadpfData[i][iI.valPrevPF])  || 0;
  }

  // 4. Consolidar antes de 2022
  const PRE_YEAR = 2022;
  let preCobra = null, preInad = null;
  const fCobra = {}, fInad = {};
  const allYears = new Set([
    ...Object.keys(cobrafixByYear).map(Number),
    ...Object.keys(inadpfByYear).map(Number),
  ]);
  for (const yr of [...allYears].sort()) {
    if (yr < PRE_YEAR) {
      const cb = cobrafixByYear[yr], ip = inadpfByYear[yr];
      if (cb) { if (!preCobra) preCobra = { titulos: 0, valor: 0 }; preCobra.titulos += cb.titulos; preCobra.valor += cb.valor; }
      if (ip) { if (!preInad)  preInad  = { valReceber: 0, qtReceber: 0, valPrevPF: 0 }; preInad.valReceber += ip.valReceber; preInad.qtReceber += ip.qtReceber; preInad.valPrevPF += ip.valPrevPF; }
    } else {
      if (cobrafixByYear[yr]) fCobra[yr] = cobrafixByYear[yr];
      if (inadpfByYear[yr])   fInad[yr]  = inadpfByYear[yr];
    }
  }

  // 5. Year cards
  const yearCards = {};
  if (preCobra || preInad) yearCards['pre2022'] = _buildYearCard(preCobra, preInad);
  const recentYears = new Set([...Object.keys(fCobra).map(Number), ...Object.keys(fInad).map(Number)]);
  for (const yr of [...recentYears].sort()) yearCards[yr] = _buildYearCard(fCobra[yr]||null, fInad[yr]||null);

  return {
    total: {
      valorPrevisto: totalValor,
      titulos:       Math.round(totalTitulos),
      ticketMedio:   totalTitulos > 0 ? totalValor / totalTitulos : 0,
    },
    yearCards,
  };
}

function _buildYearCard(cb, ipRaw) {
  return {
    cobrafix: cb ? { titulos: cb.titulos, valor: cb.valor } : null,
    inadpf: ipRaw ? {
      inadAcumulada:   ipRaw.valReceber,
      titulosAReceber: ipRaw.qtReceber,
      ticketMedio:     ipRaw.qtReceber > 0 ? ipRaw.valReceber / ipRaw.qtReceber : 0,
      pctInad:         ipRaw.valPrevPF  > 0 ? (ipRaw.valReceber / ipRaw.valPrevPF * 100) : 0,
    } : null,
  };
}

// ── Debug temporário ──────────────────────────────────────────
function debugKpi() {
  const matData = SpreadsheetApp.openById(MATRICULAS_SHEET_ID)
    .getSheetByName('db_all_since2020').getDataRange().getValues();
  const matH = matData[0].map(_norm);
  const mI = { unidade: matH.indexOf('unidade'), mes: matH.indexOf('mes'), ano: matH.indexOf('ano'), mat: matH.findIndex(h => h === 'matriculas') };
  const curYear = new Date().getFullYear();
  const samples = [];
  for (let i = 1; i < matData.length && samples.length < 5; i++) {
    if (+matData[i][mI.ano] === curYear) samples.push({ u: matData[i][mI.unidade], m: matData[i][mI.mes], a: matData[i][mI.ano], v: matData[i][mI.mat] });
  }
  const result = { indices: mI, curYear, amostras_ano_atual: samples, total_linhas: matData.length };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// ── Utilitários de cache ──────────────────────────────────────
function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['kpi_ALL','inad_ALL']);
  Logger.log('Cache limpo.');
}

// ── Pré-aquecimento de cache ──────────────────────────────────
// Gatilho recomendado: a cada 30 min (ScriptApp > Gatilhos)
function warmCache() {
  try {
    const cache    = CacheService.getScriptCache();
    const kpiData  = _computeKpiData(null);
    const inadData = _computeInadData(null);

    const kpiResult  = JSON.stringify({ ok: true, kpi:  kpiData  });
    const inadResult = JSON.stringify({ ok: true, inad: inadData });

    try { cache.put('kpi_ALL',  kpiResult,  3600); } catch(e) {}
    try { cache.put('inad_ALL', inadResult, 3600); } catch(e) {}

    Logger.log('warmCache OK — kpi=' + kpiResult.length + ' bytes, inad=' + inadResult.length + ' bytes');
  } catch(err) {
    Logger.log('warmCache ERRO: ' + err.message);
  }
}

// ── Redirect / Acesso negado ──────────────────────────────────
function _redirectToHub() {
  const appUrl = ScriptApp.getService().getUrl();
  const appUrlWithQuery = appUrl + (appUrl.includes('?') ? '&' : '?') + '_=1';
  const target = HUB_URL + '?next=' + encodeURIComponent(appUrlWithQuery);
  return HtmlService.createHtmlOutput(
    `<!DOCTYPE html><html><head>
      <meta http-equiv="refresh" content="0;url=${target}">
      <style>body{font-family:sans-serif;background:#0a1628;color:#89afd4;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}</style>
    </head><body>
      <div style="text-align:center">
        <p>Redirecionando para autenticação…</p>
        <a href="${target}" style="color:#fff">Clique aqui se não for redirecionado</a>
      </div>
      <script>try{window.top.location.replace('${target}');}catch(e){window.location.replace('${target}');}<\/script>
    </body></html>`
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _accessDenied(email) {
  return HtmlService.createHtmlOutput(
    `<!DOCTYPE html><html><body style="font-family:sans-serif;background:#0a1628;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0">
      <div style="text-align:center;max-width:400px">
        <div style="font-size:2rem;margin-bottom:16px">🔒</div>
        <h2 style="margin:0 0 8px">Acesso não autorizado</h2>
        <p style="color:#89afd4;margin:0">${email} não tem acesso a este painel.</p>
      </div>
    </body></html>`
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Diagnóstico ───────────────────────────────────────────────
function diagnostico() {
  const t0 = new Date();
  const results = { inicio: t0.toISOString() };
  try {
    const ss       = SpreadsheetApp.openById(HUB_SHEET_ID);
    const sessSheet = ss.getSheetByName('SESSOES');
    if (sessSheet) {
      const allData = sessSheet.getDataRange().getValues();
      const h = allData[0];
      results.sessoes          = 'OK';
      results.sessoes_colunas  = h.join(' | ');
      results.sessoes_col_acessos = h[7] ? 'OK (col H = ' + h[7] + ')' : 'AUSENTE';
      // Mostra as últimas 3 sessões ativas com role e acessos
      const now = new Date();
      const amostras = [];
      for (let i = allData.length - 1; i >= 1 && amostras.length < 3; i--) {
        const exp = allData[i][6] ? new Date(allData[i][6]) : null;
        if (exp && exp > now) amostras.push({
          email:   String(allData[i][1]).substring(0,30),
          role:    String(allData[i][3]||''),
          acessos: String(allData[i][7]||'(vazio)'),
        });
      }
      results.sessoes_amostras = amostras;
    } else {
      results.sessoes = 'ERRO: aba não encontrada';
    }
  } catch(e) { results.hub = 'ERRO: ' + e.message; }

  try {
    const ss = SpreadsheetApp.openById(MATRICULAS_SHEET_ID);
    const s  = ss.getSheetByName('db_all_since2020');
    if (s) {
      const h = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
      results.matriculas         = 'OK — ' + s.getLastRow() + ' linhas';
      results.matriculas_headers = h.join(' | ');
      results.matriculas_norm    = h.map(_norm).join(' | ');
    } else { results.matriculas = 'ERRO: aba não encontrada'; }
  } catch(e) { results.matriculas = 'ERRO: ' + e.message; }

  try {
    const ss = SpreadsheetApp.openById(CANCELADOS_SHEET_ID);
    const s  = ss.getSheetByName('db_Cancelados_All');
    if (s) {
      const h = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
      results.cancelados         = 'OK — ' + s.getLastRow() + ' linhas';
      results.cancelados_headers = h.join(' | ');
      results.cancelados_norm    = h.map(_norm).join(' | ');
    } else { results.cancelados = 'ERRO: aba não encontrada'; }
  } catch(e) { results.cancelados = 'ERRO: ' + e.message; }

  try {
    const ss = SpreadsheetApp.openById(INADIMPLENCIA_SHEET_ID);
    ['inad_total+cobrafix','db_baixados_max','db_inadpf_max'].forEach(n => {
      const s = ss.getSheetByName(n);
      results['inad_' + n] = s ? 'OK — ' + s.getLastRow() + ' linhas' : 'ERRO';
    });
  } catch(e) { results.inadimplencia = 'ERRO: ' + e.message; }

  results.tempo_total_ms = new Date() - t0;
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}
