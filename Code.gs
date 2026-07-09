// ============================================================
//  Boa Semana — Painel do Diretor
//  Code.gs
// ============================================================

const HUB_URL             = 'https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec';
const HUB_SHEET_ID        = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const PREENCHIMENTO_SHEET_ID = '13hjshnh3EBcOAT3EmRDB5JO0FB51knfmT9PdCpoa4TU';
const INADIMPLENCIA_SHEET_ID = '1ubH_01diSh8r2djLk0bNbS1adcPQYYEN_UJBF3EF5c4';
const CLASS_AVERAGE_SHEET_ID = '1O5mYkfiFKpSd0aFxWVWjXJFxlbw6jYnDZX0SbiDda5M';
const ESTATISTICA_SHEET_ID   = '1qiafd1roeusjkfXkOTPp2SdnXMVLdDHvYk4lDRwl4vE';
const NPS_SHEET_ID           = '1JImJD3_KxbOYZ0g7b7ibCau9dMw7g__LrC4cXlqhnnU';

// Bump isto a cada mudança na lógica de _computeKpiData/_computeInadData — invalida
// automaticamente todo cache antigo (de qualquer unidade), sem precisar rodar clearCache().
// v9: weekly da inadimplência passou a retornar { years: [{ano, points}] } em vez de y2025/y2026
// v10: respostas da BOL na tabela do NPS só aparecem para BOL e acesso total
const CACHE_VERSION = 'v10';

// ── Memoização de leituras de planilha (por execução) ─────────
// SpreadsheetApp é o gargalo: cada getValues() de uma aba grande custa segundos.
// Este memo garante que cada aba seja lida no máximo UMA vez por execução — essencial
// em sendWeeklyEmails(), que computa KPI/Inad/NPS para 27 unidades + "todas" e, sem isso,
// relia as mesmas abas ~28 vezes (risco real de estourar o limite de 6 min do Apps Script).
// O objeto global zera a cada execução, então não há risco de dados velhos entre execuções.
const _sheetDataMemo = {};
function _getSheetData(ssId, sheetName) {
  const key = ssId + '|' + sheetName;
  if (!_sheetDataMemo[key]) {
    _sheetDataMemo[key] = SpreadsheetApp.openById(ssId).getSheetByName(sheetName).getDataRange().getValues();
  }
  return _sheetDataMemo[key];
}

const NO_COBRAFIX = new Set(['BF','CH','DT','IP','IT','MR','NL','TQ','LJ','PC','PN']);
const MEU_ACESSO  = 'boa semana'; // identificador na col H (ACESSOS) da aba SESSOES
const ALL_UNITS   = ['BF','BG','BOL','CG','CH','CP','CX','DT','FG','GR','IG','IP','IT','LJ','MR','NI','NL','NT','PC','PN','PO','RC','TJ','TQ','VO','VP','VQ'];

const _norm = s => String(s||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');

// Algumas bases gravam "ONLINE" ou "Métodos Online" em vez da sigla "BOL" para a mesma unidade
const UNIT_CODE_ALIASES = { 'online': 'BOL', 'metodos online': 'BOL', 'bol': 'BOL' };
// Usado onde a coluna já vem com a sigla (não o nome completo) — unifica as variações acima
const _unitCode = raw => UNIT_CODE_ALIASES[_norm(raw)] || String(raw||'').trim().toUpperCase();

// ── Equivalência nome completo da unidade (col Unidade da aba Preenchimento) → sigla ──
const UNIDADE_SIGLA = {
  'bol':                       'BOL',
  'online':                    'BOL',
  'metodos online':            'BOL',
  'botafogo':                  'BF',
  'cachambi':                  'CH',
  'campo grande':              'CG',
  'caxias':                    'CX',
  'copacabana':                'CP',
  'downtown':                  'DT',
  'freguesia':                 'FG',
  'fg assessoria':             'FG',
  'grajau':                    'GR',
  'ilha do governador':        'IG',
  'ilha':                      'IG',
  'ipanema':                   'IP',
  'itaipu':                    'IT',
  'meier':                     'MR',
  'niteroi':                   'NT',
  'nova iguacu':               'NI',
  'novo leblon':               'NL',
  'parque olimpico':           'PO',
  'pechincha':                 'PC',
  'peninsula':                 'PN',
  'polo brasas - bangu':       'BG',
  'bangu':                     'BG',
  'polo brasas - laranjeiras': 'LJ',
  'laranjeiras':               'LJ',
  'recreio':                   'RC',
  'taquara':                   'TQ',
  'tq assessoria':             'TQ',
  'tijuca':                    'TJ',
  'vila da penha':             'VP',
  'vila olimpia':              'VO',
  'vila valqueire':            'VQ',
};
const _sigla = nome => UNIDADE_SIGLA[_norm(nome)] || null;

// Motivos de cancelamento que NÃO contam como cancelado
const CANCEL_EXCLUDE_MOTIVOS = new Set(['termino book 10', 'acerto de sistema']);

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
  tmpl.hubUrl       = HUB_URL;

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
    const cacheKey = 'kpi_' + CACHE_VERSION + '_' + (units === null ? 'ALL' : units.slice().sort().join(','));
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
    const cacheKey = 'inad_' + CACHE_VERSION + '_' + (units === null ? 'ALL' : units.slice().sort().join(','));
    const cached   = cache.get(cacheKey);
    if (cached) return cached;

    const result = JSON.stringify({ ok: true, inad: _computeInadData(units) });
    try { cache.put(cacheKey, result, 3600); } catch(e) {}
    return result;
  } catch(err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

function getClassAverageOnly(token, unitFilter) {
  try {
    const auth = _authAndAccess(token);
    if (!auth) return JSON.stringify({ ok: false, error: 'Sessão inválida ou sem acesso.' });
    const units = _resolveUnits(auth.access, unitFilter);
    if (units === false) return JSON.stringify({ ok: false, error: 'Acesso não autorizado.' });

    const cache    = CacheService.getScriptCache();
    const cacheKey = 'classavg_' + CACHE_VERSION + '_' + (units === null ? 'ALL' : units.slice().sort().join(','));
    const cached   = cache.get(cacheKey);
    if (cached) return cached;

    const weekly = _computeClassAverageWeekly(units);
    const cards  = _computeEstatisticaCards(units);
    cards.atual  = weekly.points.length ? weekly.points[weekly.points.length - 1].valor : null;

    const result = JSON.stringify({ ok: true, classAverage: { points: weekly.points, cards } });
    try { cache.put(cacheKey, result, 3600); } catch(e) {}
    return result;
  } catch(err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

function getNpsOnly(token, unitFilter) {
  try {
    const auth = _authAndAccess(token);
    if (!auth) return JSON.stringify({ ok: false, error: 'Sessão inválida ou sem acesso.' });
    const units = _resolveUnits(auth.access, unitFilter);
    if (units === false) return JSON.stringify({ ok: false, error: 'Acesso não autorizado.' });

    const cache    = CacheService.getScriptCache();
    const cacheKey = 'nps_' + CACHE_VERSION + '_' + (units === null ? 'ALL' : units.slice().sort().join(','));
    const cached   = cache.get(cacheKey);
    if (cached) return cached;

    const result = JSON.stringify({ ok: true, nps: _computeNpsData(units) });
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
        ? unidadeRaw.split('|').map(u => _unitCode(u)).filter(Boolean)
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
  const f = _unitCode(unitFilter);
  if (access.unidades === null) return [f]; // acesso total → qualquer unidade ok
  return access.unidades.includes(f) ? [f] : false;
}

// ── KPI: Matrículas + Cancelados + Saldo ─────────────────────
// Comparação sempre "mesmo intervalo de dias" no ano anterior (ex.: dia 1 a 3
// de julho/26 vs dia 1 a 3 de julho/25), não o mês/trimestre/ano fechado.
function _computeKpiData(units) {
  const now      = new Date();
  const curYear  = now.getFullYear();
  const curMonth = now.getMonth() + 1;
  const curQ     = Math.ceil(curMonth / 3);
  const prevYear = curYear - 1;
  const qStart   = (curQ - 1) * 3 + 1;
  const today    = new Date(curYear, now.getMonth(), now.getDate());

  const { matDay, canDay } = _loadMatCanIndexes(units, curYear, prevYear);

  function sumRange(idx, start, end) {
    let t = 0;
    const d = new Date(start);
    while (d <= end) {
      t += idx[`${d.getFullYear()}_${d.getMonth() + 1}_${d.getDate()}`] || 0;
      d.setDate(d.getDate() + 1);
    }
    return t;
  }
  function shiftYear(d, delta) { return new Date(d.getFullYear() + delta, d.getMonth(), d.getDate()); }
  function pct(cur, prev) { return prev ? +((cur - prev) / prev * 100).toFixed(1) : null; }

  const mesStart = new Date(curYear, curMonth - 1, 1);
  const triStart = new Date(curYear, qStart - 1, 1);
  const anoStart = new Date(curYear, 0, 1);
  const todayP   = shiftYear(today, -1);

  const mMes = sumRange(matDay, mesStart, today);
  const mTri = sumRange(matDay, triStart, today);
  const mAno = sumRange(matDay, anoStart, today);
  const cMes = sumRange(canDay, mesStart, today);
  const cTri = sumRange(canDay, triStart, today);
  const cAno = sumRange(canDay, anoStart, today);

  const mMesP = sumRange(matDay, shiftYear(mesStart, -1), todayP);
  const mTriP = sumRange(matDay, shiftYear(triStart, -1), todayP);
  const mAnoP = sumRange(matDay, shiftYear(anoStart, -1), todayP);
  const cMesP = sumRange(canDay, shiftYear(mesStart, -1), todayP);
  const cTriP = sumRange(canDay, shiftYear(triStart, -1), todayP);
  const cAnoP = sumRange(canDay, shiftYear(anoStart, -1), todayP);

  const sparkline = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(curYear, curMonth - 1 - i, 1);
    const y = d.getFullYear(), m = d.getMonth() + 1;
    const monthStart = new Date(y, m - 1, 1);
    const monthEnd   = new Date(y, m, 0);
    sparkline.push({ y, m, mat: sumRange(matDay, monthStart, monthEnd), can: sumRange(canDay, monthStart, monthEnd) });
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

// ── Matrículas + Cancelados: lê direto da aba "Preenchimento" (base bruta) ──
// mat: coluna "Tipo de Registro" (E) = Matrícula | Matrícula Turma Fechada; data = "Data da matrícula" (H)
// can: coluna "Tipo de Registro Saída" (T) = Cancelado, exceto motivo (F) = Término Book 10 | Acerto de sistema; data = "Data de Cancelamento" (U)
function _loadMatCanIndexes(units, curYear, prevYear) {
  const data = _getSheetData(PREENCHIMENTO_SHEET_ID, 'Preenchimento');

  const header = data[0].map(_norm);
  const cI = {
    unidade:      header.indexOf('unidade'),
    tipoReg:      header.indexOf('tipo de registro'),
    motivoCancel: header.indexOf('motivo do cancelamento'),
    dataMat:      header.indexOf('data da matricula'),
    tipoRegSaida: header.indexOf('tipo de registro saida'),
    dataCancel:   header.indexOf('data de cancelamento'),
  };

  const unitSet = units ? new Set(units) : null; // null = sem filtro (todas)
  // Indexado por dia exato (não por mês) para permitir comparação "mesmo intervalo de dias" ano a ano
  const matDay = {}, canDay = {};

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const sigla = _sigla(r[cI.unidade]);
    if (!sigla || (unitSet && !unitSet.has(sigla))) continue;

    const tipoReg = _norm(r[cI.tipoReg]);
    if (tipoReg === 'matricula' || tipoReg === 'matricula turma fechada') {
      const d = r[cI.dataMat] instanceof Date ? r[cI.dataMat] : new Date(r[cI.dataMat]);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        if (y === curYear || y === prevYear) {
          const key = `${y}_${d.getMonth() + 1}_${d.getDate()}`;
          matDay[key] = (matDay[key] || 0) + 1;
        }
      }
    }

    const tipoSaida = _norm(r[cI.tipoRegSaida]);
    if (tipoSaida === 'cancelado' && !CANCEL_EXCLUDE_MOTIVOS.has(_norm(r[cI.motivoCancel]))) {
      const d = r[cI.dataCancel] instanceof Date ? r[cI.dataCancel] : new Date(r[cI.dataCancel]);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        if (y === curYear || y === prevYear) {
          const key = `${y}_${d.getMonth() + 1}_${d.getDate()}`;
          canDay[key] = (canDay[key] || 0) + 1;
        }
      }
    }
  }

  return { matDay, canDay };
}

// ── Inadimplência ─────────────────────────────────────────────
function _computeInadData(units) {
  const unitSet    = units ? new Set(units) : null;
  // cobraUnits: unidades que usam cobrafix; null=todas exceto NO_COBRAFIX; []=nenhuma (só NO_COBRAFIX solicitadas)
  const cobraUnits = units ? units.filter(u => !NO_COBRAFIX.has(u)) : null;

  // 1. Totais gerais
  const totalData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'inad_total+cobrafix');
  const totalH    = totalData[0].map(_norm);
  // Header real da aba (confirmado via debugInadTotal): Data Relatório, Unidade, Mês, Ano,
  // Títulos, Valor Previsto, Mês/Ano de vencimento — colunas simples, sem "Ajustada"
  const tI = {
    unidade:   totalH.findIndex(h => h === 'unidade'),
    titulos:   totalH.findIndex(h => h === 'titulos'),
    valorPrev: totalH.findIndex(h => h.includes('valor') && h.includes('previsto')),
    dataRel:   totalH.findIndex(h => h.includes('data') && h.includes('relat')),
  };
  // "Unidade" pode vir como nome completo ("Botafogo") ou já como sigla ("BF"/"BOL") — tenta os dois
  const _totalUnit = raw => _sigla(raw) || _unitCode(raw);

  const latestDate = {};
  for (let i = 1; i < totalData.length; i++) {
    const unit = _totalUnit(totalData[i][tI.unidade]);
    if (unitSet && !unitSet.has(unit)) continue;
    const d = totalData[i][tI.dataRel] instanceof Date
      ? totalData[i][tI.dataRel] : new Date(totalData[i][tI.dataRel]);
    if (isNaN(d.getTime())) continue;
    if (!latestDate[unit] || d > latestDate[unit]) latestDate[unit] = d;
  }

  let totalTitulos = 0, totalValor = 0;
  for (let i = 1; i < totalData.length; i++) {
    const unit = _totalUnit(totalData[i][tI.unidade]);
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
    const baixData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'db_baixados_max');
    const baixH    = baixData[0].map(_norm);
    const bI = {
      unidadeAj: baixH.findIndex(h => h.includes('unidade') && h.includes('ajust')),
      previsto:  baixH.findIndex(h => h === 'previsto'),
      dataVcto:  baixH.findIndex(h => h === 'data_vcto'),
    };
    for (let i = 1; i < baixData.length; i++) {
      const unit = _unitCode(baixData[i][bI.unidadeAj]);
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
  const inadpfData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'db_inadpf_max');
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
    const unit = _unitCode(inadpfData[i][iI.unidadeAj]);
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

  // 6. Séries semanais (ano anterior / ano atual) — últimas 10 datas de relatório
  const weekly = _computeInadWeekly(unitSet);

  return {
    total: {
      valorPrevisto: totalValor,
      titulos:       Math.round(totalTitulos),
      ticketMedio:   totalTitulos > 0 ? totalValor / totalTitulos : 0,
    },
    yearCards,
    weekly,
  };
}

// ── Inadimplência semanal: soma db_inadimplência_pf + db_baixados por data de relatório ──
// pf: col F = ano, col I = valor, col U = data do relatório, col V = unidade
// baixados: ano tirado da col E (data da inadimplência), col H = valor, col J = data do relatório, col K = unidade
// Mostra só as 10 datas de relatório mais recentes (mesmo eixo X pras duas séries de ano).
// Os anos são derivados da data atual (ano anterior + ano atual) — não precisa atualizar na virada do ano.
function _computeInadWeekly(unitSet) {
  const curYear = new Date().getFullYear();
  const YEARS = [curYear - 1, curYear];
  const sums = {}; YEARS.forEach(y => sums[y] = {});
  const allDates = new Set();

  function addRow(ano, dataRelRaw, valor) {
    if (YEARS.indexOf(ano) === -1) return;
    const key = _dateKey(dataRelRaw);
    if (!key) return;
    sums[ano][key] = (sums[ano][key] || 0) + valor;
    allDates.add(key);
  }

  const pfData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'db_inadimplência_pf');
  for (let i = 1; i < pfData.length; i++) {
    const r = pfData[i];
    const unit = _unitCode(r[21]); // V
    if (unitSet && !unitSet.has(unit)) continue;
    addRow(parseInt(r[5]), r[20], parseFloat(r[8]) || 0); // F, U, I
  }

  const baixData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'db_baixados');
  for (let i = 1; i < baixData.length; i++) {
    const r = baixData[i];
    const unit = _unitCode(r[10]); // K
    if (unitSet && !unitSet.has(unit)) continue;
    const dIndiv = r[4] instanceof Date ? r[4] : new Date(r[4]); // E
    if (isNaN(dIndiv.getTime())) continue;
    addRow(dIndiv.getFullYear(), r[9], parseFloat(r[7]) || 0); // J, H
  }

  const last10 = [...allDates].sort().slice(-10);
  const series = ano => last10.map(key => ({ date: key, label: _dateLabel(key), valor: sums[ano][key] || 0 }));

  return { labels: last10.map(_dateLabel), years: YEARS.map(ano => ({ ano, points: series(ano) })) };
}

function _dateKey(raw) {
  const d = raw instanceof Date ? raw : new Date(raw);
  if (isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ── Class Average semanal (últimas 10 datas de relatório) ──
// Aba NEW_Class_Average: Unidade, Turmas, Nº Alunos, Class Average, Data
// Ao combinar unidades (ex.: "Todas as unidades"), usa média ponderada (soma Alunos / soma Turmas)
// por data de relatório, em vez de somar os valores de Class Average — senão distorce a métrica.
function _computeClassAverageWeekly(units) {
  const unitSet = units ? new Set(units) : null;
  const data = _getSheetData(CLASS_AVERAGE_SHEET_ID, 'NEW_Class_Average');

  const header = data[0].map(_norm);
  const cI = {
    unidade: header.indexOf('unidade'),
    turmas:  header.indexOf('turmas'),
    alunos:  header.findIndex(h => h.includes('aluno')),
    data:    header.indexOf('data'),
  };

  const sums = {}; // dateKey -> { turmas, alunos }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const sigla = _sigla(r[cI.unidade]);
    if (!sigla || (unitSet && !unitSet.has(sigla))) continue;

    const key = _dateKey(r[cI.data]);
    if (!key) continue;

    if (!sums[key]) sums[key] = { turmas: 0, alunos: 0 };
    sums[key].turmas += parseFloat(r[cI.turmas]) || 0;
    sums[key].alunos += parseFloat(r[cI.alunos]) || 0;
  }

  const last10 = Object.keys(sums).sort().slice(-10);
  const points = last10.map(key => {
    const s = sums[key];
    return { date: key, label: _dateLabel(key), valor: s.turmas > 0 ? s.alunos / s.turmas : 0 };
  });

  return { points };
}

// ── Class Average Regular / Parceiro (aba "Estatística" — 1 linha por unidade por mês) ──
// Usa só a linha mais recente (maior Carimbo de data/hora) de CADA unidade, igual ao
// padrão já usado em _computeInadData pra "Inadimplência Total" (latestDate por unidade).
// Regular  = soma(Nº de Alunos Regulares) / soma(Nº de Turmas Regulares)
// Parceiro = soma(alunos em turmas fechadas/escolas-empresas) / soma(turmas fechadas/escolas-empresas)
function _computeEstatisticaCards(units) {
  const unitSet = units ? new Set(units) : null;
  const data = _getSheetData(ESTATISTICA_SHEET_ID, 'Estatística');

  const latest = {}; // sigla -> { row, ts }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const sigla = _unitCode(r[3]); // Unidade — já vem como sigla nessa aba
    if (!sigla || (unitSet && !unitSet.has(sigla))) continue;
    const ts = r[0] instanceof Date ? r[0] : new Date(r[0]); // Carimbo de data/hora
    if (isNaN(ts.getTime())) continue;
    if (!latest[sigla] || ts > latest[sigla].ts) latest[sigla] = { row: r, ts };
  }

  let alunosRegulares = 0, turmasRegulares = 0;
  let alunosParceiro = 0, turmasParceiro = 0;

  Object.values(latest).forEach(({ row: r }) => {
    alunosRegulares += parseFloat(r[39]) || 0; // Nº de Alunos Regulares
    turmasRegulares += parseFloat(r[18]) || 0; // Nº de Turmas Regulares

    alunosParceiro += (parseFloat(r[9])  || 0)  // Nº de Alunos em Turmas Fechadas na Unidade
                     + (parseFloat(r[10]) || 0)  // Nº de Alunos em Turmas de Escolas/Empresas
                     + (parseFloat(r[11]) || 0)  // ...turmas fechadas custeado pelo aluno
                     + (parseFloat(r[12]) || 0); // ...turmas abertas custeado pelo aluno

    turmasParceiro += (parseFloat(r[19]) || 0)  // Nº de Turmas Fechadas na Unidade
                     + (parseFloat(r[20]) || 0)  // N° de Turmas de Escolas/Empresas custeada pela empresa
                     + (parseFloat(r[21]) || 0)  // ...turmas fechadas custeado pelo aluno
                     + (parseFloat(r[22]) || 0); // ...turmas abertas custeado pelo aluno
  });

  return {
    regular:  turmasRegulares > 0 ? alunosRegulares / turmasRegulares : 0,
    parceiro: turmasParceiro  > 0 ? alunosParceiro  / turmasParceiro  : 0,
  };
}

// ── NPS ─────────────────────────────────────────────────────
// Aba "C - A e K" (pesquisa de cancelamento, Adults + Kids — todas as unidades exceto BOL):
//   0 Carimbo (data/hora real de envio) · 2 Nome · 5 Livro em que parou (Book) · 7 Motivo do Cancelamento
//   9 Nota (Qual é a chance de indicar) · 10 Críticas/Sugestões/Elogios
//   12 Data Resposta (dia 1º do mês do lote — NÃO é a data de envio, só serve p/ agregação mensal)
//   13 Pesquisa (Adults/Kids) · 14 Unidade Ajustada (sigla)
// Aba "C - BOL" (pesquisa de cancelamento exclusiva da unidade BOL/Online — sem separação Adults/Kids,
// vira o medidor "Cancelados"; sem coluna de Carimbo, por isso usa a própria "Data Resposta"):
//   1 Nome · 4 Motivo do Cancelamento · 6 Nota · 7 Críticas/Sugestões/Elogios
//   8 Em que livro parou (Book) · 10 Data Resposta (data real da resposta nessa aba)
// Aba "N - All" (pesquisa de mudança de nível — comum a todas as unidades, inclusive BOL):
//   9 Nota · 12 Data Resposta · 14 UNIDADE AJUSTADA (sigla)
// NPS = (%promotores [nota 9-10] − %detratores [nota 0-6]) × 100
function _computeNpsData(units) {
  const unitSet = units ? new Set(units) : null;

  const cakData  = _getSheetData(NPS_SHEET_ID, 'C - A e K');
  const nAllData = _getSheetData(NPS_SHEET_ID, 'N - All');
  const bolData  = _getSheetData(NPS_SHEET_ID, 'C - BOL');

  const cakRows = cakData.slice(1).filter(r => {
    const unit = _unitCode(r[14]);
    return unit && (!unitSet || unitSet.has(unit));
  });
  const adultsRows = cakRows.filter(r => _norm(r[13]).includes('adult'));
  const kidsRows   = cakRows.filter(r => _norm(r[13]).includes('kids'));

  const nAllRows = nAllData.slice(1).filter(r => {
    const unit = _unitCode(r[14]);
    return unit && (!unitSet || unitSet.has(unit));
  });

  // C - BOL: toda a aba é da unidade BOL (não tem coluna de unidade) — só entra se BOL estiver no escopo
  const bolApplicable  = !unitSet || unitSet.has('BOL');
  const bolOnlyFilter  = !!(unitSet && unitSet.size === 1 && unitSet.has('BOL'));
  const bolRows = bolApplicable ? bolData.slice(1) : [];

  // Filtrando só por BOL, Adults/Kids (que são da C - A e K, não tem BOL) ficam zerados
  const adults     = bolOnlyFilter ? { nps: null, total: 0, promoters: 0, detractors: 0 } : _npsScore(adultsRows, 9);
  const kids       = bolOnlyFilter ? { nps: null, total: 0, promoters: 0, detractors: 0 } : _npsScore(kidsRows, 9);
  const nivel      = _npsScore(nAllRows, 9);
  const cancelados = _npsScore(bolRows, 6);

  // Data da última resposta — respeita o filtro de unidade, combina as pesquisas
  // C - A e K / N - All usam o Carimbo (data real de envio); C - BOL não tem Carimbo, usa a Data Resposta
  let lastResponse = null;
  [...cakRows, ...nAllRows].forEach(r => {
    const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (!isNaN(d.getTime()) && (!lastResponse || d > lastResponse)) lastResponse = d;
  });
  bolRows.forEach(r => {
    const d = r[10] instanceof Date ? r[10] : new Date(r[10]);
    if (!isNaN(d.getTime()) && (!lastResponse || d > lastResponse)) lastResponse = d;
  });

  // Tabela de respostas recentes — últimos 15 dias. Todos veem as respostas de todas as
  // unidades, EXCETO as da BOL: comentários da BOL só aparecem para a própria BOL e para
  // quem tem acesso a todas as unidades (bolApplicable).
  // C - A e K usa o Carimbo (col 0); C - BOL usa a Data Resposta (col 10, única data confiável na aba)
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 15);
  const recentRaw = [
    ...cakData.slice(1).map(r => ({ r, d: r[0]  instanceof Date ? r[0]  : new Date(r[0]),  src: 'cak' })),
    ...(bolApplicable ? bolData.slice(1).map(r => ({ r, d: r[10] instanceof Date ? r[10] : new Date(r[10]), src: 'bol' })) : []),
  ];
  const recent = recentRaw
    .filter(({ d }) => !isNaN(d.getTime()) && d >= cutoff)
    .sort((a, b) => b.d - a.d)
    .map(({ r, d, src }) => src === 'bol' ? ({
      unidade:  'BOL',
      data:     _dateKey(d),
      nome:     String(r[1]||''),
      book:     String(r[8]||''),
      nota:     r[6] === '' || r[6] === null || r[6] === undefined ? null : parseFloat(r[6]),
      motivo:   String(r[4]||''),
      criticas: String(r[7]||''),
    }) : ({
      unidade:  _unitCode(r[14]),
      data:     _dateKey(d),
      nome:     String(r[2]||''),
      book:     String(r[5]||''),
      nota:     r[9] === '' || r[9] === null || r[9] === undefined ? null : parseFloat(r[9]),
      motivo:   String(r[7]||''),
      criticas: String(r[10]||''),
    }))
    // Garantia extra: se alguma linha da C - A e K vier marcada como BOL/Online,
    // ela também fica restrita a quem pode ver a BOL
    .filter(row => bolApplicable || row.unidade !== 'BOL');

  return {
    adults, kids, mudancaNivel: nivel, cancelados,
    // Controla quais medidores o frontend exibe: unidade BOL só mostra Cancelados (não Adults/Kids),
    // as demais unidades só mostram Adults/Kids (não Cancelados); "todas as unidades" mostra os 4
    showAdultsKids:  !bolOnlyFilter,
    showCancelados:  bolApplicable,
    lastResponse: lastResponse ? _dateKey(lastResponse) : null,
    recent,
  };
}

function _npsScore(rows, notaIdx) {
  let promoters = 0, detractors = 0, total = 0;
  rows.forEach(r => {
    const nota = parseFloat(r[notaIdx]);
    if (isNaN(nota)) return;
    total++;
    if (nota >= 9) promoters++;
    else if (nota <= 6) detractors++;
  });
  return {
    nps: total > 0 ? +(((promoters - detractors) / total) * 100).toFixed(1) : null,
    total, promoters, detractors,
  };
}

// Diagnóstico do NPS — roda e olha o Log de Execução
function debugNps() {
  const result = {};
  try {
    const ss = SpreadsheetApp.openById(NPS_SHEET_ID);
    ['C - A e K', 'N - All'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      const key = name.replace(/[^a-zA-Z]/g, '_');
      if (!sheet) { result[key + '_aba'] = 'ERRO: aba não encontrada'; return; }
      const data = sheet.getDataRange().getValues();
      result[key + '_aba']      = 'OK';
      result[key + '_linhas']   = data.length;
      result[key + '_header']   = data[0];
      result[key + '_amostras'] = data.slice(1, 6).map(r => ({
        unidade: r[14],
        pesquisa: name === 'C - A e K' ? r[13] : '(n/a)',
        nota: r[9],
        dataResposta: r[12], dataRespostaType: typeof r[12],
      }));
    });
  } catch(e) { result.erro = e.message; }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function _dateLabel(key) {
  const p = key.split('-');
  return `${p[2]}/${p[1]}`;
}

// Diagnóstico do gráfico de Class Average — roda e olha o Log de Execução
function debugClassAverage() {
  const result = {};
  try {
    const ss    = SpreadsheetApp.openById(CLASS_AVERAGE_SHEET_ID);
    const sheet = ss.getSheetByName('NEW_Class_Average');
    result.aba = sheet ? 'OK' : 'ERRO: aba não encontrada';
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      result.linhas = data.length;
      result.header = data[0];

      const siglaMissing = {};
      let datasValidas = 0, datasInvalidas = 0;
      const samples = [];
      const header = data[0].map(_norm);
      const cI = {
        unidade: header.indexOf('unidade'),
        turmas:  header.indexOf('turmas'),
        alunos:  header.findIndex(h => h.includes('aluno')),
        data:    header.indexOf('data'),
      };
      result.indices = cI;

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const sigla = _sigla(r[cI.unidade]);
        if (!sigla) siglaMissing[String(r[cI.unidade])] = (siglaMissing[String(r[cI.unidade])] || 0) + 1;
        const key = _dateKey(r[cI.data]);
        key ? datasValidas++ : datasInvalidas++;
        if (samples.length < 5) samples.push({
          unidadeRaw: r[cI.unidade], sigla,
          turmas: r[cI.turmas], alunos: r[cI.alunos],
          dataRaw: r[cI.data], dataType: typeof r[cI.data],
        });
      }
      result.amostras          = samples;
      result.siglaMissingTop   = Object.entries(siglaMissing).sort((a,b)=>b[1]-a[1]).slice(0,10);
      result.datasValidas      = datasValidas;
      result.datasInvalidas    = datasInvalidas;
    }
  } catch(e) { result.erro = e.message; }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Diagnóstico dos cards Class Average Regular/Parceiro — roda e olha o Log de Execução
function debugEstatistica() {
  const result = {};
  try {
    const ss    = SpreadsheetApp.openById(ESTATISTICA_SHEET_ID);
    const sheet = ss.getSheetByName('Estatística');
    result.aba = sheet ? 'OK' : 'ERRO: aba não encontrada';
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      result.linhas = data.length;
      result.header = data[0];

      const siglaMissing = {};
      let carimbosValidos = 0, carimbosInvalidos = 0;
      const samples = [];
      const latest = {};
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const sigla = _unitCode(r[3]);
        if (!sigla) siglaMissing[String(r[3])] = (siglaMissing[String(r[3])] || 0) + 1;
        const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
        if (isNaN(ts.getTime())) { carimbosInvalidos++; } else {
          carimbosValidos++;
          if (sigla && (!latest[sigla] || ts > latest[sigla].ts)) latest[sigla] = { ts, row: r };
        }
        if (samples.length < 5) samples.push({
          unidadeRaw: r[3], sigla, carimboRaw: r[0], carimboType: typeof r[0],
          turmasRegulares: r[18], alunosRegulares: r[39],
        });
      }
      result.amostras        = samples;
      result.siglaMissingTop = Object.entries(siglaMissing).sort((a,b)=>b[1]-a[1]).slice(0,10);
      result.carimbosValidos   = carimbosValidos;
      result.carimbosInvalidos = carimbosInvalidos;
      result.linhaMaisRecentePorUnidade = Object.fromEntries(
        Object.entries(latest).map(([sigla, v]) => [sigla, {
          carimbo: v.ts, turmasRegulares: v.row[18], alunosRegulares: v.row[39],
        }])
      );
    }
  } catch(e) { result.erro = e.message; }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Diagnóstico da Inadimplência Total (aba inad_total+cobrafix) — roda e olha o Log de Execução
function debugInadTotal() {
  const result = {};
  try {
    const ss    = SpreadsheetApp.openById(INADIMPLENCIA_SHEET_ID);
    const sheet = ss.getSheetByName('inad_total+cobrafix');
    if (!sheet) { result.erro = 'aba não encontrada'; Logger.log(JSON.stringify(result)); return result; }

    const data  = sheet.getDataRange().getValues();
    result.linhas = data.length;
    result.header = data[0];

    const normH = data[0].map(_norm);
    const iUnidade = normH.findIndex(h => h === 'unidade');
    const iTitulos = normH.findIndex(h => h === 'titulos');
    const iValor   = normH.findIndex(h => h.includes('valor') && h.includes('previsto'));
    const iData    = normH.findIndex(h => h.includes('data') && h.includes('relat'));
    result.indices = { iUnidade, iTitulos, iValor, iData };

    const _totalUnit = raw => _sigla(raw) || _unitCode(raw);
    let validDates = 0, invalidDates = 0, blankRows = 0;
    const samples = [];
    const latestDate = {};
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[iUnidade] && !r[iData]) { blankRows++; continue; } // linha vazia no fim da planilha
      const unit = _totalUnit(r[iUnidade]);
      const d    = r[iData] instanceof Date ? r[iData] : new Date(r[iData]);
      if (isNaN(d.getTime())) {
        invalidDates++;
      } else {
        validDates++;
        if (!latestDate[unit] || d > latestDate[unit]) latestDate[unit] = d;
      }
      if (samples.length < 20) samples.push({
        linha: i + 1, unidadeRaw: r[iUnidade], unit,
        dataRelRaw: r[iData], dataRelType: typeof r[iData], isDateObj: r[iData] instanceof Date,
        parsedValid: !isNaN(d.getTime()), titulos: r[iTitulos], valor: r[iValor],
      });
    }
    result.amostrasPrimeirasLinhasComDados = samples;
    result.linhasEmBranco = blankRows;
    result.datasValidas   = validDates;
    result.datasInvalidas = invalidDates;
    result.latestDatePorUnidade = Object.fromEntries(
      Object.entries(latestDate).map(([u, d]) => [u, Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'yyyy-MM-dd')])
    );
  } catch(e) { result.erro = e.message; }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Diagnóstico do gráfico semanal — roda e olha o Log de Execução
function debugInadWeekly() {
  const ss = SpreadsheetApp.openById(INADIMPLENCIA_SHEET_ID);
  const result = {};

  try {
    const pfSheet = ss.getSheetByName('db_inadimplência_pf');
    result.pf_aba = pfSheet ? 'OK' : 'ERRO: aba não encontrada';
    if (pfSheet) {
      const pfData = pfSheet.getDataRange().getValues();
      result.pf_linhas  = pfData.length;
      result.pf_header  = pfData[0];
      const anoCounts = {};
      let validDates = 0, invalidDates = 0;
      const samples = [];
      for (let i = 1; i < pfData.length; i++) {
        const r = pfData[i];
        const ano = parseInt(r[5]);
        anoCounts[ano] = (anoCounts[ano] || 0) + 1;
        const d = r[20] instanceof Date ? r[20] : new Date(r[20]);
        isNaN(d.getTime()) ? invalidDates++ : validDates++;
        if (samples.length < 5) samples.push({ unidade: r[21], ano: r[5], valor: r[8], dataRelRaw: r[20], dataRelType: typeof r[20] });
      }
      result.pf_amostras          = samples;
      result.pf_anoCounts         = anoCounts;
      result.pf_datasRelValidas   = validDates;
      result.pf_datasRelInvalidas = invalidDates;
    }
  } catch(e) { result.pf_erro = e.message; }

  try {
    const baixSheet = ss.getSheetByName('db_baixados');
    result.baixados_aba = baixSheet ? 'OK' : 'ERRO: aba não encontrada';
    if (baixSheet) {
      const baixData = baixSheet.getDataRange().getValues();
      result.baixados_linhas = baixData.length;
      result.baixados_header = baixData[0];
      const anoCounts = {};
      let validIndiv = 0, invalidIndiv = 0, validRel = 0, invalidRel = 0;
      const samples = [];
      for (let i = 1; i < baixData.length; i++) {
        const r = baixData[i];
        const dIndiv = r[4] instanceof Date ? r[4] : new Date(r[4]);
        if (isNaN(dIndiv.getTime())) { invalidIndiv++; } else { validIndiv++; const ano = dIndiv.getFullYear(); anoCounts[ano] = (anoCounts[ano] || 0) + 1; }
        const dRel = r[9] instanceof Date ? r[9] : new Date(r[9]);
        isNaN(dRel.getTime()) ? invalidRel++ : validRel++;
        if (samples.length < 5) samples.push({ unidade: r[10], dataIndivRaw: r[4], dataIndivType: typeof r[4], valor: r[7], dataRelRaw: r[9], dataRelType: typeof r[9] });
      }
      result.baixados_amostras           = samples;
      result.baixados_anoCounts          = anoCounts;
      result.baixados_datasIndivValidas  = validIndiv;
      result.baixados_datasIndivInvalidas = invalidIndiv;
      result.baixados_datasRelValidas    = validRel;
      result.baixados_datasRelInvalidas  = invalidRel;
    }
  } catch(e) { result.baixados_erro = e.message; }

  Logger.log(JSON.stringify(result, null, 2));
  return result;
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
  const curYear  = new Date().getFullYear();
  const prevYear = curYear - 1;
  const { matDay, canDay } = _loadMatCanIndexes(null, curYear, prevYear);
  const result = { curYear, prevYear, matDay, canDay };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Diagnóstico detalhado da leitura da aba Preenchimento — roda e olha o Log de Execução
function debugPreenchimento() {
  const data = SpreadsheetApp.openById(PREENCHIMENTO_SHEET_ID)
    .getSheetByName('Preenchimento').getDataRange().getValues();

  const headerRaw = data[0];
  const header    = headerRaw.map(_norm);
  const cI = {
    unidade:      header.indexOf('unidade'),
    tipoReg:      header.indexOf('tipo de registro'),
    motivoCancel: header.indexOf('motivo do cancelamento'),
    dataMat:      header.indexOf('data da matricula'),
    tipoRegSaida: header.indexOf('tipo de registro saida'),
    dataCancel:   header.indexOf('data de cancelamento'),
  };

  const curYear = new Date().getFullYear();
  const prevYear = curYear - 1;

  const samples = [];
  const siglaMissing = {};
  const tipoRegValues = {}, tipoRegSaidaValues = {};
  let matMatches = 0, canMatches = 0, dateInvalidMat = 0, dateInvalidCan = 0;
  let matNoSigla = 0, canNoSigla = 0;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const unidadeRaw = r[cI.unidade];
    const sigla = _sigla(unidadeRaw);
    if (!sigla) siglaMissing[String(unidadeRaw)] = (siglaMissing[String(unidadeRaw)] || 0) + 1;

    const tipoRegRaw = r[cI.tipoReg];
    tipoRegValues[String(tipoRegRaw)] = (tipoRegValues[String(tipoRegRaw)] || 0) + 1;
    const tipoRegSaidaRaw = r[cI.tipoRegSaida];
    tipoRegSaidaValues[String(tipoRegSaidaRaw)] = (tipoRegSaidaValues[String(tipoRegSaidaRaw)] || 0) + 1;

    if (samples.length < 10) {
      samples.push({
        unidadeRaw, sigla,
        tipoReg: tipoRegRaw, tipoRegNorm: _norm(tipoRegRaw),
        dataMatRaw: r[cI.dataMat], dataMatType: typeof r[cI.dataMat],
        tipoRegSaida: tipoRegSaidaRaw,
        motivoCancel: r[cI.motivoCancel],
        dataCancelRaw: r[cI.dataCancel], dataCancelType: typeof r[cI.dataCancel],
      });
    }

    const tipoReg = _norm(tipoRegRaw);
    if (tipoReg === 'matricula' || tipoReg === 'matricula turma fechada') {
      matMatches++;
      if (!sigla) matNoSigla++;
      const d = r[cI.dataMat] instanceof Date ? r[cI.dataMat] : new Date(r[cI.dataMat]);
      if (isNaN(d.getTime())) dateInvalidMat++;
    }
    const tipoSaida = _norm(tipoRegSaidaRaw);
    if (tipoSaida === 'cancelado') {
      canMatches++;
      if (!sigla) canNoSigla++;
      const d = r[cI.dataCancel] instanceof Date ? r[cI.dataCancel] : new Date(r[cI.dataCancel]);
      if (isNaN(d.getTime())) dateInvalidCan++;
    }
  }

  const result = {
    headerRaw, header, indices: cI,
    totalLinhas: data.length,
    curYear, prevYear,
    matMatches, canMatches, matNoSigla, canNoSigla, dateInvalidMat, dateInvalidCan,
    siglaMissingTop: Object.entries(siglaMissing).sort((a,b)=>b[1]-a[1]).slice(0,10),
    tipoRegValuesTop: Object.entries(tipoRegValues).sort((a,b)=>b[1]-a[1]).slice(0,10),
    tipoRegSaidaValuesTop: Object.entries(tipoRegSaidaValues).sort((a,b)=>b[1]-a[1]).slice(0,10),
    samples,
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// ── Utilitários de cache ──────────────────────────────────────
// Prefixos de todas as chaves de cache usadas pelos endpoints getXOnly
const CACHE_PREFIXES = ['kpi_', 'inad_', 'classavg_', 'nps_'];

// Limpa só as chaves ALL da versão atual. Pra invalidar cache de TODAS as unidades
// de uma vez (ex.: depois de uma mudança grande), basta trocar CACHE_VERSION no topo do arquivo.
function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(CACHE_PREFIXES.map(p => p + CACHE_VERSION + '_ALL'));
  Logger.log('Cache limpo (versão ' + CACHE_VERSION + ').');
}

// ── Pré-aquecimento de cache ──────────────────────────────────
// Gatilho recomendado: a cada 30 min (ScriptApp > Gatilhos)
// Aquece os 4 datasets da visão "todas as unidades", nos MESMOS formatos que os
// endpoints getKpiOnly/getInadOnly/getClassAverageOnly/getNpsOnly gravam e leem.
function warmCache() {
  try {
    const cache = CacheService.getScriptCache();
    const sizes = {};
    const put = (prefix, payload) => {
      try { cache.put(prefix + CACHE_VERSION + '_ALL', payload, 3600); } catch(e) {}
      sizes[prefix] = payload.length;
    };

    put('kpi_',  JSON.stringify({ ok: true, kpi:  _computeKpiData(null) }));
    put('inad_', JSON.stringify({ ok: true, inad: _computeInadData(null) }));

    const weekly = _computeClassAverageWeekly(null);
    const cards  = _computeEstatisticaCards(null);
    cards.atual  = weekly.points.length ? weekly.points[weekly.points.length - 1].valor : null;
    put('classavg_', JSON.stringify({ ok: true, classAverage: { points: weekly.points, cards } }));

    put('nps_', JSON.stringify({ ok: true, nps: _computeNpsData(null) }));

    Logger.log('warmCache OK — ' + Object.keys(sizes).map(k => k + '=' + sizes[k] + ' bytes').join(', '));
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
    const ss = SpreadsheetApp.openById(PREENCHIMENTO_SHEET_ID);
    const s  = ss.getSheetByName('Preenchimento');
    if (s) {
      const h = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
      results.preenchimento         = 'OK — ' + s.getLastRow() + ' linhas';
      results.preenchimento_headers = h.join(' | ');
      results.preenchimento_norm    = h.map(_norm).join(' | ');
    } else { results.preenchimento = 'ERRO: aba não encontrada'; }
  } catch(e) { results.preenchimento = 'ERRO: ' + e.message; }

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

// ── Email semanal para diretores ───────────────────────────────
// Toda segunda-feira ao meio-dia (horário de Brasília), cada diretor recebe um resumo
// com os KPIs da própria unidade; adriane/bruno/peter recebem a visão de todas as unidades.

// Trava de segurança: enquanto for false, checkAndSendWeeklyEmails() (o gatilho automático)
// não envia nada — só retorna. Mude para true quando terminar os ajustes e quiser ligar de
// verdade o envio semanal. Isso NÃO afeta testSendWeeklyEmail() nem sendWeeklyEmails() rodadas
// manualmente pelo editor — só bloqueia o disparo automático pelo gatilho.
const WEEKLY_EMAILS_ENABLED = false;

const DIRETORES_UNIDADE = {
  BF:  [{ email: 'dirbf@brasas.com',     nome: 'Paco' }],
  BG:  [{ email: 'dirbg@brasas.com',     nome: 'Lilly' }],
  BOL: [{ email: 'natasha@brasas.com',   nome: 'Natasha' }, { email: 'alexander@brasas.com', nome: 'Alex' }],
  CG:  [{ email: 'dircg@brasas.com',     nome: 'Rafa' }],
  CH:  [{ email: 'dirch@brasas.com',     nome: 'Anne' }],
  CP:  [{ email: 'dircp@brasas.com',     nome: 'Angie' }],
  CX:  [{ email: 'dircx@brasas.com',     nome: 'Monika' }],
  DT:  [{ email: 'dirdt@brasas.com',     nome: 'Renny' }],
  FG:  [{ email: 'dirfg@brasas.com',     nome: 'Nelson' }],
  GR:  [{ email: 'dirgr@brasas.com',     nome: 'Elica' }],
  IG:  [{ email: 'dirig@brasas.com',     nome: 'Ana' }, { email: 'marcelo.ig@brasas.com', nome: 'Marcelo' }],
  IP:  [{ email: 'dirip@brasas.com',     nome: 'Angie' }],
  IT:  [{ email: 'dirit@brasas.com',     nome: 'Isa' }],
  LJ:  [{ email: 'dirlj@brasas.com',     nome: 'Dennis' }],
  MR:  [{ email: 'dirmr@brasas.com',     nome: 'Joseph' }],
  NI:  [{ email: 'dirni@brasas.com',     nome: 'Clay' }],
  NL:  [{ email: 'dirnl@brasas.com',     nome: 'Kat' }],
  NT:  [{ email: 'dirnt@brasas.com',     nome: 'Kay' }],
  PC:  [{ email: 'dirpc@brasas.com',     nome: 'Ashley' }],
  PN:  [{ email: 'dirpn@brasas.com',     nome: 'Ashley' }],
  PO:  [{ email: 'dirpo@brasas.com',     nome: 'James' }],
  RC:  [{ email: 'dirrc@brasas.com',     nome: 'Emerson' }],
  TJ:  [{ email: 'dirtj@brasas.com',     nome: 'Elica' }],
  TQ:  [{ email: 'dirtq@brasas.com',     nome: 'Paula' }],
  VO:  [{ email: 'dirvo@brasas.com',     nome: 'Wess' }],
  VP:  [{ email: 'dirvp@brasas.com',     nome: 'Brooke' }],
  VQ:  [{ email: 'dirvq@brasas.com',     nome: 'Rachel' }],
};
const EMAILS_TODAS_UNIDADES = [
  { email: 'adriane@brasas.com', nome: 'Adriane' },
  { email: 'bruno@brasas.com',   nome: 'Bruno' },
  { email: 'peter@brasas.com',   nome: 'Peter' },
];

// Junta nomes em uma saudação natural: "Ana", "Ana e Marcelo", "Ana, Bruno e Peter"
function _joinNames(names) {
  if (names.length <= 1) return names[0] || '';
  if (names.length === 2) return names.join(' e ');
  return names.slice(0, -1).join(', ') + ' e ' + names[names.length - 1];
}

// Rode esta função UMA VEZ manualmente pelo editor do Apps Script para instalar o gatilho.
// O gatilho roda de hora em hora; checkAndSendWeeklyEmails() decide, usando o horário de
// Brasília, se é o momento certo de disparar — assim não depende do fuso-horário do projeto.
function setupWeeklyEmailTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkAndSendWeeklyEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkAndSendWeeklyEmails').timeBased().everyHours(1).create();
}

// Disparado pelo gatilho horário criado em setupWeeklyEmailTrigger().
// Só envia às segundas-feiras, às 12h no horário de Brasília, e no máximo uma vez por dia
// (guarda a data do último envio em Properties para não duplicar se o gatilho disparar de novo).
function checkAndSendWeeklyEmails() {
  if (!WEEKLY_EMAILS_ENABLED) return; // trava de segurança — ver comentário em WEEKLY_EMAILS_ENABLED

  const tz      = 'America/Sao_Paulo';
  const now     = new Date();
  const weekday = Utilities.formatDate(now, tz, 'u'); // 1 = segunda-feira
  const hour    = Number(Utilities.formatDate(now, tz, 'H'));
  if (weekday !== '1' || hour !== 12) return;

  const props    = PropertiesService.getScriptProperties();
  const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  if (props.getProperty('weeklyEmailSentOn') === todayKey) return;

  sendWeeklyEmails();
  props.setProperty('weeklyEmailSentOn', todayKey);
}

// Envia o resumo semanal para cada diretor de unidade + a visão de todas as unidades.
// Pode ser rodada manualmente pelo editor do Apps Script para testar/reenviar.
function sendWeeklyEmails() {
  Object.keys(DIRETORES_UNIDADE).forEach(sigla => {
    try {
      const dirs = DIRETORES_UNIDADE[sigla];
      const nome = _joinNames(dirs.map(d => d.nome));
      const html = _buildWeeklyEmailHtml([sigla], sigla, nome);
      MailApp.sendEmail({
        to:       dirs.map(d => d.email).join(','),
        subject:  `Boa Semana - ${sigla}`,
        htmlBody: html,
      });
    } catch(e) { Logger.log('Erro ao enviar email da unidade ' + sigla + ': ' + e.message); }
  });

  // Visão de todas as unidades: um email individual para cada destinatário, com o nome
  // no assunto e na saudação. O corpo é montado uma vez só (com um placeholder no lugar
  // do nome) para não recomputar KPI/Inad/NPS três vezes.
  try {
    const htmlTemplate = _buildWeeklyEmailHtml(null, 'Todas as unidades', '__NOME__');
    EMAILS_TODAS_UNIDADES.forEach(dest => {
      try {
        MailApp.sendEmail({
          to:       dest.email,
          subject:  `Boa Semana - ${dest.nome}`,
          htmlBody: htmlTemplate.replace('__NOME__', dest.nome),
        });
      } catch(e) { Logger.log('Erro ao enviar email de todas as unidades para ' + dest.email + ': ' + e.message); }
    });
  } catch(e) { Logger.log('Erro ao enviar email de todas as unidades: ' + e.message); }
}

// Teste rápido: manda o resumo de UMA unidade só para o seu próprio email (não para o diretor real).
// Troque SIGLA_TESTE/EMAIL_TESTE abaixo e rode esta função pelo editor do Apps Script.
function testSendWeeklyEmail() {
  const SIGLA_TESTE = 'RC';                    // sigla da unidade (ex: 'RC', 'BOL') ou 'ALL' p/ todas as unidades
  const EMAIL_TESTE = 'adriane@brasas.com';    // para onde o email de teste vai

  const sigla = String(SIGLA_TESTE||'').trim().toUpperCase();
  const isAll = sigla === 'ALL' || sigla === 'TODAS';
  if (!isAll && !DIRETORES_UNIDADE[sigla]) {
    Logger.log('Unidade "' + sigla + '" inválida. Use uma sigla de DIRETORES_UNIDADE ou "ALL".');
    return;
  }

  const units = isAll ? null : [sigla];
  const label = isAll ? 'Todas as unidades' : sigla;
  // No modo ALL o envio real é individual (um email por destinatário, nome no assunto);
  // o teste usa o nome de quem bater com EMAIL_TESTE, senão o do primeiro da lista
  const destAll = EMAILS_TODAS_UNIDADES.find(d => d.email === EMAIL_TESTE) || EMAILS_TODAS_UNIDADES[0];
  const nome  = isAll
    ? destAll.nome
    : _joinNames(DIRETORES_UNIDADE[sigla].map(d => d.nome));
  const html  = _buildWeeklyEmailHtml(units, label, nome);

  MailApp.sendEmail({
    to:       EMAIL_TESTE,
    subject:  isAll ? `[TESTE] Boa Semana - ${nome}` : `[TESTE] Boa Semana - ${label}`,
    htmlBody: html,
  });
  Logger.log('Email de teste enviado para ' + EMAIL_TESTE + ' — unidade: ' + label);
}

function _dateBR(key) {
  if (!key) return '';
  const p = key.split('-');
  return `${p[2]}/${p[1]}/${p[0]}`;
}

// Cores da barra do medidor de NPS no email (tema claro)
function _npsColorHex(val) {
  if (val === null || val === undefined) return '#94a3b8';
  if (val >= 70) return '#16a34a';
  if (val < 40)  return '#dc2626';
  return '#f59e0b';
}

function _fmtIntEmail(n) { return Math.round(n||0).toLocaleString('pt-BR'); }
function _fmtBRLEmail(n) { return 'R$ ' + (n||0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function _fmtPctEmail(n) { return (n === null || n === undefined) ? '—' : Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%'; }

// Uma métrica dentro do card de ano, igual ao ".year-metric" do painel
function _yearMetricHtml(label, value, accent) {
  return `<td style="width:50%;text-align:center;padding:10px 6px;">
    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:#4a617c;margin-bottom:2px;">${label}</div>
    <div style="font-size:16px;font-weight:700;color:${accent ? '#dc2626' : '#0a1628'};">${value}</div>
  </td>`;
}

// Card "Inadimplência por Ano" de um ano específico, igual ao ".year-card" do painel
function _yearCardEmailHtml(label, yc) {
  if (!yc) return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
    <tr><td style="text-align:center;padding:10px 16px 8px;font-size:12px;font-weight:700;letter-spacing:.04em;color:#1f3a5c;background:#eef3f8;border-bottom:1px solid #d6e0ea;border-radius:15px 15px 0 0;">${label}</td></tr>
    <tr><td style="text-align:center;padding:24px 16px;font-size:13px;color:#64748b;">Nenhum dado encontrado.</td></tr>
  </table>`;

  const hasI = yc.inadpf   !== null;
  const hasC = yc.cobrafix !== null;

  let metrics;
  if (hasI && hasC) {
    metrics = [
      ['Inad. Acumulada',        _fmtBRLEmail(yc.inadpf.inadAcumulada),   true],
      ['Títulos a Receber',      _fmtIntEmail(yc.inadpf.titulosAReceber), false],
      ['Títulos Cobrafix',       _fmtIntEmail(yc.cobrafix.titulos),       false],
      ['Valor Títulos Cobrafix', _fmtBRLEmail(yc.cobrafix.valor),         false],
      ['Inadimplência %',        _fmtPctEmail(yc.inadpf.pctInad),         false],
      ['Ticket Médio',           _fmtBRLEmail(yc.inadpf.ticketMedio),     true],
    ];
  } else if (hasI) {
    metrics = [
      ['Inad. Acumulada',   _fmtBRLEmail(yc.inadpf.inadAcumulada),   true],
      ['Títulos a Receber', _fmtIntEmail(yc.inadpf.titulosAReceber), false],
      ['Ticket Médio',      _fmtBRLEmail(yc.inadpf.ticketMedio),     true],
      ['Inadimplência %',   _fmtPctEmail(yc.inadpf.pctInad),         false],
    ];
  } else if (hasC) {
    metrics = [
      ['Títulos Cobrafix',       _fmtIntEmail(yc.cobrafix.titulos)],
      ['Valor Títulos Cobrafix', _fmtBRLEmail(yc.cobrafix.valor)],
    ];
  } else {
    metrics = [];
  }

  const rows = [];
  for (let i = 0; i < metrics.length; i += 2) {
    const a = metrics[i], b = metrics[i + 1];
    rows.push(`<tr>${_yearMetricHtml(a[0], a[1], a[2])}${b ? _yearMetricHtml(b[0], b[1], b[2]) : '<td style="width:50%;"></td>'}</tr>`);
  }

  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
    <tr><td colspan="2" style="text-align:center;padding:10px 16px 8px;font-size:12px;font-weight:700;letter-spacing:.04em;color:#1f3a5c;background:#eef3f8;border-bottom:1px solid #d6e0ea;border-radius:15px 15px 0 0;">${label}</td></tr>
    ${rows.join('')}
  </table>`;
}

// Pílula de seção, igual ao ".section-title" do painel
function _sectionTitleHtml(label) {
  return `<div style="text-align:center;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:#1f3a5c;background:#eef3f8;border:1px solid #d6e0ea;border-radius:12px;padding:10px 20px;">${label}</div>`;
}

// Badge de variação %, igual ao ".kpi-card-badge" do painel (verde/vermelho conforme invert)
function _pctBadgeHtml(pct, invert) {
  if (pct === null || pct === undefined) return '';
  if (+pct === 0) return `<div style="display:inline-block;margin-top:6px;font-size:11px;font-weight:600;padding:2px 7px;border-radius:20px;background:#e8eef4;color:#4a617c;">0,0%</div>`;
  const isUp   = pct > 0;
  const isGood = invert ? !isUp : isUp;
  const bg     = isGood ? 'rgba(22,163,74,.14)' : 'rgba(220,38,38,.12)';
  const color  = isGood ? '#16a34a' : '#dc2626';
  const sign   = isUp ? '▲' : '▼';
  return `<div style="display:inline-block;margin-top:6px;font-size:11px;font-weight:600;padding:2px 7px;border-radius:20px;background:${bg};color:${color};">${sign} ${Math.abs(pct).toFixed(1).replace('.',',')}%</div>`;
}

// Card com 3 colunas (Mês/Trimestre/Ano), igual ao ".kpi-group" do painel
function _kpiGroupHtml(groupLabel, data, invert) {
  const cols = [
    { period: 'Este Mês',       d: data.mes },
    { period: 'Este Trimestre', d: data.tri },
    { period: 'Este Ano',       d: data.ano },
  ];
  const tds = cols.map((c, i) => {
    const val    = c.d ? c.d.val : 0;
    const pct    = c.d ? c.d.pct : null;
    const isNeg  = val < 0;
    const border = i < cols.length - 1 ? 'border-right:1px solid #e2e8f0;' : '';
    return `<td style="padding:16px 8px 18px;text-align:center;${border}">
      <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:#4a617c;margin-bottom:8px;">${c.period}</div>
      <div style="font-size:22px;font-weight:700;letter-spacing:-.02em;color:${isNeg ? '#dc2626' : '#0a1628'};">${_fmtIntEmail(val)}</div>
      ${_pctBadgeHtml(pct, invert)}
    </td>`;
  }).join('');

  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
    <tr><td colspan="3" style="text-align:center;padding:11px 16px 9px;font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#1f3a5c;background:#eef3f8;border-bottom:1px solid #d6e0ea;border-radius:15px 15px 0 0;">${groupLabel}</td></tr>
    <tr>${tds}</tr>
  </table>`;
}

// Card único, igual ao ".inad-kpi" do painel
function _statCardHtml(label, value) {
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
    <tr><td style="padding:18px 20px;text-align:center;">
      <div style="font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#4a617c;margin-bottom:8px;">${label}</div>
      <div style="font-size:22px;font-weight:700;letter-spacing:-.02em;color:#0a1628;">${value}</div>
    </td></tr>
  </table>`;
}

// Medidor de NPS com barra (-100 a 100), igual ao ".nps-meter-card" do painel — construído com
// larguras de <td> (não position:absolute) pra funcionar em clientes de email como o Outlook
function _npsMeterHtml(label, data) {
  const val      = (data && data.nps !== null && data.nps !== undefined) ? data.nps : null;
  const valLabel = val === null ? '—' : val.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

  let barHtml;
  if (val === null) {
    barHtml = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="background:#e2e8f0;border-radius:4px;height:8px;line-height:8px;font-size:0;">&nbsp;</td></tr></table>`;
  } else {
    const clamped   = Math.max(-100, Math.min(100, val));
    const pct       = (clamped + 100) / 200 * 100;
    const color     = _npsColorHex(val);
    const leftPct   = clamped >= 0 ? 50 : pct;
    const widthPct  = Math.abs(pct - 50);
    const rightPct  = 100 - leftPct - widthPct;
    const track     = 'background:#e2e8f0;height:8px;line-height:8px;font-size:0;';
    barHtml = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
      ${leftPct  > 0.5 ? `<td width="${leftPct}%"  style="${track}">&nbsp;</td>` : ''}
      ${widthPct > 0.5 ? `<td width="${widthPct}%" style="background:${color};height:8px;line-height:8px;font-size:0;">&nbsp;</td>` : ''}
      ${rightPct > 0.5 ? `<td width="${rightPct}%" style="${track}">&nbsp;</td>` : ''}
    </tr></table>`;
  }

  const subLabel = (!data || !data.total) ? 'sem respostas' : `${data.total} resposta${data.total === 1 ? '' : 's'}`;
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
    <tr><td style="padding:18px 18px 16px;text-align:center;">
      <div style="font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#4a617c;margin-bottom:8px;">${label}</div>
      <div style="font-size:24px;font-weight:700;letter-spacing:-.02em;color:#0a1628;margin-bottom:12px;">${valLabel}</div>
      ${barHtml}
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px;"><tr>
        <td style="font-size:10px;color:#64748b;text-align:left;">-100</td>
        <td style="font-size:10px;color:#64748b;text-align:center;">0</td>
        <td style="font-size:10px;color:#64748b;text-align:right;">100</td>
      </tr></table>
      <div style="font-size:11px;color:#64748b;margin-top:8px;">${subLabel}</div>
    </td></tr>
  </table>`;
}

// Grade 2x2 dos medidores de NPS
function _npsMetersGridHtml(cards) {
  const rows = [];
  for (let i = 0; i < cards.length; i += 2) {
    const left  = cards[i];
    const right = cards[i + 1];
    rows.push(`<tr>
      <td style="width:50%;padding:0 8px 16px 0;vertical-align:top;">${_npsMeterHtml(left.label, left.data)}</td>
      <td style="width:50%;padding:0 0 16px 8px;vertical-align:top;">${right ? _npsMeterHtml(right.label, right.data) : ''}</td>
    </tr>`);
  }
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${rows.join('')}</table>`;
}

// ── Ranking de inadimplência por unidade ──────────────────────
// Mesma lógica da "Inadimplência Total": para cada unidade, considera só as linhas da
// data de relatório mais recente daquela unidade (aba inad_total+cobrafix), e ordena
// do maior para o menor valor previsto. Usado apenas no email "todas as unidades".
function _computeInadRanking() {
  const totalData = _getSheetData(INADIMPLENCIA_SHEET_ID, 'inad_total+cobrafix');
  const totalH    = totalData[0].map(_norm);
  const tI = {
    unidade:   totalH.findIndex(h => h === 'unidade'),
    titulos:   totalH.findIndex(h => h === 'titulos'),
    valorPrev: totalH.findIndex(h => h.includes('valor') && h.includes('previsto')),
    dataRel:   totalH.findIndex(h => h.includes('data') && h.includes('relat')),
  };
  const _totalUnit = raw => _sigla(raw) || _unitCode(raw);

  const latestDate = {};
  for (let i = 1; i < totalData.length; i++) {
    const unit = _totalUnit(totalData[i][tI.unidade]);
    if (!unit) continue;
    const d = totalData[i][tI.dataRel] instanceof Date
      ? totalData[i][tI.dataRel] : new Date(totalData[i][tI.dataRel]);
    if (isNaN(d.getTime())) continue;
    if (!latestDate[unit] || d > latestDate[unit]) latestDate[unit] = d;
  }

  const byUnit = {};
  for (let i = 1; i < totalData.length; i++) {
    const unit = _totalUnit(totalData[i][tI.unidade]);
    if (!unit || !latestDate[unit]) continue;
    const d = totalData[i][tI.dataRel] instanceof Date
      ? totalData[i][tI.dataRel] : new Date(totalData[i][tI.dataRel]);
    if (isNaN(d.getTime()) || d.getTime() !== latestDate[unit].getTime()) continue;
    if (!byUnit[unit]) byUnit[unit] = { titulos: 0, valor: 0 };
    byUnit[unit].titulos += parseFloat(totalData[i][tI.titulos])   || 0;
    byUnit[unit].valor   += parseFloat(totalData[i][tI.valorPrev]) || 0;
  }

  return Object.keys(byUnit)
    .map(unidade => ({ unidade, valor: byUnit[unidade].valor, titulos: Math.round(byUnit[unidade].titulos) }))
    .sort((a, b) => b.valor - a.valor);
}

// Escapa conteúdo vindo das planilhas (nomes, motivos, críticas) antes de inserir no HTML do email
function _escEmail(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Tabela de ranking de inadimplência por unidade (maior → menor valor previsto).
// As 3 maiores inadimplências ficam em vermelho; linha final com o total geral.
function _inadRankingEmailHtml(ranking) {
  if (!ranking || !ranking.length) {
    return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
      <tr><td style="text-align:center;padding:20px;font-size:12px;color:#64748b;">Nenhum dado encontrado.</td></tr>
    </table>`;
  }

  const th = (label, align) => `<th style="padding:9px 10px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:#4a617c;background:#eef3f8;border-bottom:1px solid #d6e0ea;text-align:${align || 'left'};">${label}</th>`;

  const body = ranking.map((r, i) => {
    const top3      = i < 3;
    const valColor  = top3 ? '#dc2626' : '#0a1628';
    const rowBg     = top3 ? 'background:rgba(220,38,38,.04);' : '';
    return `<tr>
      <td style="padding:9px 10px;font-size:12px;font-weight:700;color:${top3 ? '#dc2626' : '#64748b'};border-bottom:1px solid #e2e8f0;${rowBg}">${i + 1}º</td>
      <td style="padding:9px 10px;font-size:12px;font-weight:700;color:#1f3a5c;border-bottom:1px solid #e2e8f0;${rowBg}">${_escEmail(r.unidade)}</td>
      <td style="padding:9px 10px;font-size:12px;font-weight:700;color:${valColor};border-bottom:1px solid #e2e8f0;text-align:right;${rowBg}">${_fmtBRLEmail(r.valor)}</td>
      <td style="padding:9px 10px;font-size:12px;color:#1f3a5c;border-bottom:1px solid #e2e8f0;text-align:right;${rowBg}">${_fmtIntEmail(r.titulos)}</td>
    </tr>`;
  }).join('');

  const totalValor   = ranking.reduce((s, r) => s + r.valor, 0);
  const totalTitulos = ranking.reduce((s, r) => s + r.titulos, 0);
  const totalRow = `<tr>
    <td colspan="2" style="padding:10px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#4a617c;background:#eef3f8;">Total</td>
    <td style="padding:10px;font-size:12px;font-weight:700;color:#0a1628;background:#eef3f8;text-align:right;">${_fmtBRLEmail(totalValor)}</td>
    <td style="padding:10px;font-size:12px;font-weight:700;color:#0a1628;background:#eef3f8;text-align:right;">${_fmtIntEmail(totalTitulos)}</td>
  </tr>`;

  return `<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;border-collapse:separate;">
    <tr>${th('#')}${th('Unidade')}${th('Valor Previsto', 'right')}${th('Títulos', 'right')}</tr>
    ${body}
    ${totalRow}
  </table>`;
}

// Tabela de respostas recentes do NPS (últimos 15 dias), igual à ".nps-table" do painel.
// Assim como no painel, mostra as respostas de TODAS as unidades (sem filtro).
function _npsTableEmailHtml(rows) {
  if (!rows || !rows.length) {
    return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;">
      <tr><td style="text-align:center;padding:20px;font-size:12px;color:#64748b;">Sem respostas nos últimos 15 dias.</td></tr>
    </table>`;
  }

  const MAX_LINHAS = 30; // evita emails gigantes (Gmail corta mensagens acima de ~102KB)
  const shown = rows.slice(0, MAX_LINHAS);

  const th = label => `<th style="padding:8px 6px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:#4a617c;background:#eef3f8;border-bottom:1px solid #d6e0ea;text-align:left;">${label}</th>`;
  const td = (content, extra) => `<td style="padding:8px 6px;font-size:11px;color:#1f3a5c;border-bottom:1px solid #e2e8f0;vertical-align:top;${extra || ''}">${content}</td>`;

  const body = shown.map(r => {
    const notaColor = r.nota === null ? '#64748b' : (r.nota >= 9 ? '#16a34a' : (r.nota <= 6 ? '#dc2626' : '#1f3a5c'));
    return `<tr>
      ${td(_escEmail(r.unidade))}
      ${td(_dateBR(r.data), 'white-space:nowrap;')}
      ${td(_escEmail(r.nome))}
      ${td(_escEmail(r.book))}
      ${td(r.nota === null ? '—' : r.nota, `font-weight:700;color:${notaColor};`)}
      ${td(_escEmail(r.motivo))}
      ${td(_escEmail(r.criticas))}
    </tr>`;
  }).join('');

  const more = rows.length > MAX_LINHAS
    ? `<tr><td colspan="7" style="text-align:center;padding:10px;font-size:11px;color:#64748b;">+ ${rows.length - MAX_LINHAS} resposta(s) — veja todas no painel completo</td></tr>`
    : '';

  return `<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d6e0ea;border-radius:16px;border-collapse:separate;">
    <tr>${th('Unidade')}${th('Data')}${th('Nome')}${th('Book')}${th('Nota')}${th('Motivo do Cancelamento')}${th('Críticas, Sugestões e Elogios')}</tr>
    ${body}
    ${more}
  </table>`;
}

function _buildWeeklyEmailHtml(units, unidadeLabel, nomeDiretor) {
  const kpi  = _computeKpiData(units);
  const inad = _computeInadData(units);
  const nps  = _computeNpsData(units);

  const npsCards = [];
  if (nps.showAdultsKids) {
    npsCards.push({ label: 'Adults', data: nps.adults });
    npsCards.push({ label: 'Kids',   data: nps.kids });
  }
  if (nps.showCancelados) npsCards.push({ label: 'Cancelados BOL', data: nps.cancelados });
  npsCards.push({ label: 'Mudança de Nível', data: nps.mudancaNivel });

  const inadCardsHtml = [
    { label: 'Valor Previsto',        value: _fmtBRLEmail(inad.total.valorPrevisto) },
    { label: 'Quantidade de Títulos', value: _fmtIntEmail(inad.total.titulos) },
    { label: 'Ticket Médio',          value: _fmtBRLEmail(inad.total.ticketMedio) },
  ].map(it => `<tr><td style="padding-bottom:12px;">${_statCardHtml(it.label, it.value)}</td></tr>`).join('');

  const anoAtual   = String(new Date().getFullYear());
  const yearCardAno = _yearCardEmailHtml(anoAtual, inad.yearCards[anoAtual] || null);

  // Ranking por unidade: só para quem recebe a visão de todas as unidades (units === null)
  const rankingHtml = units === null ? `
    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Ranking de Inadimplência por Unidade')}</td></tr>
    <tr><td style="padding-bottom:22px;">${_inadRankingEmailHtml(_computeInadRanking())}</td></tr>` : '';

  return `
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;padding:32px 0;">
  <tr><td align="center">
  <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;font-family:Arial,Helvetica,sans-serif;">

    <tr><td style="text-align:center;padding-bottom:24px;">
      <div style="font-size:22px;font-weight:700;color:#0a1628;letter-spacing:-.02em;">Boa semana${nomeDiretor ? ', ' + nomeDiretor : ''}!</div>
      <div style="font-size:12px;color:#4a617c;font-style:italic;margin-top:4px;">BRASAS English Course</div>
      <div style="font-size:13px;color:#4a617c;margin-top:14px;">Unidade: <strong style="color:#0a1628;">${unidadeLabel}</strong> &middot; ${Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy')}</div>
    </td></tr>

    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Matrículas')}</td></tr>
    <tr><td style="padding-bottom:22px;">${_kpiGroupHtml('Matrículas', kpi.matriculas, false)}</td></tr>

    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Cancelados')}</td></tr>
    <tr><td style="padding-bottom:22px;">${_kpiGroupHtml('Cancelados', kpi.cancelados, true)}</td></tr>

    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Saldo')}</td></tr>
    <tr><td style="padding-bottom:22px;">${_kpiGroupHtml('Saldo', kpi.saldo, false)}</td></tr>

    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Inadimplência Total')}</td></tr>
    <tr><td style="padding-bottom:22px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">${inadCardsHtml}</table></td></tr>

    <tr><td style="padding-bottom:14px;">${_sectionTitleHtml('Inadimplência ' + anoAtual)}</td></tr>
    <tr><td style="padding-bottom:22px;">${yearCardAno}</td></tr>
    ${rankingHtml}

    <tr><td style="padding-bottom:14px;padding-top:12px;">${_sectionTitleHtml('NPS')}</td></tr>
    <tr><td style="padding-bottom:6px;">${_npsMetersGridHtml(npsCards)}</td></tr>
    ${nps.lastResponse ? `<tr><td style="text-align:center;font-size:12px;color:#64748b;padding-bottom:20px;">Última resposta: ${_dateBR(nps.lastResponse)}</td></tr>` : ''}

    <tr><td style="padding-bottom:14px;padding-top:4px;">${_sectionTitleHtml('Respostas recentes do NPS — últimos 15 dias')}</td></tr>
    <tr><td style="padding-bottom:22px;">${_npsTableEmailHtml(nps.recent)}</td></tr>

    <tr><td style="text-align:center;padding-top:8px;">
      <a href="${HUB_URL}" style="display:inline-block;background:#1e3a5f;color:#ffffff;text-decoration:none;padding:12px 24px;border-radius:8px;font-size:14px;font-weight:600;">Abrir o painel completo →</a>
    </td></tr>

  </table>
  </td></tr>
  </table>`;
}
