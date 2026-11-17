/**
 * Recebe os dados do Health Connect enviados pelo app Android "HC Webhook" (com.hcwebhook.app).
 * O Mi Fitness grava no Health Connect; o HC Webhook manda de tempos em tempos um POST JSON para
 * <URL /exec>?k=<chave do usuário>. A chave é gerada no app Saúde (Ajustes > Relógio) e vale só para gravar.
 *
 * O que entra:
 *  - exercise / exercise_sessions → aba Atividades (sem duplicar com histórico, Strava ou envio repetido)
 *  - steps → Diario.passos (soma do dia)      - sleep → Diario.sono (horas, dia em que acordou)
 *  - resting_heart_rate (ou o menor heart_rate do dia) → Diario.fc_rep
 * Os intervalos recebidos ficam na aba HC_bruto (um por início+fim), assim reenvio ou atualização
 * do mesmo intervalo substitui o valor em vez de somar de novo.
 */
var HC_TIPOS = { RUNNING: 'Corrida', RUNNING_TREADMILL: 'Corrida', WALKING: 'Caminhada', HIKING: 'Trilha', BIKING: 'Bicicleta',
  BIKING_STATIONARY: 'Bicicleta', STRENGTH_TRAINING: 'Musculação', WEIGHTLIFTING: 'Musculação', CALISTHENICS: 'Funcional',
  HIGH_INTENSITY_INTERVAL_TRAINING: 'Funcional', EXERCISE_CLASS: 'Funcional', BOOT_CAMP: 'Funcional', SOCCER: 'Futebol',
  FOOTBALL_AMERICAN: 'Futebol', SWIMMING_POOL: 'Natação', SWIMMING_OPEN_WATER: 'Natação', YOGA: 'Alongamento',
  PILATES: 'Alongamento', STRETCHING: 'Alongamento', ELLIPTICAL: 'Elíptico', STAIR_CLIMBING: 'Escada', STAIR_CLIMBING_MACHINE: 'Escada',
  ROWING: 'Remo', ROWING_MACHINE: 'Remo', VOLLEYBALL: 'Vôlei', BASKETBALL: 'Basquete', TENNIS: 'Tênis', GOLF: 'Golfe',
  DANCING: 'Dança', OTHER_WORKOUT: 'Treino livre' };
var HC_MAX_ITENS = 5000;

// ---------- Chave por usuário (pelo app) ----------
function hcGerarLink_(u) {
  var p = props_(), velha = p.getProperty('HCW_USER_' + u.id);
  if (velha) p.deleteProperty('HCW_' + velha);
  var k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  p.setProperty('HCW_' + k, u.id); p.setProperty('HCW_USER_' + u.id, k);
  auditLog_('update', 'healthconnect', u.id, null, { link: 'gerado' }, '', u.id);
  var url = ScriptApp.getService().getUrl();
  return { ok: true, url: url + '?k=' + k, hc: hcStatus_(u.id) };
}
function hcDesligar_(u) {
  var p = props_(), k = p.getProperty('HCW_USER_' + u.id);
  if (k) p.deleteProperty('HCW_' + k);
  p.deleteProperty('HCW_USER_' + u.id);
  auditLog_('update', 'healthconnect', u.id, null, { link: 'desligado' }, '', u.id);
  return { ok: true, hc: hcStatus_(u.id) };
}
function hcStatus_(uid) {
  var p = props_();
  return { ligado: !!p.getProperty('HCW_USER_' + uid), ultima: p.getProperty('HCW_LAST_' + uid) || '' };
}

// ---------- Recebimento ----------
function hcDono_(k) { // chave -> usuário ativo, ou null
  if (!/^[a-f0-9]{48}$/.test(String(k || ''))) return null;
  var p = props_(), uid = p.getProperty('HCW_' + k), u = uid ? userById_(uid) : null;
  return u && u.ativo && p.getProperty('HCW_USER_' + uid) === k ? u : null;
}
var HC_DIAS = 60, HC_POR_HORA = 40;
function hcWebhook_(k, body) {
  var u = hcDono_(k);
  if (!u) return { ok: false, error: 'unauthorized' };
  var cache = CacheService.getScriptCache(), rk = 'hc_rate_' + u.id, n = parseInt(cache.get(rk) || '0', 10);
  if (n >= HC_POR_HORA) return { ok: false, error: 'limite', message: 'Muitos envios nesta hora.' };
  cache.put(rk, String(n + 1), 3600);
  if (!body || typeof body !== 'object') return { ok: false, error: 'requisicao_invalida' };
  var exs = arr_(body.exercise).concat(arr_(body.exercise_sessions)), passos = arr_(body.steps), sono = arr_(body.sleep);
  var fcr = arr_(body.resting_heart_rate), fc = arr_(body.heart_rate);
  if (exs.length + passos.length + sono.length + fcr.length + fc.length > HC_MAX_ITENS) return { ok: false, error: 'grande_demais' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return { ok: false, error: 'ocupado' }; } // o app tenta de novo
  try {
    if (!hcDono_(k)) return { ok: false, error: 'unauthorized' }; // revogada enquanto esperava a trava
    var ss = SpreadsheetApp.getActiveSpreadsheet(), res = { ok: true, treinos: 0, dias: 0 };
    res.treinos = hcTreinos_(ss, u.id, exs);
    var bruto = hcBruto_(ss), datas = {};
    passos.forEach(function (r) {
      var i = iso_(r.start_time), f = iso_(r.end_time), n = Number(r.count);
      if (!i || !f || f < i || !janela_(i) || !(n >= 0) || n > 100000) return;
      bruto.upsert(u.id, 'passos', i.toISOString(), f.toISOString(), Math.round(n)); datas[diaSP_(i)] = 1;
    });
    sono.forEach(function (r) {
      var f = iso_(r.session_end_time || r.end_time), s = Number(r.duration_seconds);
      if (!f || !janela_(f) || !(s > 0) || s > 86400) return;
      // a sessão é identificada pelo fim (o início não vem): duração corrigida substitui, não soma
      bruto.upsert(u.id, 'sono', '', f.toISOString(), Math.round(s)); datas[diaSP_(f)] = 1;
    });
    // FC de repouso: usa a de repouso; se não vier, o menor batimento do dia (como no histórico)
    fcr.concat(fc.map(function (r) { return { bpm: r.bpm, time: r.time, minimo: true }; })).forEach(function (r) {
      var t = iso_(r.time), b = Number(r.bpm);
      if (!t || !janela_(t) || !(b >= 25 && b <= 150)) return;
      var d = diaSP_(t);
      bruto.upsertMin(u.id, r.minimo ? 'fc_min' : 'fc_rep', d, Math.round(b)); datas[d] = 1;
    });
    bruto.flush();
    res.dias = hcAtualizaDiario_(ss, u.id, Object.keys(datas), bruto);
    props_().setProperty('HCW_LAST_' + u.id, Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    if (res.treinos || res.dias) auditLog_('import', 'healthconnect', '', null, res, '', u.id);
    return res;
  } catch (err) {
    return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) };
  } finally { lock.releaseLock(); }
}

function hcTreinos_(ss, uid, exs) {
  if (!exs.length) return 0;
  var imp = importados_(ss, uid), novos = [], lote = {};
  exs.forEach(function (r) {
    var i = iso_(r.start_time), f = iso_(r.end_time);
    if (!i || !f || f <= i || !janela_(i)) return;
    var ini = Math.floor(i.getTime() / 1000), chave = 'hc:' + ini;
    // mesmo treino reenviado (chave) ou já vindo do histórico/Strava (início a até 3 min) não entra
    if (imp.ids[chave] || lote[ini] || ehDuplicado_(imp.todos, ini)) return;
    var seg = Number(r.duration_seconds) > 0 ? Number(r.duration_seconds) : (f - i) / 1000;
    var tipoHC = String(r.type || '').toUpperCase().replace(/^EXERCISE_TYPE_/, '');
    var v = validate_(SCHEMA.atividade, {
      data: diaSP_(i), tipo: HC_TIPOS[tipoHC] || nomeTipo_(tipoHC), minutos: Math.max(1, Math.round(seg / 60)),
      distancia: Number(r.distance_meters) > 0 ? Math.round(Number(r.distance_meters) / 10) / 100 : '',
      passos: Number(r.steps) > 0 ? Math.round(Number(r.steps)) : '',
      kcal: Number(r.active_calories_kcal) > 0 ? Math.round(Number(r.active_calories_kcal)) : '',
      fc: '', intensidade: '', obs: 'Relógio (Health Connect)'
    }, true);
    if (!v.ok) return;
    lote[ini] = 1; novos.push({ ini: ini, origem: 'Health Connect', chave: chave, values: v.values });
  });
  gravarComRegistro_(ss, imp.sheet, uid, novos);
  return novos.length;
}

// Recalcula passos (soma), sono (soma das sessões que terminaram no dia) e FC de repouso nas datas tocadas.
function hcAtualizaDiario_(ss, uid, datas, bruto) {
  if (!datas.length) return 0;
  var sc = SCHEMA.diario, sh = sheetFor_(ss, sc), n = sc.campos.length, pos = {};
  sc.campos.forEach(function (c, i) { pos[c[0]] = COL.CAMPOS + i; });
  var linhas = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 3).getDisplayValues().forEach(function (r, i) {
    if (String(r[1]) === uid) linhas[dataTxt_(r[2])] = i + 2;
  });
  var tot = bruto.totais(uid), feitos = 0, hoje = hoje_(), novos = [];
  datas.forEach(function (d) {
    if (!validDate_(d) || d > hoje) return;
    var t = tot[d] || {}, val = {};
    if (t.passos != null && t.passos <= 150000) val.passos = t.passos;
    if (t.sono != null && t.sono <= 86400) val.sono = Math.round(t.sono / 360) / 10;
    var fcv = t.fc_rep != null ? t.fc_rep : t.fc_min;
    if (fcv != null) val.fc_rep = fcv;
    if (!Object.keys(val).length) return;
    var row = linhas[d];
    if (row) {
      Object.keys(val).forEach(function (k) { sh.getRange(row, pos[k]).setValue(val[k]); });
      sh.getRange(row, n + 4).setValue(new Date());
    } else {
      var v = validate_(sc, { data: d, sono: val.sono == null ? '' : val.sono, agua_l: '', passos: val.passos == null ? '' : val.passos,
        pa_sis: '', pa_dia: '', fc_rep: val.fc_rep == null ? '' : val.fc_rep, obs: 'Relógio (Health Connect)' }, true);
      if (!v.ok) return;
      novos.push(v.values);
    }
    feitos++;
  });
  gravarLote_(ss, sc, uid, novos);
  return feitos;
}

// Aba HC_bruto: Usuário | Tipo | Início | Fim | Valor | Recebido em. Chave = usuário+tipo+início+fim.
function hcBruto_(ss) {
  var sh = ss.getSheetByName('HC_bruto');
  if (!sh) {
    sh = ss.insertSheet('HC_bruto');
    sh.getRange(1, 1, 1, 6).setValues([['Usuário', 'Tipo', 'Início', 'Fim', 'Valor', 'Recebido em']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var last = sh.getLastRow(), rows = last >= 2 ? sh.getRange(2, 1, last - 1, 5).getDisplayValues() : [], idx = {}, mud = {}, novos = [], base = rows.length;
  rows.forEach(function (r, i) { idx[r[0] + '|' + r[1] + '|' + r[2] + '|' + r[3]] = i; });
  function put(uid, tipo, ini, fim, valor) {
    var k = uid + '|' + tipo + '|' + ini + '|' + fim;
    if (idx[k] != null) {
      if (String(rows[idx[k]][4]) === String(valor)) return;
      rows[idx[k]][4] = String(valor);
      if (idx[k] < base) mud[idx[k]] = 1; // linha nova do mesmo envio: só muda o valor antes de gravar
    }
    else { idx[k] = rows.length; rows.push([uid, tipo, ini, fim, String(valor)]); novos.push(rows.length - 1); }
  }
  return {
    upsert: function (uid, tipo, ini, fim, valor) { put(uid, tipo, String(ini), String(fim), valor); },
    upsertMin: function (uid, tipo, dia, valor) { // um por dia: guarda o menor
      var k = uid + '|' + tipo + '|' + dia + '|';
      if (idx[k] != null && Number(rows[idx[k]][4]) <= valor) return;
      put(uid, tipo, dia, '', valor);
    },
    flush: function () {
      var now = new Date();
      Object.keys(mud).forEach(function (i) { sh.getRange(+i + 2, 5, 1, 2).setValues([[rows[i][4], now]]); });
      if (novos.length) {
        var r0 = sh.getLastRow() + 1;
        sh.getRange(r0, 1, novos.length, 6).setNumberFormat('@');
        sh.getRange(r0, 1, novos.length, 6).setValues(novos.map(function (i) { return rows[i].concat([String(now.toISOString())]); }));
      }
      mud = {}; novos = []; base = rows.length;
    },
    totais: function (uid) {
      var t = {};
      rows.forEach(function (r) {
        if (r[0] !== uid) return;
        var v = Number(r[4]);
        if (r[1] === 'passos') { var d = diaSP_(iso_(r[2])); if (!d) return; t[d] = t[d] || {}; t[d].passos = (t[d].passos || 0) + v; }
        else if (r[1] === 'sono') { var d2 = diaSP_(iso_(r[3])); if (!d2) return; t[d2] = t[d2] || {}; t[d2].sono = (t[d2].sono || 0) + v; }
        else if (r[1] === 'fc_rep' || r[1] === 'fc_min') { t[r[2]] = t[r[2]] || {}; t[r[2]][r[1]] = v; }
      });
      return t;
    }
  };
}

// ---------- Utilitários ----------
function janela_(d) { var t = d.getTime(), agora = Date.now(); return t > agora - HC_DIAS * 86400000 && t < agora + 86400000; }
function arr_(x) { return Array.isArray(x) ? x : []; }
function iso_(s) {
  if (!s || !/^\d{4}-\d{2}-\d{2}T/.test(String(s))) return null;
  var d = new Date(String(s));
  return isNaN(d.getTime()) ? null : d;
}
function diaSP_(d) { return d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : ''; }
function nomeTipo_(t) {
  if (!t) return 'Atividade';
  var s = t.toLowerCase().replace(/_/g, ' ');
  return s.charAt(0).toUpperCase() + s.slice(1);
}
