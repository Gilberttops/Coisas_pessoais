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
  DANCING: 'Dança', OTHER_WORKOUT: 'Musculação' }; // no relógio do Gilberto, treino livre é academia
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
  var nFases = sono.reduce(function (a, r) { return a + arr_(r.stages).length; }, 0);
  if (exs.length + passos.length + sono.length + fcr.length + fc.length + nFases > HC_MAX_ITENS) return { ok: false, error: 'grande_demais' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return { ok: false, error: 'ocupado' }; } // o app tenta de novo
  try {
    if (!hcDono_(k)) return { ok: false, error: 'unauthorized' }; // revogada enquanto esperava a trava
    var ss = SpreadsheetApp.getActiveSpreadsheet(), res = { ok: true, treinos: 0, dias: 0 };
    res.treinos = hcTreinos_(ss, u.id, exs);
    var bruto = hcBruto_(ss), datas = {}, sonoDias = {};
    passos.forEach(function (r) {
      var i = iso_(r.start_time), f = iso_(r.end_time), n = Number(r.count);
      if (!i || !f || f < i || !janela_(i) || !(n >= 0) || n > 100000) return;
      bruto.upsert(u.id, 'passos', i.toISOString(), f.toISOString(), Math.round(n)); datas[diaSP_(i)] = 1;
    });
    sono.forEach(function (r) {
      var f = iso_(r.session_end_time || r.end_time), s = Number(r.duration_seconds);
      if (!f || !janela_(f) || !(s > 0) || s > 86400) return;
      // a sessão é identificada pelo fim (o início não vem): duração corrigida substitui, não soma
      bruto.upsert(u.id, 'sono', '', f.toISOString(), Math.round(s));
      arr_(r.stages).forEach(function (g) { // fases: profundo, leve, REM, acordado
        var fi = iso_(g.start_time), ff = iso_(g.end_time), tp = faseSono_(g.stage);
        if (fi && ff && ff > fi && tp && ff - fi <= 86400000) bruto.upsert(u.id, 'fase_' + tp, fi.toISOString(), ff.toISOString(), Math.round((ff - fi) / 1000));
      });
      // a noite pode virar outro dia ao juntar versões: marca início, fim e o dia seguinte
      [f.getTime() - s * 1000, f.getTime(), f.getTime() + 43200000].forEach(function (x) { var d = diaSP_(new Date(x)); datas[d] = 1; sonoDias[d] = 1; });
    });
    // FC de repouso: só a de repouso. O heart_rate do Mi Fitness vem quase só dos treinos,
    // então o "menor do dia" dava 104-107 bpm; ele fica para FC média/máx no futuro.
    fcr.forEach(function (r) {
      var t = iso_(r.time), b = Number(r.bpm);
      if (!t || !janela_(t) || !(b >= 25 && b <= 150)) return;
      var d = diaSP_(t);
      bruto.upsertMin(u.id, 'fc_rep', d, Math.round(b)); datas[d] = 1;
    });
    fc.forEach(function (r) { // pontos de FC (o Mi Fitness grava quase só nos treinos): servem para FC dos treinos
      var t = iso_(r.time), b = Number(r.bpm);
      if (t && janela_(t) && b >= 30 && b <= 230) bruto.upsert(u.id, 'fc', t.toISOString(), '', Math.round(b));
    });
    bruto.flush();
    res.dias = hcAtualizaDiario_(ss, u.id, Object.keys(datas), bruto, sonoDias);
    if (Object.keys(sonoDias).length) hcSonoFases_(ss, u.id, Object.keys(sonoDias), bruto);
    res.fc_treinos = hcFcTreinos_(ss, u.id, bruto);
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
function hcAtualizaDiario_(ss, uid, datas, bruto, sonoDias) {
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
    if (t.fc_rep != null) val.fc_rep = t.fc_rep;
    var row = linhas[d];
    // dia que recebeu sono antes, mas a noite agora conta para outro dia: limpa (só linha do relógio)
    if (row && sonoDias && sonoDias[d] && val.sono == null && /^Relógio/.test(String(sh.getRange(row, pos.obs).getValue()))) {
      sh.getRange(row, pos.sono).setValue(''); feitos++;
    }
    if (!Object.keys(val).length) return;
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
      var t = {}, sonos = [];
      rows.forEach(function (r) {
        if (r[0] !== uid) return;
        var v = Number(r[4]);
        if (r[1] === 'passos') { var d = diaSP_(iso_(r[2])); if (!d) return; t[d] = t[d] || {}; t[d].passos = (t[d].passos || 0) + v; }
        else if (r[1] === 'sono') { var f = iso_(r[3]); if (f && v > 0) sonos.push([f.getTime() - v * 1000, f.getTime()]); }
        else if (r[1] === 'fc_rep') { t[r[2]] = t[r[2]] || {}; t[r[2]].fc_rep = v; }
      });
      // O Mi Fitness grava várias versões da mesma noite (mesmo início, fim crescente).
      // Junta os intervalos que se sobrepõem e soma só o tempo coberto, no dia em que acordou.
      uniao_(sonos).forEach(function (b) {
        var d = diaSP_(new Date(b[1])); t[d] = t[d] || {};
        t[d].sono = (t[d].sono || 0) + (b[1] - b[0]) / 1000;
      });
      return t;
    },
    linhas: function (uid) { return rows.filter(function (r) { return r[0] === uid; }); }
  };
}

// ---------- Fases do sono (aba Sono) ----------
// Sono: Usuário | Data (acordou) | Dormiu | Acordou | Total (h) | Profundo (h) | Leve (h) | REM (h) | Acordado (h) | Alterado em
function hcSonoFases_(ss, uid, dias, bruto) {
  var quer = {}; dias.forEach(function (d) { quer[d] = 1; });
  var sonos = [], fases = {};
  bruto.linhas(uid).forEach(function (r) {
    var v = Number(r[4]), f = iso_(r[3]);
    if (r[1] === 'sono' && f && v > 0) sonos.push([f.getTime() - v * 1000, f.getTime()]);
    else if (/^fase_/.test(r[1])) { var i = iso_(r[2]); if (i && f) (fases[r[1].slice(5)] = fases[r[1].slice(5)] || []).push([i.getTime(), f.getTime()]); }
  });
  Object.keys(fases).forEach(function (k) { fases[k] = uniao_(fases[k]); });
  var sh = ss.getSheetByName('Sono');
  if (!sh) {
    sh = ss.insertSheet('Sono');
    sh.getRange(1, 1, 1, 10).setValues([['Usuário', 'Data', 'Dormiu', 'Acordou', 'Total (h)', 'Profundo (h)', 'Leve (h)', 'REM (h)', 'Acordado (h)', 'Alterado em']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var idx = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(function (r, i) { idx[r[0] + '|' + r[1]] = i + 2; });
  var porDia = {};
  uniao_(sonos).forEach(function (b) { // a noite principal do dia é o bloco mais longo
    var d = diaSP_(new Date(b[1]));
    if (quer[d] && (!porDia[d] || b[1] - b[0] > porDia[d][1] - porDia[d][0])) porDia[d] = b;
  });
  dias.forEach(function (d) { // a noite passou a contar para outro dia: apaga a linha velha
    var row = idx[uid + '|' + d];
    if (row && !porDia[d]) sh.getRange(row, 3, 1, 8).setValues([['', '', '', '', '', '', '', new Date()]]);
  });
  var h = function (ms) { return Math.round(ms / 360000) / 10; }, hm = function (t) { return Utilities.formatDate(new Date(t), TZ, 'HH:mm'); };
  Object.keys(porDia).forEach(function (d) {
    var b = porDia[d], tem = false, lin = [uid, d, hm(b[0]), hm(b[1]), h(b[1] - b[0])];
    // versões repetidas da noite fazem as fases se sobreporem: se a soma passar da noite, reduz na proporção
    var ks = ['profundo', 'leve', 'rem', 'acordado'], ms = ks.map(function (k) { return sobrepoe_(fases[k] || [], b); });
    var soma = ms.reduce(function (a, x) { return a + x; }, 0), fat = soma > b[1] - b[0] ? (b[1] - b[0]) / soma : 1;
    ms.forEach(function (x) { if (x) tem = true; lin.push(x ? h(x * fat) : ''); });
    lin.push(new Date());
    var row = idx[uid + '|' + d];
    if (!row) { row = sh.getLastRow() + 1; idx[uid + '|' + d] = row; }
    sh.getRange(row, 1, 1, 4).setNumberFormat('@');
    sh.getRange(row, 1, 1, 10).setValues([lin]);
  });
}
function faseSono_(x) {
  var s = String(x == null ? '' : x).toLowerCase();
  if (s === '5' || /deep/.test(s)) return 'profundo';
  if (s === '4' || /light/.test(s)) return 'leve';
  if (s === '6' || /rem/.test(s)) return 'rem';
  if (s === '1' || s === '7' || /awake/.test(s)) return 'acordado';
  return '';
}
function uniao_(lista) { // junta intervalos [ini, fim] sobrepostos
  var out = [];
  lista.slice().sort(function (a, b) { return a[0] - b[0]; }).forEach(function (s) {
    var b = out[out.length - 1];
    if (b && s[0] <= b[1]) b[1] = Math.max(b[1], s[1]); else out.push([s[0], s[1]]);
  });
  return out;
}
function sobrepoe_(ints, b) { // ms de ints dentro do bloco b
  var t = 0; ints.forEach(function (x) { var i = Math.max(x[0], b[0]), f = Math.min(x[1], b[1]); if (f > i) t += f - i; }); return t;
}

// ---------- FC média e máxima dos treinos ----------
// Usa os pontos de FC guardados no HC_bruto que caem dentro do treino (últimos N dias). FC média + "FC máx N" na Obs.
function hcFcTreinos_(ss, uid, bruto, dias) {
  var pts = [];
  bruto.linhas(uid).forEach(function (r) { if (r[1] === 'fc') { var t = iso_(r[2]); if (t) pts.push([t.getTime(), Number(r[4])]); } });
  if (!pts.length) return 0;
  var imp = ss.getSheetByName('Importados'); if (!imp || imp.getLastRow() < 2) return 0;
  var desde = Date.now() / 1000 - (dias || 3) * 86400, treinos = [];
  imp.getRange(2, 1, imp.getLastRow() - 1, 6).getValues().forEach(function (r) {
    if (String(r[5]) === uid && /^hc:/.test(String(r[4])) && Number(r[0]) > desde) treinos.push({ ini: Number(r[0]) * 1000, id: String(r[2]) });
  });
  if (!treinos.length) return 0;
  var sc = SCHEMA.atividade, sh = sheetFor_(ss, sc), pos = {};
  sc.campos.forEach(function (c, i) { pos[c[0]] = COL.CAMPOS + i; });
  var linha = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r, i) { if (String(r[1]) === uid) linha[String(r[0])] = i + 2; });
  var n = 0;
  treinos.forEach(function (t) {
    var row = linha[t.id]; if (!row) return;
    var fim = t.ini + Number(sh.getRange(row, pos.minutos).getValue() || 0) * 60000;
    var bs = pts.filter(function (p) { return p[0] >= t.ini && p[0] <= fim; }).map(function (p) { return p[1]; });
    if (bs.length < 2) return;
    var media = Math.round(bs.reduce(function (a, b) { return a + b; }, 0) / bs.length), max = Math.max.apply(null, bs);
    var cObs = sh.getRange(row, pos.obs), obs = String(cObs.getValue() || ''), cFc = sh.getRange(row, pos.fc);
    // só mexe na FC se ela veio do relógio (vazia ou já marcada com "FC máx"); recalcula a cada envio
    if (cFc.getValue() !== '' && cFc.getValue() != null && !/FC máx/.test(obs)) return;
    cFc.setValue(media);
    cObs.setValue(/FC máx \d+/.test(obs) ? obs.replace(/FC máx \d+/, 'FC máx ' + max) : (obs ? obs + ' · ' : '') + 'FC máx ' + max);
    n++;
  });
  return n;
}

// Rodar UMA vez pelo editor depois da correção do sono/FC (17/11/2026): refaz o Diário com o HC_bruto
// e apaga a FC de repouso que tinha vindo do menor batimento dos treinos.
function hcRecalcularDiario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), bruto = hcBruto_(ss), sh = ss.getSheetByName('HC_bruto');
  var rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues() : [];
  var porUser = {}, fcMin = {}, sonoDia = {};
  rows.forEach(function (r) {
    porUser[r[0]] = porUser[r[0]] || {};
    if (r[1] === 'passos') porUser[r[0]][diaSP_(iso_(r[2]))] = 1;
    if (r[1] === 'sono') { porUser[r[0]][diaSP_(iso_(r[3]))] = 1; sonoDia[r[0] + '|' + diaSP_(iso_(r[3]))] = 1; }
    if (r[1] === 'fc_min') fcMin[r[0] + '|' + r[2]] = String(r[4]);
  });
  var sc = SCHEMA.diario, dsh = sheetFor_(ss, sc), col = COL.CAMPOS;
  sc.campos.forEach(function (c, i) { if (c[0] === 'fc_rep') col = COL.CAMPOS + i; });
  var colObs = COL.CAMPOS;
  sc.campos.forEach(function (c, i) { if (c[0] === 'obs') colObs = COL.CAMPOS + i; });
  if (dsh.getLastRow() >= 2) dsh.getRange(2, 1, dsh.getLastRow() - 1, col).getDisplayValues().forEach(function (r, i) {
    var v = fcMin[r[1] + '|' + dataTxt_(r[2])];
    if (v != null && String(r[col - 1]) === v && /^Relógio/.test(String(dsh.getRange(i + 2, colObs).getValue()))) dsh.getRange(i + 2, col).setValue('');
  });
  var colSono = COL.CAMPOS;
  sc.campos.forEach(function (c, i) { if (c[0] === 'sono') colSono = COL.CAMPOS + i; });
  Object.keys(porUser).forEach(function (uid) {
    var datas = Object.keys(porUser[uid]).filter(Boolean), tot = bruto.totais(uid);
    // dia que só tinha um pedaço da noite seguinte (ex.: 10/11 com 1,3 h): limpa o sono
    if (dsh.getLastRow() >= 2) dsh.getRange(2, 1, dsh.getLastRow() - 1, 3).getDisplayValues().forEach(function (r, i) {
      var d = dataTxt_(r[2]);
      if (r[1] === uid && sonoDia[uid + '|' + d] && !(tot[d] && tot[d].sono != null)) dsh.getRange(i + 2, colSono).setValue('');
    });
    Logger.log(uid + ': ' + hcAtualizaDiario_(ss, uid, datas, bruto) + ' dias');
    hcSonoFases_(ss, uid, datas, bruto);
    hcFcTreinos_(ss, uid, bruto, HC_DIAS);
  });
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
