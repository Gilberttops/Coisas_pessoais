/**
 * Integração Strava (treinos do relógio via Mi Fitness → Strava) e importação do histórico do Mi Fitness.
 * Usa as funções de Code.gs: SCHEMA, validate_, sheetFor_, auditLog_, props_, TZ, hoje_, parseNumCampo_.
 *
 * PRIMEIRO USO (depois da implantação do app):
 * 1. strava.com/settings/api → criar app. "Authorization Callback Domain" = script.google.com
 * 2. Menu Saúde > Strava: configurar chaves (Client ID e Client Secret)
 * 3. Menu Saúde > Strava: conectar conta → abrir o link → Autorizar
 * 4. Menu Saúde > Strava: ativar sincronização automática (de hora em hora)
 */
var STRAVA = { AUTH: 'https://www.strava.com/oauth/authorize', TOKEN: 'https://www.strava.com/oauth/token', API: 'https://www.strava.com/api/v3' };
var STRAVA_TIPOS = { Run: 'Corrida', TrailRun: 'Corrida', VirtualRun: 'Corrida', Walk: 'Caminhada', Hike: 'Trilha', Ride: 'Bicicleta',
  VirtualRide: 'Bicicleta', MountainBikeRide: 'Bicicleta', GravelRide: 'Bicicleta', EBikeRide: 'Bicicleta', WeightTraining: 'Musculação',
  Workout: 'Funcional', Crossfit: 'Funcional', HighIntensityIntervalTraining: 'Funcional', Soccer: 'Futebol', Swim: 'Natação',
  Yoga: 'Alongamento', Pilates: 'Alongamento', StairStepper: 'Escada', Elliptical: 'Elíptico', Rowing: 'Remo', Volleyball: 'Vôlei',
  Basketball: 'Basquete', Golf: 'Golfe', Tennis: 'Tênis' };
var DUP_SEGUNDOS = 180; // mesmo treino vindo de duas fontes: início com até 3 min de diferença

// ---------- Menu (getUi() primeiro: chamadas pelo app web falham antes de fazer qualquer coisa) ----------
function configurarStrava() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('Strava: Client ID', 'Em strava.com/settings/api', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK || !/^\d+$/.test(a.getResponseText().trim())) { ui.alert('Client ID inválido. Nada foi alterado.'); return; }
  var b = ui.prompt('Strava: Client Secret', 'Copie o "Client Secret" da mesma página.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK || !/^[a-f0-9]{20,64}$/i.test(b.getResponseText().trim())) { ui.alert('Client Secret inválido. Nada foi alterado.'); return; }
  var p = props_();
  p.setProperty('STRAVA_CLIENT_ID', a.getResponseText().trim());
  p.setProperty('STRAVA_CLIENT_SECRET', b.getResponseText().trim());
  ['STRAVA_REFRESH', 'STRAVA_ACCESS', 'STRAVA_EXP'].forEach(function (k) { p.deleteProperty(k); });
  ui.alert('Chaves salvas. Agora use "Conectar conta".');
}
function conectarStrava() {
  var ui = SpreadsheetApp.getUi(), id = props_().getProperty('STRAVA_CLIENT_ID');
  if (!id) { ui.alert('Configure as chaves primeiro.'); return; }
  var url = ScriptApp.getService().getUrl();
  if (!url) { ui.alert('Implante o app da Web primeiro (Implantar > Nova implantação).'); return; }
  var state = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('strava_state', state, 900);
  var link = STRAVA.AUTH + '?client_id=' + encodeURIComponent(id) + '&response_type=code&approval_prompt=force&scope=activity:read_all' +
    '&redirect_uri=' + encodeURIComponent(url) + '&state=' + state;
  ui.showModalDialog(HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Abra o link, autorize e volte aqui. O link vale 15 minutos.</p>' +
    '<p><a style="font:600 16px sans-serif" target="_blank" href="' + link + '">Autorizar no Strava</a></p>').setWidth(380).setHeight(140), 'Conectar Strava');
}
function ativarSyncStrava() {
  var ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncStravaAgendado') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncStravaAgendado').timeBased().everyHours(1).create();
  ui.alert('Sincronização automática ligada: de hora em hora.');
}
function sincronizarStravaMenu() {
  var ui = SpreadsheetApp.getUi(), r = sincronizarStrava_(true);
  ui.alert(r.ok ? 'Strava: ' + r.novos + ' treino(s) novo(s).' : 'Strava: ' + (r.message || r.error));
}
// Gatilho de hora em hora. Chamado de outro jeito (ex.: google.script.run), não faz nada:
// exige o triggerUid de um gatilho existente deste projeto.
function syncStravaAgendado(e) {
  var uid = e && e.triggerUid ? String(e.triggerUid) : '';
  if (!uid || !ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === uid; })) return;
  sincronizarStrava_(false);
}

// ---------- OAuth ----------
function stravaCallback_(e) {
  var p = e.parameter, cache = CacheService.getScriptCache(), st = cache.get('strava_state');
  function pagina(t) { return HtmlService.createHtmlOutput('<p style="font:16px sans-serif;padding:24px">' + t + '</p>').setTitle('Strava'); }
  if (!st || !p.state || !safeEquals_(p.state, st)) return pagina('Link expirado ou inválido. Gere um novo pelo menu da planilha.');
  cache.remove('strava_state');
  if (p.error || !p.code) return pagina('Autorização cancelada.');
  if (String(p.scope || '').indexOf('activity:read') < 0) return pagina('Faltou permitir a leitura das atividades. Tente de novo e deixe a opção marcada.');
  var r = tokenRequest_({ code: p.code, grant_type: 'authorization_code' });
  if (!r.ok) return pagina('Falha ao conectar: ' + r.message);
  var pr = props_();
  if (!pr.getProperty('STRAVA_AFTER')) pr.setProperty('STRAVA_AFTER', String(Math.floor(Date.now() / 1000) - 30 * 86400));
  return pagina('Strava conectado. Pode fechar esta aba.');
}
function tokenRequest_(extra) {
  var pr = props_(), payload = { client_id: pr.getProperty('STRAVA_CLIENT_ID'), client_secret: pr.getProperty('STRAVA_CLIENT_SECRET') };
  Object.keys(extra).forEach(function (k) { payload[k] = extra[k]; });
  var res = UrlFetchApp.fetch(STRAVA.TOKEN, { method: 'post', payload: payload, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { ok: false, message: 'Strava respondeu ' + res.getResponseCode() };
  var j = JSON.parse(res.getContentText());
  pr.setProperty('STRAVA_ACCESS', j.access_token);
  pr.setProperty('STRAVA_REFRESH', j.refresh_token);
  pr.setProperty('STRAVA_EXP', String(j.expires_at));
  return { ok: true };
}
function accessToken_() {
  var pr = props_();
  if (!pr.getProperty('STRAVA_REFRESH')) return null;
  if (parseInt(pr.getProperty('STRAVA_EXP') || '0', 10) - 120 > Date.now() / 1000) return pr.getProperty('STRAVA_ACCESS');
  var r = tokenRequest_({ grant_type: 'refresh_token', refresh_token: pr.getProperty('STRAVA_REFRESH') });
  return r.ok ? pr.getProperty('STRAVA_ACCESS') : null;
}
function stravaGet_(path, tok) {
  var res = UrlFetchApp.fetch(STRAVA.API + path, { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
  var c = res.getResponseCode();
  if (c === 429) throw new Error('Limite de consultas do Strava. Tenta de novo em 15 min.');
  if (c !== 200) throw new Error('Strava respondeu ' + c);
  return JSON.parse(res.getContentText());
}
function stravaStatus_() {
  var pr = props_();
  return { configurado: !!pr.getProperty('STRAVA_CLIENT_ID'), conectado: !!pr.getProperty('STRAVA_REFRESH'), ultima: pr.getProperty('STRAVA_LAST') || '' };
}

// ---------- Sincronização ----------
function sincronizarStrava_(forcar) {
  var cache = CacheService.getScriptCache();
  if (!forcar && cache.get('strava_busy')) return { ok: true, novos: 0 };
  cache.put('strava_busy', '1', 120);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'ocupado', message: 'Planilha ocupada, tente de novo.' }; }
  var pr = props_();
  try {
    var tok = accessToken_();
    if (!tok) return { ok: false, error: 'strava_desconectado', message: 'Strava não conectado.' };
    var cursor = parseInt(pr.getProperty('STRAVA_AFTER') || '0', 10);
    var desde = Math.max(0, cursor - 3 * 86400); // revê 3 dias: pega treino enviado com atraso; o ID barra repetidos
    var ss = SpreadsheetApp.getActiveSpreadsheet(), imp = importados_(ss), novos = [], maxIni = cursor, aviso = '';
    try {
      for (var page = 1; page <= 10; page++) {
        var lista = stravaGet_('/athlete/activities?per_page=50&page=' + page + '&after=' + desde, tok);
        for (var k = 0; k < lista.length; k++) {
          var a = lista[k], ini = Math.floor(Date.parse(a.start_date) / 1000), sid = 'strava:' + a.id;
          if (isNaN(ini)) continue;
          if (imp.ids[sid] || ehDuplicado_(imp.outros, ini)) { maxIni = Math.max(maxIni, ini); continue; }
          var kcal = a.calories;
          if (kcal == null) kcal = stravaGet_('/activities/' + encodeURIComponent(a.id), tok).calories; // erro interrompe: tenta na próxima
          var v = validate_(SCHEMA.atividade, {
            data: Utilities.formatDate(new Date(ini * 1000), TZ, 'yyyy-MM-dd'),
            tipo: STRAVA_TIPOS[a.sport_type] || STRAVA_TIPOS[a.type] || String(a.sport_type || a.type || 'Atividade'),
            minutos: Math.max(1, Math.round((a.elapsed_time || a.moving_time || 60) / 60)),
            distancia: a.distance ? Math.round(a.distance / 10) / 100 : '',
            passos: '', kcal: kcal ? Math.round(kcal) : '', fc: a.average_heartrate ? Math.round(a.average_heartrate) : '',
            intensidade: '', obs: 'Strava'
          }, true);
          maxIni = Math.max(maxIni, ini);
          if (!v.ok) continue;
          imp.ids[sid] = 1;
          novos.push({ ini: ini, sid: sid, values: v.values });
        }
        if (lista.length < 50) break;
      }
    } catch (e) {
      aviso = String(e && e.message ? e.message : e); // grava o que já veio; o resto entra na próxima rodada
    }
    gravarComRegistro_(ss, imp.sheet, novos.map(function (n) { return { ini: n.ini, origem: 'Strava', chave: n.sid, values: n.values }; }));
    // A API devolve do mais antigo para o mais novo quando há "after": o cursor só passa do que foi processado.
    pr.setProperty('STRAVA_AFTER', String(Math.max(cursor, maxIni)));
    pr.setProperty('STRAVA_LAST', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    if (aviso && !novos.length) return { ok: false, error: 'falha', message: aviso };
    return { ok: true, novos: novos.length, aviso: aviso };
  } catch (err) {
    return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Histórico do Mi Fitness ----------
// Antes: Arquivo > Importar > Fazer upload de mifit_atividades.csv e mifit_diario.csv,
// cada um com "Inserir nova(s) página(s)" e "Converter texto em números e datas" = Não.
function importarHistoricoMiFitness() {
  var ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActiveSpreadsheet();
  var shA = ss.getSheetByName('mifit_atividades'), shD = ss.getSheetByName('mifit_diario');
  if (!shA && !shD) { ui.alert('Importe antes os arquivos mifit_atividades.csv e mifit_diario.csv (Arquivo > Importar > Inserir nova página).'); return; }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { ui.alert('Planilha ocupada, tente de novo.'); return; }
  var rel = [];
  try {
    if (shA) {
      var imp = importados_(ss), chaves = imp.todos, lote = {}, novos = [], ruins = 0, dup = 0; // lote: dentro do arquivo, só início idêntico
      linhas_(shA).forEach(function (r) {
        var ini = parseInt(r.inicio, 10);
        if (!ini) { ruins++; return; }
        if (lote[ini] || ehDuplicado_(chaves, ini)) { dup++; return; }
        var dm = parseNumCampo_(r.distancia_m, true);
        var v = validate_(SCHEMA.atividade, { data: dataTxt_(r.data), tipo: r.tipo, minutos: r.minutos, distancia: dm > 0 ? Math.round(dm / 10) / 100 : '',
          passos: r.passos, kcal: r.kcal, fc: r.fc, intensidade: '', obs: r.obs || 'Mi Fitness' }, true);
        if (!v.ok) { ruins++; return; }
        lote[ini] = 1; novos.push({ ini: ini, values: v.values });
      });
      gravarComRegistro_(ss, imp.sheet, novos.map(function (n) { return { ini: n.ini, origem: 'Mi Fitness', chave: 'mifit:' + n.ini, values: n.values }; }));
      rel.push('Treinos: ' + novos.length + ' importados, ' + dup + ' já existiam, ' + ruins + ' recusados');
    }
    if (shD) {
      var shDia = sheetFor_(ss, SCHEMA.diario), existentes = {}, novosD = [], ruinsD = 0, dupD = 0;
      if (shDia.getLastRow() >= 2) shDia.getRange(2, 2, shDia.getLastRow() - 1, 1).getDisplayValues().forEach(function (x) { existentes[dataTxt_(x[0])] = 1; });
      linhas_(shD).forEach(function (r) {
        var d = dataTxt_(r.data);
        if (existentes[d]) { dupD++; return; }
        var sm = parseNumCampo_(r.sono_min, true);
        var v = validate_(SCHEMA.diario, { data: d, sono: sm > 0 ? Math.round(sm / 6) / 10 : '', agua_l: '', passos: r.passos, pa_sis: '', pa_dia: '', fc_rep: r.fc_rep, obs: 'Mi Fitness' }, true);
        if (!v.ok) { ruinsD++; return; }
        existentes[d] = 1; novosD.push(v.values);
      });
      gravarLote_(ss, SCHEMA.diario, novosD);
      rel.push('Diário: ' + novosD.length + ' dias importados, ' + dupD + ' já existiam, ' + ruinsD + ' recusados');
    }
    auditLog_('import', 'mifitness', '', null, { resumo: rel }, '');
  } finally {
    lock.releaseLock();
  }
  ui.alert(rel.join('\n') + '\n\nPode apagar as abas mifit_atividades e mifit_diario.');
}

// ---------- Utilitários ----------
function linhas_(sh) {
  if (sh.getLastRow() < 2) return [];
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues(), h = v[0].map(function (x) { return String(x).trim(); });
  return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = String(r[i]).trim(); }); return o; });
}
function dataTxt_(s) {
  s = String(s || '').trim();
  var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : s;
}
function importados_(ss) {
  var sh = ss.getSheetByName('Importados');
  if (!sh) {
    sh = ss.insertSheet('Importados');
    sh.getRange(1, 1, 1, 5).setValues([['Início (epoch)', 'Origem', 'ID', 'Quando', 'Chave']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var out = { sheet: sh, todos: [], outros: [], ids: {} };
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
    var ini = Number(r[0]);
    if (r[4]) out.ids[String(r[4])] = 1;
    if (ini > 0) { out.todos.push(ini); if (r[1] !== 'Strava') out.outros.push(ini); }
  });
  return out;
}
function ehDuplicado_(chaves, ini) {
  for (var i = 0; i < chaves.length; i++) if (Math.abs(chaves[i] - ini) <= DUP_SEGUNDOS) return true;
  return false;
}
// Grava os treinos e o registro em Importados. Se o registro falhar, desfaz os treinos (evita duplicar depois).
function gravarComRegistro_(ss, shImp, itens) {
  if (!itens.length) return [];
  var sh = sheetFor_(ss, SCHEMA.atividade), ids = gravarLote_(ss, SCHEMA.atividade, itens.map(function (x) { return x.values; }));
  try {
    var now = new Date();
    shImp.getRange(shImp.getLastRow() + 1, 1, itens.length, 5).setValues(itens.map(function (x, i) { return [x.ini, x.origem, ids[i], now, x.chave]; }));
  } catch (e) {
    sh.deleteRows(sh.getLastRow() - itens.length + 1, itens.length);
    throw e;
  }
  auditLog_('import', 'atividade', ids[0] + '..' + ids[ids.length - 1], null, { origem: itens[0].origem, qtd: itens.length }, '');
  return ids;
}
// Grava várias linhas de uma vez, com IDs sequenciais. Chamar dentro do lock.
function gravarLote_(ss, sc, lista) {
  if (!lista.length) return [];
  var sh = sheetFor_(ss, sc), p = props_(), k = 'SEQ_' + sc.prefixo, max = parseInt(p.getProperty(k) || '0', 10), last = sh.getLastRow(), now = new Date();
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { var m = String(r[0]).match(/(\d+)$/); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  var ids = [], rows = lista.map(function (v, i) {
    var id = sc.prefixo + (max + 1 + i); ids.push(id);
    return [id].concat(sc.campos.map(function (c) { return v[c[0]]; })).concat([now, now]);
  });
  p.setProperty(k, String(max + lista.length));
  sh.getRange(last + 1, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
  return ids;
}
