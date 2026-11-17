/**
 * Saúde – acompanhamento de peso, bioimpedância, atividades e diário.
 * Padrão apps-script-seguro: senha (hash+salt), sessão assinada, limite de tentativas,
 * LockService, idempotência por opId, auditoria, validação no servidor.
 *
 * PRIMEIRO USO
 * 1. Planilha nova > Extensões > Apps Script > colar este arquivo (Code.gs) e criar o HTML "Index".
 * 2. Rodar a função "configurarPlanilha" uma vez (cria as abas). Recarregar a planilha.
 * 3. Menu "Saúde > Definir senha de acesso".
 * 4. Implantar > Nova implantação > App da Web > Executar como "Eu" > Acesso "Qualquer pessoa".
 *    Abrir a URL /exec no celular e "Adicionar à tela inicial".
 */
var AUTH = { SESSION_HOURS: 12, REMEMBER_DAYS: 60, MAX_FAILS: 5, LOCK_MINUTES: 15, MIN_PASSWORD_LENGTH: 8 };
var TZ = 'America/Sao_Paulo';

// Campos por entidade: [chave, título, tipo, min, max, 'int' = inteiro (aceita 1.500 como mil)]
var SCHEMA = {
  peso: { aba: 'Peso', prefixo: 'P', campos: [
    ['data', 'Data', 'date'], ['peso', 'Peso (kg)', 'num', 20, 300], ['gordura', 'Gordura (%)', 'num', 1, 75],
    ['agua', 'Água (%)', 'num', 20, 85], ['musculo', 'Músculo (%)', 'num', 10, 75], ['osso', 'Massa óssea (kg)', 'num', 0.5, 10],
    ['kcal', 'Kcal basal/ativa', 'num', 500, 6000, 'int'], ['cintura', 'Cintura (cm)', 'num', 40, 200], ['obs', 'Obs', 'text']] },
  atividade: { aba: 'Atividades', prefixo: 'A', campos: [
    ['data', 'Data', 'date'], ['tipo', 'Tipo', 'text'], ['minutos', 'Minutos', 'num', 1, 1440, 'int'],
    ['distancia', 'Distância (km)', 'num', 0, 500], ['passos', 'Passos', 'num', 0, 150000, 'int'], ['kcal', 'Kcal', 'num', 0, 10000, 'int'],
    ['fc', 'FC média', 'num', 30, 230, 'int'], ['intensidade', 'Intensidade', 'text'], ['obs', 'Obs', 'text']] },
  diario: { aba: 'Diario', prefixo: 'D', campos: [
    ['data', 'Data', 'date'], ['sono', 'Sono (h)', 'num', 0, 24], ['agua_l', 'Água (L)', 'num', 0, 15],
    ['passos', 'Passos do dia', 'num', 0, 150000, 'int'], ['pa_sis', 'PA sistólica', 'num', 60, 260, 'int'], ['pa_dia', 'PA diastólica', 'num', 30, 160, 'int'],
    ['fc_rep', 'FC repouso', 'num', 25, 150, 'int'], ['obs', 'Obs', 'text']] }
};
var CONFIG_KEYS = { altura: [100, 230], meta: [30, 250], nascimento: null, sexo: null };

// ---------- Autenticação ----------
function props_() { return PropertiesService.getScriptProperties(); }
function secret_() {
  var p = props_(), s = p.getProperty('AUTH_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('AUTH_SECRET', s); }
  return s;
}
function toHex_(bytes) {
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}
function hashPassword_(pw, salt) {
  return toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
function safeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function sign_(t) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(t, secret_())); }
function issueToken_(ttlMs) { var exp = Date.now() + ttlMs; return exp + '.' + sign_(String(exp)); }
function validToken_(token) {
  if (!token) return false;
  var p = String(token).split('.');
  if (p.length !== 2) return false;
  var exp = parseInt(p[0], 10);
  if (isNaN(exp) || exp < Date.now()) return false;
  return safeEquals_(p[1], sign_(String(exp)));
}
function login_(password, remember) {
  var cache = CacheService.getScriptCache();
  var fails = parseInt(cache.get('login_fails') || '0', 10);
  if (fails >= AUTH.MAX_FAILS) return { ok: false, error: 'bloqueado', message: 'Muitas tentativas. Tente em ' + AUTH.LOCK_MINUTES + ' min.' };
  var p = props_(), hash = p.getProperty('AUTH_HASH'), salt = p.getProperty('AUTH_SALT');
  if (!hash || !salt) return { ok: false, error: 'sem_senha', message: 'Defina a senha pelo menu Saúde da planilha.' };
  if (!password || !safeEquals_(hashPassword_(String(password), salt), hash)) {
    cache.put('login_fails', String(fails + 1), AUTH.LOCK_MINUTES * 60);
    return { ok: false, error: 'senha_incorreta', message: 'Senha incorreta.' };
  }
  cache.remove('login_fails');
  var ttl = remember ? AUTH.REMEMBER_DAYS * 86400000 : AUTH.SESSION_HOURS * 3600000;
  return { ok: true, token: issueToken_(ttl), exp: Date.now() + ttl };
}

// ---------- Menu ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Saúde')
    .addItem('Definir senha de acesso', 'definirSenhaAcesso')
    .addItem('Encerrar todas as sessões', 'encerrarTodasSessoes')
    .addItem('Criar/conferir abas', 'configurarPlanilha')
    .addSeparator()
    .addItem('Importar histórico do Mi Fitness', 'importarHistoricoMiFitness')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Strava')
      .addItem('1. Configurar chaves', 'configurarStrava')
      .addItem('2. Conectar conta', 'conectarStrava')
      .addItem('3. Ativar sincronização automática', 'ativarSyncStrava')
      .addItem('Sincronizar agora', 'sincronizarStravaMenu'))
    .addToUi();
}
function definirSenhaAcesso() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Definir senha de acesso', 'Mínimo ' + AUTH.MIN_PASSWORD_LENGTH + ' caracteres.', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var pw = r1.getResponseText();
  if (!pw || pw.length < AUTH.MIN_PASSWORD_LENGTH) { ui.alert('Senha muito curta.'); return; }
  var r2 = ui.prompt('Confirmar senha', 'Digite de novo:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK || r2.getResponseText() !== pw) { ui.alert('Senhas diferentes. Nada foi alterado.'); return; }
  var salt = Utilities.getUuid(), p = props_();
  p.setProperty('AUTH_SALT', salt);
  p.setProperty('AUTH_HASH', hashPassword_(pw, salt));
  p.setProperty('AUTH_SECRET', Utilities.getUuid() + Utilities.getUuid());
  CacheService.getScriptCache().remove('login_fails');
  ui.alert('Senha definida. Sessões anteriores encerradas.');
}
// Funções de menu: getUi() falha fora da planilha, então chamadas pelo app web não fazem nada.
function encerrarTodasSessoes() {
  var ui = SpreadsheetApp.getUi();
  props_().setProperty('AUTH_SECRET', Utilities.getUuid() + Utilities.getUuid());
  ui.alert('Sessões encerradas.');
}
function configurarPlanilha() {
  SpreadsheetApp.getUi();
  criarAbas_();
}
function criarAbas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (k) { sheetFor_(ss, SCHEMA[k]); });
  var cfg = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config');
    cfg.getRange(1, 1, 1, 2).setValues([['Chave', 'Valor']]).setFontWeight('bold');
    cfg.setFrozenRows(1);
  }
}
function sheetFor_(ss, sc) {
  var sh = ss.getSheetByName(sc.aba);
  if (!sh) {
    sh = ss.insertSheet(sc.aba);
    var head = ['ID'].concat(sc.campos.map(function (c) { return c[1]; })).concat(['Criado em', 'Alterado em']);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('@'); // data como texto aaaa-mm-dd (sem fuso)
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---------- Web app ----------
// Chamado pelo site no GitHub Pages: fetch com Content-Type text/plain (sem preflight de CORS).
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'requisicao_invalida' }); }
  return json_(api(body));
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  if (e && e.parameter && e.parameter.state) return stravaCallback_(e); // retorno da autorização do Strava
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Saúde')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// Único ponto de entrada chamado pela tela via google.script.run.api(body)
function api(body) {
  try {
    body = body || {};
    if (body.action === 'login') return login_(body.password, !!body.remember);
    if (!validToken_(body.t)) return { ok: false, error: 'unauthorized' };
    if (body.action === 'load') return load_();
    if (body.action === 'syncStrava') return sincronizarStrava_(false);
    if (body.action !== 'op') return { ok: false, error: 'acao_desconhecida' };
    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'ocupado', message: 'Planilha ocupada, tente de novo.' }; }
    try { return handleOp_(body); } finally { lock.releaseLock(); }
  } catch (err) {
    return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) };
  }
}

function load_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = { ok: true, config: readConfig_(ss), strava: stravaStatus_() };
  Object.keys(SCHEMA).forEach(function (k) { out[k] = readRows_(ss, SCHEMA[k]); });
  return out;
}
function readRows_(ss, sc) {
  var sh = sheetFor_(ss, sc), last = sh.getLastRow();
  if (last < 2) return [];
  var n = sc.campos.length + 1;
  return sh.getRange(2, 1, last - 1, n).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      var o = { id: String(r[0]) };
      sc.campos.forEach(function (c, i) {
        var v = r[i + 1];
        if (c[2] === 'date') v = v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
        else if (c[2] === 'num') v = v === '' ? null : parseNumBR(v);
        else v = String(v);
        o[c[0]] = (typeof v === 'number' && isNaN(v)) ? null : v;
      });
      return o;
    });
}
function readConfig_(ss) {
  var sh = ss.getSheetByName('Config'), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) o[String(r[0])] = r[1]; });
  return o;
}

// ---------- Gravação ----------
function handleOp_(body) {
  if (!/^[a-z0-9]{8,40}$/i.test(String(body.opId || ''))) return { ok: false, error: 'opid_invalido' };
  var cache = CacheService.getScriptCache(), key = 'op_' + body.opId;
  var done = cache.get(key);
  if (done) { var r0 = JSON.parse(done); r0.repeated = true; return r0; }
  var prev = findOpInAudit_(body.opId); // cache pode ter expirado: confere na auditoria
  if (prev) return { ok: true, id: prev, repeated: true };
  var result;
  try { result = runOp_(body); } catch (err) { return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) }; }
  if (result && result.ok) cache.put(key, JSON.stringify(result), 21600);
  return result;
}

function runOp_(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (body.entity === 'config') return saveConfig_(ss, body.values || {}, body.opId);
  var sc = SCHEMA[body.entity];
  if (!sc) return { ok: false, error: 'entidade_invalida' };
  var sh = sheetFor_(ss, sc), now = new Date();

  if (body.kind === 'create') {
    var v = validate_(sc, body.values || {}, true);
    if (!v.ok) return v;
    var id = nextId_(sh, sc.prefixo);
    var row = [id].concat(sc.campos.map(function (c) { return v.values[c[0]]; })).concat([now, now]);
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    auditLog_('create', body.entity, id, null, v.values, body.opId);
    return { ok: true, id: id };
  }

  var rowNum = findRow_(sh, body.id);
  if (!rowNum) return { ok: false, error: 'nao_encontrado', message: 'Registro não encontrado.' };
  var n = sc.campos.length + 1;
  var before = {};
  var cur = sh.getRange(rowNum, 1, 1, n).getValues()[0];
  sc.campos.forEach(function (c, i) { before[c[0]] = cur[i + 1] instanceof Date ? Utilities.formatDate(cur[i + 1], TZ, 'yyyy-MM-dd') : cur[i + 1]; });

  if (body.kind === 'update') {
    var merged = {};
    Object.keys(before).forEach(function (k) { merged[k] = before[k]; });
    Object.keys(body.values || {}).forEach(function (k) { merged[k] = body.values[k]; });
    var u = validate_(sc, merged, true);
    if (!u.ok) return u;
    sh.getRange(rowNum, 2, 1, sc.campos.length).setValues([sc.campos.map(function (c) { return u.values[c[0]]; })]);
    sh.getRange(rowNum, n + 2).setValue(now);
    auditLog_('update', body.entity, body.id, before, u.values, body.opId);
    return { ok: true, id: body.id };
  }
  if (body.kind === 'remove') {
    sh.deleteRow(rowNum);
    auditLog_('remove', body.entity, body.id, before, null, body.opId);
    return { ok: true, id: body.id };
  }
  return { ok: false, error: 'operacao_invalida' };
}

function validate_(sc, src, requireDate) {
  var out = {}, erros = [];
  sc.campos.forEach(function (c) {
    var k = c[0], raw = src[k];
    if (c[2] === 'date') {
      var s = String(raw == null ? '' : raw);
      if (!validDate_(s)) { if (requireDate) erros.push('Data inválida'); out[k] = ''; return; }
      if (s > hoje_()) erros.push('Data no futuro');
      out[k] = s;
    } else if (c[2] === 'num') {
      if (raw === '' || raw == null) { out[k] = ''; return; }
      var x = parseNumCampo_(raw, c[5] === 'int');
      if (isNaN(x)) { erros.push(c[1] + ': número inválido'); return; }
      if (x < c[3] || x > c[4]) { erros.push(c[1] + ' fora da faixa (' + c[3] + ' a ' + c[4] + ')'); return; }
      out[k] = Math.round(x * 100) / 100;
    } else {
      out[k] = sanitizeText_(raw);
    }
  });
  if (sc.aba === 'Peso' && out.peso === '') erros.push('Informe o peso');
  if (sc.aba === 'Atividades' && !out.tipo) erros.push('Informe o tipo de atividade');
  if (sc.aba === 'Atividades' && out.minutos === '') erros.push('Informe os minutos');
  if (sc.aba === 'Diario' && out.pa_sis !== '' && out.pa_dia !== '' && out.pa_dia >= out.pa_sis) erros.push('PA diastólica deve ser menor que a sistólica');
  if (erros.length) return { ok: false, error: 'validacao', message: erros.join('; ') };
  return { ok: true, values: out };
}
// Evita injeção de fórmula na planilha (texto começando com = + - @)
function sanitizeText_(v) {
  var s = String(v == null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function saveConfig_(ss, values, opId) {
  var sh = ss.getSheetByName('Config');
  if (!sh) { criarAbas_(); sh = ss.getSheetByName('Config'); }
  var before = readConfig_(ss), erros = [], ok = {};
  Object.keys(values).forEach(function (k) {
    if (!(k in CONFIG_KEYS)) return;
    var v = values[k];
    if (CONFIG_KEYS[k]) {
      if (v === '' || v == null) v = '';
      else {
        v = parseNumBR(v);
        if (isNaN(v) || v < CONFIG_KEYS[k][0] || v > CONFIG_KEYS[k][1]) { erros.push(k + ' fora da faixa'); return; }
      }
    } else if (k === 'nascimento') {
      v = String(v || '');
      if (v && !validDate_(v)) { erros.push('Nascimento inválido'); return; }
    } else v = (v === 'M' || v === 'F') ? v : '';
    ok[k] = v;
  });
  if (erros.length) return { ok: false, error: 'validacao', message: erros.join('; ') };
  Object.keys(ok).forEach(function (k) {
    var r = findRow_(sh, k);
    if (r) sh.getRange(r, 2).setValue(ok[k]); else sh.appendRow([k, ok[k]]);
  });
  auditLog_('update', 'config', 'config', before, values, opId);
  return { ok: true };
}

// ---------- Utilitários ----------
function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
// Chamar SEMPRE dentro do lock. Guarda o maior número já usado para não reaproveitar ID apagado.
function nextId_(sheet, prefix) {
  var p = props_(), k = 'SEQ_' + prefix, max = parseInt(p.getProperty(k) || '0', 10), last = sheet.getLastRow();
  if (last >= 2) sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var m = String(r[0]).match(/(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  p.setProperty(k, String(max + 1));
  return prefix + (max + 1);
}
function findOpInAudit_(opId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Auditoria');
  if (!sh || sh.getLastRow() < 2) return null;
  var from = Math.max(2, sh.getLastRow() - 2000), vals = sh.getRange(from, 4, sh.getLastRow() - from + 1, 4).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][3]) === opId) return String(vals[i][0]);
  return null;
}
function validDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s));
  if (!m) return false;
  var d = new Date(+m[1], +m[2] - 1, +m[3], 12);
  return d.getFullYear() === +m[1] && d.getMonth() === +m[2] - 1 && d.getDate() === +m[3] && +m[1] >= 1900;
}
function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function auditLog_(action, entity, id, before, after, opId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Auditoria');
    if (!sh) {
      sh = ss.insertSheet('Auditoria');
      sh.getRange(1, 1, 1, 7).setValues([['Quando', 'Ação', 'Entidade', 'ID', 'Antes (JSON)', 'Depois (JSON)', 'OpId']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), action, entity, id, JSON.stringify(before || null), JSON.stringify(after || null), opId || '']);
  } catch (e) { /* auditoria nunca derruba a gravação */ }
}
// Entrada digitada no app: só formatos sem ambiguidade. Recusa '1,2,3', '1.2.3' etc.
function parseNumCampo_(raw, inteiro) {
  if (typeof raw === 'number') return raw;
  var s = String(raw).replace(/\s/g, '');
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  if (inteiro && /^\d{1,3}(\.\d{3})+$/.test(s)) return parseInt(s.replace(/\./g, ''), 10); // 1.500 passos
  if (/^\d+[.,]\d+$/.test(s)) return parseFloat(s.replace(',', '.'));                     // 82,4 ou 82.4
  if (/^\d{1,3}(\.\d{3})+,\d+$/.test(s)) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
  return NaN;
}
function parseNumBR(input) {
  if (typeof input === 'number') return input;
  var s = String(input == null ? '' : input).replace(/[R$\s ]/g, '');
  if (s === '') return NaN;
  var neg = false;
  if (s[0] === '-' || (s[0] === '(' && s[s.length - 1] === ')')) { neg = true; s = s.replace(/[-()]/g, ''); }
  if (!/^[0-9.,]+$/.test(s)) return NaN;
  var lastDot = s.lastIndexOf('.'), lastComma = s.lastIndexOf(',');
  if (lastDot !== -1 && lastComma !== -1) {
    if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (lastComma !== -1) {
    if ((s.match(/,/g) || []).length > 1) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (lastDot !== -1) {
    if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  }
  var n = parseFloat(s);
  return isNaN(n) ? NaN : (neg ? -n : n);
}
