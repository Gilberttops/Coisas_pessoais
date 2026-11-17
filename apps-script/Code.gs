/**
 * Saúde – peso, bioimpedância, treinos, diário e medicação (tirzepatida/semaglutida), com vários usuários.
 * Padrão apps-script-seguro: senha por usuário (hash+salt), sessão assinada com usuário e geração,
 * limite de tentativas, LockService, idempotência por opId, auditoria, validação no servidor.
 * Cada registro tem a coluna "Usuário"; leitura e gravação só enxergam as linhas do usuário da sessão.
 *
 * PRIMEIRO USO
 * 1. Planilha nova > Extensões > Apps Script > colar Code.gs, Strava.gs (Script) e Index (HTML).
 * 2. Recarregar a planilha > menu Saúde > Criar administrador (você).
 * 3. Implantar > Nova implantação > App da Web > Executar como "Eu" > Acesso "Qualquer pessoa".
 * 4. Outros usuários: no app, Ajustes > Usuários (só administrador).
 */
var AUTH = { SESSION_HOURS: 12, REMEMBER_DAYS: 60, MAX_FAILS: 5, MAX_FAILS_GLOBAL: 25, LOCK_MINUTES: 15, MIN_PASSWORD_LENGTH: 8 };
var TZ = 'America/Sao_Paulo';
var REMEDIOS = ['Tirzepatida', 'Semaglutida'];
var LOCAIS = ['', 'Abdômen', 'Coxa', 'Braço'];

// Campos por entidade: [chave, título, tipo, min, max, 'int' = inteiro (aceita 1.500 como mil)]
// Colunas na planilha: ID | Usuário | campos... | Criado em | Alterado em
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
    ['fc_rep', 'FC repouso', 'num', 25, 150, 'int'], ['obs', 'Obs', 'text']] },
  medicacao: { aba: 'Medicacao', prefixo: 'M', campos: [
    ['data', 'Data', 'date'], ['remedio', 'Remédio', 'text'], ['dose', 'Dose (mg)', 'num', 0.05, 20],
    ['local', 'Local', 'text'], ['obs', 'Obs', 'text']] }
};
var CONFIG_KEYS = { altura: [100, 230], meta: [30, 250], nascimento: null, sexo: null };
var COL = { ID: 1, UID: 2, CAMPOS: 3 }; // a data é sempre o primeiro campo (coluna 3)

// ---------- Usuários ----------
// Aba Usuarios: ID | Nome | Login | Hash | Salt | Admin | Ativo | Geração | Criado em
function props_() { return PropertiesService.getScriptProperties(); }
function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Usuarios');
  if (!sh) {
    sh = ss.insertSheet('Usuarios');
    sh.getRange(1, 1, 1, 9).setValues([['ID', 'Nome', 'Login', 'Hash', 'Salt', 'Admin', 'Ativo', 'Geração', 'Criado em']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function readUsers_() {
  var sh = usersSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().map(function (r, i) { r.row = i + 2; return r; }).filter(function (r) { return r[0]; }).map(function (r) {
    return { row: r.row, id: String(r[0]), nome: String(r[1]), login: String(r[2]), hash: String(r[3]), salt: String(r[4]),
      admin: r[5] === 'S', ativo: r[6] === 'S', gen: String(r[7] || '1') };
  });
}
function userById_(id) { return readUsers_().filter(function (u) { return u.id === String(id); })[0] || null; }
function normLogin_(s) {
  return String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}
function validLogin_(s) { return /^[a-z0-9._-]{3,30}$/.test(s); }
function criarUsuario_(nome, login, senha, admin) {
  nome = String(nome || '').trim().slice(0, 40); login = normLogin_(login);
  if (!nome) return { ok: false, error: 'validacao', message: 'Informe o nome.' };
  if (!validLogin_(login)) return { ok: false, error: 'validacao', message: 'Usuário: 3 a 30 letras, números, ponto, hífen ou sublinhado, sem espaço.' };
  if (!senha || String(senha).length < AUTH.MIN_PASSWORD_LENGTH) return { ok: false, error: 'validacao', message: 'Senha com no mínimo ' + AUTH.MIN_PASSWORD_LENGTH + ' caracteres.' };
  var users = readUsers_();
  if (users.some(function (u) { return u.login === login; })) return { ok: false, error: 'validacao', message: 'Esse usuário já existe.' };
  var sh = usersSheet_(), max = 0;
  users.forEach(function (u) { var m = u.id.match(/(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  var id = 'U' + (max + 1), salt = Utilities.getUuid();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 9).setValues([[id, sanitizeText_(nome), login, hashPassword_(String(senha), salt), salt, admin ? 'S' : 'N', 'S', '1', new Date()]]);
  auditLog_('create', 'usuario', id, null, { nome: nome, login: login, admin: !!admin }, '', '');
  return { ok: true, id: id };
}
function setUserFields_(u, f) { // f: {hash,salt,admin,ativo,bumpGen}
  var sh = usersSheet_();
  if (f.hash) sh.getRange(u.row, 4, 1, 2).setValues([[f.hash, f.salt]]);
  if (f.admin != null) sh.getRange(u.row, 6).setValue(f.admin ? 'S' : 'N');
  if (f.ativo != null) sh.getRange(u.row, 7).setValue(f.ativo ? 'S' : 'N');
  if (f.bumpGen) sh.getRange(u.row, 8).setValue(String(parseInt(u.gen, 10) + 1));
}

// ---------- Autenticação ----------
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
// token = exp.uid.gen.assinatura  (trocar a senha do usuário aumenta a geração e derruba as sessões dele)
function issueToken_(u, ttlMs) { var exp = Date.now() + ttlMs, base = exp + '.' + u.id + '.' + u.gen; return base + '.' + sign_(base); }
function sessionUser_(token) {
  var p = String(token || '').split('.');
  if (p.length !== 4) return null;
  var exp = parseInt(p[0], 10);
  if (isNaN(exp) || exp < Date.now()) return null;
  if (!safeEquals_(p[3], sign_(p[0] + '.' + p[1] + '.' + p[2]))) return null;
  var u = userById_(p[1]);
  if (!u || !u.ativo || u.gen !== p[2]) return null;
  return u;
}
function login_(login, password, remember) {
  var cache = CacheService.getScriptCache(), key = 'login_fails_' + normLogin_(login).slice(0, 40);
  var fails = parseInt(cache.get(key) || '0', 10), global = parseInt(cache.get('login_fails_all') || '0', 10);
  if (fails >= AUTH.MAX_FAILS || global >= AUTH.MAX_FAILS_GLOBAL) return { ok: false, error: 'bloqueado', message: 'Muitas tentativas. Tente em ' + AUTH.LOCK_MINUTES + ' min.' };
  var users = readUsers_();
  if (!users.length) return { ok: false, error: 'sem_usuario', message: 'Crie o administrador pelo menu Saúde da planilha.' };
  var u = users.filter(function (x) { return x.login === normLogin_(login); })[0];
  var good = u && u.ativo && password && safeEquals_(hashPassword_(String(password), u.salt), u.hash);
  if (!good) {
    cache.put(key, String(fails + 1), AUTH.LOCK_MINUTES * 60);
    cache.put('login_fails_all', String(global + 1), AUTH.LOCK_MINUTES * 60);
    return { ok: false, error: 'senha_incorreta', message: 'Usuário ou senha incorretos.' };
  }
  cache.remove(key);
  var ttl = remember ? AUTH.REMEMBER_DAYS * 86400000 : AUTH.SESSION_HOURS * 3600000;
  return { ok: true, token: issueToken_(u, ttl), exp: Date.now() + ttl, nome: u.nome };
}

// ---------- Menu (getUi() primeiro: chamadas pelo app web falham antes de fazer qualquer coisa) ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Saúde')
    .addItem('Criar administrador (primeiro uso)', 'criarAdministrador')
    .addItem('Redefinir senha de um usuário', 'redefinirSenhaMenu')
    .addItem('Encerrar todas as sessões', 'encerrarTodasSessoes')
    .addItem('Criar/conferir abas', 'configurarPlanilha')
    .addSeparator()
    .addItem('Importar histórico do Mi Fitness', 'importarHistoricoMiFitness')
    .addItem('Ativar sincronização do Strava (de hora em hora)', 'ativarSyncStrava')
    .addToUi();
}
function pedirSenha_(ui, titulo) {
  var r1 = ui.prompt(titulo, 'Senha (mínimo ' + AUTH.MIN_PASSWORD_LENGTH + ' caracteres):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return null;
  var pw = r1.getResponseText();
  if (!pw || pw.length < AUTH.MIN_PASSWORD_LENGTH) { ui.alert('Senha muito curta.'); return null; }
  var r2 = ui.prompt(titulo, 'Digite a senha de novo:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK || r2.getResponseText() !== pw) { ui.alert('Senhas diferentes. Nada foi alterado.'); return null; }
  return pw;
}
function criarAdministrador() {
  var ui = SpreadsheetApp.getUi();
  criarAbas_();
  if (readUsers_().length) { ui.alert('Já existe usuário. Crie os demais pelo app (Ajustes > Usuários) ou use "Redefinir senha".'); return; }
  var n = ui.prompt('Criar administrador', 'Seu nome:', ui.ButtonSet.OK_CANCEL); if (n.getSelectedButton() !== ui.Button.OK) return;
  var l = ui.prompt('Criar administrador', 'Usuário para entrar (ex.: gilberto):', ui.ButtonSet.OK_CANCEL); if (l.getSelectedButton() !== ui.Button.OK) return;
  var pw = pedirSenha_(ui, 'Criar administrador'); if (!pw) return;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { var r = criarUsuario_(n.getResponseText(), l.getResponseText(), pw, true); ui.alert(r.ok ? 'Administrador criado.' : r.message); }
  finally { lock.releaseLock(); }
}
function redefinirSenhaMenu() {
  var ui = SpreadsheetApp.getUi();
  var l = ui.prompt('Redefinir senha', 'Usuário:', ui.ButtonSet.OK_CANCEL); if (l.getSelectedButton() !== ui.Button.OK) return;
  var u = readUsers_().filter(function (x) { return x.login === normLogin_(l.getResponseText()); })[0];
  if (!u) { ui.alert('Usuário não encontrado.'); return; }
  var pw = pedirSenha_(ui, 'Nova senha de ' + u.nome); if (!pw) return;
  var salt = Utilities.getUuid();
  setUserFields_(u, { hash: hashPassword_(pw, salt), salt: salt, ativo: true, bumpGen: true });
  CacheService.getScriptCache().remove('login_fails_' + u.login);
  auditLog_('update', 'usuario', u.id, null, { senha: 'redefinida (menu)' }, '', '');
  ui.alert('Senha de ' + u.nome + ' redefinida. As sessões dele(a) foram encerradas.');
}
function encerrarTodasSessoes() {
  var ui = SpreadsheetApp.getUi();
  props_().setProperty('AUTH_SECRET', Utilities.getUuid() + Utilities.getUuid());
  ui.alert('Sessões de todos os usuários encerradas.');
}
function configurarPlanilha() {
  SpreadsheetApp.getUi();
  criarAbas_();
}
function criarAbas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  usersSheet_();
  Object.keys(SCHEMA).forEach(function (k) { sheetFor_(ss, SCHEMA[k]); });
  configSheet_(ss);
}
function sheetFor_(ss, sc) {
  var sh = ss.getSheetByName(sc.aba);
  if (!sh) {
    sh = ss.insertSheet(sc.aba);
    var head = ['ID', 'Usuário'].concat(sc.campos.map(function (c) { return c[1]; })).concat(['Criado em', 'Alterado em']);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.getRange(2, COL.CAMPOS, sh.getMaxRows() - 1, 1).setNumberFormat('@'); // data como texto aaaa-mm-dd (sem fuso)
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 2).getValue()) !== 'Usuário') {
    throw new Error('Aba ' + sc.aba + ' no formato antigo (sem a coluna Usuário). Renomeie a aba e use o menu Criar/conferir abas.');
  }
  return sh;
}
function configSheet_(ss) {
  var sh = ss.getSheetByName('Config');
  if (!sh) {
    sh = ss.insertSheet('Config');
    sh.getRange(1, 1, 1, 3).setValues([['Usuário', 'Chave', 'Valor']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 1).getValue()) !== 'Usuário') {
    throw new Error('Aba Config no formato antigo. Renomeie a aba e use o menu Criar/conferir abas.');
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

// Único ponto de entrada da tela (google.script.run.api ou doPost)
function api(body) {
  try {
    body = body || {};
    if (body.action === 'login') return login_(body.login, body.password, !!body.remember);
    var u = sessionUser_(body.t);
    if (!u) return { ok: false, error: 'unauthorized' };
    switch (body.action) {
      case 'load': return load_(u);
      case 'syncStrava': return sincronizarStrava_(u.id, false);
      case 'stravaSetup': return comLock_(body, function (u) { return stravaSetup_(u, body.values || {}); });
      case 'stravaLink': return stravaLink_(u);
      case 'stravaDisconnect': return comLock_(body, function (u) { return stravaDisconnect_(u); });
      case 'changePassword': return comLock_(body, function (u) { return trocarSenha_(u, body.atual, body.nova); });
      case 'users': return u.admin ? listarUsuarios_() : { ok: false, error: 'proibido' };
      case 'userAdmin': return comLock_(body, function (u) { return u.admin ? adminUsuario_(u, body) : { ok: false, error: 'proibido' }; });
      case 'op': return comLock_(body, function (u) { return handleOp_(u, body); });
    }
    return { ok: false, error: 'acao_desconhecida' };
  } catch (err) {
    return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) };
  }
}
// Revalida a sessão dentro da trava: senha trocada, usuário desativado ou admin retirado enquanto esperava não passam.
function comLock_(body, fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'ocupado', message: 'Planilha ocupada, tente de novo.' }; }
  try { var u = sessionUser_(body.t); return u ? fn(u) : { ok: false, error: 'unauthorized' }; } finally { lock.releaseLock(); }
}

function load_(u) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true, me: { id: u.id, nome: u.nome, login: u.login, admin: u.admin }, config: readConfig_(ss, u.id), strava: stravaStatus_(u.id) };
  Object.keys(SCHEMA).forEach(function (k) { out[k] = readRows_(ss, SCHEMA[k], u.id); });
  return out;
}
function readRows_(ss, sc, uid) {
  var sh = sheetFor_(ss, sc), last = sh.getLastRow();
  if (last < 2) return [];
  var n = sc.campos.length + 2;
  return sh.getRange(2, 1, last - 1, n).getValues()
    .filter(function (r) { return r[0] !== '' && String(r[1]) === uid; })
    .map(function (r) {
      var o = { id: String(r[0]) };
      sc.campos.forEach(function (c, i) {
        var v = r[i + 2];
        if (c[2] === 'date') v = v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
        else if (c[2] === 'num') v = v === '' ? null : parseNumBR(v);
        else v = String(v);
        o[c[0]] = (typeof v === 'number' && isNaN(v)) ? null : v;
      });
      return o;
    });
}
function readConfig_(ss, uid) {
  var sh = configSheet_(ss), o = {};
  if (sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    if (String(r[0]) === uid && r[1]) o[String(r[1])] = r[2] instanceof Date ? Utilities.formatDate(r[2], TZ, 'yyyy-MM-dd') : r[2];
  });
  return o;
}

// ---------- Contas (troca de senha e administração) ----------
function trocarSenha_(u, atual, nova) {
  if (!atual || !safeEquals_(hashPassword_(String(atual), u.salt), u.hash)) return { ok: false, error: 'senha_incorreta', message: 'Senha atual incorreta.' };
  if (!nova || String(nova).length < AUTH.MIN_PASSWORD_LENGTH) return { ok: false, error: 'validacao', message: 'Nova senha com no mínimo ' + AUTH.MIN_PASSWORD_LENGTH + ' caracteres.' };
  var salt = Utilities.getUuid();
  setUserFields_(u, { hash: hashPassword_(String(nova), salt), salt: salt, bumpGen: true });
  auditLog_('update', 'usuario', u.id, null, { senha: 'trocada pelo próprio usuário' }, '', u.id);
  return { ok: true, relogin: true };
}
function listarUsuarios_() {
  return { ok: true, users: readUsers_().map(function (x) { return { id: x.id, nome: x.nome, login: x.login, admin: x.admin, ativo: x.ativo }; }) };
}
// body.kind: 'create' {nome, login, senha, admin} | 'reset' {id, senha} | 'ativo' {id, ativo} | 'admin' {id, admin}
function adminUsuario_(me, body) {
  var v = body.values || {};
  if (body.kind === 'create') return criarUsuario_(v.nome, v.login, v.senha, !!v.admin);
  var alvo = userById_(body.id);
  if (!alvo) return { ok: false, error: 'nao_encontrado', message: 'Usuário não encontrado.' };
  if (alvo.id === me.id && body.kind !== 'reset') return { ok: false, error: 'proibido', message: 'Você não pode desativar nem tirar o seu próprio acesso de administrador.' };
  if (body.kind === 'reset') {
    if (!v.senha || String(v.senha).length < AUTH.MIN_PASSWORD_LENGTH) return { ok: false, error: 'validacao', message: 'Senha com no mínimo ' + AUTH.MIN_PASSWORD_LENGTH + ' caracteres.' };
    var salt = Utilities.getUuid();
    setUserFields_(alvo, { hash: hashPassword_(String(v.senha), salt), salt: salt, bumpGen: true });
    CacheService.getScriptCache().remove('login_fails_' + alvo.login);
  } else if (body.kind === 'ativo') setUserFields_(alvo, { ativo: !!v.ativo, bumpGen: !v.ativo });
  else if (body.kind === 'admin') setUserFields_(alvo, { admin: !!v.admin, bumpGen: true });
  else return { ok: false, error: 'operacao_invalida' };
  auditLog_('update', 'usuario', alvo.id, null, { acao: body.kind, ativo: v.ativo, admin: v.admin }, '', me.id);
  return { ok: true, relogin: alvo.id === me.id };
}

// ---------- Gravação ----------
function handleOp_(u, body) {
  if (!/^[a-z0-9]{8,40}$/i.test(String(body.opId || ''))) return { ok: false, error: 'opid_invalido' };
  var cache = CacheService.getScriptCache(), key = 'op_' + u.id + '_' + body.opId;
  var done = cache.get(key);
  if (done) { var r0 = JSON.parse(done); r0.repeated = true; return r0; }
  var prev = findOpInAudit_(body.opId, u.id); // cache pode ter expirado: confere na auditoria
  if (prev) return { ok: true, id: prev, repeated: true };
  var result;
  try { result = runOp_(u, body); } catch (err) { return { ok: false, error: 'falha', message: String(err && err.message ? err.message : err) }; }
  if (result && result.ok) cache.put(key, JSON.stringify(result), 21600);
  return result;
}

function runOp_(u, body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (body.entity === 'config') return saveConfig_(ss, u.id, body.values || {}, body.opId);
  var sc = SCHEMA[body.entity];
  if (!sc || !Object.prototype.hasOwnProperty.call(SCHEMA, body.entity)) return { ok: false, error: 'entidade_invalida' };
  var sh = sheetFor_(ss, sc), now = new Date(), n = sc.campos.length;

  if (body.kind === 'create') {
    var v = validate_(sc, body.values || {}, true);
    if (!v.ok) return v;
    var id = nextId_(sh, sc.prefixo);
    var row = [id, u.id].concat(sc.campos.map(function (c) { return v.values[c[0]]; })).concat([now, now]);
    var r = sh.getLastRow() + 1;
    sh.getRange(r, COL.CAMPOS).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    auditLog_('create', body.entity, id, null, v.values, body.opId, u.id);
    return { ok: true, id: id };
  }

  var rowNum = findRow_(sh, body.id);
  if (!rowNum) return { ok: false, error: 'nao_encontrado', message: 'Registro não encontrado.' };
  var cur = sh.getRange(rowNum, 1, 1, n + 2).getValues()[0];
  if (String(cur[1]) !== u.id) return { ok: false, error: 'nao_encontrado', message: 'Registro não encontrado.' }; // de outro usuário
  var before = {};
  sc.campos.forEach(function (c, i) { before[c[0]] = cur[i + 2] instanceof Date ? Utilities.formatDate(cur[i + 2], TZ, 'yyyy-MM-dd') : cur[i + 2]; });

  if (body.kind === 'update') {
    var merged = {};
    Object.keys(before).forEach(function (k) { merged[k] = before[k]; });
    Object.keys(body.values || {}).forEach(function (k) { if (Object.prototype.hasOwnProperty.call(before, k)) merged[k] = body.values[k]; });
    var up = validate_(sc, merged, true);
    if (!up.ok) return up;
    sh.getRange(rowNum, COL.CAMPOS, 1, n).setValues([sc.campos.map(function (c) { return up.values[c[0]]; })]);
    sh.getRange(rowNum, n + 4).setValue(now);
    auditLog_('update', body.entity, body.id, before, up.values, body.opId, u.id);
    return { ok: true, id: body.id };
  }
  if (body.kind === 'remove') {
    sh.deleteRow(rowNum);
    auditLog_('remove', body.entity, body.id, before, null, body.opId, u.id);
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
  if (sc.aba === 'Medicacao') {
    if (REMEDIOS.indexOf(out.remedio) < 0) erros.push('Escolha Tirzepatida ou Semaglutida');
    if (out.dose === '') erros.push('Informe a dose');
    if (LOCAIS.indexOf(out.local) < 0) erros.push('Local inválido');
  }
  if (erros.length) return { ok: false, error: 'validacao', message: erros.join('; ') };
  return { ok: true, values: out };
}
// Evita injeção de fórmula na planilha (texto começando com = + - @)
function sanitizeText_(v) {
  var s = String(v == null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function saveConfig_(ss, uid, values, opId) {
  var sh = configSheet_(ss), before = readConfig_(ss, uid), erros = [], ok = {};
  Object.keys(values).forEach(function (k) {
    if (!Object.prototype.hasOwnProperty.call(CONFIG_KEYS, k)) return;
    var v = values[k];
    if (CONFIG_KEYS[k]) {
      if (v === '' || v == null) v = '';
      else {
        v = parseNumCampo_(v, false);
        if (isNaN(v) || v < CONFIG_KEYS[k][0] || v > CONFIG_KEYS[k][1]) { erros.push(k + ' fora da faixa'); return; }
      }
    } else if (k === 'nascimento') {
      v = String(v || '');
      if (v && !validDate_(v)) { erros.push('Nascimento inválido'); return; }
    } else v = (v === 'M' || v === 'F') ? v : '';
    ok[k] = v;
  });
  if (erros.length) return { ok: false, error: 'validacao', message: erros.join('; ') };
  var rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  Object.keys(ok).forEach(function (k) {
    var r = 0;
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === uid && String(rows[i][1]) === k) { r = i + 2; break; }
    if (r) { sh.getRange(r, 3).setNumberFormat('@').setValue(String(ok[k])); }
    else { var nr = sh.getLastRow() + 1; sh.getRange(nr, 3).setNumberFormat('@'); sh.getRange(nr, 1, 1, 3).setValues([[uid, k, String(ok[k])]]); rows.push([uid, k]); }
  });
  auditLog_('update', 'config', uid, before, ok, opId, uid);
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
function findOpInAudit_(opId, uid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Auditoria');
  if (!sh || sh.getLastRow() < 2) return null;
  var from = Math.max(2, sh.getLastRow() - 2000), vals = sh.getRange(from, 4, sh.getLastRow() - from + 1, 5).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][3]) === opId && String(vals[i][4]) === uid) return String(vals[i][0]);
  return null;
}
function validDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s));
  if (!m) return false;
  var d = new Date(+m[1], +m[2] - 1, +m[3], 12);
  return d.getFullYear() === +m[1] && d.getMonth() === +m[2] - 1 && d.getDate() === +m[3] && +m[1] >= 1900;
}
function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
// Auditoria: Quando | Ação | Entidade | ID | Antes | Depois | OpId | Usuário
function auditLog_(action, entity, id, before, after, opId, uid) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Auditoria');
    if (!sh) {
      sh = ss.insertSheet('Auditoria');
      sh.getRange(1, 1, 1, 8).setValues([['Quando', 'Ação', 'Entidade', 'ID', 'Antes (JSON)', 'Depois (JSON)', 'OpId', 'Usuário']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), action, entity, id, JSON.stringify(before || null), JSON.stringify(after || null), opId || '', uid || '']);
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
