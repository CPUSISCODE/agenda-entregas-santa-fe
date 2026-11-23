/**
 * Agenda de Entregas Futuras – Armazém Santa Fé
 * Servidor (Google Apps Script). Guarda tudo no Google Drive do dono:
 *  - planilha "Agenda de Entregas – Banco" (abas usuarios, vendas, sessoes, log)
 *  - pasta "XMLs" com os arquivos das notas
 * Implantação: Implantar > Nova implantação > App da Web
 *   Executar como: Eu   ·   Quem pode acessar: Qualquer pessoa
 * (o acesso aos dados é controlado pelo login e senha do próprio sistema)
 */
var PLANILHA_ID = '1n5-bQ59cNqadWvEplNspMhhz8xDyYh0ZU6Xlwy8QvnM'; // planilha do banco no Drive
var NOME_PASTA = 'Agenda de Entregas – Armazém Santa Fé';
var NOME_PLANILHA = 'Agenda de Entregas – Banco';
var HORAS_SESSAO = 12;
var PERFIS = ['admin', 'operador', 'consulta'];

/* ---------------- infraestrutura ---------------- */
function props_() { return PropertiesService.getScriptProperties(); }
function pasta_() {
  var id = props_().getProperty('PASTA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var pais = DriveApp.getFileById(planilha_().getId()).getParents();
  var p = pais.hasNext() ? pais.next() : DriveApp.createFolder(NOME_PASTA);
  props_().setProperty('PASTA_ID', p.getId());
  return p;
}
function pastaXml_() {
  var id = props_().getProperty('XML_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var p = pasta_(), it = p.getFoldersByName('XMLs');
  var f = it.hasNext() ? it.next() : p.createFolder('XMLs');
  props_().setProperty('XML_ID', f.getId());
  return f;
}
var CAB = {
  usuarios: ['login', 'nome', 'perfil', 'ativo', 'salt', 'hash', 'trocarSenha', 'criadoEm', 'criadoPor', 'ultimoAcesso'],
  vendas: ['id', 'cliente', 'numero', 'loja', 'emissao', 'prazo', 'status', 'json', 'atualizadoEm', 'atualizadoPor'],
  sessoes: ['token', 'login', 'expira'],
  log: ['quando', 'login', 'acao', 'detalhe']
};
var SS_ = null;
function planilha_() {
  if (SS_) return SS_;
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(PLANILHA_ID);
  Object.keys(CAB).forEach(function (nome) {
    var sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      sh.getRange(1, 1, 1, CAB[nome].length).setValues([CAB[nome]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.getRange('A:Z').setNumberFormat('@');
    }
  });
  var padrao = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Planilha1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
  SS_ = ss;
  return ss;
}
function aba_(nome) { return planilha_().getSheetByName(nome); }
function linhas_(nome) {
  var sh = aba_(nome), n = sh.getLastRow();
  if (n < 2) return [];
  var cab = CAB[nome], v = sh.getRange(2, 1, n - 1, cab.length).getValues();
  return v.map(function (r, i) { var o = { _linha: i + 2 }; cab.forEach(function (c, j) { o[c] = r[j]; }); return o; });
}
function gravar_(nome, linha, obj) {
  var cab = CAB[nome], sh = aba_(nome);
  var row = cab.map(function (c) { return obj[c] == null ? '' : String(obj[c]); });
  if (linha) sh.getRange(linha, 1, 1, cab.length).setValues([row]);
  else sh.appendRow(row);
}
function log_(login, acao, detalhe) {
  try { aba_('log').appendRow([new Date().toISOString(), login || '', acao, String(detalhe || '').slice(0, 500)]); } catch (e) {}
}

/* ---------------- senhas e sessões ---------------- */
function hash_(senha, salt) {
  var h = salt + '|' + senha;
  for (var i = 0; i < 1500; i++) {
    var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = Utilities.base64Encode(b) + salt;
  }
  return h;
}
function novoSalt_() { return Utilities.getUuid().replace(/-/g, ''); }
function senhaValida_(s) {
  if (!s || String(s).length < 6) throw erro_('A senha precisa ter pelo menos 6 caracteres.');
}
function erro_(msg, cod) { var e = new Error(msg); e.cod = cod || 'erro'; return e; }
function limpaLogin_(l) { return String(l || '').trim().toLowerCase().replace(/[^a-z0-9._-]/g, ''); }
function publico_(u) { return { login: u.login, nome: u.nome, perfil: u.perfil, ativo: String(u.ativo) === 'true', trocarSenha: String(u.trocarSenha) === 'true', ultimoAcesso: u.ultimoAcesso || '', criadoEm: u.criadoEm || '' }; }
function acharUsuario_(login) {
  login = limpaLogin_(login);
  var L = linhas_('usuarios');
  for (var i = 0; i < L.length; i++) if (L[i].login === login) return L[i];
  return null;
}
function sessao_(token) {
  if (!token) throw erro_('Faça login novamente.', 'sessao');
  var cache = CacheService.getScriptCache(), login = cache.get('s:' + token);
  if (!login) {
    var S = linhas_('sessoes'), agora = new Date().toISOString();
    for (var i = 0; i < S.length; i++) if (S[i].token === token && S[i].expira > agora) { login = S[i].login; break; }
    if (login) cache.put('s:' + token, login, 1800);
  }
  if (!login) throw erro_('Sua sessão expirou. Faça login novamente.', 'sessao');
  var u = acharUsuario_(login);
  if (!u || String(u.ativo) !== 'true') throw erro_('Usuário desativado. Fale com o administrador.', 'sessao');
  return u;
}
function exige_(u, perfis) { if (perfis.indexOf(u.perfil) < 0) throw erro_('Seu usuário não tem permissão para isso.', 'permissao'); }

/* ---------------- ações ---------------- */
var ACOES = {
  status: function () { return { temAdmin: linhas_('usuarios').length > 0 }; },

  criarAdmin: function (p) {
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      if (linhas_('usuarios').length) throw erro_('O administrador já foi criado. Faça login.');
      var login = limpaLogin_(p.login); if (!login) throw erro_('Informe o usuário.');
      senhaValida_(p.senha);
      var salt = novoSalt_();
      gravar_('usuarios', 0, { login: login, nome: String(p.nome || login).trim(), perfil: 'admin', ativo: 'true', salt: salt, hash: hash_(p.senha, salt), trocarSenha: 'false', criadoEm: new Date().toISOString(), criadoPor: 'instalação' });
      log_(login, 'criarAdmin', '');
    } finally { lock.releaseLock(); }
    return ACOES.login({ login: p.login, senha: p.senha });
  },

  login: function (p) {
    var login = limpaLogin_(p.login), cache = CacheService.getScriptCache(), k = 'f:' + login;
    var falhas = +(cache.get(k) || 0);
    if (falhas >= 5) throw erro_('Muitas tentativas erradas. Espere 10 minutos e tente de novo.');
    var u = acharUsuario_(login);
    if (!u || String(u.ativo) !== 'true' || hash_(String(p.senha || ''), u.salt) !== u.hash) {
      cache.put(k, String(falhas + 1), 600);
      log_(login, 'loginFalhou', '');
      throw erro_('Usuário ou senha incorretos.');
    }
    cache.remove(k);
    var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    var expira = new Date(Date.now() + HORAS_SESSAO * 3600e3).toISOString();
    gravar_('sessoes', 0, { token: token, login: u.login, expira: expira });
    cache.put('s:' + token, u.login, 1800);
    u.ultimoAcesso = new Date().toISOString(); gravar_('usuarios', u._linha, u);
    limparSessoes_();
    log_(u.login, 'login', '');
    return { token: token, usuario: publico_(u) };
  },

  sair: function (p) {
    CacheService.getScriptCache().remove('s:' + p.token);
    var S = linhas_('sessoes');
    for (var i = S.length - 1; i >= 0; i--) if (S[i].token === p.token) aba_('sessoes').deleteRow(S[i]._linha);
    return { ok: true };
  },

  eu: function (p, u) { return { usuario: publico_(u) }; },

  trocarSenha: function (p, u) {
    if (hash_(String(p.atual || ''), u.salt) !== u.hash) throw erro_('A senha atual não confere.');
    senhaValida_(p.nova);
    u.salt = novoSalt_(); u.hash = hash_(p.nova, u.salt); u.trocarSenha = 'false';
    gravar_('usuarios', u._linha, u); log_(u.login, 'trocarSenha', '');
    return { usuario: publico_(u) };
  },

  nomes: function () {
    var o = {}; linhas_('usuarios').forEach(function (u) { o[u.login] = u.nome; }); return o;
  },

  listar: function () {
    return { vendas: linhas_('vendas').map(function (r) { try { var v = JSON.parse(r.json); v.id = r.id; return v; } catch (e) { return null; } }).filter(Boolean), em: new Date().toISOString() };
  },

  salvar: function (p, u) {
    exige_(u, ['admin', 'operador']);
    return salvarVarios_([{ id: p.id, obj: p.obj }], u);
  },

  importar: function (p, u) {
    exige_(u, ['admin']);
    return salvarVarios_(p.itens || [], u, !!p.manterExistentes);
  },

  excluir: function (p, u) {
    exige_(u, ['admin', 'operador']);
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var L = linhas_('vendas');
      for (var i = L.length - 1; i >= 0; i--) if (L[i].id === String(p.id)) aba_('vendas').deleteRow(L[i]._linha);
      log_(u.login, 'excluir', p.id);
    } finally { lock.releaseLock(); }
    return { ok: true };
  },

  enviarArquivo: function (p, u) {
    exige_(u, ['admin', 'operador']);
    var nome = String(p.nome || 'arquivo').replace(/[\\/]/g, '_').slice(0, 150);
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.tipo || 'application/octet-stream', nome);
    var f = pastaXml_().createFile(blob);
    log_(u.login, 'arquivo', nome);
    return { id: f.getId() };
  },

  baixarArquivo: function (p) {
    var f = DriveApp.getFileById(String(p.id));
    var pais = f.getParents(), ok = false, xml = pastaXml_().getId();
    while (pais.hasNext()) if (pais.next().getId() === xml) ok = true;
    if (!ok) throw erro_('Arquivo fora da pasta da agenda.');
    var b = f.getBlob();
    return { nome: f.getName(), tipo: b.getContentType(), base64: Utilities.base64Encode(b.getBytes()) };
  },

  usuarios: function (p, u) {
    exige_(u, ['admin']);
    return { usuarios: linhas_('usuarios').map(publico_) };
  },

  salvarUsuario: function (p, u) {
    exige_(u, ['admin']);
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var login = limpaLogin_(p.login); if (!login) throw erro_('Informe o usuário (só letras, números, ponto, hífen).');
      if (PERFIS.indexOf(p.perfil) < 0) throw erro_('Perfil inválido.');
      var x = acharUsuario_(login), novo = !x;
      if (novo) { if (!p.senha) throw erro_('Informe a senha inicial.'); x = { login: login, criadoEm: new Date().toISOString(), criadoPor: u.login }; }
      if (x.login === u.login && (p.perfil !== 'admin' || p.ativo === false)) throw erro_('Você não pode tirar o seu próprio acesso de administrador.');
      x.nome = String(p.nome || login).trim(); x.perfil = p.perfil; x.ativo = p.ativo === false ? 'false' : 'true';
      if (p.senha) { senhaValida_(p.senha); x.salt = novoSalt_(); x.hash = hash_(p.senha, x.salt); x.trocarSenha = p.trocarSenha === false ? 'false' : 'true'; }
      gravar_('usuarios', novo ? 0 : x._linha, x);
      if (x.ativo !== 'true' || p.senha) encerrarSessoes_(login);
      log_(u.login, novo ? 'criarUsuario' : 'alterarUsuario', login + ' · ' + x.perfil + (p.senha ? ' · senha definida' : ''));
      return { usuario: publico_(x) };
    } finally { lock.releaseLock(); }
  },

  historico: function (p, u) {
    exige_(u, ['admin']);
    var sh = aba_('log'), n = sh.getLastRow();
    if (n < 2) return { itens: [] };
    var ini = Math.max(2, n - 199), v = sh.getRange(ini, 1, n - ini + 1, 4).getValues();
    return { itens: v.reverse().map(function (r) { return { quando: String(r[0]), login: String(r[1]), acao: String(r[2]), detalhe: String(r[3]) }; }) };
  }
};
var PUBLICAS = { status: 1, criarAdmin: 1, login: 1 };

function salvarVarios_(itens, u, manter) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = aba_('vendas'), L = linhas_('vendas'), pos = {}, agora = new Date().toISOString(), novas = [], n = 0;
    L.forEach(function (r) { pos[r.id] = r._linha; });
    itens.forEach(function (it) {
      var id = String(it.id || (it.obj && (it.obj.id || it.obj.chave)) || ''); if (!id) return;
      if (manter && pos[id]) return;
      var o = it.obj || {}; delete o.id;
      var json = JSON.stringify(o);
      if (json.length > 49000) throw erro_('A venda ' + id + ' ficou grande demais para a planilha.');
      var row = [id, o.cliente || '', o.numero || '', o.loja || '', o.emissao || '', o.prazo || '', o.status || '', json, agora, u.login];
      if (pos[id]) sh.getRange(pos[id], 1, 1, row.length).setValues([row]); else novas.push(row);
      n++;
    });
    if (novas.length) sh.getRange(sh.getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
    log_(u.login, itens.length > 1 ? 'importar' : 'salvar', itens.length > 1 ? n + ' vendas' : (itens[0] && itens[0].id));
    return { ok: true, gravadas: n };
  } finally { lock.releaseLock(); }
}
function encerrarSessoes_(login) {
  var S = linhas_('sessoes'), c = CacheService.getScriptCache();
  for (var i = S.length - 1; i >= 0; i--) if (S[i].login === login) { c.remove('s:' + S[i].token); aba_('sessoes').deleteRow(S[i]._linha); }
}
function limparSessoes_() {
  var S = linhas_('sessoes'), agora = new Date().toISOString();
  for (var i = S.length - 1; i >= 0; i--) if (S[i].expira < agora) aba_('sessoes').deleteRow(S[i]._linha);
}

/* ---------------- entrada HTTP ---------------- */
function doPost(e) {
  var r;
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var f = ACOES[p.acao];
    if (!f) throw erro_('Ação desconhecida.');
    var u = PUBLICAS[p.acao] ? null : sessao_(p.token);
    r = { ok: true, dados: f(p, u) };
  } catch (err) {
    r = { ok: false, erro: String(err && err.message || err), cod: err && err.cod || 'erro' };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, servico: 'Agenda de Entregas – Armazém Santa Fé', temAdmin: ACOES.status().temAdmin })).setMimeType(ContentService.MimeType.JSON);
}
/** Rode esta função uma vez no editor para autorizar e criar a planilha e as pastas. */
function instalar() {
  var ss = planilha_(); pastaXml_();
  Logger.log('Pronto. Planilha: ' + ss.getUrl());
}
