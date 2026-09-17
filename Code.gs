// =====================================================
// GOOGLE APPS SCRIPT - API AGENDA DA PREFEITURA
// =====================================================

// ID da planilha guardado como Script Property (Configurações do projeto no
// editor do Apps Script → Propriedades do script) — NUNCA em texto puro no
// código-fonte, que é público no GitHub (achado real, 2026-09-17: usuário
// perguntou "tem como fazer proteção" pensando num invasor lendo o
// repositório público; o ID sozinho não abre a planilha — ela já exige
// login autorizado, confirmei — mas não custa nada tirar ele daqui também).
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const LIMITE_POR_SESSAO = 5;
// Excecoes de limite por orgao (demanda maior de usuarios). Demais orgaos usam LIMITE_POR_SESSAO.
const LIMITES_ESPECIAIS = {
  'Secretaria de Educação': 10,
  'Secretaria de Saúde': 10,
  'Secretaria de Comunicação': 20
};
function limitePara(orgao) {
  return LIMITES_ESPECIAIS[orgao] || LIMITE_POR_SESSAO;
}

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    let p = Object.assign({}, e.parameter || {});
    if (e.postData && e.postData.contents) {
      try { p = Object.assign({}, p, JSON.parse(e.postData.contents)); } catch(ex) {}
    }
    const action = p.action || '';
    switch (action) {
      case 'login':               return responder(handleLogin(p), output);
      case 'getEventos':          return responder(handleGetEventos(p), output);
      case 'criarEvento':         return responder(handleCriarEvento(p), output);
      case 'atualizarEvento':     return responder(handleAtualizarEvento(p), output);
      case 'atualizarRecorrencia':return responder(handleAtualizarRecorrencia(p), output);
      case 'excluirEvento':       return responder(handleExcluirEvento(p), output);
      case 'excluirSerieRecorrente': return responder(handleExcluirSerieRecorrente(p), output);
      case 'getUsuarios':         return responder(handleGetUsuarios(p), output);
      case 'getContagemUsuarios': return responder(handleGetContagemUsuarios(p), output);
      case 'criarUsuario':        return responder(handleCriarUsuario(p), output);
      case 'resetarSenha':        return responder(handleResetarSenha(p), output);
      case 'excluirUsuario':      return responder(handleExcluirUsuario(p), output);
      case 'solicitarAcesso':     return responder(handleSolicitarAcesso(p), output);
      case 'recuperarSenha':      return responder(handleRecuperarSenha(p), output);
      case 'getSolicitacoes':     return responder(handleGetSolicitacoes(p), output);
      case 'atualizarSolicitacao':return responder(handleAtualizarSolicitacao(p), output);
      case 'excluirSolicitacao': return responder(handleExcluirSolicitacao(p), output);
      case 'setup':               return responder(criarAbas(), output);
      case 'primeiroAdmin':       return responder(handlePrimeiroAdmin(p), output);
      default: return responder({ error: 'Acao nao encontrada: ' + action }, output);
    }
  } catch (err) {
    return responder({ error: err.toString() }, output);
  }
}

function responder(data, output) {
  output.setContent(JSON.stringify(data));
  return output;
}

// ── Organograma oficial: nome completo → sigla ───────────────────────────────
var SIGLAS_ORGAOS = {
  'Gabinete do Prefeito':'PREFEITO','Chefia de Gabinete':'GABINETE',
  'Secretaria de Administração':'SECAD','Secretaria de Agricultura e Abastecimento':'SEAGRI',
  'Secretaria de Assistência Social':'SAS','Secretaria de Assuntos Jurídicos e Legislativos':'SEAJUR',
  'Secretaria de Comunicação':'SECOM','Secretaria de Cultura':'SECULT',
  'Secretaria de Desenvolvimento Econômico':'SEDEPP','Secretaria de Educação':'SEDUC',
  'Secretaria de Esporte':'SEMEPP','Secretaria de Finanças':'SEFIN',
  'Secretaria de Meio Ambiente':'SEMEA',
  'Secretaria de Mobilidade Urbana e Cooperação em Segurança Pública':'SEMOB',
  'Secretaria de Obras e Serviços Públicos':'SOSP',
  'Secretaria de Planejamento, Desenvolvimento Urbano e Habitação':'SEPLAN',
  'Secretaria de Saúde':'SESAU','Secretaria de Tecnologia da Informação':'SETEC',
  'Secretaria de Turismo':'SETUR',
  'Comissão Interna de Prevenção de Acidentes e Assédio':'CIPA','Comitê Gestor da Praça CEU':'CGPCEU',
  'Controladoria Geral do Município':'CGM','Controladoria-Geral do Município':'CGM',
  'Coordenadoria da Juventude':'JUVENTUDE','Coordenadoria da Pessoa com Deficiência':'CPD',
  'Coordenadoria de Proteção e Defesa Civil':'COMPDEC','Coordenadoria Municipal do Idoso':'IDOSOPP',
  'Instituto do Idoso de Presidente Prudente':'IDOSOPP',
  'Fundo Municipal de Defesa dos Interesses Difusos':'FMDID',
  'Fundo Social de Solidariedade de Presidente Prudente':'FUNDO',
  'Núcleo da Escola Federativa do Município de Presidente Prudente':'NEF',
  'Serviço Especializado em Engenharia de Segurança e em Medicina do Trabalho':'SESMT',
  // Secretarias Indiretas
  'INOVA PRUDENTE':'INOVA',
  'SASSOM':'SASSOM',
  'PRUDENPREV':'PRUDENPREV',
  'PRUDENCO':'PRUDENCO'
};

function getSiglaOrgao(orgao) {
  if (!orgao) return '';
  // NFC normaliza representações diferentes de acentos (ã, ç, é…) vindos da planilha
  var norm = String(orgao).trim().normalize('NFC');
  if (SIGLAS_ORGAOS[norm]) return SIGLAS_ORGAOS[norm];
  var lower = norm.toLowerCase();
  var keys = Object.keys(SIGLAS_ORGAOS);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].normalize('NFC').toLowerCase() === lower) return SIGLAS_ORGAOS[keys[i]];
  }
  // Retorna '' para que o frontend use seu próprio lookup com fallback
  return '';
}

// Helpers
function getSheet(nome) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nome);
}

function lerAba(nome) {
  const sheet = getSheet(nome);
  if (!sheet) return { headers: [], rows: [] };
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 1) return { headers: [], rows: [] };
  const headers = vals[0].map(String);
  const rows = vals.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return { headers: headers, rows: rows, sheet: sheet };
}

// Le so a coluna 'id' (nao a aba inteira) pra achar o maior id — evita reconstruir
// objeto por objeto de todas as colunas so pra descobrir um numero.
function proximoId(nome) {
  const sheet = getSheet(nome);
  if (!sheet) return 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idCol = headers.indexOf('id') + 1;
  if (idCol < 1) return 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var v = Number(ids[i][0]) || 0;
    if (v > max) max = v;
  }
  return max + 1;
}

// Log e so-escrita (nunca lido de volta pelo app) — usar o timestamp como id evita
// escanear a aba de logs (que so cresce) a cada acao registrada, inclusive as que
// nao tem nada a ver com log (login, criar evento, etc. chamam registrarLog).
function registrarLog(usuario, acao, detalhes) {
  try {
    const sheet = getSheet('logs');
    if (!sheet) return;
    sheet.appendRow([new Date().getTime(), new Date().toISOString(), usuario, acao, detalhes || '']);
  } catch (e) {}
}

// ── Cache de abas grandes (CacheService tem limite de ~100KB por chave — divide em pedaços) ──
const CACHE_CHUNK = 60000;

function cachePutGrande(chave, valor, ttlSegundos) {
  try {
    const cache = CacheService.getScriptCache();
    const texto = JSON.stringify(valor);
    const partes = Math.max(1, Math.ceil(texto.length / CACHE_CHUNK));
    const obj = {};
    for (var i = 0; i < partes; i++) {
      obj[chave + '_' + i] = texto.substring(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
    }
    obj[chave + '_n'] = String(partes);
    cache.putAll(obj, ttlSegundos);
  } catch (e) {} // cache e so otimizacao — falha aqui nunca deve quebrar a resposta real
}

function cacheGetGrande(chave) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(chave + '_n'));
    if (!n) return null;
    const chaves = [];
    for (var i = 0; i < n; i++) chaves.push(chave + '_' + i);
    const valores = cache.getAll(chaves);
    var texto = '';
    for (var j = 0; j < n; j++) {
      var v = valores[chave + '_' + j];
      if (v === undefined) return null; // pedaco expirou/faltando — trata como cache miss
      texto += v;
    }
    return JSON.parse(texto);
  } catch (e) { return null; }
}

function invalidarCacheAba(nome) {
  try { CacheService.getScriptCache().remove('aba_' + nome + '_n'); } catch (e) {}
}

// Igual lerAba(), mas com cache curto — usar so em leituras (nunca em quem vai escrever
// na planilha, pois o resultado nao inclui referencia ao Sheet quando vem do cache).
function lerAbaCache(nome, ttlSegundos) {
  const chave = 'aba_' + nome;
  const cached = cacheGetGrande(chave);
  if (cached) return cached;
  const dados = lerAba(nome);
  cachePutGrande(chave, { headers: dados.headers, rows: dados.rows }, ttlSegundos);
  return { headers: dados.headers, rows: dados.rows };
}

// Token
function gerarToken(user) {
  const payload = {
    id: user.id, login: user.login, nome: user.nome, email: user.email,
    tipo: user.tipo, orgao: user.orgao,
    is_prefeito: user.tipo === 'prefeito' || String(user.is_prefeito).toUpperCase() === 'TRUE',
    exp: Date.now() + 86400000
  };
  // Charset.UTF_8 garante que acentos/ç sejam corretamente codificados
  return Utilities.base64Encode(JSON.stringify(payload), Utilities.Charset.UTF_8);
}

function verificarToken(token) {
  if (!token) return null;
  try {
    const bytes = Utilities.base64Decode(token);
    const decoded = JSON.parse(Utilities.newBlob(bytes).getDataAsString('UTF-8'));
    if (decoded.exp < Date.now()) return null;
    return decoded;
  } catch (e) { return null; }
}

// Login
function handleLogin(data) {
  const usuario = data.usuario;
  const senha = data.senha;
  if (!usuario || !senha) return { error: 'Usuário e senha são obrigatórios' };

  const u = String(usuario).toLowerCase().trim();
  const rows = lerAba('usuarios').rows;
  var user = null;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if ((String(r.login).trim().toLowerCase() === u || String(r.email).trim().toLowerCase() === u) &&
        String(r.senha || '').trim() === String(senha || '').trim() &&
        String(r.ativo).toUpperCase() === 'TRUE') {
      user = r;
      break;
    }
  }

  if (!user) {
    registrarLog(usuario, 'login_falhou', 'Credenciais invalidas');
    return { error: 'Usuário ou senha incorretos' };
  }

  registrarLog(user.email, 'login', 'Login com sucesso');
  return {
    success: true,
    token: gerarToken(user),
    usuario: {
      id: user.id, nome: user.nome, email: user.email,
      tipo: user.tipo, orgao: user.orgao,
      is_prefeito: user.tipo === 'prefeito' || String(user.is_prefeito).toUpperCase() === 'TRUE'
    }
  };
}

// Eventos
// Corrige células que o Sheets converteu de texto p/ Date (mudando a hora digitada ao virar JSON/UTC)
function normalizarDataEvento(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm");
  }
  return v;
}

function handleGetEventos(params) {
  const usuario = verificarToken(params.token);
  if (!usuario) return { error: 'Não autorizado' };

  const eventosRows = lerAbaCache('eventos', 30).rows;
  const usuariosRows = lerAba('usuarios').rows;
  const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();

  var verTodos = params.todos === 'true' || params.todos === true;
  const filtrados = eventosRows.filter(function(e) {
    if (!e.id) return false;
    if (usuario.tipo === 'admin' || usuario.tipo === 'prefeito' || verTodos) return true;
    return String(e.orgao || '').trim() === String(usuario.orgao || '').trim();
  });

  // Enriquece publicado_por com o login atual (coluna B da aba usuarios)
  return filtrados.map(function(e) {
    const emailRef = String(e.email_publicado || '').trim().toLowerCase();
    const loginRef = String(e.publicado_por  || '').trim().toLowerCase();
    var pubUser = null;
    for (var i = 0; i < usuariosRows.length; i++) {
      var u = usuariosRows[i];
      const uEmail = String(u.email || '').toLowerCase();
      const uLogin = String(u.login || '').toLowerCase();
      if ((emailRef && uEmail === emailRef) || (loginRef && uLogin === loginRef)) {
        pubUser = u;
        break;
      }
    }
    var ev = {};
    for (var k in e) { ev[k] = e[k]; }
    ev.data_evento = normalizarDataEvento(e.data_evento, tz);
    ev.sigla_orgao = getSiglaOrgao(e.orgao); // coluna do organograma
    if (pubUser) {
      ev.publicado_por = pubUser.login; // coluna B da planilha de usuarios
    }
    return ev;
  });
}

// Impede o Sheets de auto-converter o texto digitado (ex: "2026-07-14T15:30") em Date,
// o que mudava a hora exibida ao serializar em UTC. Coluna C = data_evento.
function garantirTextoDataEvento(sheet) {
  sheet.getRange('C2:C').setNumberFormat('@');
}

// Gera as datas (yyyy-MM-dd) dos dias da semana selecionados, dentro do intervalo [inicio,fim]
function gerarDatasRecorrencia(dataInicioStr, dataFimStr, dias, tz) {
  const LIMITE_OCORRENCIAS = 150;
  const diasSet = {};
  for (var i = 0; i < dias.length; i++) diasSet[Number(dias[i])] = true;

  const inicio = new Date(dataInicioStr + 'T00:00:00');
  const fim = new Date(dataFimStr + 'T00:00:00');
  var cursor = new Date(inicio);
  var datas = [];
  while (cursor <= fim) {
    if (diasSet[cursor.getDay()]) {
      datas.push(Utilities.formatDate(cursor, tz, 'yyyy-MM-dd'));
      if (datas.length > LIMITE_OCORRENCIAS) return { erro: 'Limite de ' + LIMITE_OCORRENCIAS + ' ocorrências excedido. Reduza o período ou os dias selecionados.' };
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return { datas: datas };
}

// Valida os parametros de recorrencia e gera as datas (com hora fixa 00:00). Usado por criar e atualizar.
function validarRecorrencia(rec, tz) {
  if (rec.tipo !== 'semanal' && rec.tipo !== 'mensal') return { erro: 'Tipo de repetição inválido' };
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const inicioRec = new Date(rec.dataInicio + 'T00:00:00');
  const fimRec = new Date(rec.dataFim + 'T00:00:00');
  if (isNaN(inicioRec.getTime()) || isNaN(fimRec.getTime())) return { erro: 'Datas de repetição inválidas' };
  if (fimRec < inicioRec) return { erro: 'A data de fim deve ser depois da data de início' };
  if (inicioRec < hoje) return { erro: 'Não é permitido repetir a partir de uma data retroativa.' };
  const anoAtual = new Date().getFullYear();
  if (inicioRec.getFullYear() !== anoAtual || fimRec.getFullYear() !== anoAtual) {
    return { erro: 'A repetição só pode ocorrer dentro do ano ' + anoAtual + '.' };
  }
  const geradas = gerarDatasRecorrencia(rec.dataInicio, rec.dataFim, rec.dias, tz);
  if (geradas.erro) return { erro: geradas.erro };
  if (!geradas.datas.length) return { erro: 'Nenhuma data foi gerada com os dias da semana selecionados nesse período.' };
  return { datas: geradas.datas.map(function(d) { return d + 'T00:00'; }) };
}

// Faz upload de um anexo (base64) pra pasta compartilhada do Drive e devolve {anexo_url, anexo_nome} ou {error}.
function salvarAnexoDrive(arquivo_base64, arquivo_nome, arquivo_tipo) {
  try {
    const bytes = Utilities.base64Decode(arquivo_base64);
    const mimeType = arquivo_tipo || 'application/octet-stream';
    const blob = Utilities.newBlob(bytes, mimeType, arquivo_nome);
    const pastaNome = 'Agenda Prefeitura Anexos';
    const pastas = DriveApp.getFoldersByName(pastaNome);
    const pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(pastaNome);
    const file = pasta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { anexo_url: 'https://drive.google.com/file/d/' + file.getId() + '/view', anexo_nome: arquivo_nome };
  } catch (driveErr) {
    return { error: 'Erro ao salvar anexo. Execute autorizarDrive() no editor para liberar permissões.' };
  }
}

function handleCriarEvento(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };

  const titulo = data.titulo;
  const local = data.local;
  const responsavel = data.responsavel;
  const telefone = data.telefone;
  const observacao = data.observacao;
  if (!titulo || !observacao) return { error: 'Título e observação são obrigatórios' };

  const sheet = getSheet('eventos');
  if (!sheet) return { error: 'Aba eventos não encontrada' };
  const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();

  // Define as datas a gerar: uma unica (evento normal) ou varias (recorrencia)
  var datasEvento = [];
  var recorrenciaTipo = '';
  var recorrenciaGrupo = '';
  const rec = data.recorrencia;

  if (rec && rec.tipo && rec.dataInicio && rec.dataFim && rec.dias && rec.dias.length) {
    if (!data.arquivo_base64 || !data.arquivo_nome) return { error: 'Anexo obrigatório para eventos que se repetem (documento com a variação de local/horário).' };
    const validado = validarRecorrencia(rec, tz);
    if (validado.erro) return { error: validado.erro };
    datasEvento = validado.datas;
    recorrenciaTipo = rec.tipo;
    recorrenciaGrupo = 'rec_' + new Date().getTime();
  } else {
    const data_evento = data.data_evento;
    if (!data_evento) return { error: 'Data é obrigatória' };
    if (new Date(data_evento) < new Date()) return { error: 'Não é permitido criar eventos com data ou hora retroativa.' };
    datasEvento = [data_evento];
  }

  var anexo_url = '', anexo_nome = '';
  if (data.arquivo_base64 && data.arquivo_nome) {
    const up = salvarAnexoDrive(data.arquivo_base64, data.arquivo_nome, data.arquivo_tipo);
    if (up.error) return { error: up.error };
    anexo_url = up.anexo_url; anexo_nome = up.anexo_nome;
  }

  const now = new Date().toISOString();
  const orgaoEvento = usuario.tipo === 'prefeito' ? 'Gabinete do Prefeito' : usuario.orgao;
  const isPref = (usuario.tipo === 'prefeito' || usuario.is_prefeito) ? 'TRUE' : 'FALSE';

  garantirTextoDataEvento(sheet);
  // Trava para gerar id + gravar como uma unica operacao: sem isso, duas secretarias
  // salvando ao mesmo tempo podem ler o mesmo "maior id" e gerar linhas duplicadas.
  var idsGerados = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    for (var i = 0; i < datasEvento.length; i++) {
      const id = proximoId('eventos');
      sheet.appendRow([
        id, titulo, datasEvento[i], local || '', responsavel || '',
        telefone || '', observacao, anexo_url, anexo_nome,
        orgaoEvento, isPref, usuario.login, usuario.email, now, now, 'ativo',
        recorrenciaTipo, recorrenciaGrupo
      ]);
      idsGerados.push(id);
    }
  } finally {
    lock.releaseLock();
  }
  invalidarCacheAba('eventos');

  registrarLog(usuario.email, 'criar_evento', titulo + (idsGerados.length > 1 ? (' (recorrência x' + idsGerados.length + ')') : ''));
  return {
    success: true, id: idsGerados[0], ids: idsGerados, total: idsGerados.length,
    message: idsGerados.length > 1 ? (idsGerados.length + ' eventos criados (repetição)') : 'Evento criado com sucesso'
  };
}

function handleAtualizarEvento(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };

  const aba = lerAba('eventos');
  const rows = aba.rows;
  const headers = aba.headers;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.id)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Evento não encontrado' };

  if (usuario.tipo !== 'admin' && rows[idx].orgao !== usuario.orgao) return { error: 'Acesso negado' };

  garantirTextoDataEvento(sheet);
  const rowNum = idx + 2;
  const campos = ['titulo', 'data_evento', 'local', 'responsavel', 'telefone', 'observacao'];
  for (var c = 0; c < campos.length; c++) {
    var campo = campos[c];
    if (data[campo] !== undefined) {
      const col = headers.indexOf(campo) + 1;
      if (col > 0) sheet.getRange(rowNum, col).setValue(data[campo]);
    }
  }
  if (data.arquivo_base64 && data.arquivo_nome) {
    const up = salvarAnexoDrive(data.arquivo_base64, data.arquivo_nome, data.arquivo_tipo);
    if (up.error) return { error: up.error };
    const colUrl = headers.indexOf('anexo_url') + 1;
    const colNome = headers.indexOf('anexo_nome') + 1;
    if (colUrl > 0) sheet.getRange(rowNum, colUrl).setValue(up.anexo_url);
    if (colNome > 0) sheet.getRange(rowNum, colNome).setValue(up.anexo_nome);
  }
  const colAtu = headers.indexOf('data_atualizacao') + 1;
  if (colAtu > 0) sheet.getRange(rowNum, colAtu).setValue(new Date().toISOString());
  invalidarCacheAba('eventos');

  registrarLog(usuario.email, 'atualizar_evento', 'ID: ' + data.id);
  return { success: true, message: 'Evento atualizado' };
}

// Atualiza uma serie recorrente (ou converte um evento avulso em serie / uma serie em avulso).
// So substitui ocorrencias futuras (hoje em diante); ocorrencias passadas ficam intocadas.
function handleAtualizarRecorrencia(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };

  const aba = lerAba('eventos');
  const rows = aba.rows;
  const sheet = aba.sheet;

  const grupo = data.recorrencia_grupo;
  var idxAlvo = [];
  if (grupo) {
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].recorrencia_grupo) === String(grupo)) idxAlvo.push(i);
    }
  } else {
    for (var j = 0; j < rows.length; j++) {
      if (String(rows[j].id) === String(data.id)) { idxAlvo.push(j); break; }
    }
  }
  if (!idxAlvo.length) return { error: 'Evento não encontrado' };

  const referencia = rows[idxAlvo[0]];
  if (usuario.tipo !== 'admin' && referencia.orgao !== usuario.orgao) return { error: 'Acesso negado' };

  const titulo = data.titulo, local = data.local, responsavel = data.responsavel, telefone = data.telefone, observacao = data.observacao;
  if (!titulo || !observacao) return { error: 'Título e observação são obrigatórios' };

  const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);

  var datasEvento = [];
  var recorrenciaTipo = '';
  var recorrenciaGrupo = grupo || '';
  const rec = data.recorrencia;

  if (rec && rec.tipo && rec.dataInicio && rec.dataFim && rec.dias && rec.dias.length) {
    const validado = validarRecorrencia(rec, tz);
    if (validado.erro) return { error: validado.erro };
    datasEvento = validado.datas;
    recorrenciaTipo = rec.tipo;
    if (!recorrenciaGrupo) recorrenciaGrupo = 'rec_' + new Date().getTime();
  } else {
    if (!data.data_evento) return { error: 'Data é obrigatória' };
    if (new Date(data.data_evento) < new Date()) return { error: 'Não é permitido usar data ou hora retroativa.' };
    datasEvento = [data.data_evento];
    recorrenciaTipo = '';
    recorrenciaGrupo = '';
  }

  var anexo_url = referencia.anexo_url || '';
  var anexo_nome = referencia.anexo_nome || '';
  if (data.arquivo_base64 && data.arquivo_nome) {
    const up = salvarAnexoDrive(data.arquivo_base64, data.arquivo_nome, data.arquivo_tipo);
    if (up.error) return { error: up.error };
    anexo_url = up.anexo_url; anexo_nome = up.anexo_nome;
  }
  const orgaoEvento = referencia.orgao;
  const isPref = referencia.is_prefeito;
  const publicadoPor = referencia.publicado_por;
  const publicadoEmail = referencia.email_publicado;
  const now = new Date().toISOString();

  // Remove so as ocorrencias futuras (hoje em diante); as passadas ficam como historico.
  // Usa normalizarDataEvento pois linhas antigas (gravadas antes de um deploy correto) podem ter
  // data_evento como objeto Date em vez de texto — String(Date).substring(0,10) gera lixo tipo
  // "Sat Jul 18" em vez de "2026-07-18", quebrando a comparacao silenciosamente (sempre "nao e futuro").
  var idxRemover = idxAlvo.filter(function(i) {
    var dv = String(normalizarDataEvento(rows[i].data_evento, tz) || '').substring(0, 10);
    var d = new Date(dv + 'T00:00:00');
    return d >= hoje;
  });
  idxRemover.sort(function(a, b) { return b - a; });
  for (var k = 0; k < idxRemover.length; k++) sheet.deleteRow(idxRemover[k] + 2);

  garantirTextoDataEvento(sheet);
  var idsGerados = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    for (var m = 0; m < datasEvento.length; m++) {
      const id = proximoId('eventos');
      sheet.appendRow([
        id, titulo, datasEvento[m], local || '', responsavel || '',
        telefone || '', observacao, anexo_url, anexo_nome,
        orgaoEvento, isPref, publicadoPor, publicadoEmail, now, now, 'ativo',
        recorrenciaTipo, recorrenciaGrupo
      ]);
      idsGerados.push(id);
    }
  } finally {
    lock.releaseLock();
  }
  invalidarCacheAba('eventos');

  registrarLog(usuario.email, 'atualizar_recorrencia', titulo + ' (' + idsGerados.length + ' ocorrência(s) futura(s))');
  return { success: true, total: idsGerados.length, ids: idsGerados };
}

function handleExcluirEvento(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };

  const aba = lerAba('eventos');
  const rows = aba.rows;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.id)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Evento não encontrado' };

  if (usuario.tipo !== 'admin' && rows[idx].orgao !== usuario.orgao) return { error: 'Acesso negado' };

  sheet.deleteRow(idx + 2);
  invalidarCacheAba('eventos');
  registrarLog(usuario.email, 'excluir_evento', 'ID: ' + data.id);
  return { success: true, message: 'Evento excluído' };
}

// Exclui toda uma serie recorrente — so as ocorrencias futuras (hoje em diante), preservando o historico passado
function handleExcluirSerieRecorrente(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };

  const grupo = data.recorrencia_grupo;
  if (!grupo) return { error: 'Grupo de recorrência não informado' };

  const aba = lerAba('eventos');
  const rows = aba.rows;
  const sheet = aba.sheet;
  var idxAlvo = [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].recorrencia_grupo) === String(grupo)) idxAlvo.push(i);
  }
  if (!idxAlvo.length) return { error: 'Nenhum evento encontrado para essa série' };

  const referencia = rows[idxAlvo[0]];
  if (usuario.tipo !== 'admin' && referencia.orgao !== usuario.orgao) return { error: 'Acesso negado' };

  const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  var idxRemover = idxAlvo.filter(function(i) {
    var dv = String(normalizarDataEvento(rows[i].data_evento, tz) || '').substring(0, 10);
    var d = new Date(dv + 'T00:00:00');
    return d >= hoje;
  });
  if (!idxRemover.length) return { error: 'Não há ocorrências futuras nessa série para excluir.' };
  idxRemover.sort(function(a, b) { return b - a; });
  for (var k = 0; k < idxRemover.length; k++) sheet.deleteRow(idxRemover[k] + 2);
  invalidarCacheAba('eventos');

  registrarLog(usuario.email, 'excluir_serie_recorrente', 'grupo ' + grupo + ' (' + idxRemover.length + ' ocorrência(s))');
  return { success: true, total: idxRemover.length };
}

// Senha nao pode conter espaco (usuario deve usar caractere especial no lugar)
function temEspaco(str) {
  return /\s/.test(String(str || ''));
}

// Checa duplicidade de login/e-mail de forma normalizada (case/espaco) e cruzada:
// o login novo nao pode coincidir com um login OU e-mail existente, e vice-versa.
function normalizarIdentificador(v) {
  return String(v || '').trim().toLowerCase();
}
function loginOuEmailEmUso(login, email, rows) {
  const loginNorm = normalizarIdentificador(login);
  const emailNorm = normalizarIdentificador(email);
  for (var i = 0; i < rows.length; i++) {
    const rLogin = normalizarIdentificador(rows[i].login);
    const rEmail = normalizarIdentificador(rows[i].email);
    if (rLogin === loginNorm || rEmail === loginNorm || rLogin === emailNorm || rEmail === emailNorm) return true;
  }
  return false;
}

// Usuarios
function handleGetUsuarios(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const rows = lerAba('usuarios').rows;
  return rows.map(function(u) { const c = Object.assign({}, u); delete c.senha; return c; });
}

// So contagem agregada por orgao (sem nome/login/email) — liberado pra SECOM alem do
// admin, pra alimentar o export do dashboard sem expor dado pessoal de outro orgao.
function handleGetContagemUsuarios(data) {
  const usuario = verificarToken(data.token);
  if (!usuario) return { error: 'Não autorizado' };
  const podeVer = usuario.tipo === 'admin' || getSiglaOrgao(usuario.orgao) === 'SECOM';
  if (!podeVer) return { error: 'Acesso negado' };

  const rows = lerAba('usuarios').rows;
  const contagem = {};
  rows.forEach(function(u) {
    if (!u.orgao) return;
    contagem[u.orgao] = (contagem[u.orgao] || 0) + 1;
  });
  return contagem;
}

function handleCriarUsuario(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const login = String(data.login || '').trim();
  const nome = data.nome;
  const email = String(data.email || '').trim();
  const senha = data.senha;
  const tipo = data.tipo;
  const orgao = data.orgao;

  if (!login || !nome || !email || !senha || !tipo) return { error: 'Todos os campos são obrigatórios' };
  if (temEspaco(senha)) return { error: 'Espaço não é permitido na senha. Use um caractere especial no lugar.' };

  const tiposValidos = ['admin', 'prefeito', 'orgao'];
  if (tiposValidos.indexOf(tipo) === -1) return { error: 'Tipo inválido. Use: admin, prefeito ou orgao' };
  if (tipo === 'orgao' && !orgao) return { error: 'Órgão é obrigatório para sessão de órgão' };

  const aba = lerAba('usuarios');
  const rows = aba.rows;
  const sheet = aba.sheet;
  if (loginOuEmailEmUso(login, email, rows)) return { error: 'Login ou e-mail já existe' };

  var count = 0;
  for (var j = 0; j < rows.length; j++) {
    if (tipo === 'admin' && rows[j].tipo === 'admin') count++;
    else if (tipo === 'prefeito' && rows[j].tipo === 'prefeito') count++;
    else if (tipo === 'orgao' && rows[j].tipo === 'orgao' && rows[j].orgao === orgao) count++;
  }

  const limite = tipo === 'orgao' ? limitePara(orgao) : LIMITE_POR_SESSAO;
  if (count >= limite) {
    return { error: 'Limite de ' + limite + ' usuários atingido. Entre em contato com a SECOM.' };
  }

  const orgaoFinal = tipo === 'prefeito' ? 'Gabinete do Prefeito' : (tipo === 'admin' ? '' : orgao);

  sheet.appendRow([
    proximoId('usuarios'), login, nome, email, senha,
    tipo, orgaoFinal,
    tipo === 'prefeito' ? 'TRUE' : 'FALSE',
    'TRUE', new Date().toISOString()
  ]);

  registrarLog(admin.email, 'criar_usuario', login);
  return { success: true, message: 'Usuário criado' };
}

function handleResetarSenha(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };
  if (temEspaco(data.novaSenha)) return { error: 'Espaço não é permitido na senha. Use um caractere especial no lugar.' };

  const aba = lerAba('usuarios');
  const rows = aba.rows;
  const headers = aba.headers;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.usuarioId)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Usuário não encontrado' };

  const col = headers.indexOf('senha') + 1;
  sheet.getRange(idx + 2, col).setValue(data.novaSenha);
  registrarLog(admin.email, 'resetar_senha', 'ID: ' + data.usuarioId);
  return { success: true, message: 'Senha alterada' };
}

function handleExcluirUsuario(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const aba = lerAba('usuarios');
  const rows = aba.rows;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.id)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Usuário não encontrado' };

  sheet.deleteRow(idx + 2);
  registrarLog(admin.email, 'excluir_usuario', 'ID: ' + data.id);
  return { success: true, message: 'Usuário excluído' };
}

// Garante que a aba 'solicitacoes' exista antes de gravar nela — usado tanto
// por handleSolicitarAcesso quanto handleRecuperarSenha.
function getOuCriarSolicitacoesSheet() {
  var sheet = getSheet('solicitacoes');
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('solicitacoes');
    sheet.appendRow(['id','nome','email','login','telefone','orgao','justificativa','status','tipoSolicitacao','data_solicitacao','senha']);
  }
  return sheet;
}

// Solicitar Acesso (publico) — SEMPRE fica PENDENTE ate um admin aprovar
// (aprovarSolicitacao, index.html, ja chama criarUsuario — devidamente
// autenticado — na hora da aprovacao).
//
// Bug de segurança real corrigido (2026-09-16, usuário: "Invadiram nossa
// agenda e fizeram o login como cadastro teste"): esta função ANTES criava
// a conta em 'usuarios' na hora, direto, sem nenhuma aprovação — qualquer
// pessoa com o link público do sistema (a própria tela de login expõe
// "Criar Conta") conseguia se auto-cadastrar com QUALQUER nome/e-mail/
// telefone e ganhar acesso imediato a um órgão real, sem nenhuma
// verificação de identidade. Agora só grava uma solicitação PENDENTE —
// o acesso de verdade só existe depois que um admin revisar e aprovar
// pelo painel (que já chamava `criarUsuario`, só nunca era exercitado
// porque a conta já existia antes de chegar lá). Telefone virou
// obrigatório (pedido do usuário: "priorize... o contato telefônico e um
// email real") — dá ao admin um jeito de confirmar a identidade de quem
// pediu antes de aprovar.
function handleSolicitarAcesso(data) {
  const nome = data.nome;
  const email = String(data.email || '').trim();
  const login = String(data.login || '').trim();
  const telefone = String(data.telefone || '').trim();
  const orgao = data.orgao;
  const senha = data.senha;
  if (!nome || !email || !login || !telefone || !orgao || !senha) {
    return { error: 'Nome, e-mail, telefone, login, orgao e senha sao obrigatorios' };
  }
  if (temEspaco(senha)) return { error: 'Espaço não é permitido na senha. Use um caractere especial no lugar.' };

  const rows = lerAba('usuarios').rows;
  if (loginOuEmailEmUso(login, email, rows)) return { error: 'Login ou e-mail já está em uso' };

  var count = 0;
  for (var j = 0; j < rows.length; j++) {
    if (rows[j].tipo === 'orgao' && rows[j].orgao === orgao) count++;
  }
  const limite = limitePara(orgao);
  if (count >= limite) {
    return { error: 'Limite de ' + limite + ' usuários atingido. Entre em contato com a SECOM.' };
  }

  const solicSheet = getOuCriarSolicitacoesSheet();
  solicSheet.appendRow([
    proximoId('solicitacoes'), nome, email, login,
    telefone, orgao, '',
    'pendente', 'acesso', new Date().toISOString(), senha
  ]);

  registrarLog(email, 'solicitar_acesso', login);

  return { success: true, message: 'Solicitação enviada! Um administrador vai revisar e liberar seu acesso em breve.' };
}

// Recuperar Senha (publico) — SEMPRE fica PENDENTE ate um admin confirmar
// (resetarSenhaSolicitacao, index.html, ja chama resetarSenha — devidamente
// autenticado — na hora da confirmação; o admin define a senha final, não
// necessariamente a que a pessoa digitou aqui).
//
// Bug de segurança real corrigido (2026-09-16, mesma rodada do fix acima):
// esta função ANTES trocava a senha de QUALQUER conta na hora, só com o
// login/e-mail (público, fácil de adivinhar/tentar pra logins como
// "admin"/"educacao"/"saude") — sem enviar nenhum código, sem confirmar
// e-mail, sem NADA verificando que quem pediu é o dono de verdade da
// conta. Era um sequestro de conta completo, ao alcance de qualquer
// visitante do site. Agora só registra um PEDIDO pendente — a senha só
// muda de verdade depois que um admin confirma a identidade (por telefone/
// e-mail já cadastrado, fora do sistema) e aprova pelo painel.
function handleRecuperarSenha(data) {
  const login = String(data.login || '').trim();
  const novaSenha = data.novaSenha;
  if (!login) return { error: 'Informe seu login ou e-mail' };
  if (!novaSenha) return { error: 'Informe a nova senha' };
  if (temEspaco(novaSenha)) return { error: 'Espaço não é permitido na senha. Use um caractere especial no lugar.' };

  const rows = lerAba('usuarios').rows;
  var usuario = null;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].login === login || rows[i].email === login) { usuario = rows[i]; break; }
  }
  if (!usuario) return { error: 'Usuário não encontrado. Verifique o login ou contate o administrador.' };

  const solicSheet = getOuCriarSolicitacoesSheet();
  solicSheet.appendRow([
    proximoId('solicitacoes'), usuario.nome || '', usuario.email || '', usuario.login,
    '', usuario.orgao || '', '',
    'pendente', 'recuperacao', new Date().toISOString(), novaSenha
  ]);

  registrarLog(usuario.email, 'solicitar_recuperacao_senha', 'Pedido registrado, aguardando confirmação do admin');

  return { success: true, message: 'Pedido enviado! Um administrador vai confirmar sua identidade e liberar a nova senha em breve.' };
}

// Gerenciar Solicitacoes (admin)
function handleGetSolicitacoes(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const rows = lerAba('solicitacoes').rows;
  return rows.sort(function(a, b) { return new Date(b.data_solicitacao) - new Date(a.data_solicitacao); });
}

function handleAtualizarSolicitacao(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const aba = lerAba('solicitacoes');
  const rows = aba.rows;
  const headers = aba.headers;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.id)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Solicitação não encontrada' };

  const col = headers.indexOf('status') + 1;
  if (col > 0) sheet.getRange(idx + 2, col).setValue(data.status);

  registrarLog(admin.email, 'atualizar_solicitacao', 'ID ' + data.id + ' -> ' + data.status);
  return { success: true };
}

function handleExcluirSolicitacao(data) {
  const admin = verificarToken(data.token);
  if (!admin || admin.tipo !== 'admin') return { error: 'Acesso negado' };

  const aba = lerAba('solicitacoes');
  const rows = aba.rows;
  const sheet = aba.sheet;
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(data.id)) { idx = i; break; }
  }
  if (idx === -1) return { error: 'Solicitação não encontrada' };

  sheet.deleteRow(idx + 2);
  registrarLog(admin.email, 'excluir_solicitacao', 'ID: ' + data.id);
  return { success: true, message: 'Solicitação excluída' };
}

// Primeiro Admin
function handlePrimeiroAdmin(data) {
  const login = data.login;
  const senha = data.senha;
  const nome = data.nome;
  const email = data.email;
  if (!login || !senha) return { error: 'Login e senha são obrigatórios' };

  const aba = lerAba('usuarios');
  const rows = aba.rows;
  const sheet = aba.sheet;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].tipo === 'admin') return { error: 'Já existe um administrador. Use o painel para criar novos usuários.' };
  }

  sheet.appendRow([
    1, login, nome || login, email || (login + '@prefeitura.gov.br'),
    senha, 'admin', '', 'FALSE', 'TRUE', new Date().toISOString()
  ]);

  return { success: true, message: 'Administrador "' + login + '" criado com sucesso!' };
}

// Setup inicial das abas
function criarAbas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const config = {
    usuarios:     ['id','login','nome','email','senha','tipo','orgao','is_prefeito','ativo','data_criacao'],
    eventos:      ['id','titulo','data_evento','local','responsavel','telefone','observacao','anexo_url','anexo_nome','orgao','is_prefeito','publicado_por','email_publicado','data_publicacao','data_atualizacao','status','recorrencia_tipo','recorrencia_grupo'],
    logs:         ['id','data','usuario','acao','detalhes'],
    solicitacoes: ['id','nome','email','login','telefone','orgao','justificativa','status','tipoSolicitacao','data_solicitacao','senha']
  };
  const resultado = [];
  const nomes = Object.keys(config);
  for (var i = 0; i < nomes.length; i++) {
    const nome = nomes[i];
    if (!ss.getSheetByName(nome)) {
      const s = ss.insertSheet(nome);
      s.appendRow(config[nome]);
      resultado.push('criada: ' + nome);
    } else {
      resultado.push('ja existe: ' + nome);
    }
  }
  return { success: true, abas: resultado };
}

// Adiciona colunas de recorrencia na aba eventos, se ainda nao existirem (execute manualmente uma vez)
function adicionarColunasRecorrencia() {
  const sheet = getSheet('eventos');
  if (!sheet) { Logger.log('Aba eventos nao encontrada'); return; }
  // Colunas fixas: tem que bater exatamente com a ordem em que handleCriarEvento/handleAtualizarRecorrencia
  // escrevem via appendRow (sempre 18 valores, A ate R). Nao usar getLastColumn() aqui — se sobrar
  // qualquer dado solto mais a direita na planilha, ele infla a contagem e desalinha as colunas novas
  // do lugar onde os dados de recorrencia sao realmente gravados (foi o que causou o bug de 2026-07-18).
  const COL_TIPO = 17, COL_GRUPO = 18;
  const atualTipo = sheet.getRange(1, COL_TIPO).getValue();
  const atualGrupo = sheet.getRange(1, COL_GRUPO).getValue();
  Logger.log('Antes — coluna 17: "' + atualTipo + '" | coluna 18: "' + atualGrupo + '"');
  var mudou = [];
  if (atualTipo !== 'recorrencia_tipo') { sheet.getRange(1, COL_TIPO).setValue('recorrencia_tipo'); mudou.push('coluna 17 -> recorrencia_tipo'); }
  if (atualGrupo !== 'recorrencia_grupo') { sheet.getRange(1, COL_GRUPO).setValue('recorrencia_grupo'); mudou.push('coluna 18 -> recorrencia_grupo'); }
  Logger.log(mudou.length ? ('Corrigido: ' + mudou.join(', ')) : 'Colunas 17 e 18 ja estavam corretas.');
}

// Remove cabecalhos "recorrencia_tipo"/"recorrencia_grupo" duplicados que sobraram fora das colunas
// 17/18 (bug da versao antiga de adicionarColunasRecorrencia, que usava getLastColumn() e colocou
// esses nomes em colunas erradas, tipo 19/20). Header duplicado faz lerAba() sobrescrever o valor
// certo da coluna 17/18 com o vazio da coluna errada. Execute manualmente uma vez.
function limparColunasRecorrenciaDuplicadas() {
  const sheet = getSheet('eventos');
  if (!sheet) { Logger.log('Aba eventos nao encontrada'); return; }
  const lastCol = sheet.getLastColumn();
  var limpas = [];
  for (var col = 19; col <= lastCol; col++) {
    var valor = sheet.getRange(1, col).getValue();
    if (valor === 'recorrencia_tipo' || valor === 'recorrencia_grupo') {
      sheet.getRange(1, col).clearContent();
      limpas.push('coluna ' + col + ' (' + valor + ')');
    }
  }
  Logger.log(limpas.length ? ('Cabecalhos duplicados removidos: ' + limpas.join(', ')) : 'Nenhum cabecalho duplicado encontrado a partir da coluna 19.');
}

// Autorizar Drive (execute manualmente uma vez)
function autorizarDrive() {
  const pastaNome = 'Agenda Prefeitura Anexos';
  const pastas = DriveApp.getFoldersByName(pastaNome);
  const pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(pastaNome);
  Logger.log('Drive autorizado. Pasta: ' + pasta.getName() + ' | ID: ' + pasta.getId());
}

// Testes
function testarConexao() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Planilha: ' + ss.getName());
  const abas = ss.getSheets();
  const nomes = [];
  for (var i = 0; i < abas.length; i++) { nomes.push(abas[i].getName()); }
  Logger.log('Abas: ' + nomes.join(', '));
}

function testarLogin() {
  const rows = lerAba('usuarios').rows;
  Logger.log('Usuarios: ' + rows.length);
  for (var i = 0; i < rows.length; i++) {
    Logger.log((i+1) + ': login="' + rows[i].login + '" tipo="' + rows[i].tipo + '" ativo="' + rows[i].ativo + '"');
  }
  Logger.log('Teste admin: ' + JSON.stringify(handleLogin({ usuario: 'admin', senha: 'admin123' })));
}
