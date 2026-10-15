/**
 * Controle de Atividades - backend em Google Apps Script.
 *
 * Funciona nas suas abas de turma, com 3 ou 4 linhas de cabeçalho. O layout é detectado sozinho:
 *   - Os alunos começam na primeira linha em que a coluna A é um número e a coluna B tem um nome.
 *   - A linha logo acima do primeiro aluno guarda a DATA da atividade.
 *   - A linha 1 guarda o TIPO (Caderno, Livro do Estudante...).
 *   - As linhas entre as duas guardam o TÍTULO.
 *   - Uma coluna é de ATIVIDADE quando tem FEZ / NÃO FEZ (ou F / N) nos alunos, ou uma lista suspensa
 *     com essas opções. Por isso aparecem todas as atividades da aba, de qualquer data e em qualquer posição.
 *   - Colunas "NÃO FEZ", "FEZ" e "NOTA DE PARTICIPAÇÃO" (totais) são só lidas, nunca alteradas.
 *
 * Se o script for um projeto separado (não ligado à planilha), troque o ID abaixo pelo da sua planilha.
 */

var PLANILHA_ID_PADRAO = '1U94wmZgmsc4oLqIUx4yoMTBLFhX7h116wRzdwe_JXLM';
var PRIMEIRA_COLUNA_ATIVIDADE = 3; // coluna C
var VALORES_VALIDOS = ['FEZ', 'NÃO FEZ', ''];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Controle de Atividades')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ---------- API para o aplicativo instalável (PWA) ----------
 * O PWA fica em um site próprio e fala com este script por aqui.
 * Implante como "Executar como: Eu" e acesso "Qualquer pessoa"; só quem tem o TOKEN abaixo consegue usar.
 * Para trocar o token, edite a linha abaixo e crie uma nova versão da implantação. */
var TOKEN_API = 'TROQUE_ESTE_TOKEN';
var FUNCOES_API = ['getPlanilhas', 'selecionarPlanilha', 'getTurmas', 'getTurma', 'salvarRegistro',
                   'salvarLote', 'novaAtividade', 'editarAtividade', 'novoAluno', 'novoBimestre'];

function doPost(e) {
  var saida;
  try {
    var req = JSON.parse(e.postData.contents);
    if (!req.token || req.token !== TOKEN_API) throw new Error('Token inválido.');
    if (FUNCOES_API.indexOf(req.fn) < 0) throw new Error('Função não permitida.');
    var f = { getPlanilhas: getPlanilhas, selecionarPlanilha: selecionarPlanilha, getTurmas: getTurmas, getTurma: getTurma,
              salvarRegistro: salvarRegistro, salvarLote: salvarLote, novaAtividade: novaAtividade, editarAtividade: editarAtividade,
              novoAluno: novoAluno, novoBimestre: novoBimestre }[req.fn];
    saida = { ok: true, dados: f.apply(null, req.args || []) };
  } catch (err) {
    saida = { ok: false, erro: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Planilha em uso ---------- */

function baseId_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss.getId();
  return PLANILHA_ID_PADRAO;
}

function planilha_() {
  var id = PropertiesService.getUserProperties().getProperty('PLANILHA_ATIVA');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* cai para a planilha base */ }
  }
  return SpreadsheetApp.openById(baseId_());
}

function listaSalva_() {
  try {
    return JSON.parse(PropertiesService.getUserProperties().getProperty('LISTA_PLANILHAS') || '[]');
  } catch (e) { return []; }
}

function salvarLista_(lista) {
  PropertiesService.getUserProperties().setProperty('LISTA_PLANILHAS', JSON.stringify(lista));
}

/** Lista as planilhas de bimestre conhecidas e qual está ativa. */
function getPlanilhas() {
  var base = baseId_();
  var lista = listaSalva_();
  var vistos = {};
  var saida = [];
  var ativa = planilha_().getId();
  [{ id: base, nome: null }].concat(lista).forEach(function (p) {
    if (vistos[p.id]) return;
    vistos[p.id] = true;
    try { saida.push({ id: p.id, nome: SpreadsheetApp.openById(p.id).getName() }); } catch (e) { /* sem acesso: ignora */ }
  });
  return { ativa: ativa, planilhas: saida };
}

/** Troca a planilha em uso. */
function selecionarPlanilha(id) {
  var ss = SpreadsheetApp.openById(id); // gera erro se não houver acesso
  PropertiesService.getUserProperties().setProperty('PLANILHA_ATIVA', ss.getId());
  var lista = listaSalva_();
  if (id !== baseId_() && !lista.some(function (p) { return p.id === id; })) {
    lista.push({ id: id, nome: ss.getName() });
    salvarLista_(lista);
  }
  return true;
}

/* ---------- Detecção do layout das abas ---------- */

function ehInicioDosTotais_(texto) {
  var t = String(texto).trim().toUpperCase();
  return t === 'NÃO FEZ' || t === 'NAO FEZ' || t === 'FEZ' || t.indexOf('NOTA DE PARTICIPA') === 0;
}

/**
 * Descobre onde estão os alunos e a linha da data.
 * Devolve null se a aba não parecer uma aba de turma.
 */
function layout_(aba) {
  var ultimaLinha = aba.getLastRow();
  var ultimaCol = aba.getLastColumn();
  if (ultimaLinha < 3 || ultimaCol < 3) return null;
  var topo = Math.min(ultimaLinha, 15);
  var ab = aba.getRange(1, 1, topo, 2).getValues();
  var primeira = 0;
  var numerada = true;
  var linhaNome = 0;
  for (var r = 0; r < ab.length; r++) {
    var num = String(ab[r][0]).trim();
    var nome = String(ab[r][1]).trim();
    if (!linhaNome && nome.toUpperCase().indexOf('NOME') === 0) linhaNome = r + 1;
    if (num !== '' && !isNaN(Number(num)) && nome && nome.toUpperCase().indexOf('NOME') !== 0) {
      primeira = r + 1;
      break;
    }
  }
  if (!primeira && linhaNome) {
    // aba sem numeração na coluna A: o primeiro nome abaixo do cabeçalho "NOME"
    numerada = false;
    for (var r2 = linhaNome; r2 < ab.length; r2++) {
      var n2 = String(ab[r2][1]).trim();
      if (n2 && n2.toUpperCase().indexOf('NOME') !== 0) { primeira = r2 + 1; break; }
    }
  }
  if (primeira < 3) return null;
  var linhaData = primeira - 1;
  var cab = aba.getRange(1, 1, linhaData, ultimaCol).getValues();
  // primeira coluna de totais (usada só para decidir onde criar uma atividade nova)
  var fim = ultimaCol + 1;
  achou:
  for (var c = PRIMEIRA_COLUNA_ATIVIDADE; c <= ultimaCol; c++) {
    for (var rr = 0; rr < linhaData; rr++) {
      if (ehInicioDosTotais_(cab[rr][c - 1])) { fim = c; break achou; }
    }
  }
  return { primeira: primeira, linhaData: linhaData, fim: fim, ultimaLinha: ultimaLinha, ultimaCol: ultimaCol, cab: cab, numerada: numerada };
}

/** Filtro rápido, só pelo nome da aba: ignora abas de configuração, de tarefas e auxiliares. */
function nomeDeTurma_(nome) {
  return !(/^Config_/i.test(nome) || /Tarefas/i.test(nome) || /^r[0-9a-f]{8,}-l$/i.test(nome));
}

/** Confere o conteúdo da aba (usado só no novo bimestre, onde é preciso ter certeza). */
function ehAbaDeTurma_(aba) {
  if (!nomeDeTurma_(aba.getName())) return false;
  if (aba.getLastRow() < 3 || aba.getLastColumn() < 3) return false;
  return layout_(aba) !== null;
}

/**
 * Lista as turmas (abas) disponíveis. Não lê o conteúdo das abas: assim abre na hora,
 * e o layout de cada aba é conferido quando ela é escolhida.
 */
function getTurmas() {
  return planilha_().getSheets().filter(function (aba) {
    var oculta = false;
    try { oculta = aba.isSheetHidden(); } catch (e) { oculta = false; }
    return !oculta && nomeDeTurma_(aba.getName());
  }).map(function (a) { return a.getName(); });
}

function textoData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, planilha_().getSpreadsheetTimeZone(), 'dd/MM/yyyy');
  return String(v).trim();
}

/** dd/mm/aaaa -> número aaaammdd (0 se não for uma data). Serve para ordenar. */
function ordemData_(texto) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(texto).trim());
  if (!m) return 0;
  return Number(m[3]) * 10000 + Number(m[2]) * 100 + Number(m[1]);
}

/** Converte o que está na célula (FEZ, F, NÃO FEZ, N...) para FEZ, NÃO FEZ ou vazio. */
function normalizar_(v) {
  var t = String(v).trim().toUpperCase();
  if (t === 'FEZ' || t === 'F') return 'FEZ';
  if (t === 'NÃO FEZ' || t === 'NAO FEZ' || t === 'NÃO' || t === 'NAO' || t === 'N') return 'NÃO FEZ';
  return '';
}

/** A lista suspensa tem opções do tipo FEZ / NÃO FEZ (ou F / N)? */
function dvTemFez_(dv) {
  try {
    if (dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return false;
    return dv.getCriteriaValues()[0].some(function (v) { return normalizar_(v) !== ''; });
  } catch (e) { return false; }
}

/**
 * Descobre como a coluna grava FEZ / NÃO FEZ (por exemplo "FEZ"/"NÃO FEZ" ou "F"/"N"),
 * olhando a lista suspensa da célula do primeiro aluno.
 */
function rotulos_(aba, L, col) {
  var r = { 'FEZ': 'FEZ', 'NÃO FEZ': 'NÃO FEZ' };
  try {
    var dv = aba.getRange(L.primeira, col).getDataValidation();
    if (!dv) return r;
    var tipo = dv.getCriteriaType();
    var vals = dv.getCriteriaValues();
    var lista = [];
    if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      lista = vals[0].map(String);
    } else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      lista = [].concat.apply([], vals[0].getValues()).map(String);
    }
    lista.forEach(function (item) {
      var n = normalizar_(item);
      if (n) r[n] = item;
    });
  } catch (e) { /* mantém o padrão */ }
  return r;
}

/**
 * Lê a aba inteira uma vez e devolve: valores, alunos e as colunas de atividade
 * (de qualquer posição, inclusive atividades mais antigas).
 */
function analisar_(aba, L) {
  var dados = aba.getRange(1, 1, L.ultimaLinha, L.ultimaCol).getValues();
  var alunos = [];
  for (var lin = L.primeira; lin <= L.ultimaLinha; lin++) {
    var linha = dados[lin - 1];
    var nome = String(linha[1]).replace(/\s+/g, ' ').trim();
    var num = linha[0];
    if (!nome) continue;
    if (L.numerada && (num instanceof Date || !/^\d{1,3}$/.test(String(num).trim()))) continue;
    if (!L.numerada) num = alunos.length + 1;
    alunos.push({ row: lin, num: num, nome: nome });
  }
  var nAl = L.ultimaLinha - L.primeira + 1;
  var nCol = L.ultimaCol - PRIMEIRA_COLUNA_ATIVIDADE + 1;
  var dvs = nCol > 0 ? aba.getRange(L.primeira, PRIMEIRA_COLUNA_ATIVIDADE, nAl, nCol).getDataValidations() : [];
  var cols = [];
  for (var c = PRIMEIRA_COLUNA_ATIVIDADE; c <= L.ultimaCol; c++) {
    var achou = false;
    for (var i = 0; i < alunos.length; i++) {
      if (normalizar_(dados[alunos[i].row - 1][c - 1])) { achou = true; break; }
    }
    if (!achou && c < L.fim) {
      // coluna ainda sem registros, mas com tipo / título / data no cabeçalho (atividade recém-criada)
      for (var h = 0; h < L.linhaData; h++) {
        if (String(L.cab[h][c - 1]).trim()) { achou = true; break; }
      }
    }
    if (!achou) {
      for (var j = 0; j < dvs.length; j++) {
        var dv = dvs[j][c - PRIMEIRA_COLUNA_ATIVIDADE];
        if (dv) { achou = dvTemFez_(dv); break; }
      }
    }
    if (achou) cols.push(c);
  }
  return { dados: dados, alunos: alunos, cols: cols };
}

/* ---------- Leitura ---------- */

/** Devolve todas as atividades, os alunos com o registro de cada atividade e os totais da planilha. */
function getTurma(nomeAba) {
  var aba = planilha_().getSheetByName(nomeAba);
  if (!aba) throw new Error('Aba não encontrada: ' + nomeAba);
  var L = layout_(aba);
  if (!L) throw new Error('Não encontrei a lista de alunos nesta aba.');
  var A = analisar_(aba, L);
  var dados = A.dados;

  var ltit = linhaTitulo_(aba, L);
  var atividades = A.cols.map(function (c) {
    var tipo = String(dados[0][c - 1]).trim();
    var titulo = '';
    if (ltit) titulo = String(dados[ltit - 1][c - 1]).trim();
    for (var r = 1; !titulo && r < L.linhaData - 1; r++) {
      var t = String(dados[r][c - 1]).trim();
      if (t) { titulo = t; break; }
    }
    var data = textoData_(dados[L.linhaData - 1][c - 1]);
    return { col: c, tipo: tipo, titulo: titulo, data: data, ord: ordemData_(data) };
  });

  // colunas de totais da própria planilha (NÃO FEZ, FEZ, NOTA DE PARTICIPAÇÃO): só leitura
  var ehAtiv = {};
  A.cols.forEach(function (c) { ehAtiv[c] = true; });
  var extrasCols = [];
  for (var c = PRIMEIRA_COLUNA_ATIVIDADE; c <= L.ultimaCol; c++) {
    if (ehAtiv[c]) continue;
    for (var rr = 0; rr < L.linhaData; rr++) {
      if (ehInicioDosTotais_(L.cab[rr][c - 1])) {
        extrasCols.push({ col: c, label: String(L.cab[rr][c - 1]).trim() });
        break;
      }
    }
  }

  var alunos = A.alunos.map(function (al) {
    var linha = dados[al.row - 1];
    var valores = {};
    atividades.forEach(function (a) { valores[a.col] = normalizar_(linha[a.col - 1]); });
    var extras = extrasCols.map(function (e) {
      var v = linha[e.col - 1];
      if (typeof v === 'number') return Math.round(v * 100) / 100;
      return String(v).trim();
    });
    return { row: al.row, num: al.num, nome: al.nome, valores: valores, extras: extras };
  });

  return { aba: nomeAba, atividades: atividades, alunos: alunos, extrasCols: extrasCols };
}

/* ---------- Escrita ---------- */

function exigirAba_(nomeAba) {
  var aba = planilha_().getSheetByName(nomeAba);
  if (!aba) throw new Error('Aba não encontrada.');
  var L = layout_(aba);
  if (!L) throw new Error('Layout da aba não reconhecido.');
  return { aba: aba, L: L };
}

/** A coluna pode receber FEZ / NÃO FEZ? Nunca as colunas de totais nem células com fórmula. */
function colunaGravavel_(aba, L, coluna) {
  if (coluna < PRIMEIRA_COLUNA_ATIVIDADE || coluna > L.ultimaCol) return false;
  for (var r = 0; r < L.linhaData; r++) {
    if (ehInicioDosTotais_(L.cab[r][coluna - 1])) return false;
  }
  // só bloqueia se TODAS as células dos alunos forem fórmulas (coluna calculada)
  var nAl = L.ultimaLinha - L.primeira + 1;
  var formulas = aba.getRange(L.primeira, coluna, nAl, 1).getFormulas();
  var todas = formulas.length > 0 && formulas.every(function (f) { return f[0] !== ''; });
  if (todas) return false;
  return true;
}

/** Grava FEZ / NÃO FEZ / vazio para um aluno em uma atividade. */
function salvarRegistro(nomeAba, linha, coluna, valor) {
  if (VALORES_VALIDOS.indexOf(valor) === -1) throw new Error('Valor inválido.');
  var x = exigirAba_(nomeAba);
  if (linha < x.L.primeira || linha > x.L.ultimaLinha) throw new Error('Linha fora da lista de alunos.');
  if (!colunaGravavel_(x.aba, x.L, coluna)) throw new Error('Esta coluna não pode ser alterada.');
  if (x.aba.getRange(linha, coluna).getFormula()) throw new Error('Esta célula tem fórmula e não pode ser alterada.');
  var texto = valor === '' ? '' : rotulos_(x.aba, x.L, coluna)[valor];
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    try { x.aba.getRange(linha, coluna).setValue(texto); }
    catch (e) { throw new Error('A planilha recusou a gravação (célula protegida?): ' + e.message); }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Grava vários registros de uma vez (ex.: "marcar todos como FEZ"). itens = [{linha, valor}] */
function salvarLote(nomeAba, coluna, itens) {
  var x = exigirAba_(nomeAba);
  if (!colunaGravavel_(x.aba, x.L, coluna)) throw new Error('Esta coluna não pode ser alterada.');
  var rot = rotulos_(x.aba, x.L, coluna);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    itens.forEach(function (it) {
      if (VALORES_VALIDOS.indexOf(it.valor) === -1 || it.linha < x.L.primeira || it.linha > x.L.ultimaLinha) return;
      x.aba.getRange(it.linha, coluna).setValue(it.valor === '' ? '' : rot[it.valor]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Linha do título das atividades: a linha (entre o tipo e a data) mais usada pelas atividades existentes. */
function linhaTitulo_(aba, L) {
  if (L.linhaData < 3) return 0;
  var A = analisar_(aba, L), melhor = 2, max = -1;
  for (var r = 2; r <= L.linhaData - 1; r++) {
    var n = 0;
    A.cols.forEach(function (c) { if (String(L.cab[r - 1][c - 1]).trim()) n++; });
    if (n >= max) { max = n; melhor = r; }
  }
  return melhor;
}

/** Grava o título no cabeçalho, desfazendo células mescladas se necessário, e confere se ficou gravado. */
function gravarTitulo_(aba, L, coluna, titulo) {
  var lt = linhaTitulo_(aba, L);
  if (!lt && !titulo) return;   // sem linha de título e nada a gravar
  if (!lt) throw new Error('Esta aba não tem linha de título no cabeçalho (o cabeçalho tem só tipo e data).');
  var cel0 = aba.getRange(lt, coluna);
  try {
    var mr = cel0.getMergedRanges();
    if (mr.length && mr[0].getRow() !== lt) {          // linha escondida dentro de uma mesclagem (ex.: C1:C2)
      var prox = mr[0].getLastRow() + 1;
      if (prox <= L.linhaData - 1) lt = prox;
    }
  } catch (e) { /* sem mesclagem */ }
  var cel = aba.getRange(lt, coluna);
  try { if (cel.getMergedRanges().length && cel.getMergedRanges()[0].getRow() !== lt) cel.breakApart(); } catch (e) { /* ok */ }
  if (!titulo) { cel.clearContent(); SpreadsheetApp.flush(); return; }   // título apagado
  cel.setNumberFormat('@');
  cel.setValue(String(titulo).trim());
  SpreadsheetApp.flush();
  ajustarFonteTitulo_(aba, cel, String(titulo).trim());
  if (String(cel.getValue()).trim() !== String(titulo).trim()) {
    throw new Error('A planilha não aceitou o título na célula ' + cel.getA1Notation() + ' (célula protegida ou mesclada?).');
  }
}

/** Reduz a letra do título (mínimo 5, máximo 10) até o texto caber na célula, com quebra de linha, mantendo legível. */
function ajustarFonteTitulo_(aba, cel, texto) {
  try {
    var col = cel.getColumn();
    var larg = aba.getColumnWidth(col) - 8;
    var alt = 0, r0 = cel.getRow(), r1 = r0;
    try {
      var mr = cel.getMergedRanges();
      if (mr.length) { r0 = mr[0].getRow(); r1 = mr[0].getLastRow(); }
    } catch (e) { /* sem mesclagem */ }
    for (var r = r0; r <= r1; r++) alt += aba.getRowHeight(r);
    alt -= 4;
    var palavras = texto.split(/\s+/), escolhido = 5;
    for (var t = 10; t >= 5; t--) {
      var cw = t * 0.62;                       // largura média de um caractere em negrito (px)
      var porLinha = Math.max(1, Math.floor(larg / cw));
      var maior = 0;
      palavras.forEach(function (p) { if (p.length > maior) maior = p.length; });
      if (maior > porLinha) continue;          // palavra mais longa precisa caber inteira numa linha
      var linhas = 1, atual = 0;
      palavras.forEach(function (p) {
        var need = atual ? atual + 1 + p.length : p.length;
        if (need > porLinha) { linhas++; atual = p.length; } else atual = need;
      });
      if (linhas * t * 1.3 <= alt) { escolhido = t; break; }
    }
    cel.setFontSize(escolhido).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  } catch (e) { /* ajuste de fonte é opcional */ }
}

/** Copia o formato do cabeçalho (fonte, quebra de linha, cores) de uma atividade vizinha para a coluna informada. */
function formatarCabecalho_(aba, L, coluna) {
  var A = analisar_(aba, L), ref = 0;
  A.cols.forEach(function (c) { if (c < coluna && c > ref) ref = c; });
  if (!ref) A.cols.forEach(function (c) { if (c > coluna && (!ref || c < ref)) ref = c; });
  try {
    if (ref) aba.getRange(1, ref, L.linhaData, 1).copyTo(aba.getRange(1, coluna, L.linhaData, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } catch (e) { /* mesclagens diferentes: segue sem copiar o formato */ }
  try { for (var r = 1; r <= L.linhaData; r++) aba.getRange(r, coluna).setWrap(true); } catch (e) { /* ok */ }
}

/** Cria uma nova atividade (coluna) na turma. Usa a primeira coluna livre ou insere antes dos totais. */
function novaAtividade(nomeAba, tipo, titulo, data) {
  titulo = String(titulo || '').trim();
  var x = exigirAba_(nomeAba);
  var aba = x.aba, L = x.L;
  var nAl = L.ultimaLinha - L.primeira + 1;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var alvo = 0;
    for (var c = PRIMEIRA_COLUNA_ATIVIDADE; c < L.fim && !alvo; c++) {
      var vazia = true;
      for (var r = 0; r < L.linhaData; r++) {
        if (String(L.cab[r][c - 1]).trim()) { vazia = false; break; }
      }
      if (vazia) {
        // pula colunas com lista suspensa que não é FEZ / NÃO FEZ (ex.: PRESENTE / AUSENTE)
        var dvc = aba.getRange(L.primeira, c).getDataValidation();
        if (dvc && !dvTemFez_(dvc)) vazia = false;
      }
      if (vazia) {
        // só vale se nenhum aluno já tiver registro nessa coluna
        var reg = aba.getRange(L.primeira, c, nAl, 1).getValues();
        for (var k = 0; k < reg.length; k++) {
          if (String(reg[k][0]).trim()) { vazia = false; break; }
        }
      }
      if (vazia) alvo = c;
    }
    if (!alvo) {
      aba.insertColumnBefore(L.fim);
      alvo = L.fim;
      var modelo = alvo - 1;
      if (modelo >= PRIMEIRA_COLUNA_ATIVIDADE) {
        // copia formato e lista suspensa da coluna anterior para as células dos alunos
        var origem = aba.getRange(L.primeira, modelo, nAl, 1);
        var destino = aba.getRange(L.primeira, alvo, nAl, 1);
        origem.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        origem.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        var cabOrigem = aba.getRange(1, modelo, L.linhaData, 1);
        cabOrigem.copyTo(aba.getRange(1, alvo, L.linhaData, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    }
    formatarCabecalho_(aba, L, alvo);
    aba.getRange(1, alvo).setValue(String(tipo || ''));
    gravarTitulo_(aba, L, alvo, titulo);
    var celData = aba.getRange(L.linhaData, alvo);
    celData.setNumberFormat('@');
    celData.setValue(data || '');
    // garante a lista suspensa FEZ / NÃO FEZ nas células dos alunos desta coluna
    var rngAl = aba.getRange(L.primeira, alvo, nAl, 1);
    var dvs0 = rngAl.getDataValidations();
    if (!dvs0.some(function (r) { return r[0]; })) {
      rngAl.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['FEZ', 'NÃO FEZ'], true).setAllowInvalid(true).build());
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getTurma(nomeAba);
}

/** Edita tipo, título e data de uma atividade já criada (cabeçalho da coluna). */
function editarAtividade(nomeAba, coluna, tipo, titulo, data) {
  titulo = String(titulo || '').trim();
  var x = exigirAba_(nomeAba);
  var aba = x.aba, L = x.L;
  coluna = Number(coluna);
  if (!colunaGravavel_(aba, L, coluna)) throw new Error('Esta coluna não pode ser alterada.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    formatarCabecalho_(aba, L, coluna);
    aba.getRange(1, coluna).setValue(String(tipo || ''));
    gravarTitulo_(aba, L, coluna, titulo);
    var celData = aba.getRange(L.linhaData, coluna);
    celData.setNumberFormat('@');
    celData.setValue(data || '');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getTurma(nomeAba);
}

/** Acrescenta um aluno ao fim da lista da aba (copia formato, lista suspensa e fórmulas da última linha). */
function novoAluno(nomeAba, nome) {
  nome = String(nome || '').replace(/\s+/g, ' ').trim();
  if (!nome) throw new Error('Informe o nome do aluno.');
  var x = exigirAba_(nomeAba);
  var aba = x.aba, L = x.L;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var A = analisar_(aba, L);
    if (!A.alunos.length) throw new Error('Não encontrei alunos nesta aba.');
    var existe = A.alunos.some(function (a) { return a.nome.toUpperCase() === nome.toUpperCase(); });
    if (existe) throw new Error('Já existe um aluno com esse nome nesta aba.');
    var ult = A.alunos[A.alunos.length - 1];
    aba.insertRowAfter(ult.row);
    var nova = ult.row + 1;
    // copia formato, validação e fórmulas dos totais da linha anterior
    aba.getRange(ult.row, 1, 1, L.ultimaCol).copyTo(aba.getRange(nova, 1, 1, L.ultimaCol), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    // limpa os registros FEZ / NÃO FEZ copiados
    A.cols.forEach(function (c) { aba.getRange(nova, c).clearContent(); });
    if (L.numerada) aba.getRange(nova, 1).setValue(Number(ult.num) + 1);
    aba.getRange(nova, 2).setValue(nome);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getTurma(nomeAba);
}

/**
 * DIAGNÓSTICO: rode esta função no editor do Apps Script (menu Executar) e envie o texto do "Registro de execução".
 * Edite a linha abaixo com o nome exato da aba e o número da coluna da atividade (C = 3, D = 4 ...).
 */
function diagnostico() {
  var NOME_ABA = '6ºAno A (Wanda)';   // <- troque pelo nome da aba com problema
  var COLUNA = 3;                     // <- troque pelo número da coluna da atividade
  var aba = planilha_().getSheetByName(NOME_ABA);
  if (!aba) { Logger.log('Aba não encontrada. Abas: ' + planilha_().getSheets().map(function (s) { return '[' + s.getName() + ']'; }).join(' ')); return; }
  var L = layout_(aba);
  if (!L) { Logger.log('Layout não reconhecido.'); return; }
  var A = analisar_(aba, L);
  Logger.log('Primeira linha de aluno: ' + L.primeira + ' | linha da data: ' + L.linhaData + ' | última linha: ' + L.ultimaLinha + ' | última coluna: ' + L.ultimaCol);
  Logger.log('Alunos lidos: ' + A.alunos.length + ' | colunas de atividade: ' + A.cols.join(','));
  Logger.log('Coluna ' + COLUNA + ' gravável? ' + colunaGravavel_(aba, L, COLUNA));
  var cab = []; for (var r = 0; r < L.linhaData; r++) cab.push('[' + String(L.cab[r][COLUNA - 1]) + ']');
  Logger.log('Cabeçalho da coluna: ' + cab.join(' '));
  var cel = aba.getRange(L.primeira, COLUNA);
  Logger.log('Célula do 1º aluno: valor=[' + cel.getValue() + '] fórmula=[' + cel.getFormula() + ']');
  Logger.log('Rótulos: ' + JSON.stringify(rotulos_(aba, L, COLUNA)));
  var prot = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(aba.getProtections(SpreadsheetApp.ProtectionType.SHEET));
  Logger.log('Proteções na aba: ' + prot.length + (prot.length ? ' (' + prot.map(function (p) { return p.getDescription() + ' canEdit=' + p.canEdit(); }).join('; ') + ')' : ''));
  try {
    var teste = aba.getRange(A.alunos[0].row, COLUNA), v = teste.getValue();
    teste.setValue(v); // regrava o mesmo valor, sem alterar nada
    Logger.log('Teste de gravação: OK');
  } catch (e) { Logger.log('Teste de gravação FALHOU: ' + e.message); }
}

/** DIAGNÓSTICO GERAL: rode no editor; mostra, para cada aba da lista, o que o app entende e onde falha. */
function diagnosticoTurmas() {
  var nomes = planilha_().getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return /^\s*(2|3)/.test(n); });
  nomes.forEach(function (nome) {
    try {
      var aba = planilha_().getSheetByName(nome), L = layout_(aba);
      if (!L) { Logger.log('[' + nome + '] layout não reconhecido'); return; }
      var A = analisar_(aba, L);
      Logger.log('[' + nome + '] 1ª linha aluno=' + L.primeira + ' linha data=' + L.linhaData + ' totais a partir da col ' + L.fim + ' | alunos=' + A.alunos.length + ' | colunas atividade=' + A.cols.join(','));
      var d = getTurma(nome);
      Logger.log('   getTurma OK: ' + d.atividades.length + ' atividades; título lido: ' + d.atividades.map(function (a) { return '[' + a.tipo + '|' + a.titulo + '|' + a.data + ']'; }).join(' '));
      A.cols.slice(0, 3).forEach(function (c) { Logger.log('   col ' + c + ' gravável=' + colunaGravavel_(aba, L, c) + ' rótulos=' + JSON.stringify(rotulos_(aba, L, c))); });
      var prot = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(aba.getProtections(SpreadsheetApp.ProtectionType.SHEET));
      Logger.log('   proteções=' + prot.length);
    } catch (e) { Logger.log('[' + nome + '] ERRO: ' + e.message); }
  });
}

/* ---------- Novo bimestre ---------- */

function letraColuna_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Cria o próximo bimestre: copia a planilha em uso, apaga só os registros de atividades
 * (tipo, título, data e FEZ/NÃO FEZ) das abas de turma e passa a usar a cópia.
 * Nomes, abas Config_*, abas de Tarefas e colunas de totais ficam como estão.
 */
function novoBimestre(nome) {
  nome = String(nome || '').trim();
  if (!nome) throw new Error('Informe o nome do novo bimestre.');
  var origem = planilha_();
  var arquivo = DriveApp.getFileById(origem.getId()).makeCopy(nome);
  var copia = SpreadsheetApp.openById(arquivo.getId());

  var abasLimpas = 0;
  copia.getSheets().forEach(function (aba) {
    if (!ehAbaDeTurma_(aba)) return;
    var L = layout_(aba);
    var A = analisar_(aba, L);
    if (!A.cols.length) return;
    var faixas = [];
    A.cols.forEach(function (c) {
      var letra = letraColuna_(c);
      faixas.push(letra + '1:' + letra + L.linhaData);
      faixas.push(letra + L.primeira + ':' + letra + L.ultimaLinha);
    });
    aba.getRangeList(faixas).clearContent();
    abasLimpas++;
  });
  SpreadsheetApp.flush();

  var lista = listaSalva_();
  lista.push({ id: copia.getId(), nome: copia.getName() });
  salvarLista_(lista);
  PropertiesService.getUserProperties().setProperty('PLANILHA_ATIVA', copia.getId());
  return { id: copia.getId(), nome: copia.getName(), turmas: abasLimpas };
}
