/**
 * Instrumento Multiprofissional de Gestão de Cuidados em Cuidado Paliativo
 * Sincronização da equipe — código do Apps Script da planilha.
 *
 * Como usar: na planilha, Extensões → Apps Script, apague o que houver,
 * cole este arquivo inteiro, troque a SENHA abaixo e publique como
 * "App da Web" (Executar como: Eu · Quem pode acessar: Qualquer pessoa).
 *
 * A planilha guarda três abas, criadas automaticamente:
 *  - Pacientes: o cadastro fixo de cada paciente
 *  - Visitas: cada visita salva (dados completos, usados pelo app)
 *  - Avaliações (análise): uma linha por visita, uma coluna por campo
 */

// Troque pela senha da equipe. Use a MESMA senha no app, em cada aparelho.
const SENHA = 'TROQUE-ESTA-SENHA';

const ABA_PACIENTES = 'Pacientes';
const ABA_VISITAS = 'Visitas';
const ABA_ANALISE = 'Avaliações (análise)';
const CAB_PACIENTES = ['id', 'iniciais', 'nascimento', 'prontuario', 'atualizado_em', 'excluido', 'cadastro_json', 'historico_cadastro_json'];
const CAB_VISITAS = ['chave', 'paciente_id', 'iniciais', 'data_visita', 'preenchido_por', 'categoria', 'ts', 'dados_json'];
const CAB_ANALISE = ['chave', 'paciente_iniciais', 'nascimento', 'prontuario', 'data_visita', 'preenchido_por', 'categoria'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.senha !== SENHA) return saida_({ ok: false, erro: 'senha' });
  if (p.acao === 'listar') return saida_({ ok: true, pacientes: listar_(), excluidos: excluidos_(), agora: Date.now() });
  return saida_({ ok: true });
}

function doPost(e) {
  let corpo;
  try { corpo = JSON.parse(e.postData.contents); } catch (err) { return saida_({ ok: false, erro: 'formato' }); }
  if (!corpo || corpo.senha !== SENHA) return saida_({ ok: false, erro: 'senha' });
  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    if (corpo.acao === 'salvar') gravar_(corpo.paciente, corpo.visita);
    else if (corpo.acao === 'lote') (corpo.itens || []).forEach(function (i) { gravar_(i.paciente, i.visita); });
    else if (corpo.acao === 'excluir') excluir_(corpo.id);
    else return saida_({ ok: false, erro: 'acao' });
    SpreadsheetApp.flush();
    return saida_({ ok: true });
  } finally {
    trava.releaseLock();
  }
}

function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function aba_(nome, cab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(nome);
  if (!s) {
    s = ss.insertSheet(nome);
    s.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function linhas_(s) {
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, s.getLastColumn()).getValues();
}

function dataBr_(ts) {
  return Utilities.formatDate(new Date(Number(ts)), 'America/Araguaina', 'dd/MM/yyyy HH:mm');
}

function texto_(v) {
  if (Array.isArray(v)) return v.join('; ');
  return v == null ? '' : String(v);
}

function gravar_(pac, vis) {
  if (!pac || !pac.id) return;
  const cad = pac.cadastro || {};
  const sp = aba_(ABA_PACIENTES, CAB_PACIENTES);
  const pacs = linhas_(sp);
  let i = -1;
  for (let k = 0; k < pacs.length; k++) if (String(pacs[k][0]) === String(pac.id)) { i = k; break; }
  const tsVis = vis && vis.ts ? Number(vis.ts) : Date.now();
  // Cadastro: vale o da visita mais recente.
  if (i === -1) {
    sp.appendRow([pac.id, "'" + (cad.nome || pac.nome || ''), "'" + (cad.nascimento || ''), "'" + (cad.prontuario || pac.prontuario || ''), tsVis, '', JSON.stringify(cad), JSON.stringify(pac.cadastroLog || [])]);
  } else if (tsVis >= Number(pacs[i][4] || 0)) {
    sp.getRange(i + 2, 1, 1, 8).setValues([[pac.id, "'" + (cad.nome || pac.nome || ''), "'" + (cad.nascimento || ''), "'" + (cad.prontuario || pac.prontuario || ''), tsVis, pacs[i][5], JSON.stringify(cad), JSON.stringify(pac.cadastroLog || [])]]);
  }
  if (!vis || !vis.ts) return;

  const chave = pac.id + '-' + vis.ts;
  const sv = aba_(ABA_VISITAS, CAB_VISITAS);
  const ja = sv.getLastRow() > 1 && sv.getRange(2, 1, sv.getLastRow() - 1, 1).getValues().some(function (r) { return String(r[0]) === chave; });
  if (ja) return;
  const autor = vis.autor || {};
  sv.appendRow([chave, pac.id, "'" + (cad.nome || pac.nome || ''), dataBr_(vis.ts), autor.nome || '', autor.categoria || '', Number(vis.ts), JSON.stringify(vis.d || {})]);

  // Aba de análise: uma coluna por campo; colunas novas entram no fim.
  const sa = aba_(ABA_ANALISE, CAB_ANALISE);
  let cab = sa.getRange(1, 1, 1, sa.getLastColumn()).getValues()[0].map(String);
  const d = vis.d || {};
  const novas = Object.keys(d).filter(function (k) { return cab.indexOf(k) === -1; });
  if (novas.length) {
    sa.getRange(1, cab.length + 1, 1, novas.length).setValues([novas]).setFontWeight('bold');
    cab = cab.concat(novas);
  }
  const linha = cab.map(function (c) {
    if (c === 'chave') return chave;
    if (c === 'paciente_iniciais') return "'" + (cad.nome || pac.nome || '');
    if (c === 'nascimento') return "'" + (cad.nascimento || '');
    if (c === 'prontuario') return "'" + (cad.prontuario || '');
    if (c === 'data_visita') return dataBr_(vis.ts);
    if (c === 'preenchido_por') return autor.nome || '';
    if (c === 'categoria') return autor.categoria || '';
    return "'" + texto_(d[c]);
  });
  sa.appendRow(linha);
}

function excluir_(id) {
  const sp = aba_(ABA_PACIENTES, CAB_PACIENTES);
  const pacs = linhas_(sp);
  for (let k = 0; k < pacs.length; k++) {
    if (String(pacs[k][0]) === String(id)) sp.getRange(k + 2, 6).setValue('sim');
  }
}

function excluidos_() {
  return linhas_(aba_(ABA_PACIENTES, CAB_PACIENTES)).filter(function (r) { return String(r[5]) === 'sim'; }).map(function (r) { return String(r[0]); });
}

function listar_() {
  const pacs = linhas_(aba_(ABA_PACIENTES, CAB_PACIENTES));
  const vis = linhas_(aba_(ABA_VISITAS, CAB_VISITAS));
  const porPaciente = {};
  vis.forEach(function (r) {
    let d = {};
    try { d = JSON.parse(r[7] || '{}'); } catch (err) { d = {}; }
    const id = String(r[1]);
    (porPaciente[id] = porPaciente[id] || []).push({ ts: Number(r[6]), d: d, autor: r[4] ? { nome: String(r[4]), categoria: String(r[5]) } : null });
  });
  const lista = [];
  pacs.forEach(function (r) {
    if (String(r[5]) === 'sim') return;
    let cad = {}, log = [];
    try { cad = JSON.parse(r[6] || '{}'); } catch (err) { cad = {}; }
    try { log = JSON.parse(r[7] || '[]'); } catch (err) { log = []; }
    const id = String(r[0]);
    const versoes = (porPaciente[id] || []).sort(function (a, b) { return a.ts - b.ts; });
    if (!versoes.length) return;
    lista.push({ id: id, nome: cad.nome || String(r[1]), prontuario: cad.prontuario || String(r[3] || ''), cadastro: cad, cadastroLog: log, versoes: versoes });
  });
  return lista;
}
