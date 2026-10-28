/**
 * COI - Agendamento online (versão com painel da recepção)
 *
 * Os dias abertos e os horários ocupados ficam guardados no próprio script (não dependem de
 * ler as abas da planilha). A recepção abre dias, bloqueia dias, ocupa e libera horários pelo
 * painel (painel.html). O site público só lê a lista de dias abertos e de horários ocupados.
 */
var CFG = {
  SHEET_ID: '18ibVeTEMqOg_os5mVRNNJTr-vLSnd_SH-AIvTKWGSLc', // usada só para o registro "Pedidos do site"
  TZ: 'America/Sao_Paulo',
  LOCAL: 'COI, Shopping Barroco, sala 307, Itaipuaçu',
  DIAS_A_FRENTE: 90,
  SENHA_RECEPCAO: 'Coi',           // senha do painel da recepção (troque por uma maior)
  AVISAR_POR_EMAIL: true,
  // dia da semana (0=dom ... 6=sáb). s/e = início/fim em minutos, p = intervalo, almoco = [de, até] bloqueado
  DOC: {
    1: { n: 'Dra. Maria Odette',          s: 480, e: 940, p: 10, almoco: [690, 780] },
    4: { n: 'Dr. Antonio Carlos Hallais', s: 840, e: 960, p: 15, almoco: null },
    5: { n: 'Dra. Maria Estephania',      s: 480, e: 990, p: 15, almoco: [720, 780] },
    6: { n: 'Dr. Luis Fernando',          s: 480, e: 690, p: 15, almoco: null }
  },
  TIPOS: ['Particular', 'Unimed', 'Bradesco', 'APPAI', 'SulAmérica', 'Revisão']
};

/* ---------- utilidades ---------- */
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function hm_(m) { return pad_(Math.floor(m / 60)) + ':' + pad_(m % 60); }
function minutos_(h) { var p = String(h).split(':'); return +p[0] * 60 + +p[1]; }
function normHora_(s) {
  var m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  return m ? pad_(+m[1]) + ':' + m[2] : null;
}
function diaSemana_(iso) {
  var p = iso.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12)).getUTCDay();
}
function hoje_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function agoraMin_() { return +Utilities.formatDate(new Date(), CFG.TZ, 'H') * 60 + +Utilities.formatDate(new Date(), CFG.TZ, 'm'); }
function somarDias_(iso, n) {
  var p = iso.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n, 12));
  return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());
}
function tipoPermitido_(d, t) {
  if (!CFG.DOC[d]) return false;
  if (t === 'Particular' || t === 'Revisão') return true;
  if (t === 'Unimed') return d !== 1;
  return d === 1; // Bradesco, APPAI, SulAmérica
}
function limpar_(s, max) {
  s = String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max || 120);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function resposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function dataBr_(iso) { return iso.split('-').reverse().join('/'); }

/* horários do dia (respeita almoço e o "último horário" escolhido para o dia) */
function horariosDoDia_(iso, fim) {
  var d = diaSemana_(iso), c = CFG.DOC[d], out = [];
  if (!c) return out;
  var fimMin = fim ? minutos_(fim) : c.e;
  for (var m = c.s; m <= fimMin; m += c.p) {
    if (c.almoco && m >= c.almoco[0] && m <= c.almoco[1]) continue;
    out.push(hm_(m));
  }
  return out;
}

/* ---------- armazenamento (um registro por dia) ---------- */
// dia = { aberto, fim: 'HH:MM' ou '', slots: { 'HH:MM': [ {nome, tel, tipo, uso, origem, evento} ] }, encaixes: [ ... ] }
function lerDia_(iso) {
  var s = PropertiesService.getScriptProperties().getProperty('d:' + iso);
  return s ? JSON.parse(s) : null;
}
function gravarDia_(iso, dia) {
  PropertiesService.getScriptProperties().setProperty('d:' + iso, JSON.stringify(dia));
  CacheService.getScriptCache().remove('pub');
}
function fimDoDia_(iso, dia) { return dia.fim || hm_(CFG.DOC[diaSemana_(iso)].e); }
function todosOsDias_() {
  var props = PropertiesService.getScriptProperties().getProperties(), out = {};
  Object.keys(props).forEach(function (k) {
    if (k.indexOf('d:') === 0) { try { out[k.slice(2)] = JSON.parse(props[k]); } catch (e) {} }
  });
  return out;
}
function ocupado_(dia, hora) { return !!(dia.slots[hora] && dia.slots[hora].length); }

/* ---------- leitura da planilha (complementa o painel) ---------- */
var SUF_ = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab'];
function semAcento_(t) { return String(t || '').toLowerCase().replace(/[áàâã]/g, 'a').replace(/[éê]/g, 'e').replace(/í/g, 'i').replace(/[óôõ]/g, 'o').replace(/ú/g, 'u').replace(/ç/g, 'c'); }
// abas por data (iso) para os próximos dias com atendimento; abas sem sufixo só valem depois da primeira aba com sufixo (2026 em diante)
function abasPorDia_(ss, hoje) {
  var abas = ss.getSheets(), mapa = {}, primeira = -1, lista = [];
  abas.forEach(function (a, i) {
    var m = a.getName().match(/^(\d{1,2})\/(\d{1,2})(?!\d)\s*([A-Za-zÀ-ÿ]+)?/);
    if (!m) return;
    var suf = m[3] ? semAcento_(m[3]).slice(0, 3) : '';
    if (suf && SUF_.indexOf(suf) < 0) return;
    if (suf && primeira < 0) primeira = i;
    lista.push({ a: a, i: i, dd: +m[1], mm: +m[2], suf: suf });
  });
  for (var n = 0; n <= 60; n++) {
    var iso = somarDias_(hoje, n), wd = diaSemana_(iso);
    if (!CFG.DOC[wd]) continue;
    var dd = +iso.slice(8, 10), mm = +iso.slice(5, 7);
    lista.forEach(function (t) {
      if (t.dd !== dd || t.mm !== mm) return;
      if (t.suf ? SUF_.indexOf(t.suf) === wd : (primeira >= 0 && t.i > primeira)) mapa[iso] = t.a;
    });
  }
  return mapa;
}
// lê uma aba: devolve {horas: [horários ocupados], fim: 'HH:MM' do último horário antes do marcador (ou '')}
function lerAba_(sh, iso) {
  var ult = sh.getLastRow(), out = { horas: [], fim: '' };
  if (ult < 6) return out;
  var v = sh.getRange(6, 1, ult - 5, 11).getDisplayValues(), doc = CFG.DOC[diaSemana_(iso)].n, ultimaHora = '';
  var grade = horariosDoDia_(iso);
  for (var i = 0; i < v.length; i++) {
    var l = v[i], txt = semAcento_(l[0] + ' ' + l[2]);
    if (txt.indexOf('nao agendar') >= 0) { out.fim = ultimaHora; break; }
    var h = normHora_(l[0]), nome = String(l[2] || '').trim();
    if (h) ultimaHora = h;
    if (!h || !nome || /almo/i.test(nome)) continue;
    var k = semAcento_(l[10]);
    if (k && /(odette|hallais|estephania|luis|fernando|antonio|maria)/.test(k) && k.indexOf(semAcento_(doc.split(' ')[1])) < 0 && k.indexOf(semAcento_(doc.split(' ').slice(-1)[0])) < 0) continue; // linha de outro médico
    if (grade.indexOf(h) >= 0 && out.horas.indexOf(h) < 0) out.horas.push(h);
  }
  return out;
}
// pacientes de um dia, como estão na aba (para o painel e os lembretes)
function pacientesAba_(iso) {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID), sh = abasPorDia_(ss, iso)[iso];
    if (!sh || sh.getLastRow() < 6) return out;
    var v = sh.getRange(6, 1, sh.getLastRow() - 5, 11).getDisplayValues(), doc = CFG.DOC[diaSemana_(iso)].n;
    for (var i = 0; i < v.length; i++) {
      var l = v[i];
      if (semAcento_(l[0] + ' ' + l[2]).indexOf('nao agendar') >= 0) break;
      var h = normHora_(l[0]), nome = String(l[2] || '').trim();
      if (!h || !nome || /almo/i.test(nome)) continue;
      var k = semAcento_(l[10]);
      if (k && /(odette|hallais|estephania|luis|fernando|antonio|maria)/.test(k) && k.indexOf(semAcento_(doc.split(' ')[1])) < 0 && k.indexOf(semAcento_(doc.split(' ').slice(-1)[0])) < 0) continue;
      out.push({ nome: nome, tel: String(l[4] || '').trim(), tipo: String(l[8] || l[6] || '').trim(), hora: h, origem: 'planilha' });
    }
  } catch (e) {}
  return out;
}
function planilhaDias_(hoje) {
  var res = {};
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID), mapa = abasPorDia_(ss, hoje);
    Object.keys(mapa).forEach(function (iso) { try { res[iso] = lerAba_(mapa[iso], iso); } catch (e) {} });
  } catch (e) {}
  return res;
}
// grava o agendamento na aba do dia (se existir): usa a linha do horário se estiver vazia; senão insere uma linha
function escreverNaAba_(iso, hora, reg) {
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID), sh = abasPorDia_(ss, hoje_())[iso];
    if (!sh) return;
    var ult = sh.getLastRow(); if (ult < 6) return;
    var a = sh.getRange(6, 1, ult - 5, 3).getDisplayValues(), linha = -1, ultimaDoHorario = -1;
    for (var i = 0; i < a.length; i++) {
      if (semAcento_(a[i][0] + a[i][2]).indexOf('nao agendar') >= 0) break;
      if (normHora_(a[i][0]) === hora) { ultimaDoHorario = i + 6; if (!String(a[i][2]).trim() && linha < 0) linha = i + 6; }
    }
    if (linha < 0) { if (ultimaDoHorario < 0) return; sh.insertRowAfter(ultimaDoHorario); linha = ultimaDoHorario + 1; sh.getRange(linha, 1).setValue(hora); }
    sh.getRange(linha, 3).setValue(reg.nome);
    sh.getRange(linha, 5).setValue(reg.tel);
    sh.getRange(linha, 7).setValue(reg.tipo === 'Revisão' ? 'Revisão' : 'Consulta');
    sh.getRange(linha, 9).setValue(reg.tipo === 'Revisão' ? 'REVISÃO' : reg.tipo);
  } catch (x) {}
}

/* ---------- GET: dias abertos e horários ocupados (sem nomes) ---------- */
function doGet(e) {
  try {
    var cache = CacheService.getScriptCache(), guardado = cache.get('pub');
    if (guardado) return ContentService.createTextOutput(guardado).setMimeType(ContentService.MimeType.JSON);
    var t0 = Date.now(), hoje = hoje_(), todos = todosOsDias_(), pl = planilhaDias_(hoje), abertos = [], ocupados = {}, fins = {}, datas = {};
    Object.keys(todos).forEach(function (k) { datas[k] = 1; });
    Object.keys(pl).forEach(function (k) { datas[k] = 1; });
    Object.keys(datas).sort().forEach(function (iso) {
      var d = todos[iso], p = pl[iso];
      if (iso < hoje || !CFG.DOC[diaSemana_(iso)]) return;
      if (d ? !d.aberto : !p) return;   // bloqueado no painel, ou sem registro e sem aba na planilha
      abertos.push(iso);
      fins[iso] = (d && d.fim) || (p && p.fim) || hm_(CFG.DOC[diaSemana_(iso)].e);
      var lista = d ? Object.keys(d.slots || {}).filter(function (h) { return d.slots[h].length; }) : [];
      if (p) p.horas.forEach(function (h) { if (lista.indexOf(h) < 0) lista.push(h); });
      if (lista.length) ocupados[iso] = lista;
    });
    var json = JSON.stringify({ ok: true, abertos: abertos, fins: fins, ocupados: ocupados, ms: Date.now() - t0, leitura: Object.keys(pl).length + ' abas lidas' });
    cache.put('pub', json, 60);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return resposta_({ ok: false, erro: 'Não foi possível ler a agenda agora.', detalhe: String(err && err.message || err).slice(0, 200) });
  }
}

/* ---------- POST ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (p.site) return resposta_({ ok: true }); // campo-isca: robô
    if (p.acao === 'admin') return admin_(p);
    return agendar_(p);
  } catch (err) {
    return resposta_({ ok: false, erro: 'Não foi possível concluir agora. Tente de novo ou fale com a recepção.', detalhe: String(err && err.message || err).slice(0, 200) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* agendamento feito pelo paciente, no site */
function agendar_(p) {
  var nome = limpar_(p.nome, 80), tel = limpar_(p.tel, 25), tipo = limpar_(p.tipo, 20),
      uso = limpar_(p.uso, 30), diaIso = String(p.dia || ''), hora = String(p.hora || '');
  if (nome.length < 3) return resposta_({ ok: false, erro: 'Digite seu nome completo.' });
  if (tel.replace(/\D/g, '').length < 10) return resposta_({ ok: false, erro: 'Digite seu telefone com DDD.' });
  if (CFG.TIPOS.indexOf(tipo) < 0) return resposta_({ ok: false, erro: 'Tipo de atendimento inválido.' });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(diaIso)) return resposta_({ ok: false, erro: 'Dia inválido.' });
  var hoje = hoje_();
  if (diaIso < hoje || diaIso > somarDias_(hoje, CFG.DIAS_A_FRENTE)) return resposta_({ ok: false, erro: 'Escolha um dia dentro dos próximos ' + CFG.DIAS_A_FRENTE + ' dias.' });
  var d = diaSemana_(diaIso);
  if (!CFG.DOC[d]) return resposta_({ ok: false, erro: 'Não há atendimento neste dia.' });
  if (!tipoPermitido_(d, tipo)) return resposta_({ ok: false, erro: tipo + ' não é atendido neste dia. Escolha outro dia.' });
  var c = CFG.DOC[d], dia = lerDia_(diaIso), pl = planilhaDias_(hoje)[diaIso];
  if (!dia && pl) dia = { aberto: true, fim: pl.fim || '', slots: {}, encaixes: [] };
  if (!dia || !dia.aberto) return resposta_({ ok: false, erro: 'Não há atendimento neste dia. Escolha outro dia.', codigo: 'ocupado' });
  var reg = { nome: nome, tel: tel, tipo: tipo, uso: uso, origem: 'site' };

  if (hora === 'Encaixe') {
    dia.encaixes.push(reg);
    gravarDia_(diaIso, dia);
    registrar_(nome, tel, tipo, uso, c.n, diaIso, 'Encaixe', 'Encaixe - aguardando vaga');
    avisar_(nome, tel, tipo, c.n, diaIso, 'ENCAIXE');
    return resposta_({ ok: true, encaixe: true });
  }

  hora = normHora_(hora);
  if (!hora || horariosDoDia_(diaIso, dia.fim).indexOf(hora) < 0) return resposta_({ ok: false, erro: 'Horário inválido.' });
  if (diaIso === hoje && minutos_(hora) <= agoraMin_()) return resposta_({ ok: false, erro: 'Este horário já passou.', codigo: 'ocupado' });
  if (ocupado_(dia, hora) || (pl && pl.horas.indexOf(hora) >= 0)) return resposta_({ ok: false, erro: 'Este horário acabou de ser ocupado. Escolha outro.', codigo: 'ocupado' });

  reg.evento = criarEvento_(diaIso, hora, c, reg);
  dia.slots[hora] = [reg];
  gravarDia_(diaIso, dia);
  escreverNaAba_(diaIso, hora, reg);
  registrar_(nome, tel, tipo, uso, c.n, diaIso, hora, 'Agendado pelo site');
  avisar_(nome, tel, tipo, c.n, diaIso, hora);
  return resposta_({ ok: true });
}

function criarEvento_(diaIso, hora, c, reg) {
  try {
    var ini = Utilities.parseDate(diaIso + ' ' + hora, CFG.TZ, 'yyyy-MM-dd HH:mm');
    var ev = CalendarApp.getDefaultCalendar().createEvent(reg.nome + ' - ' + c.n + ' - ' + reg.tipo, ini, new Date(ini.getTime() + c.p * 60000), {
      location: CFG.LOCAL,
      description: 'Telefone: ' + reg.tel + '\nAtendimento: ' + reg.tipo + (reg.uso ? '\nUso de óculos/lentes: ' + reg.uso : '') + '\nOrigem: ' + (reg.origem === 'site' ? 'site' : 'recepção')
    });
    ev.removeAllReminders();
    return ev.getId();
  } catch (x) { return ''; }
}
function apagarEvento_(id) {
  if (!id) return;
  try { var ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (x) {}
}

/* ---------- painel da recepção ---------- */
function admin_(p) {
  var cache = CacheService.getScriptCache(), falhas = +(cache.get('falhas-senha') || 0);
  if (falhas >= 5) return resposta_({ ok: false, erro: 'Muitas tentativas erradas. Aguarde 10 minutos.' });
  if (String(p.senha || '') !== CFG.SENHA_RECEPCAO) {
    cache.put('falhas-senha', String(falhas + 1), 600);
    return resposta_({ ok: false, erro: 'Senha incorreta.', codigo: 'senha' });
  }
  cache.remove('falhas-senha');
  var op = p.op, hoje = hoje_(), diaIso = String(p.dia || '');

  if (op === 'resumo') {
    var todos = todosOsDias_(), lista = [];
    Object.keys(todos).sort().forEach(function (iso) {
      if (iso < somarDias_(hoje, -1) || !CFG.DOC[diaSemana_(iso)]) return;
      var d = todos[iso], n = 0;
      Object.keys(d.slots).forEach(function (h) { n += d.slots[h].length; });
      lista.push({ dia: iso, medico: CFG.DOC[diaSemana_(iso)].n, aberto: !!d.aberto, pacientes: n, encaixes: d.encaixes.length });
    });
    var pl = planilhaDias_(somarDias_(hoje, 0));
    Object.keys(pl).forEach(function (iso) {
      if (todos[iso]) return;
      lista.push({ dia: iso, medico: CFG.DOC[diaSemana_(iso)].n, aberto: true, pacientes: pl[iso].horas.length, encaixes: 0 });
    });
    lista.sort(function (a, b) { return a.dia < b.dia ? -1 : 1; });
    return resposta_({ ok: true, dias: lista, hoje: hoje });
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(diaIso)) return resposta_({ ok: false, erro: 'Escolha o dia.' });
  var wd = diaSemana_(diaIso), c = CFG.DOC[wd];
  if (!c) return resposta_({ ok: false, erro: 'Neste dia da semana não há atendimento (só segunda, quinta, sexta e sábado).' });

  if (op === 'abrir') {
    if (diaIso < hoje) return resposta_({ ok: false, erro: 'Esta data já passou.' });
    if (diaIso > somarDias_(hoje, CFG.DIAS_A_FRENTE)) return resposta_({ ok: false, erro: 'Só dá para abrir dias dentro dos próximos ' + CFG.DIAS_A_FRENTE + ' dias.' });
    var d0 = lerDia_(diaIso) || { aberto: true, fim: '', slots: {}, encaixes: [] };
    d0.aberto = true;
    var fim0 = normHora_(p.fim);
    if (fim0) d0.fim = (minutos_(fim0) === c.e) ? '' : fim0;
    gravarDia_(diaIso, d0);
    return resposta_({ ok: true });
  }

  var dia = lerDia_(diaIso), doPlanilha = [];
  if (op === 'dia') doPlanilha = pacientesAba_(diaIso);
  if (!dia && (doPlanilha.length || abasPorDia_(SpreadsheetApp.openById(CFG.SHEET_ID), diaIso)[diaIso])) dia = { aberto: true, fim: '', slots: {}, encaixes: [] };
  if (!dia) return resposta_({ ok: false, erro: 'Este dia ainda não foi aberto.' });

  if (op === 'dia') {
    var horarios = horariosDoDia_(diaIso, dia.fim).map(function (h) { return { hora: h, pacientes: dia.slots[h] || [] }; });
    Object.keys(dia.slots).forEach(function (h) {
      if (dia.slots[h].length && !horarios.some(function (x) { return x.hora === h; })) horarios.push({ hora: h, pacientes: dia.slots[h], foraDaGrade: true });
    });
    doPlanilha.forEach(function (p) {
      var x = horarios.filter(function (y) { return y.hora === p.hora; })[0];
      if (!x) { x = { hora: p.hora, pacientes: [], foraDaGrade: true }; horarios.push(x); }
      if (!x.pacientes.some(function (q) { return semAcento_(q.nome) === semAcento_(p.nome); })) x.pacientes = x.pacientes.concat([p]);
    });
    horarios.sort(function (a, b) { return a.hora < b.hora ? -1 : 1; });
    return resposta_({ ok: true, dia: diaIso, medico: c.n, intervalo: c.p, aberto: !!dia.aberto, fim: fimDoDia_(diaIso, dia), fimPadrao: hm_(c.e), horarios: horarios, encaixes: dia.encaixes });
  }

  if (op === 'fechar' || op === 'reabrir') {
    dia.aberto = (op === 'reabrir');
    gravarDia_(diaIso, dia);
    return resposta_({ ok: true });
  }

  if (op === 'fim') {
    var f = normHora_(p.fim);
    if (!f || minutos_(f) < c.s) return resposta_({ ok: false, erro: 'Horário inválido.' });
    dia.fim = (minutos_(f) === c.e) ? '' : f;
    gravarDia_(diaIso, dia);
    return resposta_({ ok: true });
  }

  if (op === 'ocupar') {
    var hora = normHora_(p.hora), nome = limpar_(p.nome, 80), tel = limpar_(p.tel, 25), tipo = limpar_(p.tipo, 20);
    if (!hora || horariosDoDia_(diaIso, dia.fim).indexOf(hora) < 0) return resposta_({ ok: false, erro: 'Horário inválido.' });
    if (nome.length < 2) return resposta_({ ok: false, erro: 'Digite o nome do paciente (ou "Bloqueado").' });
    if (ocupado_(dia, hora)) return resposta_({ ok: false, erro: 'Este horário já está ocupado.' });
    var reg = { nome: nome, tel: tel, tipo: tipo || 'Particular', uso: '', origem: 'recepcao' };
    if (tel) reg.evento = criarEvento_(diaIso, hora, c, reg);
    dia.slots[hora] = [reg];
    gravarDia_(diaIso, dia);
    return resposta_({ ok: true });
  }

  if (op === 'liberar') {
    var h = normHora_(p.hora), i = +p.i || 0;
    if (!h || !dia.slots[h] || !dia.slots[h][i]) return resposta_({ ok: false, erro: 'Horário não encontrado.' });
    var r = dia.slots[h].splice(i, 1)[0];
    apagarEvento_(r.evento);
    if (!dia.slots[h].length) delete dia.slots[h];
    gravarDia_(diaIso, dia);
    registrar_(r.nome, r.tel, r.tipo, '', c.n, diaIso, h, 'Horário liberado pela recepção');
    return resposta_({ ok: true });
  }

  if (op === 'encaixe') {
    var en = { nome: limpar_(p.nome, 80), tel: limpar_(p.tel, 25), tipo: limpar_(p.tipo, 20) || 'Particular', uso: '', origem: 'recepcao' };
    if (en.nome.length < 2) return resposta_({ ok: false, erro: 'Digite o nome do paciente.' });
    dia.encaixes.push(en);
    gravarDia_(diaIso, dia);
    return resposta_({ ok: true });
  }

  if (op === 'encaixe_remover') {
    var k = +p.i;
    if (!dia.encaixes[k]) return resposta_({ ok: false, erro: 'Encaixe não encontrado.' });
    dia.encaixes.splice(k, 1);
    gravarDia_(diaIso, dia);
    return resposta_({ ok: true });
  }

  return resposta_({ ok: false, erro: 'Operação desconhecida.' });
}

/* ---------- registro na planilha e e-mail ---------- */
function registrar_(nome, tel, tipo, uso, medico, diaIso, hora, status) {
  try {
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID), sh = ss.getSheetByName('Pedidos do site');
    if (!sh) {
      sh = ss.insertSheet('Pedidos do site');
      sh.appendRow(['Recebido em', 'Nome', 'Telefone', 'Atendimento', 'Óculos/lentes', 'Médico(a)', 'Dia', 'Horário', 'Situação']);
      sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    sh.appendRow([Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm'), nome, tel, tipo, uso, medico, dataBr_(diaIso), hora, status]);
  } catch (x) {}
}
function avisar_(nome, tel, tipo, medico, diaIso, hora) {
  if (!CFG.AVISAR_POR_EMAIL) return;
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'COI - novo agendamento: ' + nome,
      nome + '\nTelefone: ' + tel + '\n' + tipo + ' - ' + medico + '\n' + dataBr_(diaIso) + ' às ' + hora);
  } catch (x) {}
}

/** Rode UMA vez no editor: autoriza planilha, agenda e e-mail. */
function autorizar() {
  SpreadsheetApp.openById(CFG.SHEET_ID).getName();
  CalendarApp.getDefaultCalendar().getName();
  Logger.log('Autorizado. Agora rode "semente" uma vez e publique o app da Web.');
}

/**
 * Rode UMA vez no editor: importa da planilha os pacientes já marcados nos dias abaixo
 * (abas "DD/MM ..."). Não mexe em dias que já existam no painel. Edite a lista se precisar.
 */
function semente() {
  var DIAS = { '2026-10-26': '26/10', '2026-10-29': '29/10', '2026-10-30': '30/10', '2026-11-06': '06/11', '2026-11-09': '09/11' };
  var ss = SpreadsheetApp.openById(CFG.SHEET_ID), abas = ss.getSheets(), feitos = [], sem = [];
  Object.keys(DIAS).forEach(function (iso) {
    if (iso < hoje_() || lerDia_(iso)) return;
    var pref = DIAS[iso], sh = null;
    abas.forEach(function (a) { var n = a.getName(); if (n.indexOf(pref) === 0 && !/^\d{1,2}\/\d{1,2}\/?\d{4}/.test(n)) sh = a; });
    if (!sh) { sem.push(iso); return; }
    var ult = sh.getLastRow();
    if (ult < 6) { sem.push(iso); return; }
    var v = sh.getRange(6, 1, ult - 5, 11).getDisplayValues(), slots = {}, enc = [], fim = false;
    v.forEach(function (l) {
      var txt = (l[0] + ' ' + l[2]).toLowerCase();
      if (txt.indexOf('não agendar') >= 0 || txt.indexOf('nao agendar') >= 0) fim = true;
      var h = normHora_(l[0]), nome = String(l[2] || '').trim();
      if (fim || !nome || /almo/i.test(nome)) return;
      var reg = { nome: limpar_(nome, 80), tel: limpar_(l[4], 25), tipo: limpar_(l[8] || l[6] || 'Particular', 20), uso: '', origem: 'planilha' };
      if (h && horariosDoDia_(iso).indexOf(h) >= 0) { (slots[h] = slots[h] || []); if (!slots[h].length) slots[h].push(reg); else enc.push(reg); }
      else enc.push(reg);
    });
    gravarDia_(iso, { aberto: true, fim: '', slots: slots, encaixes: enc });
    feitos.push(iso + ' (' + Object.keys(slots).length + ' horários, ' + enc.length + ' para conferir como encaixe)');
  });
  Logger.log('Dias carregados: ' + (feitos.join('; ') || 'nenhum') + (sem.length ? ' | sem aba: ' + sem.join(', ') : ''));
}
