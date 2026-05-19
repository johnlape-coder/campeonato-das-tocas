// ================================================================
// 🦁 CAMPEONATO DAS TOCAS — Apps Script API
// ================================================================
// Substitui TODO o conteúdo do Dashboard_codigo.gs por este ficheiro
// Depois: Implementar → Gerir implementações → lápis → Nova versão → Implementar
// ================================================================

var SS_ID = '1oiNNOe2qJ3sRDB8waaqzWWlz6Te3PIGD6ACoxfdS9Ns';

var MEMBROS_DASH = [
  "Pai Bragança", "Top 5", "Táctico", "O Mantra Ofendido",
  "Colonizador das Beiras", "Camião do Paulinho", "Preparador Físico",
  "Mister Zé", "Securitas", "Escova Progressiva", "Mini Scout",
  "Sr. Cavilhas", "Não Binário"
];

// ── Router principal ─────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter ? (e.parameter.action || '') : '';

  // API JSON para a PWA
  if (action === 'getData') {
    var data = getDashboardData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Submissão de pontos da PWA
  if (action === 'submit') {
    var result = submeterDados(e.parameter);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Dashboard HTML (fallback)
  try {
    var html = HtmlService.createHtmlOutputFromFile('Dashboard').getContent();
    html = html.replace('__DASHBOARD_DATA__', JSON.stringify(getDashboardData()));
    return HtmlService.createHtmlOutput(html)
      .setTitle('🦁 Campeonato das Tocas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    return ContentService.createTextOutput('Erro: ' + err.toString());
  }
}

// ── Submeter dados da PWA ────────────────────────────────────
function submeterDados(params) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = encontrarFolha(ss);
    if (!sheet) return { success: false, error: 'Folha de respostas não encontrada.' };

    var dataJogo = new Date(params.dataJogo);
    if (isNaN(dataJogo.getTime())) return { success: false, error: 'Data inválida.' };

    sheet.appendRow([
      new Date(),          // Timestamp
      params.nome,         // Quem és?
      dataJogo,            // Data do jogo
      params.competicao,   // Competição
      params.acoes         // O que tens a registar?
    ]);

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── Obter dados do dashboard ─────────────────────────────────
function getDashboardData() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = encontrarFolha(ss);

    var empty = MEMBROS_DASH.map(function(m) {
      return { nome: m, pontos: 0, presencas: 0, anfitrioes: 0, jogosFora: 0, drulas: 0 };
    });

    if (!sheet || sheet.getLastRow() < 2) {
      return { classificacao: empty, premios: calcPremios(empty), actividade: [], updated: formatarData(new Date()), epoca: calcularEpoca() };
    }

    var dados      = sheet.getDataRange().getValues();
    var cabecalhos = dados[0].map(function(h) { return h.toString().toLowerCase(); });
    var colNome    = encontrarColuna(cabecalhos, 'quem');
    var colData    = encontrarColuna(cabecalhos, 'data');
    var colComp    = encontrarColuna(cabecalhos, 'competi');
    var colAcoes   = encontrarColuna(cabecalhos, 'registar');

    if (colNome < 0 || colData < 0 || colAcoes < 0) {
      return { classificacao: empty, premios: calcPremios(empty), actividade: [], updated: formatarData(new Date()), epoca: calcularEpoca() };
    }

    var eventosPorMembro = {};
    MEMBROS_DASH.forEach(function(m) { eventosPorMembro[m] = []; });
    var actividadeRaw = [];

    for (var r = 1; r < dados.length; r++) {
      var linha     = dados[r];
      var nome      = linha[colNome] ? linha[colNome].toString().trim() : '';
      if (!nome || !eventosPorMembro[nome]) continue;

      var dataJogo  = new Date(linha[colData]);
      var comp      = colComp >= 0 ? linha[colComp].toString().split('—')[0].trim() : '';
      var acoes     = linha[colAcoes] ? linha[colAcoes].toString() : '';
      var timestamp = new Date(linha[0]);

      eventosPorMembro[nome].push({ timestamp: timestamp, dataJogo: dataJogo, acoes: acoes });
      actividadeRaw.push({
        nome:    nome,
        dataJogo: Utilities.formatDate(dataJogo, Session.getScriptTimeZone(), 'dd/MM'),
        dataISO:  Utilities.formatDate(dataJogo, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        comp:    comp,
        acoes:   acoes,
        ts:      timestamp.getTime()
      });
    }

    actividadeRaw.sort(function(a, b) { return b.ts - a.ts; });

    // Calcular pontos
    var resultados = [];
    MEMBROS_DASH.forEach(function(membro) {
      var eventos = eventosPorMembro[membro];
      eventos.sort(function(a, b) {
        var d = a.dataJogo - b.dataJogo;
        return d !== 0 ? d : a.timestamp - b.timestamp;
      });

      var pontos = 0, presencas = 0, anfitrioes = 0, jogosFora = 0, drulas = 0;
      eventos.forEach(function(ev) {
        var a = ev.acoes;
        if (a.indexOf('Drula') >= 0)           { pontos = 0; drulas++; return; }
        if (a.indexOf('Estive na Toca') >= 0)  { pontos += 0.5; presencas++;  }
        if (a.indexOf('Anfitrião') >= 0)       { pontos += 3;   anfitrioes++; }
        if (a.indexOf('jogo fora') >= 0)       { pontos += 2;   jogosFora++;  }
        if (a.indexOf('lampiões') >= 0 || a.indexOf('tripeiros') >= 0) { pontos -= 2; }
      });

      resultados.push({ nome: membro, pontos: Math.max(pontos, 0), presencas: presencas, anfitrioes: anfitrioes, jogosFora: jogosFora, drulas: drulas });
    });

    resultados.sort(function(a, b) { return b.pontos - a.pontos; });

    return {
      classificacao: resultados,
      premios:       calcPremios(resultados),
      actividade:    actividadeRaw.slice(0, 30),
      updated:       formatarData(new Date()),
      epoca:         calcularEpoca()
    };

  } catch(e) {
    return { error: e.toString(), classificacao: [], premios: null, actividade: [], updated: '', epoca: calcularEpoca() };
  }
}

// ── Prémios ──────────────────────────────────────────────────
function calcPremios(resultados) {
  if (!resultados || resultados.length === 0) return null;
  var porPontos  = resultados.slice().sort(function(a, b) { return b.pontos - a.pontos; });
  var porAnf     = resultados.slice().sort(function(a, b) { return b.anfitrioes - a.anfitrioes; });
  var porPres    = resultados.slice().sort(function(a, b) { return a.presencas - b.presencas; });
  return {
    campeao:     { nome: porPontos[0].nome,                   val: porPontos[0].pontos + ' pts' },
    despromoção: { nome: porPontos[porPontos.length-1].nome,  val: porPontos[porPontos.length-1].pontos + ' pts' },
    tocaDeOuro:  { nome: porAnf[0].nome,                      val: porAnf[0].anfitrioes + 'x anfitrião' },
    tocaRabolho: { nome: porPres[0].nome,                     val: porPres[0].presencas + ' presenças' }
  };
}

// ── Helpers ──────────────────────────────────────────────────
function encontrarFolha(ss) {
  var folhas = ss.getSheets();
  for (var i = 0; i < folhas.length; i++) {
    var n = folhas[i].getName();
    if (n.indexOf('Respostas') === 0 || n.indexOf('Form Responses') === 0 || n.indexOf('Form_Responses') === 0) {
      return folhas[i];
    }
  }
  return null;
}

function encontrarColuna(cabecalhos, termo) {
  for (var i = 0; i < cabecalhos.length; i++) {
    if (cabecalhos[i].indexOf(termo.toLowerCase()) >= 0) return i;
  }
  return -1;
}

function calcularEpoca() {
  var mes = new Date().getMonth() + 1;
  var ano = new Date().getFullYear();
  var ini = mes >= 7 ? ano : ano - 1;
  return String(ini).slice(2) + '/' + String(ini + 1).slice(2);
}

function formatarData(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}
