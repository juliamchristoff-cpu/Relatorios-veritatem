/**
 * Relatórios Veritatem — exporta os dados do Google Ads para o Google Sheets.
 *
 * Funciona em uma conta de cliente OU em uma conta de administrador (MCC).
 * Na MCC, cria uma planilha para cada conta de cliente.
 *
 * Cada planilha se chama "Veritatem Google Ads - <nome da conta>" e tem as abas
 * campanhas, grupos, palavras, dispositivos e info, com os dados DIA A DIA dos
 * últimos DIAS dias. Assim o Claude monta o relatório de qualquer período
 * (semana ou mês) só filtrando as datas.
 *
 * Programe para rodar todos os dias de manhã (Frequência: Diariamente).
 */

var CONFIG = {
  // Quantos dias para trás guardar (62 cobre o mês anterior inteiro + a semana atual).
  DIAS: 62,

  // Início do nome das planilhas. Não altere: é por esse nome que o Claude encontra a planilha.
  PREFIXO: 'Veritatem Google Ads - ',

  // ID da pasta do Google Drive onde as planilhas ficam (compartilhada com a equipe).
  // É o código que aparece no endereço da pasta: drive.google.com/drive/folders/<ID>
  // Deixe vazio ('') para salvar em "Meu Drive".
  PASTA_ID: '1Yqv55UJHBmK1trQbQafSyylTlv8SZOO5',

  // Só para MCC: IDs das contas que entram (ex.: ['123-456-7890']). Vazio = todas as contas ativas.
  CONTAS: []
};

var CONSULTAS = {
  campanhas:
    'SELECT segments.date, campaign.name, campaign.advertising_channel_type, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.interactions, ' +
    'metrics.conversions, metrics.all_conversions, ' +
    'metrics.top_impression_percentage, metrics.absolute_top_impression_percentage ' +
    'FROM campaign',
  grupos:
    'SELECT segments.date, campaign.name, ad_group.name, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.interactions, ' +
    'metrics.conversions, metrics.all_conversions, metrics.absolute_top_impression_percentage ' +
    'FROM ad_group',
  palavras:
    'SELECT segments.date, ad_group.name, ad_group_criterion.keyword.text, ' +
    'ad_group_criterion.keyword.match_type, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.interactions, ' +
    'metrics.conversions, metrics.all_conversions, metrics.absolute_top_impression_percentage ' +
    'FROM keyword_view',
  dispositivos:
    'SELECT segments.date, segments.device, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.conversions ' +
    'FROM campaign'
};

function main() {
  if (ehMcc()) {
    var seletor = AdsManagerApp.accounts();
    if (CONFIG.CONTAS.length) {
      seletor = seletor.withIds(CONFIG.CONTAS);
    }
    var contas = seletor.get();
    while (contas.hasNext()) {
      var conta = contas.next();
      AdsManagerApp.select(conta);
      try {
        exportarConta();
      } catch (e) {
        Logger.log('Erro na conta ' + conta.getName() + ': ' + e);
      }
    }
  } else {
    exportarConta();
  }
}

function ehMcc() {
  try {
    return typeof AdsManagerApp !== 'undefined' && AdsManagerApp.accounts() !== null;
  } catch (e) {
    return false;
  }
}

function exportarConta() {
  var conta = AdsApp.currentAccount();
  var nome = conta.getName() || conta.getCustomerId();
  var fuso = conta.getTimeZone();

  var fim = new Date();
  fim.setDate(fim.getDate() - 1);
  var inicio = new Date();
  inicio.setDate(inicio.getDate() - CONFIG.DIAS);
  var d1 = Utilities.formatDate(inicio, fuso, 'yyyy-MM-dd');
  var d2 = Utilities.formatDate(fim, fuso, 'yyyy-MM-dd');

  var planilha = abrirPlanilha(CONFIG.PREFIXO + nome);

  for (var aba in CONSULTAS) {
    var consulta = CONSULTAS[aba] +
      " WHERE segments.date BETWEEN '" + d1 + "' AND '" + d2 + "' AND metrics.impressions > 0";
    var folha = planilha.getSheetByName(aba) || planilha.insertSheet(aba);
    folha.clear();
    AdsApp.report(consulta).exportToSheet(folha);
  }

  var info = planilha.getSheetByName('info') || planilha.insertSheet('info');
  info.clear();
  info.appendRow(['conta', nome]);
  info.appendRow(['id', conta.getCustomerId()]);
  info.appendRow(['moeda', conta.getCurrencyCode()]);
  info.appendRow(['dados de', d1]);
  info.appendRow(['dados até', d2]);
  info.appendRow(['atualizado em', Utilities.formatDate(new Date(), fuso, 'yyyy-MM-dd HH:mm')]);

  // Remove a aba vazia padrão de planilhas novas.
  ['Sheet1', 'Página1', 'Planilha1'].forEach(function (n) {
    var s = planilha.getSheetByName(n);
    if (s && planilha.getSheets().length > 1) {
      planilha.deleteSheet(s);
    }
  });

  Logger.log('OK: ' + nome + ' (' + d1 + ' a ' + d2 + ') -> ' + planilha.getUrl());
}

function abrirPlanilha(nome) {
  var arquivos = DriveApp.getFilesByName(nome);
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    if (!arquivo.isTrashed()) {
      return SpreadsheetApp.open(arquivo);
    }
  }
  var planilha = SpreadsheetApp.create(nome);
  if (CONFIG.PASTA_ID) {
    var novo = DriveApp.getFileById(planilha.getId());
    novo.moveTo(DriveApp.getFolderById(CONFIG.PASTA_ID));
  }
  return planilha;
}
