/**
 * CONFIGURAÇÕES PLUGLEAD
 */
const PLUGLEAD_CONFIG = {
  url: ""
};

function registrarLog(nome, contato, dias, mensagem, tipo, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaLog = ss.getSheetByName("Log_Envios");
  
  if (!abaLog) {
    abaLog = ss.insertSheet("Log_Envios");
    abaLog.appendRow(["Data/Hora", "Nome", "Contato", "Dias", "Mensagem", "Tipo", "Status"]);
    abaLog.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  abaLog.appendRow([new Date(), nome, contato, dias, mensagem, tipo, status]);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✈ Enviar PlugLead')
      .addItem('Enviar Linha Selecionada (Single)', 'enviarApenasLinhaSelecionada')
      .addSeparator()
      .addItem('Enviar Todas as Linhas (Bulk)', 'iteraExcelEnviaWebhookUmbler')
      .addToUi();
}

function enviaHojeOuNao(dias, quinzenal, mensal){
  if(quinzenal){
    var aptoPara14dias = (dias - 17) >= 0 && (dias - 17) % 14 === 0;
    return aptoPara14dias;
  }else if(mensal){
    var aptoPara28dias = (dias - 17) >= 0 && (dias - 17) % 28 === 0;
    return aptoPara28dias;
  }else{
    return true
  }
}

/**
 * Função BULK: Itera por toda a aba enviando para o Webhook.
 */
function iteraExcelEnviaWebhookUmbler() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("TESTCASE");
  var dados = aba.getDataRange().getValues();

  var contagemProcessados = 0;
  var atualizacoesStatus = [];

  planilha.toast("Iniciando envio via PlugLead...", "Automação", 3);

  for (var i = 1; i < dados.length; i++) {
    var dadosLinha = dados[i];

    var nome           = dadosLinha[1];   // Coluna B
    var dias           = dadosLinha[4];   // Coluna E
    var mensagem       = dadosLinha[8];   // Coluna I
    var contato        = dadosLinha[9];   // Coluna J
    var statusCheckbox = dadosLinha[10]; // Coluna K
    var quinzenal      = dadosLinha[12]; // Coluna M
    var mensal         = dadosLinha[13]; // Coluna N
  
    if(enviaHojeOuNao(dias, quinzenal, mensal) == false) continue;

    if (statusCheckbox === true || !nome) {
      atualizacoesStatus.push([statusCheckbox]);
      continue;
    }

    if (!mensagem || !contato) {
      atualizacoesStatus.push([false]);
      continue;
    }

    var telefoneLimpo = contato.toString().replace(/\D/g, '');

    var payload = {
      contato: telefoneLimpo,
      nome: nome || "",
      mensagem: mensagem,
      dias: dias,
      origem: "Google Sheets Bulk"
    };

    var opcoes = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    Utilities.sleep(1000);
    try {
      var response = UrlFetchApp.fetch(PLUGLEAD_CONFIG.url, opcoes);
      var statusCode = response.getResponseCode();

      if (statusCode === 200 || statusCode === 201) {
        registrarLog(nome, telefoneLimpo, dias, mensagem, "Bulk", "✅Enviado");
        contagemProcessados++;
        atualizacoesStatus.push([true]);
      } else {
        registrarLog(nome, telefoneLimpo, dias, mensagem, "Bulk", "Erro Webhook " + statusCode);
        atualizacoesStatus.push([false]);
      }
    } catch (e) {
      console.log("Erro na linha " + (i + 1) + ": " + e);
      atualizacoesStatus.push([false]);
    }
  }

  if (atualizacoesStatus.length > 0) {
    aba.getRange(2, 11, atualizacoesStatus.length, 1).setValues(atualizacoesStatus);
  }
  planilha.toast("Finalizado! " + contagemProcessados + " enviados via Webhook ✅", "Sucesso", 10);
}

/**
 * Função SINGLE: Envio individual para PlugLead.
 */
function enviarApenasLinhaSelecionada() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var linhaAtiva = aba.getActiveCell().getRow();

  if (linhaAtiva === 1) {
    SpreadsheetApp.getUi().alert("❌ Erro: Selecione uma linha de dados.");
    return;
  }

  var dadosLinha = aba.getRange(linhaAtiva, 1, 1, 14).getValues()[0];

  var nome           = dadosLinha[1];   
  var dias           = dadosLinha[4];   
  var mensagem       = dadosLinha[8];   
  var contato        = dadosLinha[9];   
  var statusCheckbox = dadosLinha[10];  
  var quinzenal      = dadosLinha[12];  
  var mensal         = dadosLinha[13];  

  if (enviaHojeOuNao(dias, quinzenal, mensal) === false) {
    SpreadsheetApp.getUi().alert("⚠️ Aviso: Esta linha não cumpre os requisitos de data.");
    return;
  }

  if (statusCheckbox === true) {
    planilha.toast("Esta linha já consta como enviada ✅", "Aviso", 4);
    return;
  }

  if (!nome || !mensagem || !contato) {
    SpreadsheetApp.getUi().alert("❌ Erro: Dados incompletos.");
    return;
  }

  var telefoneLimpo = contato.toString().replace(/\D/g, '');

  var payload = {
    contato: telefoneLimpo,
    nome: nome || "",
    mensagem: mensagem,
    dias: dias,
    origem: "Google Sheets Single"
  };

  var opcoes = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    planilha.toast("Disparando Webhook...", "PlugLead", 2);
    var response = UrlFetchApp.fetch(PLUGLEAD_CONFIG.url, opcoes);
    var statusCode = response.getResponseCode();

    if (statusCode === 200 || statusCode === 201) {
      registrarLog(nome, telefoneLimpo, dias, mensagem, "Single", "✅Enviado");
      aba.getRange(linhaAtiva, 11).setValue(true); 
      planilha.toast("Webhook processado! ✅", "Sucesso", 5);
    } else {
      registrarLog(nome, telefoneLimpo, dias, mensagem, "Single", "Erro Webhook " + statusCode);
      SpreadsheetApp.getUi().alert("Erro no Webhook (Status " + statusCode + ")");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro técnico: " + e.toString());
  }
}
