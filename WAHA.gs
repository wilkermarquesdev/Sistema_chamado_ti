/**
 * Configurações da API WAHA
 */
var WAHA_CONFIG = {
  baseUrl: "https://whatsapp.cpdevolucoes.online",
  session: "teste",
  // Se você configurou uma API Key no WAHA, preencha abaixo. Se não, deixe em branco.
  apiKey: "89f54eeb8c0549068251e01542473fb3" 
};

/**
 * Envia uma mensagem de texto via API WAHA
 * @param {string} chatId - ID da conversa (ex: "135454796058709@lid" ou "556298351405@c.us")
 * @param {string} mensagem - Texto formatado da mensagem
 */
function enviarMensagemWAHA(chatId, mensagem) {
  var url = WAHA_CONFIG.baseUrl + "/api/sendText";
  
  var payload = {
    session: WAHA_CONFIG.session,
    chatId: chatId,
    text: mensagem
  };

  var headers = {
    "Content-Type": "application/json"
  };

  if (WAHA_CONFIG.apiKey && WAHA_CONFIG.apiKey.trim() !== "") {
    headers["X-Api-Key"] = WAHA_CONFIG.apiKey;
  }

  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      Logger.log("Mensagem WhatsApp enviada com sucesso: " + responseBody);
      return "Mensagem enviada com sucesso via WhatsApp!";
    } else {
      throw new Error("Erro WAHA (" + responseCode + "): " + responseBody);
    }
  } catch (e) {
    Logger.log("Falha ao enviar via WAHA: " + e.message);
    throw new Error("Falha no envio WhatsApp: " + e.message);
  }
}

/**
 * Monta o texto do relatório e dispara o envio para a conversa especificada.
 */
function enviarRelatorioPorWhatsApp() {
  var chatIdDestino = "135454796058709@lid";
  
  var planilha = getSheet(NOME_DA_ABA_CHAMADOS);
  var ultimaLinha = planilha.getLastRow();
  
  if (ultimaLinha <= 1) {
    return enviarMensagemWAHA(chatIdDestino, "📊 *Relatório de Chamados*\n\nNenhum chamado registrado no sistema até o momento.");
  }

  var dados = planilha.getRange(2, 1, ultimaLinha - 1, 11).getValues();
  
  var hoje = new Date();
  var seteDiasAtras = new Date();
  seteDiasAtras.setDate(hoje.getDate() - 7);

  var totalSemana = 0;
  var concluidosSemana = 0;
  var pendentesSemana = 0;

  var chamadosPorSetor = {};
  var chamadosPorTecnico = {};

  for (var i = 0; i < dados.length; i++) {
    var dataInicioRaw = dados[i][5];
    var status = String(dados[i][8]).trim().toUpperCase();
    var setor = String(dados[i][2]).trim() || "Não informado";
    var tecnico = String(dados[i][9]).trim() || "Não atribuído";

    var dataInicio = null;
    if (dataInicioRaw instanceof Date) {
      dataInicio = dataInicioRaw;
    } else if (typeof dataInicioRaw === "string" && dataInicioRaw.trim() !== "") {
      var partesEspaco = dataInicioRaw.split(" ");
      var partesData = partesEspaco[0].split("/");
      if (partesData.length === 3) {
        dataInicio = new Date(partesData[2], partesData[1] - 1, partesData[0]);
      }
    }

    if (dataInicio && dataInicio >= seteDiasAtras) {
      totalSemana++;
      if (status === "CONCLUÍDO") {
        concluidosSemana++;
      } else {
        pendentesSemana++;
      }

      chamadosPorSetor[setor] = (chamadosPorSetor[setor] || 0) + 1;
      chamadosPorTecnico[tecnico] = (chamadosPorTecnico[tecnico] || 0) + 1;
    }
  }

  // Formatação em Markdown nativo do WhatsApp
  var textoRelatorio = "📊 *RELATÓRIO SEMANAL DE CHAMADOS*\n" +
    "🗓 _Período: últimos 7 dias até " + Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy") + "_\n\n" +
    "📈 *Resumo Geral*\n" +
    "• Total de chamados: *" + totalSemana + "*\n" +
    "• Concluídos: *" + concluidosSemana + "* ✅\n" +
    "• Pendentes: *" + pendentesSemana + "* ⏳\n\n" +
    "🏢 *Demanda por Setor*\n";

  for (var s in chamadosPorSetor) {
    textoRelatorio += "• " + s + ": " + chamadosPorSetor[s] + " chamado(s)\n";
  }

  textoRelatorio += "\n👨‍💻 *Atendimentos por Técnico*\n";
  for (var t in chamadosPorTecnico) {
    textoRelatorio += "• " + t + ": " + chamadosPorTecnico[t] + " chamado(s)\n";
  }

  textoRelatorio += "\n_Mensagem enviada automaticamente pelo Sistema de Chamados._";

  return enviarMensagemWAHA(chatIdDestino, textoRelatorio);
}
