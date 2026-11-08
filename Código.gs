/**
 * SISTEMA DE GESTÃO DE CHAMADOS TI, ESTOQUE E COMPRAS
 * Integração inteligente de O.S. com fluxo de Reposição
 */

const GEMINI_MODEL = "gemini-2.5-flash";
const EMAIL_GESTOR = "wilker.marques.dev@gmail.com"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema de Gestão de Chamados & Estoque TI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterFolhaChamados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let s of sheets) {
    const nome = s.getName().trim().toLowerCase();
    if (nome === 'tabela_1' || nome === 'chamados' || nome.includes('chamado')) return s;
  }
  return sheets[0];
}

function obterOuCriarAba(nomeAba, cabecalhos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
    if (cabecalhos && cabecalhos.length > 0) sheet.appendRow(cabecalhos);
  }
  return sheet;
}

function formatarDataHora(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const pad = n => String(n).padStart(2, '0');
    return `${pad(val.getDate())}/${pad(val.getMonth() + 1)}/${val.getFullYear()} ${pad(val.getHours())}:${pad(val.getMinutes())}:${pad(val.getSeconds())}`;
  }
  return val.toString();
}

function parseDataBR(strData) {
  if (strData instanceof Date) return strData;
  if (!strData || typeof strData !== 'string') return null;
  const partes = strData.trim().split(' ');
  const dmy = partes[0].split('/');
  if (dmy.length !== 3) return null;
  let h = 0, m = 0, s = 0;
  if (partes[1]) {
    const hms = partes[1].split(':');
    h = parseInt(hms[0] || 0, 10);
    m = parseInt(hms[1] || 0, 10);
    s = parseInt(hms[2] || 0, 10);
  }
  return new Date(dmy[2], dmy[1] - 1, dmy[0], h, m, s);
}

function normalizarTexto(txt) {
  if (!txt) return "";
  return txt.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
}

// ==========================================
// 1. FLUXO DE CHAMADOS (O.S.)
// ==========================================

function obterProximaOS() {
  const sheet = obterFolhaChamados();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "OS-0001";
  
  const ultimaOS = sheet.getRange(lastRow, 1).getValue().toString();
  const match = ultimaOS.match(/OS-(\d+)/i);
  if (match) return "OS-" + String(parseInt(match[1], 10) + 1).padStart(4, '0');
  return "OS-" + String(lastRow).padStart(4, '0');
}

function salvarChamadoPendente(dados) {
  const sheet = obterFolhaChamados();
  const os = obterProximaOS();
  sheet.appendRow([os, dados.nome, dados.setor, dados.problema, "", dados.inicio || formatarDataHora(new Date()), "", "", "PENDENTE", "", "", ""]);
  return "Chamado registado com sucesso sob a numeração " + os;
}

function buscarStatusChamado(termoOS) {
  const sheet = obterFolhaChamados();
  if (sheet.getLastRow() <= 1) return null;
  const data = sheet.getDataRange().getValues();
  let numBusca = termoOS.toString().toUpperCase().trim();
  if (!numBusca.startsWith("OS-")) {
    const numLimpo = numBusca.replace(/\D/g, '');
    if (numLimpo) numBusca = "OS-" + numLimpo.padStart(4, '0');
  }

  for (let i = 1; i < data.length; i++) {
    const osLinha = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    if (osLinha === numBusca || osLinha.endsWith(termoOS.trim())) {
      return {
        os: data[i][0], nome: data[i][1], setor: data[i][2], problema: data[i][3], resolutiva: data[i][4],
        inicio: formatarDataHora(data[i][5]), fim: data[i][6] ? formatarDataHora(data[i][6]) : "",
        tempoTotal: data[i][7], status: data[i][8], tecnico: data[i][9], tipoAtendimento: data[i][10], observacao: data[i][11]
      };
    }
  }
  return null;
}

function buscarChamadosPendentes() {
  try {
    const sheet = obterFolhaChamados();
    if (sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    const resultado = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const statusBruto = (data[i][8] || "").toString(); 
      const statusLimpo = normalizarTexto(statusBruto);
      
      if (statusLimpo !== "CONCLUIDO" && statusLimpo !== "") {
        resultado.push({
          os: data[i][0].toString(), nome: (data[i][1] || "").toString(), setor: (data[i][2] || "").toString(), problema: (data[i][3] || "").toString(), resolutiva: (data[i][4] || "").toString(),
          inicio: formatarDataHora(data[i][5]), fim: data[i][6] ? formatarDataHora(data[i][6]) : "", status: statusBruto || "PENDENTE", tecnico: (data[i][9] || "").toString(), tipoAtendimento: (data[i][10] || "").toString(), observacao: (data[i][11] || "").toString()
        });
      }
    }
    return resultado.reverse();
  } catch (err) { return []; }
}

function localizarLinhaPorOS(osProcurada) {
  const sheet = obterFolhaChamados();
  if (!sheet || sheet.getLastRow() <= 1) return { sheet: sheet, rowIndex: -1 };
  const data = sheet.getDataRange().getValues();
  const osAlvo = osProcurada.toString().trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toUpperCase() === osAlvo) return { sheet: sheet, rowIndex: i + 1 };
  }
  return { sheet: sheet, rowIndex: -1 };
}

function atualizarTecnicoChamado(os, tecnico) {
  const { sheet, rowIndex } = localizarLinhaPorOS(os);
  if (rowIndex !== -1) { sheet.getRange(rowIndex, 10).setValue(tecnico); return "Técnico atualizado para " + tecnico; }
  throw new Error("Chamado não encontrado.");
}

function concluirChamado(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sheet, rowIndex } = localizarLinhaPorOS(dados.os);
    if (rowIndex === -1) throw new Error("Chamado não encontrado.");

    // Validação restrita de estoque (Se não houver peça, o código interrompe e avisa)
    if (dados.idItemEstoque && dados.qtdItemEstoque > 0) {
      darBaixaEstoqueInterno(dados.idItemEstoque, dados.qtdItemEstoque, dados.os, dados.tecnicoResp);
    }

    const dFim = new Date();
    const dataFimStr = formatarDataHora(dFim);
    const dataInicioStr = sheet.getRange(rowIndex, 6).getValue(); 
    let duracaoFormatada = "";
    if (dataInicioStr) {
      const dInicio = parseDataBR(dataInicioStr);
      if (dInicio && dFim > dInicio) {
        const diffMin = Math.round((dFim - dInicio) / (1000 * 60));
        duracaoFormatada = `${Math.floor(diffMin / 60)}h ${diffMin % 60}m`;
      }
    }

    sheet.getRange(rowIndex, 5).setValue(dados.resolutiva);        
    sheet.getRange(rowIndex, 7).setValue(dataFimStr);              
    sheet.getRange(rowIndex, 8).setValue(duracaoFormatada);        
    sheet.getRange(rowIndex, 9).setValue("CONCLUIDO");             
    sheet.getRange(rowIndex, 11).setValue(dados.tipoAtendimento);  
    
    return "Chamado " + dados.os + " concluído com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function enviarParaAssistencia(dados) {
  const { sheet, rowIndex } = localizarLinhaPorOS(dados.os);
  if (rowIndex !== -1) {
    const obsAtual = sheet.getRange(rowIndex, 12).getValue();
    const registro = `[${formatarDataHora(new Date())}] Enviado para Assistência:\n- Equipamento: ${dados.equipamento}\n- Patrimônio: ${dados.patrimonio}\n- Detalhes: ${dados.detalhes}`;
    sheet.getRange(rowIndex, 9).setValue("ASSISTENCIA TECNICA");
    sheet.getRange(rowIndex, 12).setValue((obsAtual ? obsAtual + "\n\n" : "") + registro);
    return "Envio para assistência técnica registado!";
  }
  throw new Error("Chamado não encontrado.");
}

function retornarDaAssistencia(dados) {
  const { sheet, rowIndex } = localizarLinhaPorOS(dados.os);
  if (rowIndex !== -1) {
    const obsAtual = sheet.getRange(rowIndex, 12).getValue();
    const novaObs = (obsAtual ? obsAtual + "\n" : "") + `[${formatarDataHora(new Date())}] Retorno da Assistência: ${dados.laudo}`;
    sheet.getRange(rowIndex, 9).setValue("PENDENTE");
    sheet.getRange(rowIndex, 12).setValue(novaObs);
    return "Retorno registado. Chamado recolocado na fila pendente.";
  }
  throw new Error("Chamado não encontrado.");
}

// ==========================================
// 2. MÓDULO DE ESTOQUE (KARDEX E BASE)
// ==========================================

function registrarKardex(tipo, idItem, descricao, qtd, tecnico, os, obs) {
  const sheet = obterOuCriarAba('Movimentacoes', ["Data", "Tipo", "ID_Item", "Descricao", "Quantidade", "Tecnico", "OS", "Observacao"]);
  sheet.appendRow([formatarDataHora(new Date()), tipo, idItem, descricao, qtd, tecnico || "Sistema", os || "-", obs || "-"]);
}

function avaliarAlertaMinimo(itemDesc, saldo, minimo) {
  if (saldo <= minimo) {
    try { MailApp.sendEmail(EMAIL_GESTOR, `[ALERTA ESTOQUE] Peça Crítica: ${itemDesc}`, `Atenção!\n\nO item "${itemDesc}" atingiu a margem de segurança.\nSaldo Atual: ${saldo}\nMínimo: ${minimo}\n\nProvidencie reposição.`); } catch(e) {}
  }
}

function buscarEstoqueCompleto() {
  const sheet = obterOuCriarAba('Estoque', ["ID_Item", "Descricao", "Categoria", "Qtd_Atual", "Qtd_Minima", "Localizacao", "Custo_Unitario", "Unidade", "Fornecedor"]);
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const estoque = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      estoque.push({
        id: data[i][0].toString(), descricao: data[i][1] ? data[i][1].toString() : "", categoria: data[i][2] ? data[i][2].toString() : "Geral",
        qtdAtual: Number(data[i][3]) || 0, qtdMinima: Number(data[i][4]) || 0, localizacao: data[i][5] ? data[i][5].toString() : "TI",
        custoUnitario: Number(data[i][6]) || 0, unidade: data[i][7] ? data[i][7].toString() : "Un", fornecedor: data[i][8] ? data[i][8].toString() : ""
      });
    }
  }
  return estoque;
}

function salvarItemEstoque(dados) {
  const sheet = obterOuCriarAba('Estoque', ["ID_Item", "Descricao", "Categoria", "Qtd_Atual", "Qtd_Minima", "Localizacao", "Custo_Unitario", "Unidade", "Fornecedor"]);
  const data = sheet.getDataRange().getValues();
  let idGerado = dados.id || "PEC-" + String((data.length > 1 ? data.length : 1)).padStart(3, '0');

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idGerado) {
      const qtdAnterior = Number(data[i][3]) || 0;
      const qtdNova = Number(dados.qtdAtual) || 0;
      sheet.getRange(i + 1, 2, 1, 8).setValues([[dados.descricao, dados.categoria, qtdNova, Number(dados.qtdMinima) || 0, dados.localizacao, Number(dados.custoUnitario) || 0, dados.unidade, dados.fornecedor]]);
      if (qtdNova !== qtdAnterior) registrarKardex(qtdNova > qtdAnterior ? "ENTRADA" : "SAIDA", idGerado, dados.descricao, Math.abs(qtdNova - qtdAnterior), dados.tecnicoResp, "-", "Ajuste manual");
      avaliarAlertaMinimo(dados.descricao, qtdNova, Number(dados.qtdMinima) || 0);
      return "Item " + idGerado + " atualizado!";
    }
  }
  sheet.appendRow([idGerado, dados.descricao, dados.categoria, Number(dados.qtdAtual) || 0, Number(dados.qtdMinima) || 0, dados.localizacao, Number(dados.custoUnitario) || 0, dados.unidade, dados.fornecedor]);
  registrarKardex("ENTRADA_INICIAL", idGerado, dados.descricao, Number(dados.qtdAtual) || 0, dados.tecnicoResp, "-", "Cadastro inicial");
  return "Item " + idGerado + " registado no stock!";
}

function darBaixaEstoqueInterno(idItem, quantidade, os, tecnicoResp) {
  const sheetEstoque = obterOuCriarAba('Estoque', ["ID_Item", "Descricao", "Categoria", "Qtd_Atual", "Qtd_Minima", "Localizacao", "Custo_Unitario", "Unidade", "Fornecedor"]);
  const data = sheetEstoque.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idItem.toString()) {
      const saldoAtual = Number(data[i][3]) || 0;
      const qtdPedida = Number(quantidade) || 0;
      
      // ERRO CRÍTICO AQUI SE NÃO HOUVER SALDO (Impede fechar o chamado)
      if (saldoAtual < qtdPedida) throw new Error(`Stock Insuficiente! A peça "${data[i][1]}" tem apenas ${saldoAtual} un. disponíveis. Pausa a O.S. para gerar um Pedido de Compra.`);
      
      const novoSaldo = saldoAtual - qtdPedida;
      sheetEstoque.getRange(i + 1, 4).setValue(novoSaldo);
      registrarKardex("SAIDA", idItem, data[i][1], qtdPedida, tecnicoResp, os, "Baixa na Conclusão da O.S.");
      avaliarAlertaMinimo(data[i][1], novoSaldo, Number(data[i][4]) || 0);

      const { sheet: sheetChamados, rowIndex } = localizarLinhaPorOS(os);
      if (rowIndex !== -1) {
        const obsAtual = sheetChamados.getRange(rowIndex, 12).getValue();
        sheetChamados.getRange(rowIndex, 12).setValue((obsAtual ? obsAtual + "\n" : "") + `[${formatarDataHora(new Date())}] Peça aplicada: ${data[i][1]} (${qtdPedida} un.)`);
      }
      return true;
    }
  }
  throw new Error("Item de stock inválido ou não encontrado.");
}

// ==========================================
// 3. MÓDULO DE COMPRAS (INTEGRAÇÃO O.S. <-> ESTOQUE)
// ==========================================

function buscarListaCompras() {
  const sheet = obterOuCriarAba('ListaCompras', ["ID_Compra", "Data", "ID_Item", "Descricao", "Qtd", "OS_Vinculada", "Status"]);
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const compras = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][6] === "PENDENTE") {
      compras.push({ idCompra: data[i][0], data: formatarDataHora(data[i][1]), idItem: data[i][2], descricao: data[i][3], qtd: data[i][4], os: data[i][5] });
    }
  }
  return compras;
}

function solicitarCompraPeca(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sheet, rowIndex } = localizarLinhaPorOS(dados.os);
    if (rowIndex === -1) throw new Error("O.S. não encontrada.");

    // Pausa a O.S.
    const obsAtual = sheet.getRange(rowIndex, 12).getValue();
    const msg = `[${formatarDataHora(new Date())}] Aguardando Compra de Peça: ${dados.qtd}x ${dados.descricao}`;
    sheet.getRange(rowIndex, 9).setValue("AGUARDANDO PECA");
    sheet.getRange(rowIndex, 12).setValue(obsAtual ? obsAtual + "\n" + msg : msg);

    // Regista Pedido de Compra
    const sheetCompras = obterOuCriarAba('ListaCompras', ["ID_Compra", "Data", "ID_Item", "Descricao", "Qtd", "OS_Vinculada", "Status"]);
    const idCompra = "COMP-" + new Date().getTime().toString().slice(-6);
    
    // Se a peça não tiver ID (texto livre), regista como "NOVO_ITEM"
    const idItemReal = dados.idItem || "ITEM_EXTRA"; 
    
    sheetCompras.appendRow([idCompra, new Date(), idItemReal, dados.descricao, dados.qtd, dados.os, "PENDENTE"]);

    return "Peça enviada para a Lista de Compras. O.S. pausada com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function registrarChegadaCompra(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // 1. Marca como COMPRADO na Lista de Compras
    const sheetCompras = obterOuCriarAba('ListaCompras', ["ID_Compra", "Data", "ID_Item", "Descricao", "Qtd", "OS_Vinculada", "Status"]);
    const dataCompras = sheetCompras.getDataRange().getValues();
    for (let i = 1; i < dataCompras.length; i++) {
      if (dataCompras[i][0] === dados.idCompra) {
        sheetCompras.getRange(i + 1, 7).setValue("COMPRADO");
        break;
      }
    }

    // 2. Dá Entrada no Estoque
    const sheetEstoque = obterOuCriarAba('Estoque', ["ID_Item", "Descricao", "Categoria", "Qtd_Atual", "Qtd_Minima", "Localizacao", "Custo_Unitario", "Unidade", "Fornecedor"]);
    const dataEstoque = sheetEstoque.getDataRange().getValues();
    let itemAtualizado = false;
    
    if (dados.idItem !== "ITEM_EXTRA") {
      for (let i = 1; i < dataEstoque.length; i++) {
        if (dataEstoque[i][0] === dados.idItem) {
          const saldoAtual = Number(dataEstoque[i][3]) || 0;
          sheetEstoque.getRange(i + 1, 4).setValue(saldoAtual + Number(dados.qtd));
          if (dados.custo) sheetEstoque.getRange(i + 1, 7).setValue(dados.custo);
          itemAtualizado = true;
          break;
        }
      }
    }
    
    if (!itemAtualizado && dados.idItem !== "ITEM_EXTRA") {
       sheetEstoque.appendRow([dados.idItem, dados.descricao, "Geral", dados.qtd, 1, "TI", dados.custo || 0, "Un", ""]);
    }

    registrarKardex("ENTRADA_COMPRA", dados.idItem, dados.descricao, dados.qtd, dados.tecnicoResp, dados.osVinculada, "Chegada de Pedido de Compra");

    // 3. Destrava a O.S. (Muda para Aguardando Instalação)
    if (dados.osVinculada && dados.osVinculada !== "-") {
      const { sheet, rowIndex } = localizarLinhaPorOS(dados.osVinculada);
      if (rowIndex !== -1) {
         const obsAtual = sheet.getRange(rowIndex, 12).getValue();
         const msg = `[${formatarDataHora(new Date())}] A Peça requerida chegou do fornecedor. Pode efetuar a instalação.`;
         sheet.getRange(rowIndex, 9).setValue("AGUARDANDO INSTALACAO PECA");
         sheet.getRange(rowIndex, 12).setValue(obsAtual ? obsAtual + "\n" + msg : msg);
      }
    }
    return "Entrada registada! Estoque e O.S. foram atualizados automaticamente.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. AUXILIARES, CADASTROS E DASHBOARD
// ==========================================

function buscarColunaSimples(nomeAba) {
  const sheet = obterOuCriarAba(nomeAba, ["Nome"]);
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < data.length; i++) if (data[i][0]) lista.push(data[i][0].toString().trim());
  return lista;
}

function removerItemColuna(nomeAba, colunaIdx, valor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][colunaIdx - 1] && data[i][colunaIdx - 1].toString().trim() === valor.toString().trim()) { sheet.deleteRow(i + 1); break; }
  }
}

function buscarUsuarios() {
  const sheet = obterOuCriarAba('Usuarios', ["Nome", "Setor"]);
  if (sheet.getLastRow() <= 1) return [];
  const dados = sheet.getDataRange().getValues();
  const usuarios = [];
  for (let i = 1; i < dados.length; i++) if (dados[i][0]) usuarios.push({ nome: dados[i][0].toString().trim(), setor: dados[i][1] ? dados[i][1].toString().trim() : "" });
  return usuarios;
}
function adicionarUsuario(nome, setor) { obterOuCriarAba('Usuarios', ["Nome", "Setor"]).appendRow([nome, setor || ""]); return true; }
function removerUsuario(nome) { removerItemColuna('Usuarios', 1, nome); }
function buscarSetores() { return buscarColunaSimples('Setores'); }
function adicionarSetor(nome) { obterOuCriarAba('Setores', ["Setor"]).appendRow([nome]); return true; }
function removerSetor(nome) { removerItemColuna('Setores', 1, nome); }
function buscarTecnicos() { return buscarColunaSimples('Tecnicos'); }
function adicionarTecnico(nome) { obterOuCriarAba('Tecnicos', ["Nome"]).appendRow([nome]); return true; }
function removerTecnico(nome) { removerItemColuna('Tecnicos', 1, nome); }

function obterDadosDashboard(filtros) {
  try {
    const sheet = obterFolhaChamados();
    const statusCounts = {}, tecnicoCounts = {}, setorCounts = {}, tipoCounts = {}, tendenciaDias = {};
    const mesesSet = new Set(), tecnicosSet = new Set(), setoresSet = new Set(), chamadosAtrasados = [];

    if (sheet.getLastRow() <= 1) return { totalGeral: 0, totalPendentes: 0, totalAguardando: 0, totalInstalacao: 0, totalConcluidos: 0, totalManutencao: 0, tmaFormatado: "0h 0m", statusCounts, tecnicoCounts, setorCounts, tipoCounts, tendenciaDias, mesesDisponiveis: [], tecnicosDisponiveis: [], setoresDisponiveis: [], alertas: [] };

    const data = sheet.getDataRange().getValues();
    let totalGeral = 0, totalPendentes = 0, totalAguardando = 0, totalInstalacao = 0, totalConcluidos = 0, totalManutencao = 0, somaMinutosConcluidos = 0, totalConcluidosComTempo = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const dataInicioStr = formatarDataHora(data[i][5]); 
      let mesAno = "Outros", diaAbertura = "Data";
      if (dataInicioStr) {
        const partesData = dataInicioStr.split(' ')[0].split('/');
        if (partesData.length === 3) { mesAno = `${partesData[1]}/${partesData[2]}`; diaAbertura = `${partesData[0]}/${partesData[1]}`; mesesSet.add(mesAno); }
      }

      const statusNorm = (data[i][8] || '').toString().trim().toUpperCase(), tecnico = (data[i][9] || 'Não Atribuído').toString().trim(), setor = (data[i][2] || 'Geral').toString().trim(), tipo = (data[i][10] || 'Não especificado').toString().trim(); 
      tecnicosSet.add(tecnico); setoresSet.add(setor);

      if (filtros && filtros.mes && filtros.mes !== "GERAL" && mesAno !== filtros.mes) continue;
      if (filtros && filtros.tecnico && filtros.tecnico !== "TODOS" && tecnico !== filtros.tecnico) continue;
      if (filtros && filtros.setor && filtros.setor !== "TODOS" && setor !== filtros.setor) continue;

      totalGeral++;
      if (statusNorm === "PENDENTE") { totalPendentes++; chamadosAtrasados.push({ os: data[i][0], setor: setor, inicio: dataInicioStr }); }
      else if (statusNorm === "AGUARDANDO PECA") totalAguardando++;
      else if (statusNorm === "ASSISTENCIA TECNICA") totalManutencao++;
      else if (statusNorm === "AGUARDANDO INSTALACAO PECA") totalInstalacao++;
      else if (statusNorm === "CONCLUIDO") totalConcluidos++;

      statusCounts[statusNorm || "OUTRO"] = (statusCounts[statusNorm || "OUTRO"] || 0) + 1;
      tecnicoCounts[tecnico] = (tecnicoCounts[tecnico] || 0) + 1;
      setorCounts[setor] = (setorCounts[setor] || 0) + 1;
      
      if (statusNorm === "CONCLUIDO") {
        tipoCounts[tipo] = (tipoCounts[tipo] || 0) + 1;
        tendenciaDias[diaAbertura] = (tendenciaDias[diaAbertura] || 0) + 1;
        if (data[i][5] && data[i][6]) {
          try {
            const dInicio = parseDataBR(data[i][5]), dFim = parseDataBR(data[i][6]);
            if (dInicio && dFim && dFim > dInicio) { somaMinutosConcluidos += (dFim - dInicio) / (1000 * 60); totalConcluidosComTempo++; }
          } catch(e) {}
        }
      }
    }
    let tmaFormatado = "0h 0m";
    if (totalConcluidosComTempo > 0) { const mediaMinutos = Math.round(somaMinutosConcluidos / totalConcluidosComTempo); tmaFormatado = `${Math.floor(mediaMinutos / 60)}h ${mediaMinutos % 60}m`; }
    return { totalGeral, totalPendentes, totalAguardando, totalInstalacao, totalConcluidos, totalManutencao, tmaFormatado, statusCounts, tecnicoCounts, setorCounts, tipoCounts, tendenciaDias, alertas: chamadosAtrasados.slice(0, 5), mesesDisponiveis: Array.from(mesesSet).sort().reverse(), tecnicosDisponiveis: Array.from(tecnicosSet).sort(), setoresDisponiveis: Array.from(setoresSet).sort() };
  } catch (err) { return null; }
}

function analisarDashboardComIA(filtroMes) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("Chave GEMINI_API_KEY não configurada.");
  const dados = obterDadosDashboard({mes: filtroMes, tecnico: "TODOS", setor: "TODOS"});
  const estoque = buscarEstoqueCompleto();
  const prompt = `Você é o Gestor Estratégico de TI. Analise os dados do sistema (Período: ${filtroMes}): Pendentes: ${dados.totalPendentes} | Faltam Peças: ${dados.totalAguardando} | Concluídos: ${dados.totalConcluidos} | TMA: ${dados.tmaFormatado} | Chamados por Setor: ${JSON.stringify(dados.setorCounts)}. Estoque Crítico: ${JSON.stringify(estoque.filter(item => item.qtdAtual <= item.qtdMinima))}. Elabore um Diagnóstico Estratégico em Markdown abordando Resumo, Setores Críticos, Análise de Stock e 3 Recomendações.`;
  const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true });
  const responseData = JSON.parse(response.getContentText());
  if (responseData.error) throw new Error(responseData.error.message);
  return responseData.candidates[0].content.parts[0].text;
}
