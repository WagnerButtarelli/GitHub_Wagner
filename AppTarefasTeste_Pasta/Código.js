/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * 
 * SGT v2.1 - Google Apps Script Backend
 * Sistema de Designação e Controle de Tarefas
 * Planilha com abas: "Acessos", "Permissões", "Tarefas"
 * 
 * ATENÇÃO: Ative o serviço avançado "Drive API" no Editor de Apps Script.
 */

// 1. Renderizador da Web App
function doGet(e) {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SGT - Designação e Controle de Tarefas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return htmlOutput;
}

// 2. Obter planilha ativa (compatível com container-bound)
function obterPlanilhaAtiva() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    var active = SpreadsheetApp.getActive();
    if (!active) throw new Error("Execute este script dentro de uma Google Sheets.");
    return active;
  }
}

// 3. Inicializar planilha (cria abas e cabeçalhos se não existirem)
function inicializarPlanilha() {
  var ss = obterPlanilhaAtiva();
  
  // Aba Acessos – colunas: Nome, Identificação, Usuário, Senha, Perfil, Nome da Loja, E-mail, Telefone, Status, Último Acesso
  var abaAcessos = ss.getSheetByName("Acessos");
  if (!abaAcessos) {
    abaAcessos = ss.insertSheet("Acessos");
    var cabecalhos = ["Nome", "Identificação", "Usuário", "Senha", "Perfil", "Nome da Loja", "E-mail", "Telefone", "Status", "Último Acesso"];
    abaAcessos.appendRow(cabecalhos);
    abaAcessos.getRange("A1:J1").setFontWeight("bold").setBackground("#e2e8f0");
    // Seed com dados de exemplo
    abaAcessos.appendRow(["Edwin Alves", "7939999", "Edwinson", "1", "Gerente", "Loja Centro", "", "", "Ativo", ""]);
    abaAcessos.appendRow(["Wagner Oliveira", "123", "Wagner", "1276", "Gerente", "Loja Norte", "", "", "Ativo", ""]);
    abaAcessos.appendRow(["Francisco Almei", "8888", "Francisco", "3", "Executor", "Loja Sul", "", "", "Ativo", ""]);
  }
  
  // Aba Permissões
  var abaPermissoes = ss.getSheetByName("Permissões");
  if (!abaPermissoes) {
    abaPermissoes = ss.insertSheet("Permissões");
    var cabecalhosPerm = ["Perfil", "Dashboard", "Tarefas", "Relatório", "Visualizar", "Incluir", "Excluir", "Exportar"];
    abaPermissoes.appendRow(cabecalhosPerm);
    abaPermissoes.getRange("A1:H1").setFontWeight("bold").setBackground("#cbd5e1");
    abaPermissoes.appendRow(["Gerente", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE"]);
    abaPermissoes.appendRow(["Cadastrador/Executor", "FALSE", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE"]);
    abaPermissoes.appendRow(["Consultor", "FALSE", "FALSE", "TRUE", "TRUE", "FALSE", "FALSE", "TRUE"]);
    abaPermissoes.appendRow(["Executor", "FALSE", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE"]);
  }
  
  // Aba Tarefas
  var abaTarefas = ss.getSheetByName("Tarefas");
  if (!abaTarefas) {
    abaTarefas = ss.insertSheet("Tarefas");
    var cabecalhosTar = ["ID", "Tipo da Tarefa", "Executor", "Nome da Loja", "Data Limite", "Descrição", "Foto (Drive URL)", "Status", "Observação", "Foto Execução"];
    abaTarefas.appendRow(cabecalhosTar);
    abaTarefas.getRange("A1:J1").setFontWeight("bold").setBackground("#94a3b8");
    abaTarefas.appendRow(["T-01", "Inventário de Estoque", "Edwinson", "Loja Centro", "2026-06-25", "Contagem anual e auditoria geral de estoque.", "", "Concluída", "", ""]);
    abaTarefas.appendRow(["T-02", "Limpeza / Organização", "Wagner", "Loja Norte", "2026-06-28", "Limpeza profunda dos aparelhos de ar-condicionado.", "", "Vencida", "", ""]);
    abaTarefas.appendRow(["T-03", "Manutenção Geral", "Francisco", "Loja Sul", "2026-07-02", "Ajuste na porta giratória frontal.", "", "Pendente", "", ""]);
  } else {
    var header = abaTarefas.getRange(1, 1, 1, abaTarefas.getLastColumn()).getValues()[0];
    if (header.indexOf("Observação") === -1) {
      abaTarefas.getRange(1, 9).setValue("Observação");
      abaTarefas.getRange(1, 10).setValue("Foto Execução");
    }
  }
}

// 4. Ler Acessos
function lerAcessos() {
  inicializarPlanilha();
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Acessos");
  var lr = sheet.getLastRow();
  if (lr < 2) return [];
  
  var values = sheet.getRange(2, 1, lr - 1, 10).getValues();
  return values.map(function(row, index) {
    return {
      id: String(index + 1),
      nome: String(row[0]),
      identificacao: String(row[1]),
      usuario: String(row[2]),
      senha: String(row[3]),
      perfil: String(row[4]),
      loja: String(row[5]),
      email: String(row[6]),
      telefone: String(row[7]),
      status: String(row[8]),
      ultimoAcesso: String(row[9])
    };
  });
}

// 5. Ler Permissões
function lerPermissoes() {
  inicializarPlanilha();
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Permissões");
  var lr = sheet.getLastRow();
  if (lr < 2) return [];
  
  var values = sheet.getRange(2, 1, lr - 1, 8).getValues();
  return values.map(function(row) {
    return {
      perfil: String(row[0]),
      dashboard: String(row[1]).toUpperCase() === 'TRUE',
      tarefas: String(row[2]).toUpperCase() === 'TRUE',
      relatorio: String(row[3]).toUpperCase() === 'TRUE',
      visualizar: String(row[4]).toUpperCase() === 'TRUE',
      incluir: String(row[5]).toUpperCase() === 'TRUE',
      excluir: String(row[6]).toUpperCase() === 'TRUE',
      exportar: String(row[7]).toUpperCase() === 'TRUE'
    };
  });
}

// 6. Ler Tarefas (com auto-vencimento)
function lerTarefas() {
  inicializarPlanilha();
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Tarefas");
  var lr = sheet.getLastRow();
  if (lr < 2) return [];
  
  var values = sheet.getRange(2, 1, lr - 1, 10).getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);
  
  return values.map(function(row) {
    var rawDate = row[4];
    var dataStr = "";
    if (rawDate instanceof Date) {
      var y = rawDate.getFullYear();
      var m = ("0" + (rawDate.getMonth() + 1)).slice(-2);
      var d = ("0" + rawDate.getDate()).slice(-2);
      dataStr = y + "-" + m + "-" + d;
    } else {
      dataStr = String(rawDate);
    }
    
    var status = String(row[7]);
    if (status === 'Pendente' && dataStr) {
      var partes = dataStr.split('-');
      if (partes.length === 3) {
        var dataLimite = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
        dataLimite.setHours(0,0,0,0);
        if (dataLimite < hoje) status = 'Vencida';
      }
    }
    
    return {
      id: String(row[0]),
      tipoTarefa: String(row[1]),
      executor: String(row[2]),
      nomeLoja: String(row[3]),
      dataLimite: dataStr,
      descricao: String(row[5]),
      foto: String(row[6]),
      status: status,
      observacao: String(row[8] || ''),
      fotoExecucao: String(row[9] || '')
    };
  });
}

// 7. Autenticar Usuário
function autenticarUsuario(usuarioInput, senhaInput) {
  var acessos = lerAcessos();
  var matched = acessos.filter(function(u) {
    return u.usuario.toLowerCase() === usuarioInput.trim().toLowerCase() && u.senha === senhaInput;
  });
  
  if (matched.length === 0) {
    return { sucesso: false, mensagem: "Usuário ou senha incorretos." };
  }
  
  var user = matched[0];
  if (user.status !== "Ativo") {
    return { sucesso: false, mensagem: "Esta conta está inativa no sistema." };
  }
  
  var cache = CacheService.getUserCache();
  if (cache) {
    cache.put("sessao_usuario_id", user.id, 7200);
    cache.put("sessao_usuario_perfil", user.perfil, 7200);
  }
  
  return {
    sucesso: true,
    usuario: user
  };
}

// 8. Cadastrar Tarefa
function cadastrarTarefa(tarefaObj, base64Foto) {
  inicializarPlanilha();
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Tarefas");
  
  var idFotoDriveUrl = "";
  if (base64Foto && base64Foto.indexOf("base64,") !== -1) {
    try {
      var partes = base64Foto.split("base64,");
      var contentType = partes[0].split(":")[1].split(";")[0];
      var decodedBytes = Utilities.base64Decode(partes[1]);
      var blob = Utilities.newBlob(decodedBytes, contentType, "comprovante_" + tarefaObj.id + ".jpg");
      
      var pastas = DriveApp.getFoldersByName("SGT_Comprovantes");
      var pasta;
      if (pastas.hasNext()) pasta = pastas.next();
      else pasta = DriveApp.createFolder("SGT_Comprovantes");
      
      var arquivo = pasta.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      idFotoDriveUrl = arquivo.getUrl();
    } catch (e) {
      idFotoDriveUrl = "Erro ao salvar foto: " + e.message;
    }
  }
  
  sheet.appendRow([
    tarefaObj.id,
    tarefaObj.tipoTarefa,
    tarefaObj.executor,
    tarefaObj.nomeLoja,
    tarefaObj.dataLimite,
    tarefaObj.descricao,
    idFotoDriveUrl,
    tarefaObj.status,
    tarefaObj.observacao || '',
    tarefaObj.fotoExecucao || ''
  ]);
  
  return { sucesso: true, urlFoto: idFotoDriveUrl };
}

// 9. Atualizar Status
function atualizarStatusTarefa(id, status) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Tarefas");
  var lr = sheet.getLastRow();
  if (lr < 2) return false;
  
  var ids = sheet.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sheet.getRange(i + 2, 8).setValue(status);
      return true;
    }
  }
  return false;
}

// 10. Excluir Tarefa
function excluirTarefaPlanilha(id) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Tarefas");
  var lr = sheet.getLastRow();
  if (lr < 2) return false;
  
  var ids = sheet.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

// 11. Incluir Novo Acesso
function incluirNovoAcesso(acessoObj) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Acessos");
  var lastRow = sheet.getLastRow();
  var novoId = lastRow;
  if (lastRow < 2) novoId = 1;
  
  sheet.appendRow([
    acessoObj.nome,
    acessoObj.identificacao || '',
    acessoObj.usuario,
    acessoObj.senha,
    acessoObj.perfil,
    acessoObj.loja,
    acessoObj.email || '',
    acessoObj.telefone || '',
    "Ativo",
    ""
  ]);
  return { sucesso: true };
}

// 12. Atualizar Observação e Foto de Execução
function atualizarTarefaComObservacao(id, observacao, fotoExecucaoBase64) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Tarefas");
  var lr = sheet.getLastRow();
  if (lr < 2) return false;
  
  var ids = sheet.getRange(2, 1, lr - 1, 1).getValues();
  var rowIndex = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return false;
  
  sheet.getRange(rowIndex, 9).setValue(observacao || '');
  
  if (fotoExecucaoBase64 && fotoExecucaoBase64.indexOf("base64,") !== -1) {
    try {
      var partes = fotoExecucaoBase64.split("base64,");
      var contentType = partes[0].split(":")[1].split(";")[0];
      var decodedBytes = Utilities.base64Decode(partes[1]);
      var blob = Utilities.newBlob(decodedBytes, contentType, "execucao_" + id + ".jpg");
      
      var pastas = DriveApp.getFoldersByName("SGT_Execucoes");
      var pasta;
      if (pastas.hasNext()) pasta = pastas.next();
      else pasta = DriveApp.createFolder("SGT_Execucoes");
      
      var arquivo = pasta.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = arquivo.getUrl();
      sheet.getRange(rowIndex, 10).setValue(url);
    } catch (e) {
      sheet.getRange(rowIndex, 10).setValue("Erro ao salvar foto: " + e.message);
    }
  }
  return true;
}

//// 13. Obter espaço da pasta do Drive
function obterEspacoPasta(pastaId) {
  try {
    // Verifica se a pasta existe
    var folder = DriveApp.getFolderById(pastaId);
    var files = folder.getFiles();
    var totalSize = 0;
    while (files.hasNext()) {
      var file = files.next();
      totalSize += file.getSize();
    }
    // Obtém a cota total do Drive
    var about = Drive.About.get({ fields: 'storageQuota' });
    var quotaTotal = about.storageQuota.limit;
    var quotaUsed = about.storageQuota.usage;
    var percent = quotaTotal > 0 ? (totalSize / quotaTotal) * 100 : 0;
    
    return {
      folderSize: totalSize,
      quotaTotal: quotaTotal,
      quotaUsed: quotaUsed,
      percent: percent
    };
  } catch (e) {
    return { error: e.message };
  }
}

// 14. Excluir Acesso (colaborador)
function excluirAcesso(id) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Acessos");
  var lr = sheet.getLastRow();
  if (lr < 2) return false;
  var row = parseInt(id) + 1; // id é o índice 0-based
  if (row >= 2 && row <= lr) {
    sheet.deleteRow(row);
    return true;
  }
  return false;
}

// 15. Atualizar Acesso (colaborador)
function atualizarAcesso(acessoObj) {
  var ss = obterPlanilhaAtiva();
  var sheet = ss.getSheetByName("Acessos");
  var row = parseInt(acessoObj.id) + 1;
  var lr = sheet.getLastRow();
  if (row < 2 || row > lr) return false;
  
  sheet.getRange(row, 1).setValue(acessoObj.nome);
  sheet.getRange(row, 2).setValue(acessoObj.identificacao || '');
  sheet.getRange(row, 3).setValue(acessoObj.usuario);
  sheet.getRange(row, 4).setValue(acessoObj.senha);
  sheet.getRange(row, 5).setValue(acessoObj.perfil);
  sheet.getRange(row, 6).setValue(acessoObj.loja);
  sheet.getRange(row, 7).setValue(acessoObj.email || '');
  sheet.getRange(row, 8).setValue(acessoObj.telefone || '');
  sheet.getRange(row, 9).setValue(acessoObj.status || 'Ativo');
  // Coluna J (Último Acesso) não é alterada
  return true;
}