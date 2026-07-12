// ==================== IMPORTAÇÕES E CONFIGURAÇÕES ====================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ValidadePro - Gestão Inteligente')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Acessos') {
      sheet.appendRow(['Nome', 'Usuário', 'Senha', 'Setor', 'Perfil', 'Status', 'Último Acesso', 'Loja']);
    } else if (name === 'Estoque') {
      // ===== ITEM 20: cabeçalho completo com colunas de auditoria =====
      // ===== ITEM 22: coluna "Data Rebaixa" (data em que Quebra virou 'Sim') =====
      sheet.appendRow([
        'Código', 'Produto', 'Quantidade', 'Data Validade', 'Categoria', 'Subcategoria',
        'Preço Custo', 'Preço Venda', 'Fornecedor', 'Estabelecimento', 'Quebra',
        'Cadastrado Por', 'Função', 'Data Cadastro', 'Resolvido', 'Data Rebaixa'
      ]);
    } else if (name === 'Categoria_Subcategoria') {
      sheet.appendRow(['Categoria', 'Subcategoria']);
    } else if (name === 'Estabelecimentos') {
      sheet.appendRow(['Estabelecimento']);
    } else if (name === 'Fornecedores') {
      sheet.appendRow(['Fornecedor']);
    } else if (name === 'Perfis') {
      sheet.appendRow(['Perfil', 'Dashboard', 'Estoque', 'Relatorios', 'Scanner', 'Configuracoes']);
      sheet.appendRow(['Gerente', 'FULL', 'FULL', 'FULL', 'FULL', 'FULL']);
      sheet.appendRow(['Consultor', 'VIEW', 'VIEW', 'VIEW', 'VIEW', 'VIEW']);
      sheet.appendRow(['Cadastrador', 'NONE', 'EDIT', 'NONE', 'EDIT', 'NONE']);
    }
  }

  // ===== Migração para colunas de auditoria na planilha Estoque (caso já exista) =====
  if (name === 'Estoque') {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Verifica se a coluna 12 (Cadastrado Por) não existe ou está vazia
    if (headers.length < 12 || headers[11] !== 'Cadastrado Por') {
      // Adiciona cabeçalhos faltantes a partir da coluna 12
      var colStart = headers.length + 1;
      var novosCabecalhos = [];
      if (colStart <= 12) novosCabecalhos.push('Cadastrado Por');
      if (colStart <= 13) novosCabecalhos.push('Função');
      if (colStart <= 14) novosCabecalhos.push('Data Cadastro');
      if (colStart <= 15) novosCabecalhos.push('Resolvido');
      if (novosCabecalhos.length > 0) {
        sheet.getRange(1, colStart, 1, novosCabecalhos.length).setValues([novosCabecalhos]);
      }
    }
    // Garante que a coluna "Resolvido" exista (pode ter sido adicionada antes)
    if (headers.length < 15 || headers[14] !== 'Resolvido') {
      sheet.getRange(1, 15).setValue('Resolvido');
    }
    // ===== ITEM 22: garante a coluna "Data Rebaixa" (16ª coluna) =====
    var headersAtualizados = sheet.getRange(1, 1, 1, 16).getValues()[0];
    if (headersAtualizados.length < 16 || headersAtualizados[15] !== 'Data Rebaixa') {
      sheet.getRange(1, 16).setValue('Data Rebaixa');
    }
    // ===== FIM ITEM 22 =====
  }

  // ===== Migração para coluna Loja na planilha Acessos =====
  if (name === 'Acessos' && sheet.getLastColumn() < 8) {
    sheet.getRange(1, 8).setValue('Loja');
  }

  return sheet;
}

// ==================== UTILITÁRIO DE ORDENAÇÃO ====================
function ordenarAlfabetico(lista) {
  return lista.sort(function(a, b) {
    return String(a).localeCompare(String(b), 'pt-BR', { sensitivity: 'base' });
  });
}