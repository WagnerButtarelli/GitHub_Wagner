// ==================== CONFIGURAÇÕES PRINCIPAIS ====================
function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('Index');
    var output = template.evaluate()
      .setTitle('ValidadePro - Gestão Inteligente')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return output;
  } catch (erro) {
    // Retorna uma página de erro visível para diagnóstico
    return HtmlService.createHtmlOutput(
      '<h1 style="color:red">Erro ao carregar o app:</h1>' +
      '<pre style="background:#f0f0f0;padding:20px">' + erro.toString() + '</pre>' +
      '<p>Verifique se todos os arquivos HTML existem: Index, Styles, Secoes, Modais, Script</p>'
    );
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (erro) {
    return '<div style="color:red;padding:20px;background:#ffe0e0;border:1px solid red;margin:10px;">' +
           '<strong>❌ ERRO AO CARREGAR ARQUIVO: ' + filename + '</strong><br>' +
           'Verifique se o arquivo "' + filename + '.html" existe no projeto.<br>' +
           'Detalhe: ' + erro.toString() + '</div>';
  }
}

function getSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Acessos') {
      sheet.appendRow(['Nome', 'Usuário', 'Senha', 'Setor', 'Perfil', 'Status', 'Último Acesso']);
    } else if (name === 'Estoque') {
      sheet.appendRow(['Código', 'Produto', 'Quantidade', 'Data Validade', 'Categoria', 'Subcategoria', 'Preço Custo', 'Preço Venda', 'Fornecedor', 'Estabelecimento', 'Quebra']);
    } else if (name === 'Categoria_Subcategoria') {
      sheet.appendRow(['Categoria', 'Subcategoria']);
    } else if (name === 'Estabelecimentos') {
      sheet.appendRow(['Estabelecimento']);
    } else if (name === 'Fornecedores') {
      sheet.appendRow(['Fornecedor']);
    } else if (name === 'Pesquisas') {
      sheet.appendRow(['Data/Hora', 'Código', 'Descrição', 'Quantidade', 'Data Validade']);
    }
  }
  return sheet;
}