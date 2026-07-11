// ==================== IMPORTAÇÕES E CONFIGURAÇÕES ====================
function doGet(e) {
return HtmlService.createTemplateFromFile('Index').evaluate()
.setTitle('ValidadePro - Gestão Inteligente')
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}
// Necessário para que Index.html possa incluir os demais arquivos .html (Styles, Dashboard, Estoque, etc.)
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
sheet.appendRow(['Código', 'Produto', 'Quantidade', 'Data Validade', 'Categoria', 'Subcategoria', 'Preço Custo', 'Preço Venda', 'Fornecedor', 'Estabelecimento', 'Quebra']);
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
// Auto-migração: planilha Acessos já existente sem a coluna Loja ganha o cabeçalho automaticamente
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