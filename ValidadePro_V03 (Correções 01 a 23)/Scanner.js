// ==================== SCANNER ====================
function obterOuCriarAbaPesquisas(ss) {
var planilha = ss.getSheetByName("Pesquisas");
if (!planilha) {
planilha = ss.insertSheet("Pesquisas");
planilha.appendRow(["Data/Hora", "Código", "Descrição", "Quantidade", "Data Validade"]);
planilha.setColumnWidth(1, 140);
planilha.setColumnWidth(2, 120);
planilha.setColumnWidth(3, 250);
planilha.setColumnWidth(4, 100);
planilha.setColumnWidth(5, 120);
}
return planilha;
}
function salvarDadosNoSheets(codigoBarras) {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var planilha = obterOuCriarAbaPesquisas(ss);
var descricao = "Não localizado na internet";
var codigoLimpo = codigoBarras.trim();
var quantidadeEstoque = "";
var dataValidade = "";
var mensagemEstoque = "";
// ===== ITEM 23: categoria/subcategoria do produto (preenchidas se já houver cadastro anterior) =====
var categoriaEstoque = "";
var subcategoriaEstoque = "";
// ===== FIM ITEM 23 =====
if (codigoLimpo.indexOf("http") !== 0) {
try {
var urlApi = "https://world.openfoodfacts.org/api/v0/product/" + encodeURIComponent(codigoLimpo) + ".json";
var opcoes = {
"muteHttpExceptions": true,
"headers": {
"User-Agent": "ScannerAppGoogleSheets - Vinicius - Versao2.0"
}
};
var resposta = UrlFetchApp.fetch(urlApi, opcoes);
if (resposta.getResponseCode() === 200) {
var dados = JSON.parse(resposta.getContentText());
if (dados.status === 1 && dados.product) {
descricao = dados.product.product_name_pt || dados.product.product_name || "Produto sem nome cadastrado";
}
} else {
descricao = "Não localizado (Status HTTP: " + resposta.getResponseCode() + ")";
}
} catch (erroApi) {
descricao = "Erro na Busca: " + erroApi.message;
}
try {
var planilhaEstoque = ss.getSheetByName("Estoque");
if (planilhaEstoque) {
var dadosEstoque = planilhaEstoque.getDataRange().getValues();
var encontrado = false;
for (var i = 1; i < dadosEstoque.length; i++) {
var codigoEstoque = String(dadosEstoque[i][0]).trim();
if (codigoEstoque === codigoLimpo) {
quantidadeEstoque = dadosEstoque[i][2];
dataValidade = dadosEstoque[i][3];
if (dataValidade instanceof Date) {
dataValidade = Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), "dd/MM/yyyy");
}
// ===== ITEM 23: recupera Categoria e Subcategoria já cadastradas para este código =====
categoriaEstoque = dadosEstoque[i][4] || "";
subcategoriaEstoque = dadosEstoque[i][5] || "";
// ===== FIM ITEM 23 =====
encontrado = true;
break;
}
}
if (encontrado) {
mensagemEstoque = "Produto encontrado no estoque!";
} else {
mensagemEstoque = "Produto não encontrado no estoque";
quantidadeEstoque = "N/A";
dataValidade = "N/A";
}
} else {
mensagemEstoque = "Planilha 'Estoque' não encontrada";
quantidadeEstoque = "N/A";
dataValidade = "N/A";
}
} catch (erroEstoque) {
mensagemEstoque = "Erro ao consultar estoque: " + erroEstoque.message;
quantidadeEstoque = "Erro";
dataValidade = "Erro";
}
} else {
descricao = "Link / URL Institucional";
mensagemEstoque = "Não aplicável para links";
quantidadeEstoque = "N/A";
dataValidade = "N/A";
}
var dataHoraAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
planilha.appendRow([dataHoraAtual, codigoLimpo, descricao, quantidadeEstoque, dataValidade]);
SpreadsheetApp.flush();
return {
sucesso: true,
codigo: codigoLimpo,
descricao: descricao,
quantidade: String(quantidadeEstoque),
dataValidade: String(dataValidade),
mensagemEstoque: mensagemEstoque,
// ===== ITEM 23: envia Categoria/Subcategoria para o frontend pré-preencher o formulário =====
categoria: categoriaEstoque,
subcategoria: subcategoriaEstoque
// ===== FIM ITEM 23 =====
};
} catch (erro) {
return {
sucesso: false,
mensagem: erro.toString()
};
}
}
function obterUltimasPesquisas() {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var planilha = ss.getSheetByName("Pesquisas");
if (!planilha) {
return [];
}
var ultimaLinha = planilha.getLastRow();
if (ultimaLinha <= 1) {
return [];
}
var numLinhas = Math.min(10, ultimaLinha - 1);
var linhaInicial = ultimaLinha - numLinhas + 1;
if (linhaInicial < 2) {
linhaInicial = 2;
numLinhas = ultimaLinha - 1;
}
var dadosBrutos = planilha.getRange(linhaInicial, 1, numLinhas, 5).getValues();
var dadosFormatados = [];
for (var i = 0; i < dadosBrutos.length; i++) {
var linha = [];
for (var j = 0; j < dadosBrutos[i].length; j++) {
var valor = dadosBrutos[i][j];
if (valor instanceof Date) {
valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
} else if (valor === null || valor === undefined) {
valor = "";
} else {
valor = String(valor);
}
linha.push(valor);
}
dadosFormatados.push(linha);
}
return dadosFormatados.slice().reverse();
} catch (e) {
return [];
}
}