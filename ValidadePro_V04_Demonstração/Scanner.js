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
// ===== ITEM 24: fonte adicional de pesquisa de produtos por código de barras (Cosmos Bluesoft) =====
// Documentação: https://cosmos.bluesoft.com.br/api
// A API da Cosmos exige DOIS dados que ficam disponíveis no seu painel logado em cosmos.bluesoft.com.br:
//   1) o token de acesso (cabeçalho "X-Cosmos-Token")
//   2) um User-Agent específico da sua conta (cabeçalho "User-Agent")
// Para não expor essas credenciais no código-fonte, ambas são lidas de PropertiesService
// (Extensões > Propriedades do projeto > Propriedades do script), usando as chaves
// "COSMOS_TOKEN" e "COSMOS_USER_AGENT". Caso o token não esteja configurado, esta fonte é
// simplesmente ignorada e o fluxo original (Open Food Facts) continua funcionando normalmente.
function buscarProdutoCosmos(codigoLimpo) {
try {
var token = PropertiesService.getScriptProperties().getProperty('COSMOS_TOKEN');
if (!token) return { descricao: null, debug: 'COSMOS_TOKEN não configurado nas Propriedades do script.' };
var userAgent = PropertiesService.getScriptProperties().getProperty('COSMOS_USER_AGENT');
if (!userAgent) {
return { descricao: null, debug: 'COSMOS_USER_AGENT não configurado nas Propriedades do script. A Cosmos exige um User-Agent específico da sua conta, disponível no painel do site junto com o token.' };
}
// ===== Endpoint correto: /gtins/{codigo}.json (e não /products/) =====
var urlCosmos = "https://api.cosmos.bluesoft.com.br/gtins/" + encodeURIComponent(codigoLimpo) + ".json";
var opcoesCosmos = {
"muteHttpExceptions": true,
"headers": {
"X-Cosmos-Token": token,
"User-Agent": userAgent
}
};
var respostaCosmos = UrlFetchApp.fetch(urlCosmos, opcoesCosmos);
var statusCosmos = respostaCosmos.getResponseCode();
if (statusCosmos === 200) {
var dadosCosmos = JSON.parse(respostaCosmos.getContentText());
if (dadosCosmos && dadosCosmos.description) {
return { descricao: dadosCosmos.description, debug: 'HTTP 200 - produto encontrado.' };
}
return { descricao: null, debug: 'HTTP 200, mas sem campo "description" na resposta: ' + respostaCosmos.getContentText().substring(0, 300) };
}
// 401/403 = token ou User-Agent inválidos | 404 = produto não cadastrado na Cosmos | 429 = limite de chamadas excedido
var motivo = statusCosmos === 401 || statusCosmos === 403 ? ' (token ou User-Agent inválidos)' : statusCosmos === 404 ? ' (produto não cadastrado na Cosmos)' : statusCosmos === 429 ? ' (limite de chamadas da API excedido)' : '';
return { descricao: null, debug: 'HTTP ' + statusCosmos + motivo + ' - ' + respostaCosmos.getContentText().substring(0, 300) };
} catch (erroCosmos) {
return { descricao: null, debug: 'Erro ao consultar: ' + erroCosmos.toString() };
}
}
// ===== FIM ITEM 24 =====
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
// ===== ITEM 24: diagnóstico da consulta à Cosmos, devolvido ao navegador para facilitar depuração =====
var cosmosDebug = "Não consultada (Open Food Facts já retornou um nome válido)";
// ===== FIM ITEM 24 =====
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
// ===== ITEM 24: tenta a Cosmos Bluesoft como segunda fonte, caso a Open Food Facts não tenha retornado um nome útil =====
var precisaTentarCosmos = (descricao === "Não localizado na internet") ||
descricao.indexOf("Não localizado") === 0 ||
descricao.indexOf("Erro na Busca") === 0 ||
descricao === "Produto sem nome cadastrado";
if (precisaTentarCosmos) {
var resultadoCosmos = buscarProdutoCosmos(codigoLimpo);
cosmosDebug = resultadoCosmos.debug;
if (resultadoCosmos.descricao) {
descricao = resultadoCosmos.descricao;
}
} else {
cosmosDebug = "Não consultada (Open Food Facts já retornou um nome válido: \"" + descricao + "\")";
}
// ===== FIM ITEM 24 =====
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
subcategoria: subcategoriaEstoque,
// ===== FIM ITEM 23 =====
// ===== ITEM 24: diagnóstico da consulta à Cosmos (visível no console do navegador) =====
cosmosDebug: cosmosDebug
// ===== FIM ITEM 24 =====
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