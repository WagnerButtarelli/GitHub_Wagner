// ==================== PRODUTOS ====================
function parseDataLocal(dataStr) {
if (!dataStr) return '';
var partes = String(dataStr).split('-');
return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}
function getProdutos() {
try {
var sheet = getSheetByName('Estoque');
var dados = sheet.getDataRange().getValues();
var produtos = [];
var hoje = new Date();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0]) {
var dataValidade = dados[i][3];
var diasRestantes = null;
var status = 'seguro';
if (dataValidade && dataValidade instanceof Date) {
var diffTime = dataValidade - hoje;
diasRestantes = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
if (diasRestantes < 0) status = 'vencido';
else if (diasRestantes <= 3) status = 'supercritico';
else if (diasRestantes <= 7) status = 'critico';
else if (diasRestantes <= 30) status = 'alerta';
else status = 'seguro';
}
produtos.push({
codigo: dados[i][0],
produto: dados[i][1],
quantidade: dados[i][2],
dataValidade: dataValidade ? Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
categoria: dados[i][4],
subcategoria: dados[i][5] || '',
precoCusto: dados[i][6],
precoVenda: dados[i][7],
fornecedor: dados[i][8],
estabelecimento: dados[i][9],
quebra: dados[i][10] === 'Sim' ? 'Sim' : 'Não',
diasRestantes: diasRestantes,
status: status,
linha: i + 1
});
}
}
// Ordena sempre pela data de validade mais próxima de vencer (produtos sem data ficam ao final)
produtos.sort(function(a, b) {
if (!a.dataValidade && !b.dataValidade) return 0;
if (!a.dataValidade) return 1;
if (!b.dataValidade) return -1;
return new Date(a.dataValidade) - new Date(b.dataValidade);
});
return { sucesso: true, produtos: produtos };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarProduto(produto, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', produto.linha ? 'update' : 'create');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Estoque');
var quebraValor = produto.quebra === 'Sim' ? 'Sim' : 'Não';
if (produto.linha) {
sheet.getRange(produto.linha, 1, 1, 11).setValues([[
produto.codigo,
produto.produto,
produto.quantidade,
produto.dataValidade ? parseDataLocal(produto.dataValidade) : '',
produto.categoria,
produto.subcategoria || '',
produto.precoCusto,
produto.precoVenda,
produto.fornecedor,
produto.estabelecimento,
quebraValor
]]);
} else {
sheet.appendRow([
produto.codigo,
produto.produto,
produto.quantidade,
produto.dataValidade ? parseDataLocal(produto.dataValidade) : '',
produto.categoria,
produto.subcategoria || '',
produto.precoCusto,
produto.precoVenda,
produto.fornecedor,
produto.estabelecimento,
quebraValor
]);
}
return { sucesso: true, mensagem: 'Produto salvo com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirProduto(linha, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Estoque');
sheet.deleteRow(linha);
return { sucesso: true, mensagem: 'Produto excluído com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}