// ==================== CATEGORIAS E SUBCATEGORIAS (fonte única: Categoria_Subcategoria) ====================
function getCategorias() {
try {
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
var vistos = {};
var categorias = [];
for (var i = 1; i < dados.length; i++) {
var cat = dados[i][0];
if (cat && !vistos[cat]) { vistos[cat] = true; categorias.push(cat); }
}
return { sucesso: true, categorias: ordenarAlfabetico(categorias) };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarCategoria(categoria, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'create');
if (!check.permitido) return check;
}
if (!categoria) return { sucesso: false, mensagem: 'Informe o nome da categoria!' };
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoria) return { sucesso: false, mensagem: 'Categoria já existe!' };
}
sheet.appendRow([categoria, '']);
return { sucesso: true, mensagem: 'Categoria cadastrada com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function editarCategoria(categoriaAntiga, categoriaNova, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'update');
if (!check.permitido) return check;
}
if (!categoriaNova) return { sucesso: false, mensagem: 'Informe o novo nome da categoria!' };
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
if (categoriaNova !== categoriaAntiga) {
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoriaNova) return { sucesso: false, mensagem: 'Já existe uma categoria com esse nome!' };
}
}
var encontrada = false;
for (var j = 1; j < dados.length; j++) {
if (dados[j][0] === categoriaAntiga) { sheet.getRange(j + 1, 1).setValue(categoriaNova); encontrada = true; }
}
if (!encontrada) return { sucesso: false, mensagem: 'Categoria não encontrada!' };
var sheetEstoque = getSheetByName('Estoque');
var dadosEstoque = sheetEstoque.getDataRange().getValues();
for (var k = 1; k < dadosEstoque.length; k++) {
if (dadosEstoque[k][4] === categoriaAntiga) sheetEstoque.getRange(k + 1, 5).setValue(categoriaNova);
}
return { sucesso: true, mensagem: 'Categoria atualizada com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirCategoria(categoria, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
var linhas = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoria) linhas.push(i + 1);
}
if (linhas.length === 0) return { sucesso: false, mensagem: 'Categoria não encontrada!' };
for (var k = linhas.length - 1; k >= 0; k--) sheet.deleteRow(linhas[k]);
return { sucesso: true, mensagem: 'Categoria e subcategorias vinculadas foram excluídas!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function getSubcategorias() {
try {
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
var subcategorias = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] && dados[i][1]) {
subcategorias.push({ categoria: dados[i][0], subcategoria: dados[i][1] });
}
}
subcategorias.sort(function(a, b) { return String(a.subcategoria).localeCompare(String(b.subcategoria), 'pt-BR', { sensitivity: 'base' }); });
return { sucesso: true, subcategorias: subcategorias };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarSubcategoria(categoria, subcategoria, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'create');
if (!check.permitido) return check;
}
if (!categoria) return { sucesso: false, mensagem: 'Selecione uma categoria vinculada!' };
if (!subcategoria) return { sucesso: false, mensagem: 'Informe o nome da subcategoria!' };
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoria && dados[i][1] === subcategoria) {
return { sucesso: false, mensagem: 'Subcategoria já existe para esta categoria!' };
}
}
sheet.appendRow([categoria, subcategoria]);
return { sucesso: true, mensagem: 'Subcategoria cadastrada com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function editarSubcategoria(categoria, subcategoriaAntiga, subcategoriaNova, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'update');
if (!check.permitido) return check;
}
if (!subcategoriaNova) return { sucesso: false, mensagem: 'Informe o novo nome da subcategoria!' };
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
if (subcategoriaNova !== subcategoriaAntiga) {
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoria && dados[i][1] === subcategoriaNova) return { sucesso: false, mensagem: 'Já existe uma subcategoria com esse nome para esta categoria!' };
}
}
var encontrada = false;
for (var j = 1; j < dados.length; j++) {
if (dados[j][0] === categoria && dados[j][1] === subcategoriaAntiga) { sheet.getRange(j + 1, 2).setValue(subcategoriaNova); encontrada = true; }
}
if (!encontrada) return { sucesso: false, mensagem: 'Subcategoria não encontrada!' };
var sheetEstoque = getSheetByName('Estoque');
var dadosEstoque = sheetEstoque.getDataRange().getValues();
for (var k = 1; k < dadosEstoque.length; k++) {
if (dadosEstoque[k][4] === categoria && dadosEstoque[k][5] === subcategoriaAntiga) sheetEstoque.getRange(k + 1, 6).setValue(subcategoriaNova);
}
return { sucesso: true, mensagem: 'Subcategoria atualizada com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirSubcategoria(categoria, subcategoria, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Categoria_Subcategoria');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === categoria && dados[i][1] === subcategoria) {
sheet.deleteRow(i + 1);
return { sucesso: true, mensagem: 'Subcategoria excluída!' };
}
}
return { sucesso: false, mensagem: 'Subcategoria não encontrada!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== ESTABELECIMENTOS ====================
function getEstabelecimentos() {
try {
var sheet = getSheetByName('Estabelecimentos');
var dados = sheet.getDataRange().getValues();
var estabelecimentos = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][0]) estabelecimentos.push(dados[i][0]);
}
return { sucesso: true, estabelecimentos: ordenarAlfabetico(estabelecimentos) };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function editarEstabelecimento(antigo, novo, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'update');
if (!check.permitido) return check;
}
if (!novo) return { sucesso: false, mensagem: 'Informe o novo nome do estabelecimento!' };
var sheet = getSheetByName('Estabelecimentos');
var dados = sheet.getDataRange().getValues();
if (novo !== antigo) {
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === novo) return { sucesso: false, mensagem: 'Já existe um estabelecimento com esse nome!' };
}
}
var encontrado = false;
for (var j = 1; j < dados.length; j++) {
if (dados[j][0] === antigo) { sheet.getRange(j + 1, 1).setValue(novo); encontrado = true; }
}
if (!encontrado) return { sucesso: false, mensagem: 'Estabelecimento não encontrado!' };
var sheetEstoque = getSheetByName('Estoque');
var dadosEstoque = sheetEstoque.getDataRange().getValues();
for (var k = 1; k < dadosEstoque.length; k++) {
if (dadosEstoque[k][9] === antigo) sheetEstoque.getRange(k + 1, 10).setValue(novo);
}
return { sucesso: true, mensagem: 'Estabelecimento atualizado com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarEstabelecimento(estabelecimento, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'create');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Estabelecimentos');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === estabelecimento) return { sucesso: false, mensagem: 'Estabelecimento já existe!' };
}
sheet.appendRow([estabelecimento]);
return { sucesso: true, mensagem: 'Estabelecimento cadastrado com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirEstabelecimento(estabelecimento, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Estabelecimentos');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === estabelecimento) {
sheet.deleteRow(i + 1);
return { sucesso: true, mensagem: 'Estabelecimento excluído!' };
}
}
return { sucesso: false, mensagem: 'Estabelecimento não encontrado!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== FORNECEDORES ====================
function getFornecedores() {
try {
var sheet = getSheetByName('Fornecedores');
var dados = sheet.getDataRange().getValues();
var fornecedores = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][0]) fornecedores.push(dados[i][0]);
}
return { sucesso: true, fornecedores: ordenarAlfabetico(fornecedores) };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function editarFornecedor(antigo, novo, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'update');
if (!check.permitido) return check;
}
if (!novo) return { sucesso: false, mensagem: 'Informe o novo nome do fornecedor!' };
var sheet = getSheetByName('Fornecedores');
var dados = sheet.getDataRange().getValues();
if (novo !== antigo) {
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === novo) return { sucesso: false, mensagem: 'Já existe um fornecedor com esse nome!' };
}
}
var encontrado = false;
for (var j = 1; j < dados.length; j++) {
if (dados[j][0] === antigo) { sheet.getRange(j + 1, 1).setValue(novo); encontrado = true; }
}
if (!encontrado) return { sucesso: false, mensagem: 'Fornecedor não encontrado!' };
var sheetEstoque = getSheetByName('Estoque');
var dadosEstoque = sheetEstoque.getDataRange().getValues();
for (var k = 1; k < dadosEstoque.length; k++) {
if (dadosEstoque[k][8] === antigo) sheetEstoque.getRange(k + 1, 9).setValue(novo);
}
return { sucesso: true, mensagem: 'Fornecedor atualizado com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarFornecedor(fornecedor, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'create');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Fornecedores');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === fornecedor) return { sucesso: false, mensagem: 'Fornecedor já existe!' };
}
sheet.appendRow([fornecedor]);
return { sucesso: true, mensagem: 'Fornecedor cadastrado com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirFornecedor(fornecedor, perfil) {
try {
if (perfil) {
var check = verificarPermissao(perfil, 'Estoque', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Fornecedores');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === fornecedor) {
sheet.deleteRow(i + 1);
return { sucesso: true, mensagem: 'Fornecedor excluído!' };
}
}
return { sucesso: false, mensagem: 'Fornecedor não encontrado!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
