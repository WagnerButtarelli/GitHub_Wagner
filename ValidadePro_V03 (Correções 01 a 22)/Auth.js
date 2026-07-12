// ==================== AUTENTICAÇÃO ====================
function autenticarUsuario(usuario, senha) {
try {
var sheet = getSheetByName('Acessos');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
var nome = dados[i][0];
var user = dados[i][1];
var pass = dados[i][2];
var setor = dados[i][3];
var perfil = dados[i][4] || 'Consultor';
var status = dados[i][5];
var loja = dados[i][7] || '';
if (user === usuario && pass.toString() === senha) {
if (status !== 'Ativo') {
return { sucesso: false, mensagem: 'Usuário Inativo' };
}
sheet.getRange(i + 1, 7).setValue(new Date());
return {
sucesso: true,
usuario: {
nome: nome,
usuario: user,
setor: setor,
perfil: perfil,
status: status,
loja: loja
}
};
}
}
return { sucesso: false, mensagem: 'Usuário ou senha inválidos!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== CONTROLE DE PERMISSÕES (dinâmico, lido da planilha Perfis) ====================
var SECOES_PERMISSAO = ['Dashboard', 'Estoque', 'Relatorios', 'Scanner', 'Configuracoes'];
function getPermissoesUsuario(perfil) {
var sheet = getSheetByName('Perfis');
var dados = sheet.getDataRange().getValues();
var linhaPerfil = null;
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === perfil) { linhaPerfil = dados[i]; break; }
}
// Perfil não encontrado na planilha: aplica somente-visualização por segurança
if (!linhaPerfil) linhaPerfil = [perfil, 'VIEW', 'VIEW', 'VIEW', 'VIEW', 'VIEW'];
var resultado = {};
for (var s = 0; s < SECOES_PERMISSAO.length; s++) {
var nivel = linhaPerfil[s + 1] || 'NONE';
resultado[SECOES_PERMISSAO[s]] = { nivel: nivel, permitido: nivel !== 'NONE' };
}
return resultado;
}
function verificarPermissao(perfil, secao, acao) {
var perms = getPermissoesUsuario(perfil);
var permissao = perms[secao];
if (!permissao || !permissao.permitido) {
return { permitido: false, mensagem: 'Acesso negado para esta seção.' };
}
if (acao === 'delete' && permissao.nivel !== 'FULL') {
return { permitido: false, mensagem: 'Permissão insuficiente para excluir. Somente perfis com Acesso Total podem excluir.' };
}
if ((acao === 'create' || acao === 'update') && permissao.nivel === 'VIEW') {
return { permitido: false, mensagem: 'Permissão insuficiente. Seu perfil permite apenas visualização.' };
}
return { permitido: true, nivel: permissao.nivel };
}
// ==================== PERFIS DE ACESSO (CRUD) ====================
function getPerfis() {
try {
var sheet = getSheetByName('Perfis');
var dados = sheet.getDataRange().getValues();
var perfis = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][0]) {
perfis.push({
perfil: dados[i][0],
Dashboard: dados[i][1] || 'NONE',
Estoque: dados[i][2] || 'NONE',
Relatorios: dados[i][3] || 'NONE',
Scanner: dados[i][4] || 'NONE',
Configuracoes: dados[i][5] || 'NONE',
linha: i + 1
});
}
}
perfis.sort(function(a, b) { return String(a.perfil).localeCompare(String(b.perfil), 'pt-BR', { sensitivity: 'base' }); });
return { sucesso: true, perfis: perfis };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarPerfil(dadosPerfil, perfilLogado) {
try {
if (perfilLogado) {
var check = verificarPermissao(perfilLogado, 'Configuracoes', dadosPerfil.linha ? 'update' : 'create');
if (!check.permitido) return check;
}
var nome = (dadosPerfil.perfil || '').toString().trim();
if (!nome) return { sucesso: false, mensagem: 'Informe o nome do perfil!' };
var sheet = getSheetByName('Perfis');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][0] === nome && (i + 1) !== dadosPerfil.linha) {
return { sucesso: false, mensagem: 'Já existe um perfil com esse nome!' };
}
}
var linhaValores = [
nome,
dadosPerfil.Dashboard || 'NONE',
dadosPerfil.Estoque || 'NONE',
dadosPerfil.Relatorios || 'NONE',
dadosPerfil.Scanner || 'NONE',
dadosPerfil.Configuracoes || 'NONE'
];
if (dadosPerfil.linha) {
var nomeAntigo = dados[dadosPerfil.linha - 1][0];
sheet.getRange(dadosPerfil.linha, 1, 1, 6).setValues([linhaValores]);
if (nomeAntigo !== nome) {
var sheetAcessos = getSheetByName('Acessos');
var dadosAcessos = sheetAcessos.getDataRange().getValues();
for (var k = 1; k < dadosAcessos.length; k++) {
if (dadosAcessos[k][4] === nomeAntigo) sheetAcessos.getRange(k + 1, 5).setValue(nome);
}
}
} else {
sheet.appendRow(linhaValores);
}
return { sucesso: true, mensagem: 'Perfil salvo com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirPerfil(linha, perfilLogado) {
try {
if (perfilLogado) {
var check = verificarPermissao(perfilLogado, 'Configuracoes', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Perfis');
var dados = sheet.getDataRange().getValues();
if (!dados[linha - 1]) return { sucesso: false, mensagem: 'Perfil não encontrado!' };
var nomePerfil = dados[linha - 1][0];
var sheetAcessos = getSheetByName('Acessos');
var dadosAcessos = sheetAcessos.getDataRange().getValues();
for (var k = 1; k < dadosAcessos.length; k++) {
if (dadosAcessos[k][4] === nomePerfil) {
return { sucesso: false, mensagem: 'Não é possível excluir: existem usuários cadastrados com este perfil.' };
}
}
sheet.deleteRow(linha);
return { sucesso: true, mensagem: 'Perfil excluído com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== USUÁRIOS DO SISTEMA (CRUD) ====================
function getUsuarios() {
try {
var sheet = getSheetByName('Acessos');
var dados = sheet.getDataRange().getValues();
var usuarios = [];
for (var i = 1; i < dados.length; i++) {
if (dados[i][1]) {
var ultimoAcesso = dados[i][6];
usuarios.push({
nome: dados[i][0],
usuario: dados[i][1],
setor: dados[i][3] || '',
perfil: dados[i][4] || '',
status: dados[i][5] || 'Ativo',
ultimoAcesso: (ultimoAcesso instanceof Date) ? Utilities.formatDate(ultimoAcesso, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
loja: dados[i][7] || '',
linha: i + 1
});
}
}
usuarios.sort(function(a, b) { return String(a.nome).localeCompare(String(b.nome), 'pt-BR', { sensitivity: 'base' }); });
return { sucesso: true, usuarios: usuarios };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function salvarUsuario(dadosUsuario, perfilLogado) {
try {
if (perfilLogado) {
var check = verificarPermissao(perfilLogado, 'Configuracoes', dadosUsuario.linha ? 'update' : 'create');
if (!check.permitido) return check;
}
var nome = (dadosUsuario.nome || '').toString().trim();
var login = (dadosUsuario.usuario || '').toString().trim();
if (!nome) return { sucesso: false, mensagem: 'Informe o nome!' };
if (!login) return { sucesso: false, mensagem: 'Informe o usuário (login)!' };
if (!dadosUsuario.perfil) return { sucesso: false, mensagem: 'Selecione um perfil!' };
if (!dadosUsuario.linha && !dadosUsuario.senha) return { sucesso: false, mensagem: 'Informe a senha!' };
var sheet = getSheetByName('Acessos');
var dados = sheet.getDataRange().getValues();
for (var i = 1; i < dados.length; i++) {
if (dados[i][1] === login && (i + 1) !== dadosUsuario.linha) {
return { sucesso: false, mensagem: 'Já existe um usuário com esse login!' };
}
}
var status = dadosUsuario.status === 'Inativo' ? 'Inativo' : 'Ativo';
if (dadosUsuario.linha) {
var senhaAtual = dados[dadosUsuario.linha - 1][2];
var senhaFinal = dadosUsuario.senha ? dadosUsuario.senha : senhaAtual;
sheet.getRange(dadosUsuario.linha, 1, 1, 6).setValues([[nome, login, senhaFinal, dadosUsuario.setor || '', dadosUsuario.perfil, status]]);
sheet.getRange(dadosUsuario.linha, 8).setValue(dadosUsuario.loja || '');
} else {
sheet.appendRow([nome, login, dadosUsuario.senha, dadosUsuario.setor || '', dadosUsuario.perfil, status, '', dadosUsuario.loja || '']);
}
return { sucesso: true, mensagem: 'Usuário salvo com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
function excluirUsuario(linha, perfilLogado) {
try {
if (perfilLogado) {
var check = verificarPermissao(perfilLogado, 'Configuracoes', 'delete');
if (!check.permitido) return check;
}
var sheet = getSheetByName('Acessos');
var dados = sheet.getDataRange().getValues();
if (!dados[linha - 1]) return { sucesso: false, mensagem: 'Usuário não encontrado!' };
var totalUsuarios = 0;
for (var i = 1; i < dados.length; i++) { if (dados[i][1]) totalUsuarios++; }
if (totalUsuarios <= 1) return { sucesso: false, mensagem: 'Não é possível excluir o único usuário do sistema!' };
sheet.deleteRow(linha);
return { sucesso: true, mensagem: 'Usuário excluído com sucesso!' };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}