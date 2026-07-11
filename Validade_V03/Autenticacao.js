// ==================== AUTENTICAÇÃO E PERMISSÕES ====================
function autenticarUsuario(usuario, senha) {
  try {
    var sheet = getSheetByName('Acessos');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][1] === usuario && dados[i][2].toString() === senha) {
        if (dados[i][5] !== 'Ativo') return { sucesso: false, mensagem: 'Usuário Inativo' };
        sheet.getRange(i + 1, 7).setValue(new Date());
        return { 
          sucesso: true, 
          usuario: { 
            nome: dados[i][0], 
            usuario: dados[i][1], 
            setor: dados[i][3], 
            perfil: dados[i][4] || 'Consultor', 
            status: dados[i][5] 
          } 
        };
      }
    }
    return { sucesso: false, mensagem: 'Usuário ou senha inválidos!' };
  } catch (e) { 
    return { sucesso: false, mensagem: 'Erro: ' + e.toString() }; 
  }
}

function getPermissoesUsuario(perfil) {
  var permissoes = {
    'Gerente': { 'Dashboard': 'FULL', 'Estoque': 'FULL', 'Relatorios': 'FULL', 'Scanner': 'FULL', 'Configuracoes': 'FULL' },
    'Consultor': { 'Dashboard': 'VIEW', 'Estoque': 'VIEW', 'Relatorios': 'VIEW', 'Scanner': 'VIEW', 'Configuracoes': 'VIEW' },
    'Cadastrador': { 'Dashboard': 'NONE', 'Estoque': 'EDIT', 'Relatorios': 'NONE', 'Scanner': 'EDIT', 'Configuracoes': 'NONE' }
  };
  var perms = permissoes[perfil] || permissoes['Consultor'];
  var resultado = {};
  for (var secao in perms) {
    resultado[secao] = { nivel: perms[secao], permitido: perms[secao] !== 'NONE' };
  }
  return resultado;
}

function verificarPermissao(perfil, secao, acao) {
  var perms = getPermissoesUsuario(perfil);
  var permissao = perms[secao];
  if (!permissao || !permissao.permitido) return { permitido: false, mensagem: 'Acesso negado.' };
  if (acao === 'delete' && permissao.nivel !== 'FULL') return { permitido: false, mensagem: 'Somente Gerentes podem excluir.' };
  if ((acao === 'create' || acao === 'update') && permissao.nivel === 'VIEW') return { permitido: false, mensagem: 'Apenas visualização.' };
  return { permitido: true, nivel: permissao.nivel };
}