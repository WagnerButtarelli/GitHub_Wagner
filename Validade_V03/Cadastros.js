// ==================== CATEGORIA E SUBCATEGORIA (UNIFICADO) ====================
function getCategorias() {
  try {
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    var cats = {};
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0]) cats[dados[i][0]] = true;
    }
    var lista = Object.keys(cats);
    lista.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
    return { sucesso: true, categorias: lista };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function getSubcategorias() {
  try {
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    var subs = [];
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] && dados[i][1]) {
        subs.push({ categoria: dados[i][0], subcategoria: dados[i][1] });
      }
    }
    subs.sort(function(a, b) {
      var c = a.categoria.localeCompare(b.categoria, 'pt-BR');
      return c !== 0 ? c : a.subcategoria.localeCompare(b.subcategoria, 'pt-BR');
    });
    return { sucesso: true, subcategorias: subs };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function salvarCategoria(categoria, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'create').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === categoria) return { sucesso: false, mensagem: 'Categoria já existe!' };
    }
    sheet.appendRow([categoria, '']);
    return { sucesso: true, mensagem: 'Categoria cadastrada!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function editarCategoria(nomeAntigo, nomeNovo, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'update').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === nomeAntigo) sheet.getRange(i + 1, 1).setValue(nomeNovo);
    }
    var estSheet = getSheetByName('Estoque');
    var estDados = estSheet.getDataRange().getValues();
    for (var j = 1; j < estDados.length; j++) {
      if (estDados[j][4] === nomeAntigo) estSheet.getRange(j + 1, 5).setValue(nomeNovo);
    }
    return { sucesso: true, mensagem: 'Categoria atualizada!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function excluirCategoria(categoria, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'delete').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) {
      if (dados[i][0] === categoria) sheet.deleteRow(i + 1);
    }
    return { sucesso: true, mensagem: 'Categoria excluída!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function salvarSubcategoria(categoria, subcategoria, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'create').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === categoria && dados[i][1] === subcategoria) {
        return { sucesso: false, mensagem: 'Subcategoria já existe!' };
      }
    }
    sheet.appendRow([categoria, subcategoria]);
    return { sucesso: true, mensagem: 'Subcategoria cadastrada!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function editarSubcategoria(categoria, nomeAntigo, nomeNovo, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'update').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === categoria && dados[i][1] === nomeAntigo) {
        sheet.getRange(i + 1, 2).setValue(nomeNovo);
        var estSheet = getSheetByName('Estoque');
        var estDados = estSheet.getDataRange().getValues();
        for (var j = 1; j < estDados.length; j++) {
          if (estDados[j][4] === categoria && estDados[j][5] === nomeAntigo) {
            estSheet.getRange(j + 1, 6).setValue(nomeNovo);
          }
        }
        return { sucesso: true, mensagem: 'Subcategoria atualizada!' };
      }
    }
    return { sucesso: false, mensagem: 'Não encontrada!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function excluirSubcategoria(categoria, subcategoria, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'delete').permitido) return { permitido: false };
    var sheet = getSheetByName('Categoria_Subcategoria');
    var dados = sheet.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) {
      if (dados[i][0] === categoria && dados[i][1] === subcategoria) {
        sheet.deleteRow(i + 1);
      }
    }
    return { sucesso: true, mensagem: 'Subcategoria excluída!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

// ==================== ESTABELECIMENTOS ====================
function getEstabelecimentos() {
  try {
    var sheet = getSheetByName('Estabelecimentos');
    var dados = sheet.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < dados.length; i++) if (dados[i][0]) lista.push(dados[i][0]);
    lista.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
    return { sucesso: true, estabelecimentos: lista };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function salvarEstabelecimento(nome, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'create').permitido) return { permitido: false };
    var sheet = getSheetByName('Estabelecimentos');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) if (dados[i][0] === nome) return { sucesso: false, mensagem: 'Já existe!' };
    sheet.appendRow([nome]);
    return { sucesso: true, mensagem: 'Cadastrado!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function editarEstabelecimento(nomeAntigo, nomeNovo, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'update').permitido) return { permitido: false };
    var sheet = getSheetByName('Estabelecimentos');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === nomeAntigo) {
        sheet.getRange(i + 1, 1).setValue(nomeNovo);
        var estSheet = getSheetByName('Estoque');
        var estDados = estSheet.getDataRange().getValues();
        for (var j = 1; j < estDados.length; j++) {
          if (estDados[j][9] === nomeAntigo) estSheet.getRange(j + 1, 10).setValue(nomeNovo);
        }
        return { sucesso: true, mensagem: 'Atualizado!' };
      }
    }
    return { sucesso: false, mensagem: 'Não encontrado!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function excluirEstabelecimento(nome, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'delete').permitido) return { permitido: false };
    var sheet = getSheetByName('Estabelecimentos');
    var dados = sheet.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) if (dados[i][0] === nome) sheet.deleteRow(i + 1);
    return { sucesso: true, mensagem: 'Excluído!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

// ==================== FORNECEDORES ====================
function getFornecedores() {
  try {
    var sheet = getSheetByName('Fornecedores');
    var dados = sheet.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < dados.length; i++) if (dados[i][0]) lista.push(dados[i][0]);
    lista.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
    return { sucesso: true, fornecedores: lista };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function salvarFornecedor(nome, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'create').permitido) return { permitido: false };
    var sheet = getSheetByName('Fornecedores');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) if (dados[i][0] === nome) return { sucesso: false, mensagem: 'Já existe!' };
    sheet.appendRow([nome]);
    return { sucesso: true, mensagem: 'Cadastrado!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function editarFornecedor(nomeAntigo, nomeNovo, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'update').permitido) return { permitido: false };
    var sheet = getSheetByName('Fornecedores');
    var dados = sheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === nomeAntigo) {
        sheet.getRange(i + 1, 1).setValue(nomeNovo);
        var estSheet = getSheetByName('Estoque');
        var estDados = estSheet.getDataRange().getValues();
        for (var j = 1; j < estDados.length; j++) {
          if (estDados[j][8] === nomeAntigo) estSheet.getRange(j + 1, 9).setValue(nomeNovo);
        }
        return { sucesso: true, mensagem: 'Atualizado!' };
      }
    }
    return { sucesso: false, mensagem: 'Não encontrado!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function excluirFornecedor(nome, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'delete').permitido) return { permitido: false };
    var sheet = getSheetByName('Fornecedores');
    var dados = sheet.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) if (dados[i][0] === nome) sheet.deleteRow(i + 1);
    return { sucesso: true, mensagem: 'Excluído!' };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}