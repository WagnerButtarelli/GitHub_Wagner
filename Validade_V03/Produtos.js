// ==================== PRODUTOS ====================
function getProdutos() {
  try {
    var sheet = getSheetByName('Estoque');
    if (sheet.getLastColumn() < 11) sheet.getRange(1, 11).setValue('Quebra');
    var dados = sheet.getDataRange().getValues();
    var produtos = [];
    var hoje = new Date(); 
    hoje.setHours(0, 0, 0, 0);
    
    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][0] && !dados[i][1]) continue;
      
      var dataValidade = dados[i][3];
      var diasRestantes = null;
      var status = 'seguro';
      var sortDate = 9999999999999;
      
      if (dataValidade && dataValidade instanceof Date) {
        var dv = new Date(dataValidade); 
        dv.setHours(0, 0, 0, 0);
        sortDate = dv.getTime();
        diasRestantes = Math.ceil((dv - hoje) / (1000 * 60 * 60 * 24));
        
        if (diasRestantes < 0) status = 'vencido';
        else if (diasRestantes <= 3) status = 'supercritico';
        else if (diasRestantes <= 7) status = 'critico';
        else if (diasRestantes <= 30) status = 'alerta';
      }
      
      produtos.push({
        codigo: dados[i][0],
        produto: dados[i][1],
        quantidade: dados[i][2],
        dataValidade: dataValidade ? Utilities.formatDate(new Date(dataValidade), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        categoria: dados[i][4],
        subcategoria: dados[i][5] || '',
        precoCusto: dados[i][6],
        precoVenda: dados[i][7],
        fornecedor: dados[i][8],
        estabelecimento: dados[i][9],
        quebra: dados[i][10] || 'Não',
        diasRestantes: diasRestantes,
        status: status,
        linha: i + 1,
        _sortDate: sortDate
      });
    }
    
    produtos.sort(function(a, b) { return a._sortDate - b._sortDate; });
    for (var k = 0; k < produtos.length; k++) delete produtos[k]._sortDate;
    
    return { sucesso: true, produtos: produtos };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function salvarProduto(produto, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', produto.linha ? 'update' : 'create').permitido) {
      return { permitido: false };
    }
    var sheet = getSheetByName('Estoque');
    var row = [
      produto.codigo, produto.produto, produto.quantidade,
      produto.dataValidade ? new Date(produto.dataValidade) : '',
      produto.categoria, produto.subcategoria || '',
      produto.precoCusto || 0, produto.precoVenda || 0,
      produto.fornecedor, produto.estabelecimento,
      produto.quebra || 'Não'
    ];
    
    if (produto.linha) sheet.getRange(produto.linha, 1, 1, 11).setValues([row]);
    else sheet.appendRow(row);
    
    return { sucesso: true, mensagem: 'Produto salvo!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function atualizarQuebra(linha, valor, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'update').permitido) return { permitido: false };
    getSheetByName('Estoque').getRange(linha, 11).setValue(valor);
    return { sucesso: true };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function excluirProduto(linha, perfil) {
  try {
    if (perfil && !verificarPermissao(perfil, 'Estoque', 'delete').permitido) return { permitido: false };
    getSheetByName('Estoque').deleteRow(linha);
    return { sucesso: true, mensagem: 'Excluído!' };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}
