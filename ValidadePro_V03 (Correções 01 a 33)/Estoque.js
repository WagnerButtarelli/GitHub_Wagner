// ==================== PRODUTOS ====================
function parseDataLocal(dataStr) {
  if (!dataStr) return '';
  var partes = String(dataStr).split('-');
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}

function getProdutos(perfil, loja) {
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
        var dataCadastro = dados[i][13];
        var dataRebaixa = dados[i][15];
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
          cadastradoPor: dados[i][11] || '',
          funcao: dados[i][12] || '',
          dataCadastro: dataCadastro instanceof Date ? Utilities.formatDate(dataCadastro, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          resolvido: dados[i][14] === 'Sim' ? 'Sim' : 'Não',
          dataRebaixa: dataRebaixa instanceof Date ? Utilities.formatDate(dataRebaixa, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          diasRestantes: diasRestantes,
          status: status,
          linha: i + 1
        });
      }
    }

    // ===== FILTRO POR LOJA (se não for Master) =====
    if (perfil !== 'Master' && loja) {
      produtos = produtos.filter(function(p) {
        return p.estabelecimento === loja;
      });
    }

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

function salvarProduto(produto, perfil, loja) {
  try {
    var linhaAtual = produto.linha ? parseInt(produto.linha, 10) : null;

    if (perfil) {
      var check = verificarPermissao(perfil, 'Estoque', linhaAtual ? 'update' : 'create');
      if (!check.permitido) return check;
    }

    // --- Validação e limpeza do código de barras ---
    var codigoLimpo = String(produto.codigo || '')
      .replace(/\s/g, '')
      .replace(/[^0-9]/g, '');
    if (!codigoLimpo) {
      return { sucesso: false, mensagem: 'Código de barras inválido! Utilize apenas números.' };
    }
    produto.codigo = codigoLimpo;

    // ===== Força a loja do usuário se não for Master =====
    if (perfil !== 'Master' && loja) {
      produto.estabelecimento = loja;
    }

    var sheet = getSheetByName('Estoque');
    var dados = sheet.getDataRange().getValues();

    // --- Verificação de duplicidade ---
    var dataValidade = produto.dataValidade ? parseDataLocal(produto.dataValidade) : '';
    var dataValidadeStr = dataValidade ? Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    var dataValidadeDisplay = dataValidade ? Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'sem data';

    for (var i = 1; i < dados.length; i++) {
      if (linhaAtual && (i + 1) === linhaAtual) continue;
      var codigoExistente = String(dados[i][0]).trim();
      var dataExistente = dados[i][3];
      var dataExistenteStr = dataExistente instanceof Date ? Utilities.formatDate(dataExistente, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      if (codigoExistente === codigoLimpo && dataExistenteStr === dataValidadeStr) {
        var nomeExistente = dados[i][1] || 'produto sem nome';
        return {
          sucesso: false,
          mensagem: 'O produto "' + nomeExistente + '", código ' + codigoLimpo + ', na data de validade ' + dataValidadeDisplay + ', já existe cadastrado !'
        };
      }
    }

    var quebraValor = produto.quebra === 'Sim' ? 'Sim' : 'Não';
    var resolvidoValor = produto.resolvido === 'Sim' ? 'Sim' : 'Não';

    if (linhaAtual) {
      // Edição: não altera Cadastrado Por, Função e Data Cadastro
      sheet.getRange(linhaAtual, 1, 1, 11).setValues([[
        produto.codigo,
        produto.produto,
        produto.quantidade,
        dataValidade,
        produto.categoria,
        produto.subcategoria || '',
        produto.precoCusto,
        produto.precoVenda,
        produto.fornecedor,
        produto.estabelecimento,
        quebraValor
      ]]);
      sheet.getRange(linhaAtual, 15).setValue(resolvidoValor);

      // Atualiza Data Rebaixa
      var linhaAnterior = dados[linhaAtual - 1] || [];
      var quebraAnterior = linhaAnterior[10] === 'Sim' ? 'Sim' : 'Não';
      var dataRebaixaAnterior = linhaAnterior[15] || '';
      var dataRebaixaValor;
      if (quebraValor === 'Sim') {
        dataRebaixaValor = (quebraAnterior === 'Sim' && dataRebaixaAnterior) ? dataRebaixaAnterior : new Date();
      } else {
        dataRebaixaValor = '';
      }
      sheet.getRange(linhaAtual, 16).setValue(dataRebaixaValor);
    } else {
      var cadastradoPor = String(produto.cadastradoPor || '').trim();
      var funcaoUsuario = String(produto.funcao || '').trim();
      var dataRebaixaNovo = quebraValor === 'Sim' ? new Date() : '';
      sheet.appendRow([
        produto.codigo,
        produto.produto,
        produto.quantidade,
        dataValidade,
        produto.categoria,
        produto.subcategoria || '',
        produto.precoCusto,
        produto.precoVenda,
        produto.fornecedor,
        produto.estabelecimento,
        quebraValor,
        cadastradoPor,
        funcaoUsuario,
        new Date(),
        resolvidoValor,
        dataRebaixaNovo
      ]);
    }
    return { sucesso: true, mensagem: 'Produto salvo com sucesso!' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
  }
}

function excluirProduto(linha, perfil, loja) {
  try {
    if (perfil) {
      var check = verificarPermissao(perfil, 'Estoque', 'delete');
      if (!check.permitido) return check;
    }

    var sheet = getSheetByName('Estoque');
    var dados = sheet.getDataRange().getValues();
    if (!dados[linha - 1]) {
      return { sucesso: false, mensagem: 'Produto não encontrado!' };
    }

    // ===== Verifica se o produto pertence à loja (se não for Master) =====
    if (perfil !== 'Master' && loja) {
      var estabelecimentoProduto = dados[linha - 1][9] || '';
      if (estabelecimentoProduto !== loja) {
        return { sucesso: false, mensagem: 'Você não tem permissão para excluir produtos de outra loja.' };
      }
    }

    sheet.deleteRow(linha);
    return { sucesso: true, mensagem: 'Produto excluído com sucesso!' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
  }
}

// ========== EXCLUSÃO EM LOTE DOS ITENS MARCADOS COMO RESOLVIDOS ==========
function excluirProdutosResolvidos(perfil, loja) {
  try {
    if (perfil) {
      var check = verificarPermissao(perfil, 'Estoque', 'delete');
      if (!check.permitido) return check;
    }

    var sheet = getSheetByName('Estoque');
    var dados = sheet.getDataRange().getValues();
    var linhas = [];
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] && dados[i][14] === 'Sim') {
        // ===== Filtro por loja (se não for Master) =====
        if (perfil !== 'Master' && loja) {
          var estabelecimentoProduto = dados[i][9] || '';
          if (estabelecimentoProduto !== loja) continue;
        }
        linhas.push(i + 1);
      }
    }

    if (linhas.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum produto marcado como resolvido foi encontrado para sua loja.' };
    }

    for (var k = linhas.length - 1; k >= 0; k--) {
      sheet.deleteRow(linhas[k]);
    }
    return { sucesso: true, mensagem: linhas.length + ' produto(s) resolvido(s) excluído(s) com sucesso!', total: linhas.length };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
  }
}