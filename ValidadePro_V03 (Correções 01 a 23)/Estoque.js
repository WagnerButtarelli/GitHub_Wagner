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
          // ===== ITEM 20: expõe os dados de auditoria para o frontend =====
          cadastradoPor: dados[i][11] || '',
          funcao: dados[i][12] || '',
          dataCadastro: dataCadastro instanceof Date ? Utilities.formatDate(dataCadastro, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          // ===== FIM ITEM 20 =====
          resolvido: dados[i][14] === 'Sim' ? 'Sim' : 'Não',
          // ===== ITEM 22: data em que Quebra foi marcada como 'Sim' =====
          dataRebaixa: dataRebaixa instanceof Date ? Utilities.formatDate(dataRebaixa, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          // ===== FIM ITEM 22 =====
          diasRestantes: diasRestantes,
          status: status,
          linha: i + 1
        });
      }
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

function salvarProduto(produto, perfil) {
  try {
    // ===== ITEM 21: normaliza a linha para número logo no início =====
    // produto.linha chega do frontend como string (valor de input hidden).
    // Sem essa conversão, a comparação "i + 1 === produto.linha" nunca é
    // verdadeira (número !== string) e a própria linha nunca é ignorada
    // na checagem de duplicidade durante a edição.
    var linhaAtual = produto.linha ? parseInt(produto.linha, 10) : null;
    // ===== FIM ITEM 21 =====

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

    var sheet = getSheetByName('Estoque');
    var dados = sheet.getDataRange().getValues();

    // --- Verificação de duplicidade (código + data de validade) ---
    var dataValidade = produto.dataValidade ? parseDataLocal(produto.dataValidade) : '';
    var dataValidadeStr = dataValidade ? Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    var dataValidadeDisplay = dataValidade ? Utilities.formatDate(dataValidade, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'sem data';

    for (var i = 1; i < dados.length; i++) {
      // ===== ITEM 21: agora compara número com número, ignorando a própria linha na edição =====
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

      // ===== ITEM 22: atualiza a Data Rebaixa conforme a transição do campo Quebra =====
      // - Se estava 'Não' e virou 'Sim': grava a data/hora atual (marca a baixa agora).
      // - Se já estava 'Sim' e continua 'Sim': preserva a data original da baixa.
      // - Se virou 'Não': limpa a data (produto não está mais marcado como quebra).
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
      // ===== FIM ITEM 22 =====
    } else {
      // ===== ITEM 20: preencher colunas de auditoria =====
      var cadastradoPor = String(produto.cadastradoPor || '').trim();
      var funcaoUsuario = String(produto.funcao || '').trim();
      // ===== ITEM 22: se o produto já nasce marcado como Rebaixa=Sim, grava a data agora =====
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
        // colunas 12, 13, 14, 15 e 16
        cadastradoPor,                 // L – Cadastrado Por
        funcaoUsuario,                  // M – Função
        new Date(),                     // N – Data Cadastro
        resolvidoValor,                  // O – Resolvido
        dataRebaixaNovo                  // P – Data Rebaixa
      ]);
      // ===== FIM ITEM 20 / ITEM 22 =====
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

// ========== EXCLUSÃO EM LOTE DOS ITENS MARCADOS COMO RESOLVIDOS ==========
function excluirProdutosResolvidos(perfil) {
  try {
    if (perfil) {
      var check = verificarPermissao(perfil, 'Estoque', 'delete');
      if (!check.permitido) return check;
    }
    var sheet = getSheetByName('Estoque');
    var dados = sheet.getDataRange().getValues();
    var linhas = [];
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] && dados[i][14] === 'Sim') linhas.push(i + 1);
    }
    if (linhas.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum produto marcado como resolvido foi encontrado.' };
    }
    for (var k = linhas.length - 1; k >= 0; k--) {
      sheet.deleteRow(linhas[k]);
    }
    return { sucesso: true, mensagem: linhas.length + ' produto(s) resolvido(s) excluído(s) com sucesso!', total: linhas.length };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
  }
}