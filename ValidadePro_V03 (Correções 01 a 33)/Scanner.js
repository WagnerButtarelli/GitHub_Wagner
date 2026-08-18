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

// ===== BUSCA DIRETA NA PLANILHA ESTOQUE =====
function buscarProdutoPorCodigo(codigoBarras) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilhaEstoque = ss.getSheetByName("Estoque");
    if (!planilhaEstoque) {
      Logger.log("Planilha 'Estoque' não encontrada.");
      return { encontrado: false, mensagem: "Planilha 'Estoque' não encontrada." };
    }
    var dadosEstoque = planilhaEstoque.getDataRange().getValues();
    var codigoLimpo = codigoBarras.trim();
    for (var i = 1; i < dadosEstoque.length; i++) {
      var codigoEstoque = String(dadosEstoque[i][0]).trim();
      if (codigoEstoque === codigoLimpo) {
        var produto = {
          codigo: dadosEstoque[i][0],
          nome: dadosEstoque[i][1] || "",
          quantidade: dadosEstoque[i][2] || "",
          dataValidade: dadosEstoque[i][3] || "",
          categoria: dadosEstoque[i][4] || "",
          subcategoria: dadosEstoque[i][5] || "",
          precoCusto: dadosEstoque[i][6] || "",
          precoVenda: dadosEstoque[i][7] || "",
          fornecedor: dadosEstoque[i][8] || "",
          estabelecimento: dadosEstoque[i][9] || "",
          quebra: dadosEstoque[i][10] === "Sim" ? "Sim" : "Não"
        };
        if (produto.dataValidade instanceof Date) {
          produto.dataValidade = Utilities.formatDate(produto.dataValidade, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          produto.dataValidade = String(produto.dataValidade || "");
        }
        Logger.log("Produto encontrado no estoque: " + produto.nome);
        return { encontrado: true, produto: produto };
      }
    }
    Logger.log("Código " + codigoLimpo + " não encontrado no estoque.");
    return { encontrado: false, mensagem: "Código não encontrado no estoque." };
  } catch (erro) {
    Logger.log("Erro ao buscar no estoque: " + erro.message);
    return { encontrado: false, mensagem: "Erro ao buscar no estoque: " + erro.message };
  }
}

function salvarDadosNoSheets(codigoBarras) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilha = obterOuCriarAbaPesquisas(ss);
    var codigoLimpo = codigoBarras.trim();
    var descricao = "";
    var quantidadeEstoque = "";
    var dataValidade = "";
    var mensagemEstoque = "";
    var categoriaEstoque = "";
    var subcategoriaEstoque = "";
    var encontradoNoEstoque = false;

    // ===== 1º PASSO: BUSCAR NA PLANILHA ESTOQUE =====
    var resultadoEstoque = buscarProdutoPorCodigo(codigoLimpo);
    if (resultadoEstoque.encontrado) {
      var prod = resultadoEstoque.produto;
      encontradoNoEstoque = true;
      descricao = prod.nome || "Produto sem nome cadastrado no estoque";
      quantidadeEstoque = prod.quantidade || "";
      dataValidade = prod.dataValidade || "";
      categoriaEstoque = prod.categoria || "";
      subcategoriaEstoque = prod.subcategoria || "";
      mensagemEstoque = "Produto encontrado no estoque!";
      Logger.log("Usando dados do estoque para o código " + codigoLimpo);
    } else {
      // ===== 2º PASSO: SE NÃO ENCONTRADO, CONSULTAR API EXTERNA =====
      mensagemEstoque = "Produto não encontrado no estoque. Buscando na base externa...";
      Logger.log("Iniciando consulta à API Open Food Facts para o código: " + codigoLimpo);
      try {
        // ===== URL CORRETA PARA OPEN FOOD FACTS =====
        var urlApi = "https://world.openfoodfacts.org/api/v0/product/" + encodeURIComponent(codigoLimpo) + ".json";
        var opcoes = {
          "muteHttpExceptions": true,
          "headers": {
            "User-Agent": "ScannerAppGoogleSheets - Vinicius - Versao2.0"
          }
        };
        var resposta = UrlFetchApp.fetch(urlApi, opcoes);
        var statusCode = resposta.getResponseCode();
        Logger.log("Resposta da API - Status HTTP: " + statusCode);
        if (statusCode === 200) {
          var dados = JSON.parse(resposta.getContentText());
          if (dados.status === 1 && dados.product) {
            descricao = dados.product.product_name_pt || dados.product.product_name || "Produto sem nome cadastrado na base externa";
            Logger.log("Descrição obtida da API: " + descricao);
          } else {
            descricao = "Não localizado na base externa (status da API: " + dados.status + ")";
            Logger.log("API retornou status diferente de 1: " + dados.status);
          }
        } else {
          descricao = "Não localizado (Status HTTP: " + statusCode + ")";
          Logger.log("Erro HTTP: " + statusCode);
        }
      } catch (erroApi) {
        descricao = "Erro na Busca: " + erroApi.message;
        Logger.log("Exceção ao chamar API: " + erroApi.message);
      }
      if (!descricao || descricao === "") {
        descricao = "Produto não identificado";
        Logger.log("Descrição vazia, usando valor padrão.");
      }
    }

    // ===== REGISTRA NO HISTÓRICO =====
    var dataHoraAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var descricaoHistorico = descricao;
    var qtdHistorico = quantidadeEstoque || "N/A";
    var valHistorico = dataValidade || "N/A";
    planilha.appendRow([dataHoraAtual, codigoLimpo, descricaoHistorico, qtdHistorico, valHistorico]);
    SpreadsheetApp.flush();
    Logger.log("Histórico atualizado com sucesso.");

    return {
      sucesso: true,
      codigo: codigoLimpo,
      descricao: descricao,
      quantidade: String(quantidadeEstoque || ""),
      dataValidade: String(dataValidade || ""),
      mensagemEstoque: mensagemEstoque,
      categoria: categoriaEstoque,
      subcategoria: subcategoriaEstoque,
      encontradoNoEstoque: encontradoNoEstoque
    };
  } catch (erro) {
    Logger.log("Erro geral em salvarDadosNoSheets: " + erro.message);
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