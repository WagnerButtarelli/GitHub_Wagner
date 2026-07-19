/**
 * MEU COMÉRCIO
 * App de controle de vendas e estoque para pequenos comércios.
 * Backend Google Apps Script.
 * Versão com integração ao leitor externo de código de barras.
 */

// ======================= CONFIGURAÇÃO =======================

var CONFIG = {
  SPREADSHEET_ID: '1U22la4Ro-HA07idwR078fN45IkcYzeNlQCExCT20ukw',
  ABA_ESTOQUE: 'Estoque',
  ABA_VENDAS: 'Vendas',
  ABA_ACESSOS: 'acessos',
  ABA_CONFIG: 'Configuracoes',
  DIAS_VENCIMENTO_CRITICO: 7,
  QUANTIDADE_MINIMA_CRITICA: 3
};

var CABECALHO_ESTOQUE = [
  'ID', 'Código de Barras', 'Nome do Produto', 'Data de Cadastro',
  'Quantidade', 'Data de Validade', 'Valor de Compra', 'Valor de Venda',
  'Categoria', 'Subcategoria'
];

var CABECALHO_VENDAS = [
  'ID Venda', 'Data da Venda', 'Código de Barras', 'Nome do Produto',
  'Quantidade', 'Valor Unitário', 'Custo Unitário', 'Valor Total',
  'Forma de Pagamento', 'Valor Recebido', 'Troco'
];

var CABECALHO_ACESSOS = [
  'ID', 'Nome', 'Usuário', 'Senha', 'Status'
];

var CABECALHO_CONFIG = [
  'Chave', 'Valor'
];

// ======================= ENTRADA DO APP =======================

function doGet(e) {
  var parametros = (e && e.parameter) || {};

  if (parametros.acao === 'salvarCodigoEscaneado') {
    return salvarCodigoEscaneado_(parametros.sessionId, parametros.codigo, parametros.contexto);
  }

  var template = HtmlService.createTemplateFromFile('Index');
  template.execUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('Meu Comércio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function salvarCodigoEscaneado_(sessionId, codigo, contexto) {
  if (sessionId && codigo) {
    var cache = CacheService.getScriptCache();
    cache.put('scan_' + sessionId, JSON.stringify({
      codigo: String(codigo).trim(),
      contexto: String(contexto || 'cadastro').trim()
    }), 300);
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function buscarCodigoEscaneado(sessionId) {
  if (!sessionId) return null;
  var cache = CacheService.getScriptCache();
  var valor = cache.get('scan_' + sessionId);
  if (!valor) return null;
  cache.remove('scan_' + sessionId);
  return JSON.parse(valor);
}

// ======================= INFRAESTRUTURA DE PLANILHA =======================

function abrirPlanilha_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function obterAba_(nome, cabecalho) {
  var ss = abrirPlanilha_();
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalho);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
  }
  return aba;
}

function inicializarSistema() {
  var abaEstoque = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  obterAba_(CONFIG.ABA_VENDAS, CABECALHO_VENDAS);
  obterAba_(CONFIG.ABA_CONFIG, CABECALHO_CONFIG);
  var abaAcessos = obterAba_(CONFIG.ABA_ACESSOS, CABECALHO_ACESSOS);

  var headersEstoque = abaEstoque.getRange(1, 1, 1, abaEstoque.getLastColumn()).getValues()[0];
  if (headersEstoque.length < 9 || headersEstoque[8] !== 'Categoria') {
    abaEstoque.getRange(1, 9).setValue('Categoria');
  }
  if (headersEstoque.length < 10 || headersEstoque[9] !== 'Subcategoria') {
    abaEstoque.getRange(1, 10).setValue('Subcategoria');
  }

  if (abaAcessos.getLastRow() === 1) {
    abaAcessos.appendRow([gerarId_(), 'Administrador', 'admin', 'admin', 'Ativo']);
  }

  return { ok: true, tituloApp: obterTituloApp_() };
}

function gerarId_() {
  return Utilities.getUuid().split('-')[0];
}

// ======================= LOGIN & CONFIGURAÇÕES =======================

function validarLogin(usuario, senha) {
  var aba = obterAba_(CONFIG.ABA_ACESSOS, CABECALHO_ACESSOS);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][2]) === String(usuario) && String(dados[i][3]) === String(senha)) {
      if (String(dados[i][4]) !== 'Ativo') {
        return { ok: false, erro: 'Usuário inativo.' };
      }
      return { ok: true, nome: dados[i][1], tituloApp: obterTituloApp_() };
    }
  }
  return { ok: false, erro: 'Usuário ou senha inválidos.' };
}

function obterTituloApp_() {
  var aba = obterAba_(CONFIG.ABA_CONFIG, CABECALHO_CONFIG);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === 'TITULO_APP') return dados[i][1] || 'Meu Comércio';
  }
  return 'Meu Comércio';
}

function salvarTituloApp(novoTitulo) {
  var aba = obterAba_(CONFIG.ABA_CONFIG, CABECALHO_CONFIG);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === 'TITULO_APP') {
      aba.getRange(i + 1, 2).setValue(novoTitulo);
      return { ok: true, tituloApp: novoTitulo };
    }
  }
  aba.appendRow(['TITULO_APP', novoTitulo]);
  return { ok: true, tituloApp: novoTitulo };
}

// ======================= CRUD ACESSOS (USUÁRIOS) =======================

function listarUsuarios() {
  var aba = obterAba_(CONFIG.ABA_ACESSOS, CABECALHO_ACESSOS);
  var valores = aba.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < valores.length; i++) {
    if (!valores[i][0]) continue;
    lista.push({
      id: valores[i][0],
      nome: valores[i][1],
      usuario: valores[i][2],
      senha: valores[i][3],
      status: valores[i][4]
    });
  }
  return lista;
}

function salvarUsuario(dados) {
  var aba = obterAba_(CONFIG.ABA_ACESSOS, CABECALHO_ACESSOS);
  if (dados.id) {
    var valores = aba.getDataRange().getValues();
    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][0]) === String(dados.id)) {
        aba.getRange(i + 1, 2).setValue(dados.nome);
        aba.getRange(i + 1, 3).setValue(dados.usuario);
        aba.getRange(i + 1, 4).setValue(dados.senha);
        aba.getRange(i + 1, 5).setValue(dados.status);
        return { ok: true };
      }
    }
  } else {
    aba.appendRow([gerarId_(), dados.nome, dados.usuario, dados.senha, dados.status]);
  }
  return { ok: true };
}

function excluirUsuario(id) {
  var aba = obterAba_(CONFIG.ABA_ACESSOS, CABECALHO_ACESSOS);
  var valores = aba.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      aba.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Usuário não encontrado.' };
}

// ======================= CADASTRO DE ESTOQUE =======================

function buscarProdutoPorCodigo(codigo) {
  codigo = String(codigo).trim();
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var dados = aba.getDataRange().getValues();

  for (var i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][1]).trim() === codigo) {
      return {
        encontrado: true,
        origem: 'estoque',
        nome: dados[i][2],
        valorCompra: dados[i][6],
        valorVenda: dados[i][7],
        categoria: dados[i][8] || '',
        subcategoria: dados[i][9] || '',
        quantidadeDisponivel: Number(dados[i][4]) || 0
      };
    }
  }

  var nomeExterno = buscarNomeOpenFoodFacts_(codigo);
  if (nomeExterno) {
    return { encontrado: true, origem: 'openfoodfacts', nome: nomeExterno, valorCompra: '', valorVenda: '', categoria: '', subcategoria: '' };
  }

  return { encontrado: false };
}

function buscarDadosCompletosPorCodigo(codigo) {
  return buscarProdutoPorCodigo(codigo);
}

// ===== FUNÇÃO CORRIGIDA: tenta br e depois world =====
function buscarNomeOpenFoodFacts_(codigo) {
  try {
    var dominios = [
      'https://br.openfoodfacts.org/api/v0/product/',
      'https://world.openfoodfacts.org/api/v0/product/'
    ];
    for (var d = 0; d < dominios.length; d++) {
      var url = dominios[d] + encodeURIComponent(codigo) + '.json';
      var resposta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resposta.getResponseCode() === 200) {
        var json = JSON.parse(resposta.getContentText());
        if (json.status === 1 && json.product) {
          return json.product.product_name || json.product.product_name_pt || null;
        }
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

function cadastrarProduto(dados) {
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var id = gerarId_();
  aba.appendRow([
    id,
    String(dados.codigoBarras).trim(),
    dados.nome,
    new Date(),
    Number(dados.quantidade),
    dados.dataValidade ? new Date(dados.dataValidade) : '',
    Number(dados.valorCompra) || 0,
    Number(dados.valorVenda) || 0,
    dados.categoria || '',
    dados.subcategoria || ''
  ]);
  return { ok: true, id: id };
}

function atualizarProduto(id, dados) {
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var valores = aba.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      var linha = i + 1;
      aba.getRange(linha, 2).setValue(String(dados.codigoBarras).trim());
      aba.getRange(linha, 3).setValue(dados.nome);
      aba.getRange(linha, 5).setValue(Number(dados.quantidade));
      aba.getRange(linha, 6).setValue(dados.dataValidade ? new Date(dados.dataValidade) : '');
      aba.getRange(linha, 7).setValue(Number(dados.valorCompra) || 0);
      aba.getRange(linha, 8).setValue(Number(dados.valorVenda) || 0);
      aba.getRange(linha, 9).setValue(dados.categoria || '');
      aba.getRange(linha, 10).setValue(dados.subcategoria || '');
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Lote não encontrado.' };
}

function excluirProduto(id) {
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var valores = aba.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      aba.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Lote não encontrado.' };
}

function listarEstoque() {
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var valores = aba.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < valores.length; i++) {
    var linha = valores[i];
    if (!linha[0]) continue;
    lista.push({
      id: linha[0],
      codigoBarras: linha[1],
      nome: linha[2],
      dataCadastro: formatarData_(linha[3]),
      quantidade: linha[4],
      dataValidade: formatarData_(linha[5]),
      valorCompra: linha[6],
      valorVenda: linha[7],
      categoria: linha[8] || '',
      subcategoria: linha[9] || ''
    });
  }
  lista.sort(function (a, b) { return a.nome.localeCompare(b.nome); });
  return lista;
}

// ======================= PAINEL DE VENDAS (PDV) =======================

function buscarProdutoParaVenda(codigo) {
  codigo = String(codigo).trim();
  var aba = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var valores = aba.getDataRange().getValues();
  var lotes = [];

  for (var i = 1; i < valores.length; i++) {
    var linha = valores[i];
    if (String(linha[1]).trim() === codigo && Number(linha[4]) > 0) {
      lotes.push({
        id: linha[0],
        validade: linha[5] ? new Date(linha[5]).getTime() : Infinity,
        quantidade: Number(linha[4]),
        valorCompra: Number(linha[6]) || 0,
        valorVenda: Number(linha[7]) || 0,
        nome: linha[2]
      });
    }
  }

  if (lotes.length > 0) {
    lotes.sort(function (a, b) { return a.validade - b.validade; });
    var quantidadeTotal = lotes.reduce(function (soma, l) { return soma + l.quantidade; }, 0);
    return {
      encontrado: true,
      origem: 'estoque',
      codigoBarras: codigo,
      nome: lotes[0].nome,
      valorVenda: lotes[0].valorVenda,
      quantidadeDisponivel: quantidadeTotal
    };
  }

  // Consulta externa se não houver estoque
  var nomeExterno = buscarNomeOpenFoodFacts_(codigo);
  if (nomeExterno) {
    return {
      encontrado: true,
      origem: 'openfoodfacts',
      codigoBarras: codigo,
      nome: nomeExterno,
      valorVenda: 0,
      quantidadeDisponivel: 0
    };
  }

  return { encontrado: false };
}

function registrarVenda(carrinho, formaPagamento, valorRecebido) {
  if (!carrinho || carrinho.length === 0) {
    return { ok: false, erro: 'Carrinho vazio.' };
  }

  var abaEstoque = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var abaVendas = obterAba_(CONFIG.ABA_VENDAS, CABECALHO_VENDAS);
  var valoresEstoque = abaEstoque.getDataRange().getValues();
  var agora = new Date();
  var valorTotalVenda = 0;
  var idVenda = gerarId_();

  for (var c = 0; c < carrinho.length; c++) {
    var item = carrinho[c];
    var codigo = String(item.codigoBarras).trim();
    var quantidadeNecessaria = Number(item.quantidade);
    var custoTotalItem = 0;
    var quantidadeRestante = quantidadeNecessaria;

    var indicesLotes = [];
    for (var i = 1; i < valoresEstoque.length; i++) {
      if (String(valoresEstoque[i][1]).trim() === codigo && Number(valoresEstoque[i][4]) > 0) {
        indicesLotes.push(i);
      }
    }
    indicesLotes.sort(function (a, b) {
      var va = valoresEstoque[a][5] ? new Date(valoresEstoque[a][5]).getTime() : Infinity;
      var vb = valoresEstoque[b][5] ? new Date(valoresEstoque[b][5]).getTime() : Infinity;
      return va - vb;
    });

    for (var k = 0; k < indicesLotes.length && quantidadeRestante > 0; k++) {
      var idx = indicesLotes[k];
      var disponivelLote = Number(valoresEstoque[idx][4]);
      var consumir = Math.min(disponivelLote, quantidadeRestante);
      custoTotalItem += consumir * (Number(valoresEstoque[idx][6]) || 0);
      valoresEstoque[idx][4] = disponivelLote - consumir;
      quantidadeRestante -= consumir;
    }

    if (quantidadeRestante > 0) {
      return { ok: false, erro: 'Estoque insuficiente para o produto: ' + (item.nome || codigo) };
    }

    var custoUnitarioMedio = custoTotalItem / quantidadeNecessaria;
    var valorTotalItem = quantidadeNecessaria * Number(item.valorVenda);
    valorTotalVenda += valorTotalItem;

    abaVendas.appendRow([
      idVenda, agora, codigo, item.nome, quantidadeNecessaria,
      Number(item.valorVenda), custoUnitarioMedio, valorTotalItem,
      formaPagamento, '', ''
    ]);
  }

  abaEstoque.getRange(1, 1, valoresEstoque.length, CABECALHO_ESTOQUE.length).setValues(valoresEstoque);

  var troco = 0;
  if (formaPagamento === 'Dinheiro' && valorRecebido) {
    troco = Number(valorRecebido) - valorTotalVenda;
  }

  var ultimaLinha = abaVendas.getLastRow();
  abaVendas.getRange(ultimaLinha, 10).setValue(valorRecebido || '');
  abaVendas.getRange(ultimaLinha, 11).setValue(troco || 0);

  return {
    ok: true,
    idVenda: idVenda,
    valorTotal: valorTotalVenda,
    troco: troco
  };
}

// ======================= DASHBOARD =======================

function formatarData_(valor) {
  if (!valor) return '';
  return Utilities.formatDate(new Date(valor), Session.getScriptTimeZone() || 'GMT-3', 'dd/MM/yyyy');
}

function chaveDia_(data) {
  return Utilities.formatDate(new Date(data), Session.getScriptTimeZone() || 'GMT-3', 'yyyy-MM-dd');
}

function obterDadosDashboard(periodo) {
  var abaVendas = obterAba_(CONFIG.ABA_VENDAS, CABECALHO_VENDAS);
  var abaEstoque = obterAba_(CONFIG.ABA_ESTOQUE, CABECALHO_ESTOQUE);
  var vendas = abaVendas.getDataRange().getValues();
  var estoque = abaEstoque.getDataRange().getValues();

  var limites = calcularLimitesPeriodo_(periodo);

  var totalVendidoPeriodo = 0;
  var lucroPeriodo = 0;
  var porDia = {};
  var porDiaMes = {};

  var hoje = new Date();
  var inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var totalVendidoMes = 0;
  var lucroMes = 0;

  var seteDiasAtras = new Date();
  seteDiasAtras.setDate(seteDiasAtras.getDate() - 6);
  seteDiasAtras.setHours(0, 0, 0, 0);

  var vendasPorProduto = {};

  for (var i = 1; i < vendas.length; i++) {
    var linha = vendas[i];
    if (!linha[1]) continue;
    var dataVenda = new Date(linha[1]);
    var codigo = String(linha[2]).trim();
    var nome = linha[3];
    var quantidade = Number(linha[4]) || 0;
    var valorTotal = Number(linha[7]) || 0;
    var lucro = valorTotal - (Number(linha[4]) * Number(linha[6]) || 0);

    if (dataVenda >= limites.inicio && dataVenda <= limites.fim) {
      totalVendidoPeriodo += valorTotal;
      lucroPeriodo += lucro;
    }

    if (dataVenda >= inicioMes) {
      totalVendidoMes += valorTotal;
      lucroMes += lucro;

      var chaveMes = chaveDia_(dataVenda);
      if (!porDiaMes[chaveMes]) porDiaMes[chaveMes] = { vendas: 0, lucro: 0 };
      porDiaMes[chaveMes].vendas += valorTotal;
      porDiaMes[chaveMes].lucro += lucro;
    }

    if (dataVenda >= seteDiasAtras) {
      var chave = chaveDia_(dataVenda);
      if (!porDia[chave]) porDia[chave] = { vendas: 0, lucro: 0 };
      porDia[chave].vendas += valorTotal;
      porDia[chave].lucro += lucro;
    }

    if (codigo) {
      if (!vendasPorProduto[codigo]) {
        vendasPorProduto[codigo] = { nome: nome, quantidade: 0, totalVendido: 0 };
      }
      vendasPorProduto[codigo].quantidade += quantidade;
      vendasPorProduto[codigo].totalVendido += valorTotal;
    }
  }

  var serieDias = [];
  for (var d = 6; d >= 0; d--) {
    var dia = new Date();
    dia.setDate(dia.getDate() - d);
    var chaveDia = chaveDia_(dia);
    var registro = porDia[chaveDia] || { vendas: 0, lucro: 0 };
    serieDias.push({
      dia: Utilities.formatDate(dia, Session.getScriptTimeZone() || 'GMT-3', 'dd/MM'),
      vendas: registro.vendas,
      lucro: registro.lucro
    });
  }

  var serieMes = [];
  var ultimoDiaMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate();
  for (var dm = 1; dm <= ultimoDiaMes; dm++) {
    var diaMes = new Date(hoje.getFullYear(), hoje.getMonth(), dm);
    var chaveDm = chaveDia_(diaMes);
    var registroDm = porDiaMes[chaveDm] || { vendas: 0, lucro: 0 };
    serieMes.push({
      dia: Utilities.formatDate(diaMes, Session.getScriptTimeZone() || 'GMT-3', 'dd/MM'),
      vendas: registroDm.vendas,
      lucro: registroDm.lucro
    });
  }

  var totalSeguro = 0, totalCritico = 0;
  var itensCriticos = [];
  var limiteVencimento = new Date();
  limiteVencimento.setDate(limiteVencimento.getDate() + CONFIG.DIAS_VENCIMENTO_CRITICO);
  var previsaoPerdaTotal = 0;

  for (var j = 1; j < estoque.length; j++) {
    var linhaE = estoque[j];
    if (!linhaE[0]) continue;
    var quantidade = Number(linhaE[4]) || 0;
    if (quantidade <= 0) continue;

    var validade = linhaE[5] ? new Date(linhaE[5]) : null;
    var diasParaVencer = validade ? Math.ceil((validade - hoje) / 86400000) : null;
    var critico = (quantidade <= CONFIG.QUANTIDADE_MINIMA_CRITICA) ||
                  (validade !== null && validade <= limiteVencimento);

    if (critico) {
      totalCritico++;
      var valorInvestido = quantidade * (Number(linhaE[6]) || 0);
      var previsaoPerda = (validade !== null && diasParaVencer <= CONFIG.DIAS_VENCIMENTO_CRITICO)
                          ? valorInvestido : 0;
      previsaoPerdaTotal += previsaoPerda;
      itensCriticos.push({
        codigoBarras: linhaE[1],
        nome: linhaE[2],
        quantidade: quantidade,
        valorInvestido: valorInvestido,
        diasParaVencer: diasParaVencer,
        previsaoPerda: previsaoPerda
      });
    } else {
      totalSeguro++;
    }
  }

  itensCriticos.sort(function (a, b) {
    var da = a.diasParaVencer === null ? Infinity : a.diasParaVencer;
    var db = b.diasParaVencer === null ? Infinity : b.diasParaVencer;
    return da - db;
  });

  var mapaProdutos = {};
  for (var k = 1; k < estoque.length; k++) {
    var linhaE2 = estoque[k];
    if (!linhaE2[0]) continue;
    var cod = String(linhaE2[1]).trim();
    if (!cod) continue;
    mapaProdutos[cod] = {
      codigoBarras: cod,
      nome: linhaE2[2],
      quantidadeVendida: 0,
      totalVendido: 0
    };
  }

  for (var codigo in vendasPorProduto) {
    if (mapaProdutos[codigo]) {
      mapaProdutos[codigo].quantidadeVendida = vendasPorProduto[codigo].quantidade;
      mapaProdutos[codigo].totalVendido = vendasPorProduto[codigo].totalVendido;
    } else {
      mapaProdutos[codigo] = {
        codigoBarras: codigo,
        nome: vendasPorProduto[codigo].nome || codigo,
        quantidadeVendida: vendasPorProduto[codigo].quantidade,
        totalVendido: vendasPorProduto[codigo].totalVendido
      };
    }
  }

  var arrayProdutos = Object.values(mapaProdutos);
  arrayProdutos.sort(function (a, b) {
    return b.quantidadeVendida - a.quantidadeVendida;
  });

  var topProdutos = arrayProdutos.slice(0, 5);
  var todosProdutosOrdenados = arrayProdutos.slice();
  todosProdutosOrdenados.sort(function (a, b) {
    return a.quantidadeVendida - b.quantidadeVendida;
  });
  var bottomProdutos = todosProdutosOrdenados.slice(0, 5);

  return {
    periodo: {
      totalVendido: totalVendidoPeriodo,
      lucro: lucroPeriodo
    },
    mensal: {
      totalVendido: totalVendidoMes,
      lucro: lucroMes
    },
    grafico7dias: serieDias,
    graficoMes: serieMes,
    estoque: {
      totalSeguro: totalSeguro,
      totalCritico: totalCritico,
      itensCriticos: itensCriticos,
      previsaoPerdaTotal: previsaoPerdaTotal
    },
    topProdutos: topProdutos,
    bottomProdutos: bottomProdutos
  };
}

function calcularLimitesPeriodo_(periodo) {
  var hoje = new Date();
  var inicio, fim;

  if (periodo && typeof periodo === 'object' && periodo.inicio && periodo.fim) {
    inicio = new Date(periodo.inicio + 'T00:00:00');
    fim = new Date(periodo.fim + 'T23:59:59');
  } else if (periodo === '7dias') {
    inicio = new Date();
    inicio.setDate(inicio.getDate() - 6);
    inicio.setHours(0, 0, 0, 0);
    fim = new Date();
    fim.setHours(23, 59, 59, 999);
  } else if (periodo === 'mes') {
    inicio = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    fim = new Date();
    fim.setHours(23, 59, 59, 999);
  } else {
    inicio = new Date();
    inicio.setHours(0, 0, 0, 0);
    fim = new Date();
    fim.setHours(23, 59, 59, 999);
  }

  return { inicio: inicio, fim: fim };
}