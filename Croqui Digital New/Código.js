// ============================================================
// CONFIGURAÇÕES GLOBAIS
// ============================================================
var PASTA_FOTOS = "1Qm_f7H7F00h4IvrJz-mGA8XdpP6sruZv"; // Substitua pelo ID da sua pasta
var NOME_PLANILHA_DADOS   = "Dados";
var NOME_PLANILHA_USUARIOS = "Usuarios";
var NOME_PLANILHA_CONFIG  = "Config";

// ============================================================
// DOGET - Servir a página HTML
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Croqui Digital')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// AUTENTICAÇÃO (com criação automática do admin)
// ============================================================
function autenticar(usuario, senha) {
  try {
    var sheet = getOrCreateSheet(NOME_PLANILHA_USUARIOS, ["Usuario", "Senha", "Nome"]);
    var dados = sheet.getDataRange().getValues();
    
    // Se não houver nenhum usuário, cria o administrador padrão
    if (dados.length <= 1) {
      sheet.appendRow(["E1", "1", "Administrador"]);
      dados = sheet.getDataRange().getValues();
    }
    
    // Procura usuário (case-insensitive, removendo espaços)
    usuario = usuario.trim();
    senha   = senha.trim();
    for (var i = 1; i < dados.length; i++) {
      var user = dados[i][0] ? dados[i][0].toString().trim() : "";
      var pass = dados[i][1] ? dados[i][1].toString().trim() : "";
      if (user.toLowerCase() === usuario.toLowerCase() && pass === senha) {
        return { success: true, nome: dados[i][2] || user };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false, erro: e.toString() };
  }
}

// CRUD de usuários (apenas E1)
function listarUsuarios() {
  var sheet = getOrCreateSheet(NOME_PLANILHA_USUARIOS, ["Usuario", "Senha", "Nome"]);
  var dados = sheet.getDataRange().getValues();
  var usuarios = [];
  for (var i = 1; i < dados.length; i++) {
    usuarios.push({
      usuario: dados[i][0] ? dados[i][0].toString().trim() : "",
      senha:   dados[i][1] ? dados[i][1].toString().trim() : "",
      nome:    dados[i][2] ? dados[i][2].toString().trim() : ""
    });
  }
  return usuarios;
}

function salvarUsuario(usuario, senha, nome) {
  var sheet = getOrCreateSheet(NOME_PLANILHA_USUARIOS, ["Usuario", "Senha", "Nome"]);
  // Remove se já existe (case-insensitive)
  var dados = sheet.getDataRange().getValues();
  for (var i = dados.length - 1; i >= 1; i--) {
    var user = dados[i][0] ? dados[i][0].toString().trim() : "";
    if (user.toLowerCase() === usuario.trim().toLowerCase()) {
      sheet.deleteRow(i+1);
    }
  }
  sheet.appendRow([usuario.trim(), senha.trim(), nome.trim()]);
  return "OK";
}

function excluirUsuario(usuario) {
  var sheet = getOrCreateSheet(NOME_PLANILHA_USUARIOS, ["Usuario", "Senha", "Nome"]);
  var dados = sheet.getDataRange().getValues();
  for (var i = dados.length - 1; i >= 1; i--) {
    var user = dados[i][0] ? dados[i][0].toString().trim() : "";
    if (user.toLowerCase() === usuario.trim().toLowerCase()) {
      sheet.deleteRow(i+1);
      return "OK";
    }
  }
  return "Não encontrado";
}

// ============================================================
// CONFIGURAÇÕES DOS CAMPOS
// ============================================================
function getConfigCampos() {
  var sheet = getOrCreateSheet(NOME_PLANILHA_CONFIG, ["Nome", "Habilitado"]);
  var dados = sheet.getDataRange().getValues();
  var config = [];
  for (var i = 1; i < dados.length; i++) {
    config.push({ nome: dados[i][0] ? dados[i][0].toString().trim() : "", habilitado: dados[i][1] === true });
  }
  return config;
}

function setConfigCampos(campos) {
  var sheet = getOrCreateSheet(NOME_PLANILHA_CONFIG, ["Nome", "Habilitado"]);
  sheet.clearContents();
  sheet.appendRow(["Nome", "Habilitado"]);
  for (var i = 0; i < campos.length; i++) {
    sheet.appendRow([campos[i].nome, campos[i].habilitado]);
  }
  return "OK";
}

// ============================================================
// SALVAR PONTO (online)
// ============================================================
function salvarPonto(dadosPonto) {
  try {
    var pasta = DriveApp.getFolderById(PASTA_FOTOS);
    var sheet = getOrCreateSheet(NOME_PLANILHA_DADOS,
      ["DataHora", "Usuario", "Projeto", "Latitude", "Longitude", "URL_Foto", "Campos"]);

    var projeto = dadosPonto.projeto || "SemProjeto";
    var contador = getProximoContador(projeto);
    var nomeFoto = projeto + "_" + contador + ".jpg";

    // Decodifica Base64
    var base64 = dadosPonto.fotoBase64;
    var parts = base64.split(',');
    var contentType = parts[0].match(/image\/(.*);/)[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), 'image/' + contentType, nomeFoto);
    var arquivo = pasta.createFile(blob);
    var url = arquivo.getUrl();

    var camposJSON = JSON.stringify(dadosPonto.campos || {});

    sheet.appendRow([
      new Date(),
      dadosPonto.usuario,
      dadosPonto.projeto,
      dadosPonto.latitude,
      dadosPonto.longitude,
      url,
      camposJSON
    ]);

    return { success: true, contador: contador };
  } catch (e) {
    return { success: false, erro: e.toString() };
  }
}

function getProximoContador(projeto) {
  var prop = PropertiesService.getScriptProperties();
  var chave = "contador_" + projeto;
  var atual = parseInt(prop.getProperty(chave) || "0");
  var novo = atual + 1;
  prop.setProperty(chave, novo.toString());
  return novo;
}

// ============================================================
// SINCRONIZAR MÚLTIPLOS PONTOS
// ============================================================
function sincronizarPontos(pontos) {
  var resultados = [];
  for (var i = 0; i < pontos.length; i++) {
    var resp = salvarPonto(pontos[i]);
    resultados.push(resp);
  }
  return resultados;
}

// ============================================================
// EXPORTAÇÕES
// ============================================================
function exportarDXF() {
  var sheet = getOrCreateSheet(NOME_PLANILHA_DADOS,
    ["DataHora", "Usuario", "Projeto", "Latitude", "Longitude", "URL_Foto", "Campos"]);
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return null;

  var dxf = "0\nSECTION\n2\nENTITIES\n";
  for (var i = 1; i < dados.length; i++) {
    var lat = dados[i][3];
    var lon = dados[i][4];
    if (lat && lon) {
      dxf += "0\nPOINT\n8\nColetas\n10\n" + lon + "\n20\n" + lat + "\n";
    }
  }
  dxf += "0\nENDSEC\n0\nEOF\n";
  return Utilities.base64Encode(dxf);
}

function exportarCSV(projetoFiltro) {
  var sheet = getOrCreateSheet(NOME_PLANILHA_DADOS,
    ["DataHora", "Usuario", "Projeto", "Latitude", "Longitude", "URL_Foto", "Campos"]);
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return null;

  var linhas = [];
  var cabecalho = ["DataHora","Usuario","Projeto","Latitude","Longitude","URL_Foto"];
  var camposExtras = [];

  for (var i = 1; i < dados.length; i++) {
    if (projetoFiltro && dados[i][2] !== projetoFiltro) continue;
    var camposJSON = dados[i][6];
    if (camposJSON) {
      try {
        var obj = JSON.parse(camposJSON);
        for (var chave in obj) {
          if (camposExtras.indexOf(chave) === -1) camposExtras.push(chave);
        }
      } catch(e) {}
    }
  }
  cabecalho = cabecalho.concat(camposExtras);
  linhas.push(cabecalho.join(','));

  for (var i = 1; i < dados.length; i++) {
    if (projetoFiltro && dados[i][2] !== projetoFiltro) continue;
    var linha = [
      dados[i][0], dados[i][1], dados[i][2], dados[i][3], dados[i][4], dados[i][5]
    ];
    var obj = {};
    var camposJSON = dados[i][6];
    if (camposJSON) {
      try { obj = JSON.parse(camposJSON); } catch(e) {}
    }
    for (var j = 0; j < camposExtras.length; j++) {
      linha.push(obj[camposExtras[j]] || '');
    }
    linhas.push(linha.join(','));
  }
  return Utilities.base64Encode(linhas.join('\n'));
}

function exportarExcel(projetoFiltro) {
  var sheet = getOrCreateSheet(NOME_PLANILHA_DADOS,
    ["DataHora", "Usuario", "Projeto", "Latitude", "Longitude", "URL_Foto", "Campos"]);
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return null;

  var tempSS = SpreadsheetApp.create("TempExport");
  var tempSheet = tempSS.getActiveSheet();

  var cabecalho = ["DataHora","Usuario","Projeto","Latitude","Longitude","URL_Foto"];
  var camposExtras = [];
  for (var i = 1; i < dados.length; i++) {
    if (projetoFiltro && dados[i][2] !== projetoFiltro) continue;
    var camposJSON = dados[i][6];
    if (camposJSON) {
      try {
        var obj = JSON.parse(camposJSON);
        for (var chave in obj) {
          if (camposExtras.indexOf(chave) === -1) camposExtras.push(chave);
        }
      } catch(e) {}
    }
  }
  cabecalho = cabecalho.concat(camposExtras);
  tempSheet.appendRow(cabecalho);

  for (var i = 1; i < dados.length; i++) {
    if (projetoFiltro && dados[i][2] !== projetoFiltro) continue;
    var linha = [
      dados[i][0], dados[i][1], dados[i][2], dados[i][3], dados[i][4], dados[i][5]
    ];
    var obj = {};
    var camposJSON = dados[i][6];
    if (camposJSON) {
      try { obj = JSON.parse(camposJSON); } catch(e) {}
    }
    for (var j = 0; j < camposExtras.length; j++) {
      linha.push(obj[camposExtras[j]] || '');
    }
    tempSheet.appendRow(linha);
  }

  var file = DriveApp.getFileById(tempSS.getId());
  var blob = file.getBlob();
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  return Utilities.base64Encode(blob.getBytes());
}

// ============================================================
// UTILITÁRIO
// ============================================================
function getOrCreateSheet(nome, cabecalho) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    if (cabecalho) sheet.appendRow(cabecalho);
  }
  return sheet;
}