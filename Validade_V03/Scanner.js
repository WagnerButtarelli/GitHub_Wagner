// ==================== SCANNER ====================
function salvarDadosNoSheets(codigoBarras) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilha = getSheetByName('Pesquisas');
    var descricao = 'Não localizado';
    var codigoLimpo = codigoBarras.trim();
    var qtd = 'N/A';
    var dv = 'N/A';
    var msg = 'Não aplicável';
    
    if (codigoLimpo.indexOf('http') !== 0) {
      try {
        var resp = UrlFetchApp.fetch('https://world.openfoodfacts.org/api/v0/product/' + codigoLimpo + '.json', { 
          muteHttpExceptions: true 
        });
        if (resp.getResponseCode() === 200) {
          var dados = JSON.parse(resp.getContentText());
          if (dados.status === 1) {
            descricao = dados.product.product_name_pt || dados.product.product_name || 'Sem nome';
          }
        }
      } catch (e) {}
      
      var estSheet = ss.getSheetByName('Estoque');
      if (estSheet) {
        var dadosEst = estSheet.getDataRange().getValues();
        for (var i = 1; i < dadosEst.length; i++) {
          if (String(dadosEst[i][0]).trim() === codigoLimpo) {
            qtd = dadosEst[i][2];
            dv = dadosEst[i][3] instanceof Date ? Utilities.formatDate(dadosEst[i][3], Session.getScriptTimeZone(), 'dd/MM/yyyy') : dadosEst[i][3];
            msg = 'Encontrado no estoque!';
            break;
          }
        }
        if (msg !== 'Encontrado no estoque!') msg = 'Não encontrado no estoque';
      }
    }
    
    planilha.appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      codigoLimpo, descricao, qtd, dv
    ]);
    
    return { 
      sucesso: true, codigo: codigoLimpo, descricao: descricao, 
      quantidade: String(qtd), dataValidade: String(dv), mensagemEstoque: msg 
    };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function obterUltimasPesquisas() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pesquisas');
    if (!planilha || planilha.getLastRow() <= 1) return [];
    
    var numLinhas = Math.min(10, planilha.getLastRow() - 1);
    var dados = planilha.getRange(planilha.getLastRow() - numLinhas + 1, 1, numLinhas, 5).getValues();
    
    return dados.map(function(l) {
      return l.map(function(v) {
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        return v === null || v === undefined ? '' : String(v);
      });
    }).reverse();
  } catch (e) { 
    return []; 
  }
}