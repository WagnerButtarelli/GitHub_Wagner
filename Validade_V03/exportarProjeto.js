 function exportarProjeto() {
  // Pega o ID do projeto atual automaticamente
  var projectId = ScriptApp.getScriptId();
  
  // Chama a Apps Script API
  var url = "https://script.googleapis.com/v1/projects/" + projectId + "/content";
  
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  
  // Verifica erro
  if (response.getResponseCode() !== 200) {
    Logger.log("❌ ERRO: " + response.getContentText());
    Logger.log("Verifique se a Apps Script API está habilitada (Passo 1)");
    return;
  }
  
  var content = JSON.parse(response.getContentText());
  var files = content.files;
  
  // Monta o texto de saída
  var output = "EXPORTAÇÃO DO PROJETO\n";
  output += "Data: " + new Date().toLocaleString("pt-BR") + "\n";
  output += "Total de arquivos: " + files.length + "\n";
  output += "=".repeat(60) + "\n\n";
  
  files.forEach(function(file) {
    var extensao = file.type === "SERVER_JS" ? ".gs" : 
                   file.type === "HTML" ? ".html" : ".json";
    
    output += "▶ ARQUIVO: " + file.name + extensao + "\n";
    output += "  Tipo: " + file.type + "\n";
    output += "-".repeat(60) + "\n";
    output += file.source + "\n";
    output += "=".repeat(60) + "\n\n";
  });
  
  // Cria o arquivo no Drive
  var nomeArquivo = "Backup_GAS_" + new Date().toISOString().slice(0,10) + ".txt";
  var arquivo = DriveApp.createFile(nomeArquivo, output, MimeType.PLAIN_TEXT);
  
  Logger.log("✅ SUCESSO!");
  Logger.log("📁 Arquivo criado: " + arquivo.getName());
  Logger.log("🔗 Acesse em: " + arquivo.getUrl());
  
  return arquivo.getUrl();
}