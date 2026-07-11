// ==================== RELATÓRIOS ====================
function gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
  try {
    var result = getProdutos();
    if (!result.sucesso) return result;
    
    var filtrados = result.produtos.filter(function(p) {
      if (tipo === 'vencidos' && p.status !== 'vencido') return false;
      if (tipo === 'proximos' && (p.diasRestantes === null || p.diasRestantes < 0 || p.diasRestantes > dias)) return false;
      if (categoria && p.categoria !== categoria) return false;
      if (subcategoria && p.subcategoria !== subcategoria) return false;
      if (estabelecimento && p.estabelecimento !== estabelecimento) return false;
      if (fornecedor && p.fornecedor !== fornecedor) return false;
      return true;
    });
    
    return { sucesso: true, produtos: filtrados };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function exportarRelatorioCSV(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
  try {
    var result = gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor);
    if (!result.sucesso) return result;
    
    var csv = 'Codigo;Produto;Qtd;Validade;Status;Quebra;Categoria;Subcategoria;Custo;Venda;Fornecedor;Loja;Dias\n';
    result.produtos.forEach(function(p) {
      csv += [
        '"' + p.codigo + '"', '"' + p.produto + '"', p.quantidade, p.dataValidade,
        p.status, p.quebra, '"' + p.categoria + '"', '"' + p.subcategoria + '"',
        (p.precoCusto || 0).toString().replace('.', ','), (p.precoVenda || 0).toString().replace('.', ','),
        '"' + p.fornecedor + '"', '"' + p.estabelecimento + '"',
        p.diasRestantes !== null ? p.diasRestantes : ''
      ].join(';') + '\n';
    });
    
    return { sucesso: true, csv: csv, total: result.produtos.length };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}

function exportarRelatorioPDF(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
  try {
    var result = gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor);
    if (!result.sucesso) return result;
    
    var produtos = result.produtos;
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>';
    html += 'body{font-family:Arial,sans-serif;margin:40px;color:#333;}';
    html += 'table{width:100%;border-collapse:collapse;font-size:10px;}';
    html += 'th{background:#667eea;color:white;padding:8px;text-align:left;}';
    html += 'td{padding:6px;border-bottom:1px solid #eee;}';
    html += '.badge{padding:2px 6px;border-radius:10px;font-size:8px;font-weight:bold;}';
    html += '.badge-vencido{background:#fee2e2;color:#991b1b;}';
    html += '.badge-supercritico{background:#fecaca;color:#7f1d1d;}';
    html += '.badge-critico{background:#fef3c7;color:#92400e;}';
    html += '.badge-alerta{background:#fef9c3;color:#854d0e;}';
    html += '.badge-seguro{background:#d1fae5;color:#065f46;}';
    html += '</style></head><body>';
    html += '<h1>ValidadePro - Relatorio</h1>';
    html += '<p>Total: ' + produtos.length + ' produtos</p>';
    html += '<table><thead><tr><th>Codigo</th><th>Produto</th><th>Qtd</th><th>Validade</th><th>Dias</th><th>Status</th><th>Quebra</th><th>Categoria</th></tr></thead><tbody>';
    
    produtos.forEach(function(p) {
      var dataStr = p.dataValidade ? Utilities.formatDate(new Date(p.dataValidade), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '-';
      var diasStr = p.diasRestantes !== null ? p.diasRestantes : '-';
      var statusLabel = p.status === 'vencido' ? 'Vencido' : p.status === 'supercritico' ? 'Supercritico' : p.status === 'critico' ? 'Critico' : p.status === 'alerta' ? 'Alerta' : 'Seguro';
      
      html += '<tr>';
      html += '<td>' + p.codigo + '</td><td>' + p.produto + '</td><td>' + p.quantidade + '</td>';
      html += '<td>' + dataStr + '</td><td>' + diasStr + '</td>';
      html += '<td><span class="badge badge-' + p.status + '">' + statusLabel + '</span></td>';
      html += '<td>' + p.quebra + '</td><td>' + p.categoria + '</td>';
      html += '</tr>';
    });
    
    html += '</tbody></table></body></html>';
    var blob = Utilities.newBlob(html, 'text/html', 'relatorio.html');
    return { sucesso: true, pdf: Utilities.base64Encode(blob.getAs('application/pdf').getBytes()), total: produtos.length };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}