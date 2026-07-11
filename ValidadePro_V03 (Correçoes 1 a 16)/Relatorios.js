// ==================== RELATÓRIOS ====================
function gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
try {
var result = getProdutos();
if (!result.sucesso) return result;
var produtos = result.produtos;
var filtrados = [];
produtos.forEach(function(p) {
if (tipo === 'vencidos' && p.status !== 'vencido') return;
if (tipo === 'proximos' && (p.diasRestantes === null || p.diasRestantes < 0 || p.diasRestantes > dias)) return;
if (categoria && p.categoria !== categoria) return;
if (subcategoria && p.subcategoria !== subcategoria) return;
if (estabelecimento && p.estabelecimento !== estabelecimento) return;
if (fornecedor && p.fornecedor !== fornecedor) return;
filtrados.push(p);
});
return { sucesso: true, produtos: filtrados };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== EXPORTAÇÃO CSV ====================
function exportarRelatorioCSV(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
try {
var result = gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor);
if (!result.sucesso) return result;
var produtos = result.produtos;
var csv = 'Codigo;Produto;Quantidade;Data Validade;Categoria;Subcategoria;Preco Custo;Preco Venda;Fornecedor;Estabelecimento;Rebaixa;Status;Dias Restantes\n';
produtos.forEach(function(p) {
csv += [
'"' + (p.codigo || '') + '"',
'"' + (p.produto || '') + '"',
p.quantidade,
p.dataValidade,
'"' + (p.categoria || '') + '"',
'"' + (p.subcategoria || '') + '"',
(p.precoCusto || 0).toString().replace('.', ','),
(p.precoVenda || 0).toString().replace('.', ','),
'"' + (p.fornecedor || '') + '"',
'"' + (p.estabelecimento || '') + '"',
p.Rebaixa || 'Não',
p.status,
p.diasRestantes !== null ? p.diasRestantes : ''
].join(';') + '\n';
});
return { sucesso: true, csv: csv, total: produtos.length };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
// ==================== EXPORTAÇÃO PDF ====================
function exportarRelatorioPDF(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor) {
try {
var result = gerarRelatorio(tipo, dias, categoria, subcategoria, estabelecimento, fornecedor);
if (!result.sucesso) return result;
var produtos = result.produtos;
var titulo = tipo === 'vencidos' ? 'Produtos Vencidos' : tipo === 'proximos' ? 'Produtos Proximos do Vencimento (' + dias + ' dias)' : 'Todos os Produtos';
var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>';
html += 'body{font-family:Arial,sans-serif;margin:40px;color:#333;}';
html += 'h1{color:#1a1a2e;font-size:24px;margin-bottom:5px;}';
html += 'h2{color:#667eea;font-size:14px;margin-top:0;font-weight:normal;}';
html += 'table{width:100%;border-collapse:collapse;margin-top:20px;font-size:11px;}';
html += 'th{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:10px;text-align:left;font-size:10px;text-transform:uppercase;}';
html += 'td{padding:8px 10px;border-bottom:1px solid #eee;}';
html += 'tr:nth-child(even){background:#f8f9fa;}';
html += '.status-vencido{color:#ef4444;font-weight:bold;}';
html += '.status-supercritico{color:#f97316;font-weight:bold;}';
html += '.status-critico{color:#f59e0b;font-weight:bold;}';
html += '.status-alerta{color:#fbbf24;font-weight:bold;}';
html += '.status-seguro{color:#10b981;font-weight:bold;}';
html += '.footer{margin-top:30px;font-size:10px;color:#999;text-align:center;border-top:1px solid #eee;padding-top:15px;}';
html += '.badge{display:inline-block;padding:3px 8px;border-radius:12px;font-size:9px;font-weight:bold;text-transform:uppercase;}';
html += '.badge-vencido{background:#fee2e2;color:#991b1b;}';
html += '.badge-supercritico{background:#ffedd5;color:#9a3412;}';
html += '.badge-critico{background:#fef3c7;color:#92400e;}';
html += '.badge-alerta{background:#fef9c3;color:#854d0e;}';
html += '.badge-seguro{background:#d1fae5;color:#065f46;}';
html += '</style></head><body>';
html += '<h1>ValidadePro - Relatorio Preventivo</h1>';
html += '<h2>' + titulo + ' | Gerado em: ' + hoje + ' | Total: ' + produtos.length + ' produto(s)</h2>';
html += '<table><thead><tr>';
html += '<th>Codigo</th><th>Produto</th><th>Qtd</th><th>Validade</th><th>Dias</th><th>Status</th><th>Quebra</th>';
html += '<th>Categoria</th><th>Subcategoria</th><th>Custo</th><th>Venda</th><th>Fornecedor</th><th>Local</th>';
html += '</tr></thead><tbody>';
produtos.forEach(function(p) {
var dataStr = p.dataValidade ? Utilities.formatDate(new Date(p.dataValidade), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '-';
var diasStr = p.diasRestantes !== null ? p.diasRestantes + ' dias' : '-';
var statusClass = 'badge-' + p.status;
var statusLabel = p.status === 'vencido' ? 'Vencido' : p.status === 'supercritico' ? 'Supercritico' : p.status === 'critico' ? 'Critico' : p.status === 'alerta' ? 'Alerta' : 'Seguro';
var custo = 'R$ ' + (parseFloat(p.precoCusto) || 0).toFixed(2).replace('.', ',');
var venda = 'R$ ' + (parseFloat(p.precoVenda) || 0).toFixed(2).replace('.', ',');
html += '<tr>';
html += '<td><strong>' + (p.codigo || '-') + '</strong></td>';
html += '<td>' + (p.produto || '-') + '</td>';
html += '<td>' + p.quantidade + '</td>';
html += '<td>' + dataStr + '</td>';
html += '<td>' + diasStr + '</td>';
html += '<td><span class="badge ' + statusClass + '">' + statusLabel + '</span></td>';
html += '<td>' + (p.quebra || 'Não') + '</td>';
html += '<td>' + (p.categoria || '-') + '</td>';
html += '<td>' + (p.subcategoria || '-') + '</td>';
html += '<td>' + custo + '</td>';
html += '<td>' + venda + '</td>';
html += '<td>' + (p.fornecedor || '-') + '</td>';
html += '<td>' + (p.estabelecimento || '-') + '</td>';
html += '</tr>';
});
html += '</tbody></table>';
html += '<div class="footer">ValidadePro - Gestao Inteligente de Estoque | Relatorio gerado automaticamente</div>';
html += '</body></html>';
var blob = Utilities.newBlob(html, 'text/html', 'relatorio.html');
var pdf = blob.getAs('application/pdf');
var base64 = Utilities.base64Encode(pdf.getBytes());
return { sucesso: true, pdf: base64, total: produtos.length };
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
