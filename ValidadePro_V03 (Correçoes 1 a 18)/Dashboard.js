// ==================== DASHBOARD ====================
function getDashboardData() {
try {
var result = getProdutos();
if (!result.sucesso) return result;
var produtos = result.produtos;
var hoje = new Date();
var total = produtos.length;
var vencidos = 0, supercritico = 0, critico = 0, alerta = 0, seguro = 0;
var valorPerda = 0, valorPerdaProximo = 0;
var categorias = {}, estabelecimentos = {}, fornecedores = {}, vencimentosPorMes = {};
produtos.forEach(function(p) {
if (p.status === 'vencido') { vencidos++; valorPerda += (parseFloat(p.precoCusto) || 0) * (parseInt(p.quantidade) || 0); }
else if (p.status === 'supercritico') { supercritico++; valorPerdaProximo += (parseFloat(p.precoCusto) || 0) * (parseInt(p.quantidade) || 0); }
else if (p.status === 'critico') { critico++; valorPerdaProximo += (parseFloat(p.precoCusto) || 0) * (parseInt(p.quantidade) || 0); }
else if (p.status === 'alerta') { alerta++; valorPerdaProximo += (parseFloat(p.precoCusto) || 0) * (parseInt(p.quantidade) || 0); }
else { seguro++; }
if (p.categoria) categorias[p.categoria] = (categorias[p.categoria] || 0) + 1;
if (p.estabelecimento) estabelecimentos[p.estabelecimento] = (estabelecimentos[p.estabelecimento] || 0) + 1;
if (p.fornecedor) fornecedores[p.fornecedor] = (fornecedores[p.fornecedor] || 0) + 1;
if (p.dataValidade) {
var data = new Date(p.dataValidade);
var mesAno = Utilities.formatDate(data, Session.getScriptTimeZone(), 'MM/yyyy');
vencimentosPorMes[mesAno] = (vencimentosPorMes[mesAno] || 0) + 1;
}
});
return {
sucesso: true,
resumo: { total: total, vencidos: vencidos, supercritico: supercritico, critico: critico, alerta: alerta, seguro: seguro, valorPerda: valorPerda.toFixed(2), valorPerdaProximo: valorPerdaProximo.toFixed(2) },
categorias: categorias,
estabelecimentos: estabelecimentos,
fornecedores: fornecedores,
vencimentosPorMes: vencimentosPorMes,
produtos: produtos
};
} catch (e) {
return { sucesso: false, mensagem: 'Erro: ' + e.toString() };
}
}
