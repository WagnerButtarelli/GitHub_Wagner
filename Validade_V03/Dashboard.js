// ==================== DASHBOARD ====================
function getDashboardData() {
  try {
    var result = getProdutos();
    if (!result.sucesso) return result;
    
    var produtos = result.produtos;
    var total = produtos.length;
    var vencidos = 0, supercritico = 0, critico = 0, alerta = 0, seguro = 0;
    var valorPerda = 0, valorPerdaProximo = 0;
    var categorias = {}, estabelecimentos = {}, vencimentosPorMes = {};
    
    produtos.forEach(function(p) {
      var custoQtd = (parseFloat(p.precoCusto) || 0) * (parseInt(p.quantidade) || 0);
      
      if (p.status === 'vencido') { vencidos++; valorPerda += custoQtd; }
      else if (p.status === 'supercritico') { supercritico++; valorPerdaProximo += custoQtd; }
      else if (p.status === 'critico') { critico++; valorPerdaProximo += custoQtd; }
      else if (p.status === 'alerta') { alerta++; valorPerdaProximo += custoQtd; }
      else seguro++;
      
      if (p.categoria) categorias[p.categoria] = (categorias[p.categoria] || 0) + 1;
      if (p.estabelecimento) estabelecimentos[p.estabelecimento] = (estabelecimentos[p.estabelecimento] || 0) + 1;
      
      if (p.dataValidade) {
        var mesAno = Utilities.formatDate(new Date(p.dataValidade), Session.getScriptTimeZone(), 'MM/yyyy');
        vencimentosPorMes[mesAno] = (vencimentosPorMes[mesAno] || 0) + 1;
      }
    });
    
    return { 
      sucesso: true, 
      resumo: { 
        total: total, vencidos: vencidos, supercritico: supercritico,
        critico: critico, alerta: alerta, seguro: seguro, 
        valorPerda: valorPerda.toFixed(2), valorPerdaProximo: valorPerdaProximo.toFixed(2) 
      }, 
      categorias: categorias, 
      estabelecimentos: estabelecimentos, 
      vencimentosPorMes: vencimentosPorMes 
    };
  } catch (e) { 
    return { sucesso: false, mensagem: e.toString() }; 
  }
}