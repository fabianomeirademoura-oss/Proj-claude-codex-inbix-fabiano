'use strict';
// HTML e CSV do estoque e da origem do faturamento — espelho de app/lib/PaginasEstoque.ps1 e PaginasOrigem.ps1.

const N = require('./numeros');
const R = require('./regras');
const E = require('./regras_estoque');
const P = require('./paginas');
const esc = N.esc;
const MES = P.NOMES_MES_LONGO;

function frasePrazoParado(v) {
  return `Parado é a linha de estoque (um produto numa filial) que tem pelo menos 1 unidade e está há ${v.DiasParado} dias ou mais sem nenhuma entrada nem saída, contados até a data da foto (${N.fmtData(v.DataBase)}); linha sem nenhuma data de movimento também conta como parada.`;
}

function ops(a) {
  if (a === null) return '<span class="nota">CRM indisponível</span>';
  if (a.Qtd === 0) return '<span class="nota">nenhuma</span>';
  return `<strong>${a.Qtd}</strong> <span class="id">${N.moeda(a.Valor)}</span>`;
}

function outrasFiliais(outras, comSaida) {
  // "Estoque é por produto E por filial": o mesmo produto nas outras filiais.
  return outras.map((o) => {
    const saida = comSaida ? (o.UltimaSaida !== null ? `, última saída ${N.fmtData(o.UltimaSaida)}` : ', sem saída em 2026') : '';
    const cls = o.Quantidade > o.Minimo ? 'item quebra forte' : 'item quebra';
    return `<span class="${cls}">${esc(o.Filial)}: ${o.Quantidade} (mín. ${o.Minimo})${saida}</span>`;
  }).join('');
}

function movimento(l) {
  if (l.DiasSemMovimento === null) return '<span class="marca atraso">sem movimento registrado</span>';
  return `${l.DiasSemMovimento} dias`;
}

function queryEstoque(v, idFilial) {
  let q = `dias=${v.DiasParado}`;
  if (idFilial) q += `&amp;filial=${esc(idFilial)}`;
  return q;
}

function paginaEstoque(v, usuario, filtroFilial, avisoPrazo) {
  const db = N.fmtData(v.DataBase);
  const t = v.Total;
  const q = queryEstoque(v, filtroFilial);
  const filialFiltro = filtroFilial ? v.Filiais.find((f) => N.igual(f.Id, filtroFilial)) : null;
  const nomeFiltro = filialFiltro ? filialFiltro.Nome : null;
  const opcoesFilial = '<option value="">todas</option>' + v.Filiais.map((f) => `<option value="${esc(f.Id)}"${N.igual(f.Id, filtroFilial) ? ' selected' : ''}>${esc(f.Nome)}</option>`).join('');
  const aviso = avisoPrazo ? `<p class="parcial" role="alert">${esc(avisoPrazo)}</p>` : '';
  const semCrm = !v.ComCrm ? '<p class="parcial">O CRM não pôde ser carregado: a coluna de oportunidades abertas está indisponível. Veja a página Pipeline para os erros.</p>' : '';
  const noFiltro = (l) => !filtroFilial || N.igual(l.IdFilial, filtroFilial);

  // 1. Matriz categoria × filial
  const cab = v.Filiais.map((f) => `<th class="num"><a href="/estoque?dias=${v.DiasParado}&amp;filial=${esc(f.Id)}#parados">${esc(f.Nome)}</a></th>`).join('');
  const cel = (c) => `<td class="num">${N.moeda(c.Valor)}${c.Parado > 0 ? `<span class="item alerta-txt">parado ${N.moeda(c.Parado)}</span>` : ''}</td>`;
  const linhasMatriz = v.Categorias.map((cat) => {
    const celulas = v.Filiais.map((f) => cel(v.Celulas.get(`${cat}|${f.Id}`))).join('');
    const tc = v.TotCategoria.get(cat);
    const pct = t.Valor > 0 ? N.pct(N.fracao(tc.Valor, t.Valor)) : '—';
    return `<tr><td>${esc(cat)}</td>${celulas}${cel(tc)}<td class="num">${pct}</td></tr>`;
  });
  const rodape = v.Filiais.map((f) => cel(v.TotFilial.get(f.Id))).join('');

  // 2. Parados
  const parados = v.Parados.filter((p) => noFiltro(p.Linha));
  let valorParadosLista = 0;
  for (const p of parados) valorParadosLista += p.Linha.Valor;
  const linhasParados = parados.map((p) => {
    const l = p.Linha;
    const desc = l.Produto.Descontinuado ? ' <span class="marca desligado">descontinuado</span>' : '';
    return `<tr><td class="nw">${esc(l.Produto.Id)}</td><td>${esc(l.Produto.Nome)}${desc}<span class="item">${esc(l.Produto.Categoria)}</span></td><td>${esc(l.Filial)}</td><td class="num">${l.Quantidade}</td><td class="num">${N.moeda(l.Valor)}</td><td class="nw">${N.fmtData(l.UltimaEntrada)}</td><td class="nw">${N.fmtData(l.UltimaSaida)}</td><td class="num">${movimento(l)}</td><td class="num">${ops(p.Abertas)}</td><td>${outrasFiliais(p.Outras, true)}</td></tr>`;
  });
  const vazioParados = !parados.length ? '<tr><td colspan="10" class="nota">Nenhuma linha parada com este prazo.</td></tr>' : '';

  // 3. Sem estoque em nenhuma filial
  const linhasSem = v.SemEstoque.map((s) => {
    const desc = s.Produto.Descontinuado ? ' <span class="marca desligado">descontinuado</span>' : '';
    const mins = s.Linhas.map((l) => `<span class="item">${esc(l.Filial)}: mín. ${l.Minimo}</span>`).join('');
    const cls = s.Abertas && s.Abertas.Qtd > 0 ? ' class="atrasada"' : '';
    const saida = s.UltimaSaida !== null ? N.fmtData(s.UltimaSaida) : '<span class="nota">sem saída em 2026</span>';
    return `<tr${cls}><td class="nw">${esc(s.Produto.Id)}</td><td>${esc(s.Produto.Nome)}${desc}</td><td>${esc(s.Produto.Categoria)}</td><td>${mins}</td><td class="nw">${saida}</td><td class="num">${ops(s.Abertas)}</td></tr>`;
  });

  // 4. Abaixo do mínimo
  const abaixo = v.Abaixo.filter((a) => noFiltro(a.Linha));
  const linhasAbaixo = abaixo.map((a) => {
    const l = a.Linha;
    const zerada = l.Quantidade === 0 ? ' <span class="marca atraso">zerado</span>' : '';
    return `<tr><td>${esc(l.Filial)}</td><td class="nw">${esc(l.Produto.Id)}</td><td>${esc(l.Produto.Nome)}<span class="item">${esc(l.Produto.Categoria)}</span></td><td class="num">${l.Quantidade}${zerada}</td><td class="num">${l.Minimo}</td><td class="num forte">${a.Falta}</td><td>${outrasFiliais(a.Outras, false)}</td><td class="num">${ops(a.Abertas)}</td></tr>`;
  });
  const sufixoFiltro = nomeFiltro ? ` em ${esc(nomeFiltro)} <a class="nota" href="/estoque?dias=${v.DiasParado}">ver todas as filiais</a>` : '';
  const semComOp = v.SemEstoque.filter((s) => s.Abertas && s.Abertas.Qtd).length;

  const corpo = `<section class="filtros">
  <h1>Estoque</h1>
  <form method="get" action="/estoque">
    <label>Parado a partir de <input type="number" name="dias" min="1" max="${E.DIAS_PARADO_MAXIMO}" step="1" value="${v.DiasParado}" class="curto"> dias sem movimento</label>
    <label>Filial nas listas <select name="filial">${opcoesFilial}</select></label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Foto do estoque de <strong>${db}</strong> (${esc(v.NomeArquivo)}). Dias contados até ${db}, não até hoje. Estoque é por produto <em>e</em> por filial: cada linha abaixo é um produto numa filial. Padrão do prazo: ${E.DIAS_PARADO_PADRAO} dias (REGRAS_NEGOCIO.md §5).</p>
</section>
${aviso}
${semCrm}
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Dinheiro no pátio (valor imobilizado)</span><span class="valor">${N.moeda(t.Valor)}</span><span class="det">${t.Unidades} unidades · quantidade × custo médio</span></div>
  <div class="cartao alerta"><span class="rotulo">Parado há ${v.DiasParado}+ dias</span><span class="valor">${N.moeda(t.Parado)}</span><span class="det">${v.Parados.length} linhas (produto, filial)</span></div>
  <div class="cartao"><span class="rotulo">Sem estoque em nenhuma filial</span><span class="valor">${v.SemEstoque.length} produtos</span><span class="det">${semComOp} com oportunidade aberta no CRM</span></div>
  <div class="cartao"><span class="rotulo">Abaixo do estoque mínimo</span><span class="valor">${v.Abaixo.length} linhas</span><span class="det">quantidade &lt; mínimo na filial</span></div>
</section>

<section class="bloco">
  <h2>1. Quanto dinheiro está parado no pátio, por filial e por categoria</h2>
  <p class="nota">Valor imobilizado = quantidade × custo médio de cada linha. Em vermelho, a parte que está parada há ${v.DiasParado} dias ou mais. Clique numa filial para ver os parados e os itens a repor dela.</p>
  <table>
    <thead><tr><th>Categoria</th>${cab}<th class="num">Total</th><th class="num">% do total</th></tr></thead>
    <tbody>${linhasMatriz.join('\n')}</tbody>
    <tfoot><tr><td>Total</td>${rodape}${cel(t)}<td class="num">100,0%</td></tr></tfoot>
  </table>
</section>

<section class="bloco" id="parados">
  <h2>2. Produtos parados${sufixoFiltro}</h2>
  <p class="definicao"><strong>Definição usada:</strong> ${esc(frasePrazoParado(v))} Para usar outro prazo, mude o campo no topo da página.</p>
  <p class="nota">${parados.length} linhas, ${N.moeda(valorParadosLista)}. Mais tempo parado primeiro. "Oportunidades abertas" é do produto no CRM, em qualquer filial: parado não quer dizer sem demanda. A última coluna mostra o mesmo produto nas outras filiais; em negrito, onde há mais que o mínimo.</p>
  <table>
    <thead><tr><th>ID</th><th>Produto</th><th>Filial</th><th class="num">Qtd.</th><th class="num">Valor parado</th><th>Última entrada</th><th>Última saída</th><th class="num">Sem movimento</th><th class="num">Oport. abertas</th><th>Nas outras filiais</th></tr></thead>
    <tbody>${linhasParados.join('\n')}${vazioParados}</tbody>
  </table>
</section>

<section class="bloco">
  <h2>3. O que está sem estoque em lugar nenhum</h2>
  <p class="nota">Produtos com quantidade zero nas ${v.Filiais.length} filiais somadas. Em destaque, os que têm oportunidade aberta no CRM: há cliente esperando e não há máquina no pátio.</p>
  <table>
    <thead><tr><th>ID</th><th>Produto</th><th>Categoria</th><th>Estoque mínimo</th><th>Última saída</th><th class="num">Oport. abertas</th></tr></thead>
    <tbody>${linhasSem.join('\n')}</tbody>
  </table>
</section>

<section class="bloco">
  <div class="cab-bloco"><h2>4. Abaixo do estoque mínimo: precisa repor${sufixoFiltro}</h2><a class="botao" href="/estoque.csv?${q}">Baixar CSV</a></div>
  <p class="nota">${abaixo.length} linhas em que a quantidade da filial é menor que o mínimo da filial. "Faltam" = mínimo − quantidade. Antes de comprar, veja se outra filial tem sobra do mesmo produto (em negrito).</p>
  <table>
    <thead><tr><th>Filial</th><th>ID</th><th>Produto</th><th class="num">Qtd.</th><th class="num">Mínimo</th><th class="num">Faltam</th><th>Nas outras filiais</th><th class="num">Oport. abertas</th></tr></thead>
    <tbody>${linhasAbaixo.join('\n')}</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md §5):</strong>
  tudo por linha (produto, filial) ·
  valor imobilizado = quantidade × custo médio ·
  parado = quantidade &gt; 0 e dias sem entrada nem saída ≥ prazo ·
  sem estoque = soma das filiais = 0 ·
  abaixo do mínimo = quantidade &lt; mínimo na linha ·
  dias contados até a data da foto, nunca até hoje.
</footer>`;
  return P.layout(`Estoque · foto de ${db}`, corpo, usuario, 'estoque');
}

function csvEstoque(v) {
  const f = N.csvDinheiro;
  const semEstoque = new Set(v.SemEstoque.map((s) => s.Produto.Id));
  const linhas = [`ID Produto;Produto;Categoria;Status do produto;Filial;Quantidade;Estoque mínimo;Custo médio;Valor imobilizado;Última entrada;Última saída;Dias sem movimento;Parado (prazo ${v.DiasParado} dias);Abaixo do mínimo;Produto sem estoque em nenhuma filial;Linha na planilha`];
  for (const l of N.ordenar(v.Linhas, (x) => x.Produto.Id, (x) => x.IdFilial)) {
    linhas.push([l.Produto.Id, l.Produto.Nome, l.Produto.Categoria, l.Produto.Status, l.Filial, l.Quantidade, l.Minimo, f(l.CustoMedio), f(l.Valor),
      N.fmtData(l.UltimaEntrada), N.fmtData(l.UltimaSaida), l.DiasSemMovimento === null ? 'sem movimento' : l.DiasSemMovimento,
      E.linhaParada(l, v.DiasParado) ? 'sim' : 'não', E.abaixoMinimo(l) ? 'sim' : 'não', semEstoque.has(l.Produto.Id) ? 'sim' : 'não', l.Linha].map((x) => x ?? '').join(';'));
  }
  const t = v.Total;
  linhas.push(`TOTAL;;;;;${t.Unidades};;;${f(t.Valor)};;;;${v.Parados.length} linhas, ${f(t.Parado)};${v.Abaixo.length} linhas;${v.SemEstoque.length} produtos;`);
  return linhas.join('\n') + '\n';
}

function paginaOrigem(base, v, usuario) {
  const meses = R.mesesDisponiveis(base);
  const pct = v.FracaoCrm !== null ? N.pct(v.FracaoCrm) : '—';
  const pctSem = v.FracaoSem !== null ? N.pct(v.FracaoSem) : '—';
  const resumo = v.FracaoCrm !== null
    ? `<strong>${pct} do faturamento nasceu em oportunidades registradas no CRM.</strong> São ${v.QtdComCrm} vendas faturadas, somando ${N.moeda(v.ComCrm)}.`
    : 'Sem faturamento no período. A participação do CRM não se aplica.';
  const parcial = v.Parcial ? `<p class="parcial">Faturamento parcial até ${N.fmtData(v.DataBase)}.</p>` : '';
  const categorias = v.CategoriasSem.map((c) => `<tr><td>${esc(c.Categoria)}</td><td class="num">${c.Quantidade}</td><td class="num">${N.moeda(c.Valor)}</td></tr>`);
  const vazioCat = !v.CategoriasSem.length ? '<tr><td colspan="3">Nenhuma venda faturada sem oportunidade neste período.</td></tr>' : '';
  const linhas = v.Linhas.map((l) => {
    const s = l.Venda, o = l.Oportunidade;
    const origem = o ? `${esc(o.Id)}<span class="item">${esc(o.Etapa)}</span><span class="item">Origem: ${esc(o.Origem)} · linha ${o.Linha} do CRM</span>` : 'Sem oportunidade registrada';
    const classe = !l.Faturada ? ' class="cancelada"' : '';
    const status = l.Faturada ? 'Faturada' : 'Cancelada · fora do faturamento';
    return `<tr${classe}><td class="nw">${esc(s.Id)}<span class="item">${N.fmtData(s.Data)} · linha ${s.Linha}</span></td><td>${esc(s.Cliente)}<span class="item quebra">${esc(s.Produto)}</span></td><td>${origem}</td><td class="num">${N.moeda(s.Valor)}</td><td>${status}</td></tr>`;
  });
  const vazio = !v.Linhas.length ? '<tr><td colspan="5">Nenhuma venda no período.</td></tr>' : '';
  const corpo = `<section class="filtros">
  <h1>Origem do faturamento</h1>
  <form method="get" action="/origem-vendas">
    <label>De <select name="de">${P.opcoesMeses(meses, v.MesInicial)}</select></label>
    <label>até <select name="ate">${P.opcoesMeses(meses, v.MesFinal)}</select></label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Período: ${MES[v.MesInicial]} a ${MES[v.MesFinal]} de 2026 · vendas até ${N.fmtData(v.DataBase)} · foto do CRM: ${N.fmtData(v.DataBaseCrm)}.</p>
</section>
${parcial}
<section class="bloco"><p>${resumo}</p><p class="nota">Participação = valor das vendas faturadas com oportunidade vinculada ÷ valor de todas as vendas faturadas do período. Usamos o valor vendido, já com desconto, e não o valor estimado da oportunidade.</p></section>
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Faturamento total</span><span class="valor">${N.moeda(v.Total)}</span><span class="det">${v.QtdFaturadas} vendas faturadas</span></div>
  <div class="cartao destaque"><span class="rotulo">Com origem no CRM · ${pct}</span><span class="valor">${N.moeda(v.ComCrm)}</span><span class="det">${v.QtdComCrm} vendas vinculadas por ID Oportunidade</span></div>
  <div class="cartao"><span class="rotulo">Sem oportunidade · ${pctSem}</span><span class="valor">${N.moeda(v.SemOportunidade)}</span><span class="det">${v.QtdSemOportunidade} vendas sem vínculo registrado</span></div>
</section>
<section class="bloco">
  <h2>O que são as vendas sem oportunidade?</h2>
  <p>São vendas cuja coluna <strong>ID Oportunidade está vazia</strong>. Nesta base didática, são descritas como <strong>vendas de balcão</strong>: faturamentos sem uma oportunidade registrada no CRM.</p>
  <p>Isso não significa ausência de trabalho comercial, nem que sejam somente peças. A composição abaixo mostra as categorias efetivamente vendidas sem esse vínculo.</p>
  <p class="nota">Não inferimos vínculos por cliente, produto ou nome. Uma venda sem ID continua sem oportunidade registrada, mesmo que exista uma oportunidade parecida. Este indicador mede a origem documentada no CRM; não é taxa de conversão do funil nem mede, sozinho, a influência de todo o trabalho comercial.</p>
  <table><thead><tr><th>Categoria sem oportunidade</th><th class="num">Vendas faturadas</th><th class="num">Faturamento</th></tr></thead><tbody>${categorias.join('\n')}${vazioCat}</tbody><tfoot><tr><td>Total sem oportunidade</td><td class="num">${v.QtdSemOportunidade}</td><td class="num">${N.moeda(v.SemOportunidade)}</td></tr></tfoot></table>
</section>
<section class="bloco">
  <h2>Conferência: cada venda e sua oportunidade</h2>
  <p class="nota">${v.Linhas.length} registros no período. As ${v.QtdCanceladas} canceladas, no valor de ${N.moeda(v.ValorCancelado)}, aparecem riscadas e não entram nos indicadores. Todos os vendedores estão incluídos, inclusive o histórico dos desligados.</p>
  <p class="fonte">Fontes: vendas_2026_jan-ago.xlsx, aba Vendas; crm_oportunidades.xlsx, aba Oportunidades. Os números de linha permitem conferir o vínculo nas planilhas.</p>
  <details><summary>Ver vendas e vínculos (${v.Linhas.length} registros)</summary>
  <table><thead><tr><th>Venda / data / linha</th><th>Cliente / produto</th><th>Oportunidade / etapa / origem</th><th class="num">Valor da venda</th><th>Status</th></tr></thead><tbody>${linhas.join('\n')}${vazio}</tbody></table>
  </details>
</section>
<footer class="regras">REGRAS_NEGOCIO.md §0.1: vínculo pelo ID · §1: somente faturadas no realizado; vendas com e sem oportunidade contam no faturamento · §2.1: histórico dos desligados preservado. O filtro usa a data da venda, não a data de criação da oportunidade.</footer>`;
  return P.layout('Origem do faturamento', corpo, usuario, 'origem');
}

module.exports = { paginaEstoque, csvEstoque, paginaOrigem };
