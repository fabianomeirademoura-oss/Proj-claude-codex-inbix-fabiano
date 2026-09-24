'use strict';
// HTML e CSV do pipeline e do funil — espelho de app/lib/PaginasPipeline.ps1 e PaginasFunil.ps1.

const N = require('./numeros');
const R = require('./regras');
const P = require('./paginas');
const esc = N.esc;
const MES = P.NOMES_MES_LONGO;

function classes(lista) { return lista.length ? ` class="${lista.join(' ')}"` : ''; }

function paginaPipeline(base, vis, usuario, filtroVendedor) {
  const d = vis.Desempenho;
  const t = vis.Total;
  const q = P.queryPeriodo(d, false);
  const meses = R.mesesDisponiveis(base);
  const periodo = `${MES[d.MesInicial]} a ${MES[d.MesFinal]} de 2026`;
  const dbCrm = N.fmtData(vis.DataBaseCrm);
  const euId = usuario.idVendedor;

  const parcial = d.Parcial
    ? `<p class="parcial">Meta e realizado parciais até ${N.fmtData(d.DataBase)}: ${MES[d.MesFinal]} ainda não terminou e é comparado com a meta cheia do mês.</p>` : '';

  const orfas = vis.Abertas.filter((a) => a.Orfa);
  let avisoOrfas = '';
  if (orfas.length) {
    const grupos = new Map();
    for (const a of orfas) { if (!grupos.has(a.Vendedor.Id)) grupos.set(a.Vendedor.Id, []); grupos.get(a.Vendedor.Id).push(a); }
    const donos = [...grupos.values()].map((g) => {
      const v = g[0].Vendedor;
      return `${esc(v.Nome)} (desligado em ${N.fmtData(v.Desligamento)}): ${g.map((a) => esc(a.Oportunidade.Id)).join(', ')}`;
    }).join(' · ');
    avisoOrfas = `<p class="aviso-orfa" role="alert"><strong>${orfas.length} oportunidade(s) órfã(s), ${N.moeda(vis.ValorOrfas)}.</strong> ${donos}. Continuam no pipeline e nos totais; a reatribuição é uma ação do gestor (REGRAS_NEGOCIO.md §2.4).</p>`;
  }

  const temFora = t.ForaMesesMeta !== 0;
  const notaFora = temFora ? `<p class="nota">${N.moeda(t.ForaMesesMeta)} foram vendidos em meses sem meta do vendedor: ficam fora do realizado desta tabela (REGRAS_NEGOCIO.md §4.1).</p>` : '';

  const linhasVend = vis.Linhas.map((l) => {
    const v = l.Vendedor;
    const cls = [];
    if (v.Id === euId) cls.push('voce');
    if (N.igual(filtroVendedor, v.Id)) cls.push('selecionado');
    const atr = l.QtdAtrasadas ? `<span class="marca atraso">${l.QtdAtrasadas}</span>` : '0';
    const gap = l.Gap === null ? '—' : l.Gap <= 0 ? '<span class="nota">meta batida</span>' : N.moeda(l.Gap);
    const orfa = !v.Ativo && l.QtdAbertas ? ' <span class="marca atraso">carteira órfã</span>' : '';
    return `<tr${classes(cls)}><td><a href="/pipeline?${q}&amp;vendedor=${esc(v.Id)}#oportunidades">${esc(v.Nome)}</a> ${P.marcas(v)}${orfa}</td><td>${esc(v.Filial)}</td><td class="num">${N.moeda(l.Meta)}</td><td class="num">${N.moeda(l.RealizadoMesesMeta)}</td><td class="num">${N.pct(l.Atingimento)}</td><td class="num">${gap}</td><td class="num">${l.QtdAbertas}</td><td class="num">${atr}</td><td class="num">${N.moeda(l.Pipeline)}</td><td class="num forte">${N.moeda(l.Ponderado)}</td></tr>`;
  });
  const gapTotal = t.Gap === null ? '—' : t.Gap <= 0 ? 'meta batida' : N.moeda(t.Gap);

  const lista = vis.Abertas.filter((a) => !filtroVendedor || N.igual(a.Vendedor.Id, filtroVendedor));
  const linhaFiltro = filtroVendedor ? vis.Linhas.find((l) => N.igual(l.Vendedor.Id, filtroVendedor)) : null;
  const nomeFiltro = linhaFiltro ? linhaFiltro.Vendedor.Nome : null;
  const tituloLista = nomeFiltro ? `Oportunidades abertas de ${esc(nomeFiltro)} <a class="nota" href="/pipeline?${q}#oportunidades">ver todas</a>` : 'Oportunidades abertas';
  const qtdAtrLista = lista.filter((a) => a.Atrasada).length;
  const linhasOp = lista.map((a) => {
    const o = a.Oportunidade, v = a.Vendedor, c = a.Cliente;
    const cls = [];
    if (a.Atrasada) cls.push('atrasada');
    if (v.Id === euId) cls.push('voce');
    const sit = a.Atrasada ? `<span class="marca atraso">atrasada há ${a.DiasAtraso} dias</span>` : '<span class="nota">no prazo</span>';
    const orfa = a.Orfa ? ' <span class="marca desligado">órfã</span>' : '';
    return `<tr${classes(cls)}><td class="nw">${esc(o.Id)}</td><td>${esc(v.Nome)}${orfa}</td><td>${esc(c.Nome)} <span class="id">${esc(c.Cidade)}/${esc(c.UF)}</span></td><td>${esc(o.Produto)}</td><td class="nw">${esc(o.Etapa)}</td><td class="num">${N.prob(o.Probabilidade)}</td><td class="num">${N.moeda(o.Valor)}</td><td class="num">${N.moeda(a.Ponderado)}</td><td class="nw">${N.fmtData(o.Previsao)}</td><td class="nw">${sit}</td></tr>`;
  });

  const corpo = `<section class="filtros">
  <h1>Pipeline por vendedor</h1>
  <form method="get" action="/pipeline">
    <label>Meta e realizado de <select name="de">${P.opcoesMeses(meses, d.MesInicial)}</select></label>
    <label>até <select name="ate">${P.opcoesMeses(meses, d.MesFinal)}</select></label>
    ${filtroVendedor ? `<input type="hidden" name="vendedor" value="${esc(filtroVendedor)}">` : ''}
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Meta e realizado: <strong>${periodo}</strong>, data-base das vendas ${N.fmtData(d.DataBase)}. Pipeline: <strong>foto do CRM de ${dbCrm}</strong> (${esc(vis.NomeArquivoCrm)}), não filtrada pelo período. Atraso contado até ${dbCrm}, não até hoje.</p>
</section>
${parcial}
${avisoOrfas}
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">${N.moeda(t.Meta)}</span><span class="det">realizado ${N.moeda(t.RealizadoMesesMeta)} · ${N.pct(t.Atingimento)}</span></div>
  <div class="cartao"><span class="rotulo">Pipeline aberto</span><span class="valor">${N.moeda(t.Pipeline)}</span><span class="det">${t.QtdAbertas} oportunidades · soma do valor estimado</span></div>
  <div class="cartao destaque"><span class="rotulo">Pipeline ponderado</span><span class="valor">${N.moeda(t.Ponderado)}</span><span class="det">valor estimado × probabilidade da etapa</span></div>
  <div class="cartao alerta"><span class="rotulo">Previsão de fechamento vencida</span><span class="valor">${t.QtdAtrasadas} oportunidades</span><span class="det">${N.moeda(t.ValorAtrasado)} em pipeline atrasado</span></div>
</section>

<section class="bloco">
  <h2>Uma linha por vendedor</h2>
  <p class="nota">Todos os vendedores, inclusive desligados, para que o pipeline órfão apareça e o total feche (REGRAS_NEGOCIO.md §7.4). Não é ranking: a ordem é pelo ID. Clique num nome para filtrar as oportunidades.</p>
  <table>
    <thead><tr><th>Vendedor</th><th>Filial</th><th class="num">Meta do período</th><th class="num">Realizado</th><th class="num">Atingimento</th><th class="num">Falta p/ meta</th><th class="num">Abertas</th><th class="num">Atrasadas</th><th class="num">Pipeline aberto</th><th class="num">Ponderado</th></tr></thead>
    <tbody>${linhasVend.join('\n')}</tbody>
    <tfoot><tr><td>Total da empresa</td><td></td><td class="num">${N.moeda(t.Meta)}</td><td class="num">${N.moeda(t.RealizadoMesesMeta)}</td><td class="num">${N.pct(t.Atingimento)}</td><td class="num">${gapTotal}</td><td class="num">${t.QtdAbertas}</td><td class="num">${t.QtdAtrasadas}</td><td class="num">${N.moeda(t.Pipeline)}</td><td class="num">${N.moeda(t.Ponderado)}</td></tr></tfoot>
  </table>
  ${notaFora}
</section>

<section class="bloco" id="oportunidades">
  <div class="cab-bloco"><h2>${tituloLista}</h2><a class="botao" href="/pipeline.csv">Baixar CSV</a></div>
  <p class="nota">${lista.length} abertas, ${qtdAtrLista} com previsão de fechamento anterior a ${dbCrm} (em destaque, as mais atrasadas primeiro). Aberta = etapa diferente de Fechada Ganha e Fechada Perdida.</p>
  <table>
    <thead><tr><th>Oportunidade</th><th>Vendedor</th><th>Cliente</th><th>Produto</th><th>Etapa</th><th class="num">Prob.</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th>Previsão</th><th>Situação</th></tr></thead>
    <tbody>${linhasOp.join('\n')}</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  aberta = etapa ≠ Fechada Ganha e ≠ Fechada Perdida (§7) ·
  pipeline = soma do valor estimado; ponderado = valor × probabilidade (§7) ·
  atrasada = previsão anterior à data-base do CRM (§0.3, §7) ·
  meta e realizado só nos meses com meta, só vendas faturadas (§1, §4.1) ·
  empresa = soma ÷ soma (§4.2) ·
  órfãs de desligados ficam no pipeline (§2.4).
</footer>`;
  return P.layout(`Pipeline · CRM de ${dbCrm}`, corpo, usuario, 'pipeline');
}

function csvPipeline(vis) {
  const f = N.csvDinheiro;
  const linhas = ['ID Oportunidade;ID Vendedor;Vendedor;Status do vendedor;ID Cliente;Cliente;ID Produto;Produto;Etapa;Probabilidade;Valor estimado;Ponderado;Previsão de fechamento;Atrasada;Dias de atraso;Órfã;Linha na planilha'];
  for (const a of vis.Abertas) {
    const o = a.Oportunidade;
    linhas.push([o.Id, a.Vendedor.Id, a.Vendedor.Nome, a.Vendedor.Status, a.Cliente.Id, a.Cliente.Nome, o.IdProduto, o.Produto, o.Etapa,
      N.csvProb(o.Probabilidade), f(o.Valor), f(a.Ponderado), N.fmtData(o.Previsao),
      a.Atrasada ? 'sim' : 'não', a.DiasAtraso, a.Orfa ? 'sim' : 'não', o.Linha].map((x) => x ?? '').join(';'));
  }
  const t = vis.Total;
  linhas.push(`TOTAL (${t.QtdAbertas} abertas, ${t.QtdAtrasadas} atrasadas);;;;;;;;;;${f(t.Pipeline)};${f(t.Ponderado)};;;;;`);
  return linhas.join('\n') + '\n';
}

// ---------- funil ----------

function qtdTexto(n, singular, plural) { return n === 1 ? `1 ${singular}` : `${n} ${plural}`; }

function queryFunil(filtro, medida, troca = {}) {
  const v = { vendedor: filtro.IdVendedor, filial: filtro.IdFilial, categoria: filtro.Categoria, medida, ...troca };
  const partes = [];
  for (const k of ['vendedor', 'filial', 'categoria', 'medida']) if (v[k]) partes.push(`${k}=${N.escaparUri(v[k])}`);
  return esc(partes.join('&'));
}

function svgFunil(etapas, medida) {
  const n = (x) => N.formatarDouble(x, 0, 1);
  const medir = (e) => (medida === 'valor' ? e.Valor : e.Qtd);
  const medirAtraso = (e) => (medida === 'valor' ? e.ValorAtrasado : e.QtdAtrasadas);
  let max = 0;
  for (const e of etapas) { const m = medir(e); if (m > max) max = m; }
  const larguraMax = 400, centro = 395, alt = 50, passo = 64;
  const partes = [];
  let y = 8;
  for (const e of etapas) {
    const m = medir(e);
    const w = max > 0 ? (larguraMax * m) / max : 0;
    const x = centro - w / 2;
    const ym = y + alt / 2;
    const prob = e.Probabilidade !== null ? ` · ${N.prob(e.Probabilidade)}` : '';
    partes.push(`<text class="f-rotulo" x="0" y="${n(ym - 3)}">${esc(e.Etapa)}</text>`);
    partes.push(`<text class="f-sub" x="0" y="${n(ym + 15)}">probabilidade${prob}</text>`);
    if (m > 0) {
      partes.push(`<rect class="f-barra" x="${n(x)}" y="${y}" width="${n(w)}" height="${alt}" rx="4"><title>${esc(e.Etapa)}: ${e.Qtd} abertas, ${N.moeda(e.Valor)}</title></rect>`);
      const wa = (larguraMax * medirAtraso(e)) / max;
      if (wa > 0) partes.push(`<rect class="f-atraso" x="${n(x)}" y="${y}" width="${n(wa)}" height="${alt}" rx="4"><title>${esc(e.Etapa)}: ${e.QtdAtrasadas} atrasadas, ${N.moeda(e.ValorAtrasado)}</title></rect>`);
    } else {
      partes.push(`<line class="f-vazio" x1="${n(centro - 20)}" x2="${n(centro + 20)}" y1="${ym}" y2="${ym}"></line>`);
    }
    const principal = medida === 'valor' ? N.moeda(e.Valor) : qtdTexto(e.Qtd, 'aberta', 'abertas');
    const secundario = medida === 'valor'
      ? `${qtdTexto(e.Qtd, 'aberta', 'abertas')} · ${qtdTexto(e.QtdAtrasadas, 'atrasada', 'atrasadas')}`
      : `${N.moeda(e.Valor)} · ${qtdTexto(e.QtdAtrasadas, 'atrasada', 'atrasadas')}`;
    partes.push(`<text class="f-rotulo" x="800" y="${n(ym - 3)}" text-anchor="end">${esc(principal)}</text>`);
    partes.push(`<text class="f-sub" x="800" y="${n(ym + 15)}" text-anchor="end">${esc(secundario)}</text>`);
    y += passo;
  }
  const altura = y - passo + alt + 8;
  const resumo = etapas.map((e) => `${e.Etapa}: ${e.Qtd} abertas, ${e.QtdAtrasadas} atrasadas, ${N.moeda(e.Valor)}`).join('; ');
  return `<svg class="funil" viewBox="0 0 800 ${altura}" role="img" aria-label="Funil de oportunidades abertas por etapa. ${esc(resumo)}">${partes.join('')}</svg>`;
}

function paginaFunil(v, usuario, medida) {
  const f = v.Filtro;
  const t = v.Total;
  const dbCrm = N.fmtData(v.DataBaseCrm);
  const euId = usuario.idVendedor;
  const sel = (a, b) => (N.igual(a, b) ? ' selected' : '');

  const opVend = '<option value="">todos</option>' + v.Vendedores.map((x) => `<option value="${esc(x.Id)}"${sel(x.Id, f.IdVendedor)}>${esc(x.Nome)}${!x.Ativo ? ' (desligado)' : ''}</option>`).join('');
  const opFilial = '<option value="">todas</option>' + v.Filiais.map((x) => `<option value="${esc(x.Id)}"${sel(x.Id, f.IdFilial)}>${esc(x.Nome)}</option>`).join('');
  const opCat = '<option value="">todas</option>' + v.Categorias.map((x) => `<option value="${esc(x)}"${sel(x, f.Categoria)}>${esc(x)}</option>`).join('');
  const chkQtd = medida !== 'valor' ? ' checked' : '';
  const chkVal = medida === 'valor' ? ' checked' : '';

  const recorte = [];
  if (f.IdVendedor) { const vd = v.Vendedores.find((x) => N.igual(x.Id, f.IdVendedor)); recorte.push(`vendedor <strong>${esc(vd.Nome)}</strong> ${P.marcas(vd)}`); }
  if (f.IdFilial) recorte.push(`filial do vendedor <strong>${esc(v.Filiais.find((x) => N.igual(x.Id, f.IdFilial)).Nome)}</strong>`);
  if (f.Categoria) recorte.push(`categoria <strong>${esc(f.Categoria)}</strong>`);
  const textoRecorte = recorte.length ? `Mostrando: ${recorte.join(' · ')} · <a href="/funil?${queryFunil({}, medida)}">limpar filtros</a>` : 'Mostrando: todo o CRM, sem filtro.';

  const linhasEtapa = v.Etapas.map((e) => {
    const pct = t.Valor > 0 ? N.pct(N.fracao(e.Valor, t.Valor)) : '—';
    const atr = e.QtdAtrasadas ? `<span class="marca atraso">${e.QtdAtrasadas}</span> <span class="id">${N.moeda(e.ValorAtrasado)}</span>` : '0';
    const prob = e.Probabilidade !== null ? N.prob(e.Probabilidade) : '—';
    return `<tr><td>${esc(e.Etapa)}</td><td class="num">${prob}</td><td class="num">${e.Qtd}</td><td class="num">${atr}</td><td class="num">${N.moeda(e.Valor)}</td><td class="num forte">${N.moeda(e.Ponderado)}</td><td class="num">${pct}</td></tr>`;
  });

  let maxMotivo = 0;
  for (const m of v.Motivos) if (m.Qtd > maxMotivo) maxMotivo = m.Qtd;
  const linhasMotivo = v.Motivos.map((m) => `<tr><td>${esc(m.Motivo)}</td><td><meter min="0" max="${maxMotivo}" value="${m.Qtd}"></meter></td><td class="num">${m.Qtd}</td><td class="num">${N.moeda(m.Valor)}</td></tr>`);
  const conv = v.Conversao !== null ? N.pct(v.Conversao) : '—';
  const convValor = v.ConversaoValor !== null ? N.pct(v.ConversaoValor) : '—';
  const tabelaMotivos = v.Motivos.length
    ? `<table><thead><tr><th>Motivo da perda</th><th></th><th class="num">Perdidas</th><th class="num">Valor estimado</th></tr></thead><tbody>${linhasMotivo.join('\n')}</tbody></table>`
    : '<p class="nota">Nenhuma oportunidade perdida neste recorte.</p>';

  const linhasOp = v.Abertas.map((a) => {
    const o = a.Oportunidade, vd = a.Vendedor, c = a.Cliente;
    const cls = [];
    if (a.Atrasada) cls.push('atrasada');
    if (vd.Id === euId) cls.push('voce');
    const sit = a.Atrasada ? `<span class="marca atraso">atrasada há ${a.DiasAtraso} dias</span>` : '<span class="nota">no prazo</span>';
    const orfa = a.Orfa ? ' <span class="marca desligado">órfã</span>' : '';
    return `<tr${classes(cls)}><td class="nw">${esc(o.Etapa)}</td><td class="nw">${esc(o.Id)}</td><td><a href="/funil?${queryFunil(f, medida, { vendedor: vd.Id })}">${esc(vd.Nome)}</a>${orfa} <span class="id">${esc(vd.Filial)}</span></td><td>${esc(c.Nome)}</td><td>${esc(o.Produto)}<span class="item">${esc(a.Produto.Categoria)}</span></td><td class="num">${N.moeda(o.Valor)}</td><td class="num">${N.moeda(a.Ponderado)}</td><td class="nw">${N.fmtData(o.Previsao)}</td><td class="nw">${sit}</td></tr>`;
  });
  const vazio = !v.Abertas.length ? '<tr><td colspan="9" class="nota">Nenhuma oportunidade aberta neste recorte.</td></tr>' : '';
  const rotuloMedida = medida === 'valor' ? 'valor estimado' : 'quantidade de oportunidades';

  const corpo = `<section class="filtros">
  <h1>Funil de vendas</h1>
  <form method="get" action="/funil">
    <label>Vendedor <select name="vendedor">${opVend}</select></label>
    <label>Filial <select name="filial">${opFilial}</select></label>
    <label>Categoria <select name="categoria">${opCat}</select></label>
    <span class="check">Largura das barras:
      <label class="check"><input type="radio" name="medida" value="qtd"${chkQtd}> quantidade</label>
      <label class="check"><input type="radio" name="medida" value="valor"${chkVal}> valor</label>
    </span>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">${textoRecorte}</p>
  <p class="fonte">Foto do CRM de <strong>${dbCrm}</strong> (${esc(v.NomeArquivoCrm)}), não filtrada por período. Atraso contado até ${dbCrm}. Filial = filial do vendedor; categoria = categoria do produto no catálogo.</p>
</section>
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Oportunidades abertas</span><span class="valor">${t.Qtd}</span><span class="det">${N.moeda(t.Valor)} em valor estimado</span></div>
  <div class="cartao destaque"><span class="rotulo">Pipeline ponderado</span><span class="valor">${N.moeda(t.Ponderado)}</span><span class="det">valor estimado × probabilidade da etapa</span></div>
  <div class="cartao alerta"><span class="rotulo">Previsão de fechamento vencida</span><span class="valor">${qtdTexto(t.QtdAtrasadas, 'aberta', 'abertas')}</span><span class="det">${N.moeda(t.ValorAtrasado)} em pipeline atrasado</span></div>
  <div class="cartao"><span class="rotulo">Conversão das fechadas</span><span class="valor">${conv}</span><span class="det">${v.Ganhas.Qtd} ganhas ÷ ${v.Ganhas.Qtd + v.Perdidas.Qtd} fechadas</span></div>
</section>

<section class="bloco">
  <h2>Onde estão as oportunidades abertas hoje</h2>
  <p class="definicao">Cada barra é o total de oportunidades <strong>abertas</strong> que estão hoje naquela etapa, medido por ${rotuloMedida}. O CRM guarda só a etapa atual: o funil não mostra quantas oportunidades passaram por cada etapa (REGRAS_NEGOCIO.md §8).</p>
  <p class="legenda"><span class="cor cor-dia"></span> no prazo <span class="cor cor-atraso"></span> previsão de fechamento já passou</p>
  ${svgFunil(v.Etapas, medida)}
  <table>
    <thead><tr><th>Etapa</th><th class="num">Prob.</th><th class="num">Abertas</th><th class="num">Atrasadas</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th class="num">% do valor aberto</th></tr></thead>
    <tbody>${linhasEtapa.join('\n')}</tbody>
    <tfoot><tr><td>Total aberto</td><td></td><td class="num">${t.Qtd}</td><td class="num">${t.QtdAtrasadas}</td><td class="num">${N.moeda(t.Valor)}</td><td class="num">${N.moeda(t.Ponderado)}</td><td class="num">${t.Valor > 0 ? '100,0%' : '—'}</td></tr></tfoot>
  </table>
</section>

<section class="rankings">
  <div class="bloco">
    <h2>Fechadas: ganhas × perdidas</h2>
    <table>
      <thead><tr><th></th><th class="num">Oportunidades</th><th class="num">Valor estimado</th></tr></thead>
      <tbody>
        <tr><td>Fechadas Ganhas</td><td class="num">${v.Ganhas.Qtd}</td><td class="num">${N.moeda(v.Ganhas.Valor)}</td></tr>
        <tr><td>Fechadas Perdidas</td><td class="num">${v.Perdidas.Qtd}</td><td class="num">${N.moeda(v.Perdidas.Valor)}</td></tr>
      </tbody>
      <tfoot><tr><td>Conversão</td><td class="num">${conv}</td><td class="num">${convValor}</td></tr></tfoot>
    </table>
    <p class="nota">Conversão = ganhas ÷ (ganhas + perdidas), em quantidade e, ao lado, em valor estimado. Abertas não entram.</p>
  </div>
  <div class="bloco">
    <h2>Por que perdemos</h2>
    ${tabelaMotivos}
  </div>
</section>

<section class="bloco" id="abertas">
  <h2>Oportunidades abertas deste recorte</h2>
  <p class="nota">${v.Abertas.length} abertas, ${t.QtdAtrasadas} com previsão de fechamento anterior a ${dbCrm} (em destaque). Da etapa mais avançada para a menos avançada; dentro da etapa, atrasadas primeiro. Clique num vendedor para filtrar o funil por ele.</p>
  <table>
    <thead><tr><th>Etapa</th><th>Oportunidade</th><th>Vendedor</th><th>Cliente</th><th>Produto</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th>Previsão</th><th>Situação</th></tr></thead>
    <tbody>${linhasOp.join('\n')}${vazio}</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  aberta = etapa ≠ Fechada Ganha e ≠ Fechada Perdida (§7) ·
  barra = abertas na etapa atual, sem histórico de etapas (§8) ·
  atrasada = previsão anterior à data-base do CRM (§7) ·
  conversão = ganhas ÷ (ganhas + perdidas) (§8) ·
  filial = filial do vendedor (§8.4) ·
  órfãs de desligados ficam no funil (§2.4).
</footer>`;
  return P.layout(`Funil de vendas · CRM de ${dbCrm}`, corpo, usuario, 'funil');
}

module.exports = { paginaPipeline, csvPipeline, paginaFunil };
