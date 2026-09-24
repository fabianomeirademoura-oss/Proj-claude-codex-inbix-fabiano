'use strict';
// HTML e CSV — espelho de app/lib/Paginas.ps1. Só formata: todo número vem de regras.js.
// Os textos seguem o PowerShell caractere por caractere (testes/paridade_node.ps1 compara as duas versões).

const fs = require('fs');
const path = require('path');
const N = require('./numeros');
const R = require('./regras');

const NOMES_MES = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const NOMES_MES_LONGO = ['', 'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const esc = N.esc;

function formatarMeses(meses) {
  if (!meses || !meses.length) return 'nenhum';
  const partes = [];
  let inicio = meses[0], anterior = meses[0];
  for (let i = 1; i <= meses.length; i++) {
    const atual = i < meses.length ? meses[i] : -1;
    if (atual !== anterior + 1) {
      partes.push(inicio === anterior ? NOMES_MES[inicio] : `${NOMES_MES[inicio]}–${NOMES_MES[anterior]}`);
      inicio = atual;
    }
    anterior = atual;
  }
  return partes.join(', ');
}

// §2.3 e §3.3: desligado e admitido no ano aparecem marcados.
function marcas(v) {
  if (!v.Ativo) return `<span class="marca desligado">desligado em ${N.fmtData(v.Desligamento)}</span>`;
  if (v.Admissao !== null && N.anoDe(v.Admissao) === 2026) return `<span class="marca admitido">admissão em ${N.fmtMesAno(v.Admissao)}</span>`;
  return '';
}

function medidor(f) {
  if (f === null) return '';
  const v = N.formatarDouble(Math.min(N.fracaoParaNumero(f), 1.5), 4, 4);
  return `<meter min="0" max="1.5" low="0.8" high="0.9999" optimum="1.2" value="${v}"></meter>`;
}

function queryPeriodo(d, incluirDesligados) {
  let q = `de=${d.MesInicial}&amp;ate=${d.MesFinal}`;
  if (incluirDesligados) q += '&amp;desligados=1';
  return q;
}

function opcoesMeses(meses, selecionado) {
  return meses.map((m) => `<option value="${m}"${m === selecionado ? ' selected' : ''}>${NOMES_MES_LONGO[m]}</option>`).join('');
}

function layout(titulo, corpo, usuario, secao = 'vendas') {
  const ativo = (s) => (s === secao ? ' class="ativo" aria-current="page"' : '');
  const topo = usuario ? `<header class="topo">
  <nav class="nav">
    <span class="marca-app">Horizonte Máquinas</span>
    <a href="/"${ativo('vendas')}>Vendas e metas</a>
    <a href="/origem-vendas"${ativo('origem')}>Origem do faturamento</a>
    <a href="/pipeline"${ativo('pipeline')}>Pipeline</a>
    <a href="/funil"${ativo('funil')}>Funil</a>
    <a href="/estoque"${ativo('estoque')}>Estoque</a>
    ${N.igual(usuario.perfil, 'diretoria') ? `<a href="/importar"${ativo('importar')}>Importar</a>` : ''}
  </nav>
  <div class="usuario">${esc(usuario.nome)}
    <form method="post" action="/sair"><button type="submit" class="link">Sair</button></form>
  </div>
</header>` : '';
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(titulo)}</title>
<link rel="stylesheet" href="/estilo.css">
</head>
<body>
${topo}
<main>
${corpo}
</main>
</body>
</html>`;
}

function paginaLogin(mensagem, loginDigitado) {
  const aviso = mensagem ? `<p class="erro" role="alert">${esc(mensagem)}</p>` : '';
  const corpo = `<section class="login">
  <h1>Horizonte Máquinas</h1>
  <p class="sub">Desempenho comercial 2026</p>
  ${aviso}
  <form method="post" action="/entrar">
    <label>Usuário (e-mail)<input name="login" type="text" autocomplete="username" required value="${esc(loginDigitado)}"></label>
    <label>Senha<input name="senha" type="password" autocomplete="current-password" required></label>
    <button type="submit">Entrar</button>
  </form>
</section>`;
  return layout('Entrar · Horizonte Máquinas', corpo, null);
}

function paginaErroImportacao(erros, usuario, secao = 'vendas') {
  const itens = erros.slice(0, 50).map((e) => `<li>${esc(e)}</li>`).join('\n');
  const corpo = `<section class="aviso-erro">
  <h1>Os dados não puderam ser carregados</h1>
  <p>A importação encontrou ${erros.length} problema(s). Nenhum número é exibido até que as planilhas sejam corrigidas (regras do REGRAS_NEGOCIO.md).</p>
  <ul>${itens}</ul>
</section>`;
  return layout('Erro de importação', corpo, usuario, secao);
}

function paginaPainel(base, d, usuario, incluirDesligados, fuso) {
  const e = d.Empresa;
  const q = queryPeriodo(d, incluirDesligados);
  const meses = R.mesesDisponiveis(base);
  const marcado = incluirDesligados ? ' checked' : '';
  const periodo = `${NOMES_MES_LONGO[d.MesInicial]} a ${NOMES_MES_LONGO[d.MesFinal]} de 2026`;
  const fontes = base.Arquivos.map((a) => `${esc(a.Nome)} (${N.fmtDataHora(a.Modificado, fuso)})`).join(' · ');
  const euId = usuario.idVendedor;
  const classeVoce = (v) => (v.Id === euId ? ' class="voce"' : '');

  const parcial = d.Parcial
    ? `<p class="parcial">Resultado parcial até ${N.fmtData(d.DataBase)}: ${NOMES_MES_LONGO[d.MesFinal]} ainda não terminou e é comparado com a meta cheia do mês.</p>` : '';
  const foraMeta = e.ForaMesesMeta !== 0
    ? `<p class="nota">${N.moeda(e.ForaMesesMeta)} foram vendidos em meses em que o vendedor não tinha meta: entram no faturamento, não no atingimento (REGRAS_NEGOCIO.md §4.1).</p>` : '';

  const linhasFat = R.rankingFaturamento(d, incluirDesligados).map((l, i) => {
    const v = l.Vendedor;
    const part = e.Realizado > 0 ? N.pct(N.fracao(l.Realizado, e.Realizado)) : '—';
    return `<tr${classeVoce(v)}><td class="pos">${i + 1}</td><td><a href="/vendedor?id=${esc(v.Id)}&amp;${q}">${esc(v.Nome)}</a> ${marcas(v)}</td><td>${esc(v.Filial)}</td><td class="num">${l.MesesComMeta.length}</td><td class="num">${N.moeda(l.Realizado)}</td><td class="num">${part}</td></tr>`;
  });
  const linhasAt = R.rankingAtingimento(d, incluirDesligados).map((l, i) => {
    const v = l.Vendedor;
    return `<tr${classeVoce(v)}><td class="pos">${i + 1}</td><td><a href="/vendedor?id=${esc(v.Id)}&amp;${q}">${esc(v.Nome)}</a> ${marcas(v)}</td><td class="num">${N.moeda(l.Meta)}</td><td class="num">${N.moeda(l.RealizadoMesesMeta)}</td><td class="num"><span class="at">${medidor(l.Atingimento)} ${N.pct(l.Atingimento)}</span></td></tr>`;
  });
  const semMeta = d.Linhas.filter((l) => (incluirDesligados || l.Vendedor.Ativo) && l.Atingimento === null);
  const notaSemMeta = semMeta.length
    ? `<p class="nota">Sem meta no período (fora do ranking de atingimento): ${semMeta.map((l) => esc(l.Vendedor.Nome)).join(', ')}.</p>` : '';
  const foraRanking = d.Linhas.filter((l) => !l.Vendedor.Ativo);
  const notaDesligados = !incluirDesligados && foraRanking.length
    ? `<p class="nota">Desligados não entram nos rankings (REGRAS_NEGOCIO.md §2.3): ${foraRanking.map((l) => esc(l.Vendedor.Nome)).join(', ')}. Os números deles estão no quadro de conferência e nos totais.</p>` : '';

  // Quadro de conferência: todos os vendedores, total = empresa
  const temFora = d.Linhas.some((l) => l.ForaMesesMeta !== 0);
  const cabFora = temFora ? '<th class="num">Realizado fora dos meses c/ meta</th>' : '';
  const rotuloReal = temFora ? 'Realizado nos meses c/ meta' : 'Realizado';
  const linhasQuadro = N.ordenar(d.Linhas, (l) => l.Vendedor.Id).map((l) => {
    const v = l.Vendedor;
    const celFora = temFora ? `<td class="num">${N.moeda(l.ForaMesesMeta)}</td>` : '';
    return `<tr${classeVoce(v)}><td>${esc(v.Id)}</td><td><a href="/vendedor?id=${esc(v.Id)}&amp;${q}">${esc(v.Nome)}</a> ${marcas(v)}</td><td>${esc(v.Filial)}</td><td class="nw">${formatarMeses(l.MesesComMeta)}</td><td class="num">${N.moeda(l.Meta)}</td><td class="num">${N.moeda(l.RealizadoMesesMeta)}</td>${celFora}<td class="num">${N.pct(l.Atingimento)}</td><td class="num">${l.QtdVendas}</td></tr>`;
  });
  const totFora = temFora ? `<td class="num">${N.moeda(e.ForaMesesMeta)}</td>` : '';

  const corpo = `<section class="filtros">
  <form method="get" action="/">
    <label>De <select name="de">${opcoesMeses(meses, d.MesInicial)}</select></label>
    <label>até <select name="ate">${opcoesMeses(meses, d.MesFinal)}</select></label>
    <label class="check"><input type="checkbox" name="desligados" value="1"${marcado}> Incluir desligados nos rankings</label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Período: <strong>${periodo}</strong> · data-base das vendas: ${N.fmtData(d.DataBase)} · fontes: ${fontes}</p>
</section>
${parcial}
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Realizado (vendas faturadas)</span><span class="valor">${N.moeda(e.Realizado)}</span><span class="det">${e.QtdVendas} vendas · ticket médio ${N.moeda(e.TicketMedio)}</span></div>
  <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">${N.moeda(e.Meta)}</span><span class="det">soma só dos meses em que cada vendedor tinha meta</span></div>
  <div class="cartao destaque"><span class="rotulo">Atingimento da empresa</span><span class="valor">${N.pct(e.Atingimento)}</span><span class="det">realizado ÷ meta (soma ÷ soma)</span></div>
  <div class="cartao neutro"><span class="rotulo">Canceladas (fora do realizado)</span><span class="valor">${N.moeda(d.ValorCanceladas)}</span><span class="det">${d.QtdCanceladas} vendas com Status = Cancelada</span></div>
</section>
${foraMeta}
<section class="rankings">
  <div class="bloco">
    <h2>Ranking por faturamento</h2>
    <table>
      <thead><tr><th>#</th><th>Vendedor</th><th>Filial</th><th class="num">Meses c/ meta</th><th class="num">Realizado</th><th class="num">% do total</th></tr></thead>
      <tbody>${linhasFat.join('\n')}</tbody>
    </table>
  </div>
  <div class="bloco">
    <h2>Ranking por atingimento</h2>
    <table>
      <thead><tr><th>#</th><th>Vendedor</th><th class="num">Meta</th><th class="num">Realizado</th><th class="num">Atingimento</th></tr></thead>
      <tbody>${linhasAt.join('\n')}</tbody>
    </table>
    ${notaSemMeta}
  </div>
</section>
${notaDesligados}
<section class="bloco">
  <div class="cab-bloco">
    <h2>Quadro de conferência por vendedor</h2>
    <a class="botao" href="/conferencia.csv?${q}">Baixar CSV</a>
  </div>
  <p class="nota">Todos os vendedores, inclusive desligados. A linha de total é a da empresa e deve bater com a soma das planilhas. Clique num nome para ver mês a mês e venda a venda.</p>
  <table>
    <thead><tr><th>ID</th><th>Vendedor</th><th>Filial</th><th>Meses c/ meta</th><th class="num">Meta</th><th class="num">${rotuloReal}</th>${cabFora}<th class="num">Atingimento</th><th class="num">Vendas</th></tr></thead>
    <tbody>${linhasQuadro.join('\n')}</tbody>
    <tfoot><tr><td></td><td>Total da empresa</td><td></td><td></td><td class="num">${N.moeda(e.Meta)}</td><td class="num">${N.moeda(e.RealizadoMesesMeta)}</td>${totFora}<td class="num">${N.pct(e.Atingimento)}</td><td class="num">${e.QtdVendas}</td></tr></tfoot>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  realizado = só vendas com Status = Faturada (§1) ·
  meta e realizado somados só nos meses em que o vendedor tem meta (§4.1) ·
  empresa = soma ÷ soma, com desligados (§4.2) ·
  rankings só com ativos, salvo filtro (§2.3) ·
  filial = filial do vendedor (§1).
</footer>`;
  return layout(`Desempenho comercial · ${periodo}`, corpo, usuario);
}

function paginaVendedor(base, d, detalhe, linha, usuario, incluirDesligados) {
  const v = linha.Vendedor;
  const q = queryPeriodo(d, incluirDesligados);
  const situacao = v.Ativo ? `Ativo desde ${N.fmtData(v.Admissao)}` : `Desligado em ${N.fmtData(v.Desligamento)} (admitido em ${N.fmtData(v.Admissao)})`;
  const linhasMes = detalhe.Meses.map((m) => {
    const meta = m.Meta === null ? '<span class="nota">sem meta</span>' : N.moeda(m.Meta);
    const canc = m.QtdCanceladas ? `${m.QtdCanceladas} · ${N.moeda(m.ValorCanceladas)}` : '—';
    return `<tr><td>${NOMES_MES_LONGO[m.Mes]}</td><td class="num">${meta}</td><td class="num">${N.moeda(m.Realizado)}</td><td class="num">${m.Meta !== null ? N.pct(m.Atingimento) : '—'}</td><td class="num">${m.QtdVendas}</td><td class="num">${canc}</td></tr>`;
  });
  const linhasVenda = detalhe.Vendas.map((s) => {
    const real = R.vendaRealizada(s);
    const cls = real ? '' : ' class="cancelada"';
    const status = real ? 'Faturada' : 'Cancelada <span class="marca desligado">fora do realizado</span>';
    return `<tr${cls}><td>${esc(s.Id)}</td><td class="num">${s.Linha}</td><td>${N.fmtData(s.Data)}</td><td>${esc(s.Cliente)}</td><td>${esc(s.Produto)}</td><td>${status}</td><td class="num">${N.moeda(s.Valor)}</td></tr>`;
  });

  const corpo = `<p><a href="/?${q}">← Voltar ao painel</a></p>
<section class="bloco">
  <h1>${esc(v.Nome)} <span class="id">${esc(v.Id)}</span> ${marcas(v)}</h1>
  <p class="fonte">${esc(v.Filial)} · ${situacao} · período: ${NOMES_MES_LONGO[d.MesInicial]} a ${NOMES_MES_LONGO[d.MesFinal]} de 2026</p>
  <section class="cartoes">
    <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">${N.moeda(linha.Meta)}</span><span class="det">meses com meta: ${formatarMeses(linha.MesesComMeta)}</span></div>
    <div class="cartao"><span class="rotulo">Realizado</span><span class="valor">${N.moeda(linha.RealizadoMesesMeta)}</span><span class="det">${linha.QtdVendas} vendas faturadas</span></div>
    <div class="cartao destaque"><span class="rotulo">Atingimento</span><span class="valor">${N.pct(linha.Atingimento)}</span><span class="det">realizado ÷ meta dos mesmos meses</span></div>
  </section>
</section>
<section class="bloco">
  <h2>Mês a mês</h2>
  <table>
    <thead><tr><th>Mês</th><th class="num">Meta</th><th class="num">Realizado</th><th class="num">Atingimento</th><th class="num">Vendas</th><th class="num">Canceladas</th></tr></thead>
    <tbody>${linhasMes.join('\n')}</tbody>
    <tfoot><tr><td>Período</td><td class="num">${N.moeda(linha.Meta)}</td><td class="num">${N.moeda(linha.Realizado)}</td><td class="num">${N.pct(linha.Atingimento)}</td><td class="num">${linha.QtdVendas}</td><td></td></tr></tfoot>
  </table>
</section>
<section class="bloco">
  <h2>Vendas do período</h2>
  <p class="nota">"Linha" é o número da linha em vendas_2026_jan-ago.xlsx, para achar a venda direto na planilha.</p>
  <table>
    <thead><tr><th>ID Venda</th><th class="num">Linha</th><th>Data</th><th>Cliente</th><th>Produto</th><th>Status</th><th class="num">Valor total</th></tr></thead>
    <tbody>${linhasVenda.join('\n')}</tbody>
  </table>
</section>`;
  return layout(`${v.Nome} · Desempenho`, corpo, usuario);
}

function csvConferencia(d) {
  const f = N.csvDinheiro, p = N.csvPct;
  const linhas = ['ID Vendedor;Vendedor;Filial;Status;Meses com meta;Meta;Realizado nos meses com meta;Realizado fora dos meses com meta;Realizado total;Atingimento (%);Vendas faturadas'];
  for (const l of N.ordenar(d.Linhas, (x) => x.Vendedor.Id)) {
    const v = l.Vendedor;
    linhas.push([v.Id, v.Nome, v.Filial, v.Status, formatarMeses(l.MesesComMeta), f(l.Meta), f(l.RealizadoMesesMeta), f(l.ForaMesesMeta), f(l.Realizado), p(l.Atingimento), l.QtdVendas].map((x) => x ?? '').join(';'));
  }
  const e = d.Empresa;
  linhas.push(`TOTAL;Empresa;;;;${f(e.Meta)};${f(e.RealizadoMesesMeta)};${f(e.ForaMesesMeta)};${f(e.Realizado)};${p(e.Atingimento)};${e.QtdVendas}`);
  linhas.push(`CANCELADAS (fora do realizado);;;;;;;;${f(d.ValorCanceladas)};;${d.QtdCanceladas}`);
  return linhas.join('\n') + '\n';
}

// O CSS tem uma fonte só: a função Get-Css de app/lib/Paginas.ps1.
let cssEmCache = null;
function css() {
  if (cssEmCache === null) {
    const ps = fs.readFileSync(path.join(__dirname, '..', '..', 'app', 'lib', 'Paginas.ps1'), 'utf8').replace(/\r\n/g, '\n');
    const m = /function Get-Css \{\n\s*return @'\n([\s\S]*?)\n'@/.exec(ps);
    if (!m) throw new Error('CSS não encontrado em app/lib/Paginas.ps1 (função Get-Css)');
    cssEmCache = m[1];
  }
  return cssEmCache;
}

module.exports = {
  NOMES_MES_LONGO, formatarMeses, marcas, queryPeriodo, opcoesMeses, layout, paginaLogin, paginaErroImportacao,
  paginaPainel, paginaVendedor, csvConferencia, css,
};
