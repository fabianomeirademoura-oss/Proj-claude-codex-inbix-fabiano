'use strict';
// Tela do histórico de alterações — espelho de app/lib/PaginasHistorico.ps1 (REGRAS_NEGOCIO.md §11.3).
// Só formata: as linhas vêm de R.lerAlteracoes. Textos idênticos aos do PowerShell (paridade).

const N = require('./numeros');
const R = require('./regras');
const P = require('./paginas');

const esc = N.esc;
const TIPOS_ALTERACAO = { vendedor: 'Vendedor', meta: 'Meta', produto: 'Produto', oportunidade: 'Oportunidade' };

function podeVerHistorico(usuario) { return N.igual(usuario.perfil, 'diretoria') || N.igual(usuario.perfil, 'gerente'); }   // §11.3

// Da alteração mais recente para a mais antiga (o arquivo só recebe acréscimos, em ordem).
function historicoFiltrado(alteracoes, tipo, registro) {
  const r = registro.toLowerCase();
  return [...alteracoes].reverse().filter((a) => (!tipo || a.Tipo === tipo) && (!registro || a.Registro.toLowerCase().includes(r)));
}

function valor(v) { return v === '' ? '—' : esc(v); }

function paginaHistorico(alteracoes, tipo, registro, usuario) {
  const linhas = historicoFiltrado(alteracoes, tipo, registro);
  const opcoes = '<option value="">todos</option>' + Object.keys(TIPOS_ALTERACAO).map((k) => `<option value="${k}"${k === tipo ? ' selected' : ''}>${TIPOS_ALTERACAO[k]}</option>`).join('');
  const q = `tipo=${N.escaparUri(tipo)}&amp;registro=${N.escaparUri(registro)}`;
  const corpoTabela = linhas.map((a) => {
    const t = Object.prototype.hasOwnProperty.call(TIPOS_ALTERACAO, a.Tipo) ? TIPOS_ALTERACAO[a.Tipo] : a.Tipo;
    return `<tr><td class="nw">${esc(a.Quando)}</td><td>${esc(a.Nome)}<span class="item id">${esc(a.Login)} · ${esc(a.Perfil)}</span></td><td>${esc(a['Ação'])}</td><td class="nw">${esc(t)} <code>${esc(a.Registro)}</code></td><td>${esc(a.Campo)}</td><td>${valor(a.Anterior)}</td><td>${valor(a.Novo)}</td><td class="nw"><code>${esc(a.Lote)}</code></td></tr>`;
  });
  const tabela = linhas.length
    ? `<table><thead><tr><th>Quando</th><th>Quem</th><th>Ação</th><th>Registro</th><th>Campo</th><th>Anterior</th><th>Novo</th><th>Lote</th></tr></thead><tbody>${corpoTabela.join('\n')}</tbody></table>`
    : alteracoes.length ? '<p class="nota">Nenhuma alteração com esse filtro.</p>' : '<p class="nota">Nenhuma alteração registrada até agora.</p>';
  const corpo = `<section class="filtros">
  <h1>Histórico de alterações</h1>
  <p class="fonte">Quem alterou o quê pelas ferramentas de escrita (REGRAS_NEGOCIO.md §11): uma linha por campo, com o valor anterior e o novo. Nada é apagado; desfazer é uma nova alteração, que também aparece aqui. As alterações do mesmo lote foram confirmadas juntas.</p>
  <form method="get" action="/historico">
    <label>Tipo <select name="tipo">${opcoes}</select></label>
    <label>Registro <input name="registro" type="text" value="${esc(registro)}" placeholder="ex.: V011, P066, OP-0130"></label>
    <button type="submit">Filtrar</button>
    <a href="/historico.csv?${q}">Baixar CSV</a>
  </form>
</section>
<section class="bloco">
  <p class="nota">${linhas.length} de ${alteracoes.length} alterações registradas, da mais recente para a mais antiga.</p>
  ${tabela}
</section>
<footer class="regras">REGRAS_NEGOCIO.md §11.1: nada é gravado sem confirmação · §11.2: só gerentes (na própria filial) e a diretoria alteram · §11.3: registro por campo, só com acréscimos · §11.4: as planilhas originais não mudam; o painel aplica o registro por cima delas.</footer>`;
  return P.layout('Histórico · Horizonte Máquinas', corpo, usuario, 'historico');
}

const aspas = (t) => '"' + String(t ?? '').replace(/"/g, '""') + '"';
function csvHistorico(alteracoes, tipo, registro) {
  let s = R.COLUNAS_ALTERACOES.map(aspas).join(';') + '\r\n';
  for (const a of historicoFiltrado(alteracoes, tipo, registro)) s += R.COLUNAS_ALTERACOES.map((c) => aspas(a[c])).join(';') + '\r\n';
  return s;
}

function paginaSemAcessoHistorico(usuario) {
  return P.layout('Sem acesso', '<section class="aviso-erro"><h1>Sem acesso</h1><p>Só a diretoria e os gerentes veem o histórico de alterações (REGRAS_NEGOCIO.md §11.3).</p></section>', usuario, 'historico');
}

module.exports = { TIPOS_ALTERACAO, podeVerHistorico, historicoFiltrado, paginaHistorico, csvHistorico, paginaSemAcessoHistorico };
