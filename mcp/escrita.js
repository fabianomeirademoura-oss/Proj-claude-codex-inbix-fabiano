'use strict';
// Ferramentas de escrita do servidor MCP (REGRAS_NEGOCIO.md §11).
//
// Cada ferramenta de escrita só PREPARA: valida, devolve o que vai mudar (atual → novo) e um código.
// Nada é gravado aí. Quem grava é confirmar_alteracao, com o código, depois do "sim" da pessoa (§11.1).
// A confirmação refaz tudo; se o resultado mudou, não grava e devolve o novo resumo.
//
// Gravar = acrescentar linhas em dados/importacoes/alteracoes.csv, uma por campo (§11.3). As planilhas
// originais não mudam; o painel e o MCP aplicam esse registro por cima delas (§11.4).

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const N = require('../web/lib/numeros');
const R = require('../web/lib/regras');
const F = require('./ferramentas');

const { ErroFerramenta, reais, data, mesTexto, vendedorResumo, semAcento, texto, inteiro, logico, resolverVendedor } = F;
const VALIDADE_MS = 15 * 60 * 1000;   // §11.1: o código vale 15 minutos
const FUSO = 'America/Sao_Paulo';
const PERFIS_ESCRITA = ['diretoria', 'gerente'];

// ---------- quem está usando (§11.2) ----------

// A conta vem da configuração do Claude Desktop (HORIZONTE_USUARIO) e é conferida em usuarios.json a
// cada pedido: se o perfil mudar ou a conta sumir, a escrita para na hora.
function criarIdentidade(login, arquivoUsuarios) {
  return () => {
    if (!login) return null;
    let lista;
    try {
      const v = JSON.parse(fs.readFileSync(arquivoUsuarios, 'utf8').replace(/^﻿/, ''));
      lista = Array.isArray(v) ? v : [v];
    } catch (e) {
      throw new ErroFerramenta(`Não consegui ler as contas do painel (${arquivoUsuarios}): ${e.message}`);
    }
    const u = lista.find((x) => String(x.login).toLowerCase() === String(login).toLowerCase());
    if (!u) throw new ErroFerramenta(`A conta '${login}', configurada no Claude Desktop, não existe em ${path.basename(arquivoUsuarios)}.`);
    return { login: u.login, nome: u.nome, perfil: String(u.perfil || '').toLowerCase(), idFilial: u.idFilial || null, idVendedor: u.idVendedor || null };
  };
}

function exigirEscrita(ctx) {
  const u = ctx.usuario();
  if (!u) throw new ErroFerramenta('Nenhuma conta configurada para escrever. Defina HORIZONTE_USUARIO no Claude Desktop (docs/MCP.md). A leitura continua livre.');
  if (!PERFIS_ESCRITA.includes(String(u.perfil).toLowerCase())) {
    throw new ErroFerramenta(`A conta '${u.login}' tem perfil '${u.perfil}' e só pode ler. Escrever exige perfil gerente ou diretoria (REGRAS_NEGOCIO.md §11.2).`);
  }
  if (u.perfil === 'gerente' && !u.idFilial) throw new ErroFerramenta(`A conta de gerente '${u.login}' não tem filial em usuarios.json.`);
  return u;
}

// §11.2: diretoria altera tudo; gerente, só vendedores da própria filial (pela filial do cadastro).
function exigirFilial(u, vend, oque) {
  if (u.perfil === 'diretoria' || vend.IdFilial === u.idFilial) return;
  throw new ErroFerramenta(`${oque}: ${vend.Id} ${vend.Nome} é da filial ${vend.Filial} (${vend.IdFilial}), e a conta '${u.login}' é gerente da filial ${u.idFilial}. Gerente só altera a própria filial (REGRAS_NEGOCIO.md §11.2).`);
}

// ---------- valores ----------

// Dinheiro digitado: número, "455000.50" ou "455.000,50". Até 2 casas (§0.4). Devolve centavos.
function lerDinheiro(v, nome) {
  if (v === undefined || v === null || v === '') return null;
  let s = typeof v === 'number' ? String(v) : String(v).trim().replace(/^R\$\s*/i, '');
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  const d = N.lerDecimal(s);
  if (d === null) throw new ErroFerramenta(`'${nome}' não é um valor em reais: '${v}'.`);
  let { n, escala } = d;
  while (escala > 2 && n % 10n === 0n) { n /= 10n; escala--; }
  if (escala > 2) throw new ErroFerramenta(`'${nome}' tem mais de 2 casas decimais: '${v}'.`);
  return Number(N.paraEscala(d, 2));
}
function decimalTexto(c) { return c === null || c === undefined ? '' : N.formatarEscalado(c, 2, false).replace(',', '.'); }
function moeda(c) { return c === null || c === undefined ? 'sem meta' : N.moeda(c); }

function agora() {
  const p = {};
  for (const x of new Intl.DateTimeFormat('en-GB', { timeZone: FUSO, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(new Date())) p[x.type] = x.value;
  return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
}

// ---------- preparar: cada ação devolve { acao, titulo, mudancas, efeitos } ----------
// mudança = { Tipo, Registro, Campo, Anterior, Novo, descricao } — as quatro primeiras vão para o registro.

function mudanca(tipo, registro, campo, anterior, novo, descricao) {
  return { Tipo: tipo, Registro: registro, Campo: campo, Anterior: anterior ?? '', Novo: novo ?? '', descricao };
}

function vendedorPorArg(fonte, args, nome) {
  const busca = texto(args, nome);
  if (!busca) throw new ErroFerramenta(`Informe '${nome}' (ID, como V003, ou nome).`);
  return resolverVendedor(fonte.comercial(), busca);
}

function oportunidadesAbertas(fonte, idVendedor) {
  return N.ordenar(fonte.pipeline().Oportunidades.filter((o) => o.Aberta && o.IdVendedor === idVendedor), (o) => o.Id);
}

function inativarVendedor(fonte, args, u) {
  const v = vendedorPorArg(fonte, args, 'vendedor');
  exigirFilial(u, v, 'Inativar');
  if (!v.Ativo) throw new ErroFerramenta(`${v.Id} ${v.Nome} já está inativo (desligado em ${N.fmtData(v.Desligamento)}).`);
  const textoData = texto(args, 'data_desligamento');
  if (!textoData) throw new ErroFerramenta("Informe 'data_desligamento' (AAAA-MM-DD ou DD/MM/AAAA). Ela é obrigatória (REGRAS_NEGOCIO.md §11.5).");
  const dia = N.lerData(textoData);
  if (dia === null) throw new ErroFerramenta(`Data de desligamento inválida: '${textoData}'. Use AAAA-MM-DD ou DD/MM/AAAA.`);

  // §11.5 e §2.6: não pode ser antes da admissão, da última venda nem da criação da última oportunidade.
  const base = fonte.comercial();
  const limites = [];
  if (v.Admissao !== null) limites.push(['à admissão', v.Admissao]);
  const vendas = base.Vendas.filter((x) => x.IdVendedor === v.Id);
  if (vendas.length) limites.push(['à última venda', Math.max(...vendas.map((x) => x.Data))]);
  const ops = fonte.pipeline().Oportunidades.filter((o) => o.IdVendedor === v.Id);
  if (ops.length) limites.push(['à criação da última oportunidade', Math.max(...ops.map((o) => o.Criacao))]);
  for (const [oque, d] of limites) {
    if (dia < d) throw new ErroFerramenta(`A data de desligamento (${N.fmtData(dia)}) não pode ser anterior ${oque} de ${v.Nome} (${N.fmtData(d)}) (REGRAS_NEGOCIO.md §2.6 e §11.5).`);
  }

  const mudancas = [
    mudanca('vendedor', v.Id, 'Status', 'Ativo', 'Inativo', `Status de ${v.Id} ${v.Nome}`),
    mudanca('vendedor', v.Id, 'Data de Desligamento', '', N.fmtDataIso(dia), `Data de desligamento de ${v.Id}`),
  ];
  // §11.5: metas dos meses depois do mês do desligamento passam a "sem meta".
  const anoD = N.anoDe(dia), mesD = N.mesDe(dia);
  const futuras = N.ordenar(base.Metas.filter((m) => m.IdVendedor === v.Id && (m.Ano > anoD || (m.Ano === anoD && m.Mes > mesD))), (m) => m.Ano, (m) => m.Mes);
  for (const m of futuras) {
    const reg = `${v.Id}/${m.Ano}-${String(m.Mes).padStart(2, '0')}`;
    mudancas.push(mudanca('meta', reg, 'Meta (R$)', decimalTexto(m.Valor), '', `Meta de ${v.Id} em ${String(m.Mes).padStart(2, '0')}/${m.Ano}: ${moeda(m.Valor)} → sem meta`));
  }
  const abertas = oportunidadesAbertas(fonte, v.Id);
  let valorAbertas = 0;
  for (const o of abertas) valorAbertas += o.Valor;
  const efeitos = [
    `${v.Nome} perde o acesso ao painel (§2.2).`,
    futuras.length ? `${futuras.length} meta(s) depois de ${String(mesD).padStart(2, '0')}/${anoD} passam a "sem meta" e saem da meta da filial e da empresa (§11.5, §4.2).` : 'Nenhuma meta depois do mês do desligamento.',
    abertas.length ? `${abertas.length} oportunidade(s) aberta(s) viram órfãs (${abertas.map((o) => o.Id).join(', ')}; ${N.moeda(valorAbertas)}). Elas não são transferidas sozinhas: use transferir_oportunidades (§2.4).` : 'Nenhuma oportunidade aberta fica órfã.',
    'O histórico de vendas e metas até o desligamento continua valendo (§2.1).',
  ];
  return { acao: 'inativar_vendedor', titulo: `Inativar ${v.Id} ${v.Nome} em ${N.fmtData(dia)}`, mudancas, efeitos };
}

function reativarVendedor(fonte, args, u, ctx) {
  const v = vendedorPorArg(fonte, args, 'vendedor');
  exigirFilial(u, v, 'Reativar');
  if (v.Ativo) throw new ErroFerramenta(`${v.Id} ${v.Nome} já está ativo.`);
  const base = fonte.comercial();
  const mudancas = [
    mudanca('vendedor', v.Id, 'Status', 'Inativo', 'Ativo', `Status de ${v.Id} ${v.Nome}`),
    mudanca('vendedor', v.Id, 'Data de Desligamento', N.fmtDataIso(v.Desligamento), '', `Data de desligamento de ${v.Id} (fica no histórico)`),
  ];
  const semMeta = [];
  for (let m = 1; m <= 12; m++) if (!base.Metas.some((x) => x.IdVendedor === v.Id && x.Ano === R.ANO_METAS && x.Mes === m)) semMeta.push(m);
  const orfas = oportunidadesAbertas(fonte, v.Id);
  const temConta = ctx.contaDoVendedor(v.Id);
  const efeitos = [
    semMeta.length ? `As metas não voltam sozinhas: ${v.Nome} continua "sem meta" em ${semMeta.map((m) => mesTexto(m)).join(', ')}. Use alterar_meta (§11.6).` : 'Todos os meses de 2026 já têm meta.',
    orfas.length ? `${orfas.length} oportunidade(s) aberta(s) deixam de ser órfãs (${orfas.map((o) => o.Id).join(', ')}).` : 'Nenhuma oportunidade órfã dele.',
    temConta ? 'A conta dele no painel volta a entrar.' : `Ele ainda não tem conta no painel: depois de confirmar, rode app/criar_usuarios.ps1 -Completar (§11.6).`,
  ];
  return { acao: 'reativar_vendedor', titulo: `Reativar ${v.Id} ${v.Nome}`, mudancas, efeitos };
}

function alterarMeta(fonte, args, u) {
  const v = vendedorPorArg(fonte, args, 'vendedor');
  exigirFilial(u, v, 'Alterar meta');
  const mes = inteiro(args, 'mes', 1, 12);
  if (mes === null) throw new ErroFerramenta("Informe 'mes' (1 a 12).");
  const ano = inteiro(args, 'ano', 1900, 9999);
  if (ano !== null && ano !== R.ANO_METAS) throw new ErroFerramenta(`Só há metas de ${R.ANO_METAS}.`);
  const base = fonte.comercial();
  // §11.7: só meses abertos — do mês da data-base de vendas em diante.
  const primeiroAberto = N.anoDe(base.DataBase) < R.ANO_METAS ? 1 : N.anoDe(base.DataBase) > R.ANO_METAS ? 13 : N.mesDe(base.DataBase);
  if (mes < primeiroAberto) {
    throw new ErroFerramenta(`${mesTexto(mes)} já fechou: a data-base de vendas é ${N.fmtData(base.DataBase)}, e só os meses de ${mesTexto(primeiroAberto)} a 12/${R.ANO_METAS} podem mudar (REGRAS_NEGOCIO.md §11.7).`);
  }
  if (v.Desligamento !== null && (N.anoDe(v.Desligamento) < R.ANO_METAS || (N.anoDe(v.Desligamento) === R.ANO_METAS && mes > N.mesDe(v.Desligamento)))) {
    throw new ErroFerramenta(`${v.Nome} foi desligado em ${N.fmtData(v.Desligamento)}: não pode ter meta em ${mesTexto(mes)} (REGRAS_NEGOCIO.md §11.7).`);
  }
  const semMeta = logico(args, 'sem_meta') === true;
  const valor = lerDinheiro(args.valor, 'valor');
  if (semMeta && valor !== null) throw new ErroFerramenta("Informe 'valor' ou 'sem_meta', não os dois.");
  if (!semMeta && valor === null) throw new ErroFerramenta("Informe 'valor' (a meta nova, em reais) ou 'sem_meta: true' para tirar a meta do mês.");
  if (!semMeta && valor <= 0) throw new ErroFerramenta('A meta tem de ser maior que zero. Para tirar a meta do mês, use sem_meta (REGRAS_NEGOCIO.md §3.1 e §11.7).');
  const atual = base.Metas.find((m) => m.IdVendedor === v.Id && m.Ano === R.ANO_METAS && m.Mes === mes);
  const valorAtual = atual ? atual.Valor : null;
  const novo = semMeta ? null : valor;
  if (valorAtual === novo) throw new ErroFerramenta(`Nada a alterar: a meta de ${v.Nome} em ${mesTexto(mes)} já é ${moeda(valorAtual)}.`);
  const reg = `${v.Id}/${R.ANO_METAS}-${String(mes).padStart(2, '0')}`;
  const mudancas = [mudanca('meta', reg, 'Meta (R$)', decimalTexto(valorAtual), decimalTexto(novo), `Meta de ${v.Id} ${v.Nome} em ${mesTexto(mes)}: ${moeda(valorAtual)} → ${moeda(novo)}`)];
  const efeitos = [
    valorAtual === null ? 'O mês não tinha meta: passa a ter (§11.7).' : novo === null ? 'O mês passa a "sem meta": sai do atingimento dele e da meta da filial e da empresa (§3.1, §4.2).' : 'A meta da filial e da empresa no mês muda na mesma diferença (§4.2).',
  ];
  return { acao: 'alterar_meta', titulo: `Meta de ${v.Id} ${v.Nome} em ${mesTexto(mes)}`, mudancas, efeitos };
}

const CAMPOS_PRODUTO = [
  ['nome', 'Produto'], ['categoria', 'Categoria'], ['marca', 'Marca'], ['modelo', 'Modelo'],
  ['preco_tabela', 'Preço de Tabela'], ['custo_unitario', 'Custo Unitário'], ['unidade', 'Unidade'], ['status', 'Status'],
];

function cadastrarOuEditarProduto(fonte, args) {
  const cadastro = fonte.catalogo();
  const categorias = N.unicosOrdenados(cadastro.map((r) => r.Categoria).filter(Boolean));
  const id = texto(args, 'produto_id');
  const existente = id ? cadastro.find((r) => N.igual(r['ID Produto'], id)) : null;
  if (id && !existente) throw new ErroFerramenta(`O produto '${id}' não existe. Para cadastrar um produto novo, não informe produto_id: o código é o próximo livre (REGRAS_NEGOCIO.md §11.8).`);

  // Valores pedidos, já no formato da planilha.
  const pedidos = {};
  for (const [param, coluna] of CAMPOS_PRODUTO) {
    if (args[param] === undefined || args[param] === null) continue;
    if (param === 'preco_tabela' || param === 'custo_unitario') {
      const c = lerDinheiro(args[param], param);
      if (c === null) continue;
      if (c < 0) throw new ErroFerramenta(`'${param}' não pode ser negativo (REGRAS_NEGOCIO.md §11.8).`);
      pedidos[coluna] = decimalTexto(c);
      continue;
    }
    const t = String(args[param]).replace(/[\r\n]+/g, ' ').trim();
    if (param === 'categoria') {
      const cat = categorias.find((c) => semAcento(c) === semAcento(t));
      if (!cat) throw new ErroFerramenta(`Categoria '${t}' não existe no catálogo. Categoria nova é combinada antes com o coordenador (REGRAS_NEGOCIO.md §11.8).`, { categorias });
      pedidos[coluna] = cat;
    } else if (param === 'status') {
      const st = ['Ativo', 'Descontinuado'].find((x) => N.igual(x, t));
      if (!st) throw new ErroFerramenta(`Status '${t}' inválido: use Ativo ou Descontinuado (§5.5).`);
      pedidos[coluna] = st;
    } else if (param === 'nome' && !t) {
      throw new ErroFerramenta('O nome do produto não pode ficar vazio.');
    } else {
      pedidos[coluna] = t;
    }
  }

  const mudancas = [];
  const efeitos = [];
  let registro, titulo;
  if (existente) {
    registro = existente['ID Produto'];
    const mesmoValor = (coluna, a, b) => (coluna === 'Preço de Tabela' || coluna === 'Custo Unitário') ? N.dinheiro(a) === N.dinheiro(b) : String(a ?? '') === String(b ?? '');
    for (const [, coluna] of CAMPOS_PRODUTO) {
      if (!(coluna in pedidos)) continue;
      const atual = existente[coluna] ?? '';
      if (mesmoValor(coluna, atual, pedidos[coluna])) continue;
      mudancas.push(mudanca('produto', registro, coluna, String(atual), pedidos[coluna], `${coluna} de ${registro}: ${atual === '' ? '(vazio)' : atual} → ${pedidos[coluna] === '' ? '(vazio)' : pedidos[coluna]}`));
    }
    if (!mudancas.length) throw new ErroFerramenta(`Nada a alterar em ${registro} ${existente.Produto}: informe ao menos um campo com valor diferente.`);
    titulo = `Editar ${registro} ${existente.Produto}`;
    if ('Categoria' in pedidos && mudancas.some((m) => m.Campo === 'Categoria')) efeitos.push(`O estoque e o funil passam a contar ${registro} em ${pedidos.Categoria} (§5, §8.4).`);
    if (mudancas.some((m) => m.Campo === 'Status' && m.Novo === 'Descontinuado')) efeitos.push('Descontinuado sai do catálogo de venda, mas continua no estoque e na lista de parados (§5.5).');
  } else {
    for (const obrigatorio of ['Produto', 'Categoria', 'Preço de Tabela', 'Unidade']) {
      if (!(obrigatorio in pedidos) || pedidos[obrigatorio] === '') {
        throw new ErroFerramenta(`Produto novo exige nome, categoria, preco_tabela e unidade; falta '${obrigatorio}' (REGRAS_NEGOCIO.md §11.8).`, { categorias });
      }
    }
    if (!('Status' in pedidos)) pedidos.Status = 'Ativo';
    let maior = 0;
    for (const r of cadastro) { const m = /^P(\d+)$/i.exec(r['ID Produto'] || ''); if (m) maior = Math.max(maior, parseInt(m[1], 10)); }
    registro = `P${String(maior + 1).padStart(3, '0')}`;
    for (const [, coluna] of CAMPOS_PRODUTO) {
      if (pedidos[coluna] === undefined || pedidos[coluna] === '') continue;
      mudancas.push(mudanca('produto', registro, coluna, '', pedidos[coluna], `${coluna} de ${registro}: ${pedidos[coluna]}`));
    }
    titulo = `Cadastrar ${registro} ${pedidos.Produto}`;
    efeitos.push(`O produto nasce com o código ${registro} e sem linha de estoque: aparece em "sem estoque" até a próxima foto (§5, §11.8).`);
    const parecido = cadastro.find((r) => semAcento(r.Produto || '') === semAcento(pedidos.Produto));
    if (parecido) efeitos.push(`Atenção: já existe ${parecido['ID Produto']} com o mesmo nome (${parecido.Produto}).`);
  }
  return { acao: 'cadastrar_ou_editar_produto', titulo, mudancas, efeitos };
}

function transferirOportunidades(fonte, args, u) {
  const de = vendedorPorArg(fonte, args, 'de_vendedor');
  const para = vendedorPorArg(fonte, args, 'para_vendedor');
  exigirFilial(u, de, 'Transferir (dono atual)');
  exigirFilial(u, para, 'Transferir (novo dono)');
  if (de.Id === para.Id) throw new ErroFerramenta('O novo dono tem de ser diferente do dono atual.');
  if (!para.Ativo) throw new ErroFerramenta(`${para.Id} ${para.Nome} está desligado: o novo dono tem de ser um vendedor ativo (REGRAS_NEGOCIO.md §11.9).`);
  const pipe = fonte.pipeline();
  let escolhidas;
  const lista = Array.isArray(args.oportunidades) ? args.oportunidades.map(String) : texto(args, 'oportunidades') ? texto(args, 'oportunidades').split(/[\s,;]+/) : [];
  const ids = lista.map((x) => x.trim()).filter(Boolean);
  if (ids.length) {
    const problemas = [];
    escolhidas = [];
    for (const id of [...new Set(ids.map((x) => x.toUpperCase()))]) {
      const o = pipe.Oportunidades.find((x) => N.igual(x.Id, id));
      if (!o) problemas.push(`${id}: não existe`);
      else if (o.IdVendedor !== de.Id) problemas.push(`${o.Id}: o dono atual é ${o.IdVendedor}, não ${de.Id}`);
      else if (!o.Aberta) problemas.push(`${o.Id}: está ${o.Etapa} (só oportunidades abertas são transferidas, §11.9)`);
      else escolhidas.push(o);
    }
    if (problemas.length) throw new ErroFerramenta('Algumas oportunidades não podem ser transferidas; nada foi preparado.', { problemas });
  } else {
    escolhidas = oportunidadesAbertas(fonte, de.Id);
    if (!escolhidas.length) throw new ErroFerramenta(`${de.Id} ${de.Nome} não tem oportunidades abertas.`);
  }
  escolhidas = N.ordenar(escolhidas, (o) => o.Id);
  let valor = 0;
  const mudancas = escolhidas.map((o) => {
    valor += o.Valor;
    return mudanca('oportunidade', o.Id, 'ID Vendedor', de.Id, para.Id, `Dono de ${o.Id} (${o.Produto}, ${N.moeda(o.Valor)}, ${o.Etapa}): ${de.Id} ${de.Nome} → ${para.Id} ${para.Nome}`);
  });
  const efeitos = [`${escolhidas.length} oportunidade(s), ${N.moeda(valor)} em valor estimado, passam para ${para.Nome} no pipeline e no funil (§7).`];
  if (de.IdFilial !== para.IdFilial) efeitos.push(`Elas mudam de filial nas visões por filial: de ${de.Filial} para ${para.Filial} (§8.4).`);
  if (!de.Ativo) efeitos.push(`Deixam de ser órfãs de ${de.Nome}, desligado em ${N.fmtData(de.Desligamento)} (§2.4).`);
  return { acao: 'transferir_oportunidades', titulo: `Transferir ${escolhidas.length} oportunidade(s) de ${de.Id} para ${para.Id}`, mudancas, efeitos };
}

// ---------- preparar → confirmar ----------

const ACOES = {
  inativar_vendedor: inativarVendedor,
  reativar_vendedor: reativarVendedor,
  alterar_meta: alterarMeta,
  cadastrar_ou_editar_produto: cadastrarOuEditarProduto,
  transferir_oportunidades: transferirOportunidades,
};

// Monta o resumo, conferindo que a base continua válida depois da alteração (§11.1).
function montar(fonte, acao, args, ctx) {
  const u = exigirEscrita(ctx);
  const plano = ACOES[acao](fonte, args, u, ctx);
  const validacao = fonte.validar(plano.mudancas.map((m) => ({ Lote: 'pendente', Tipo: m.Tipo, Registro: m.Registro, Campo: m.Campo, Anterior: m.Anterior, Novo: m.Novo })));
  if (validacao.erros.length) {
    throw new ErroFerramenta('A alteração deixaria as planilhas inválidas; nada foi preparado (REGRAS_NEGOCIO.md §11.1).', { erros: validacao.erros.slice(0, 20) });
  }
  return { usuario: u, plano };
}

function assinaturaPlano(plano) { return JSON.stringify(plano.mudancas.map((m) => [m.Tipo, m.Registro, m.Campo, m.Anterior, m.Novo])); }

function resumo(plano, u, codigo, validoAte) {
  return {
    situacao: 'AGUARDANDO CONFIRMAÇÃO — nada foi gravado',
    acao: plano.acao,
    titulo: plano.titulo,
    quem: { login: u.login, nome: u.nome, perfil: u.perfil, filial: u.idFilial },
    mudancas: plano.mudancas.map((m) => ({ tipo: m.Tipo, registro: m.Registro, campo: m.Campo, atual: m.Anterior === '' ? null : m.Anterior, novo: m.Novo === '' ? null : m.Novo, descricao: m.descricao })),
    efeitos: plano.efeitos,
    codigo_confirmacao: codigo,
    valido_ate: validoAte,
    como_confirmar: `Mostre este resumo à pessoa e pergunte se ela confirma. Só depois de um "sim" explícito chame confirmar_alteracao com o código ${codigo}. Se ela não confirmar, não faça nada: o código expira sozinho.`,
  };
}

function preparar(acao) {
  return (fonte, args, ctx) => {
    const { usuario, plano } = montar(fonte, acao, args, ctx);
    const codigo = crypto.randomBytes(4).toString('hex').toUpperCase();
    const expira = Date.now() + VALIDADE_MS;
    ctx.pendentes.set(codigo, { acao, args, login: usuario.login, assinatura: assinaturaPlano(plano), expira });
    return resumo(plano, usuario, codigo, new Date(expira).toISOString());
  };
}

function gravar(fonte, plano, u, lote) {
  const arq = R.arquivoAlteracoes(fonte.dirImportacoes);
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  const aspas = (t) => '"' + String(t ?? '').replace(/[\r\n]+/g, ' ').replace(/"/g, '""') + '"';
  const quando = agora();
  let texto = '';
  for (const m of plano.mudancas) {
    const linha = { Quando: quando, Lote: lote, Login: u.login, Nome: u.nome, Perfil: u.perfil, 'Ação': plano.acao, Tipo: m.Tipo, Registro: m.Registro, Campo: m.Campo, Anterior: m.Anterior, Novo: m.Novo };
    texto += R.COLUNAS_ALTERACOES.map((c) => aspas(linha[c])).join(';') + '\r\n';
  }
  // §11.3: só acréscimos. O cabeçalho (com BOM, para o Excel) só na criação do arquivo.
  if (!fs.existsSync(arq)) texto = '﻿' + R.COLUNAS_ALTERACOES.map(aspas).join(';') + '\r\n' + texto;
  fs.appendFileSync(arq, texto, 'utf8');
  return quando;
}

async function confirmarAlteracao(fonte, args, ctx) {
  const codigo = (texto(args, 'codigo') || '').toUpperCase();
  if (!codigo) throw new ErroFerramenta("Informe 'codigo', o código de confirmação devolvido pela ferramenta que preparou a alteração.");
  const p = ctx.pendentes.get(codigo);
  if (!p) throw new ErroFerramenta(`Código ${codigo} desconhecido ou já usado. Prepare a alteração de novo.`);
  if (p.expira < Date.now()) { ctx.pendentes.delete(codigo); throw new ErroFerramenta(`O código ${codigo} expirou (vale 15 minutos). Prepare a alteração de novo.`); }
  const u0 = exigirEscrita(ctx);
  if (u0.login.toLowerCase() !== p.login.toLowerCase()) throw new ErroFerramenta(`O código ${codigo} foi pedido por outra conta (${p.login}).`);

  // §11.1: a confirmação refaz tudo. Se a alteração mudou, não grava e devolve o resumo novo.
  ctx.pendentes.delete(codigo);
  const { usuario, plano } = montar(fonte, p.acao, p.args, ctx);
  if (assinaturaPlano(plano) !== p.assinatura) {
    const novo = preparar(p.acao)(fonte, p.args, ctx);
    throw new ErroFerramenta('Os dados mudaram desde o resumo e a alteração ficou diferente. Nada foi gravado. Mostre o resumo novo e peça outra confirmação.', { resumo_novo: novo });
  }

  // Quando o cliente sabe perguntar direto à pessoa (elicitation), a confirmação sai dela, não do modelo.
  if (ctx.perguntar) {
    const linhas = plano.mudancas.slice(0, 12).map((m) => `• ${m.descricao}`);
    if (plano.mudancas.length > 12) linhas.push(`• … e mais ${plano.mudancas.length - 12}`);
    const resposta = await ctx.perguntar(`${plano.titulo}\n${linhas.join('\n')}\n\nConfirma a gravação?`);
    if (resposta !== true) {
      return { situacao: 'NÃO GRAVADO — a pessoa não confirmou', acao: plano.acao, titulo: plano.titulo };
    }
  }

  const quando = gravar(fonte, plano, usuario, codigo);
  return {
    situacao: 'GRAVADO',
    acao: plano.acao,
    titulo: plano.titulo,
    quando,
    quem: { login: usuario.login, nome: usuario.nome, perfil: usuario.perfil },
    lote: codigo,
    campos_gravados: plano.mudancas.length,
    mudancas: plano.mudancas.map((m) => ({ tipo: m.Tipo, registro: m.Registro, campo: m.Campo, anterior: m.Anterior === '' ? null : m.Anterior, novo: m.Novo === '' ? null : m.Novo })),
    efeitos: plano.efeitos,
    historico: `Registrado em ${path.join('dados', 'importacoes', 'alteracoes.csv')}; aparece na tela Histórico do painel (/historico) e em ver_historico.`,
  };
}

function verHistorico(fonte, args) {
  const tipo = texto(args, 'tipo');
  const registro = texto(args, 'registro');
  const limite = inteiro(args, 'limite', 1, F.LIMITE_MAXIMO) ?? F.LIMITE_PADRAO;
  const tipos = ['vendedor', 'meta', 'produto', 'oportunidade'];
  if (tipo && !tipos.includes(tipo.toLowerCase())) throw new ErroFerramenta(`Tipo desconhecido: '${tipo}'.`, { tipos });
  const todas = R.lerAlteracoes(fonte.dirImportacoes);
  const filtradas = [...todas].reverse().filter((a) => (!tipo || a.Tipo === tipo.toLowerCase()) && (!registro || a.Registro.toLowerCase().includes(registro.toLowerCase())));
  return {
    total_no_registro: todas.length,
    encontradas: filtradas.length,
    mostradas: Math.min(limite, filtradas.length),
    alteracoes: filtradas.slice(0, limite).map((a) => ({
      quando: a.Quando, lote: a.Lote, quem: { login: a.Login, nome: a.Nome, perfil: a.Perfil }, acao: a['Ação'],
      tipo: a.Tipo, registro: a.Registro, campo: a.Campo, anterior: a.Anterior === '' ? null : a.Anterior, novo: a.Novo === '' ? null : a.Novo,
    })),
    tela: 'No painel: menu Histórico (/historico), para diretoria e gerentes.',
  };
}

// ---------- catálogo ----------

const txt = (description) => ({ type: 'string', description });
const int = (description, minimum, maximum) => ({ type: 'integer', description, minimum, maximum });
const PREPARA = 'Só prepara: devolve o que vai mudar (atual → novo) e um código de confirmação. NADA é gravado até confirmar_alteracao, que só pode ser chamada depois de a pessoa dizer "sim". Exige conta gerente (própria filial) ou diretoria.';
const soPrepara = { readOnlyHint: true, destructiveHint: false, idempotentHint: false, openWorldHint: false };

const FERRAMENTAS_ESCRITA = [
  {
    name: 'inativar_vendedor',
    title: 'Inativar vendedor (prepara)',
    description: `Prepara o desligamento de um vendedor: status Inativo, data de desligamento, metas dos meses seguintes viram "sem meta" e as oportunidades abertas ficam órfãs. ${PREPARA}`,
    inputSchema: { type: 'object', properties: { vendedor: txt('ID (ex.: V004) ou nome.'), data_desligamento: txt('Data do desligamento: AAAA-MM-DD ou DD/MM/AAAA.') }, required: ['vendedor', 'data_desligamento'] },
    annotations: soPrepara,
    executar: preparar('inativar_vendedor'),
  },
  {
    name: 'reativar_vendedor',
    title: 'Reativar vendedor (prepara)',
    description: `Prepara a volta de um vendedor desligado: status Ativo e data de desligamento em branco. As metas não voltam sozinhas. ${PREPARA}`,
    inputSchema: { type: 'object', properties: { vendedor: txt('ID (ex.: V011) ou nome.') }, required: ['vendedor'] },
    annotations: soPrepara,
    executar: preparar('reativar_vendedor'),
  },
  {
    name: 'alterar_meta',
    title: 'Alterar meta (prepara)',
    description: `Prepara a mudança da meta de um vendedor num mês aberto de 2026 (do mês da data-base de vendas até dezembro): valor novo maior que zero, ou sem_meta para tirar a meta do mês. ${PREPARA}`,
    inputSchema: {
      type: 'object',
      properties: {
        vendedor: txt('ID (ex.: V012) ou nome.'),
        mes: int('Mês de 1 a 12 (só meses abertos).', 1, 12),
        valor: { type: ['number', 'string'], description: 'Meta nova em reais (ex.: 455000 ou "455.000,50"). Até 2 casas.' },
        sem_meta: { type: 'boolean', description: 'true para tirar a meta do mês (fica "sem meta", não zero).' },
      },
      required: ['vendedor', 'mes'],
    },
    annotations: soPrepara,
    executar: preparar('alterar_meta'),
  },
  {
    name: 'cadastrar_ou_editar_produto',
    title: 'Cadastrar ou editar produto (prepara)',
    description: `Sem produto_id: prepara um produto novo (código = próximo livre; exige nome, categoria, preco_tabela e unidade). Com produto_id: prepara a edição só dos campos informados. A categoria tem de existir. ${PREPARA}`,
    inputSchema: {
      type: 'object',
      properties: {
        produto_id: txt('Código do produto a editar (ex.: P008). Vazio = produto novo.'),
        nome: txt('Nome do produto.'),
        categoria: txt('Categoria existente (ex.: Tratores).'),
        marca: txt('Marca.'),
        modelo: txt('Modelo.'),
        preco_tabela: { type: ['number', 'string'], description: 'Preço de tabela em reais, ≥ 0.' },
        custo_unitario: { type: ['number', 'string'], description: 'Custo unitário em reais, ≥ 0.' },
        unidade: txt('Unidade (ex.: UN).'),
        status: txt('Ativo ou Descontinuado.'),
      },
    },
    annotations: soPrepara,
    executar: preparar('cadastrar_ou_editar_produto'),
  },
  {
    name: 'transferir_oportunidades',
    title: 'Transferir oportunidades (prepara)',
    description: `Prepara a troca do dono de oportunidades ABERTAS para um vendedor ativo. Sem lista, transfere todas as abertas do dono atual (ex.: as órfãs de um desligado). ${PREPARA}`,
    inputSchema: {
      type: 'object',
      properties: {
        de_vendedor: txt('Dono atual: ID ou nome.'),
        para_vendedor: txt('Novo dono (ativo): ID ou nome.'),
        oportunidades: txt('IDs separados por vírgula (ex.: "OP-0130, OP-0140"). Vazio = todas as abertas do dono atual.'),
      },
      required: ['de_vendedor', 'para_vendedor'],
    },
    annotations: soPrepara,
    executar: preparar('transferir_oportunidades'),
  },
  {
    name: 'confirmar_alteracao',
    title: 'Confirmar e gravar alteração',
    description: 'GRAVA uma alteração preparada antes. Chame SOMENTE depois de mostrar o resumo à pessoa e ela responder "sim" explicitamente; nunca por conta própria. Refaz todas as validações; se algo mudou, não grava e devolve o resumo novo. Registra quem, quando, o registro, o valor anterior e o novo.',
    inputSchema: { type: 'object', properties: { codigo: txt('Código de confirmação devolvido pela ferramenta que preparou a alteração.') }, required: ['codigo'] },
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    executar: confirmarAlteracao,
  },
  {
    name: 'ver_historico',
    title: 'Ver histórico de alterações',
    description: 'Histórico das alterações gravadas (da mais recente para a mais antiga): quando, quem, ação, registro, campo, valor anterior e novo. Filtre por tipo (vendedor, meta, produto, oportunidade) e/ou registro (ex.: V011, P066).',
    inputSchema: {
      type: 'object',
      properties: {
        tipo: txt('vendedor, meta, produto ou oportunidade.'),
        registro: txt('Parte do ID do registro (ex.: V011, V003/2026-10, OP-0130).'),
        limite: int(`Máximo de linhas (padrão ${F.LIMITE_PADRAO}).`, 1, F.LIMITE_MAXIMO),
      },
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    executar: verHistorico,
  },
].map((f) => ({ ...f, annotations: { title: f.title, ...f.annotations } }));

module.exports = { FERRAMENTAS_ESCRITA, criarIdentidade, lerDinheiro };
