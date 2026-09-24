'use strict';
// Regras do pipeline do CRM e do funil — espelho de app/lib/RegrasPipeline.ps1 (REGRAS_NEGOCIO.md §7 e §8).

const fs = require('fs');
const path = require('path');
const { lerAba } = require('./xlsx');
const N = require('./numeros');
const R = require('./regras');

const DATA_BASE_CRM = N.dia(2026, 8, 31);   // §0.3: o arquivo não traz a data da foto
const ETAPAS_FECHADAS = ['Fechada Ganha', 'Fechada Perdida'];
const ARQUIVOS_PIPELINE = ['crm_oportunidades.xlsx', 'clientes.xlsx', 'vendedores.xlsx', 'produtos.xlsx'];
const ETAPAS_FUNIL = ['Prospecção', 'Qualificação', 'Proposta Enviada', 'Negociação'];   // §8, de cima para baixo

function assinaturaPipeline(dirDados) { return ARQUIVOS_PIPELINE.map((n) => R.assinaturaArquivo(path.join(dirDados, n), n)).join(';'); }

function aberta(etapa) { return !N.contem(ETAPAS_FECHADAS, etapa); }   // §7

function carregarBasePipeline(dirDados, comercial) {
  const erros = [];
  const caminhos = {};
  for (const nome of ARQUIVOS_PIPELINE) {
    const p = path.join(dirDados, nome);
    if (!R.existe(p)) erros.push(`Arquivo não encontrado: ${p}`);
    caminhos[nome] = p;
  }
  if (erros.length) return { Erros: erros };

  const fil = R.mapaFiliais(caminhos['vendedores.xlsx']);

  const produtoPorId = new Map();   // só o necessário para a categoria do funil (§8.4)
  for (const r of lerAba(caminhos['produtos.xlsx'], 'Produtos')) {
    if (r['ID Produto']) produtoPorId.set(r['ID Produto'], { Id: r['ID Produto'], Nome: r.Produto, Categoria: r.Categoria });
  }

  const clientePorId = new Map();
  for (const r of lerAba(caminhos['clientes.xlsx'], 'Clientes')) {
    const onde = `clientes.xlsx, linha ${r._Linha}`;
    const id = r['ID Cliente'];
    if (!id) { erros.push(`${onde}: ID Cliente vazio`); continue; }
    if (clientePorId.has(id)) { erros.push(`${onde}: ID Cliente ${id} duplicado`); continue; }
    const fa = r['Filial de Atendimento'];
    const idFilial = fa ? fil.idPorNome.get(fa.normalize('NFC')) : undefined;
    if (!idFilial) { erros.push(`${onde} (${id}): filial de atendimento desconhecida '${fa ?? ''}'`); continue; }
    clientePorId.set(id, { Id: id, Nome: r.Cliente, Tipo: r.Tipo, Cidade: r.Cidade, UF: r.UF, IdFilial: idFilial, Filial: fil.nome.get(idFilial) });
  }

  const oportunidades = [];
  const ids = new Set();
  for (const r of lerAba(caminhos['crm_oportunidades.xlsx'], 'Oportunidades')) {
    const onde = `crm_oportunidades.xlsx, linha ${r._Linha}`;
    const id = r['ID Oportunidade'];
    if (!id) { erros.push(`${onde}: ID Oportunidade vazio`); continue; }
    if (ids.has(id)) { erros.push(`${onde}: ID Oportunidade ${id} duplicado`); continue; }
    ids.add(id);
    const vend = comercial.VendedorPorId.get(r['ID Vendedor']);
    const cli = clientePorId.get(r['ID Cliente']);
    const valor = N.dinheiro(r['Valor Estimado']);
    const prob = N.lerDecimal(r.Probabilidade);
    const criacao = N.lerData(r['Data de Criação']);
    const prev = r['Previsão de Fechamento'];
    const previsao = prev ? N.lerData(prev) : null;
    const ab = aberta(r.Etapa);
    let falhou = false;
    if (!vend) { erros.push(`${onde} (${id}): ID Vendedor '${r['ID Vendedor'] ?? ''}' não existe no cadastro`); falhou = true; }
    if (!cli) { erros.push(`${onde} (${id}): ID Cliente '${r['ID Cliente'] ?? ''}' não existe em clientes.xlsx`); falhou = true; }
    if (!produtoPorId.has(r['ID Produto'] ?? '')) { erros.push(`${onde} (${id}): ID Produto '${r['ID Produto'] ?? ''}' não existe em produtos.xlsx`); falhou = true; }
    if (!r.Etapa) { erros.push(`${onde} (${id}): Etapa vazia`); falhou = true; }
    if (valor === null || valor < 0) { erros.push(`${onde} (${id}): Valor Estimado inválido '${r['Valor Estimado'] ?? ''}'`); falhou = true; }
    if (prob === null || prob.n < 0n || prob.n > 10n ** BigInt(prob.escala)) { erros.push(`${onde} (${id}): Probabilidade inválida '${r.Probabilidade ?? ''}' (esperado fração de 0 a 1)`); falhou = true; }
    if (criacao === null) { erros.push(`${onde} (${id}): Data de Criação inválida '${r['Data de Criação'] ?? ''}'`); falhou = true; }
    if (prev && previsao === null) { erros.push(`${onde} (${id}): Previsão de Fechamento inválida '${prev}'`); falhou = true; }
    if (ab && !prev) { erros.push(`${onde} (${id}): oportunidade aberta sem Previsão de Fechamento`); falhou = true; }
    if (falhou) continue;
    // §2.6: oportunidade nova depois do desligamento do dono é erro. §3.4: não validar criação ≥ admissão.
    if (vend.Desligamento !== null && criacao > vend.Desligamento) {
      erros.push(`${onde} (${id}): criada em ${N.fmtData(criacao)} para vendedor desligado em ${N.fmtData(vend.Desligamento)}`);
      continue;
    }
    oportunidades.push({
      Id: id, Criacao: criacao, IdVendedor: vend.Id, IdCliente: cli.Id, IdProduto: r['ID Produto'], Produto: r.Produto,
      Quantidade: r.Quantidade, Valor: valor, Etapa: r.Etapa, Probabilidade: prob, Previsao: previsao, Origem: r.Origem,
      MotivoPerda: r['Motivo da Perda'], Aberta: ab, Linha: r._Linha,
    });
  }

  return {
    Erros: erros, DataBase: DATA_BASE_CRM, Oportunidades: oportunidades, ClientePorId: clientePorId, ProdutoPorId: produtoPorId,
    Filiais: fil.lista, NomeArquivo: 'crm_oportunidades.xlsx',
  };
}

function ponderado(o) { return Number(N.divArred(BigInt(o.Valor) * o.Probabilidade.n, 10n ** BigInt(o.Probabilidade.escala))); }

function abertaComContexto(o, vend, pipeline, extra) {
  const db = pipeline.DataBase;
  const atrasada = o.Previsao < db;
  return Object.assign({
    Oportunidade: o, Vendedor: vend, Cliente: pipeline.ClientePorId.get(o.IdCliente), Ponderado: ponderado(o),
    Atrasada: atrasada, DiasAtraso: atrasada ? db - o.Previsao : 0, Orfa: !vend.Ativo,   // §2.4
  }, extra);
}

function visaoPipeline(comercial, desempenho, pipeline) {
  // Uma linha por vendedor: meta e realizado do período (§4) + pipeline da foto do CRM (§7).
  let abertas = pipeline.Oportunidades.filter((o) => o.Aberta).map((o) => abertaComContexto(o, comercial.VendedorPorId.get(o.IdVendedor), pipeline));
  abertas = N.ordenar(abertas, [(a) => a.Atrasada, true], (a) => a.Oportunidade.Previsao, (a) => a.Oportunidade.Id);

  const porVendedor = new Map();
  for (const a of abertas) {
    const id = a.Vendedor.Id;
    if (!porVendedor.has(id)) porVendedor.set(id, { Qtd: 0, Valor: 0, Ponderado: 0, QtdAtrasadas: 0, ValorAtrasado: 0 });
    const p = porVendedor.get(id);
    p.Qtd++; p.Valor += a.Oportunidade.Valor; p.Ponderado += a.Ponderado;
    if (a.Atrasada) { p.QtdAtrasadas++; p.ValorAtrasado += a.Oportunidade.Valor; }
  }

  // §7.4: todos os vendedores (órfãs de inativos entram nos totais).
  const linhas = N.ordenar(desempenho.Linhas, (l) => l.Vendedor.Id).map((l) => {
    const p = porVendedor.get(l.Vendedor.Id);
    return {
      Vendedor: l.Vendedor, MesesComMeta: l.MesesComMeta, Meta: l.Meta, RealizadoMesesMeta: l.RealizadoMesesMeta,
      ForaMesesMeta: l.ForaMesesMeta, Atingimento: l.Atingimento, Gap: l.Meta > 0 ? l.Meta - l.RealizadoMesesMeta : null,   // §4.3
      QtdAbertas: p ? p.Qtd : 0, Pipeline: p ? p.Valor : 0, Ponderado: p ? p.Ponderado : 0,
      QtdAtrasadas: p ? p.QtdAtrasadas : 0, ValorAtrasado: p ? p.ValorAtrasado : 0,
    };
  });

  const tot = { Meta: 0, RealizadoMesesMeta: 0, ForaMesesMeta: 0, QtdAbertas: 0, Pipeline: 0, Ponderado: 0, QtdAtrasadas: 0, ValorAtrasado: 0 };
  for (const l of linhas) for (const k of Object.keys(tot)) tot[k] += l[k];
  tot.Atingimento = desempenho.Empresa.Atingimento;   // §4.2: soma ÷ soma
  tot.Gap = tot.Meta > 0 ? tot.Meta - tot.RealizadoMesesMeta : null;

  const orfas = abertas.filter((a) => a.Orfa);
  let valorOrfas = 0;
  for (const a of orfas) valorOrfas += a.Oportunidade.Valor;
  return {
    DataBaseCrm: pipeline.DataBase, Desempenho: desempenho, Linhas: linhas, Total: tot, Abertas: abertas,
    QtdOrfas: orfas.length, ValorOrfas: valorOrfas, NomeArquivoCrm: pipeline.NomeArquivo,
  };
}

function visaoFunil(comercial, pipeline, idVendedor, idFilial, categoria) {
  // §8: abertas por etapa ATUAL + fechadas à parte. Filtros em E; filial = filial do vendedor (§8.4).
  const probEtapa = new Map();
  for (const o of pipeline.Oportunidades) if (o.Aberta && !probEtapa.has(o.Etapa)) probEtapa.set(o.Etapa, o.Probabilidade);
  const extras = N.ordenar([...probEtapa.keys()].filter((e) => !N.contem(ETAPAS_FUNIL, e)), (e) => N.probParaNumero(probEtapa.get(e)), (e) => e);
  const etapas = new Map();
  for (const nome of [...ETAPAS_FUNIL, ...extras]) {
    etapas.set(nome, { Etapa: nome, Probabilidade: probEtapa.has(nome) ? probEtapa.get(nome) : null, Qtd: 0, Valor: 0, Ponderado: 0, QtdAtrasadas: 0, ValorAtrasado: 0 });
  }

  const abertas = [];
  const ganhas = { Qtd: 0, Valor: 0 };
  const perdidas = { Qtd: 0, Valor: 0 };
  const motivos = new Map();
  for (const o of pipeline.Oportunidades) {
    const vend = comercial.VendedorPorId.get(o.IdVendedor);
    const prod = pipeline.ProdutoPorId.get(o.IdProduto);
    if (idVendedor && !N.igual(o.IdVendedor, idVendedor)) continue;
    if (idFilial && !N.igual(vend.IdFilial, idFilial)) continue;
    if (categoria && !N.igual(prod.Categoria, categoria)) continue;
    if (N.igual(o.Etapa, 'Fechada Ganha')) { ganhas.Qtd++; ganhas.Valor += o.Valor; continue; }
    if (N.igual(o.Etapa, 'Fechada Perdida')) {
      perdidas.Qtd++; perdidas.Valor += o.Valor;
      const m = o.MotivoPerda ? o.MotivoPerda : '(sem motivo informado)';
      if (!motivos.has(m)) motivos.set(m, { Motivo: m, Qtd: 0, Valor: 0 });
      motivos.get(m).Qtd++; motivos.get(m).Valor += o.Valor;
      continue;
    }
    const a = abertaComContexto(o, vend, pipeline, { Produto: prod });
    abertas.push(a);
    const e = etapas.get(o.Etapa);
    e.Qtd++; e.Valor += o.Valor; e.Ponderado += a.Ponderado;
    if (a.Atrasada) { e.QtdAtrasadas++; e.ValorAtrasado += o.Valor; }
  }

  const ordem = new Map([...etapas.keys()].map((k, i) => [k, i]));
  const lista = N.ordenar(abertas, [(a) => ordem.get(a.Oportunidade.Etapa), true], [(a) => a.Atrasada, true], (a) => a.Oportunidade.Previsao, (a) => a.Oportunidade.Id);

  const tot = { Qtd: 0, Valor: 0, Ponderado: 0, QtdAtrasadas: 0, ValorAtrasado: 0 };
  for (const e of etapas.values()) for (const k of Object.keys(tot)) tot[k] += e[k];
  const fechadas = ganhas.Qtd + perdidas.Qtd;
  const valorFechadas = ganhas.Valor + perdidas.Valor;

  return {
    DataBaseCrm: pipeline.DataBase, NomeArquivoCrm: pipeline.NomeArquivo, Etapas: [...etapas.values()], Abertas: lista, Total: tot,
    Ganhas: ganhas, Perdidas: perdidas,
    Conversao: fechadas ? N.fracao(ganhas.Qtd, fechadas) : null,
    ConversaoValor: valorFechadas > 0 ? N.fracao(ganhas.Valor, valorFechadas) : null,
    Motivos: N.ordenar([...motivos.values()], [(m) => m.Qtd, true], (m) => m.Motivo),
    Filtro: { IdVendedor: idVendedor, IdFilial: idFilial, Categoria: categoria },
    Vendedores: N.ordenar(comercial.Vendedores, (v) => v.Nome),
    Filiais: pipeline.Filiais,
    Categorias: categoriasCrm(pipeline),   // peças não passam pelo CRM
  };
}

function categoriasCrm(pipeline) {
  return N.unicosOrdenados(pipeline.Oportunidades.map((o) => pipeline.ProdutoPorId.get(o.IdProduto).Categoria));
}

module.exports = { ETAPAS_FUNIL, assinaturaPipeline, carregarBasePipeline, visaoPipeline, visaoFunil, categoriasCrm };
