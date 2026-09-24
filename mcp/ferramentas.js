'use strict';
// Ferramentas do servidor MCP (só leitura). Cada uma recebe parâmetros simples e devolve um objeto
// que vira JSON. Nenhuma regra nova: os cálculos vêm de web/lib/ (espelho de app/lib/Regras*.ps1),
// e aqui só se escolhe o recorte e se formata a resposta. Nada é gravado em disco.
//
// Nomes digitados pela pessoa (vendedor, produto, filial) são convertidos para o ID antes de
// qualquer cálculo; dali em diante tudo se liga pelo ID (REGRAS_NEGOCIO.md §0.1).

const path = require('path');
const { lerAba } = require('../web/lib/xlsx');
const N = require('../web/lib/numeros');
const R = require('../web/lib/regras');
const RP = require('../web/lib/regras_pipeline');
const RE = require('../web/lib/regras_estoque');

const NOMES_MES = ['', 'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const LIMITE_PADRAO = 50;
const LIMITE_MAXIMO = 300;

// Erro que vira resposta para quem chamou (parâmetro inválido, nome ambíguo, planilha com erro).
class ErroFerramenta extends Error {
  constructor(mensagem, detalhe) { super(mensagem); this.detalhe = detalhe; }
}

// ---------- carga com cache: recarrega quando alguma planilha muda (como web/app.js) ----------

function criarFonte(dirDados) {
  const dirImportacoes = path.join(dirDados, 'importacoes');   // §9.8: importações confirmadas, se existirem
  const cache = {};
  const lembrar = (chave, assinatura, carregar) => {
    if (cache[chave] && cache[chave].assinatura === assinatura) return cache[chave].valor;
    let valor;
    try { valor = carregar(); } catch (e) { valor = { Erros: [`Falha ao ler as planilhas: ${e.message}`] }; }
    cache[chave] = { assinatura, valor };
    return valor;
  };
  const exigir = (base) => {
    // Planilha inválida: nada de números, só a lista de erros (como a tela do painel).
    if (base.Erros.length) throw new ErroFerramenta('As planilhas têm erros de validação; corrija antes de consultar.', { erros: base.Erros });
    return base;
  };
  const comercial = () => exigir(lembrar('comercial', R.assinaturaBase(dirDados, dirImportacoes), () => R.carregarBaseComercial(dirDados, dirImportacoes)));
  const pipeline = () => {
    const base = comercial();
    const a = `${RP.assinaturaPipeline(dirDados)}#${R.assinaturaBase(dirDados, dirImportacoes)}`;
    return exigir(lembrar('pipeline', a, () => RP.carregarBasePipeline(dirDados, base)));
  };
  const estoque = () => exigir(lembrar('estoque', RE.assinaturaEstoque(dirDados, dirImportacoes), () => RE.carregarBaseEstoque(dirDados, dirImportacoes)));
  const catalogo = () => {
    // Colunas do cadastro que as regras não carregam (marca, modelo, preço de tabela).
    const p = path.join(dirDados, 'produtos.xlsx');
    return lembrar('catalogo', R.assinaturaArquivo(p, 'produtos.xlsx'), () => lerAba(p, 'Produtos'));
  };
  return { dirDados, comercial, pipeline, estoque, catalogo };
}

// ---------- formatação da resposta ----------

// Centavos -> reais com 2 casas (o número só é exibido; as contas continuam em centavos, §0.4).
function reais(c) { return c === null || c === undefined ? null : Number(N.formatarEscalado(c, 2, false).replace(',', '.')); }
function data(n) { return n === null || n === undefined ? null : N.fmtDataIso(n); }
function mesTexto(mes) { return `${String(mes).padStart(2, '0')}/${R.ANO_METAS}`; }

// §0.5: atingimento guardado como fração, exibido com 1 casa no formato brasileiro.
function percentual(f) {
  if (f === null || f === undefined) return { fracao: null, texto: 'sem meta' };
  return { fracao: Number(N.divArred(f.n * 10000n, f.d)) / 10000, texto: N.pct(f) };
}

// §2.3 e §3.3: desligado e admitido no ano aparecem marcados.
function marcacao(v) {
  if (!v.Ativo) return `desligado em ${N.fmtData(v.Desligamento)}`;
  if (v.Admissao !== null && N.anoDe(v.Admissao) === R.ANO_METAS) return `admissão em ${N.fmtMesAno(v.Admissao)}`;
  return null;
}
function vendedorResumo(v) { return { id: v.Id, nome: v.Nome, filial: v.IdFilial, ativo: v.Ativo, marcacao: marcacao(v) }; }

// ---------- entrada ----------

function semAcento(s) { return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
// Todas as palavras digitadas aparecem no texto, sem diferenciar maiúsculas nem acentos.
function casaPalavras(texto, busca) {
  const alvo = semAcento(texto || '');
  return semAcento(busca).split(/\s+/).filter(Boolean).every((p) => alvo.includes(p));
}

function texto(args, nome) {
  const v = args[nome];
  if (v === undefined || v === null) return null;
  if (typeof v !== 'string' && typeof v !== 'number') throw new ErroFerramenta(`O parâmetro '${nome}' deve ser um texto.`);
  const s = String(v).trim();
  return s === '' ? null : s;
}
function inteiro(args, nome, min, max) {
  const v = args[nome];
  if (v === undefined || v === null || v === '') return null;
  const n = typeof v === 'number' ? v : /^\s*[+-]?\d+\s*$/.test(String(v)) ? parseInt(v, 10) : NaN;
  if (!Number.isInteger(n) || n < min || n > max) throw new ErroFerramenta(`O parâmetro '${nome}' deve ser um número inteiro de ${min} a ${max}.`);
  return n;
}
function logico(args, nome) {
  const v = args[nome];
  if (v === undefined || v === null || v === '') return null;
  if (typeof v === 'boolean') return v;
  if (/^(true|sim|1)$/i.test(String(v))) return true;
  if (/^(false|nao|não|0)$/i.test(String(v))) return false;
  throw new ErroFerramenta(`O parâmetro '${nome}' deve ser verdadeiro ou falso.`);
}

// Vendedor por ID (V003) ou por nome (parte do nome basta). Ambíguo ou inexistente vira erro com as opções.
function resolverVendedor(base, busca) {
  const lista = N.ordenar(base.Vendedores, (v) => v.Id);
  const porId = lista.find((v) => N.igual(v.Id, busca));
  if (porId) return porId;
  const achados = lista.filter((v) => casaPalavras(v.Nome, busca));
  if (achados.length === 1) return achados[0];
  const opcoes = (achados.length ? achados : lista).map(vendedorResumo);
  throw new ErroFerramenta(achados.length
    ? `Mais de um vendedor corresponde a '${busca}'. Informe o ID ou o nome completo.`
    : `Nenhum vendedor corresponde a '${busca}'.`, { vendedores: opcoes });
}

// Filial por ID (F01) ou nome (Cascavel, Chapeco).
function resolverFilial(filiais, busca) {
  const f = filiais.find((x) => N.igual(x.Id, busca) || semAcento(x.Nome) === semAcento(busca));
  if (!f) throw new ErroFerramenta(`Filial desconhecida: '${busca}'.`, { filiais: filiais.map((x) => ({ id: x.Id, nome: x.Nome })) });
  return f;
}

function periodoValido(base, mesInicial, mesFinal) {
  const meses = R.mesesDisponiveis(base);
  const max = meses.length ? meses[meses.length - 1] : 0;
  if (!max) throw new ErroFerramenta(`Não há vendas de ${R.ANO_METAS} na base.`);
  const de = mesInicial ?? 1;
  const ate = mesFinal ?? max;
  if (de > max || ate > max) throw new ErroFerramenta(`As vendas vão até ${N.fmtData(base.DataBase)}: escolha meses de 1 a ${max}.`);
  if (de > ate) throw new ErroFerramenta('O mês inicial não pode ser depois do mês final.');
  return [de, ate];
}

// ---------- 1. consultar_vendedor ----------

function consultarVendedor(fonte, args) {
  const busca = texto(args, 'vendedor');
  if (!busca) throw new ErroFerramenta("Informe o parâmetro 'vendedor' (ID, como V003, ou nome).");
  const base = fonte.comercial();
  const v = resolverVendedor(base, busca);
  const [de, ate] = periodoValido(base, inteiro(args, 'mes_inicial', 1, 12), inteiro(args, 'mes_final', 1, 12));

  const d = R.desempenho(base, de, ate);
  const l = d.Linhas.find((x) => x.Vendedor.Id === v.Id);
  const det = R.detalheVendedor(base, v.Id, de, ate);
  const canceladas = det.Vendas.filter((x) => !R.vendaRealizada(x));
  let valorCanceladas = 0;
  for (const x of canceladas) valorCanceladas += x.Valor;

  return {
    vendedor: {
      id: v.Id, nome: v.Nome, email: v.Email, filial: { id: v.IdFilial, nome: v.Filial }, status: v.Status, ativo: v.Ativo,
      admissao: data(v.Admissao), desligamento: data(v.Desligamento), marcacao: marcacao(v),
    },
    periodo: {
      mes_inicial: de, mes_final: ate, texto: `${mesTexto(de)} a ${mesTexto(ate)}`,
      data_base_vendas: data(d.DataBase),
      parcial: d.Parcial ? `parcial até ${N.fmtData(d.DataBase).slice(0, 5)}` : null,   // §4.5: meta cheia, sem pró-rata
    },
    meses_com_meta: l.MesesComMeta,
    meta: reais(l.Meta),
    realizado: reais(l.RealizadoMesesMeta),
    atingimento: percentual(l.Atingimento),
    falta_para_meta: l.Meta > 0 ? reais(l.Meta - l.RealizadoMesesMeta) : null,   // §4.3
    faturamento_no_periodo: reais(l.Realizado),
    faturamento_fora_dos_meses_com_meta: reais(l.ForaMesesMeta),
    vendas_realizadas: l.QtdVendas,
    ticket_medio: l.QtdVendas ? reais(Number(N.divArred(l.Realizado, l.QtdVendas))) : null,
    canceladas_fora_do_realizado: { quantidade: canceladas.length, valor: reais(valorCanceladas) },
    por_mes: det.Meses.map((m) => ({
      mes: m.Mes, nome: NOMES_MES[m.Mes], meta: reais(m.Meta), realizado: reais(m.Realizado),
      atingimento: percentual(m.Atingimento), vendas_realizadas: m.QtdVendas,
    })),
    criterio: 'Realizado = vendas com Status Faturada (§1), só nos meses em que há meta; atingimento = realizado ÷ meta desses mesmos meses (§4.1). Mês sem linha de meta é "sem meta", não meta zero (§3.1).',
  };
}

// ---------- 2. buscar_produto ----------

function buscarProduto(fonte, args) {
  const busca = texto(args, 'busca');
  const categoria = texto(args, 'categoria');
  const limite = inteiro(args, 'limite', 1, LIMITE_MAXIMO) ?? LIMITE_PADRAO;
  if (!busca && !categoria) throw new ErroFerramenta("Informe 'busca' (código ou parte do nome) e/ou 'categoria'.");

  const cadastro = fonte.catalogo();
  const categorias = N.unicosOrdenados(cadastro.map((r) => r.Categoria));
  let cat = null;
  if (categoria) {
    cat = categorias.find((c) => semAcento(c) === semAcento(categoria)) || null;
    if (!cat) {
      const parecidas = categorias.filter((c) => casaPalavras(c, categoria));
      if (parecidas.length !== 1) throw new ErroFerramenta(`Categoria desconhecida: '${categoria}'.`, { categorias });
      cat = parecidas[0];
    }
  }

  // Estoque atual por produto (soma das filiais na foto em vigor, §5); sem foto válida, fica de fora.
  let unidades = null, dataFoto = null;
  try {
    const e = fonte.estoque();
    unidades = new Map();
    for (const l of e.Linhas) unidades.set(l.Produto.Id, (unidades.get(l.Produto.Id) || 0) + l.Quantidade);
    dataFoto = data(e.DataBase);
  } catch (e) { if (!(e instanceof ErroFerramenta)) throw e; }

  const porCodigo = busca ? cadastro.filter((r) => N.igual(r['ID Produto'], busca)) : [];
  const achados = (porCodigo.length ? porCodigo : cadastro.filter((r) => !busca || casaPalavras(`${r.Produto} ${r.Marca || ''} ${r.Modelo || ''}`, busca)))
    .filter((r) => !cat || N.igual(r.Categoria, cat));
  const produtos = N.ordenar(achados, (r) => r['ID Produto']).slice(0, limite).map((r) => ({
    id: r['ID Produto'], produto: r.Produto, categoria: r.Categoria, marca: r.Marca, modelo: r.Modelo,
    preco_tabela: reais(N.dinheiro(r['Preço de Tabela'])), unidade: r.Unidade, status: r.Status,
    descontinuado: N.igual(r.Status, 'Descontinuado'),   // §5.5: fora do catálogo de venda, mas continua no estoque
    estoque_unidades: unidades ? unidades.get(r['ID Produto']) || 0 : null,
  }));

  return {
    busca, categoria: cat, encontrados: achados.length, mostrados: produtos.length,
    foto_estoque: dataFoto, produtos,
    ...(achados.length ? {} : { categorias }),
  };
}

// ---------- 3. ver_estoque ----------

function verEstoque(fonte, args) {
  const buscaProduto = texto(args, 'produto');
  const buscaFilial = texto(args, 'filial');
  const dias = inteiro(args, 'dias_parado', 1, RE.DIAS_PARADO_MAXIMO) ?? RE.DIAS_PARADO_PADRAO;   // §5.0
  const soParados = logico(args, 'somente_parados') === true;
  const e = fonte.estoque();

  let abertas = null;   // §5.4: oportunidades abertas do produto ao lado
  try { abertas = RE.abertasPorProduto(fonte.pipeline()); } catch (err) { if (!(err instanceof ErroFerramenta)) throw err; }

  const filial = buscaFilial ? resolverFilial(e.Filiais, buscaFilial) : null;
  let produtos = null;
  if (buscaProduto) {
    const porId = e.Produtos.filter((p) => N.igual(p.Id, buscaProduto));
    produtos = porId.length ? porId : e.Produtos.filter((p) => casaPalavras(p.Nome, buscaProduto));
    if (!produtos.length) throw new ErroFerramenta(`Nenhum produto corresponde a '${buscaProduto}'. Use buscar_produto para achar o código.`);
  }
  const ids = produtos ? new Set(produtos.map((p) => p.Id)) : null;

  const linhas = N.ordenar(e.Linhas.filter((l) => (!ids || ids.has(l.Produto.Id)) && (!filial || l.IdFilial === filial.Id)
    && (!soParados || RE.linhaParada(l, dias))), (l) => l.Produto.Id, (l) => l.IdFilial);

  const resumo = { linhas: linhas.length, unidades: 0, valor_imobilizado: 0, parados: { linhas: 0, valor: 0 }, abaixo_do_minimo: 0 };
  const saida = linhas.map((l) => {
    const parado = RE.linhaParada(l, dias);
    resumo.unidades += l.Quantidade; resumo.valor_imobilizado += l.Valor;
    if (parado) { resumo.parados.linhas++; resumo.parados.valor += l.Valor; }
    if (RE.abaixoMinimo(l)) resumo.abaixo_do_minimo++;
    const ops = abertas === null ? null : abertas.get(l.Produto.Id) || { Qtd: 0, Valor: 0 };
    return {
      produto: { id: l.Produto.Id, nome: l.Produto.Nome, categoria: l.Produto.Categoria, descontinuado: l.Produto.Descontinuado },
      filial: { id: l.IdFilial, nome: l.Filial },
      quantidade: l.Quantidade, estoque_minimo: l.Minimo, abaixo_do_minimo: RE.abaixoMinimo(l),
      custo_medio: reais(l.CustoMedio), valor_imobilizado: reais(l.Valor),
      ultima_entrada: data(l.UltimaEntrada), ultima_saida: data(l.UltimaSaida),
      dias_parado: l.DiasSemMovimento,   // dias sem movimento até a data da foto; null = sem movimento registrado
      sem_movimento_registrado: l.UltimoMovimento === null,
      parado,
      oportunidades_abertas_do_produto: ops === null ? null : { quantidade: ops.Qtd, valor: reais(ops.Valor) },
    };
  });
  resumo.valor_imobilizado = reais(resumo.valor_imobilizado);
  resumo.parados.valor = reais(resumo.parados.valor);

  return {
    foto: { data: data(e.DataBase), arquivo: e.NomeArquivo },
    filtro: { produto: buscaProduto, filial: filial ? filial.Id : null, somente_parados: soParados, dias_parado: dias },
    definicao_parado: `Parado = linha (produto, filial) com quantidade maior que zero e ${dias} dias ou mais sem entrada nem saída até ${N.fmtData(e.DataBase)}, a data da foto; sem nenhuma data registrada também conta como parado (§5).`,
    resumo,
    linhas: saida,
    ...(abertas === null ? { aviso: 'CRM indisponível: as oportunidades abertas do produto não foram calculadas.' } : {}),
  };
}

// ---------- 4. ver_meta ----------

function linhaMeta(l, mes, db) {
  const futuro = mes > N.mesDe(db) && N.anoDe(db) === R.ANO_METAS;
  return {
    vendedor: vendedorResumo(l.Vendedor),
    meta: l.MesesComMeta.length ? reais(l.Meta) : null,
    realizado: reais(l.RealizadoMesesMeta),
    // Mês depois da data-base ainda não tem vendas: sem atingimento (não é 0%).
    atingimento: futuro ? { fracao: null, texto: 'mês ainda não começou nos dados' } : percentual(l.Atingimento),
    falta_para_meta: l.Meta > 0 ? reais(l.Meta - l.RealizadoMesesMeta) : null,
    vendas_realizadas: l.QtdVendas,
  };
}

function situacaoMes(mes, db) {
  if (N.anoDe(db) !== R.ANO_METAS || mes > N.mesDe(db)) return `mês posterior à data-base de vendas (${N.fmtData(db)})`;
  if (R.parcial(mes, db)) return `parcial até ${N.fmtData(db).slice(0, 5)} (meta cheia do mês, sem pró-rata, §4.5)`;
  return 'mês fechado';
}

function verMeta(fonte, args) {
  const busca = texto(args, 'vendedor');
  const mes = inteiro(args, 'mes', 1, 12);
  const ano = inteiro(args, 'ano', 1900, 9999);
  if (ano !== null && ano !== R.ANO_METAS) throw new ErroFerramenta(`Só há metas de ${R.ANO_METAS}.`);
  const base = fonte.comercial();
  const db = base.DataBase;

  if (busca) {
    const v = resolverVendedor(base, busca);
    const meses = mes ? [mes] : Array.from({ length: 12 }, (_, i) => i + 1);
    return {
      vendedor: vendedorResumo(v), ano: R.ANO_METAS, data_base_vendas: data(db),
      meses: meses.map((m) => {
        const l = R.desempenho(base, m, m).Linhas.find((x) => x.Vendedor.Id === v.Id);
        const { vendedor, ...resto } = linhaMeta(l, m, db);
        return { mes: m, nome: NOMES_MES[m], situacao: situacaoMes(m, db), ...resto };
      }),
      criterio: 'Meta mensal da planilha de metas; mês sem linha de meta é "sem meta" (§3.1). Realizado = vendas faturadas do mês (§1).',
    };
  }

  // Sem vendedor: todos no mês (padrão: o mês da data-base de vendas), com a empresa = soma ÷ soma (§4.2).
  const m = mes ?? N.mesDe(db);
  const d = R.desempenho(base, m, m);
  const futuro = situacaoMes(m, db).startsWith('mês posterior');
  return {
    ano: R.ANO_METAS, mes: m, nome: NOMES_MES[m], situacao: situacaoMes(m, db), data_base_vendas: data(db),
    empresa: {
      meta: reais(d.Empresa.Meta), realizado: reais(d.Empresa.RealizadoMesesMeta),
      atingimento: futuro ? { fracao: null, texto: 'mês ainda não começou nos dados' } : percentual(d.Empresa.Atingimento),
      falta_para_meta: d.Empresa.Meta > 0 ? reais(d.Empresa.Meta - d.Empresa.RealizadoMesesMeta) : null,
    },
    vendedores: N.ordenar(d.Linhas, (l) => l.Vendedor.Id).map((l) => linhaMeta(l, m, db)),
    criterio: 'Lista por ID, não é ranking. A empresa soma realizado e meta de todos, inclusive desligados, e divide um pelo outro (§4.2).',
  };
}

// ---------- 5. listar_oportunidades ----------

function listarOportunidades(fonte, args) {
  const buscaVendedor = texto(args, 'vendedor');
  const buscaEtapa = texto(args, 'etapa');
  const vencidas = logico(args, 'vencidas');
  const limite = inteiro(args, 'limite', 1, LIMITE_MAXIMO) ?? LIMITE_PADRAO;
  const base = fonte.comercial();
  const p = fonte.pipeline();
  const vend = buscaVendedor ? resolverVendedor(base, buscaVendedor) : null;

  // Etapa: sem filtro = só abertas (o pipeline, §7); 'todas' inclui as fechadas.
  const etapas = N.unicosOrdenados(p.Oportunidades.map((o) => o.Etapa));
  let etapa = null;
  if (buscaEtapa && semAcento(buscaEtapa) !== 'todas') {
    if (semAcento(buscaEtapa) === 'abertas') etapa = null;
    else {
      etapa = etapas.find((e) => semAcento(e) === semAcento(buscaEtapa)) || null;
      if (!etapa) {
        const parecidas = etapas.filter((e) => casaPalavras(e, buscaEtapa));
        if (parecidas.length !== 1) throw new ErroFerramenta(`Etapa desconhecida: '${buscaEtapa}'.`, { etapas: [...etapas, 'abertas', 'todas'] });
        etapa = parecidas[0];
      }
    }
  }
  const todas = buscaEtapa !== null && semAcento(buscaEtapa) === 'todas';

  const db = p.DataBase;
  const escolhidas = p.Oportunidades.filter((o) => {
    if (vend && o.IdVendedor !== vend.Id) return false;
    if (etapa ? !N.igual(o.Etapa, etapa) : !todas && !o.Aberta) return false;
    const atrasada = o.Aberta && o.Previsao < db;   // §7: ATRASADA ⇔ aberta e previsão < data-base do CRM
    if (vencidas !== null && atrasada !== vencidas) return false;
    return true;
  });
  // Abertas vencidas primeiro, do maior atraso para o menor; depois pela previsão e pelo ID (§10.4).
  const ordenadas = N.ordenar(escolhidas, [(o) => o.Aberta && o.Previsao < db, true], (o) => o.Previsao, (o) => o.Id);

  const resumo = { quantidade: 0, valor: 0, ponderado_das_abertas: 0, vencidas: { quantidade: 0, valor: 0 } };
  const lista = ordenadas.map((o) => {
    const v = base.VendedorPorId.get(o.IdVendedor);
    const cli = p.ClientePorId.get(o.IdCliente);
    const prod = p.ProdutoPorId.get(o.IdProduto);
    const atrasada = o.Aberta && o.Previsao < db;
    const pond = o.Aberta ? Number(N.divArred(BigInt(o.Valor) * o.Probabilidade.n, 10n ** BigInt(o.Probabilidade.escala))) : null;
    resumo.quantidade++; resumo.valor += o.Valor;
    if (o.Aberta) resumo.ponderado_das_abertas += pond;
    if (atrasada) { resumo.vencidas.quantidade++; resumo.vencidas.valor += o.Valor; }
    return {
      id: o.Id, etapa: o.Etapa, aberta: o.Aberta,
      vendedor: vendedorResumo(v), orfa: o.Aberta && !v.Ativo,   // §2.4: aberta de vendedor desligado
      cliente: { id: cli.Id, nome: cli.Nome, cidade: cli.Cidade, uf: cli.UF },
      produto: { id: o.IdProduto, nome: o.Produto, categoria: prod.Categoria },
      quantidade: o.Quantidade === null ? null : Number(o.Quantidade),
      valor_estimado: reais(o.Valor),
      probabilidade: N.probParaNumero(o.Probabilidade),
      ponderado: reais(pond),
      criacao: data(o.Criacao), previsao_fechamento: data(o.Previsao),
      vencida: atrasada, dias_de_atraso: atrasada ? db - o.Previsao : 0,
      origem: o.Origem, motivo_da_perda: o.MotivoPerda || null,
    };
  });
  resumo.valor = reais(resumo.valor);
  resumo.ponderado_das_abertas = reais(resumo.ponderado_das_abertas);
  resumo.vencidas.valor = reais(resumo.vencidas.valor);

  return {
    data_base_crm: data(db),
    filtro: { vendedor: vend ? vend.Id : null, etapa: etapa || (todas ? 'todas' : 'abertas'), vencidas },
    resumo, mostradas: Math.min(limite, lista.length),
    oportunidades: lista.slice(0, limite),
    criterio: `Aberta = etapa diferente de Fechada Ganha e Fechada Perdida; vencida = aberta com previsão de fechamento antes de ${N.fmtData(db)}, a data-base do CRM (§7). O CRM é uma foto: não depende do período das metas.`,
  };
}

// ---------- catálogo das ferramentas (nome, descrição, parâmetros) ----------

const somenteLeitura = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };
const txt = (description) => ({ type: 'string', description });
const int = (description, minimum, maximum) => ({ type: 'integer', description, minimum, maximum });

const FERRAMENTAS = [
  {
    name: 'consultar_vendedor',
    title: 'Consultar vendedor',
    description: 'Dados de um vendedor (filial, status, admissão, desligamento), meta, realizado, atingimento e falta para a meta num período de meses inteiros de 2026, com a abertura mês a mês. Sem período, usa o acumulado do ano até a data-base de vendas.',
    inputSchema: {
      type: 'object',
      properties: {
        vendedor: txt('ID (ex.: V003) ou nome, inteiro ou em parte (ex.: "Mariana").'),
        mes_inicial: int('Primeiro mês do período (1 a 12). Padrão: 1.', 1, 12),
        mes_final: int('Último mês do período (1 a 12). Padrão: o mês da data-base de vendas.', 1, 12),
      },
      required: ['vendedor'],
    },
    executar: consultarVendedor,
  },
  {
    name: 'buscar_produto',
    title: 'Buscar produto',
    description: 'Procura produtos no cadastro por código (ex.: P008), por parte do nome, marca ou modelo, e/ou por categoria. Devolve preço de tabela, status (descontinuado ou não) e as unidades em estoque somando as filiais.',
    inputSchema: {
      type: 'object',
      properties: {
        busca: txt('Código do produto ou palavras do nome, da marca ou do modelo.'),
        categoria: txt('Categoria (ex.: "Tratores"). Pode ser usada sozinha para listar a categoria.'),
        limite: int(`Máximo de produtos na resposta (padrão ${LIMITE_PADRAO}).`, 1, LIMITE_MAXIMO),
      },
    },
    executar: buscarProduto,
  },
  {
    name: 'ver_estoque',
    title: 'Ver estoque',
    description: 'Linhas de estoque (produto, filial) da foto em vigor: quantidade, mínimo, custo médio, valor imobilizado, últimas entrada e saída, dias parado e se está parado, com as oportunidades abertas do produto ao lado. Filtre por produto e/ou filial; sem filtros, devolve o estoque inteiro.',
    inputSchema: {
      type: 'object',
      properties: {
        produto: txt('Código do produto (ex.: P008) ou parte do nome.'),
        filial: txt('Filial: F01, F02, F03 ou o nome (Cascavel, Chapecó, Passo Fundo).'),
        dias_parado: int(`Prazo para considerar parado, em dias (padrão ${RE.DIAS_PARADO_PADRAO}).`, 1, RE.DIAS_PARADO_MAXIMO),
        somente_parados: { type: 'boolean', description: 'true para devolver só as linhas paradas.' },
      },
    },
    executar: verEstoque,
  },
  {
    name: 'ver_meta',
    title: 'Ver meta',
    description: 'Meta mensal de 2026 com o realizado e o atingimento do mês. Com vendedor e mês: aquele mês. Só com vendedor: os 12 meses dele. Só com mês (ou sem nada, que usa o mês da data-base): todos os vendedores e o total da empresa.',
    inputSchema: {
      type: 'object',
      properties: {
        vendedor: txt('ID (ex.: V012) ou nome, inteiro ou em parte.'),
        mes: int('Mês de 1 a 12.', 1, 12),
        ano: int(`Ano (só há metas de ${R.ANO_METAS}).`, R.ANO_METAS, R.ANO_METAS),
      },
    },
    executar: verMeta,
  },
  {
    name: 'listar_oportunidades',
    title: 'Listar oportunidades',
    description: 'Oportunidades do CRM (foto na data-base do CRM), com valor, probabilidade, ponderado, previsão e atraso. Sem etapa, lista só as abertas (o pipeline). As vencidas vêm primeiro, da mais atrasada para a menos.',
    inputSchema: {
      type: 'object',
      properties: {
        vendedor: txt('ID (ex.: V011) ou nome do vendedor dono da oportunidade.'),
        etapa: txt('Prospecção, Qualificação, Proposta Enviada, Negociação, Fechada Ganha, Fechada Perdida, "abertas" (padrão) ou "todas".'),
        vencidas: { type: 'boolean', description: 'true: só abertas com previsão de fechamento vencida; false: só as que não estão vencidas.' },
        limite: int(`Máximo de oportunidades na lista (padrão ${LIMITE_PADRAO}); o resumo sempre conta todas.`, 1, LIMITE_MAXIMO),
      },
    },
    executar: listarOportunidades,
  },
].map((f) => ({ ...f, annotations: { title: f.title, ...somenteLeitura } }));

// Executa uma ferramenta: { ok: true, resultado } ou { ok: false, erro, detalhe }.
function executar(fonte, nome, args) {
  const f = FERRAMENTAS.find((x) => x.name === nome);
  if (!f) return { ok: false, erro: `Ferramenta desconhecida: ${nome}` };
  try {
    return { ok: true, resultado: f.executar(fonte, args || {}) };
  } catch (e) {
    if (e instanceof ErroFerramenta) return { ok: false, erro: e.message, detalhe: e.detalhe };
    throw e;
  }
}

module.exports = { FERRAMENTAS, criarFonte, executar };
