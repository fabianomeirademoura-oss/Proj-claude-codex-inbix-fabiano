'use strict';
// Regras de estoque — espelho de app/lib/RegrasEstoque.ps1 (REGRAS_NEGOCIO.md §5).
// Estoque é sempre por linha (produto, filial); o produto no agregado só entra em "sem estoque".

const fs = require('fs');
const path = require('path');
const { lerAba } = require('./xlsx');
const N = require('./numeros');
const R = require('./regras');

const DIAS_PARADO_PADRAO = 180;
const DIAS_PARADO_MAXIMO = 3650;
const PADRAO_ARQUIVO = /^estoque_(\d{4}-\d{2}-\d{2})\.xlsx$/i;
const PADRAO_IMPORTADO = /^estoque_(\d{4}-\d{2}-\d{2})__(\d{8}-\d{6})\.xlsx$/i;   // §9.7
const ARQUIVOS_FIXOS = ['produtos.xlsx', 'vendedores.xlsx'];

function fotosEstoque(dirDados, dirImportacoes) {
  // §5: a foto vale na data do nome do arquivo. §9.8: vale a mais recente; no empate, a importação mais recente.
  const fotos = [];
  const listar = (pasta, padrao, importada) => {
    if (!R.existe(pasta)) return;
    for (const nome of fs.readdirSync(pasta)) {
      const caminho = path.join(pasta, nome);
      if (!/^estoque_.*\.xlsx$/i.test(nome) || !fs.statSync(caminho).isFile()) continue;
      const m = padrao.exec(nome);
      if (!m) continue;
      const data = N.lerData(m[1]);
      if (data !== null) fotos.push({ nome, caminho, Data: data, Importada: importada, Carimbo: importada ? m[2] : '' });
    }
  };
  listar(dirDados, PADRAO_ARQUIVO, false);
  if (dirImportacoes) listar(path.join(dirImportacoes, 'estoque'), PADRAO_IMPORTADO, true);
  return N.ordenar(fotos, [(f) => f.Data, true], [(f) => f.Importada, true], [(f) => f.Carimbo, true]);
}

function assinaturaEstoque(dirDados, dirImportacoes) {
  const partes = ARQUIVOS_FIXOS.map((n) => R.assinaturaArquivo(path.join(dirDados, n), n));
  for (const f of fotosEstoque(dirDados, dirImportacoes)) partes.push(R.assinaturaArquivo(f.caminho, f.caminho));
  return partes.join(';');
}

function quantidade(texto) { const n = N.inteiro(texto); return n === null || n < 0 ? null : n; }

function carregarBaseEstoque(dirDados, dirImportacoes) {
  const erros = [];
  for (const nome of ARQUIVOS_FIXOS) if (!R.existe(path.join(dirDados, nome))) erros.push(`Arquivo não encontrado: ${path.join(dirDados, nome)}`);
  const foto = fotosEstoque(dirDados, dirImportacoes)[0];
  if (!foto) erros.push(`Nenhuma foto de estoque (estoque_AAAA-MM-DD.xlsx) em ${dirDados}`);
  if (erros.length) return { Erros: erros };

  const filiais = [];
  const idFilialPorNome = new Map();
  for (const f of lerAba(path.join(dirDados, 'vendedores.xlsx'), 'Filiais')) {
    idFilialPorNome.set(f.Filial.normalize('NFC'), f['ID Filial']);
    filiais.push({ Id: f['ID Filial'], Nome: f.Filial });
  }

  const produtos = [];
  const produtoPorId = new Map();
  for (const r of lerAba(path.join(dirDados, 'produtos.xlsx'), 'Produtos')) {
    const onde = `produtos.xlsx, linha ${r._Linha}`;
    const id = r['ID Produto'];
    if (!id) { erros.push(`${onde}: ID Produto vazio`); continue; }
    if (produtoPorId.has(id)) { erros.push(`${onde}: ID Produto ${id} duplicado`); continue; }
    if (!r.Categoria) { erros.push(`${onde} (${id}): Categoria vazia`); continue; }
    if (!N.contem(['Ativo', 'Descontinuado'], r.Status)) { erros.push(`${onde} (${id}): Status inválido '${r.Status ?? ''}'`); continue; }
    const p = { Id: id, Nome: r.Produto, Categoria: r.Categoria, Unidade: r.Unidade, Status: r.Status, Descontinuado: N.igual(r.Status, 'Descontinuado') };
    produtos.push(p);
    produtoPorId.set(id, p);
  }

  const linhas = [];
  const chaves = new Set();
  for (const r of lerAba(foto.caminho, 'Estoque')) {
    const onde = `${foto.nome}, linha ${r._Linha}`;
    const prod = produtoPorId.get(r['ID Produto']);
    const idFilial = r.Filial ? idFilialPorNome.get(r.Filial.normalize('NFC')) : undefined;
    const qtd = quantidade(r.Quantidade);
    const min = quantidade(r['Estoque Mínimo']);
    const custo = N.dinheiro(r['Custo Médio']);
    const te = r['Data da Última Entrada'], ts = r['Data da Última Saída'];
    const entrada = te ? N.lerData(te) : null;
    const saida = ts ? N.lerData(ts) : null;
    let falhou = false;
    if (!prod) { erros.push(`${onde}: ID Produto '${r['ID Produto'] ?? ''}' não existe em produtos.xlsx`); falhou = true; }
    if (!idFilial) { erros.push(`${onde}: filial desconhecida '${r.Filial ?? ''}'`); falhou = true; }
    if (qtd === null) { erros.push(`${onde}: Quantidade inválida '${r.Quantidade ?? ''}'`); falhou = true; }
    if (min === null) { erros.push(`${onde}: Estoque Mínimo inválido '${r['Estoque Mínimo'] ?? ''}'`); falhou = true; }
    if (custo === null || custo < 0) { erros.push(`${onde}: Custo Médio inválido '${r['Custo Médio'] ?? ''}'`); falhou = true; }
    if (te && entrada === null) { erros.push(`${onde}: Data da Última Entrada inválida '${te}'`); falhou = true; }
    if (ts && saida === null) { erros.push(`${onde}: Data da Última Saída inválida '${ts}'`); falhou = true; }
    if (falhou) continue;
    const chave = `${prod.Id}|${idFilial}`;
    if (chaves.has(chave)) { erros.push(`${onde}: produto ${prod.Id} repetido na filial ${r.Filial}`); continue; }
    chaves.add(chave);
    // §5: último movimento = a mais recente entre entrada e saída; as duas em branco = sem movimento registrado.
    const ultimo = entrada !== null && saida !== null ? (entrada > saida ? entrada : saida) : entrada !== null ? entrada : saida;
    linhas.push({
      Produto: prod, IdFilial: idFilial, Filial: r.Filial, Quantidade: qtd, Minimo: min, CustoMedio: custo,
      Valor: qtd * custo,   // §5: valor imobilizado da linha
      UltimaEntrada: entrada, UltimaSaida: saida, UltimoMovimento: ultimo,
      DiasSemMovimento: ultimo !== null ? foto.Data - ultimo : null, Linha: r._Linha,
    });
  }

  return { Erros: erros, DataBase: foto.Data, NomeArquivo: foto.nome, Filiais: filiais, Produtos: produtos, ProdutoPorId: produtoPorId, Linhas: linhas };
}

// §5: PARADO ⇔ Quantidade > 0 E (sem movimento registrado OU dias sem movimento ≥ prazo).
function linhaParada(l, dias) { return l.Quantidade > 0 && (l.DiasSemMovimento === null || l.DiasSemMovimento >= dias); }
function abaixoMinimo(l) { return l.Quantidade < l.Minimo; }

function abertasPorProduto(pipeline) {
  // §5.4: oportunidades abertas do produto (todas as filiais). null = CRM indisponível.
  if (!pipeline || pipeline.Erros.length) return null;
  const mapa = new Map();
  for (const o of pipeline.Oportunidades) {
    if (!o.Aberta) continue;
    if (!mapa.has(o.IdProduto)) mapa.set(o.IdProduto, { Qtd: 0, Valor: 0 });
    mapa.get(o.IdProduto).Qtd++;
    mapa.get(o.IdProduto).Valor += o.Valor;
  }
  return mapa;
}

function visaoEstoque(estoque, diasParado, abertas) {
  const semOp = { Qtd: 0, Valor: 0 };
  const opsDe = (id) => (abertas === null ? null : abertas.has(id) ? abertas.get(id) : semOp);

  const linhasPorProduto = new Map();
  for (const l of estoque.Linhas) {
    if (!linhasPorProduto.has(l.Produto.Id)) linhasPorProduto.set(l.Produto.Id, []);
    linhasPorProduto.get(l.Produto.Id).push(l);
  }
  const outras = (l) => N.ordenar(linhasPorProduto.get(l.Produto.Id).filter((x) => x.IdFilial !== l.IdFilial), (x) => x.IdFilial);

  // 1. Dinheiro no pátio: valor imobilizado por categoria × filial, com a parte parada.
  const novo = () => ({ Valor: 0, Parado: 0, Unidades: 0 });
  let categorias = N.unicosOrdenados(estoque.Produtos.map((p) => p.Categoria));
  const celulas = new Map();
  for (const c of categorias) for (const f of estoque.Filiais) celulas.set(`${c}|${f.Id}`, novo());
  const totFilial = new Map(estoque.Filiais.map((f) => [f.Id, novo()]));
  const totCategoria = new Map(categorias.map((c) => [c, novo()]));
  const total = novo();

  const parados = [];
  const abaixo = [];
  for (const l of estoque.Linhas) {
    const parado = linhaParada(l, diasParado);
    for (const acum of [celulas.get(`${l.Produto.Categoria}|${l.IdFilial}`), totFilial.get(l.IdFilial), totCategoria.get(l.Produto.Categoria), total]) {
      acum.Valor += l.Valor; acum.Unidades += l.Quantidade;
      if (parado) acum.Parado += l.Valor;
    }
    if (parado) parados.push({ Linha: l, Abertas: opsDe(l.Produto.Id), Outras: outras(l) });
    if (abaixoMinimo(l)) abaixo.push({ Linha: l, Falta: l.Minimo - l.Quantidade, Abertas: opsDe(l.Produto.Id), Outras: outras(l) });
  }
  categorias = N.ordenar(categorias, [(c) => totCategoria.get(c).Valor, true], (c) => c);

  // 2. Sem estoque: soma da quantidade nas filiais = 0 (produto sem linha na foto conta como zerado).
  let semEstoque = [];
  for (const p of estoque.Produtos) {
    const ls = linhasPorProduto.get(p.Id) || [];
    let soma = 0;
    for (const l of ls) soma += l.Quantidade;
    if (soma !== 0) continue;
    let ultSaida = null;
    for (const l of ls) if (l.UltimaSaida !== null && (ultSaida === null || l.UltimaSaida > ultSaida)) ultSaida = l.UltimaSaida;
    semEstoque.push({ Produto: p, Linhas: N.ordenar(ls, (l) => l.IdFilial), UltimaSaida: ultSaida, Abertas: opsDe(p.Id) });
  }
  semEstoque = N.ordenar(semEstoque, [(s) => (s.Abertas ? s.Abertas.Qtd : 0), true], (s) => s.Produto.Id);

  return {
    DataBase: estoque.DataBase, NomeArquivo: estoque.NomeArquivo, DiasParado: diasParado, Filiais: estoque.Filiais, Categorias: categorias,
    Celulas: celulas, TotFilial: totFilial, TotCategoria: totCategoria, Total: total,
    Parados: N.ordenar(parados, [(p) => (p.Linha.DiasSemMovimento === null ? 2147483647 : p.Linha.DiasSemMovimento), true], (p) => p.Linha.Produto.Id, (p) => p.Linha.IdFilial),
    SemEstoque: semEstoque,
    Abaixo: N.ordenar(abaixo, (a) => a.Linha.IdFilial, [(a) => a.Falta, true], (a) => a.Linha.Produto.Id),
    ComCrm: abertas !== null, Linhas: estoque.Linhas,
  };
}

module.exports = { DIAS_PARADO_PADRAO, DIAS_PARADO_MAXIMO, assinaturaEstoque, carregarBaseEstoque, linhaParada, abaixoMinimo, abertasPorProduto, visaoEstoque };
