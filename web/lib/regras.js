'use strict';
// Regras de negócio da Horizonte Máquinas — espelho de app/lib/Regras.ps1.
// Decide o que é venda realizada, meta do período e atingimento. As páginas só formatam.

const fs = require('fs');
const path = require('path');
const { lerAba } = require('./xlsx');
const N = require('./numeros');

const ANO_METAS = 2026;
const ARQUIVOS_BASE = { Vendedores: 'vendedores.xlsx', Metas: 'metas_2026.xlsx', Vendas: 'vendas_2026_jan-ago.xlsx' };
const PADRAO_VENDA_IMPORTADA = /^\d{8}-\d{6}__.+\.xlsx$/i;   // REGRAS_NEGOCIO.md §9.7

function existe(p) { try { fs.statSync(p); return true; } catch { return false; } }
function assinaturaArquivo(p, nome) {
  try { const s = fs.statSync(p); return `${nome}|${s.mtimeMs}|${s.size}`; } catch { return `${nome}|ausente`; }
}

function arquivosVendasImportados(dirImportacoes) {
  // §9.8: em ordem cronológica (o nome começa pelo carimbo AAAAMMDD-HHMMSS).
  if (!dirImportacoes) return [];
  const pasta = path.join(dirImportacoes, 'vendas');
  if (!existe(pasta)) return [];
  return N.ordenar(fs.readdirSync(pasta).filter((n) => /\.xlsx$/i.test(n) && PADRAO_VENDA_IMPORTADA.test(n) && fs.statSync(path.join(pasta, n)).isFile()), (n) => n)
    .map((n) => ({ nome: n, caminho: path.join(pasta, n) }));
}

function assinaturaBase(dirDados, dirImportacoes) {
  const partes = Object.values(ARQUIVOS_BASE).map((n) => assinaturaArquivo(path.join(dirDados, n), n));
  for (const f of arquivosVendasImportados(dirImportacoes)) partes.push(assinaturaArquivo(f.caminho, f.nome));
  return [...partes, ...assinaturaAlteracoes(dirImportacoes)].join(';');
}

// ---------- Alterações de cadastro (REGRAS_NEGOCIO.md §11) ----------
// O registro das alterações confirmadas é também o que se aplica por cima da base (§11.4):
// cada linha põe o valor Novo no Campo do Registro, na ordem do arquivo.

const COLUNAS_ALTERACOES = ['Quando', 'Lote', 'Login', 'Nome', 'Perfil', 'Ação', 'Tipo', 'Registro', 'Campo', 'Anterior', 'Novo'];
// Colunas de produtos.xlsx: um produto cadastrado por alteração (§11.8) nasce com todas elas.
const COLUNAS_PRODUTOS = ['ID Produto', 'Produto', 'Categoria', 'Marca', 'Modelo', 'Preço de Tabela', 'Custo Unitário', 'Unidade', 'Status'];

function arquivoAlteracoes(dirImportacoes) { return dirImportacoes ? path.join(dirImportacoes, 'alteracoes.csv') : null; }

function assinaturaAlteracoes(dirImportacoes) {
  const arq = arquivoAlteracoes(dirImportacoes);
  return arq && existe(arq) ? [assinaturaArquivo(arq, 'alteracoes.csv')] : [];
}

// Uma linha de CSV com ';', campos entre aspas e aspas dobradas (o que o Import-Csv do PowerShell lê).
function camposCsv(linha) {
  const campos = [];
  let i = 0;
  while (i <= linha.length) {
    let valor = '';
    if (linha[i] === '"') {
      i++;
      for (;;) {
        const j = linha.indexOf('"', i);
        if (j < 0) { valor += linha.slice(i); i = linha.length; break; }
        valor += linha.slice(i, j);
        if (linha[j + 1] === '"') { valor += '"'; i = j + 2; } else { i = j + 1; break; }
      }
      const fim = linha.indexOf(';', i);
      if (fim < 0) i = linha.length + 1; else i = fim + 1;
    } else {
      const fim = linha.indexOf(';', i);
      if (fim < 0) { valor = linha.slice(i); i = linha.length + 1; } else { valor = linha.slice(i, fim); i = fim + 1; }
    }
    campos.push(valor);
  }
  return campos;
}

function lerAlteracoes(dirImportacoes) {
  const arq = arquivoAlteracoes(dirImportacoes);
  if (!arq || !existe(arq)) return [];
  const linhas = fs.readFileSync(arq, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).filter((l) => l !== '');
  if (!linhas.length) return [];
  const cab = camposCsv(linhas[0]);
  return linhas.slice(1).map((l) => {
    const c = camposCsv(l);
    const o = {};
    cab.forEach((k, i) => { o[k] = c[i] ?? ''; });
    return o;
  });
}

// Registro de uma meta: "V003/2026-09" (vendedor/ano-mês).
function chaveMeta(r) {
  const ano = N.inteiro(r.Ano);
  const mes = N.inteiro(r['Mês']);
  if (!r['ID Vendedor'] || ano === null || mes === null) return null;
  return `${r['ID Vendedor']}/${String(ano).padStart(4, '0')}-${String(mes).padStart(2, '0')}`;
}

// Aplica as alterações de um tipo às linhas lidas de uma planilha. Devolve as linhas em vigor.
// colunas: colunas de uma linha nova (só produto e meta podem nascer de uma alteração).
function mesclarAlteracoes(linhas, alteracoes, tipo, chave, colunas, erros) {
  const lista = [...linhas];
  const porChave = new Map();
  for (const r of linhas) { const k = chave(r); if (k && !porChave.has(k)) porChave.set(k, r); }
  const semMeta = new Map();
  for (const a of alteracoes) {
    if (a.Tipo !== tipo) continue;
    let r = porChave.get(a.Registro);
    if (!r) {
      if (!colunas.length) { erros.push(`alteracoes.csv (${a.Lote}): ${tipo} '${a.Registro}' não existe`); continue; }
      r = {};
      for (const c of colunas) r[c] = '';
      if (tipo === 'produto') r['ID Produto'] = a.Registro;
      if (tipo === 'meta') {
        const m = /^(.+)\/(\d{4})-(\d{2})$/.exec(a.Registro);
        if (!m) { erros.push(`alteracoes.csv (${a.Lote}): meta '${a.Registro}' inválida`); continue; }
        r['ID Vendedor'] = m[1]; r.Ano = String(parseInt(m[2], 10)); r['Mês'] = String(parseInt(m[3], 10));
      }
      r._Linha = 0;
      lista.push(r);
      porChave.set(a.Registro, r);
    }
    r[a.Campo] = a.Novo;
    if (tipo === 'meta') semMeta.set(a.Registro, a.Novo === '');   // §11.7: valor vazio volta a "sem meta"
  }
  if (tipo === 'meta') return lista.filter((r) => { const k = chave(r); return !(k && semMeta.get(k)); });
  return lista;
}

function mapaFiliais(caminhoVendedores) {
  const idPorNome = new Map();
  const nome = new Map();
  const lista = [];
  for (const f of lerAba(caminhoVendedores, 'Filiais')) {
    idPorNome.set(f.Filial.normalize('NFC'), f['ID Filial']);
    nome.set(f['ID Filial'], f.Filial);
    lista.push({ Id: f['ID Filial'], Nome: f.Filial });
  }
  return { idPorNome, nome, lista };
}

// alteracoesAdicionais: alterações ainda não gravadas, aplicadas depois do registro (só para validar antes de confirmar, §11.1).
function carregarBaseComercial(dirDados, dirImportacoes, vendasAdicionais, alteracoesAdicionais) {
  const erros = [];
  const caminhos = {};
  for (const [chave, nome] of Object.entries(ARQUIVOS_BASE)) {
    const p = path.join(dirDados, nome);
    if (!existe(p)) erros.push(`Arquivo não encontrado: ${p}`);
    caminhos[chave] = p;
  }
  if (erros.length) return { Erros: erros };

  // Filiais: §0.2 — nome vira ID Filial; nome desconhecido é erro.
  const fil = mapaFiliais(caminhos.Vendedores);

  // §11.4: alterações confirmadas por cima do cadastro e das metas (só com as importações, como o painel).
  const alteracoes = [...lerAlteracoes(dirImportacoes), ...(alteracoesAdicionais || [])];

  const vendedores = [];
  const vendedorPorId = new Map();
  for (const r of mesclarAlteracoes(lerAba(caminhos.Vendedores, 'Vendedores'), alteracoes, 'vendedor', (x) => x['ID Vendedor'], [], erros)) {
    const onde = `vendedores.xlsx, linha ${r._Linha}`;
    const id = r['ID Vendedor'];
    if (!id) { erros.push(`${onde}: ID Vendedor vazio`); continue; }
    if (vendedorPorId.has(id)) { erros.push(`${onde}: ID Vendedor ${id} duplicado`); continue; }
    const idFilial = r.Filial ? fil.idPorNome.get(r.Filial.normalize('NFC')) : undefined;
    if (!idFilial) erros.push(`${onde}: filial desconhecida '${r.Filial ?? ''}'`);
    if (!N.contem(['Ativo', 'Inativo'], r.Status)) erros.push(`${onde}: Status inválido '${r.Status ?? ''}'`);
    const admissao = N.lerData(r['Data de Admissão']);
    const desligamento = r['Data de Desligamento'] ? N.lerData(r['Data de Desligamento']) : null;
    if (N.igual(r.Status, 'Inativo') && desligamento === null) erros.push(`${onde}: vendedor Inativo sem Data de Desligamento válida`);
    const v = {
      Id: id, Nome: r.Nome, Email: r['E-mail'], IdFilial: idFilial ?? null, Filial: fil.nome.get(idFilial) ?? null,
      Status: r.Status, Ativo: N.igual(r.Status, 'Ativo'), Admissao: admissao, Desligamento: desligamento,
    };
    vendedores.push(v);
    vendedorPorId.set(id, v);
  }

  // Metas: uma linha por (vendedor, ano, mês). Ausência de linha = sem meta (§3.1).
  const metas = [];
  const chavesMeta = new Set();
  for (const r of mesclarAlteracoes(lerAba(caminhos.Metas, 'Metas'), alteracoes, 'meta', chaveMeta, ['ID Vendedor', 'Vendedor', 'Ano', 'Mês', 'Meta (R$)'], erros)) {
    const onde = `metas_2026.xlsx, linha ${r._Linha}`;
    const id = r['ID Vendedor'];
    const ano = N.inteiro(r.Ano);
    const mes = N.inteiro(r['Mês']);
    const valor = N.dinheiro(r['Meta (R$)']);
    if (!vendedorPorId.has(id)) { erros.push(`${onde}: ID Vendedor '${id ?? ''}' não existe no cadastro`); continue; }
    if (ano === null || mes === null || mes < 1 || mes > 12) { erros.push(`${onde}: Ano/Mês inválido`); continue; }
    if (valor === null || valor < 0) { erros.push(`${onde}: Meta inválida '${r['Meta (R$)'] ?? ''}'`); continue; }
    const chave = `${id}|${ano}|${mes}`;
    if (chavesMeta.has(chave)) { erros.push(`${onde}: meta duplicada para ${id} em ${mes}/${ano}`); continue; }
    chavesMeta.add(chave);
    metas.push({ IdVendedor: id, Ano: ano, Mes: mes, Valor: valor });
  }

  // Vendas: a base e, na ordem, cada arquivo importado. §1 e §9.8: upsert por ID Venda — a última versão vale.
  const fontes = [{ caminho: caminhos.Vendas, nome: ARQUIVOS_BASE.Vendas }, ...arquivosVendasImportados(dirImportacoes)];
  for (const c of vendasAdicionais || []) if (c) fontes.push({ caminho: c, nome: path.basename(c) });
  const vendaPorId = new Map();
  for (const fonte of fontes) {
    const idsVenda = new Set();   // ID repetido dentro do MESMO arquivo é erro (§9.4)
    for (const r of lerAba(fonte.caminho, 'Vendas')) {
      const onde = `${fonte.nome}, linha ${r._Linha}`;
      const id = r['ID Venda'];
      if (!id) { erros.push(`${onde}: ID Venda vazio`); continue; }
      if (idsVenda.has(id)) { erros.push(`${onde}: ID Venda ${id} duplicado`); continue; }
      idsVenda.add(id);
      const data = N.lerData(r.Data);
      const valor = N.dinheiro(r['Valor Total']);
      const vend = vendedorPorId.get(r['ID Vendedor']);
      const idFilial = r.Filial ? fil.idPorNome.get(r.Filial.normalize('NFC')) : undefined;
      let falhou = false;
      if (data === null) { erros.push(`${onde} (${id}): Data inválida '${r.Data ?? ''}'`); falhou = true; }
      if (valor === null) { erros.push(`${onde} (${id}): Valor Total inválido '${r['Valor Total'] ?? ''}'`); falhou = true; }
      if (!vend) { erros.push(`${onde} (${id}): ID Vendedor '${r['ID Vendedor'] ?? ''}' não existe no cadastro`); falhou = true; }
      if (!idFilial) { erros.push(`${onde} (${id}): filial desconhecida '${r.Filial ?? ''}'`); falhou = true; }
      if (!N.contem(['Faturada', 'Cancelada'], r.Status)) { erros.push(`${onde} (${id}): Status inválido '${r.Status ?? ''}'`); falhou = true; }
      if (falhou) continue;
      // §2.6: venda posterior ao desligamento é erro de importação.
      if (vend.Desligamento !== null && data > vend.Desligamento) {
        erros.push(`${onde} (${id}): venda em ${N.fmtData(data)} de vendedor desligado em ${N.fmtData(vend.Desligamento)}`);
        continue;
      }
      vendaPorId.set(id, {   // ID já existente mantém a posição, como no [ordered] do PowerShell
        Id: id, Data: data, IdVendedor: vend.Id, IdFilial: idFilial, Cliente: r.Cliente, IdCliente: r['ID Cliente'],
        Produto: r.Produto, IdProduto: r['ID Produto'], Categoria: r.Categoria, Quantidade: r.Quantidade, Valor: valor,
        Status: r.Status, IdOportunidade: r['ID Oportunidade'], Linha: r._Linha, Arquivo: fonte.nome,
      });
    }
  }
  const vendas = [...vendaPorId.values()];

  let dataBase = null;
  for (const v of vendas) if (dataBase === null || v.Data > dataBase) dataBase = v.Data;
  const arquivos = Object.keys(ARQUIVOS_BASE).map((k) => ({ Nome: path.basename(caminhos[k]), Modificado: fs.statSync(caminhos[k]).mtime }));
  for (const f of arquivosVendasImportados(dirImportacoes)) arquivos.push({ Nome: f.nome, Modificado: fs.statSync(f.caminho).mtime });
  if (lerAlteracoes(dirImportacoes).length) arquivos.push({ Nome: 'alteracoes.csv', Modificado: fs.statSync(arquivoAlteracoes(dirImportacoes)).mtime });

  return { Erros: erros, Vendedores: vendedores, VendedorPorId: vendedorPorId, Metas: metas, Vendas: vendas, DataBase: dataBase, Arquivos: arquivos };
}

function vendaRealizada(v) { return N.igual(v.Status, 'Faturada'); }   // §1: só Status = Faturada

function mesesDisponiveis(base) {
  // Metas são mensais (§4): de janeiro até o mês da data-base de vendas.
  if (base.DataBase === null || N.anoDe(base.DataBase) !== ANO_METAS) return [];
  const saida = [];
  for (let m = 1; m <= N.mesDe(base.DataBase); m++) saida.push(m);
  return saida;
}

function parcial(mesFinal, db) { return mesFinal === N.mesDe(db) && N.mesDe(db + 1) === N.mesDe(db); }   // §4.5

function desempenho(base, mesInicial, mesFinal) {
  const noPeriodo = (m) => m >= mesInicial && m <= mesFinal;
  const metasPorVendedor = new Map();
  for (const m of base.Metas) {
    if (m.Ano !== ANO_METAS || !noPeriodo(m.Mes)) continue;
    if (!metasPorVendedor.has(m.IdVendedor)) metasPorVendedor.set(m.IdVendedor, new Map());
    metasPorVendedor.get(m.IdVendedor).set(m.Mes, m.Valor);
  }
  const realizadas = new Map();
  let qtdCanceladas = 0, valorCanceladas = 0;
  for (const v of base.Vendas) {
    if (N.anoDe(v.Data) !== ANO_METAS || !noPeriodo(N.mesDe(v.Data))) continue;
    if (!vendaRealizada(v)) { qtdCanceladas++; valorCanceladas += v.Valor; continue; }
    if (!realizadas.has(v.IdVendedor)) realizadas.set(v.IdVendedor, []);
    realizadas.get(v.IdVendedor).push(v);
  }

  const linhas = base.Vendedores.map((vend) => {
    const metasMes = metasPorVendedor.get(vend.Id) || new Map();
    let meta = 0;
    for (const x of metasMes.values()) meta += x;
    let realizado = 0, realizadoMesesMeta = 0, qtd = 0;
    for (const venda of realizadas.get(vend.Id) || []) {
      realizado += venda.Valor; qtd++;
      if (metasMes.has(N.mesDe(venda.Data))) realizadoMesesMeta += venda.Valor;
    }
    return {
      Vendedor: vend,
      MesesComMeta: [...metasMes.keys()].sort((a, b) => a - b),
      Meta: meta,
      Realizado: realizado,                   // faturamento no período
      RealizadoMesesMeta: realizadoMesesMeta, // numerador do atingimento (§4.1)
      ForaMesesMeta: realizado - realizadoMesesMeta,
      QtdVendas: qtd,
      Atingimento: meta > 0 ? N.fracao(realizadoMesesMeta, meta) : null,   // §4.1: sem meta => "sem meta"
    };
  });

  // §4.2: empresa = soma ÷ soma, com inativos.
  let tr = 0, trm = 0, tm = 0, tq = 0;
  for (const l of linhas) { tr += l.Realizado; trm += l.RealizadoMesesMeta; tm += l.Meta; tq += l.QtdVendas; }
  const db = base.DataBase;
  return {
    MesInicial: mesInicial, MesFinal: mesFinal, DataBase: db, Parcial: parcial(mesFinal, db), Linhas: linhas,
    Empresa: {
      Realizado: tr, RealizadoMesesMeta: trm, ForaMesesMeta: tr - trm, Meta: tm, QtdVendas: tq,
      TicketMedio: tq ? Number(N.divArred(tr, tq)) : null,
      Atingimento: tm > 0 ? N.fracao(trm, tm) : null,
    },
    QtdCanceladas: qtdCanceladas, ValorCanceladas: valorCanceladas,
  };
}

function detalheVendedor(base, idVendedor, mesInicial, mesFinal) {
  const metaMes = new Map();
  for (const m of base.Metas) if (N.igual(m.IdVendedor, idVendedor) && m.Ano === ANO_METAS) metaMes.set(m.Mes, m.Valor);
  const vendas = N.ordenar(base.Vendas.filter((v) => N.igual(v.IdVendedor, idVendedor) && N.anoDe(v.Data) === ANO_METAS &&
    N.mesDe(v.Data) >= mesInicial && N.mesDe(v.Data) <= mesFinal), (v) => v.Data, (v) => v.Id);
  const meses = [];
  for (let mes = mesInicial; mes <= mesFinal; mes++) {
    let realizado = 0, qtd = 0, qtdCanc = 0, valorCanc = 0;
    for (const v of vendas) {
      if (N.mesDe(v.Data) !== mes) continue;
      if (vendaRealizada(v)) { realizado += v.Valor; qtd++; } else { qtdCanc++; valorCanc += v.Valor; }
    }
    const meta = metaMes.has(mes) ? metaMes.get(mes) : null;
    meses.push({ Mes: mes, Meta: meta, Realizado: realizado, QtdVendas: qtd, QtdCanceladas: qtdCanc, ValorCanceladas: valorCanc,
      Atingimento: meta > 0 ? N.fracao(realizado, meta) : null });
  }
  return { Meses: meses, Vendas: vendas };
}

// §2.3: rankings mostram só ativos, salvo pedido explícito.
function rankingFaturamento(d, incluirDesligados) {
  return N.ordenar(d.Linhas.filter((l) => incluirDesligados || l.Vendedor.Ativo), [(l) => l.Realizado, true], (l) => l.Vendedor.Id);
}
function rankingAtingimento(d, incluirDesligados) {
  return N.ordenar(d.Linhas.filter((l) => (incluirDesligados || l.Vendedor.Ativo) && l.Atingimento !== null),
    [(l) => N.fracaoParaNumero(l.Atingimento), true], (l) => l.Vendedor.Id);
}

module.exports = {
  ANO_METAS, COLUNAS_ALTERACOES, COLUNAS_PRODUTOS, arquivoAlteracoes, assinaturaAlteracoes, lerAlteracoes, camposCsv, chaveMeta, mesclarAlteracoes,
  assinaturaBase, assinaturaArquivo, carregarBaseComercial, mapaFiliais, vendaRealizada, mesesDisponiveis,
  parcial, desempenho, detalheVendedor, rankingFaturamento, rankingAtingimento, existe,
};
