'use strict';
// Origem do faturamento — espelho de app/lib/RegrasOrigem.ps1 (REGRAS_NEGOCIO.md §1 e §0.1).
// Uma linha por venda, ligada exclusivamente pelo ID Oportunidade.

const N = require('./numeros');
const R = require('./regras');

function visaoOrigem(base, pipeline, mesInicial, mesFinal) {
  if (base.Erros.length || pipeline.Erros.length) return { Erros: [...base.Erros, ...pipeline.Erros] };
  const erros = [];
  const porId = new Map();
  for (const o of pipeline.Oportunidades) {
    if (porId.has(o.Id)) { erros.push(`ID Oportunidade duplicado no CRM: ${o.Id}`); continue; }
    porId.set(o.Id, o);
  }
  let com = 0, sem = 0, cancelado = 0, qCom = 0, qSem = 0, qCancelado = 0;
  const linhas = [];
  const categorias = new Map();
  for (const v of N.ordenar(base.Vendas, (x) => x.Data, (x) => x.Id)) {
    const mes = N.mesDe(v.Data);
    if (N.anoDe(v.Data) !== R.ANO_METAS || mes < mesInicial || mes > mesFinal) continue;
    let o = null;
    if (v.IdOportunidade) {
      if (!porId.has(v.IdOportunidade)) erros.push(`Venda ${v.Id}, linha ${v.Linha}: ID Oportunidade '${v.IdOportunidade}' não encontrado no CRM.`);
      else o = porId.get(v.IdOportunidade);
    }
    const faturada = R.vendaRealizada(v);
    linhas.push({ Venda: v, Oportunidade: o, Faturada: faturada });
    if (!faturada) { qCancelado++; cancelado += v.Valor; continue; }
    if (o) { qCom++; com += v.Valor; } else if (!v.IdOportunidade) {
      qSem++; sem += v.Valor;
      const c = v.Categoria ?? '';
      if (!categorias.has(c)) categorias.set(c, { Categoria: c, Quantidade: 0, Valor: 0 });
      categorias.get(c).Quantidade++; categorias.get(c).Valor += v.Valor;
    }
  }
  // Não publicar totais incompletos nem tratar ID inválido como venda de balcão.
  if (erros.length) return { Erros: erros };
  const total = com + sem;
  return {
    Erros: [], MesInicial: mesInicial, MesFinal: mesFinal, DataBase: base.DataBase, DataBaseCrm: pipeline.DataBase,
    Total: total, ComCrm: com, SemOportunidade: sem, QtdComCrm: qCom, QtdSemOportunidade: qSem, QtdFaturadas: qCom + qSem,
    FracaoCrm: total > 0 ? N.fracao(com, total) : null, FracaoSem: total > 0 ? N.fracao(sem, total) : null,
    QtdCanceladas: qCancelado, ValorCancelado: cancelado, Linhas: linhas,
    CategoriasSem: N.ordenar([...categorias.values()], [(c) => c.Valor, true], (c) => c.Categoria),
    Parcial: R.parcial(mesFinal, base.DataBase),
  };
}

module.exports = { visaoOrigem };
