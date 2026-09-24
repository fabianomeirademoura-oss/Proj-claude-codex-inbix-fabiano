'use strict';
// Números, datas e textos com o mesmo resultado da versão PowerShell.
// Dinheiro é inteiro em centavos e frações são pares de BigInt: nada passa por
// ponto flutuante nas contas de negócio (REGRAS_NEGOCIO.md §0.4). O arredondamento
// é sempre "metade para longe do zero", como o decimal do .NET.

// ---------- decimais ----------

const RE_DECIMAL = /^\s*([+-]?)(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?\s*$/;

// Texto invariante ("8891.280000000001", "1E-3") -> { n: BigInt, escala } exato, ou null.
function lerDecimal(texto) {
  if (texto === null || texto === undefined) return null;
  const m = RE_DECIMAL.exec(String(texto));
  if (!m) return null;
  const inteira = m[2] || '';
  const fracao = m[3] || '';
  if (!inteira && !fracao) return null;
  let n = BigInt(inteira + fracao);
  let escala = fracao.length - (m[4] ? parseInt(m[4], 10) : 0);
  if (escala < 0) { n *= 10n ** BigInt(-escala); escala = 0; }
  if (m[1] === '-') n = -n;
  return { n, escala };
}

// num ÷ den arredondado para inteiro, metade para longe do zero.
function divArred(num, den) {
  num = BigInt(num); den = BigInt(den);
  if (den < 0n) { num = -num; den = -den; }
  const neg = num < 0n;
  if (neg) num = -num;
  let q = num / den;
  if ((num % den) * 2n >= den) q += 1n;
  return neg ? -q : q;
}

// Decimal -> BigInt com `casas` casas decimais implícitas (arredondado).
function paraEscala(dec, casas) {
  if (dec.escala <= casas) return dec.n * 10n ** BigInt(casas - dec.escala);
  return divArred(dec.n, 10n ** BigInt(dec.escala - casas));
}

// ConvertTo-Dinheiro: centavos (Number inteiro) ou null.
function dinheiro(texto) {
  const d = lerDecimal(texto);
  return d === null ? null : Number(paraEscala(d, 2));
}

// ConvertTo-Inteiro: só aceita valor sem parte fracionária.
function inteiro(texto) {
  const d = lerDecimal(texto);
  if (d === null) return null;
  const p = 10n ** BigInt(d.escala);
  if (d.n % p !== 0n) return null;
  const v = Number(d.n / p);
  return v >= -2147483648 && v <= 2147483647 ? v : null;
}

// Fração exata a/b (b > 0), usada em atingimento, participação e conversão.
function fracao(a, b) { return { n: BigInt(a), d: BigInt(b) }; }
function fracaoParaNumero(f) { return Number(f.n) / Number(f.d); }

// ---------- formatação pt-BR (N2, N1, N0, 0.00) ----------

function formatarEscalado(v, casas, agrupar) {
  v = BigInt(v);
  const neg = v < 0n;
  if (neg) v = -v;
  const s = v.toString().padStart(casas + 1, '0');
  let int = casas ? s.slice(0, -casas) : s;
  const frac = casas ? s.slice(-casas) : '';
  if (agrupar) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + int + (casas ? ',' + frac : '');
}

// Format-Moeda
function moeda(centavos) {
  if (centavos === null || centavos === undefined) return '—';
  return 'R$ ' + formatarEscalado(centavos, 2, true);
}

// Format-Pct: fração -> "12,3%"; sem fração -> "sem meta".
function pct(f) {
  if (f === null || f === undefined) return 'sem meta';
  return formatarEscalado(divArred(f.n * 1000n, f.d), 1, true) + '%';
}

// CSV: valor com 2 casas e vírgula, sem separador de milhar ('0.00' pt-BR).
function csvDinheiro(centavos) {
  if (centavos === null || centavos === undefined) return '';
  return formatarEscalado(centavos, 2, false);
}
function csvPct(f) {
  if (f === null || f === undefined) return 'sem meta';
  return formatarEscalado(divArred(f.n * 10000n, f.d), 2, false);
}

// Probabilidade (decimal) -> "25%" (Format-Prob) e "0,25" (CSV).
function prob(p) { return formatarEscalado(divArred(p.n * 100n, 10n ** BigInt(p.escala)), 0, true) + '%'; }
function csvProb(p) { return formatarEscalado(paraEscala(p, 2), 2, false); }
function probParaNumero(p) { return Number(p.n) / 10 ** p.escala; }

// Número de ponto flutuante com formato personalizado do .NET ('0.#', '0.0000'):
// 15 algarismos significativos, depois arredonda para as casas pedidas. Ponto decimal.
function formatarDouble(x, minCasas, maxCasas) {
  const dec = lerDecimal(Math.abs(x).toPrecision(15));
  let s = paraEscala(dec, maxCasas).toString().padStart(maxCasas + 1, '0');
  let int = maxCasas ? s.slice(0, -maxCasas) : s;
  let frac = maxCasas ? s.slice(-maxCasas) : '';
  while (frac.length > minCasas && frac.endsWith('0')) frac = frac.slice(0, -1);
  return (x < 0 && /[1-9]/.test(int + frac) ? '-' : '') + int + (frac ? '.' + frac : '');
}

// ---------- datas (dia = número de dias desde 1970-01-01, sem fuso) ----------

const MS_DIA = 86400000;
function dia(ano, mes, d) { return Date.UTC(ano, mes - 1, d) / MS_DIA; }
function partesData(n) { const t = new Date(n * MS_DIA); return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()]; }
function anoDe(n) { return partesData(n)[0]; }
function mesDe(n) { return partesData(n)[1]; }

// ConvertTo-Data: 'yyyy-MM-dd' ou 'dd/MM/yyyy', data de calendário válida.
function lerData(texto) {
  if (texto === null || texto === undefined) return null;
  const s = String(texto);
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  let a, me, d;
  if (m) { a = +m[1]; me = +m[2]; d = +m[3]; } else {
    m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
    if (!m) return null;
    d = +m[1]; me = +m[2]; a = +m[3];
  }
  if (a < 1 || me < 1 || me > 12 || d < 1) return null;
  const n = dia(a, me, d);
  const [a2, me2, d2] = partesData(n);
  return a2 === a && me2 === me && d2 === d ? n : null;
}

const d2 = (x) => String(x).padStart(2, '0');
function fmtData(n) {
  if (n === null || n === undefined) return '';
  const [a, m, d] = partesData(n);
  return `${d2(d)}/${d2(m)}/${a}`;
}
function fmtDataIso(n) { const [a, m, d] = partesData(n); return `${a}-${d2(m)}-${d2(d)}`; }
function fmtMesAno(n) { const [a, m] = partesData(n); return `${d2(m)}/${a}`; }

// Data e hora de modificação de arquivo, no fuso do painel (a versão PowerShell usa a hora local).
function fmtDataHora(data, fuso) {
  const p = {};
  for (const x of new Intl.DateTimeFormat('en-GB', { timeZone: fuso, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }).formatToParts(data)) p[x.type] = x.value;
  return `${p.day}/${p.month}/${p.year} ${p.hour}:${p.minute}`;
}

// ---------- textos ----------

// WebUtility.HtmlEncode: também codifica os caracteres 160–255 (acentos) e os fora do BMP como &#N;.
function esc(texto) {
  if (texto === null || texto === undefined) return '';
  let s = '';
  for (const ch of String(texto)) {
    const c = ch.codePointAt(0);
    if (ch === '<') s += '&lt;';
    else if (ch === '>') s += '&gt;';
    else if (ch === '&') s += '&amp;';
    else if (ch === '"') s += '&quot;';
    else if (ch === "'") s += '&#39;';
    else if ((c >= 160 && c <= 255) || c > 0xffff) s += `&#${c};`;
    else s += ch;
  }
  return s;
}

// Uri.EscapeDataString: só letras, dígitos e -._~ ficam como estão.
function escaparUri(texto) {
  return encodeURIComponent(String(texto)).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());
}

// Comparações do PowerShell: -eq e -in não diferenciam maiúsculas; ordenação de texto pela cultura.
function igual(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return a === b || (a == null && b == null);
  return String(a).toLowerCase() === String(b).toLowerCase();
}
function contem(lista, v) { return lista.some((x) => igual(x, v)); }

const COLACAO = new Intl.Collator('pt-BR', { sensitivity: 'accent' });
function compararValores(a, b) {
  if (a === b) return 0;
  if (a === null || a === undefined) return -1;
  if (b === null || b === undefined) return 1;
  if (typeof a === 'string' || typeof b === 'string') return COLACAO.compare(String(a), String(b));
  if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

// Sort-Object com várias chaves: ordenar(lista, [fn, desc], [fn], ...). Estável.
function ordenar(lista, ...chaves) {
  const ks = chaves.map((k) => (typeof k === 'function' ? [k, false] : k));
  return [...lista].sort((x, y) => {
    for (const [f, desc] of ks) {
      const c = compararValores(f(x), f(y));
      if (c) return desc ? -c : c;
    }
    return 0;
  });
}

// Sort-Object -Unique de textos (sem diferenciar maiúsculas).
function unicosOrdenados(lista) {
  const ord = ordenar(lista.filter((x) => x !== null && x !== undefined), (x) => x);
  const saida = [];
  for (const x of ord) if (!saida.length || COLACAO.compare(saida[saida.length - 1], x) !== 0) saida.push(x);
  return saida;
}

module.exports = {
  lerDecimal, divArred, paraEscala, dinheiro, inteiro, fracao, fracaoParaNumero,
  formatarEscalado, moeda, pct, csvDinheiro, csvPct, prob, csvProb, probParaNumero, formatarDouble,
  dia, partesData, anoDe, mesDe, lerData, fmtData, fmtDataIso, fmtMesAno, fmtDataHora,
  esc, escaparUri, igual, contem, ordenar, unicosOrdenados, compararValores,
};
