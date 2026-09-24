'use strict';
// Leitor mínimo de .xlsx sem dependências, espelho de app/lib/Xlsx.ps1.
// Devolve as linhas de uma aba como objetos, com a primeira linha como cabeçalho.
// Números vêm como texto invariante; células com formato de data vêm como "yyyy-MM-dd".

const fs = require('fs');
const zlib = require('zlib');

function lerZip(caminho) {
  const buf = fs.readFileSync(caminho);
  let fim = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { fim = i; break; }
  }
  if (fim < 0) throw new Error(`Arquivo não é um .xlsx válido: ${caminho}`);
  const qtd = buf.readUInt16LE(fim + 10);
  let p = buf.readUInt32LE(fim + 16);
  const entradas = new Map();
  for (let k = 0; k < qtd; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error(`Diretório do .xlsx corrompido: ${caminho}`);
    const metodo = buf.readUInt16LE(p + 10);
    const tamComp = buf.readUInt32LE(p + 20);
    const lenNome = buf.readUInt16LE(p + 28);
    const lenExtra = buf.readUInt16LE(p + 30);
    const lenCom = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const nome = buf.toString('utf8', p + 46, p + 46 + lenNome);
    entradas.set(nome, { metodo, tamComp, local });
    p += 46 + lenNome + lenExtra + lenCom;
  }
  return {
    texto(nome) {
      const e = entradas.get(nome);
      if (!e) return null;
      const ini = e.local + 30 + buf.readUInt16LE(e.local + 26) + buf.readUInt16LE(e.local + 28);
      const dados = buf.subarray(ini, ini + e.tamComp);
      const bruto = e.metodo === 0 ? dados : zlib.inflateRawSync(dados);
      return bruto.toString('utf8').replace(/^﻿/, '');
    },
  };
}

function decodificar(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, e) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e];
  });
}

function atributos(texto) {
  const a = {};
  const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(texto))) a[m[1]] = decodificar(m[2] !== undefined ? m[2] : m[3]);
  return a;
}

// Texto de um trecho de XML (InnerText).
function textoInterno(xml) { return decodificar(xml.replace(/<[^>]*>/g, '')); }

function indiceColuna(ref) {
  let n = 0;
  for (const c of ref.replace(/\d/g, '')) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

function formatoData(id, personalizados) {
  if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
  if (personalizados.has(id)) {
    const codigo = personalizados.get(id).replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[dmyhs]/i.test(codigo);
  }
  return false;
}

// DateTime.FromOADate(v).ToString('yyyy-MM-dd')
function dataOA(v) {
  const ms = Math.trunc(v * 86400000 + (v >= 0 ? 0.5 : -0.5));
  const t = new Date(Date.UTC(1899, 11, 30) + ms);
  const d2 = (x) => String(x).padStart(2, '0');
  return `${t.getUTCFullYear()}-${d2(t.getUTCMonth() + 1)}-${d2(t.getUTCDate())}`;
}

const P = '(?:[\\w.-]+:)?';   // prefixo de namespace opcional

function lerAba(caminho, aba) {
  const zip = lerZip(caminho);

  const compartilhadas = [];
  const xmlStrings = zip.texto('xl/sharedStrings.xml');
  if (xmlStrings) {
    const reSi = new RegExp(`<${P}si(?:\\s[^>]*)?(?:/>|>([\\s\\S]*?)</${P}si>)`, 'g');
    const reT = new RegExp(`<${P}t(?:\\s[^>]*)?(?:/>|>([\\s\\S]*?)</${P}t>)`, 'g');
    const reRph = new RegExp(`<${P}rPh\\b[\\s\\S]*?</${P}rPh>`, 'g');
    let m;
    while ((m = reSi.exec(xmlStrings))) {
      const corpo = (m[1] || '').replace(reRph, '');
      let s = '', t;
      reT.lastIndex = 0;
      while ((t = reT.exec(corpo))) s += decodificar(t[1] || '');
      compartilhadas.push(s);
    }
  }

  const estilosData = new Set();
  const xmlEstilos = zip.texto('xl/styles.xml');
  if (xmlEstilos) {
    const personalizados = new Map();
    const blocoFmts = new RegExp(`<${P}numFmts\\b[\\s\\S]*?</${P}numFmts>`).exec(xmlEstilos);
    if (blocoFmts) {
      const re = new RegExp(`<${P}numFmt\\b([^>]*?)/?>`, 'g');
      let m;
      while ((m = re.exec(blocoFmts[0]))) { const a = atributos(m[1]); personalizados.set(parseInt(a.numFmtId, 10), a.formatCode || ''); }
    }
    const blocoXfs = new RegExp(`<${P}cellXfs\\b[\\s\\S]*?</${P}cellXfs>`).exec(xmlEstilos);
    if (blocoXfs) {
      const re = new RegExp(`<${P}xf\\b([^>]*?)/?>`, 'g');
      let m, i = 0;
      while ((m = re.exec(blocoXfs[0]))) {
        if (formatoData(parseInt(atributos(m[1]).numFmtId || '0', 10), personalizados)) estilosData.add(i);
        i++;
      }
    }
  }

  const xmlLivro = zip.texto('xl/workbook.xml');
  let idRel = null, achou = false;
  const reSheet = new RegExp(`<${P}sheet\\b([^>]*?)/?>`, 'g');
  let m;
  while ((m = reSheet.exec(xmlLivro))) {
    const a = atributos(m[1]);
    if (a.name === aba) { achou = true; idRel = Object.keys(a).filter((k) => /:id$/.test(k)).map((k) => a[k])[0]; }
  }
  if (!achou) throw new Error(`Aba '${aba}' não encontrada em ${caminho}`);

  const xmlRel = zip.texto('xl/_rels/workbook.xml.rels');
  let alvo = null;
  const reRel = /<(?:[\w.-]+:)?Relationship\b([^>]*?)\/?>/g;
  while ((m = reRel.exec(xmlRel))) { const a = atributos(m[1]); if (a.Id === idRel) alvo = a.Target; }
  alvo = alvo.startsWith('/') ? alvo.replace(/^\/+/, '') : 'xl/' + alvo;

  const xmlAba = zip.texto(alvo);
  let cabecalho = null;
  const linhas = [];
  const reRow = new RegExp(`<${P}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${P}row>)`, 'g');
  const reC = new RegExp(`<${P}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${P}c>)`, 'g');
  const reV = new RegExp(`<${P}v(?:\\s[^>]*)?>([\\s\\S]*?)</${P}v>`);
  const reIs = new RegExp(`<${P}is(?:\\s[^>]*)?>([\\s\\S]*?)</${P}is>`);
  let r;
  while ((r = reRow.exec(xmlAba))) {
    const numLinha = parseInt(atributos(r[1]).r, 10);
    const valores = new Map();
    let maior = -1;
    reC.lastIndex = 0;
    let c;
    while ((c = reC.exec(r[2] || ''))) {
      const a = atributos(c[1]);
      const col = indiceColuna(a.r);
      const tipo = a.t || '';
      const corpo = c[2] || '';
      const mv = reV.exec(corpo);
      const bruto = mv ? decodificar(mv[1]) : null;
      let valor;
      if (tipo === 's') valor = compartilhadas[parseInt(bruto, 10)];
      else if (tipo === 'inlineStr') { const mi = reIs.exec(corpo); valor = mi ? textoInterno(mi[1]) : ''; }
      else if (tipo === 'e') throw new Error(`Célula com erro (${bruto}) em ${aba}!${a.r} de ${caminho}`);
      else if (bruto !== null && tipo !== 'str' && a.s && estilosData.has(parseInt(a.s, 10))) valor = dataOA(parseFloat(bruto));
      else valor = bruto;
      valores.set(col, valor === undefined ? null : valor);
      if (col > maior) maior = col;
    }
    if (cabecalho === null) {
      if (maior < 0) continue;
      cabecalho = [];
      for (let k = 0; k <= maior; k++) cabecalho.push(String(valores.get(k) ?? '').trim());
      continue;
    }
    let vazia = true;
    const obj = {};
    for (let k = 0; k < cabecalho.length; k++) {
      let v = valores.has(k) ? valores.get(k) : null;
      if (v !== null) { v = String(v).trim(); if (v !== '') vazia = false; else v = null; }
      obj[cabecalho[k]] = v;
    }
    if (!vazia) { obj._Linha = numLinha; linhas.push(obj); }
  }
  return linhas;
}

module.exports = { lerAba };
