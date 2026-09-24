'use strict';
// Autenticação — mesmas senhas (PBKDF2-SHA256) e regras de app/lib/Auth.ps1, com uma diferença:
// na Vercel cada visita pode cair numa instância nova, sem memória. Por isso a sessão não fica
// guardada no servidor: o cookie leva login e validade, assinados com HMAC (HORIZONTE_SEGREDO).

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const DURACAO_SESSAO_MS = 8 * 3600 * 1000;
const MAX_FALHAS = 5;
const DURACAO_BLOQUEIO_MS = 5 * 60 * 1000;

// Contas públicas: a senha é divulgada de propósito (conta "teste" dos alunos). Só o hash fica no repositório.
const ARQUIVO_PUBLICAS = path.join(__dirname, '..', 'contas_publicas.json');

// Sem HORIZONTE_SEGREDO, a assinatura usa uma chave fixa e conhecida e SÓ as contas públicas entram:
// forjar um cookie não daria mais acesso do que a senha divulgada já dá. Contas com senha de verdade
// (HORIZONTE_USUARIOS ou HORIZONTE_USUARIOS_ARQUIVO) exigem o segredo configurado.
const COM_SEGREDO = !!process.env.HORIZONTE_SEGREDO;
const segredo = process.env.HORIZONTE_SEGREDO || 'painel-horizonte:somente-contas-publicas';
if (!COM_SEGREDO && (process.env.HORIZONTE_USUARIOS || process.env.HORIZONTE_USUARIOS_ARQUIVO)) {
  console.warn('HORIZONTE_SEGREDO não definido: só as contas públicas (web/contas_publicas.json) podem entrar.');
}

function lerJson(texto) { const v = JSON.parse(texto.replace(/^﻿/, '')); return Array.isArray(v) ? v : [v]; }

function usuarios() {
  const publicas = lerJson(fs.readFileSync(ARQUIVO_PUBLICAS, 'utf8')).map((u) => ({ ...u, publica: true }));
  let outras = [];
  if (!COM_SEGREDO) outras = [];
  else if (process.env.HORIZONTE_USUARIOS) outras = lerJson(process.env.HORIZONTE_USUARIOS);
  else if (process.env.HORIZONTE_USUARIOS_ARQUIVO) outras = lerJson(fs.readFileSync(process.env.HORIZONTE_USUARIOS_ARQUIVO, 'utf8'));
  return [...outras, ...publicas];
}

function acharUsuario(login) {
  const l = String(login).toLowerCase();
  return usuarios().find((u) => String(u.login).toLowerCase() === l) || null;
}

function testarSenha(u, senha) {
  const esperado = Buffer.from(u.hash, 'base64');
  const obtido = crypto.pbkdf2Sync(senha, Buffer.from(u.sal, 'base64'), parseInt(u.iteracoes, 10), 32, 'sha256');
  return esperado.length === obtido.length && crypto.timingSafeEqual(esperado, obtido);
}

function novoRegistroSenha(senha, iteracoes = 120000) {
  const sal = crypto.randomBytes(16);
  return { iteracoes, sal: sal.toString('base64'), hash: crypto.pbkdf2Sync(senha, sal, iteracoes, 32, 'sha256').toString('base64') };
}

// Bloqueio por tentativas: por instância (na Vercel é o melhor possível sem banco de dados).
// Contas públicas não bloqueiam: a senha delas é conhecida e o bloqueio travaria a turma inteira.
const falhas = new Map();
function bloqueado(login) { const f = falhas.get(login); return !!(f && f.ate && f.ate > Date.now()); }
function registrarFalha(login) {
  const u = acharUsuario(login);
  if (u && u.publica) return;
  let f = falhas.get(login);
  if (!f || (f.ate && f.ate <= Date.now())) f = { qtd: 0, ate: null };
  f.qtd++;
  if (f.qtd >= MAX_FALHAS) f.ate = Date.now() + DURACAO_BLOQUEIO_MS;
  falhas.set(login, f);
}

const b64 = (s) => Buffer.from(s, 'utf8').toString('base64url');
function assinar(texto) { return crypto.createHmac('sha256', segredo).update(texto).digest('base64url'); }

function novaSessao(u) {
  falhas.delete(String(u.login).toLowerCase());
  const corpo = `${b64(u.login)}.${Date.now() + DURACAO_SESSAO_MS}`;
  return `${corpo}.${assinar(corpo)}`;
}

function sessao(token) {
  if (!token) return null;
  const partes = token.split('.');
  if (partes.length !== 3) return null;
  const corpo = `${partes[0]}.${partes[1]}`;
  const esperado = Buffer.from(assinar(corpo));
  const recebido = Buffer.from(partes[2]);
  if (esperado.length !== recebido.length || !crypto.timingSafeEqual(esperado, recebido)) return null;
  if (!(Number(partes[1]) > Date.now())) return null;
  const u = acharUsuario(Buffer.from(partes[0], 'base64url').toString('utf8'));
  return u ? { Usuario: u } : null;
}

module.exports = { acharUsuario, testarSenha, novoRegistroSenha, bloqueado, registrarFalha, novaSessao, sessao };
