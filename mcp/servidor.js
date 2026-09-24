#!/usr/bin/env node
'use strict';
// Servidor MCP do painel da Horizonte Máquinas, para o Claude Desktop.
// Fala o protocolo MCP pela entrada e saída padrão (stdio): uma mensagem JSON-RPC 2.0 por linha.
// Sem dependências, como o resto da versão Node. Leitura em ferramentas.js; escrita em escrita.js (§11).
//
// Uso: node mcp/servidor.js            (lê as planilhas de dados/, ao lado desta pasta)
// Variáveis de ambiente:
//   HORIZONTE_DADOS              pasta das planilhas (padrão: dados/ ao lado desta pasta)
//   HORIZONTE_USUARIO            login da conta que escreve (gerente ou diretoria). Sem ele, só leitura.
//   HORIZONTE_ARQUIVO_USUARIOS   contas do painel (padrão: app/usuarios.json ao lado desta pasta)
// A saída padrão é só do protocolo: mensagens para gente vão para a saída de erro (o log do Claude Desktop).

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { FERRAMENTAS, criarFonte, executar } = require('./ferramentas');
const { FERRAMENTAS_ESCRITA, criarIdentidade } = require('./escrita');

const VERSOES = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05'];   // da mais nova para a mais antiga
const DIR_DADOS = path.resolve(process.env.HORIZONTE_DADOS || path.join(__dirname, '..', 'dados'));
const ARQUIVO_USUARIOS = path.resolve(process.env.HORIZONTE_ARQUIVO_USUARIOS || path.join(__dirname, '..', 'app', 'usuarios.json'));
const LOGIN = (process.env.HORIZONTE_USUARIO || '').trim();
const TODAS = [...FERRAMENTAS, ...FERRAMENTAS_ESCRITA];
const fonte = criarFonte(DIR_DADOS);

const INSTRUCOES = [
  'Dados comerciais da Horizonte Máquinas (vendas e metas de 2026, CRM, estoque).',
  'As regras de negócio estão em REGRAS_NEGOCIO.md no repositório; as respostas citam a seção (ex.: §4.1).',
  'Valores em reais com 2 casas; atingimento como fração e como texto (89,3%).',
  'Toda resposta traz a data-base que usou: os cálculos nunca usam a data de hoje (§0.3).',
  'ESCRITA (§11): as ferramentas inativar_vendedor, reativar_vendedor, alterar_meta, cadastrar_ou_editar_produto e transferir_oportunidades só PREPARAM e nada gravam.',
  'Mostre sempre à pessoa o resumo (valor atual → valor novo e os efeitos) e pergunte se ela confirma.',
  'Só chame confirmar_alteracao depois de um "sim" explícito dela para aquele resumo; nunca confirme por conta própria nem junte confirmações.',
  'Toda gravação fica no histórico (ver_historico e a tela Histórico do painel).',
].join(' ');

function log(msg) { process.stderr.write(`[horizonte-mcp] ${msg}\n`); }
function enviar(msg) { process.stdout.write(JSON.stringify(msg) + '\n'); }
function responder(id, result) { enviar({ jsonrpc: '2.0', id, result }); }
function falhar(id, code, message) { enviar({ jsonrpc: '2.0', id, error: { code, message } }); }

// ---------- pedidos do servidor ao cliente (elicitation) ----------

let capacidadesCliente = {};
let proximoId = 1;
const aguardando = new Map();
function pedirAoCliente(method, params) {
  const id = `srv-${proximoId++}`;
  return new Promise((ok, falha) => {
    aguardando.set(id, { ok, falha });
    enviar({ jsonrpc: '2.0', id, method, params });
  });
}

// §11.1: se o Claude Desktop sabe perguntar direto à pessoa, a confirmação sai dela, não do modelo.
async function perguntar(mensagem) {
  const r = await pedirAoCliente('elicitation/create', {
    message: mensagem,
    requestedSchema: { type: 'object', properties: { confirmar: { type: 'boolean', title: 'Confirmo a gravação', description: 'Marque para gravar esta alteração.' } }, required: ['confirmar'] },
  });
  return !!(r && r.action === 'accept' && r.content && r.content.confirmar === true);
}

// ---------- contexto das ferramentas de escrita ----------

const ctx = {
  usuario: criarIdentidade(LOGIN, ARQUIVO_USUARIOS),
  pendentes: new Map(),   // código → alteração preparada (só na memória deste processo)
  get perguntar() { return capacidadesCliente.elicitation ? perguntar : null; },
  contaDoVendedor(idVendedor) {
    try {
      const v = JSON.parse(fs.readFileSync(ARQUIVO_USUARIOS, 'utf8').replace(/^﻿/, ''));
      return (Array.isArray(v) ? v : [v]).some((u) => u.idVendedor === idVendedor);
    } catch { return false; }
  },
};

async function tratar(msg) {
  const { id, method, params } = msg;
  const notificacao = id === undefined || id === null;
  switch (method) {
    case 'initialize': {
      const pedida = params && params.protocolVersion;
      capacidadesCliente = (params && params.capabilities) || {};
      return responder(id, {
        protocolVersion: VERSOES.includes(pedida) ? pedida : VERSOES[0],
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: 'horizonte-maquinas', title: 'Painel Horizonte Máquinas', version: '2.0.0' },
        instructions: INSTRUCOES,
      });
    }
    case 'ping':
      return responder(id, {});
    case 'tools/list':
      return responder(id, { tools: TODAS.map(({ name, title, description, inputSchema, annotations }) => ({ name, title, description, inputSchema, annotations })) });
    case 'tools/call': {
      const nome = params && params.name;
      if (!TODAS.some((f) => f.name === nome)) return falhar(id, -32602, `Ferramenta desconhecida: ${nome}`);
      const r = await executar(TODAS, fonte, nome, params.arguments, ctx);
      const corpo = r.ok ? r.resultado : { erro: r.erro, ...(r.detalhe || {}) };
      return responder(id, { content: [{ type: 'text', text: JSON.stringify(corpo, null, 2) }], isError: !r.ok });
    }
    default:
      if (notificacao) return undefined;   // notifications/initialized, notifications/cancelled etc.: nada a responder
      return falhar(id, -32601, `Método não suportado: ${method}`);
  }
}

async function processar(m) {
  try {
    if (!m || typeof m !== 'object') return;
    // Resposta do cliente a um pedido nosso (elicitation).
    if (typeof m.method !== 'string' && m.id !== undefined && aguardando.has(m.id)) {
      const p = aguardando.get(m.id);
      aguardando.delete(m.id);
      if (m.error) p.falha(new Error(m.error.message)); else p.ok(m.result);
      return;
    }
    if (typeof m.method !== 'string') {
      if (m.id !== undefined && m.id !== null && m.result === undefined && m.error === undefined) falhar(m.id, -32600, 'Requisição inválida');
      return;
    }
    await tratar(m);
  } catch (e) {
    log(`erro em ${m && m.method}: ${e.stack || e.message}`);
    if (m && m.id !== undefined && m.id !== null && typeof m.method === 'string') falhar(m.id, -32603, `Erro interno: ${e.message}`);
  }
}

const entrada = readline.createInterface({ input: process.stdin, terminal: false });
let emAndamento = 0;
let fechou = false;
entrada.on('line', (linha) => {
  if (!linha.trim()) return;
  let msg;
  try { msg = JSON.parse(linha); } catch { return falhar(null, -32700, 'JSON inválido'); }
  for (const m of Array.isArray(msg) ? msg : [msg]) {   // lote (versão 2025-03-26): uma resposta por linha
    emAndamento++;
    processar(m).finally(() => { emAndamento--; if (fechou && !emAndamento) process.exit(0); });
  }
});
entrada.on('close', () => { fechou = true; if (!emAndamento) process.exit(0); });
log(`pronto; dados em ${DIR_DADOS}; ${LOGIN ? `escrita como '${LOGIN}' (contas em ${ARQUIVO_USUARIOS})` : 'só leitura (HORIZONTE_USUARIO não definido)'}`);
