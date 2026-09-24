#!/usr/bin/env node
'use strict';
// Servidor MCP do painel da Horizonte Máquinas, só leitura, para o Claude Desktop.
// Fala o protocolo MCP pela entrada e saída padrão (stdio): uma mensagem JSON-RPC 2.0 por linha.
// Sem dependências, como o resto da versão Node. As ferramentas estão em ferramentas.js.
//
// Uso: node mcp/servidor.js            (lê as planilhas de dados/, ao lado desta pasta)
//      HORIZONTE_DADOS=/caminho/dados node mcp/servidor.js
// A saída padrão é só do protocolo: mensagens para gente vão para a saída de erro (o log do Claude Desktop).

const path = require('path');
const readline = require('readline');
const { FERRAMENTAS, criarFonte, executar } = require('./ferramentas');

const VERSOES = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05'];   // da mais nova para a mais antiga
const DIR_DADOS = path.resolve(process.env.HORIZONTE_DADOS || path.join(__dirname, '..', 'dados'));
const fonte = criarFonte(DIR_DADOS);

const INSTRUCOES = [
  'Dados comerciais da Horizonte Máquinas (vendas e metas de 2026, CRM, estoque), somente leitura.',
  'As regras de negócio estão em REGRAS_NEGOCIO.md no repositório; as respostas citam a seção (ex.: §4.1).',
  'Valores em reais com 2 casas; atingimento como fração e como texto (89,3%).',
  'Toda resposta traz a data-base que usou: os cálculos nunca usam a data de hoje (§0.3).',
].join(' ');

function log(msg) { process.stderr.write(`[horizonte-mcp] ${msg}\n`); }
function enviar(msg) { process.stdout.write(JSON.stringify(msg) + '\n'); }
function responder(id, result) { enviar({ jsonrpc: '2.0', id, result }); }
function falhar(id, code, message) { enviar({ jsonrpc: '2.0', id, error: { code, message } }); }

function tratar(msg) {
  const { id, method, params } = msg;
  const notificacao = id === undefined || id === null;
  switch (method) {
    case 'initialize': {
      const pedida = params && params.protocolVersion;
      return responder(id, {
        protocolVersion: VERSOES.includes(pedida) ? pedida : VERSOES[0],
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: 'horizonte-maquinas', title: 'Painel Horizonte Máquinas (leitura)', version: '1.0.0' },
        instructions: INSTRUCOES,
      });
    }
    case 'ping':
      return responder(id, {});
    case 'tools/list':
      return responder(id, { tools: FERRAMENTAS.map(({ name, title, description, inputSchema, annotations }) => ({ name, title, description, inputSchema, annotations })) });
    case 'tools/call': {
      const nome = params && params.name;
      if (!FERRAMENTAS.some((f) => f.name === nome)) return falhar(id, -32602, `Ferramenta desconhecida: ${nome}`);
      const r = executar(fonte, nome, params.arguments);
      const corpo = r.ok ? r.resultado : { erro: r.erro, ...(r.detalhe || {}) };
      return responder(id, { content: [{ type: 'text', text: JSON.stringify(corpo, null, 2) }], isError: !r.ok });
    }
    default:
      if (notificacao) return undefined;   // notifications/initialized, notifications/cancelled etc.: nada a responder
      return falhar(id, -32601, `Método não suportado: ${method}`);
  }
}

const entrada = readline.createInterface({ input: process.stdin, terminal: false });
entrada.on('line', (linha) => {
  if (!linha.trim()) return;
  let msg;
  try { msg = JSON.parse(linha); } catch { return falhar(null, -32700, 'JSON inválido'); }
  for (const m of Array.isArray(msg) ? msg : [msg]) {   // lote (versão 2025-03-26): uma resposta por linha
    try {
      if (!m || typeof m !== 'object' || typeof m.method !== 'string') {
        if (m && m.id !== undefined && m.id !== null && m.result === undefined && m.error === undefined) falhar(m.id, -32600, 'Requisição inválida');
        continue;
      }
      tratar(m);
    } catch (e) {
      log(`erro em ${m.method}: ${e.stack || e.message}`);
      if (m.id !== undefined && m.id !== null) falhar(m.id, -32603, `Erro interno: ${e.message}`);
    }
  }
});
entrada.on('close', () => process.exit(0));
log(`pronto; dados em ${DIR_DADOS}`);
