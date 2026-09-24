'use strict';
// Sobe a versão Node do painel na própria máquina (para testar antes de publicar).
// Uso: node web/servidor.js [porta]   — contas extras em HORIZONTE_USUARIOS_ARQUIVO (formato de app/usuarios.json)

const http = require('http');
const handler = require('./app');

const porta = parseInt(process.argv[2] || process.env.PORT || '8090', 10);
http.createServer(handler).listen(porta, '127.0.0.1', () => {
  console.log(`Painel (Node) no ar: http://localhost:${porta}/   (Ctrl+C para parar)`);
});
