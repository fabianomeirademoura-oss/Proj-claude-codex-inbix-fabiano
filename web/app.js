'use strict';
// Rotas do painel em Node — espelho de app/servidor.ps1, para rodar na Vercel.
// A importação de planilhas (§9) fica só na versão PowerShell: a Vercel não tem disco permanente.

const path = require('path');
const N = require('./lib/numeros');
const R = require('./lib/regras');
const RP = require('./lib/regras_pipeline');
const RE = require('./lib/regras_estoque');
const RO = require('./lib/regras_origem');
const A = require('./lib/auth');
const P = require('./lib/paginas');
const PP = require('./lib/paginas_pipeline');
const PE = require('./lib/paginas_estoque');

const DIR_DADOS = process.env.HORIZONTE_DADOS || path.join(__dirname, '..', 'dados');
const DIR_IMPORTACOES = path.join(DIR_DADOS, 'importacoes');   // §9.8: importações confirmadas, se existirem
const FUSO = process.env.HORIZONTE_FUSO || 'America/Sao_Paulo';
const HTML = 'text/html; charset=utf-8';
const TEXTO = 'text/plain; charset=utf-8';

// ---------- carga com cache: recarrega quando alguma planilha muda ----------

const cache = { base: null, aBase: null, pipeline: null, aPipeline: null, estoque: null, aEstoque: null };

function atualizarBase() {
  const a = R.assinaturaBase(DIR_DADOS, DIR_IMPORTACOES);
  if (a === cache.aBase) return cache.base;
  try { cache.base = R.carregarBaseComercial(DIR_DADOS, DIR_IMPORTACOES); } catch (e) { cache.base = { Erros: [`Falha ao ler as planilhas: ${e.message}`] }; }
  cache.aBase = a;
  return cache.base;
}
function atualizarPipeline() {
  const base = atualizarBase();
  const a = `${RP.assinaturaPipeline(DIR_DADOS)}#${cache.aBase}`;
  if (a === cache.aPipeline) return cache.pipeline;
  try { cache.pipeline = RP.carregarBasePipeline(DIR_DADOS, base); } catch (e) { cache.pipeline = { Erros: [`Falha ao ler o CRM: ${e.message}`] }; }
  cache.aPipeline = a;
  return cache.pipeline;
}
function atualizarEstoque() {
  const a = RE.assinaturaEstoque(DIR_DADOS, DIR_IMPORTACOES);
  if (a === cache.aEstoque) return cache.estoque;
  try { cache.estoque = RE.carregarBaseEstoque(DIR_DADOS, DIR_IMPORTACOES); } catch (e) { cache.estoque = { Erros: [`Falha ao ler o estoque: ${e.message}`] }; }
  cache.aEstoque = a;
  return cache.estoque;
}

// ---------- HTTP ----------

function responder(res, status, tipo, corpo, extra = {}) {
  let buf = Buffer.from(corpo, 'utf8');
  if (extra.bom) buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), buf]);
  const cab = {
    'Content-Type': tipo,
    'Content-Security-Policy': "default-src 'none'; style-src 'self'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'",
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    'Cache-Control': 'no-store',
    'Content-Length': buf.length,
  };
  if (extra.disposicao) cab['Content-Disposition'] = extra.disposicao;
  res.writeHead(status, cab);
  res.end(buf);
}

function redirecionar(res, destino, cookie) {
  const cab = { Location: destino, 'Cache-Control': 'no-store', 'Content-Length': 0 };
  if (cookie) cab['Set-Cookie'] = cookie;
  res.writeHead(303, cab);
  res.end();
}

function lerCorpo(req, limite) {
  return new Promise((ok, falha) => {
    const partes = [];
    let total = 0;
    req.on('data', (c) => { total += c.length; if (total <= limite) partes.push(c); });
    req.on('end', () => ok(total > limite ? null : Buffer.concat(partes)));
    req.on('error', falha);
  });
}

function decodificarUrl(s) {
  // WebUtility.UrlDecode: "+" vira espaço; sequência inválida fica como está.
  s = s.replace(/\+/g, ' ');
  try { return decodeURIComponent(s); } catch { return s.replace(/(%[0-9a-fA-F]{2})+/g, (m) => { try { return decodeURIComponent(m); } catch { return m; } }); }
}

async function lerFormulario(req) {
  const campos = {};
  if (Number(req.headers['content-length'] || 0) > 8192) return campos;
  const corpo = await lerCorpo(req, 8192);
  if (!corpo) return campos;
  for (const par of corpo.toString('utf8').split('&')) {
    if (!par) continue;
    const i = par.indexOf('=');
    const k = i < 0 ? par : par.slice(0, i);
    const v = i < 0 ? '' : par.slice(i + 1);
    campos[decodificarUrl(k)] = decodificarUrl(v);
  }
  return campos;
}

function cookie(req, nome) {
  for (const parte of String(req.headers.cookie || '').split(';')) {
    const i = parte.indexOf('=');
    if (i > 0 && parte.slice(0, i).trim() === nome) return parte.slice(i + 1).trim();
  }
  return null;
}

// HttpListener junta valores repetidos da query com vírgula.
function query(url) { const qs = url.searchParams; return (k) => { const v = qs.getAll(k); return v.length ? v.join(',') : null; }; }

// [int]::TryParse: espaços nas pontas e sinal; falha vira 0.
function inteiroOuZero(t) {
  if (t === null || !/^\s*[+-]?\d+\s*$/.test(t)) return { ok: false, n: 0 };
  const n = parseInt(t, 10);
  return n >= -2147483648 && n <= 2147483647 ? { ok: true, n } : { ok: false, n: 0 };
}

function periodo(q, base) {
  const meses = R.mesesDisponiveis(base);
  const max = meses.length ? meses[meses.length - 1] : 1;
  let de = q('de') ? inteiroOuZero(q('de')).n : 1;
  let ate = q('ate') ? inteiroOuZero(q('ate')).n : max;
  de = Math.max(1, Math.min(de, max));
  ate = Math.max(1, Math.min(ate, max));
  if (de > ate) [de, ate] = [ate, de];
  return [de, ate];
}

function diasParado(q) {
  // §5.0: vazio usa o padrão; inválido usa o padrão e avisa.
  const texto = q('dias');
  if (!texto) return [RE.DIAS_PARADO_PADRAO, ''];
  const r = inteiroOuZero(texto);
  if (r.ok && r.n >= 1 && r.n <= RE.DIAS_PARADO_MAXIMO) return [r.n, ''];
  return [RE.DIAS_PARADO_PADRAO, `Prazo inválido: use um número inteiro de 1 a ${RE.DIAS_PARADO_MAXIMO} dias. Mostrando o padrão de ${RE.DIAS_PARADO_PADRAO} dias.`];
}

function cookieSessao(req, valor, apagar) {
  const seguro = String(req.headers['x-forwarded-proto'] || '').split(',')[0].trim() === 'https' ? '; Secure' : '';
  return `sessao=${valor}; Path=/; HttpOnly; SameSite=Strict${seguro}${apagar ? '; Max-Age=0' : ''}`;
}

async function entrar(req, res) {
  const form = await lerFormulario(req);
  const login = String(form.login ?? '').trim().toLowerCase();
  const senha = String(form.senha ?? '');
  const generica = 'Usuário ou senha inválidos.';
  if (!login || !senha) return responder(res, 400, HTML, P.paginaLogin(generica, login));
  if (A.bloqueado(login)) return responder(res, 429, HTML, P.paginaLogin('Muitas tentativas. Aguarde alguns minutos e tente de novo.', login));
  const u = A.acharUsuario(login);
  if (!u || !A.testarSenha(u, senha)) {
    A.registrarFalha(login);
    console.log(`Login recusado: ${login}`);
    return responder(res, 401, HTML, P.paginaLogin(generica, login));
  }
  // §2.2: vendedor inativo não faz login. Confere no cadastro atual.
  if (u.idVendedor) {
    const base = atualizarBase();
    const vend = !base.Erros.length ? base.VendedorPorId.get(u.idVendedor) : null;
    if (!vend || !vend.Ativo) {
      console.log(`Login recusado (vendedor inativo ou não encontrado): ${login}`);
      return responder(res, 403, HTML, P.paginaLogin('Acesso desativado. Procure a diretoria comercial.', login));
    }
  }
  console.log(`Login: ${login}`);
  redirecionar(res, '/', cookieSessao(req, A.novaSessao(u)));
}

function importacao(res, usuario) {
  // §9.1: só a diretoria importa. Na Vercel a gravação não é possível (sem disco permanente).
  if (!N.igual(usuario.perfil, 'diretoria')) {
    return responder(res, 403, HTML, P.layout('Sem acesso', '<section class="aviso-erro"><h1>Sem acesso</h1><p>Só a diretoria importa arquivos (REGRAS_NEGOCIO.md §9.1).</p></section>', usuario, 'importar'));
  }
  responder(res, 501, HTML, P.layout('Importar', '<section class="aviso-erro"><h1>Importação indisponível nesta versão</h1><p>A versão publicada na internet só lê as planilhas do repositório: ela não tem onde gravar arquivos. Importe pela versão local do painel (app/servidor.ps1), como descrito em docs/PAINEL.md.</p></section>', usuario, 'importar'));
}

async function tratar(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const caminho = url.pathname;
  const metodo = req.method;
  const q = query(url);

  if (caminho === '/estilo.css' && metodo === 'GET') return responder(res, 200, 'text/css; charset=utf-8', P.css());
  if (caminho === '/robots.txt' && metodo === 'GET') return responder(res, 200, TEXTO, 'User-agent: *\nDisallow: /\n');

  const token = cookie(req, 'sessao');
  const sessao = A.sessao(token);

  if (caminho === '/entrar') {
    if (metodo === 'POST') return entrar(req, res);
    if (sessao) return redirecionar(res, '/');
    return responder(res, 200, HTML, P.paginaLogin('', ''));
  }
  if (caminho === '/sair' && metodo === 'POST') return redirecionar(res, '/entrar', cookieSessao(req, '', true));

  if (!sessao) return redirecionar(res, '/entrar');
  const usuario = sessao.Usuario;
  if (caminho === '/importar' || caminho.startsWith('/importar/')) return importacao(res, usuario);
  if (metodo !== 'GET') return responder(res, 405, TEXTO, 'Método não permitido');

  const base = atualizarBase();
  if (base.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(base.Erros, usuario));

  const [de, ate] = periodo(q, base);
  const incluirDesligados = q('desligados') === '1';
  const desempenho = R.desempenho(base, de, ate);

  if (caminho === '/estoque' || caminho === '/estoque.csv') {
    const estoque = atualizarEstoque();
    if (estoque.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(estoque.Erros, usuario, 'estoque'));
    const pipeline = atualizarPipeline();   // só para contar oportunidades abertas por produto (§5.4)
    const [dias, avisoPrazo] = diasParado(q);
    const visao = RE.visaoEstoque(estoque, dias, RE.abertasPorProduto(pipeline));
    if (caminho === '/estoque') {
      let filial = q('filial') || '';
      if (filial && !visao.Filiais.some((f) => N.igual(f.Id, filial))) filial = '';
      return responder(res, 200, HTML, PE.paginaEstoque(visao, usuario, filial, avisoPrazo));
    }
    return responder(res, 200, 'text/csv; charset=utf-8', PE.csvEstoque(visao), { bom: true, disposicao: `attachment; filename="estoque_${N.fmtDataIso(visao.DataBase)}_parado${dias}d.csv"` });
  }

  if (caminho === '/funil') {
    const pipeline = atualizarPipeline();
    if (pipeline.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(pipeline.Erros, usuario, 'funil'));
    // §8.4: só valores conhecidos; qualquer outro vira "sem filtro".
    let fVend = q('vendedor') || '';
    if (fVend && ![...base.VendedorPorId.keys()].some((k) => N.igual(k, fVend))) fVend = '';
    let fFilial = q('filial') || '';
    if (fFilial && !pipeline.Filiais.some((f) => N.igual(f.Id, fFilial))) fFilial = '';
    let fCat = q('categoria') || '';
    if (fCat && !N.contem([...pipeline.ProdutoPorId.values()].map((p) => p.Categoria), fCat)) fCat = '';
    const medida = N.igual(q('medida'), 'valor') ? 'valor' : 'qtd';
    const visao = RP.visaoFunil(base, pipeline, fVend, fFilial, fCat);
    return responder(res, 200, HTML, PP.paginaFunil(visao, usuario, medida));
  }

  if (caminho === '/pipeline' || caminho === '/pipeline.csv') {
    const pipeline = atualizarPipeline();
    if (pipeline.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(pipeline.Erros, usuario, 'pipeline'));
    const visao = RP.visaoPipeline(base, desempenho, pipeline);
    if (caminho === '/pipeline') {
      let filtro = q('vendedor') || '';
      if (filtro && ![...base.VendedorPorId.keys()].some((k) => N.igual(k, filtro))) filtro = '';
      return responder(res, 200, HTML, PP.paginaPipeline(base, visao, usuario, filtro));
    }
    return responder(res, 200, 'text/csv; charset=utf-8', PP.csvPipeline(visao), { bom: true, disposicao: `attachment; filename="pipeline_crm_${N.fmtDataIso(visao.DataBaseCrm)}.csv"` });
  }

  if (caminho === '/origem-vendas') {
    const pipeline = atualizarPipeline();
    if (pipeline.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(pipeline.Erros, usuario, 'origem'));
    const visao = RO.visaoOrigem(base, pipeline, de, ate);
    if (visao.Erros.length) return responder(res, 500, HTML, P.paginaErroImportacao(visao.Erros, usuario, 'origem'));
    return responder(res, 200, HTML, PE.paginaOrigem(base, visao, usuario));
  }

  if (caminho === '/') return responder(res, 200, HTML, P.paginaPainel(base, desempenho, usuario, incluirDesligados, FUSO));
  if (caminho === '/vendedor') {
    const id = q('id') || '';
    const linha = desempenho.Linhas.find((l) => N.igual(l.Vendedor.Id, id));
    if (!linha) return responder(res, 404, TEXTO, 'Vendedor não encontrado');
    const detalhe = R.detalheVendedor(base, id, de, ate);
    return responder(res, 200, HTML, P.paginaVendedor(base, desempenho, detalhe, linha, usuario, incluirDesligados));
  }
  if (caminho === '/conferencia.csv') {
    const d2 = (x) => String(x).padStart(2, '0');
    return responder(res, 200, 'text/csv; charset=utf-8', P.csvConferencia(desempenho), { bom: true, disposicao: `attachment; filename="conferencia_2026_${d2(de)}-${d2(ate)}.csv"` });
  }
  responder(res, 404, TEXTO, 'Página não encontrada');
}

async function handler(req, res) {
  try {
    await tratar(req, res);
  } catch (e) {
    console.error(`Erro em ${req.method} ${req.url}:`, e);
    try { if (!res.headersSent) responder(res, 500, TEXTO, 'Erro interno. Veja o console do servidor.'); } catch { /* resposta já iniciada */ }
  }
}

module.exports = handler;
