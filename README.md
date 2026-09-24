# Painel comercial — Horizonte Máquinas

Aplicação web com login sobre as planilhas da Horizonte Máquinas Agrícolas
(empresa fictícia do curso). Telas: vendas e metas, origem do faturamento,
pipeline do CRM, funil de vendas, estoque e importação de Excel (só diretoria).
Ela é desenvolvida em paralelo pelo Claude Code e
pelo Codex, sob coordenação humana.

## Qual pasta abrir

**Abra a raiz deste repositório tanto no Claude Code quanto no Codex.** No
computador do coordenador, ela fica em
`~/Downloads/dados_horizonte_maquinas/painel-horizonte`. Quem clonar do GitHub
abre a pasta criada pelo `git clone`.

É a raiz oficial do projeto. O Claude Code lê [CLAUDE.md](CLAUDE.md), o Codex
lê [AGENTS.md](AGENTS.md), e os dois seguem os mesmos documentos
compartilhados.

No computador do coordenador, a pasta de cima (`dados_horizonte_maquinas/`) reúne três coisas separadas; só a primeira está neste repositório:

```
dados_horizonte_maquinas/
├── painel-horizonte/     ← projeto em desenvolvimento (RAIZ OFICIAL, branch main do coordenador)
├── painel-claude/        ← pasta de trabalho do Claude Code (branches claude/*)
├── painel-codex/         ← pasta de trabalho do Codex (branches codex/*)
├── curso-claude-code/    ← base inicial: material do aluno (planilhas + README sem respostas)
└── material-professor/   ← gabarito, README original com respostas, gerador (não distribuir)
```

O painel **não lê** `curso-claude-code/` nem `material-professor/`. As planilhas
do projeto ficam em `dados/`.

## Estrutura da raiz

```
painel-horizonte/
├── CLAUDE.md            instruções de entrada do Claude Code
├── AGENTS.md            instruções de entrada do Codex
├── COORDENACAO.md       acordo de trabalho paralelo e registro de tarefas
├── REGRAS_NEGOCIO.md    fonte única das regras de negócio (vence o código)
├── PERFIL.md            contexto dos dados e casos de borda (não cria regras)
├── README.md            este arquivo
├── .gitignore           credenciais e arquivos locais fora do versionamento
├── .env                 LOCAL: logins e senhas em texto (nunca versionar nem enviar)
├── .claude/launch.json  inicialização do servidor para o preview do Claude Code
├── app/                 aplicação
│   ├── servidor.ps1         rotas HTTP (só localhost)
│   ├── iniciar.cmd          atalho do Windows para subir o servidor
│   ├── criar_usuarios.ps1   cria as contas (e o atalho .cmd)
│   ├── usuarios.json        LOCAL: hashes das senhas (nunca versionar)
│   └── lib/                 Regras*.ps1 (regras), Paginas*.ps1 (HTML), Auth, Xlsx
├── dados/               planilhas base (jan–ago/2026, foto de estoque 31/08) — nunca alteradas
│   ├── atualizacoes/        planilhas de 19/09 para importar pela tela (NÃO lidas diretamente)
│   └── importacoes/         criada na 1ª importação: arquivos confirmados, registro.csv e pendentes/
├── docs/
│   ├── MCP.md               servidor MCP (só leitura) e como conectar no Claude Desktop
│   ├── PAINEL.md            manual do painel: contas, telas e conferência no Excel
│   └── RELATORIO_DIRETORIA.md  relatório semanal da diretoria em PDF
├── relatorios/          LOCAL: relatórios semanais gerados (fora do Git)
├── testes/              conferências independentes, teste de ponta a ponta e paridade Node × PowerShell
├── mcp/                 servidor MCP para o Claude Desktop: leitura e escrita com confirmação (usa as regras de web/lib/)
├── web/                 versão Node do painel, para a Vercel (espelho de app/; contas públicas em contas_publicas.json)
├── api/index.js         entrada da Vercel (chama web/app.js)
├── public/              arquivos estáticos da Vercel (só robots.txt: nada de dados/ fica público sem login)
├── vercel.json          configuração da Vercel
└── package.json         metadados da versão Node (sem dependências)
```

## Como executar

**No Windows:**

1. Na primeira vez, rode `app\criar_usuarios.cmd`.
2. Depois, `app\iniciar.cmd`.
3. Abra http://localhost:8080.

**No macOS**, com PowerShell 7 (`brew install powershell`), a partir desta pasta:

```
cd <pasta do projeto>
pwsh -NoProfile -File app/servidor.ps1 -Porta 8080
```

- **Primeira vez:** antes do servidor, rode `pwsh -NoProfile -File app/criar_usuarios.ps1`.
- **Login:** `diretoria`, ou o e-mail de um vendedor ativo. As senhas estão no `.env`.
- **Detalhes:** contas, telas e como conferir cada número no Excel estão em [docs/PAINEL.md](docs/PAINEL.md).

### Versão web (Vercel)

A Vercel não roda PowerShell. Para publicar o painel na internet, ele tem uma
segunda versão em Node, em `web/`, sem dependências. Ela lê as mesmas planilhas
de `dados/` e gera as mesmas telas, byte a byte (`testes/paridade_node.ps1`).
A versão PowerShell continua sendo a de referência. Regra nova ou tela nova
entra primeiro nela e depois é espelhada em `web/`, arquivo por arquivo
(`Regras.ps1` → `web/lib/regras.js`, `Paginas.ps1` → `web/lib/paginas.js` etc.).

- **Publicação:** cada merge na `main` publica sozinho. A Vercel está ligada ao
  repositório (`vercel.json` manda todas as rotas para `api/index.js`).
- **Conta dos alunos:** `teste` / `teste`, perfil `leitura`. Ela vê todas as
  telas e não importa. O hash dessa conta está em `web/contas_publicas.json`.
- **Sem importação na internet:** a Vercel não tem disco permanente. A versão
  publicada mostra a base de `dados/` (janeiro a agosto) e importa só na versão
  local (§9).
- **Outras contas na Vercel (opcional):** configure `HORIZONTE_SEGREDO` (texto
  aleatório longo) e `HORIZONTE_USUARIOS` (o conteúdo de um `usuarios.json`)
  nas variáveis de ambiente do projeto. Sem o segredo, só as contas públicas
  entram.
- **Rodar a versão Node na máquina:** `node web/servidor.js 8090`.

## Testes

Rode a partir desta pasta. No Windows, cada `conferir*.ps1` tem um `.cmd` com o
mesmo nome em `testes\`.

| Comando | O que confere | Resultado esperado |
|---|---|---|
| `pwsh -NoProfile -File testes/conferir.ps1` | Vendas, metas e atingimento em três períodos, contra somas diretas nas planilhas e REGRAS_NEGOCIO.md §6 | 212 conferências |
| `pwsh -NoProfile -File testes/conferir_pipeline.ps1` | Pipeline por vendedor e funil em 26 recortes | 608 conferências |
| `pwsh -NoProfile -File testes/conferir_estoque.ps1` | Estoque em quatro prazos de "parado" | 205 conferências |
| `pwsh -NoProfile -File testes/conferir_origem.ps1` | Origem do faturamento: vínculo venda–oportunidade por ID (tarefa #2, Codex) | 375 conferências |
| `pwsh -NoProfile -File testes/conferir_importacao.ps1` | Importação de vendas e de estoque numa cópia temporária de `dados/`: resumo, gravação, reimportação, recusas | 70 conferências |
| `pwsh -NoProfile -File testes/conferir_relatorio_diretoria.ps1` | Relatório semanal da diretoria: itens 1 a 5 contra as planilhas brutas e a §6, PDF e `calculo.json` idênticos em duas execuções, recusa de riscos inválidos, setembro importado | 237 conferências |
| `pwsh -NoProfile -File testes/teste_e2e.ps1` | Sobe o servidor com contas temporárias, numa cópia temporária de `dados/`, e testa login, permissões, telas, CSV e importação por HTTP | 73 verificações |
| `pwsh -NoProfile -File testes/conferir_mcp.ps1` | Servidor MCP pelo protocolo: vendedores, metas, pipeline, estoque e produtos contra a versão PowerShell e a §6, filtros e erros de entrada, numa cópia temporária de `dados/` (precisa do `node`) | 1.503 conferências |
| `pwsh -NoProfile -File testes/conferir_escrita_mcp.ps1` | Ferramentas de escrita do MCP (§11): permissões, recusas, nada gravado antes da confirmação, código de uso único, confirmação que refaz a validação, registro linha a linha e aplicação pela versão PowerShell, numa cópia temporária (precisa do `node`) | 226 conferências |
| `pwsh -NoProfile -File testes/paridade_node.ps1` | Sobe as versões PowerShell e Node lado a lado e compara status, tipo e corpo de 139 telas e CSV nas contas `teste`, vendedor, gerente e diretoria, em dois cenários: a base e a base com um registro de alterações (precisa do `node`) | 4.547 verificações |

As seis conferências somam **1.707** comparações. Todas devem terminar com
"Todas as N … bateram". Nenhum teste grava nos dados reais: as conferências
leem só a base (§9.9) e o teste de importação e o e2e trabalham em cópias.

## Como o código muda (dois agentes)

Cada agente trabalha na sua pasta, na sua branch e com a sua assinatura. Nada
entra na `main` sem pull request revisado pelo coordenador. Em cada commit dá
para saber quem o fez: o autor é `Claude Code (agente)` ou `Codex (agente)`, e
a mensagem termina com `Agente: Claude Code` ou `Agente: Codex`. Um gancho
local (`.githooks/commit-msg`) e a checagem "autoria" do GitHub recusam commits
fora da regra. As regras completas estão em [COORDENACAO.md](COORDENACAO.md),
regras 10 e 11.

Para ver quem fez o quê: `git log --format='%h %an  %s'`, ou a aba Commits do
GitHub.

## Estágio atual

### Implementado

| Funcionalidade | Onde | Regras |
|---|---|---|
| Login, contas por vendedor ativo, bloqueio após 5 erros, sessão de 8 h | `app/lib/Auth.ps1`, `servidor.ps1` | §2.2 |
| Vendas e metas: realizado, atingimento, rankings, quadro de conferência, detalhe por vendedor, CSV | `/`, `/vendedor`, `/conferencia.csv` | §1–§4 |
| Pipeline por vendedor (meta, realizado, pipeline, ponderado, atrasadas, órfãs), CSV | `/pipeline`, `/pipeline.csv` | §7 |
| Funil de vendas com filtros por vendedor, filial e categoria | `/funil` | §8 |
| Estoque: dinheiro no pátio por filial × categoria, parados (prazo configurável), sem estoque, abaixo do mínimo, CSV | `/estoque`, `/estoque.csv` | §5 |
| Origem do faturamento: vendas vinculadas ao CRM por `ID Oportunidade` | `/origem-vendas` | §0.1, §1 |
| Importação de Excel só pela diretoria: vendas incrementais (upsert por `ID Venda`) e estoque substitutivo, com resumo e confirmação antes de gravar; recusa arquivo sem coluna obrigatória, dizendo qual | `/importar` | §9 |
| Validação das planilhas, com lista de erros na tela no lugar dos números | todas as telas | §0.2, §2.6, §5, §7.5 |
| Servidor MCP para o Claude Desktop, com respostas em JSON ([docs/MCP.md](docs/MCP.md)). Leitura: `consultar_vendedor`, `buscar_produto`, `ver_estoque`, `ver_meta`, `listar_oportunidades` e `ver_historico`. Escrita: `inativar_vendedor`, `reativar_vendedor`, `alterar_meta`, `cadastrar_ou_editar_produto` e `transferir_oportunidades`, que só preparam, e `confirmar_alteracao`, que grava depois do "sim". Escrita só para gerente (própria filial) e diretoria | `mcp/` | §1–§5, §7, §11 |
| Histórico das alterações feitas pelo MCP (quem, quando, registro, valor anterior e novo), com filtro e CSV, para diretoria e gerentes; o painel aplica as alterações por cima das planilhas | `/historico`, `/historico.csv` | §11 |
| Contas de gerente por filial e criação só das contas que faltam | `app/criar_usuarios.ps1 -Completar` | §11.2 |
| Relatório semanal da diretoria em PDF: ranking, abaixo de 80%, pipeline por etapa, previsão vencida, maior saída e parados, e três riscos analisados pelo Claude Code ([docs/RELATORIO_DIRETORIA.md](docs/RELATORIO_DIRETORIA.md)) | `/relatorio-diretoria` no Claude Code, ou `app/relatorio_diretoria.ps1` | §1–§5, §7 |

**O que o painel carrega:**

- **Vendas:** `dados/vendas_2026_jan-ago.xlsx` e, por cima, as importações confirmadas em `dados/importacoes/vendas/`, em ordem (§9.8). Sem importação, o período vai de janeiro a agosto e a data-base de vendas é 31/08/2026. Depois de importar setembro, vai até setembro e a data-base passa a 19/09/2026, com o aviso "parcial até 19/09" (§4.5).
- **Estoque:** a foto de data mais recente entre `dados/estoque_*.xlsx` e as importadas em `dados/importacoes/estoque/`. Hoje é a de 31/08/2026.
- **CRM:** a foto de 31/08/2026. O CRM não é importado pela tela.
- **Alterações de cadastro** (vendedores, metas, produtos e donos de oportunidades) confirmadas pelo Claude Desktop: `dados/importacoes/alteracoes.csv`, aplicado por cima de tudo, na ordem em que foi gravado (§11.4). As planilhas originais não mudam.

### Previsto (ainda não implementado)

| Requisito | Regra | Observação |
|---|---|---|
| Reatribuição de oportunidades órfãs pela tela do painel | §2.4 | Já é possível pelo Claude Desktop (`transferir_oportunidades`, §11.9); falta a tela. |
| Visão do gerente restrita à própria filial | — | Sem regra aprovada. Hoje todos veem tudo. |

**Sobre `dados/atualizacoes/`:** o painel não lê essa pasta diretamente. Os dois
arquivos (vendas de setembro e estoque de 19/09) são a matéria-prima da aula 6.
Eles entram no painel **pela tela Importar**, que valida, mostra o resumo e pede
confirmação. Não copie esses arquivos para `dados/` como atalho: o estoque de
19/09 passaria a valer sem as validações da importação (§9).

**Para desfazer uma importação**, retire o arquivo de `dados/importacoes/vendas/`
ou de `dados/importacoes/estoque/`. O painel recarrega sozinho.

**Fora do escopo aprovado:**

- pró-rata de metas por padrão (§4.5);
- deduplicação, baixa ou fechamento automático de oportunidades (§7.6);
- pesos do ponderado calculados pelo histórico de conversão (o CRM não guarda o histórico de etapas, §8.1).

Qualquer uma dessas mudanças exige uma regra nova em `REGRAS_NEGOCIO.md`.

## Credenciais e distribuição

- `.env` (senhas em texto) e `app/usuarios.json` (hashes) são **locais**. Os dois estão no `.gitignore` e não devem ser enviados por chat, por e-mail ou para o material do curso.
- **Senhas novas para todos:** rode `app/criar_usuarios.ps1 -Recriar`.
- **Material do curso:** o dos alunos é `../curso-claude-code/`, sem gabarito e sem respostas. O do professor fica em `../material-professor/`.
- **Respostas dentro do projeto:** esta pasta também tem números de conferência (REGRAS_NEGOCIO.md §6) e o perfil dos dados. Não é material do aluno sem uma revisão.

## Histórico da organização

Em 23/09/2026:

- O projeto saiu de `curso-claude-code/Curso Claude com edicoes até aula 4/dados_horizonte_maquinas/` para esta pasta.
- As regras saíram do `CLAUDE.md` para o `REGRAS_NEGOCIO.md`.

Os detalhes estão em [COORDENACAO.md](COORDENACAO.md), tarefa #1.
