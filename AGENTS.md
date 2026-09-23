# AGENTS.md — Instruções de entrada para o Codex

Projeto: **painel comercial da Horizonte Máquinas** (vendas e metas, pipeline,
funil e estoque). Este repositório é a **raiz oficial**, e você trabalha na sua própria pasta de
trabalho, `../painel-codex`. O
Claude Code trabalha no mesmo projeto, com as instruções em
[CLAUDE.md](CLAUDE.md).

## Leia antes de qualquer tarefa

1. **[COORDENACAO.md](COORDENACAO.md):** o acordo de trabalho entre os dois
   agentes e as regras de Git. Consulte as tarefas em andamento (`gh pr list`)
   e **registre a sua num pull request em rascunho antes de editar**.
2. **[REGRAS_NEGOCIO.md](REGRAS_NEGOCIO.md):** a fonte única das regras de
   negócio. Se o código e esse arquivo divergirem, o arquivo vence. Não copie
   regras para cá nem para outros documentos: referencie a seção (ex.: §5).
3. **[PERFIL.md](PERFIL.md):** o contexto dos dados e os casos de borda. Ele
   descreve os dados e **não cria regras**. Sugestões marcadas como
   exploratórias não estão aprovadas.
4. **[README.md](README.md):** a estrutura, como executar, os testes e o que já
   está implementado versus o que está previsto.

## Regras de trabalho

- Quando uma pergunta de negócio não estiver coberta em `REGRAS_NEGOCIO.md`,
  **pergunte antes de inventar um critério**.
- Mudanças em regras de negócio, na estrutura de pastas, em dependências e em
  interfaces compartilhadas são combinadas antes com o coordenador, salvo
  quando a tarefa já as autorizar explicitamente. Uma mudança de regra edita
  primeiro `REGRAS_NEGOCIO.md` e só depois o código.
- Trabalhe só dentro desta pasta. `../curso-claude-code/` (material do aluno) e
  `../material-professor/` (gabarito e respostas) não fazem parte do projeto em
  desenvolvimento e não são editados sem pedido explícito.
- Nunca mostre, copie ou versione o conteúdo de `.env` ou de
  `app/usuarios.json`.

## Git: sua pasta, sua branch, sua assinatura

- **Trabalhe em `../painel-codex`**, a sua pasta de trabalho. Ela já assina os commits
  como `Codex (agente)`. A pasta `painel-horizonte/` é a `main` do coordenador, e
  `../painel-claude` é a do outro agente: não edite nenhuma das duas.
- Cada tarefa começa numa branch nova:
  `git fetch origin && git switch -c codex/<tarefa> origin/main`.
- **Todo commit termina com a linha `Agente: Codex`.** Sem ela, o commit é
  recusado.
- Registre a tarefa logo no início, abrindo um pull request em rascunho com o
  modelo preenchido: `git push -u origin HEAD && gh pr create --draft --base main`.
- Nunca faça commit nem push na `main`, e nunca use `push --force`,
  `reset --hard`, `clean` ou `stash` fora da sua branch. Para atualizar a sua
  branch com a `main`, use `git rebase origin/main`.
- Detalhes: COORDENACAO.md, regras 10 e 11.

## Convenções técnicas

- PowerShell sem dependências externas: Windows PowerShell 5.1 (`.cmd`) e
  PowerShell 7 (`pwsh`) no macOS.
- Arquivos `.ps1` são gravados em **UTF-8 com BOM**. Sem o BOM, o Windows
  PowerShell 5.1 estraga os acentos. Preserve o BOM ao editar.
- Regras ficam em `app/lib/Regras*.ps1`. As páginas (`app/lib/Paginas*.ps1`)
  só formatam.
- A CSP do servidor proíbe script e estilo inline: estilos vão em `Get-Css`
  (`app/lib/Paginas.ps1`).
- Dinheiro é `decimal`, nunca `double` (REGRAS_NEGOCIO.md §0.4).

## Antes de concluir

Rode as verificações, a partir desta pasta:

```
pwsh -NoProfile -File testes/conferir.ps1
pwsh -NoProfile -File testes/conferir_pipeline.ps1
pwsh -NoProfile -File testes/conferir_estoque.ps1
pwsh -NoProfile -File testes/conferir_origem.ps1
pwsh -NoProfile -File testes/conferir_importacao.ps1
pwsh -NoProfile -File testes/teste_e2e.ps1
```

Os números esperados de cada verificação estão no README.md, seção \"Testes\".
Depois, preencha a seção \"Entrega\" do seu pull request (os arquivos alterados,
as verificações executadas, as limitações e os impactos para o Claude Code) e marque-o
como pronto para revisão.
