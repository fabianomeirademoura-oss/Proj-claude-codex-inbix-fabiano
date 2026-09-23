# AGENTS.md — Instruções de entrada para o Codex

Projeto: **painel comercial da Horizonte Máquinas** (vendas e metas, pipeline,
funil e estoque). Esta pasta (`painel-horizonte/`) é a **raiz oficial**. O
Claude Code trabalha no mesmo projeto, com as instruções em
[CLAUDE.md](CLAUDE.md).

## Leia antes de qualquer tarefa

1. **[COORDENACAO.md](COORDENACAO.md):** o acordo de trabalho com o Claude Code
   e o registro de tarefas. Consulte as tarefas em andamento e **registre a sua
   antes de editar**.
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
pwsh -NoProfile -File testes/teste_e2e.ps1
```

Depois, registre em `COORDENACAO.md` os arquivos alterados, as verificações
executadas, as limitações e os impactos para o Claude Code.
