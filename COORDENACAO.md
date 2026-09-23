# COORDENACAO.md — Acordo de trabalho paralelo (Claude Code + Codex)

Dois agentes trabalham neste projeto: o **Claude Code** (instruções em
`CLAUDE.md`) e o **Codex** (instruções em `AGENTS.md`), sob a coordenação do
responsável pelo projeto. Este arquivo tem o acordo e o registro de tarefas.

## Regras

1. **O coordenador** define as prioridades e atribui as tarefas a cada agente.
   Nenhum agente atribui tarefa ao outro.
2. **Antes de editar**, cada agente registra aqui sua tarefa, seu escopo e os
   arquivos ou módulos que pretende alterar.
3. **Antes de começar**, cada agente consulta as tarefas em andamento. O
   registro é um aviso de coordenação, não um bloqueio técnico.
4. **Não editar arquivos reservados pelo outro agente** sem combinar com o
   coordenador a transferência ou a divisão do trabalho.
5. **Se a tarefa exigir mudanças fora do escopo registrado**, atualizar o
   registro e verificar possíveis conflitos antes de prosseguir.
6. **Não desfazer, sobrescrever nem limpar alterações do outro agente.**
   Alterações inesperadas devem ser preservadas e comunicadas.
7. **Combinar antes com o coordenador** as mudanças em regras de negócio
   (`REGRAS_NEGOCIO.md`), na estrutura de pastas, em dependências e em
   interfaces compartilhadas. A exceção é quando a tarefa já as autoriza
   explicitamente.
8. **Evitar** reformatações gerais e refatorações fora do escopo.
9. **Ao concluir**, registrar os arquivos alterados, as verificações executadas,
   as limitações e os impactos para o outro agente.
10. **Git:** se o projeto usar Git, preferir uma branch e um worktree por agente
    para tarefas simultâneas, com integração revisada. Se os dois trabalharem
    na mesma pasta (situação atual, sem Git), manter a separação explícita de
    arquivos pelo registro abaixo.

**Interfaces compartilhadas**, que exigem cuidado redobrado: `app/lib/Regras*.ps1`
(funções usadas pelas páginas e pelos testes), `app/lib/Paginas.ps1` (layout,
menu e CSS de todas as páginas), `app/servidor.ps1` (rotas) e os documentos de
entrada (`CLAUDE.md`, `AGENTS.md`, `REGRAS_NEGOCIO.md`, este arquivo).

## Como registrar

Acrescente uma linha na tabela ao começar e atualize o status ao terminar.
Status possíveis: `planejada`, `em andamento`, `concluída`, `pausada`,
`cancelada`. Detalhes da entrega (verificações, limitações, impactos) vão na
seção "Entregas", abaixo da tabela.

## Registro de tarefas

| # | Agente | Tarefa | Escopo | Arquivos previstos | Status |
|---|---|---|---|---|---|
| 1 | Claude Code | Organizar a estrutura do projeto e centralizar as regras | Pastas do projeto, do material didático e do professor; documentos de entrada; caminhos do app e dos testes. Sem mudar critérios de cálculo e sem implementar a importação de setembro. | Toda a raiz `painel-horizonte/` (estrutura, docs, caminhos em `app/` e `testes/`), `../curso-claude-code/README.md`, `../material-professor/`, `../.claude/launch.json`, `../LEIA-ME.md` | concluída |
| 2 | Codex | Origem do faturamento: vendas vinculadas ao CRM | Nova página com período, participação no faturamento e vínculo venda–oportunidade por ID; explicação das vendas sem oportunidade. Integração mínima de menu e rota; sem mudar critérios existentes. | app/lib/RegrasOrigem.ps1; app/lib/PaginasOrigem.ps1; app/lib/Paginas.ps1 (menu); app/servidor.ps1 (carga e rota); testes/conferir_origem.ps1; testes/teste_e2e.ps1; docs/PAINEL.md; COORDENACAO.md | concluída |
| 3 | Claude Code | Tela de importação de Excel (vendas incrementais, estoque substitutivo) | Tela `/importar` só para a diretoria, com resumo antes de gravar e confirmação; recusa arquivo sem coluna obrigatória. Arquivos aceitos ficam em `dados/importacoes/`, sem alterar as planilhas originais. Nova §9 no REGRAS_NEGOCIO.md, com as decisões aprovadas pelo coordenador em 23/09. Os carregadores só aplicam importações quando o servidor pede; as conferências continuam sobre a base. | Novos: app/lib/RegrasImportacao.ps1, app/lib/PaginasImportacao.ps1, testes/conferir_importacao.ps1 (+ .cmd), testes/XlsxTeste.ps1. Alterados: app/lib/Regras.ps1 e app/lib/RegrasEstoque.ps1 (parâmetro opcional de importações), app/servidor.ps1 (rota e upload), app/lib/Paginas.ps1 (menu e CSS), testes/teste_e2e.ps1, REGRAS_NEGOCIO.md (§0.3 nota, §9), README.md, docs/PAINEL.md, PERFIL.md (estado do caso 7), .gitignore, COORDENACAO.md | concluída |
| 4 | Claude Code | Versionar o projeto no GitHub | `git init` na raiz, primeiro commit e envio para `fabianomeirademoura-oss/Proj-claude-codex-inbix-fabiano` (público, decisão do coordenador em 23/09). `dados/` entra; `dados/importacoes/`, `.env` e `app/usuarios.json` ficam fora. Sem mudar código nem regras. | .git/ (novo), .gitignore, COORDENACAO.md | em andamento |

## Entregas

### #1 — Claude Code — Organização da estrutura (23/09/2026)

- **Estrutura:**
  - O projeto saiu de `curso-claude-code/Curso Claude com edicoes até aula 4/dados_horizonte_maquinas/` para `painel-horizonte/`.
  - Dados em `dados/`; testes em `testes/`; manual do painel em `docs/PAINEL.md` (antigo `app/LEIAME.md`).
- **Duplicatas:** as 10 planilhas da cópia interna eram idênticas às da base, conferidas byte a byte. Elas ficaram como `dados/` do projeto. O `gerar_dados.py` duplicado, também idêntico, foi removido do projeto e mantido uma única vez, em `material-professor/`.
- **Material didático:**
  - `curso-claude-code/` ficou só com o material do aluno.
  - Gabarito, script gerador e o README original completo foram para `material-professor/`, sem alteração de conteúdo.
- **Regras:**
  - O contrato saiu de `CLAUDE.md` para `REGRAS_NEGOCIO.md`, com as mesmas seções e a mesma numeração.
  - A única adição ao corpo do contrato é uma nota em §0.3 (data-base de vendas atual: 31/08).
  - As referências "CLAUDE.md §x" no código, nas telas e nos testes passaram a "REGRAS_NEGOCIO.md §x".
  - O `PERFIL.md` separa agora as decisões aprovadas das sugestões exploratórias.
- **Caminhos:** o app e `criar_usuarios` leem `dados/`; os testes carregam `../app/lib`. A variável `HORIZONTE_DADOS` continua valendo.
- **Verificações:** estão no README.md, seção "Testes".
  - Conferências: 212 (desempenho), 608 (pipeline/funil) e 205 (estoque), total de 1.025.
  - Teste de ponta a ponta: 49.
  - Inicialização conferida pelo HTTP na porta de teste.
- **Limitações:**
  - Os `.cmd` não foram executados, porque o ambiente é macOS; os `.ps1` equivalentes rodaram com `pwsh`.
  - Um servidor iniciado antes da reorganização aponta para a pasta antiga e precisa ser reiniciado.
- **Impactos para o Codex:**
  - Abrir `painel-horizonte/`.
  - Os caminhos antigos (`app/conferir*.ps1`, `app/LEIAME.md`, `dados_horizonte_maquinas/curso-claude-code/`) não existem mais.
  - Referenciar regras sempre como `REGRAS_NEGOCIO.md §x`.

### #2 — Codex — Origem do faturamento (23/09/2026)

- **Entrega:** nova página `/origem-vendas`, acessível pelo menu, com filtro de
  meses, valor e participação do CRM, composição por categoria das vendas sem
  oportunidade e tabela de todas as vendas do período com ID, oportunidade,
  etapa, origem no CRM e linhas das planilhas. Canceladas aparecem fora dos totais.
- **Critério:** REGRAS_NEGOCIO.md §0.1 e §1; vínculo exclusivamente por ID
  Oportunidade, valor efetivamente faturado. Sem inferência por cliente/produto.
  ID preenchido inexistente impede a exibição de totais incompletos.
- **Resultado jan–ago:** R$ 15.301.000,00 / R$ 45.295.119,16 = 33,8% do CRM;
  95 vendas vinculadas e 628 sem oportunidade. A explicação distingue a
  descrição didática de balcão da ausência de registro no CRM.
- **Arquivos alterados:** novos `app/lib/RegrasOrigem.ps1`,
  `app/lib/PaginasOrigem.ps1`, `testes/conferir_origem.ps1`; integração em
  `app/servidor.ps1` e somente menu em `app/lib/Paginas.ps1`; extensão de
  `testes/teste_e2e.ps1`; manual `docs/PAINEL.md`; este registro.
- **Verificações:** 212 desempenho + 608 pipeline/funil + 205 estoque +
  375 origem + 56 HTTP = 1.456 aprovadas. Casos adicionais: canceladas,
  período vazio, valor zero, vínculo inexistente/duplicado, HTML escapado,
  filtro jul–ago e reconciliação com faturamento do painel. Layout revisado
  no navegador. Arquivos PowerShell gravados em UTF-8 com BOM.
- **Limitações:** validado no macOS com PowerShell 7; Windows PowerShell 5.1
  não executado. Importações de setembro permanecem fora do escopo.
- **Execução:** versão atualizada iniciada em `http://localhost:8081/` com
  autenticação existente. Processo da porta 8080 preservado; ele só carregará
  os novos módulos após um reinício coordenado.
- **Impacto para Claude Code:** reservar a nova rota e nomes das funções;
  nenhum critério de negócio alterado. O e2e passa a ter 56 verificações
  (as 49 originais mais 7). Executar também `testes/conferir_origem.ps1`.
  Arquivos desta tarefa liberados.

### #3 — Claude Code — Tela de importação de Excel (23/09/2026)

- **Entrega:** página `/importar`, só para o perfil diretoria. Fluxo em três passos: escolher o tipo e o arquivo, ver o resumo sem gravar nada, e confirmar ou cancelar.
  - Vendas: novas, atualizadas (com o que muda, coluna a coluna) e ignoradas, pelo `ID Venda`.
  - Estoque: foto em vigor × nova, linhas substituídas, quantidades que mudam e impacto nos indicadores.
  - Recusa o arquivo inteiro, dizendo exatamente o motivo: colunas que faltam, aba errada, arquivo que não é xlsx, ID repetido, erro de validação da carga, foto mais antiga ou nome de estoque fora do padrão.
- **Regra:** nova REGRAS_NEGOCIO.md §9, com as decisões do coordenador de 23/09:
  - só a diretoria importa;
  - todas as colunas da planilha atual são obrigatórias;
  - os arquivos aceitos são guardados em `dados/importacoes/`, sem alterar as planilhas originais;
  - a foto de estoque da mesma data substitui a daquela data.

  A nota da §0.3 foi ajustada. Nenhum critério de cálculo existente mudou.
- **Interfaces compartilhadas alteradas (compatíveis):**
  - `Import-BaseComercial` e `Import-BaseEstoque` ganharam o parâmetro opcional `-DirImportacoes` (e `-VendasAdicionais`, `-ArquivoFoto` e `-DataFoto`, usados só no resumo).
  - `Get-AssinaturaBase` e `Get-AssinaturaEstoque` ganharam o 2º argumento opcional.
  - Sem esses parâmetros, a carga é idêntica à anterior.
  - Cada venda carregada ganhou a propriedade `Arquivo` (de onde veio a versão em vigor). `Linha` continua sendo a linha dentro desse arquivo.
  - Em `servidor.ps1`: rota `/importar`, leitor multipart (limite de 10 MB) e passagem de `-DirImportacoes` às cargas.
  - Em `Paginas.ps1`: link "Importar" no menu, só para a diretoria, e o CSS da tela.
- **Verificações:**
  - Conferências: 212 + 608 + 205 + 375 (origem) + 70 (importação) = 1.470 aprovadas.
  - E2e: 73 verificações (as 56 anteriores mais 17). Ele agora roda sobre uma cópia temporária de `dados/`.
  - Os números batem com a documentação: setembro tem 82 vendas novas e R$ 3.850.268,28 faturados; o estoque de 19/09 muda 61 quantidades; P003 sai de "sem estoque" e P030/P034/P042 entram.
  - Telas conferidas no navegador com a CSP do servidor.
  - `.ps1` em UTF-8 com BOM.
- **Limitações:**
  - Validado no macOS com PowerShell 7; o Windows PowerShell 5.1 e os `.cmd` não foram executados.
  - Desfazer uma importação é manual: retirar o arquivo de `dados/importacoes/`.
  - O CRM não é importável.
  - O servidor em execução precisa ser reiniciado para ter a rota `/importar`.
- **Impactos para o Codex:**
  - Depois de uma importação real, `/origem-vendas` passa a incluir setembro e a data-base de vendas passa a 19/09.
  - Para as vendas de setembro, a coluna de linha da página de origem se refere ao arquivo importado (propriedade `Arquivo`), e não a `vendas_2026_jan-ago.xlsx`.
  - Os testes do Codex continuam sobre a base e não são afetados.
  - `CLAUDE.md` e `AGENTS.md` ainda listam só 4 verificações. A lista completa está no README; a atualização dos documentos de entrada fica para o coordenador decidir.
