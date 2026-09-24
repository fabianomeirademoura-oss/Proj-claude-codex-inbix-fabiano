# Servidor MCP: o painel dentro do Claude Desktop

O servidor MCP deixa o Claude Desktop consultar e alterar os dados da
Horizonte Máquinas direto na conversa. Exemplos: "como está a Mariana no ano?",
"o que está parado em Chapecó?", "passa as oportunidades do Ricardo para o
Carlos".

- **Leitura:** livre, com cinco ferramentas.
- **Escrita:** cinco ferramentas que só **preparam** a alteração, e uma que
  grava depois do "sim". Exige conta de gerente (na própria filial) ou de
  diretoria. Toda gravação fica no **histórico**. As regras estão no
  [REGRAS_NEGOCIO.md §11](../REGRAS_NEGOCIO.md).

Como o servidor está organizado:

- **Onde está:** `mcp/servidor.js` (protocolo), `mcp/ferramentas.js` (leitura) e `mcp/escrita.js` (escrita).
- **Do que precisa:** Node 20 ou mais novo. Não tem dependências nem `npm install`.
- **Regras:** as mesmas do painel. O servidor usa `web/lib/`, o espelho das
  regras da versão PowerShell. Cada resposta diz a data-base que usou e cita a
  seção do REGRAS_NEGOCIO.md.
- **Dados:** lê `dados/` e, por cima, as importações confirmadas e as
  alterações gravadas, que ficam em `dados/importacoes/` (§9.8, §11.4), como o
  painel local.
  - Outra pasta: variável `HORIZONTE_DADOS`.
  - Se uma planilha mudar, a próxima pergunta já vê a mudança.
  - Se uma planilha tiver erro de validação, a ferramenta devolve a lista de erros no lugar dos números.

## Ferramentas de leitura

Todas recebem parâmetros simples e devolvem JSON. Vendedor, produto e filial
podem ser informados pelo ID ou pelo nome, inteiro ou em parte, sem acento. O
nome vira ID antes de qualquer conta (§0.1). Nome ambíguo ou inexistente volta
como erro, com a lista de opções, e nunca com um palpite.

| Ferramenta | Parâmetros | O que devolve |
|---|---|---|
| `consultar_vendedor` | `vendedor` (obrigatório), `mes_inicial`, `mes_final` | Cadastro (filial, status, admissão, desligamento), meta, realizado, atingimento, falta para a meta, faturamento, ticket médio, canceladas e a abertura mês a mês. Sem período, usa o acumulado do ano até a data-base de vendas (§4). |
| `buscar_produto` | `busca` (código ou palavras do nome, marca ou modelo), `categoria`, `limite` | Cadastro do produto, com preço de tabela, status (descontinuado aparece marcado, §5.5) e unidades em estoque na foto em vigor. |
| `ver_estoque` | `produto`, `filial`, `dias_parado` (padrão 180), `somente_parados` | Linhas (produto, filial) da foto em vigor: quantidade, mínimo, valor imobilizado, últimas entrada e saída, dias parado e as oportunidades abertas do produto (§5, §5.4). Traz o resumo e a definição de parado com o prazo usado. |
| `ver_meta` | `vendedor`, `mes`, `ano` (só 2026) | Com vendedor e mês: aquele mês. Só com vendedor: os 12 meses. Só com mês, ou sem nada: todos os vendedores e a empresa naquele mês (padrão: o mês da data-base). Mês sem meta é "sem meta" (§3.1); mês em andamento é "parcial até dd/mm" (§4.5). |
| `listar_oportunidades` | `vendedor`, `etapa`, `vencidas`, `limite` | Oportunidades do CRM com cliente, produto, valor, probabilidade, ponderado, previsão e dias de atraso. Sem etapa, só as abertas; `etapa: "todas"` inclui as fechadas. As vencidas vêm primeiro (§7). O resumo conta todas, mesmo quando o limite corta a lista. |
| `ver_historico` | `tipo`, `registro`, `limite` | As alterações gravadas, da mais recente para a mais antiga: quando, quem, ação, registro, campo, valor anterior e novo. |

Valores em reais vêm com 2 casas. O atingimento vem como fração (`1.1489`) e
como texto (`114,9%`).

## Ferramentas de escrita (§11)

| Ferramenta | Parâmetros | O que prepara |
|---|---|---|
| `inativar_vendedor` | `vendedor`, `data_desligamento` | Status `Inativo` e a data do desligamento. A data não pode ser anterior à admissão, à última venda nem à última oportunidade dele. As metas dos meses seguintes viram "sem meta", as oportunidades abertas viram órfãs e ele perde o acesso ao painel (§11.5). |
| `reativar_vendedor` | `vendedor` | Status `Ativo` e data de desligamento em branco. As metas não voltam sozinhas. Se ele não tiver conta no painel, rode `app/criar_usuarios.ps1 -Completar` (§11.6). |
| `alterar_meta` | `vendedor`, `mes`, `valor` **ou** `sem_meta` | Meta de um mês **aberto**: do mês da data-base de vendas até dezembro. Valor maior que zero, com até 2 casas (`455000` ou `"455.000,50"`). Com `sem_meta`, o mês volta a "sem meta" (§11.7). |
| `cadastrar_ou_editar_produto` | `produto_id` (vazio = novo), `nome`, `categoria`, `marca`, `modelo`, `preco_tabela`, `custo_unitario`, `unidade`, `status` | Produto novo com o próximo código livre (exige nome, categoria existente, preço e unidade) ou a edição só dos campos informados (§11.8). |
| `transferir_oportunidades` | `de_vendedor`, `para_vendedor`, `oportunidades` (vazio = todas as abertas) | Novo dono, ativo, para oportunidades **abertas** (§11.9). É assim que as órfãs de um desligado são reatribuídas (§2.4). |
| `confirmar_alteracao` | `codigo` | **Grava** a alteração preparada. |

### Como a confirmação funciona

1. Você pede em português ("inativa o Rafael no dia 31/08").
2. O Claude chama a ferramenta, que **só prepara**. Ela devolve o que vai mudar,
   campo a campo (**valor atual → valor novo**), os efeitos (órfãs, metas que
   saem, perda de acesso) e um **código**. Nada é gravado.
3. O Claude mostra o resumo e pergunta se você confirma.
4. Só depois do seu **"sim"** ele chama `confirmar_alteracao` com o código. A
   confirmação refaz todas as validações. Se os dados mudaram nesse meio tempo e
   a alteração ficou diferente, nada é gravado e aparece o resumo novo.

Três travas garantem que ninguém grava sem querer:

- **O código vale uma vez, por 15 minutos**, só para a conta que o pediu.
- **Permissão por ferramenta no Claude Desktop.** Na primeira vez que o Claude
  quiser usar `confirmar_alteracao`, o app pergunta. **Escolha "Permitir uma
  vez", nunca "Permitir sempre"**, para essa ferramenta: assim cada gravação
  passa por você. Para as outras, que não gravam nada, "Permitir sempre" é
  seguro.
- **Pergunta direta à pessoa.** Se o Claude Desktop tiver o recurso de
  perguntar ao usuário pelo servidor (*elicitation* no protocolo MCP), a
  confirmação aparece numa janela do próprio app, e o "sim" vem de você, não do
  Claude.

### Quem pode escrever

A conta que escreve é a de `HORIZONTE_USUARIO`, na configuração do Claude
Desktop. Ela é conferida em `app/usuarios.json` a cada pedido.

| Perfil | Pode |
|---|---|
| `diretoria` | Todas as alterações. |
| `gerente` | Vendedores, metas e oportunidades da **própria filial** (nas transferências, os dois vendedores); produtos, em qualquer filial. |
| `vendedor`, `leitura` ou sem conta configurada | Só ler. |

As contas de gerente (`gerente.cascavel`, `gerente.chapeco` e
`gerente.passofundo`) são criadas por `app/criar_usuarios.ps1 -Completar`
([PAINEL.md](PAINEL.md), "Contas").

**Limite desta versão:** o servidor roda só no computador de quem usa o Claude
Desktop e confia na conta configurada ali, sem pedir senha. Quem pode editar
essa configuração também pode editar as planilhas da pasta, então a trava real
é o acesso ao computador. Cada computador deve configurar a conta da própria
pessoa.

### Onde fica o registro

Cada gravação acrescenta uma linha por campo em
`dados/importacoes/alteracoes.csv`: quando, lote (o código), login, nome,
perfil, ação, tipo, registro, campo, valor anterior e valor novo. O arquivo só
recebe acréscimos, e as planilhas originais não mudam. O painel aplica o
registro por cima delas, e a tela **Histórico** (`/historico`) mostra tudo
(§11.3, §11.4). Para desfazer, faça outra alteração (ex.: reativar quem foi
inativado). Ela também entra no histórico.

## Conectar no Claude Desktop

1. **Descubra o caminho do Node.** No Terminal, rode `which node`. O Claude
   Desktop não enxerga o `PATH` do Terminal; por isso a configuração usa o
   caminho completo (ex.: `/usr/local/bin/node`).
2. **Feche o Claude Desktop** (Cmd+Q). Ele grava preferências no arquivo de
   configuração e pode sobrescrever uma edição feita com ele aberto.
3. **Faça uma cópia** de `~/Library/Application Support/Claude/claude_desktop_config.json`
   (no Windows, `%APPDATA%\Claude\claude_desktop_config.json`).
4. **Acrescente a chave `mcpServers`** logo depois da primeira `{`, sem apagar o
   resto do arquivo, e troque os caminhos pelos da sua máquina:

   ```json
     "mcpServers": {
       "horizonte-maquinas": {
         "command": "/usr/local/bin/node",
         "args": ["/caminho/para/painel-horizonte/mcp/servidor.js"],
         "env": {
           "HORIZONTE_DADOS": "/caminho/para/painel-horizonte/dados",
           "HORIZONTE_USUARIO": "diretoria",
           "HORIZONTE_ARQUIVO_USUARIOS": "/caminho/para/painel-horizonte/app/usuarios.json"
         }
       }
     },
   ```

   - Sem `HORIZONTE_USUARIO`, o servidor só lê.
   - Um gerente põe o próprio login (ex.: `gerente.cascavel`).
   - No Windows, use barras duplas: `"C:\\Users\\voce\\painel-horizonte\\mcp\\servidor.js"`.
5. **Abra o Claude Desktop.**
6. **Confira:** numa conversa nova, clique no botão de ferramentas, abaixo da
   caixa de mensagem. O servidor `horizonte-maquinas` aparece com 12
   ferramentas.
   - Pergunte: "Qual o atingimento da Mariana no ano?". Com a base de jan–ago, a resposta certa é **114,9%** (§6).
   - Para testar a escrita sem gravar: "Prepara a meta da Camila de dezembro para R$ 310.000". Veja o resumo e diga "não".

## Se não aparecer

- **Log:** no macOS, `~/Library/Logs/Claude/mcp-server-horizonte-maquinas.log`.
  Com tudo certo, ele mostra `[horizonte-mcp] pronto; dados em …; escrita como 'diretoria' …`.
- **Teste fora do Claude Desktop**, a partir da raiz do projeto (deve listar
  as 12 ferramentas):

  ```
  printf '%s\n' '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp/servidor.js
  ```

- **Erros comuns:**
  - caminho do `node` ou do `servidor.js` errado;
  - JSON da configuração inválido (vírgula sobrando);
  - Claude Desktop não reiniciado;
  - `HORIZONTE_USUARIO` com um login que não existe em `usuarios.json`: a leitura funciona, e a escrita recusa dizendo isso.

## Testes

- `pwsh -NoProfile -File testes/conferir_mcp.ps1`: a leitura. Compara cada
  resposta com as regras da versão PowerShell e com os números da §6.
- `pwsh -NoProfile -File testes/conferir_escrita_mcp.ps1`: a escrita, com
  contas de diretoria, dois gerentes, um vendedor e sem conta. Confere:
  - as recusas de permissão e de regra;
  - que nada é gravado antes da confirmação;
  - o código de uso único;
  - a confirmação com dados que mudaram no meio do caminho;
  - a pergunta direta à pessoa (aceita e recusada);
  - o registro linha a linha;
  - que a versão PowerShell aplica o registro e que a base pura continua a da §6.
- `pwsh -NoProfile -File testes/paridade_node.ps1`: roda também um cenário com
  alterações. A tela Histórico e todas as outras telas precisam sair idênticas
  nas duas versões do painel.

Todos rodam em cópias temporárias de `dados/`.
