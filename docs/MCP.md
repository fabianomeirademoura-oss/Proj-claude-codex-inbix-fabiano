# Servidor MCP: o painel dentro do Claude Desktop

O servidor MCP deixa o Claude Desktop consultar os dados da Horizonte Máquinas
direto na conversa ("como está a Mariana no ano?", "o que está parado em
Chapecó?"). Ele **só lê**: não grava, não importa e não altera nenhuma planilha.

- **Onde está:** `mcp/servidor.js` (protocolo) e `mcp/ferramentas.js` (as cinco ferramentas).
- **Do que precisa:** Node 20 ou mais novo. Não tem dependências nem `npm install`.
- **Regras:** as mesmas do painel. O servidor usa `web/lib/`, o espelho das
  regras da versão PowerShell, e não cria cálculo novo. Cada resposta diz a
  data-base que usou e cita a seção do [REGRAS_NEGOCIO.md](../REGRAS_NEGOCIO.md).
- **Dados:** lê `dados/` e, por cima, as importações confirmadas em
  `dados/importacoes/` (§9.8), como o painel local. Outra pasta:
  variável `HORIZONTE_DADOS`. Se uma planilha mudar, a próxima pergunta já vê a
  mudança. Se ela tiver erro de validação, a ferramenta devolve a lista de erros
  no lugar dos números.

## As cinco ferramentas

Todas recebem parâmetros simples e devolvem JSON. Vendedor, produto e filial
podem ser informados pelo ID ou pelo nome, inteiro ou em parte, sem acento. O
nome vira ID antes de qualquer conta (§0.1). Nome ambíguo ou inexistente volta
como erro, com a lista de opções, e nunca com um palpite.

| Ferramenta | Parâmetros | O que devolve |
|---|---|---|
| `consultar_vendedor` | `vendedor` (obrigatório), `mes_inicial`, `mes_final` | Cadastro (filial, status, admissão, desligamento), meta, realizado, atingimento, falta para a meta, faturamento, ticket médio, canceladas e a abertura mês a mês. Sem período, usa o acumulado do ano até a data-base de vendas (§4). |
| `buscar_produto` | `busca` (código ou palavras do nome, marca ou modelo), `categoria`, `limite` | Cadastro do produto, com preço de tabela, status (descontinuado aparece marcado, §5.5) e unidades em estoque na foto em vigor. |
| `ver_estoque` | `produto`, `filial`, `dias_parado` (padrão 180), `somente_parados` | Linhas (produto, filial) da foto em vigor, com quantidade, mínimo, valor imobilizado, últimas entrada e saída, dias parado e as oportunidades abertas do produto (§5, §5.4). Traz o resumo e a definição de parado com o prazo usado. |
| `ver_meta` | `vendedor`, `mes`, `ano` (só 2026) | Vendedor e mês: aquele mês. Só vendedor: os 12 meses. Só mês, ou nada: todos os vendedores e a empresa naquele mês (padrão: o mês da data-base). Mês sem meta é "sem meta" (§3.1); mês em andamento é "parcial até dd/mm" (§4.5). |
| `listar_oportunidades` | `vendedor`, `etapa`, `vencidas`, `limite` | Oportunidades do CRM com cliente, produto, valor, probabilidade, ponderado, previsão e dias de atraso. Sem etapa, só as abertas; `etapa: "todas"` inclui as fechadas. As vencidas vêm primeiro (§7). O resumo conta todas, mesmo quando a lista é cortada pelo limite. |

Valores em reais vêm com 2 casas. O atingimento vem como fração (`1.1489`) e
como texto (`114,9%`).

## Conectar no Claude Desktop

1. **Descubra o caminho do Node.** No Terminal, rode `which node`. O Claude
   Desktop não enxerga o `PATH` do Terminal; por isso a configuração usa o
   caminho completo (ex.: `/usr/local/bin/node`).
2. **Abra a configuração.** No Claude Desktop: menu **Claude → Configurações…
   → Desenvolvedor → Editar configuração**. Isso abre o arquivo
   `~/Library/Application Support/Claude/claude_desktop_config.json` (no
   Windows, `%APPDATA%\Claude\claude_desktop_config.json`).
3. **Acrescente o servidor** dentro de `"mcpServers"`, trocando os caminhos
   pelos da sua máquina. Se o arquivo já tiver outras chaves, mantenha-as e só
   acrescente `"mcpServers"`:

   ```json
   {
     "mcpServers": {
       "horizonte-maquinas": {
         "command": "/usr/local/bin/node",
         "args": ["/caminho/para/painel-horizonte/mcp/servidor.js"],
         "env": { "HORIZONTE_DADOS": "/caminho/para/painel-horizonte/dados" }
       }
     }
   }
   ```

   No Windows, use barras duplas: `"C:\\Users\\voce\\painel-horizonte\\mcp\\servidor.js"`.
4. **Feche o Claude Desktop por completo** (Cmd+Q, não só a janela) e abra de
   novo.
5. **Confira:** numa conversa nova, clique no botão de ferramentas (o ícone de
   controles, abaixo da caixa de mensagem). O servidor `horizonte-maquinas`
   aparece com as cinco ferramentas. Pergunte, por exemplo: "Qual o atingimento
   da Mariana no ano?". A resposta certa, com a base de jan–ago, é **114,9%**
   (R$ 6.474.285,51 de R$ 5.635.000,00, §6).

Na primeira vez que usar cada ferramenta, o Claude Desktop pede permissão. Como
todas são só de leitura, dá para escolher "Permitir sempre".

## Se não aparecer

- **Log:** no macOS, `~/Library/Logs/Claude/mcp-server-horizonte-maquinas.log`.
  Com tudo certo, ele mostra `[horizonte-mcp] pronto; dados em …`.
- **Teste fora do Claude Desktop**, a partir da raiz do projeto (deve listar
  as cinco ferramentas):

  ```
  printf '%s\n' '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp/servidor.js
  ```

- **Erros comuns:** caminho do `node` ou do `servidor.js` errado; JSON da
  configuração inválido (vírgula sobrando); Claude Desktop não reiniciado.

## Testes

`pwsh -NoProfile -File testes/conferir_mcp.ps1` sobe o servidor, fala com ele
pelo protocolo e compara cada resposta com as regras da versão PowerShell e com
os números da §6: todos os vendedores em três períodos, todas as metas de
janeiro a agosto, o pipeline de cada vendedor, as 195 linhas de estoque, os
filtros e os erros de entrada. Roda numa cópia temporária de `dados/`, sem as
importações (§9.9).

## Próximos passos (fora desta versão)

- Ferramentas que escrevem (importar planilha, reatribuir órfã) ficam para
  depois, com as mesmas regras de permissão do painel (§9.1, §2.4).
- Hoje o servidor não tem login: quem tem a pasta do projeto no computador vê
  tudo. Ele roda só na máquina local, pelo Claude Desktop.
