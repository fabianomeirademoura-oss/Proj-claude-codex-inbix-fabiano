# Manual do painel — Horizonte Máquinas

Aplicação web com login: vendas e metas (jan–ago/2026), pipeline, funil e
estoque. As regras de negócio seguem o [REGRAS_NEGOCIO.md](../REGRAS_NEGOCIO.md).
A estrutura do projeto e o estado de cada funcionalidade estão no
[README.md](../README.md).

Caminhos neste manual são relativos à raiz do projeto (`painel-horizonte/`).

## Como rodar

**No Windows**, pelos atalhos em `app\` (passos abaixo). **No macOS**, com
PowerShell 7, a partir da raiz: `pwsh -NoProfile -File app/criar_usuarios.ps1`
e depois `pwsh -NoProfile -File app/servidor.ps1 -Porta 8080`.

1. **Primeira vez apenas:** dê dois cliques em `app\criar_usuarios.cmd`.
   - Cria `usuarios.json`, que guarda só o hash das senhas.
   - Grava os logins e senhas no `.env`, na raiz do projeto. Nada disso aparece na tela.
   - O `.env` está no `.gitignore`: **não versione nem envie por chat ou e-mail.**
2. Dê dois cliques em `app\iniciar.cmd` e abra http://localhost:8080.
   - Para usar outra porta: `iniciar.cmd -Porta 9000`.
3. Para parar o servidor, use Ctrl+C na janela do servidor.

**Onde ficam as planilhas:** o servidor lê de `dados\`, na raiz do projeto. Para
usar outra pasta, passe `-DirDados <pasta>` ou defina a variável
`HORIZONTE_DADOS`. A subpasta `dados\atualizacoes\` **não** é lida (veja o
README, "Estágio atual").

**Recarga automática:** se uma planilha mudar, a próxima página carregada já
reflete a mudança. Se a planilha tiver erro (vendedor inexistente, status
inválido, venda depois do desligamento, ID duplicado…), o painel mostra a lista
de erros no lugar dos números.

## Contas

| Login | Perfil |
|---|---|
| `diretoria` | Diretoria comercial |
| e-mail de cada vendedor **ativo** (ex.: `joao.almeida@horizontemaquinas.com.br`) | Vendedor; vê a própria linha destacada |
| `teste` (senha `teste`, divulgada aos alunos) | Leitura: vê todas as telas, não importa. Só existe na versão web (`web/contas_publicas.json`) |

- Todos veem o painel completo.
- Vendedor desligado não tem conta e, mesmo que tivesse, o login seria recusado (REGRAS_NEGOCIO.md §2.2).
- Cinco senhas erradas seguidas bloqueiam o login por 5 minutos. A conta `teste` não bloqueia: a senha dela é pública e o bloqueio travaria a turma inteira.
- A sessão dura 8 horas.
- Para gerar novas senhas para todos: `criar_usuarios.cmd -Recriar`. Isso reescreve o `.env` e invalida as senhas anteriores.
- Editar o `.env` à mão **não** muda nenhuma senha: o servidor só confere o hash em `usuarios.json`.

## No celular

O painel se ajusta à tela do celular:
- O menu vira uma fileira de abas que rola de lado.
- Os filtros ficam em duas colunas.
- As tabelas largas rolam de lado dentro do próprio bloco.

Ele também pode ser **instalado na tela inicial** como aplicativo (PWA). Ganha ícone próprio e abre em tela cheia, sem a barra do navegador:

- **iPhone (Safari):** abra o endereço, toque em **Compartilhar** e depois em **Adicionar à Tela de Início**.
- **Android (Chrome):** abra o endereço, toque no menu **⋮** e depois em **Instalar app** (ou **Adicionar à tela inicial**).

O aplicativo instalado mostra sempre os números atuais, porque as telas vêm da internet a cada abertura. Ele guarda só o visual e os ícones. Sem conexão, aparece o aviso "Sem conexão", e nenhum dado fica gravado no telefone. No iPhone, o app instalado tem o próprio login, separado do Safari.

## Estrutura

Em `app/`:

| Arquivo | Papel |
|---|---|
| `lib/Xlsx.ps1` | Lê as abas do .xlsx direto do arquivo, sem Excel. |
| `lib/Regras.ps1` | **Regras de vendas:** venda realizada, meta do período, atingimento e rankings. |
| `lib/RegrasPipeline.ps1` | **Regras do pipeline (REGRAS_NEGOCIO.md §7):** oportunidade aberta, pipeline, ponderado, atrasada e órfã. Lê `crm_oportunidades.xlsx` e `clientes.xlsx`. |
| `lib/PaginasPipeline.ps1` | HTML e CSV da página "Pipeline". Só formata. |
| `lib/PaginasFunil.ps1` | HTML da página "Funil" (desenho em SVG, sem JavaScript). Só formata; as regras do funil ficam em `RegrasPipeline.ps1` (`Get-VisaoFunil`). |
| `lib/RegrasEstoque.ps1` | **Regras de estoque (REGRAS_NEGOCIO.md §5):** valor imobilizado, parado (prazo configurável), sem estoque e abaixo do mínimo. Lê `produtos.xlsx` e a foto `estoque_AAAA-MM-DD.xlsx` mais recente. |
| `lib/PaginasEstoque.ps1` | HTML e CSV da página "Estoque". Só formata. |
| `lib/Auth.ps1` | Senhas (PBKDF2-SHA256, 120 mil iterações), sessões e bloqueio. |
| `lib/Paginas.ps1` | HTML e CSV. Só formata. |
| `lib/RegrasImportacao.ps1` | **Regras da importação (REGRAS_NEGOCIO.md §9):** colunas obrigatórias, validação, resumo (novas, atualizadas, ignoradas; foto substituída), pendência até a confirmação e gravação em `dados/importacoes/`. |
| `lib/PaginasImportacao.ps1` | HTML da tela "Importar". Só formata. |
| `servidor.ps1` | Rotas HTTP, restritas a `localhost`. |
| `iniciar.cmd`, `criar_usuarios.cmd/.ps1` | Atalhos para subir o servidor e criar as contas. |

Em `testes/` (cada `.ps1` tem um `.cmd` com o mesmo nome para o Windows, menos o e2e):

| Arquivo | Papel |
|---|---|
| `conferir.ps1` | Recalcula vendas e metas a partir das planilhas brutas e compara com o painel e com o REGRAS_NEGOCIO.md §6. |
| `conferir_pipeline.ps1` | Recalcula o pipeline, a meta e o realizado por vendedor e o funil a partir das planilhas brutas e compara com as páginas "Pipeline" e "Funil". |
| `conferir_estoque.ps1` | Recalcula o estoque a partir das planilhas brutas, em quatro prazos de "parado", e compara com a página "Estoque". |
| `conferir_importacao.ps1` | Importa numa cópia temporária de `dados/` e confere resumo, gravação, reimportação sem duplicar, correção de venda, substituição da foto de estoque e recusas (coluna faltando, ID repetido, foto mais antiga…). |
| `XlsxTeste.ps1` | Só para os testes: grava planilhas com defeitos de propósito. |
| `teste_e2e.ps1` | Sobe o servidor com contas temporárias, numa cópia temporária de `dados/`, e testa login, permissões, telas, CSV e importação por HTTP. |

## Como conferir cada número contra a fonte

1. **Automático:** rode `testes\conferir.cmd`.
   - São 212 comparações, em três períodos (jan–ago, mar–mai e jul–ago).
   - Cada vendedor e cada total é recalculado por somas diretas sobre as linhas das planilhas.
   - Deve terminar com "Todas as 212 conferências bateram".
2. **No Excel:** use as fórmulas abaixo.

| Número na tela | Onde | Fórmula (Excel em português) | Valor jan–ago |
|---|---|---|---|
| Realizado | vendas_2026_jan-ago.xlsx | `=SOMASES(O:O;R:R;"Faturada")` | R$ 45.295.119,16 |
| Nº de vendas | vendas | `=CONT.SES(R:R;"Faturada")` | 723 |
| Canceladas | vendas | `=SOMASES(O:O;R:R;"Cancelada")` e `=CONT.SES(R:R;"Cancelada")` | R$ 407.327,39 / 12 |
| Meta do período | metas_2026.xlsx | `=SOMASES(E:E;D:D;"<=8")` | R$ 50.720.000,00 |
| Atingimento da empresa | — | realizado ÷ meta | 89,3% |
| Realizado do vendedor | vendas | `=SOMASES(O:O;C:C;"V001";R:R;"Faturada")` | João: R$ 5.519.186,26 |
| Meta do vendedor | metas | `=SOMASES(E:E;A:A;"V001";D:D;"<=8")` | João: R$ 5.235.000,00 |

Na planilha de vendas, a coluna O é Valor Total, R é Status, C é ID Vendedor e
B é Data. Na de metas, A é ID Vendedor, D é Mês e E é Meta.

**Outros períodos:** acrescente o critério de data. Exemplo para mar–mai:

`=SOMASES(O:O;C:C;"V001";R:R;"Faturada";B:B;">="&DATA(2026;3;1);B:B;"<"&DATA(2026;6;1))`

Na planilha de metas, use `D:D;">=3";D:D;"<=5"`.

3. **Venda a venda:** clique no nome de um vendedor.
   - A página mostra o resultado mês a mês e a lista de vendas, cada uma com o **número da linha** na planilha.
   - As canceladas aparecem riscadas, com a marca "fora do realizado".
4. **Lado a lado:** o botão "Baixar CSV" exporta o quadro de conferência (separador `;`, vírgula decimal) para abrir no Excel ao lado das planilhas.

### Pipeline (página "Pipeline")

A página "Pipeline" mostra uma linha por vendedor com a meta do período, o realizado, o pipeline aberto e o pipeline ponderado. Embaixo, lista as oportunidades abertas, com as atrasadas em destaque.

- A meta e o realizado seguem o período escolhido.
- O pipeline é a foto do CRM de 31/08/2026 e **não** muda com o período (REGRAS_NEGOCIO.md §7.3).
- Aberta é toda oportunidade cuja etapa não é "Fechada Ganha" nem "Fechada Perdida".
- Atrasada é a oportunidade aberta com previsão de fechamento anterior a 31/08/2026.
- Vendedores desligados aparecem, com o alerta de carteira órfã, para que o total feche.
- Clique no nome de um vendedor para filtrar a lista de oportunidades.

**Automático:** `testes\conferir_pipeline.cmd` faz 608 comparações (pipeline e funil) e deve terminar com "Todas as 608 conferências do pipeline bateram".

**No Excel** (crm_oportunidades.xlsx: J = Valor Estimado, K = Etapa, L = Probabilidade, M = Previsão de Fechamento, C = ID Vendedor):

| Número na tela | Fórmula | Valor |
|---|---|---|
| Abertas | `=CONT.SES(K2:K260;"<>Fechada Ganha";K2:K260;"<>Fechada Perdida")` | 68 |
| Pipeline aberto | `=SOMASES(J2:J260;K2:K260;"<>Fechada Ganha";K2:K260;"<>Fechada Perdida")` | R$ 13.870.500,00 |
| Ponderado | `=SOMARPRODUTO((K2:K260<>"Fechada Ganha")*(K2:K260<>"Fechada Perdida")*J2:J260*L2:L260)` | R$ 5.605.495,00 |
| Atrasadas | `=CONT.SES(K2:K260;"<>Fechada Ganha";K2:K260;"<>Fechada Perdida";M2:M260;"<"&DATA(2026;8;31))` | 29 |
| Pipeline de um vendedor | acrescente `;C2:C260;"V001"` ao SOMASES | João: R$ 2.724.300,00 |

O botão "Baixar CSV" exporta as oportunidades abertas com o valor, o ponderado, os dias de atraso, a marca de órfã e o número da linha na planilha.

### Estoque (página "Estoque")

Responde três perguntas sobre a foto de estoque (hoje, `estoque_2026-08-31.xlsx`).
Tudo é calculado por linha, isto é, **por produto e por filial**.

1. **Quanto dinheiro está parado no pátio, por filial e por categoria.** É o valor imobilizado (quantidade × custo médio) numa tabela de categoria × filial. Em vermelho aparece a parte que está parada. A categoria vem de `produtos.xlsx`.
2. **Produtos parados.** A definição usada aparece escrita, numa frase, logo acima da lista. O **prazo é um campo no topo da página** (padrão de 180 dias, REGRAS_NEGOCIO.md §5). Mudar o prazo muda a lista, o cartão, a tabela e o CSV, sem mexer no código. Cada linha mostra as oportunidades abertas do produto e a situação do mesmo produto nas outras filiais.
3. **Sem estoque em lugar nenhum.** São os produtos com soma zero nas 3 filiais. Os que têm oportunidade aberta no CRM aparecem em destaque.
4. **Abaixo do mínimo.** São as linhas com quantidade menor que o mínimo da filial, com quantas unidades faltam e com o saldo das outras filiais (em negrito, onde há sobra).

- O filtro "Filial nas listas" vale para as listas de parados e de abaixo do mínimo.
- A foto vale na data do nome do arquivo. Havendo mais de uma `estoque_*.xlsx` na pasta de dados, vale a mais recente.
- O botão "Baixar CSV" exporta as 195 linhas com as marcas de parado, abaixo do mínimo e sem estoque, além do número da linha na planilha.

**Automático:** `testes\conferir_estoque.cmd` faz 205 comparações e deve terminar com "Todas as 205 conferências do estoque bateram".

**No Excel** (estoque_2026-08-31.xlsx: D = Filial, E = Quantidade, F = Estoque Mínimo, G = Custo Médio):

| Número na tela | Fórmula | Valor |
|---|---|---|
| Valor imobilizado | `=SOMARPRODUTO(E2:E196;G2:G196)` | R$ 18.420.192,72 |
| Imobilizado de uma filial | `=SOMARPRODUTO((D2:D196="Cascavel")*E2:E196*G2:G196)` | R$ 6.433.787,48 |
| Abaixo do mínimo | `=SOMARPRODUTO(--(E2:E196<F2:F196))` | 28 |
| Parados (180 dias) | ver o CSV, coluna "Parado" | 7 linhas, R$ 1.693.610,48 |

### Funil de vendas (página "Funil")

É uma representação visual das oportunidades abertas por etapa (Prospecção → Qualificação → Proposta Enviada → Negociação), com filtros por **vendedor**, **filial** e **categoria de produto**. As regras estão no REGRAS_NEGOCIO.md §8.

- **Cada barra** mostra as oportunidades abertas que estão **hoje** naquela etapa. A parte vermelha é a atrasada, isto é, a previsão de fechamento já passou.
- **Largura:** escolha se a largura das barras mede a quantidade de oportunidades ou o valor estimado.
- **Limite do CRM:** a planilha guarda só a etapa atual. Por isso, o funil não mostra quantas oportunidades *passaram* por cada etapa, e a tela avisa isso.
- **Fechadas:** ficam à parte, com ganhas × perdidas, a conversão (ganhas ÷ fechadas; 49,7% sem filtro) e os motivos de perda.
- **Filial:** é a filial do **vendedor**, não a do cliente.
- **Categoria:** vem de `produtos.xlsx`. Peças não passam pelo CRM, por isso não aparecem no filtro.
- **Filtros combinados:** os filtros se somam, e o endereço da página guarda a escolha. Exemplo: `/funil?filial=F02&categoria=Tratores&medida=valor`.
- **Lista das abertas:** embaixo do desenho, com as atrasadas em destaque. Clicar num vendedor filtra o funil por ele.

Sem filtro: Prospecção 16, Qualificação 13, Proposta Enviada 22, Negociação 17 (68 abertas, 29 atrasadas); 95 ganhas e 96 perdidas.


### Origem do faturamento (página "Origem do faturamento")

Acesse `/origem-vendas` pelo menu. O filtro usa o mês da venda. Cada venda é
ligada à oportunidade exclusivamente pelo `ID Oportunidade`; a tabela de
conferência mostra os IDs e as linhas das duas planilhas.

Participação do CRM = valor faturado com oportunidade vinculada / faturamento
total do período (REGRAS_NEGOCIO.md §1). Canceladas ficam visíveis na conferência,
mas fora dos indicadores. O valor usado é o da venda, não a estimativa do CRM.

Em jan–ago: 95 vendas vinculadas, R$ 15.301.000,00 (33,8%); 628 vendas sem
oportunidade registrada, R$ 29.994.119,16 (66,2%). Nesta base, as vendas sem
ID são descritas como vendas de balcão; incluem máquinas e peças. A ausência
do vínculo não comprova ausência de trabalho comercial. Não há associação
por semelhança de cliente ou produto, nem baixa automática de oportunidades.

Um ID preenchido que não existe no CRM bloqueia os indicadores desta página
com a indicação do problema; não é tratado como venda sem oportunidade.

Verificação adicional: `pwsh -NoProfile -File testes/conferir_origem.ps1`.
O teste HTTP `testes/teste_e2e.ps1` agora inclui 7 verificações desta página
(total: 56). Para carregar os módulos novos, o servidor precisa ser iniciado
novamente; mudar arquivos não recarrega funções PowerShell já carregadas.

### Importação de arquivos (página "Importar")

Só a conta `diretoria` vê o link **Importar** e consegue enviar arquivos. Um
vendedor que tente abrir `/importar` recebe "Sem acesso". As regras estão no
REGRAS_NEGOCIO.md §9.

1. **Escolha o tipo e o arquivo.**
   - **Vendas:** aba `Vendas`, com as 18 colunas da planilha de vendas.
   - **Estoque:** aba `Estoque`, com as 9 colunas, e o nome `estoque_AAAA-MM-DD.xlsx`, porque a data da foto vem do nome.
2. **Analisar arquivo.** Nada é gravado nesta etapa. A tela mostra o resumo:
   - **Vendas:** quantas linhas são **novas**, quantas são **atualizadas** (mesmo `ID Venda` com algum valor diferente; a tela lista o que muda, coluna por coluna) e quantas são **ignoradas** (já existem iguais). Mostra também como ficam o total de vendas e a data-base.
   - **Estoque:** a foto em vigor e a nova, quantas linhas são **substituídas**, quantas quantidades mudam e como ficam o valor imobilizado, os produtos sem estoque e as linhas abaixo do mínimo.
3. **Confirmar** ou **Cancelar.**
   - A confirmação refaz a validação antes de gravar.
   - O arquivo é guardado sem alteração em `dados/importacoes/vendas/` ou `dados/importacoes/estoque/`, e o painel recarrega sozinho.
   - As planilhas de `dados/` nunca são alteradas.

**Recusas.** O arquivo inteiro é recusado, e nada é gravado, quando:
- falta uma coluna obrigatória (a tela diz exatamente quais);
- a aba está errada;
- o arquivo não é um .xlsx;
- há `ID Venda` repetido no arquivo;
- há vendedor, filial ou status inválido;
- há venda depois do desligamento do vendedor;
- a foto de estoque é mais antiga que a em vigor;
- o nome do arquivo de estoque não segue o padrão.

**Histórico e desfazer.**
- Cada importação confirmada fica em `dados/importacoes/registro.csv` e aparece no histórico da tela.
- Para desfazer, retire o arquivo da pasta de importações.
- Arquivos enviados e não confirmados ficam em `dados/importacoes/pendentes/` e são descartados depois de 24 h. A confirmação só vale por 2 h.

**Aula 6:** importe `dados/atualizacoes/vendas_2026_setembro.xlsx` como vendas (82 novas, R$ 3.850.268,28 faturados) e `dados/atualizacoes/estoque_2026-09-19.xlsx` como estoque (195 linhas substituídas, 61 quantidades mudam). Importar setembro de novo dá 0 novas e 82 ignoradas.

Verificação: `pwsh -NoProfile -File testes/conferir_importacao.ps1` (70 conferências). O e2e passa a ter 73 verificações.
