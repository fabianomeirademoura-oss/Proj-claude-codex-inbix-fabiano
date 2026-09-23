# REGRAS_NEGOCIO.md — Contrato de regras de negócio da Horizonte Máquinas

Este arquivo define as regras que valem para **tudo** o que for construído sobre
os dados da Horizonte Máquinas: dashboards, relatórios, APIs, importadores e
respostas a perguntas. É a **fonte única** das regras, para o Claude Code
(`CLAUDE.md`) e para o Codex (`AGENTS.md`).

- **Se o código e este arquivo divergirem, este arquivo vence.**
- Mudar uma regra exige editar este arquivo primeiro e só depois o código.
  Mudança de regra é combinada antes com o coordenador do projeto (`COORDENACAO.md`).
- Quando uma pergunta de negócio não estiver coberta aqui, **pergunte antes de inventar um critério**.
- Algumas regras descrevem requisitos que ainda não estão implementados (por
  exemplo, as importações de setembro). O que já existe e o que está previsto
  estão no [README.md](README.md#estágio-atual).

**Onde estão os dados:** `dados/`, na raiz do projeto. O dicionário e os casos
de borda estão em [PERFIL.md](PERFIL.md); ele descreve os dados, mas não cria
regras.

---

## 0. Princípios gerais

1. **A chave é sempre o ID** (`ID Vendedor`, `ID Cliente`, `ID Produto`,
   `ID Oportunidade`, `ID Venda`). Os nomes repetidos nas planilhas servem
   só para exibição. Nunca filtre, agrupe ou ligue tabelas por nome.
2. **Filial é normalizada na importação.** Converta o nome (`Cascavel`,
   `Chapecó`, `Passo Fundo`) para `ID Filial` (`F01`, `F02`, `F03`) usando a
   aba `Filiais`. Um nome de filial desconhecido é **erro de importação**; não
   crie uma filial nova.
3. **Data-base, não relógio.**
   - Todo cálculo relativo a "hoje" usa a **data-base dos dados** envolvidos, nunca a data do sistema. Isso vale para dias parado, oportunidade atrasada e mês em andamento.
   - Datas-base atuais:
     - Vendas: 19/09/2026 (após importar setembro).
       Enquanto setembro não for importado pela tela de importação (§9), o painel carrega só janeiro a agosto e a data-base de vendas é 31/08/2026.
     - Estoque: data da foto importada (31/08/2026 ou 19/09/2026).
     - CRM: 31/08/2026.
   - Toda tela mostra a data-base que está usando.
4. **Dinheiro é decimal com 2 casas.** Nunca use ponto flutuante para somar
   dinheiro. Os arquivos têm resíduo de float (ex.: `8891.280000000001`).
   Arredonde cada valor para 2 casas na importação.
5. **Percentuais:**
   - `Probabilidade` vem como fração de 0 a 1 (0,75 = 75%).
   - Atingimento é guardado como fração e exibido com 1 casa decimal, no formato brasileiro (`89,3%`).
6. **Nada é apagado.**
   - Vendas canceladas, vendedores desligados e produtos descontinuados permanecem na base.
   - As regras abaixo dizem quando cada um entra ou não nos cálculos.

---

## 1. Venda realizada

**Venda realizada** é toda linha da base de vendas com `Status = 'Faturada'`.
Não há outra condição.

| Aspecto | Regra |
|---|---|
| Valor | `Valor Total` da linha (já com desconto), arredondado a 2 casas. |
| Data | `Data` da venda. Define o mês em que a venda conta. |
| Vendedor | `ID Vendedor` da linha. |
| Filial | `Filial` da venda, que é **a filial do vendedor**. Não usar `Clientes.Filial de Atendimento` para medir realizado. |
| Máquinas e peças | As duas contam. Relatórios podem separar por `Categoria` (`Peças e Acessórios` versus o resto), mas o total inclui tudo. |
| Origem | Vendas com e sem `ID Oportunidade` contam igual. O ID só diz se a venda nasceu no CRM. |

**Vendas canceladas** (`Status = 'Cancelada'`) ficam **fora** de:

- realizado e faturamento;
- quantidade de vendas e ticket médio;
- rankings e atingimento;
- unidades vendidas;
- "última venda" ou "última saída" de produto;
- qualquer indicador de giro.

Elas só aparecem em relatórios que tratam explicitamente de cancelamentos.

**Ticket médio** = realizado ÷ número de vendas realizadas (linhas faturadas).

**Importação de vendas** é **incremental**:

- Anexa linhas novas.
- Faz *upsert* por `ID Venda`. Reimportar o mesmo arquivo não duplica nada.
- Um `ID Venda` já existente com dados diferentes **atualiza** a linha. Ex.: uma venda que passou de Faturada para Cancelada sai do realizado.

---

## 2. Vendedor desligado (Status = Inativo)

Hoje: **V011 Ricardo Alves Pereira**, desligado em 30/04/2026.

1. **Histórico preservado.**
   - As vendas e metas dele continuam valendo nos meses em que esteve ativo.
   - Entram normalmente nos totais da empresa e da filial dele.
2. **Sem acesso.** Um vendedor inativo não faz login na aplicação.
3. **Rankings e listas de desempenho** mostram só vendedores **ativos** por padrão:
   - maior faturamento;
   - maior atingimento;
   - "mais distantes da meta";
   - "piores vendedores".
   - Inativos só aparecem com um filtro explícito ("incluir desligados") e com a marcação "desligado em dd/mm/aaaa".
4. **Oportunidades abertas de vendedor inativo** (hoje OP-0130 e OP-0140) são **órfãs**:
   - Continuam no pipeline, com alerta visível.
   - Nunca são apagadas nem reatribuídas automaticamente.
   - A reatribuição é uma ação explícita de um gestor.
5. **Atingimento** do inativo é calculado como o de qualquer vendedor (seção 4), usando só os meses em que ele tinha meta.
6. **Venda ou oportunidade nova** com data posterior ao desligamento, atribuída a um inativo, é **erro de importação**.

---

## 3. Vendedor admitido no meio do ano

Hoje: **V012 Camila Rodrigues Teixeira**, admitida em 01/07/2026, com metas de
jul a dez e metas menores no início (ramp-up).

1. **Ela só tem meta nos meses que existem na planilha de metas.** A ausência de uma linha de meta significa "sem meta", **não meta zero**.
2. **Atingimento** usa apenas os meses em que ela tem meta (seção 4). Nunca compare o realizado dela com a meta de meses anteriores à admissão.
3. **Rankings.** Ela **entra** nos rankings de ativos, com a marcação "admitida em mm/aaaa". O faturamento absoluto dela é naturalmente menor. Se o ranking for por faturamento, a tela deve deixar claro quantos meses de meta cada vendedor teve no período.
4. **Oportunidade criada antes da admissão.**
   - Existe uma: OP-0150, criada em 21/04 e fechada por ela em 01/07. É uma carteira transferida.
   - Isso **não é erro**. O CRM guarda só o dono atual. Não valide "data de criação ≥ admissão".

---

## 4. Atingimento de meta num período

Metas são **mensais**, então todo período é um conjunto de **meses inteiros**
(ex.: jan–ago/2026). Um período é inválido se não estiver alinhado a meses. A
única exceção é o mês em andamento (regra 4.5).

### 4.1 Por vendedor

```
meses_com_meta(v, P)  = meses do período P em que existe linha de meta para v
meta(v, P)            = soma de Meta (R$) de v nesses meses
realizado(v, P)       = soma das vendas realizadas (seção 1) de v nesses mesmos meses
atingimento(v, P)     = realizado(v, P) / meta(v, P)
```

- Numerador e denominador usam **os mesmos meses**.
- Se `meses_com_meta(v, P)` estiver vazio, o atingimento é **"sem meta no período"**. Não é 0%, nem infinito, nem erro.
- Vendas de um vendedor em meses sem meta não entram no atingimento dele. Elas continuam no faturamento da empresa e da filial. Hoje não há nenhum caso assim.

### 4.2 Por filial e da empresa

```
atingimento(grupo, P) = soma de realizado(v, P) dos vendedores do grupo
                        / soma de meta(v, P) dos vendedores do grupo
```

- A **filial do vendedor** vem do cadastro de vendedores.
- **Nunca faça média simples dos percentuais individuais.** Em jan–ago, a média dos percentuais dá cerca de 86,2%, mas o valor certo é 89,3%.
- **Vendedores inativos entram** no atingimento da filial e da empresa, nos meses em que tinham meta. A regra de "só ativos" vale apenas para rankings.

### 4.3 Distância da meta

`gap(v, P) = meta(v, P) − realizado(v, P)`

- Um valor positivo significa que faltou vender.
- Listas de "mais distantes da meta" seguem a regra 2.3 (só ativos).

### 4.4 Erros proibidos

- Dividir o realizado pela **meta anual** quando o período é parcial. Em jan–ago isso daria 60,0% em vez de 89,3%.
- Contar vendas canceladas no realizado.
- Usar meta zero ou meta média para meses sem meta.
- Ligar metas a vendas pelo nome do vendedor.

### 4.5 Mês em andamento

Se o período inclui o mês da data-base de vendas e ele ainda não terminou
(hoje, setembro até 19/09):

- O atingimento usa a **meta cheia do mês**.
- A tela marca o resultado como **"parcial até dd/mm"**.
- Não aplique pró-rata automaticamente. Se um pró-rata for pedido, ele deve ser mostrado como indicador separado, com o critério explícito.

---

## 5. Produto parado

"Parado" é avaliado **por linha de estoque**, ou seja, por (produto, filial), e
nunca pelo produto no agregado.

```
ultimo_movimento = a data mais recente entre Data da Última Entrada e Data da Última Saída
dias_sem_movimento = data_base_do_estoque − ultimo_movimento

PARADO  ⇔  Quantidade > 0  E  dias_sem_movimento ≥ prazo      (prazo padrão = 180 dias)
```

0. **O prazo é configurável, o critério não.**
   - O padrão é 180 dias. Quem usa a tela pode escolher outro prazo (1 a 3650 dias) sem mexer no código; o CSV segue o prazo escolhido.
   - Toda lista de parados mostra, em uma frase, a definição usada com o prazo em vigor e a data da foto.
   - Os números de conferência (§6) são sempre com o prazo padrão.
1. **Quantidade > 0 é obrigatória.** Uma linha zerada nunca está parada: ela está "sem estoque".
2. **Datas em branco:**
   - A saída em branco significa que não houve venda faturada daquele produto naquela filial em 2026.
   - Se as **duas** datas estiverem em branco e a quantidade for > 0, a linha é parada, com "sem movimento registrado".
3. **Data-base** é a da foto de estoque usada (31/08/2026 ou 19/09/2026), nunca o dia de hoje.
4. **Parado não é o mesmo que sem demanda.**
   - Toda lista de parados mostra, ao lado, o número de oportunidades abertas do produto.
   - Ex.: P008 em Cascavel está parado, mas tem 2 oportunidades.
5. **Descontinuado pode estar parado.** P002 e P020 continuam na lista de estoque e de parados. Eles só não aparecem no catálogo de venda.

Termos vizinhos (não confundir):

| Termo | Definição |
|---|---|
| Sem estoque | Soma de `Quantidade` do produto nas 3 filiais = 0. |
| Abaixo do mínimo | `Quantidade < Estoque Mínimo` na linha (produto, filial). |
| Valor imobilizado | Soma de `Quantidade × Custo Médio`. |

Na tela de estoque, **"dinheiro parado no pátio"** é o valor imobilizado de
todo o estoque, por filial e por categoria. A parte que está parada (critério
acima) aparece separada, nunca no lugar do total. A categoria vem de
`produtos.xlsx`, ligada pelo `ID Produto`.

**Importação de estoque** é **substitutiva**:

- Cada arquivo é uma foto completa numa data e substitui a foto anterior inteira.
- Guarde a data da foto. Ela vem do nome do arquivo (`estoque_AAAA-MM-DD.xlsx`). Com mais de uma foto na pasta de dados, vale a mais recente.
- Uma linha que falta na foto conta como quantidade zero.
- É **erro de importação**: `ID Produto` inexistente em `produtos.xlsx`, filial desconhecida, (produto, filial) repetido, quantidade ou mínimo que não sejam inteiros ≥ 0, custo médio inválido e data inválida.
- Recuse uma foto mais antiga que a atual.
- **Nunca** calcule estoque a partir das vendas. Os dois não fecham, porque máquina pode ser faturada sem saldo na filial.

---

## 6. Números de conferência (jan–ago/2026, data-base 31/08/2026)

Qualquer implementação destas regras tem que reproduzir estes valores.

| Indicador | Valor |
|---|---|
| Vendas realizadas (linhas) | 723 |
| Realizado jan–ago | R$ 45.295.119,16 |
| Vendas canceladas | 12 linhas, R$ 407.327,39 (fora do realizado) |
| Meta jan–ago (só meses com meta) | R$ 50.720.000,00 |
| Atingimento da empresa | 89,3% |
| Atingimento V003 Mariana (maior) | 114,9% (R$ 6.474.285,51 / R$ 5.635.000,00) |
| Atingimento V011 Ricardo (inativo, jan–abr) | 62,8% (R$ 1.487.342,56 / R$ 2.370.000,00) |
| Atingimento V012 Camila (jul–ago) | 80,1% (R$ 629.152,07 / R$ 785.000,00) |
| Linhas de estoque paradas (foto 31/08, prazo 180 dias) | 7, R$ 1.693.610,48 |
| Linhas abaixo do mínimo (foto 31/08) | 28 |
| Produtos sem estoque (foto 31/08) | 8 |
| Valor imobilizado (foto 31/08) | R$ 18.420.192,72 |
| Oportunidades abertas / atrasadas (CRM 31/08) | 68 / 29 |
| Pipeline aberto / ponderado (CRM 31/08) | R$ 13.870.500,00 / R$ 5.605.495,00 |
| Pipeline atrasado (CRM 31/08) | R$ 6.813.700,00 |
| Oportunidades órfãs (V011) | 2, R$ 255.300,00 |
| Funil (CRM 31/08): ganhas / perdidas / conversão | 95 / 96 / 49,7% |

---

## 7. Pipeline do CRM

Fonte: `crm_oportunidades.xlsx`, uma foto do CRM na data-base (hoje, 31/08/2026).

```
ABERTA     ⇔  Etapa ∉ { 'Fechada Ganha', 'Fechada Perdida' }
pipeline   =  soma de Valor Estimado das abertas
ponderado  =  soma de (Valor Estimado × Probabilidade) das abertas, cada parcela arredondada a 2 casas
ATRASADA   ⇔  ABERTA  E  Previsão de Fechamento < data-base do CRM
```

1. `Valor Estimado` é o valor **total** da oportunidade (a quantidade já está dentro). Não multiplique pela quantidade.
2. A probabilidade é a da linha, que é a da etapa (fração de 0 a 1, §0.5).
3. **O pipeline não é filtrado pelo período de metas.** É a foto do CRM na data-base. A visão por vendedor mostra lado a lado a meta e o realizado do período escolhido (§4) e o pipeline da foto.
4. O vendedor da oportunidade é o `ID Vendedor` (dono atual, §3.4). Oportunidades de vendedor inativo continuam no pipeline e nos totais, com alerta de órfã (§2.4). Por isso, a visão de pipeline mostra todos os vendedores, não só os ativos.
5. É **erro de importação**: `ID Vendedor`, `ID Cliente` ou `ID Produto` inexistente, `ID Oportunidade` duplicado, valor ou probabilidade inválidos, oportunidade aberta sem previsão de fechamento e oportunidade criada depois do desligamento do dono (§2.6).
6. Não há deduplicação automática nem baixa automática de oportunidades que parecem já ter virado venda (PERFIL.md, caso 8). Isso fica para uma regra futura.

---

## 8. Funil de vendas

Fonte: a mesma foto do CRM da §7. O funil **não** é filtrado pelo período de metas.

```
etapas do funil (de cima para baixo) = Prospecção → Qualificação → Proposta Enviada → Negociação
barra da etapa   = oportunidades ABERTAS cuja etapa atual é aquela (quantidade ou valor estimado)
conversão        = Fechadas Ganhas ÷ (Fechadas Ganhas + Fechadas Perdidas)
```

1. **O CRM guarda só a etapa atual**, sem histórico. Por isso, o funil mostra onde as oportunidades abertas estão **hoje**. Ele não mostra quantas passaram por cada etapa, e a tela diz isso. Não invente taxas de passagem entre etapas.
2. As **fechadas** ficam fora das barras. Aparecem à parte: ganhas, perdidas, a conversão (em quantidade e, separada, em valor) e os motivos de perda.
3. Dentro de cada etapa, a parte **atrasada** (§7) aparece em destaque.
4. **Filtros**, que se combinam (E):
   - **Vendedor:** o `ID Vendedor` da oportunidade (dono atual). Desligados aparecem, marcados, porque as órfãs continuam no pipeline (§2.4).
   - **Filial:** a **filial do vendedor**, pelo cadastro de vendedores, como no realizado (§1) e no atingimento (§4.2). Não é a filial de atendimento do cliente.
   - **Categoria:** a `Categoria` de `produtos.xlsx`, ligada pelo `ID Produto`. Peças não passam pelo CRM.
5. Sem filtro, o funil fecha com a §6: 68 abertas (Prospecção 16, Qualificação 13, Proposta Enviada 22, Negociação 17), 95 ganhas, 96 perdidas e conversão de 49,7%.

---

## 9. Importação de arquivos

Decisões aprovadas pelo coordenador em 23/09/2026. A tela de importação aplica
as semânticas da §1 (vendas incrementais) e da §5 (estoque substitutivo).

1. **Quem importa:** só contas com perfil **diretoria**. Vendedores não veem a
   tela nem conseguem enviar arquivos.
2. **Tipo escolhido na tela:** quem envia diz se o arquivo é de **vendas** (aba
   `Vendas`) ou de **estoque** (aba `Estoque`).
3. **Colunas obrigatórias**, com o nome exato do cabeçalho. Colunas extras são
   ignoradas. Se faltar uma ou mais, o arquivo é **recusado** e a tela diz
   exatamente quais faltaram.
   - Vendas (18): `ID Venda`, `Data`, `ID Vendedor`, `Vendedor`, `Filial`,
     `ID Cliente`, `Cliente`, `Cidade`, `UF`, `ID Produto`, `Produto`,
     `Categoria`, `Quantidade`, `Valor Unitário`, `Valor Total`,
     `Forma de Pagamento`, `ID Oportunidade`, `Status`.
   - Estoque (9): `ID Produto`, `Produto`, `Categoria`, `Filial`, `Quantidade`,
     `Estoque Mínimo`, `Custo Médio`, `Data da Última Entrada`,
     `Data da Última Saída`.
   - A coluna precisa existir. A célula pode ficar vazia onde a regra já
     permite (ex.: `ID Oportunidade`, datas do estoque).
4. **Validação completa antes do resumo.** O arquivo passa pelas mesmas
   validações da carga (§0.2, §1, §2.6, §5). Qualquer erro recusa o arquivo
   **inteiro**, com a lista de erros. Não existe importação parcial. `ID Venda`
   repetido dentro do mesmo arquivo é erro.
5. **Resumo e confirmação.** Nada é gravado antes de a pessoa ver o resumo e
   confirmar. A confirmação refaz a validação; se a base mudou nesse meio
   tempo, vale o novo resultado.
   - Vendas, comparando pelo `ID Venda` com as vendas em vigor:
     - **novas:** ID que ainda não existe;
     - **atualizadas:** ID que já existe com algum valor diferente; a linha
       nova substitui a anterior (§1);
     - **ignoradas:** ID que já existe com todos os valores iguais (dinheiro a
       2 casas, datas como data).
   - Estoque: a data da foto em vigor e a da nova, as linhas da foto que serão
     **substituídas**, as linhas da nova, as linhas com quantidade diferente e
     as linhas que deixam de existir (passam a contar como zero, §5).
6. **Data da foto de estoque:** vem do nome do arquivo,
   `estoque_AAAA-MM-DD.xlsx` (§5). Nome fora desse padrão é recusado.
   - Foto **mais antiga** que a em vigor: recusada.
   - Foto **da mesma data**: substitui a posição inteira daquela data.
   - Foto **mais nova**: passa a ser a foto em vigor.
7. **Gravação: nada é apagado (§0.6).**
   - O arquivo confirmado é guardado sem alteração em
     `dados/importacoes/vendas/` ou `dados/importacoes/estoque/`, com a data e a
     hora da importação no nome.
   - As planilhas originais de `dados/` nunca são alteradas.
   - Cada importação confirmada ganha uma linha em
     `dados/importacoes/registro.csv`: quando, quem, tipo, arquivo e contagens.
   - Desfazer é uma ação manual: retirar o arquivo da pasta de importações.
8. **Aplicação.**
   - **Vendas em vigor:** a planilha base, depois as importações em ordem
     cronológica, com upsert por `ID Venda`. A versão mais recente de cada ID é
     a que vale.
   - **Estoque em vigor:** a foto de data mais recente. Se houver empate de
     data, vale a importação mais recente sobre a planilha base.
   - A data-base de vendas passa a ser a data da venda mais recente em vigor
     (§0.3).
9. **Conferências:** os testes de conferência (§6) rodam sobre a base, sem as
   importações. Os números da §6 continuam sendo os de jan–ago e da foto de
   31/08.

---

## 11. Feedback semanal ao vendedor

Pedido do coordenador em 23/09/2026: uma mensagem de WhatsApp do gerente para
cada vendedor, com meta, atingimento, ritmo, projeção e pipeline. Os números vêm
de `app/feedback_semanal.ps1`; a skill `feedback-semanal` só redige o texto.
Meta, realizado, atingimento e gap seguem a §4; o pipeline segue a §7 e a
conversão segue a §8. Esta seção define apenas os indicadores novos.

1. **Quem recebe:** só vendedores **ativos** (§2.3). Oportunidades órfãs (§2.4)
   não entram no feedback de ninguém.
2. **Semana:** os 7 dias corridos que terminam na data-base de vendas,
   incluindo a data-base. A semana anterior são os 7 dias antes dela. Só vendas
   faturadas (§1).
3. **Ritmo do mês:** se o mês da data-base está em andamento, o atingimento do
   mês (meta cheia, §4.5) é comparado com a fração de dias corridos do mês
   (dia da data-base ÷ dias do mês). "No ritmo" se o atingimento for maior ou
   igual a essa fração. Não é pró-rata da meta: a meta continua cheia.
4. **Ano e projeção** (só meses com meta, como na §4.1):
   - meses fechados = meses com meta de janeiro até o mês anterior ao da
     data-base (ou até o mês da data-base, se ele já terminou);
   - média mensal = realizado nos meses fechados ÷ número de meses fechados;
   - projeção de fim de ano "se mantiver a média" = média mensal × número de
     meses com meta no ano;
   - necessário por mês = (meta do ano − realizado nos meses fechados) ÷ número
     de meses com meta ainda não fechados (o mês em andamento conta inteiro);
   - falta para a meta do ano = meta do ano − realizado no ano até a data-base.
   - Vendedor admitido no ano (§3) usa só os meses dele. A projeção compara a
     média do ramp-up com as metas cheias; a mensagem deve dizer isso.
5. **Pipeline:** a foto do CRM (§7), com a data-base do CRM. Se ela for
   anterior à data-base de vendas, a mensagem avisa que algumas oportunidades
   podem já ter fechado e pede a atualização do CRM. Não há baixa automática
   (§7.6).
   - Cobertura = ponderado ÷ falta para a meta do ano.
   - Cenário ponderado = realizado no ano + ponderado, comparado com a meta do
     ano. É o valor esperado do pipeline, não uma promessa.
6. **Sinal**, só para o tom da mensagem, pela projeção ÷ meta do ano: "no
   ritmo" a partir de 100%, "um pouco abaixo" de 90% a 100%, "abaixo" abaixo
   de 90%, "sem histórico" sem nenhum mês fechado com meta.
7. **Envio:** a skill só redige e salva as mensagens em
   `relatorios/feedback-semanal/` (fora do Git). Enviar é sempre uma ação do
   gerente, ou um pedido explícito dele, mensagem por mensagem.
