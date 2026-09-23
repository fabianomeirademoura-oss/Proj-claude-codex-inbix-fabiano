# Perfil dos dados — Horizonte Máquinas Agrícolas

Levantamento feito em 22/09/2026 sobre as planilhas hoje em `dados/`, na raiz do
projeto. Todos os números abaixo foram contados diretamente nos arquivos, não
copiados do README.

> **Este arquivo descreve os dados; ele não cria regras.** As regras aprovadas
> estão só em [REGRAS_NEGOCIO.md](REGRAS_NEGOCIO.md), que vence em caso de
> conflito. Nos casos de borda (seção 3), cada caso indica a regra aprovada que
> o resolve. As implicações listadas são o levantamento feito na época: o que
> estiver marcado como *sugestão exploratória* **não está aprovado** e não deve
> ser implementado sem decisão do coordenador.

- **Data-base das bases principais:** 31/08/2026.
- **Pasta `dados/atualizacoes/`:** posição de 19/09/2026. O painel não lê esta pasta diretamente: os arquivos entram pela tela Importar (REGRAS_NEGOCIO.md §9).
- **Fora da análise:** `gabarito_conferencia.xlsx` (é do instrutor) e `gerar_dados.py` (é o script que gerou os dados). Os dois ficam em `../material-professor/`, fora do projeto.

---

## 1. O que cada arquivo contém

A contagem é de linhas de dados, sem o cabeçalho.

| Arquivo | Aba | Linhas | O que é |
|---|---|---|---|
| `vendedores.xlsx` | Vendedores | 12 | Cadastro da equipe comercial (V001–V012): nome, e-mail, filial, cargo, admissão, status (Ativo/Inativo) e data de desligamento. São 11 ativos e 1 inativo. |
| `vendedores.xlsx` | Filiais | 3 | As 3 filiais (F01 Cascavel/PR, F02 Chapecó/SC, F03 Passo Fundo/RS) e o gerente comercial de cada uma. |
| `metas_2026.xlsx` | Metas | 130 | Meta mensal em R$ por vendedor, com colunas separadas de Ano e Mês. Cobre **jan–dez/2026**: 10 vendedores × 12 meses, mais V011 com 4 meses (jan–abr) e V012 com 6 meses (jul–dez). |
| `vendas_2026_jan-ago.xlsx` | Vendas | 735 | Uma linha por item vendido, de 02/01 a 31/08/2026. São 723 Faturadas e 12 Canceladas. Máquinas e peças ficam na mesma base. 95 vendas têm `ID Oportunidade` (vieram do CRM); o resto é venda de balcão. |
| `produtos.xlsx` | Produtos | 65 | Catálogo (P001–P065) com 7 categorias. 22 itens são peças e acessórios; os outros 43 são máquinas e implementos. Tem preço de tabela, custo unitário, unidade (UN/KIT/JG/BD) e status: 63 Ativos e 2 Descontinuados (P002, P020). |
| `estoque_2026-08-31.xlsx` | Estoque | 195 | Foto do estoque em 31/08: 65 produtos × 3 filiais, todas as combinações, inclusive as zeradas. Traz quantidade, estoque mínimo, custo médio, data da última entrada e data da última saída. |
| `clientes.xlsx` | Clientes | 70 | Clientes C001–C070: 52 Produtores Rurais, 8 Cooperativas, 8 Empresas Agrícolas e 2 Prefeituras. Tem cidade/UF, filial de atendimento, cultura, área (ha) e data de início ("cliente desde"). |
| `crm_oportunidades.xlsx` | Oportunidades | 259 | Oportunidades OP-0001–OP-0259, criadas de 02/11/2025 a 27/08/2026. Por etapa: 95 Fechadas Ganhas, 96 Fechadas Perdidas e 68 Abertas. Só máquinas; nenhuma peça passa pelo CRM. |

Arquivos da pasta `atualizacoes/`:

| Arquivo | Aba | Linhas | O que é |
|---|---|---|---|
| `vendas_2026_setembro.xlsx` | Vendas | 82 | Vendas de 01/09 a 19/09 (80 Faturadas e 2 Canceladas). Os IDs vão de VD-2026-00736 a 00817, continuando a sequência sem sobreposição. **Nenhuma** tem ID Oportunidade. |
| `estoque_2026-09-19.xlsx` | Estoque | 195 | Nova foto completa do estoque. Mudaram 61 quantidades, 29 datas de entrada e 60 datas de saída. O custo médio não mudou. |

Observações gerais:

- Não há espaços sobrando em nenhum texto.
- Nenhuma chave está duplicada.
- Nenhuma chave estrangeira aponta para registro inexistente.
- Os nomes repetidos nas bases (vendedor, cliente, produto, categoria, cidade/UF) batem 100% com os cadastros.

---

## 2. Como os arquivos se ligam

```
                         Filiais (ID Filial F01..F03, Filial = nome)
                               ▲ (por NOME da filial, não por ID)
          ┌────────────────────┼─────────────────────────┬──────────────────┐
          │                    │                         │                  │
     Vendedores           Clientes                   Estoque            Vendas.Filial
   (ID Vendedor)        (ID Cliente)            (ID Produto + Filial)   (= filial do vendedor)
          ▲                    ▲                         ▲
          │                    │                         │
   ┌──────┼─────────┐   ┌──────┴──────┐                  │
   │      │         │   │             │                  │
 Metas  Vendas   Oportunidades ◄──────┘             Produtos (ID Produto)
          │        (ID Oportunidade)                 ▲    ▲    ▲
          │               ▲                          │    │    │
          └───────────────┘ Vendas.ID Oportunidade   │    │    │
                            (só vendas vindas do CRM)│    │    │
          Vendas.ID Produto ─────────────────────────┘    │    │
          Oportunidades.ID Produto ───────────────────────┘    │
          Estoque.ID Produto ──────────────────────────────────┘
```

| Chave | Onde é chave primária | Onde é referenciada | Observação |
|---|---|---|---|
| `ID Vendedor` (V001…) | Vendedores | Metas, Vendas, Oportunidades | A oportunidade guarda só o dono **atual** (ver caso 4). |
| `ID Cliente` (C001…) | Clientes | Vendas, Oportunidades | |
| `ID Produto` (P001…) | Produtos | Vendas, Oportunidades, Estoque | |
| `ID Oportunidade` (OP-NNNN) | Oportunidades | Vendas (opcional) | Relação 1:1. As 95 oportunidades Fechadas Ganhas têm exatamente 1 venda cada, com o mesmo vendedor, cliente, produto e quantidade. A data da venda é igual à data de fechamento, e o valor estimado é ≥ valor faturado. |
| `ID Venda` (VD-2026-NNNNN) | Vendas | — | Sequencial em ordem de data. Setembro continua a numeração. |
| `(ID Vendedor, Ano, Mês)` | Metas | — | Chave composta. O mês é um número inteiro, não uma data. |
| `(ID Produto, Filial)` | Estoque | — | Chave composta. A **Filial é texto** ("Chapecó", com acento). |
| `Filial` (nome) | Filiais | Vendedores, Vendas, Clientes (`Filial de Atendimento`), Estoque | **O `ID Filial` só existe na aba Filiais.** Todas as outras bases ligam pelo nome. |

Relações que valem a pena registrar:

- **`Vendas.Filial` é a filial do vendedor**, não a do cliente. A venda dá baixa no estoque dessa filial: a "Data da Última Saída" do estoque bate 100% com a última venda faturada por (produto, filial do vendedor).
- **`Clientes.Filial de Atendimento`** segue a UF do cliente sem exceção: PR→Cascavel, SC→Chapecó, RS→Passo Fundo.
- **Estoque ↔ Vendas:** a última saída considera só vendas **Faturadas**. P016/Cascavel tem apenas uma venda cancelada, e a saída está em branco.

---

## 3. Dez casos de borda

Cada caso traz a evidência concreta nos dados, as implicações levantadas e a **regra aprovada** que o resolve, em `REGRAS_NEGOCIO.md`.

### 1. Vendas canceladas misturadas às faturadas
- **Evidência:** 12 canceladas em jan–ago e mais 2 em setembro. Nenhuma tem ID Oportunidade. O estoque também ignora as canceladas.
- **Implicações levantadas:** filtrar `Status = 'Faturada'` no realizado, ranking, ticket médio, giro e "última venda". Uma soma direta da coluna Valor Total fica inflada. Se no futuro uma venda puder mudar de Faturada para Cancelada, o sistema precisa refletir isso na reimportação.
- **Regra aprovada:** REGRAS_NEGOCIO.md §1 (venda realizada, canceladas fora, upsert por `ID Venda`).

### 2. Metas com meses parciais, futuros e incompletos
- **Evidência:** V011 tem meta só de jan–abr (desligado em 30/04). V012 tem meta só de jul–dez (admitida em 01/07). Todas as metas vão até dezembro. As vendas de setembro cobrem só até dia 19.
- **Implicações levantadas:**
  - O atingimento do período deve somar apenas meses que **têm meta** e **já passaram**. Dividir o realizado de jan–ago pela meta anual derruba todo mundo.
  - Setembro parcial comparado com a meta cheia de setembro precisa de tratamento explícito (pró-rata ou aviso de "mês em andamento").
  - Vendedor sem nenhuma meta no período não pode dar divisão por zero.
- **Regra aprovada:** REGRAS_NEGOCIO.md §3.1, §4.1 e §4.5. O mês em andamento usa a meta cheia, com o aviso "parcial até dd/mm". **Não há pró-rata por padrão**: um pró-rata só pode aparecer se for pedido, como indicador separado e com o critério explícito. Sem meta no período, o resultado é "sem meta no período".

### 3. Vendedor desligado com carteira aberta
- **Evidência:** V011 (Inativo desde 30/04/2026) ainda é dono de 2 oportunidades abertas, OP-0130 (Qualificação) e OP-0140 (Proposta Enviada). Ele tem histórico de vendas e metas.
- **Implicações levantadas:**
  - Não pode ter login ativo.
  - Deve aparecer no histórico e nos relatórios de jan–abr, mas sair dos rankings de "ativos".
  - As oportunidades órfãs precisam de alerta ou de um fluxo de reatribuição. Não se pode apagar o vendedor, porque há FKs apontando para ele.
- **Regra aprovada:** REGRAS_NEGOCIO.md §2. As órfãs ficam no pipeline com alerta. **Nunca** são reatribuídas automaticamente: a reatribuição é uma ação explícita de um gestor, ainda sem tela.

### 4. Oportunidade criada antes de o dono ser admitido
- **Evidência:** OP-0150 foi criada em 21/04/2026 e pertence a V012, admitida só em 01/07/2026. Foi Fechada Ganha em 01/07, e a venda saiu no nome dela.
- **Leitura provável:** carteira transferida (talvez do V011). Como o CRM guarda só o dono atual, o histórico de quem criou a oportunidade se perdeu.
- **Implicações levantadas:** uma regra rígida de "criação ≥ admissão do vendedor" rejeitaria um dado legítimo.
- **Regra aprovada:** REGRAS_NEGOCIO.md §3.4. Não validar "criação ≥ admissão".
- **Sugestão exploratória (não aprovada):** se transferência de carteira virar requisito, guardar o histórico de dono.

### 5. A "filial" depende de qual filial se está olhando
- **Evidência:** 81 vendas (73 em jan–ago e 8 em setembro) e 31 oportunidades têm vendedor de uma filial e cliente atendido por outra. Exemplo: V001 de Cascavel vende para uma cooperativa de Chapecó/SC.
- **Implicações levantadas:**
  - "Faturamento por filial" dá números diferentes se agrupar por `Vendas.Filial` ou por `Clientes.Filial de Atendimento`. É preciso escolher um critério e documentar.
  - Se gerentes só puderem ver "sua filial", é preciso definir se isso vale pelo vendedor ou pelo cliente.
  - A junção é por **nome com acento** ("Chapecó"). Normalizar para `ID Filial` na importação evita problemas de encoding e de grafia.
- **Regra aprovada:** REGRAS_NEGOCIO.md §0.2 (normalizar para `ID Filial`), §1 e §4.2 (realizado e atingimento pela filial do vendedor) e §8.4 (filtro de filial do funil pela filial do vendedor).
- **Ainda sem regra:** a visão de gerente restrita à "sua filial". Hoje todas as contas veem o painel completo.

### 6. O estoque é uma foto e não fecha com as vendas
- **Evidência:**
  - 26 vendas de setembro saíram de filiais que não tinham estoque suficiente em 31/08.
  - P001/Passo Fundo tinha 0 unidades em 31/08, vendeu 2 em setembro e continua com 0 em 19/09, sem nenhuma entrada nova.
  - Em 48 das 195 linhas, "estoque 31/08 − vendas de setembro" é diferente do "estoque 19/09".
- **Implicações levantadas:**
  - Não calcular estoque derivando das vendas.
  - Não bloquear venda por falta de saldo, porque máquinas podem ser faturadas direto da fábrica ou sob encomenda.
  - Tratar cada arquivo de estoque como a verdade naquela data, guardando a data da posição.
- **Regra aprovada:** REGRAS_NEGOCIO.md §5 (importação substitutiva; nunca calcular estoque a partir das vendas).

### 7. Duas importações com semânticas opostas
- **Vendas de setembro são incrementais** (anexar linhas):
  - Reimportar o mesmo arquivo não pode duplicar. Use `ID Venda` como chave de *upsert*.
  - Um ID que já existe com dados diferentes é uma alteração, não uma linha nova.
- **Estoque de 19/09 é substitutivo** (troca a foto inteira):
  - Linhas que sumirem do arquivo precisam ser zeradas ou apagadas, não mantidas.
  - A lista de "produtos sem estoque" muda com a troca: P003 sai da lista, e P030, P034 e P042 entram.
- **Validação comum:** conferir data e formato antes de aplicar qualquer uma das duas. Importar um estoque antigo por cima de um novo precisa ser impedido.
- **Regra aprovada:** REGRAS_NEGOCIO.md §1 (vendas incrementais, upsert por `ID Venda`) e §5 (estoque substitutivo; linha que falta conta como zero; recusar foto mais antiga).
- **Estado:** implementadas na tela Importar, com resumo e confirmação antes de gravar (REGRAS_NEGOCIO.md §9).

### 8. O CRM está defasado em relação às vendas
- **Evidência:**
  - O CRM é a foto de 31/08, mas as vendas vão até 19/09.
  - Quatro vendas de setembro **sem** ID Oportunidade têm o mesmo cliente e produto de uma oportunidade aberta:

    | Venda | Oportunidade aberta |
    |---|---|
    | VD-2026-00747 | OP-0243 |
    | VD-2026-00791 | OP-0165 |
    | VD-2026-00806 | OP-0157 |
    | VD-2026-00807 | OP-0110 e OP-0222 |

  - Há oportunidades abertas duplicadas para o mesmo cliente e produto: C027/P004 e C043/P030.
  - 29 das 68 abertas estão com previsão de fechamento anterior a 31/08 (atrasadas).
- **Implicações levantadas:** o pipeline pode estar contando valor que já virou venda, e a taxa de conversão fica subestimada.
- **Regra aprovada:** REGRAS_NEGOCIO.md §7. O atraso é medido contra a data-base do CRM, nunca contra "hoje". **Não há deduplicação automática nem baixa ou fechamento automático** de oportunidades que parecem ter virado venda (§7.6).
- **Sugestão exploratória (não aprovada):** sinalizar oportunidades "provavelmente ganhas" e apontar possíveis duplicadas para revisão humana. Qualquer uma das duas exige uma regra nova.

### 9. Produtos sem estoque, descontinuados, parados e nunca vendidos
- **Evidência:**
  - 8 produtos têm estoque zero nas 3 filiais em 31/08.
  - 7 oportunidades abertas são de produtos sem estoque: P003 ×3, P018 ×3 e P010.
  - Os descontinuados P002 e P020 ainda têm 11 unidades em estoque e estão parados. O P002 teve uma oportunidade em 2026, a OP-0087, que foi perdida.
  - 8 produtos nunca foram vendidos em 2026.
- **Implicações levantadas:**
  - Indicadores como giro, cobertura e dias de estoque precisam tratar divisão por zero e "sem venda". *Sugestão exploratória:* giro, cobertura e dias de estoque ainda não têm regra aprovada.
  - "Parado ≥ 180 dias" só faz sentido com quantidade > 0: são 7 linhas nesse critério, contra 30 se incluir as zeradas.
  - Um produto descontinuado ainda precisa aparecer no estoque, mesmo sem aparecer no catálogo de venda.
- **Regra aprovada:** REGRAS_NEGOCIO.md §5 (parado por linha, quantidade > 0, prazo padrão de 180 dias configurável na tela; sem estoque; abaixo do mínimo; descontinuado pode estar parado).

### 10. Datas vazias, datas de referência e números "sujos"
- **Datas vazias têm significado:**
  - "Data da Última Saída" em branco (45 linhas) quer dizer *sem venda faturada em 2026*, não "nunca vendeu".
  - "Data da Última Entrada" em branco (29 linhas, todas com quantidade 0) quer dizer *sem registro de entrada*.
  - Área e Propriedade vazias em cooperativas e prefeituras são esperadas.
- **Regra aprovada:** REGRAS_NEGOCIO.md §0.3 (data-base), §0.4 (dinheiro decimal com 2 casas), §0.5 (percentuais) e §5.2 (datas em branco no estoque).
- **Datas de referência:** idade de estoque e atraso do CRM devem ser calculados contra a **data-base de cada arquivo** (31/08 ou 19/09), não contra a data do sistema. Hoje é 22/09.
- **Formatos nos arquivos:**
  - As datas estão como serial do Excel.
  - A probabilidade vem como fração de 0 a 1 (0,75), não como 75.
  - Existem vendas em sábados (39) e no próprio dia da data-base (4).
- **Ruído de ponto flutuante:**
  - 12 valores totais vêm com lixo decimal. Exemplo: VD-2026-00047 = `8891.280000000001`.
  - O custo do P047 é `819.3200000000001`.
  - Valores monetários devem ser guardados como decimal com 2 casas, e as somas arredondadas no fim.
