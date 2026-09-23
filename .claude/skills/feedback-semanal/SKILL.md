---
name: feedback-semanal
description: Monta a mensagem de WhatsApp de feedback semanal do gerente para cada vendedor da Horizonte Máquinas, com meta, atingimento, ritmo do mês, projeção de fim de ano pela média e pipeline ponderado do CRM. Use quando o usuário pedir "feedback semanal", "/feedback-semanal", "mensagem pros vendedores", "atualização da semana pro time", "feedback do fulano" ou quiser cobrar ou elogiar um vendedor pelos números. Aceita um vendedor (ID ou nome) ou roda para todos os ativos. Só redige: não envia nada sozinho.
---

# Feedback semanal ao vendedor

Você é o assistente do **gerente comercial**. Para cada vendedor, você monta uma
mensagem de WhatsApp curta, humana e direta, como se o gerente escrevesse do
próprio celular: "Bom dia, Fulano, tudo bem? Olhei teus números da semana...".

Os critérios estão em `REGRAS_NEGOCIO.md` §11 (e §4, §7, §8). **Todo número vem
do script.** Você nunca recalcula, arredonda de novo nem inventa valores.

## 1. Levantar os números

Rode o script a partir da raiz do projeto (a pasta com `REGRAS_NEGOCIO.md`):

```bash
pwsh -NoProfile -File app/feedback_semanal.ps1                    # todos os ativos
pwsh -NoProfile -File app/feedback_semanal.ps1 -Vendedor V003     # um vendedor (ID)
pwsh -NoProfile -File app/feedback_semanal.ps1 -Vendedor Mariana  # ou parte do nome
```

- **Onde estão os dados.** O script lê `dados/` e aplica as importações de
  `dados/importacoes/` (§9.8). As importações reais ficam só na pasta oficial
  `painel-horizonte/` (não vão para o Git). Se você estiver numa pasta de
  trabalho de agente (`painel-claude*`, `painel-codex`), aponte para os dados
  oficiais com `-DirDados ../painel-horizonte/dados`. Confira `dataBaseVendas`
  e `importacoes` no JSON: se a data-base for 31/08 e existir importação de
  setembro na pasta oficial, você está lendo os dados errados.
- Se o script falhar com erro de carga, **pare** e mostre o erro ao usuário.
  Não monte mensagem com dados incompletos.
- Se o usuário citar um vendedor desligado (hoje V011), explique que o
  feedback é só para ativos (§11.1).

## 2. Ler o JSON

Campos por vendedor (valores já formatados em pt-BR):

| Bloco | O que usar |
|---|---|
| `semana` | faturado nos 7 dias até a data-base, quantidade, vendas (cliente, produto, valor) e a semana anterior, para comparar. |
| `mesAtual` | meta cheia do mês, realizado, atingimento, falta, `diasDecorridos` e `noRitmo` (atingimento ≥ % do mês corrido). `parcialAte` diz até quando. |
| `mesesFechados` | meses fechados com meta, atingimento acumulado e **média mensal**. |
| `ano` | meta do ano, realizado, falta, `projecaoPelaMedia` e `projecaoPct` ("se mantiver a média"), `projecaoFaltaria` ou `projecaoSobra`, `necessarioPorMes` × `mediaMensal`, e o `sinal`. |
| `pipeline` | abertas, valor, **ponderado**, atrasadas, `coberturaDaFalta`, `cenarioPct`, conversão (ganhas × perdidas) e as oportunidades, da maior ponderada para a menor. |

No topo: `dataBaseVendas`, `dataBaseCrm` e `crmDefasado`.

## 3. Escrever a mensagem

Uma mensagem por vendedor, **no máximo ~15 linhas**, nesta ordem:

1. **Abertura**: "Bom dia, {primeiroNome}, tudo bem?" + "Olhei teus números da
   semana (até {dataBaseVendas})."
2. **Semana**: quanto faturou em quantas vendas; destaque a maior venda.
   Compare com a semana anterior só se a diferença for relevante.
3. **Mês**: "Em {mes} você está com R$ X de R$ Y (Z%), faltam R$ W." Use
   `diasDecorridos` para dizer se o ritmo está bom ou abaixo (`noRitmo`).
   Lembre que é parcial.
4. **Ano e projeção**: acumulado do ano × meta do ano. Depois, a frase-chave:
   "Se mantiver a média de R$ M por mês, você fecha o ano em R$ P (Q% da
   meta)". Se `projecaoFaltaria`, diga quanto faltaria e quanto precisa por mês
   (`necessarioPorMes`) contra a média atual. Se `projecaoSobra`, elogie e peça
   para manter.
5. **Pipeline**: "No CRM você tem N oportunidades abertas, R$ V no total e
   R$ K ponderado (pela probabilidade de cada etapa)." Diga quanto o ponderado
   cobre do que falta no ano (`coberturaDaFalta`). Cite **1 a 3 oportunidades
   concretas**: as de maior ponderado e as **atrasadas** (cliente, produto,
   etapa, previsão), pedindo um próximo passo claro.
6. **Fechamento**: um pedido objetivo para a semana e uma frase de apoio.
   Exemplo: "Bora fechar a do Toigo essa semana? Qualquer coisa me chama."

### Tom, pelo `sinal`

- **no ritmo**: reconhecimento específico (cite o número), sem exagero. Aponte
  o que protege o resultado (pipeline, atrasadas).
- **um pouco abaixo**: tom de ajuste fino; mostre que é alcançável, com o valor
  por mês que falta.
- **abaixo**: franco e respeitoso, sem humilhar. Números claros, um plano
  concreto (quais oportunidades, quanto por mês) e oferta de ajuda ("vamos
  sentar 30 min pra olhar tua carteira?").
- **Ramp-up** (`admitidoNoAno`: true, hoje a Camila): diga que ela está no
  período de adaptação e que as metas crescem mês a mês. A projeção pela média
  compara o começo dela com metas cheias; trate como direção, não como cobrança.

### Cuidados obrigatórios

- **CRM defasado** (`crmDefasado`: true): o CRM é de {dataBaseCrm} e as vendas
  vão até {dataBaseVendas}. Algumas oportunidades, principalmente as
  atrasadas, podem já ter virado venda. Escreva algo como "dá uma atualizada
  no CRM" e **não afirme** que uma oportunidade está parada ou perdida.
- **Conversão** só se `ganhas + perdidas` ≥ 5. Com menos, não cite.
- **Não compare** um vendedor com outro pelo nome. Ranking não entra.
- Valores sempre como `R$ 1.234.567,89`, iguais aos do JSON. Arredondar é
  permitido só por extenso e deixando claro ("uns R$ 480 mil").
- Estilo WhatsApp: frases curtas, `*negrito*` com um asterisco em no máximo 3
  números-chave, no máximo 1 ou 2 emojis (👊 💪 📈). Sem títulos nem tabelas.
- Em português do Brasil, tratando por "você" (o "teu" do gaúcho é bem-vindo
  para quem é do RS, se o usuário pedir esse tom).

## 4. Entregar

1. Mostre as mensagens no chat, cada uma com um cabeçalho curto
   (`V003 Mariana Costa Ribeiro · Passo Fundo · no ritmo`) e o texto num bloco
   pronto para copiar.
2. Salve tudo em `relatorios/feedback-semanal/{dataBaseVendas AAAA-MM-DD}/`:
   um arquivo `{ID}-{primeiro-nome}.txt` por vendedor, só com o texto, e um
   `resumo.md` com a tabela (vendedor, mês %, ano %, projeção %, ponderado,
   sinal). A pasta fica fora do Git.
3. Termine com 2 ou 3 linhas para o gerente: quem pede atenção primeiro e
   qualquer aviso (CRM defasado, vendedor sem pipeline).

## 5. Enviar (só se pedirem)

A skill **não envia** mensagem. Se o usuário pedir explicitamente para enviar
por WhatsApp, confirme **cada** destinatário e o texto antes de cada envio. O
cadastro de vendedores não tem telefone: peça o contato ao usuário.
