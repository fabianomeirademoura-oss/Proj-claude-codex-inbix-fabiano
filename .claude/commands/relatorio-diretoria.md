---
description: Gera o relatório semanal da diretoria em PDF (itens 1 a 5 calculados, item 6 com três riscos analisados pelo Claude Code)
argument-hint: "[opções do script, ex.: -DirDados <pasta de dados>]"
---

Gere o relatório semanal da diretoria da Horizonte Máquinas. O formato é fixo e está em
`docs/RELATORIO_DIRETORIA.md`. Os itens 1 a 5 são **cálculo** do script e não são seus: não os
recalcule, não os reescreva e não os comente no PDF. O item 6 (três riscos) é **análise sua**.

Opções extras para o script (repasse nos dois passos, sem mudar): `$ARGUMENTS`

## 1. Calcular os itens 1 a 5

```
pwsh -NoProfile -File app/relatorio_diretoria.ps1 -SoCalculo $ARGUMENTS
```

No Windows: `app\relatorio_diretoria.cmd -SoCalculo`. O script usa os mesmos dados do painel: a
base de `dados/` e as importações confirmadas (REGRAS_NEGOCIO.md §9.8). Ele nunca grava em `dados/`.

- Se ele parar com erros de importação, **pare aqui** e mostre os erros. Não corrija planilhas.
- Anote a pasta de saída (`relatorios/<data-base>/`) e a impressão digital dos dados que ele imprime.

## 2. Analisar os dados desta semana e escolher três riscos

Leia `relatorios/<data-base>/calculo.json`: os itens 1 a 5 e o catálogo `indicadores` (chave →
tipo, valor, descrição). Se houver em `relatorios/` a pasta de uma semana anterior, leia o
`calculo.json` dela também: o que mudou de uma semana para a outra costuma ser o melhor sinal.

Escolha **três riscos** e siga estes critérios:

- **Diferentes entre si:** por exemplo, um de vendas e metas, um de pipeline e CRM, e um de estoque
  ou de qualidade dos dados. Não repita o que uma tabela já mostra. Diga o que pode dar errado nas
  próximas semanas se nada for feito, e quanto está em jogo em reais.
- **Um número sustenta cada risco.** Prefira a chave de um indicador do catálogo (`indicador`). O
  PDF imprime o valor do cálculo, e não um número digitado. Se o número que você precisa não está
  no catálogo, calcule-o com um script sobre as funções de `app/lib/Regras*.ps1`, sem estimar e sem
  fazer conta de cabeça sobre valores arredondados. Depois preencha `numero` e `calculo`: como
  reproduzir a conta.
- **Todo número citado na leitura** também tem que ser conferido do mesmo jeito.
- **Respeite o REGRAS_NEGOCIO.md**, citando a seção quando ajudar:
  - não apresente pró-rata como atingimento (§4.5); se usar ritmo por dias, diga que é pró-rata e qual foi o critério;
  - não invente taxa de passagem entre etapas (§8.1);
  - parado não é o mesmo que sem demanda (§5.4);
  - órfãs não são reatribuídas automaticamente (§2.4);
  - desligados ficam fora dos rankings (§2.3);
  - cada fonte tem a sua data-base (§0.3).

  Não proponha regra nova. Se um risco depender de um critério que o contrato não cobre, diga isso.
- **Escreva para a diretoria**, em português claro e sem jargão técnico. O título tem até 90
  caracteres e a leitura, até 700.

## 3. Gravar o item 6

Grave `relatorios/<data-base>/riscos.json` (UTF-8), com a impressão digital do passo 1:

```json
{
  "impressao_digital": "<a do passo 1>",
  "riscos": [
    { "titulo": "…", "indicador": "<chave do catálogo>", "leitura": "…" },
    { "titulo": "…", "indicador": "<chave do catálogo>", "leitura": "…" },
    { "titulo": "…", "numero": "R$ …", "calculo": "como reproduzir a conta", "leitura": "…" }
  ]
}
```

## 4. Gerar o PDF

```
pwsh -NoProfile -File app/relatorio_diretoria.ps1 $ARGUMENTS
```

O script recalcula os itens 1 a 5 e valida o `riscos.json`: exatamente 3 riscos, cada um com o seu
número, e a impressão digital dos dados atuais. Se ele recusar o arquivo (código 2), corrija o
arquivo e rode de novo. Só use `-SemAnalise` se a pessoa pedir.

## 5. Entregar

Responda com o caminho do PDF, as datas-base, a impressão digital e os três riscos, cada um numa
linha com o título e o número. Não faça commit de `relatorios/`: a pasta fica fora do Git.
