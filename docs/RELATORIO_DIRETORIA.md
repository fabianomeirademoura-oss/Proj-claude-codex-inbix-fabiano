# Relatório semanal da diretoria (PDF)

Um PDF por semana, sempre com as mesmas seis seções na mesma ordem. Os itens 1 a 5 são
**cálculo**: saem iguais, byte a byte, para os mesmos dados. O item 6 é **análise** do Claude Code
e aparece marcado assim no PDF. As regras de cálculo são as do [REGRAS_NEGOCIO.md](../REGRAS_NEGOCIO.md);
este documento só diz como o relatório as usa.

## Como gerar

No Claude Code, a partir da raiz do projeto:

```
/relatorio-diretoria
```

O comando ([.claude/commands/relatorio-diretoria.md](../.claude/commands/relatorio-diretoria.md))
faz os passos abaixo. Também dá para rodá-los à mão:

1. `pwsh -NoProfile -File app/relatorio_diretoria.ps1 -SoCalculo`: calcula os itens 1 a 5 e
   grava `relatorios/<data-base de vendas>/calculo.json`.
2. Alguém (o agente ou uma pessoa) escreve `relatorios/<data-base>/riscos.json` com os três riscos.
3. `pwsh -NoProfile -File app/relatorio_diretoria.ps1`: recalcula, valida os riscos e grava
   `relatorios/<data-base>/relatorio_diretoria_<data-base>.pdf`.

No Windows: `app\relatorio_diretoria.cmd`, com as mesmas opções.

| Opção | Efeito |
|---|---|
| `-SoCalculo` | Só o passo 1. |
| `-SemAnalise` | Gera o PDF sem o item 6; a seção diz "Análise não incluída nesta emissão". |
| `-DirDados <pasta>` | Outra pasta de dados (padrão: `dados/` ou `HORIZONTE_DADOS`). |
| `-SemImportacoes` | Só a base, sem as importações (é o que as conferências usam, §9.9). |
| `-Saida <pasta>`, `-Riscos <arquivo>` | Outros caminhos de saída e de riscos. |

**Dados usados:** os mesmos do painel. São a base de `dados/` mais as importações confirmadas
(§9.8), com a data-base de cada fonte (§0.3). Quando as fotos são de datas diferentes, o PDF avisa
na capa. O comando só lê `dados/` e nunca grava nada lá.

**Saída:** `relatorios/` fica fora do Git (`.gitignore`), porque traz dados do negócio e a análise
de cada semana.

## As seções

| # | Seção | Critério |
|---|---|---|
| 1 | Ranking de vendedores | Só ativos (§2.3), do maior para o menor atingimento no **acumulado do ano**. Ao lado, o **mês da data-base**, com a meta cheia e a marca "parcial até dd/mm" (§4.5). Meta, realizado e atingimento seguem a §4.1; a linha da empresa segue a §4.2, com os desligados. |
| 2 | Abaixo de 80% da meta | Ativos com atingimento acumulado **menor que 80%**. "Falta para a meta" é o gap da §4.3; "falta para 80%" é 80% da meta menos o realizado. |
| 3 | Pipeline por etapa | Foto do CRM (§7): por etapa, a quantidade, o valor bruto, o ponderado e a parte com previsão vencida. Traz um alerta para as órfãs (§2.4). |
| 4 | Previsão vencida | Oportunidades abertas com Previsão de Fechamento anterior à data-base do CRM (§7), da mais atrasada para a menos atrasada. |
| 5 | Maior saída e estoque parado | Os **dez produtos da empresa com maior faturamento** no acumulado do ano (§1), com o valor por filial da venda e o estoque atual. Depois, o estoque parado por filial (§5, prazo padrão), com as oportunidades abertas ao lado (§5.4). |
| 6 | Três riscos | Análise do agente. Cada risco traz o número que o sustenta. |

As escolhas de período (acumulado e mês lado a lado), do corte de 80% e de "maior saída = faturamento
em R$" foram decisões do coordenador em 23/09/2026 (PR #3). Elas ficam como constantes no topo de
`app/lib/RegrasRelatorioDiretoria.ps1`.

## Por que os itens 1 a 5 saem sempre iguais

- O cálculo usa as funções já conferidas de `app/lib/Regras*.ps1`. Não existe regra própria do relatório.
- Nada lê o relógio: as datas vêm das datas-base dos dados.
- Toda ordenação tem desempate pelo ID.
- Os números são formatados em pt-BR fixo, sem depender da cultura instalada na máquina.
- O PDF é escrito pelo próprio projeto (`app/lib/Pdf.ps1`, sem dependências). Ele não grava data de criação nem identificador aleatório.
- O rodapé de cada página traz a **impressão digital dos dados**: um SHA-256 do conteúdo das planilhas usadas. Com a mesma impressão digital, os itens 1 a 5 são os mesmos.

## O item 6 (`riscos.json`)

```json
{
  "impressao_digital": "<a impressa pelo passo 1>",
  "riscos": [
    { "titulo": "…", "indicador": "pipeline.atrasadas.pct_valor", "leitura": "…" },
    { "titulo": "…", "numero": "R$ 1,2 mi", "calculo": "como reproduzir a conta", "leitura": "…" },
    { "titulo": "…", "indicador": "estoque.sem_estoque_com_oportunidade.valor", "leitura": "…" }
  ]
}
```

O PDF só sai se o arquivo tiver:

- **Exatamente 3 riscos**, cada um com título (até 90 caracteres) e leitura (até 700).
- **Um número por risco**, dado de um destes jeitos:
  - `indicador`: a chave de um item do catálogo `indicadores` do `calculo.json`. O PDF imprime o valor do cálculo, e não o que foi digitado.
  - `numero` e `calculo`: o número e como reproduzi-lo. O PDF o marca como "número calculado na análise".
- **A impressão digital dos dados atuais.** Riscos escritos para outra semana são recusados.

## Conferência

`pwsh -NoProfile -File testes/conferir_relatorio_diretoria.ps1` (no Windows, o `.cmd` de mesmo nome).
O teste confere:

- os itens 1 a 5 contra somas diretas nas planilhas brutas e contra a §6;
- que duas execuções geram `calculo.json` e PDF idênticos;
- a estrutura do PDF;
- a recusa de arquivos de riscos inválidos;
- o cenário com setembro importado, numa cópia temporária de `dados/`.
