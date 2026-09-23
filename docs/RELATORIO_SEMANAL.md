# Relatório semanal da diretoria

## Um comando no agente

No Claude Code, dentro deste projeto:

```text
/relatorio-semanal
```

No Codex, com a skill do projeto carregada:

```text
$relatorio-semanal
```

Também é possível pedir: “Gere o relatório semanal seguindo
`.agents/skills/relatorio-semanal/SKILL.md`”. A skill é local ao projeto e
pode precisar de uma nova sessão para aparecer. O agente executa os cálculos,
lê os resultados, redige os três riscos e entrega o Markdown final.
Não é um agendamento e não exige uma API adicional.

## Formato e recortes

O relatório tem sempre seis seções, na ordem pedida. REGRAS_NEGOCIO.md §10:

1. Ranking dos vendedores ativos por faturamento do mês, meta cheia, realizado,
   realizado nos meses com meta e atingimento. Empate por ID crescente.
2. Ativos abaixo de 80%, sem arredondar antes de comparar. Falta em reais é
   para chegar a 100% da meta. Sem meta não vira zero por cento.
3. Pipeline aberto por etapa na foto do CRM: valor bruto e ponderado.
4. Abertas com previsão anterior à data-base do CRM, não à data do computador.
5. Até dez produtos com maior quantidade faturada nos sete dias por filial,
   desempate por ID; unidade do catálogo explícita (sem converter KIT em UN).
   Estoque parado detalhado e somado por filial, com prazo de 180 dias.
6. Três riscos selecionados e redigidos pelo agente, com evidências atuais e
   ação sugerida. Não há seleção automática de três frases predefinidas.

A janela de sete dias inclui a data da última venda em vigor. Metas e realizado
usam o mês dessa data, sem pró-rata. As datas de CRM e estoque permanecem
independentes. Sem importações, o exemplo é 25–31/08/2026, com metas de agosto;
não são dados atuais de setembro. As importações confirmadas na pasta da base
escolhida são aplicadas como no painel; `dados/atualizacoes/` não é carregada.

## Motor de cálculo no terminal

Da raiz do projeto, escolhendo explicitamente uma pasta de saída:

```powershell
pwsh -NoProfile -File app/relatorio_semanal.ps1 -PastaSaida /tmp/relatorio-diretoria
```

No Windows, use uma pasta de saída válida, por exemplo `C:\Temp\relatorio-diretoria`.
Pode-se passar `-DirDados <pasta>`; sem ele, vale `HORIZONTE_DADOS` ou `dados/`.
Não há acesso ao servidor nem necessidade de login. Quem executa já precisa
ter acesso local às planilhas. Nenhuma credencial é lida.

Essa etapa cria `.calculos.md` (itens 1–5) e `.evidencias.json`, com assinatura
SHA-256 da base calculada. **Ela não é o relatório completo.** Os cálculos,
ordenação e formatação não usam o relógio, valores aleatórios ou uma IA.
Um mês sem faturamento produz tabelas vazias explícitas; sem data-base de
vendas, o comando recusa definir uma semana arbitrária.

O agente lê os dois arquivos e grava, em UTF-8, um JSON deste formato:

```json
{
  "baseSha256": "copiar a assinatura exata do pacote desta geração",
  "autor": "Codex",
  "riscos": [
    {"titulo": "Risco identificado", "analise": "Interpretação fundamentada", "acao": "Ação sugerida", "evidencias": ["crm-atrasadas"]},
    {"titulo": "Segundo risco", "analise": "Interpretação fundamentada", "acao": "Ação sugerida", "evidencias": ["meta-abaixo80"]},
    {"titulo": "Terceiro risco", "analise": "Interpretação fundamentada", "acao": "Ação sugerida", "evidencias": ["estoque-parado"]}
  ]
}
```

Os riscos do exemplo são apenas campos do esquema, não conclusões prontas.
As evidências disponíveis vêm do JSON gerado. Finalize com:

```powershell
pwsh -NoProfile -File app/relatorio_semanal.ps1 -PastaSaida /tmp/relatorio-diretoria -AnaliseJson /tmp/relatorio-diretoria/analise.json
```

A finalização recalcula os dados e recusa assinatura antiga, menos ou mais de
três riscos, campos vazios ou evidências inexistentes. Os números de suporte
são inseridos pelo código. O texto analítico continua sujeito à revisão humana;
o validador não prova que uma interpretação econômica seja correta.

O arquivo final é `relatorio-semanal-AAAA-MM-DD.md`. Uma execução apenas de
cálculo não atualiza um relatório final anterior na mesma pasta: confira a
mensagem do comando ou use pasta nova para cada execução do agente.
Preserve o Markdown final e o pacote se precisar guardar histórico. Dados
iguais produzem cálculos iguais; a análise pode mudar. Não versione relatórios
operacionais, salvo pedido do coordenador.

## Verificação

```powershell
pwsh -NoProfile -File testes/conferir_relatorio.ps1
```

Execute também as seis verificações do projeto antes de integrar alterações.
O comando não muda as funções compartilhadas, as telas ou os servidores.
