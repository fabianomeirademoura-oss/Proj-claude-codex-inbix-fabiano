---
name: relatorio-semanal
description: Gere o relatório semanal da diretoria da Horizonte em Markdown, com cinco seções calculadas pelo PowerShell e três riscos analisados pelo agente a partir das evidências atuais.
---

# Relatório semanal

Execute a partir da raiz do projeto, sem alterar código, planilhas, importações
ou credenciais. Leia `REGRAS_NEGOCIO.md §10` e `docs/RELATORIO_SEMANAL.md`.

1. Escolha uma pasta de saída autorizada fora das fontes (se não indicada,
   use uma pasta temporária nova). Execute:
   `pwsh -NoProfile -File app/relatorio_semanal.ps1 -PastaSaida <pasta>`.
   Use `-DirDados` só se o usuário indicar outra base. Na ausência disso, o
   comando lê `HORIZONTE_DADOS` ou `dados/`, com importações confirmadas.
2. Leia integralmente o `.calculos.md` e o `.evidencias.json` produzidos.
   Não reescreva, reordene nem recalcule as seções 1–5. A semana é a dos
   últimos sete dias disponíveis, não necessariamente a semana do calendário.
3. Analise os dados e escolha três riscos relevantes, distintos, com número
   que os sustente. Escreva `analise.json` conforme o esquema do manual:
   `baseSha256` exato, `autor`, três objetos em `riscos` com `titulo`,
   `analise`, `acao` e IDs de `evidencias`. O gerador insere os números e
   respectivas fontes automaticamente; não invente números no texto livre.
   Separe hipótese de fato. Não declare tendência sem comparação histórica,
   atraso de meta em mês parcial sem critério de ritmo, estoque parado como
   ausência de demanda, ou pipeline ponderado como receita garantida. Considere
   a defasagem das fotos. Se não houver três riscos sustentáveis, declare essa
   limitação e peça orientação; não fabrique riscos para preencher o formato.
4. Execute novamente o mesmo comando, mesma base e pasta, acrescentando
   `-AnaliseJson <caminho/analise.json>`. Ele recalcula os itens 1–5 e recusa a
   análise se a assinatura não corresponder. Se os dados mudaram, execute primeiro sem `-AnaliseJson` para atualizar
   os cálculos e as evidências. Leia essa nova geração e refaça a análise;
   não copie só a assinatura. Pare e explique se
   uma nova tentativa também encontrar mudanças concorrentes.
5. Leia o `.md` final, confira seis seções e exatamente três riscos com
   evidências, e entregue o arquivo. Não entregue `.calculos.md` como relatório
   completo e não marque como concluído quando a análise não foi finalizada.

Não publica, envia a terceiros, cria agendamento, faz commit ou abre PR. A
invocação autoriza somente gerar o relatório local na pasta escolhida.
