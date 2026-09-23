# Somente formatação. Não consulta relógio, rede ou credenciais.
$script:RelPt = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
function Rel-Texto($v) { return (([string]$v).Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('|','&#124;').Replace('[','&#91;').Replace(']','&#93;').Replace('*','&#42;').Replace('_','&#95;').Replace('`','&#96;') -replace '[\r\n]+',' ') }
function Rel-Moeda($v) { return 'R$ ' + ([decimal]$v).ToString('N2',$script:RelPt) }
function Rel-Pct($v) { if ($null -eq $v) { return 'sem meta no período' }; return ([decimal]$v * 100).ToString('N1',$script:RelPt) + '%' }
function Rel-Data($v) { return $v.ToString('dd/MM/yyyy',[Globalization.CultureInfo]::InvariantCulture) }
function Get-RelatorioCalculos($D) {
    $s = New-Object Collections.Generic.List[string]
    $s.Add('# Relatório semanal da diretoria')
    $s.Add('')
    $s.Add("Semana de vendas: **$(Rel-Data $D.InicioSemana) a $(Rel-Data $D.FimSemana)** (sete dias corridos, incluindo a data-base).")
    $s.Add("Metas e realizado: **$(Rel-Data $D.InicioMes) a $(Rel-Data $D.FimSemana)**; meta mensal cheia, sem pró-rata.")
    $s.Add("CRM: **$(Rel-Data $D.DataCrm)**. Estoque: **$(Rel-Data $D.DataEstoque)**. São fotos independentes; não representam necessariamente a mesma semana.")
    if ($D.Parcial) { $s.Add("**Mês parcial até $(Rel-Data $D.FimSemana):** estar abaixo de 80% não prova atraso no ritmo mensal; não foi aplicado pró-rata.") }
    $s.Add('Os itens 1 a 5 são cálculos reproduzíveis. O item 6 é análise do agente. Sem comparação histórica, não se afirmam aumentos ou quedas semanais.')
    $s.Add('')
    $s.Add('## 1. Ranking de vendedores')
    $s.Add('')
    $s.Add('Somente ativos; ordem por faturamento decrescente, desempate por ID. Atingimento usa somente meses com meta. Histórico de desligados continua no total da empresa (REGRAS_NEGOCIO.md §§2–4).')
    $s.Add('')
    $s.Add('| Posição | ID | Vendedor | Meta mensal | Realizado | Realizado com meta | Atingimento |')
    $s.Add('|---|---|---|---:|---:|---:|---:|')
    $n=0
    foreach ($l in $D.Ranking) {
        $n++; $meta=if ($l.MesesComMeta.Count) {Rel-Moeda $l.Meta} else {'sem meta no período'}
        $nome=Rel-Texto $l.Vendedor.Nome
        if ($l.Vendedor.Admissao.Year -eq $D.FimSemana.Year -and $l.Vendedor.Admissao.Month -eq $D.FimSemana.Month) { $nome += ' (admissão neste mês)' }
        $s.Add("| $n | $($l.Vendedor.Id) | $nome | $meta | $(Rel-Moeda $l.Realizado) | $(Rel-Moeda $l.RealizadoMesesMeta) | $(Rel-Pct $l.Atingimento) |")
    }
    if (-not $n) { $s.Add('| — | — | Nenhum vendedor ativo | — | — | — | — |') }
    $s.Add('')
    $s.Add("Total da empresa, incluindo histórico dos desligados: meta $(Rel-Moeda $D.Empresa.Meta); realizado $(Rel-Moeda $D.Empresa.Realizado); atingimento $(Rel-Pct $D.Empresa.Atingimento).")
    $s.Add('')
    $s.Add('## 2. Abaixo de 80% da meta')
    $s.Add('')
    $s.Add('Atingimento estritamente menor que 80%, antes do arredondamento; sem meta fica fora. Falta em reais = meta − realizado nos meses com meta, para chegar a 100% (REGRAS_NEGOCIO.md §4.3). Ordem pela maior falta, desempate por ID.')
    $s.Add('')
    $s.Add('| ID | Vendedor | Atingimento | Falta para 100% da meta |')
    $s.Add('|---|---|---:|---:|')
    foreach ($l in $D.Abaixo) { $s.Add("| $($l.Vendedor.Id) | $(Rel-Texto $l.Vendedor.Nome) | $(Rel-Pct $l.Atingimento) | $(Rel-Moeda ($l.Meta-$l.RealizadoMesesMeta)) |") }
    if (-not $D.Abaixo.Count) { $s.Add('| — | Nenhum vendedor ativo abaixo do limite | — | — |') }
    $s.Add('')
    $s.Add('## 3. Pipeline aberto por etapa')
    $s.Add('')
    $s.Add('Foto completa do CRM, sem filtro da semana. Bruto = valor estimado; ponderado = valor × probabilidade da linha, arredondado a centavos antes da soma. Fechadas ficam fora (REGRAS_NEGOCIO.md §7).')
    $s.Add('')
    $s.Add('| Etapa | Oportunidades | Valor bruto | Valor ponderado |')
    $s.Add('|---|---:|---:|---:|')
    foreach ($e in $D.Funil.Etapas) { $s.Add("| $(Rel-Texto $e.Etapa) | $($e.Qtd) | $(Rel-Moeda $e.Valor) | $(Rel-Moeda $e.Ponderado) |") }
    $s.Add("| **Total aberto** | $($D.Funil.Total.Qtd) | $(Rel-Moeda $D.Funil.Total.Valor) | $(Rel-Moeda $D.Funil.Total.Ponderado) |")
    $s.Add('')
    $s.Add('## 4. Oportunidades abertas com previsão vencida')
    $s.Add('')
    $s.Add("Previsão anterior a $(Rel-Data $D.DataCrm), data-base do CRM. Ordem pela previsão mais antiga, desempate por ID (REGRAS_NEGOCIO.md §7).")
    $s.Add('')
    $s.Add('| ID | Vendedor | Cliente | Etapa | Previsão | Dias vencidos | Valor bruto |')
    $s.Add('|---|---|---|---|---|---:|---:|')
    foreach ($a in $D.Atrasadas) {
        $nome=Rel-Texto $a.Vendedor.Nome; if ($a.Orfa) { $nome+=' (órfã: vendedor inativo)' }
        $s.Add("| $($a.Oportunidade.Id) | $nome | $(Rel-Texto $a.Cliente.Nome) | $(Rel-Texto $a.Oportunidade.Etapa) | $(Rel-Data $a.Oportunidade.Previsao) | $($a.DiasAtraso) | $(Rel-Moeda $a.Oportunidade.Valor) |")
    }
    if (-not $D.Atrasadas.Count) { $s.Add('| — | Nenhuma oportunidade vencida | — | — | — | — | — |') }
    $s.Add('')
    $s.Add('## 5. Produtos com maior saída e estoque parado por filial')
    $s.Add('')
    $s.Add('Saída = soma da quantidade faturada nos sete dias, por produto e filial da venda. Até dez produtos por filial, quantidade decrescente e ID crescente no empate; canceladas fora. Quantidades mantêm a unidade do catálogo (UN/KIT/JG/BD), sem conversão entre embalagens. Estoque é a foto independente, nunca saldo calculado pelas vendas.')
    $s.Add('')
    foreach ($f in $D.Saidas) {
        $s.Add("### $(Rel-Texto $f.Filial.Nome) — maiores saídas")
        $s.Add('')
        $s.Add('| ID | Produto | Quantidade | Unidade | Valor faturado |')
        $s.Add('|---|---|---:|---|---:|')
        foreach ($r in $f.Produtos) { $s.Add("| $($r.Id) | $(Rel-Texto $r.Nome) | $($r.Quantidade) | $(Rel-Texto $r.Unidade) | $(Rel-Moeda $r.Valor) |") }
        if (-not $f.Produtos.Count) { $s.Add('| — | Nenhuma saída faturada nesta semana | — | — | — |') }
        $s.Add('')
    }
    $s.Add('### Estoque parado por filial')
    $s.Add('')
    $s.Add('Quantidade positiva e 180 dias ou mais sem entrada nem saída, ou sem ambas as datas. Valor = quantidade × custo médio. Oportunidades abertas referem-se ao produto em todas as filiais (REGRAS_NEGOCIO.md §5).')
    $s.Add('')
    $s.Add('| Filial | Valor total em estoque | Valor parado |')
    $s.Add('|---|---:|---:|')
    foreach ($f in $D.Filiais) { $t=$D.Estoque.TotFilial[$f.Id]; $s.Add("| $(Rel-Texto $f.Nome) | $(Rel-Moeda $t.Valor) | $(Rel-Moeda $t.Parado) |") }
    $s.Add('')
    $s.Add('| Filial | Produto | Quantidade | Dias sem movimento | Valor parado | Oportunidades abertas do produto |')
    $s.Add('|---|---|---:|---:|---:|---:|')
    foreach ($p in ($D.Estoque.Parados | Sort-Object @{Expression={$_.Linha.IdFilial}}, @{Expression={$_.Linha.Produto.Id}})) {
        $l=$p.Linha; $dias=if ($null -eq $l.DiasSemMovimento) {'sem movimento registrado'} else {$l.DiasSemMovimento}
        $s.Add("| $(Rel-Texto $l.Filial) | $($l.Produto.Id) — $(Rel-Texto $l.Produto.Nome) | $($l.Quantidade) | $dias | $(Rel-Moeda $l.Valor) | $($p.Abertas.Qtd) |")
    }
    if (-not $D.Estoque.Parados.Count) { $s.Add('| — | Nenhuma linha parada | — | — | — | — |') }
    return ($s -join "`n") + "`n"
}
function Get-EvidenciasRelatorio($D) {
    $e=New-Object Collections.Generic.List[object]
    $e.Add([ordered]@{id='meta-abaixo80';texto="$($D.Abaixo.Count) vendedores ativos abaixo de 80% no mês; faltam $(Rel-Moeda $D.GapAbaixo) para atingir suas metas cheias, no total.";fonte='Seção 2; metas_2026.xlsx e vendas em vigor.'})
    $e.Add([ordered]@{id='crm-atrasadas';texto="$($D.Funil.Total.QtdAtrasadas) das $($D.Funil.Total.Qtd) oportunidades abertas estão vencidas na foto de $(Rel-Data $D.DataCrm), somando $(Rel-Moeda $D.Funil.Total.ValorAtrasado).";fonte='Seções 3 e 4; crm_oportunidades.xlsx.'})
    $e.Add([ordered]@{id='crm-orfas';texto="$($D.Orfas) oportunidades abertas pertencem a vendedores inativos, com $(Rel-Moeda $D.ValorOrfas) em valor estimado.";fonte='CRM e cadastro de vendedores; REGRAS_NEGOCIO.md §2.4.'})
    $e.Add([ordered]@{id='estoque-parado';texto="$($D.Estoque.Parados.Count) linhas de estoque paradas, com $(Rel-Moeda $D.Estoque.Total.Parado) imobilizados na foto de $(Rel-Data $D.DataEstoque).";fonte='Seção 5; estoque em vigor.'})
    foreach ($f in $D.Filiais) {
        $e.Add([ordered]@{id="estoque-$($f.Id)";texto="$(Rel-Texto $f.Nome): $(Rel-Moeda $D.Estoque.TotFilial[$f.Id].Parado) parados de $(Rel-Moeda $D.Estoque.TotFilial[$f.Id].Valor) em estoque.";fonte='Seção 5; estoque em vigor.'})
    }
    $e.Add([ordered]@{id='datas-fontes';texto="Vendas em $(Rel-Data $D.FimSemana); CRM em $(Rel-Data $D.DataCrm); estoque em $(Rel-Data $D.DataEstoque). Defasagem do CRM frente às vendas: $(($D.FimSemana-$D.DataCrm).Days) dias.";fonte='Datas-base das cargas; REGRAS_NEGOCIO.md §0.3.'})
    return $e.ToArray()
}
function Get-RelatorioHash([string]$Texto) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Texto)))).Replace('-','').ToLowerInvariant() } finally { $sha.Dispose() }
}
function Complete-Relatorio([string]$Calculos, $Pacote, $Analise) {
    if ($Analise.baseSha256 -cne $Pacote.baseSha256) { throw 'Análise desatualizada: gere os três riscos para o pacote atual.' }
    if ([string]::IsNullOrWhiteSpace($Analise.autor) -or @($Analise.riscos).Count -ne 3) { throw 'Informe o autor e exatamente três riscos.' }
    $mapa=@{}; foreach ($e in $Pacote.evidencias) { $mapa[$e.id]=$e }
    $s=New-Object Collections.Generic.List[string]
    $s.Add($Calculos.TrimEnd());$s.Add('');$s.Add('## 6. Três riscos — análise do agente');$s.Add('')
    $s.Add("Autor: $(Rel-Texto $Analise.autor). Interpretação, não cálculo nem previsão garantida. As evidências abaixo são inseridas pelo gerador.")
    $s.Add('');$n=0
    foreach ($r in $Analise.riscos) {
        foreach ($campo in @('titulo','analise','acao')) { if ([string]::IsNullOrWhiteSpace($r.$campo)) { throw "Risco sem $campo." } }
        if (-not @($r.evidencias).Count) { throw 'Cada risco precisa de uma evidência.' }
        $n++;$s.Add("### $n. $(Rel-Texto $r.titulo)");$s.Add('');$s.Add("**Análise:** $(Rel-Texto $r.analise)");$s.Add('')
        foreach ($id in $r.evidencias) {
            if (-not $mapa.ContainsKey([string]$id)) { throw "Evidência inexistente: $id" }
            $e=$mapa[$id];$s.Add("- **Evidência [$id]:** $($e.texto) Fonte: $($e.fonte)")
        }
        $s.Add('');$s.Add("**Ação sugerida:** $(Rel-Texto $r.acao)");$s.Add('')
    }
    return ($s -join "`n") + "`n"
}
