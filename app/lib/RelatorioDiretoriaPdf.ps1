# Relatório semanal da diretoria — formatação em PDF. Só formata: todo número vem de
# RegrasRelatorioDiretoria.ps1. O layout é fixo: as mesmas seções, na mesma ordem, toda semana.
# Depende de Pdf.ps1 e RegrasRelatorioDiretoria.ps1.

# Formato brasileiro fixo, sem depender da cultura instalada na máquina (§0.5).
$script:RelNfi = New-Object Globalization.NumberFormatInfo
$script:RelNfi.NumberDecimalSeparator = ','
$script:RelNfi.NumberGroupSeparator = '.'
$script:RelCorTitulo = @(0.12, 0.2, 0.32)
$script:RelCorSuave = @(0.4, 0.4, 0.4)

function Format-RelMoeda($Valor) {
    if ($null -eq $Valor) { return '—' }
    $v = [math]::Round([decimal]$Valor, 2, [MidpointRounding]::AwayFromZero)
    $sinal = if ($v -lt 0) { '−' } else { '' }
    return "${sinal}R$ " + [math]::Abs($v).ToString('#,##0.00', $script:RelNfi)
}
function Format-RelPct($Fracao) {
    if ($null -eq $Fracao) { return 'sem meta' }
    return ([math]::Round([decimal]$Fracao * 100, 1, [MidpointRounding]::AwayFromZero)).ToString('0.0', $script:RelNfi) + '%'
}
function Format-RelInt($N) { if ($null -eq $N) { return '—' }; return ([decimal]$N).ToString('#,##0', $script:RelNfi) }
function Format-RelData($D) { if ($D) { return $D.ToString('dd/MM/yyyy') }; return '—' }
function Format-RelIndicador($Ind) {
    switch ($Ind.Tipo) {
        'dinheiro' { return (Format-RelMoeda $Ind.Valor) }
        'pct' { if ($null -eq $Ind.Valor) { return '—' }; return (Format-RelPct $Ind.Valor) }
        'inteiro' { return (Format-RelInt $Ind.Valor) }
        default { return [string]$Ind.Valor }
    }
}

function Get-RelNomeVendedor($V) {
    # §2.3 e §3.3: desligado e admitido no ano aparecem marcados.
    $marca = if (-not $V.Ativo) { " (desligado em $(Format-RelData $V.Desligamento))" }
    elseif ($V.Admissao -and $V.Admissao.Year -eq $script:AnoMetas) { " (admissão em $($V.Admissao.ToString('MM/yyyy')))" }
    return "$($V.Id) $($V.Nome)$marca"
}

function Get-RelAlturaTabela([int]$Linhas) { return ($Linhas + 1) * 14 + 4 }

function Confirm-RelBloco($Doc, [double]$Altura) {
    # Mantém o bloco inteiro na mesma página quando ele cabe numa página; se não cabe em nenhuma,
    # só garante o começo (título e algumas linhas) e deixa a tabela quebrar.
    $util = (Get-PdfFundo $Doc) - ($Doc.Margem + 16)
    [void](Confirm-PdfEspaco $Doc $(if ($Altura -le $util) { $Altura } else { 150 }))
}

function Get-RelBlocoRisco($Risco, [double]$Util) {
    # Mede o quadro de um risco antes de desenhar, para ele não quebrar no meio.
    $numero = if ($Risco.Indicador) { Format-RelIndicador $Risco.Indicador } else { $Risco.Numero }
    $fonte = if ($Risco.Indicador) { "Número do cálculo: indicador $($Risco.Chave) ($($Risco.Indicador.Descricao))." } else { "Número calculado na análise: $($Risco.Calculo)" }
    $wNum = 170
    $wTxt = $Util - $wNum - 24
    $b = [pscustomobject]@{
        WNum = $wNum; WTxt = $wTxt
        Numero = Split-PdfTexto $numero ($wNum - 16) 'F2' 15
        Leitura = Split-PdfTexto $Risco.Leitura $wTxt 'F1' 9
        Fonte = Split-PdfTexto $fonte $wTxt 'F1' 7.5
        Altura = 0
    }
    $b.Altura = [math]::Max(16 + 12.2 * $b.Leitura.Count + 10 * $b.Fonte.Count + 12, 22 + 18 * $b.Numero.Count)
    return $b
}

function Add-RelSecao($Doc, [string]$Titulo, [string]$Nota, [double]$Reserva = 60) {
    Confirm-RelBloco $Doc (60 + $Reserva)   # título não fica sozinho no pé da página
    $Doc.Y += 6
    $util = $Doc.Largura - 2 * $Doc.Margem
    Add-PdfRetangulo -Doc $Doc -X $Doc.Margem -Y $Doc.Y -Largura $util -Altura 20 -Cor $script:RelCorTitulo
    Add-PdfTexto -Doc $Doc -X ($Doc.Margem + 8) -Y ($Doc.Y + 14) -Texto $Titulo -Fonte 'F2' -Tamanho 11.5 -Cor @(1, 1, 1)
    $Doc.Y += 26
    if ($Nota) { Add-PdfParagrafo -Doc $Doc -Texto $Nota -Tamanho 8 -Cor $script:RelCorSuave -EspacoDepois 3 }
}

function Add-RelNota($Doc, [string]$Texto) {
    $Doc.Y += 3
    Add-PdfParagrafo -Doc $Doc -Texto $Texto -Tamanho 8 -Cor $script:RelCorSuave -EspacoDepois 2
}

function New-PdfRelatorioDiretoria {
    param([Parameter(Mandatory)]$Rel, $Riscos, [Parameter(Mandatory)][string]$Caminho)   # $Riscos = $null: sem o item 6
    $doc = New-PdfDocumento
    $util = $doc.Largura - 2 * $doc.Margem
    $doc | Add-Member Cabecalho "Data-base: vendas $(Format-RelData $Rel.DataBaseVendas) · CRM $(Format-RelData $Rel.DataBaseCrm) · estoque $(Format-RelData $Rel.DataBaseEstoque)"
    $doc.AoAbrirPagina = {
        param($d)
        Add-PdfTexto -Doc $d -X $d.Margem -Y ($d.Margem + 2) -Texto 'Horizonte Máquinas · Relatório semanal da diretoria' -Fonte 'F2' -Tamanho 8 -Cor $script:RelCorTitulo
        Add-PdfTextoAlinhado -Doc $d -X $d.Margem -Largura ($d.Largura - 2 * $d.Margem) -Y ($d.Margem + 2) -Texto $d.Cabecalho -Alinha 'D' -Tamanho 8 -Cor $script:RelCorSuave
        Add-PdfLinha -Doc $d -X1 $d.Margem -Y1 ($d.Margem + 7) -X2 ($d.Largura - $d.Margem) -Y2 ($d.Margem + 7) -Cor @(0.75, 0.75, 0.75)
        $d.Y = $d.Margem + 16
    }
    Add-PdfPagina $doc

    # ---------------------------------------------------------------- capa
    $mesRotulo = $Rel.RotuloMes + $(if ($Rel.MesParcial) { " (parcial até $($Rel.DataBaseVendas.ToString('dd/MM')))" })
    Add-PdfTexto -Doc $doc -X $doc.Margem -Y ($doc.Y + 20) -Texto 'Relatório semanal da diretoria' -Fonte 'F2' -Tamanho 20 -Cor $script:RelCorTitulo
    $doc.Y += 30
    Add-PdfParagrafo -Doc $doc -Texto "Semana com data-base de vendas em $(Format-RelData $Rel.DataBaseVendas). Metas no acumulado $($Rel.RotuloAno) e no mês de $mesRotulo." -Tamanho 10 -EspacoDepois 2
    if ($Rel.DataBaseCrm -ne $Rel.DataBaseVendas -or $Rel.DataBaseEstoque -ne $Rel.DataBaseVendas) {
        Add-PdfParagrafo -Doc $doc -Texto "Atenção: as fotos não são da mesma data. CRM de $(Format-RelData $Rel.DataBaseCrm) e estoque de $(Format-RelData $Rel.DataBaseEstoque); cada seção usa a data-base da sua fonte (§0.3)." -Fonte 'F2' -Tamanho 9 -Cor @(0.7, 0.15, 0.1) -EspacoDepois 2
    }
    # Faixa de destaques: quatro números dos itens 1 a 5.
    $doc.Y += 6
    $cards = @(
        @('Atingimento da empresa', (Format-RelPct $Rel.Ano.Empresa.Atingimento), $Rel.RotuloAno),
        @("Atingimento em $($Rel.RotuloMes)", (Format-RelPct $Rel.Mes.Empresa.Atingimento), $(if ($Rel.MesParcial) { "parcial até $($Rel.DataBaseVendas.ToString('dd/MM')), meta cheia" } else { 'mês completo' })),
        @('Pipeline ponderado', (Format-RelMoeda $Rel.Funil.Total.Ponderado), "de $(Format-RelMoeda $Rel.Funil.Total.Valor) bruto"),
        @('Com previsão vencida', (Format-RelMoeda $Rel.Funil.Total.ValorAtrasado), "$($Rel.Funil.Total.QtdAtrasadas) oportunidades abertas")
    )
    $wCard = ($util - 3 * 10) / 4
    for ($i = 0; $i -lt 4; $i++) {
        $x = $doc.Margem + $i * ($wCard + 10)
        Add-PdfRetangulo -Doc $doc -X $x -Y $doc.Y -Largura $wCard -Altura 52 -Cor @(0.94, 0.95, 0.97)
        Add-PdfTexto -Doc $doc -X ($x + 8) -Y ($doc.Y + 14) -Texto $cards[$i][0] -Tamanho 8 -Cor $script:RelCorSuave
        Add-PdfTexto -Doc $doc -X ($x + 8) -Y ($doc.Y + 33) -Texto $cards[$i][1] -Fonte 'F2' -Tamanho 15 -Cor $script:RelCorTitulo
        Add-PdfTexto -Doc $doc -X ($x + 8) -Y ($doc.Y + 46) -Texto (Limit-PdfTexto $cards[$i][2] ($wCard - 16) 'F1' 7.5) -Tamanho 7.5 -Cor $script:RelCorSuave
    }
    $doc.Y += 60

    # ---------------------------------------------------------------- item 1
    Add-RelSecao $doc '1. Ranking de vendedores: meta, realizado e atingimento' "Só vendedores ativos (§2.3), do maior para o menor atingimento no acumulado $($Rel.RotuloAno). Atingimento = realizado ÷ meta nos meses com meta (§4.1); vendas canceladas fora (§1). Meses = meses com meta no acumulado (§3.3). Mês: $mesRotulo, com a meta cheia do mês (§4.5). Em vermelho, abaixo de 80% no acumulado." (Get-RelAlturaTabela ($Rel.Ranking.Count + 1))
    $cols = @(
        @{ Titulo = '#'; Largura = 0.025; Alinha = 'D' }, @{ Titulo = 'Vendedor'; Largura = 0.31 }, @{ Titulo = 'Filial'; Largura = 0.075 },
        @{ Titulo = 'Meses'; Largura = 0.05; Alinha = 'C' },
        @{ Titulo = "Meta $($Rel.RotuloAno)"; Largura = 0.105; Alinha = 'D' }, @{ Titulo = 'Realizado'; Largura = 0.105; Alinha = 'D' }, @{ Titulo = 'Ating.'; Largura = 0.065; Alinha = 'D' },
        @{ Titulo = "Meta $($script:RelMesesCurtos[$Rel.MesFinal])"; Largura = 0.1; Alinha = 'D' }, @{ Titulo = 'Realizado'; Largura = 0.1; Alinha = 'D' }, @{ Titulo = 'Ating.'; Largura = 0.065; Alinha = 'D' }
    )
    $pos = 0
    $linhas = foreach ($r in $Rel.Ranking) {
        $pos++
        @{
            Estilo   = if ($null -ne $r.AtingimentoAno -and $r.AtingimentoAno -lt $Rel.Corte) { 'destaque' } else { '' }
            Celulas  = @($pos, (Get-RelNomeVendedor $r.Vendedor), $r.Vendedor.Filial, $r.MesesComMeta.Count,
                (Format-RelMoeda $r.MetaAno), (Format-RelMoeda $r.RealizadoAno), (Format-RelPct $r.AtingimentoAno),
                (Format-RelMoeda $r.MetaMes), (Format-RelMoeda $r.RealizadoMes), (Format-RelPct $r.AtingimentoMes))
        }
    }
    $e = $Rel.Ano.Empresa; $em = $Rel.Mes.Empresa
    $linhas = @($linhas) + @(@{ Estilo = 'total'; Celulas = @('', 'Empresa (com desligados, §4.2)', '', '', (Format-RelMoeda $e.Meta), (Format-RelMoeda $e.RealizadoMesesMeta), (Format-RelPct $e.Atingimento), (Format-RelMoeda $em.Meta), (Format-RelMoeda $em.RealizadoMesesMeta), (Format-RelPct $em.Atingimento)) })
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas $linhas
    if ($Rel.Desligados.Count) {
        $txt = ($Rel.Desligados | ForEach-Object { "$(Get-RelNomeVendedor $_.Vendedor): atingimento $(Format-RelPct $_.AtingimentoAno) nos meses com meta, meta $(Format-RelMoeda $_.MetaAno)" }) -join '; '
        Add-RelNota $doc "Fora do ranking (§2.3), mas dentro do total da empresa: $txt."
    }

    # ---------------------------------------------------------------- item 2
    Add-RelSecao $doc "2. Abaixo de 80% da meta no acumulado $($Rel.RotuloAno)" 'Vendedores ativos com atingimento acumulado menor que 80%, do mais distante para o mais próximo. "Falta para a meta" = meta − realizado (§4.3); "falta para 80%" = 80% da meta − realizado.' (Get-RelAlturaTabela ($Rel.Abaixo.Count + 1))
    $cols = @(
        @{ Titulo = 'Vendedor'; Largura = 0.3 }, @{ Titulo = 'Filial'; Largura = 0.1 },
        @{ Titulo = 'Meta'; Largura = 0.12; Alinha = 'D' }, @{ Titulo = 'Realizado'; Largura = 0.12; Alinha = 'D' }, @{ Titulo = 'Atingimento'; Largura = 0.1; Alinha = 'D' },
        @{ Titulo = 'Falta para a meta'; Largura = 0.13; Alinha = 'D' }, @{ Titulo = 'Falta para 80%'; Largura = 0.13; Alinha = 'D' }
    )
    $gap = [decimal]0; $f80 = [decimal]0
    $linhas = foreach ($a in $Rel.Abaixo) {
        $gap += $a.Gap; $f80 += $a.FaltaCorte
        @{ Celulas = @((Get-RelNomeVendedor $a.Linha.Vendedor), $a.Linha.Vendedor.Filial, (Format-RelMoeda $a.Linha.MetaAno), (Format-RelMoeda $a.Linha.RealizadoAno), (Format-RelPct $a.Linha.AtingimentoAno), (Format-RelMoeda $a.Gap), (Format-RelMoeda $a.FaltaCorte)) }
    }
    if ($Rel.Abaixo.Count) { $linhas = @($linhas) + @(@{ Estilo = 'total'; Celulas = @("$($Rel.Abaixo.Count) vendedor(es)", '', '', '', '', (Format-RelMoeda $gap), (Format-RelMoeda $f80)) }) }
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas @($linhas) -Vazio 'Nenhum vendedor ativo abaixo de 80% da meta no acumulado.'

    # ---------------------------------------------------------------- item 3
    Add-RelSecao $doc '3. Pipeline por etapa: valor bruto e ponderado' "Foto do CRM de $(Format-RelData $Rel.DataBaseCrm), sem filtro de período (§7.3). Bruto = soma do Valor Estimado das abertas; ponderado = soma de valor × probabilidade da etapa (§7). O CRM guarda só a etapa atual (§8.1)." (Get-RelAlturaTabela ($Rel.Funil.Etapas.Count + 1) + 30)
    $cols = @(
        @{ Titulo = 'Etapa'; Largura = 0.24 }, @{ Titulo = 'Probabilidade'; Largura = 0.1; Alinha = 'D' }, @{ Titulo = 'Oportunidades'; Largura = 0.1; Alinha = 'D' },
        @{ Titulo = 'Valor bruto'; Largura = 0.15; Alinha = 'D' }, @{ Titulo = 'Ponderado'; Largura = 0.15; Alinha = 'D' },
        @{ Titulo = 'Previsão vencida'; Largura = 0.11; Alinha = 'D' }, @{ Titulo = 'Valor vencido'; Largura = 0.15; Alinha = 'D' }
    )
    $linhas = foreach ($et in $Rel.Funil.Etapas) {
        $prob = if ($null -ne $et.Probabilidade) { Format-RelPct $et.Probabilidade } else { '—' }
        @{ Celulas = @($et.Etapa, $prob, (Format-RelInt $et.Qtd), (Format-RelMoeda $et.Valor), (Format-RelMoeda $et.Ponderado), (Format-RelInt $et.QtdAtrasadas), (Format-RelMoeda $et.ValorAtrasado)) }
    }
    $t = $Rel.Funil.Total
    $linhas = @($linhas) + @(@{ Estilo = 'total'; Celulas = @('Total aberto', '', (Format-RelInt $t.Qtd), (Format-RelMoeda $t.Valor), (Format-RelMoeda $t.Ponderado), (Format-RelInt $t.QtdAtrasadas), (Format-RelMoeda $t.ValorAtrasado)) })
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas $linhas
    if ($Rel.QtdOrfas) { Add-RelNota $doc "Inclui $($Rel.QtdOrfas) oportunidade(s) órfã(s), de vendedor desligado, somando $(Format-RelMoeda $Rel.ValorOrfas). Continuam no pipeline e só um gestor pode reatribuí-las (§2.4)." }
    Add-RelNota $doc "Fechadas no CRM: $($Rel.Funil.Ganhas.Qtd) ganhas e $($Rel.Funil.Perdidas.Qtd) perdidas; conversão de $(Format-RelPct $Rel.Funil.Conversao) em quantidade (§8)."

    # ---------------------------------------------------------------- item 4
    Add-RelSecao $doc '4. Oportunidades abertas com previsão de fechamento vencida' "Abertas com Previsão de Fechamento anterior à data-base do CRM, $(Format-RelData $Rel.DataBaseCrm) (§7), da mais atrasada para a menos atrasada. Em vermelho, as órfãs (§2.4)." (Get-RelAlturaTabela ($Rel.Atrasadas.Count + 1))
    $cols = @(
        @{ Titulo = 'Oportunidade'; Largura = 0.08 }, @{ Titulo = 'Vendedor (dono atual)'; Largura = 0.19 }, @{ Titulo = 'Cliente'; Largura = 0.19 },
        @{ Titulo = 'Produto'; Largura = 0.2 }, @{ Titulo = 'Etapa'; Largura = 0.1 }, @{ Titulo = 'Valor'; Largura = 0.09; Alinha = 'D' },
        @{ Titulo = 'Previsão'; Largura = 0.08; Alinha = 'C' }, @{ Titulo = 'Dias'; Largura = 0.07; Alinha = 'D' }
    )
    $linhas = foreach ($a in $Rel.Atrasadas) {
        $o = $a.Oportunidade
        @{
            Estilo  = if ($a.Orfa) { 'destaque' } else { '' }
            Celulas = @($o.Id, ("$($a.Vendedor.Id) $($a.Vendedor.Nome)" + $(if ($a.Orfa) { ' (órfã)' })), $a.Cliente.Nome, $o.Produto, $o.Etapa, (Format-RelMoeda $o.Valor), (Format-RelData $o.Previsao), (Format-RelInt $a.DiasAtraso))
        }
    }
    if ($Rel.Atrasadas.Count) { $linhas = @($linhas) + @(@{ Estilo = 'total'; Celulas = @('Total', "$($Rel.Atrasadas.Count) oportunidades", '', '', '', (Format-RelMoeda $t.ValorAtrasado), '', '') }) }
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas @($linhas) -Vazio 'Nenhuma oportunidade aberta com previsão vencida.'

    # ---------------------------------------------------------------- item 5
    Add-RelSecao $doc '5. Dez produtos com maior saída e estoque parado por filial' "Saída = faturamento das vendas faturadas no acumulado $($Rel.RotuloAno) (§1), pela filial da venda. Estoque atual = soma das três filiais na foto de $(Format-RelData $Rel.DataBaseEstoque)." (Get-RelAlturaTabela ($Rel.TopProdutos.Count + 1))
    $cols = @(@{ Titulo = '#'; Largura = 0.03; Alinha = 'D' }, @{ Titulo = 'Produto'; Largura = 0.27 }, @{ Titulo = 'Categoria'; Largura = 0.11 }, @{ Titulo = 'Unidades'; Largura = 0.06; Alinha = 'D' }, @{ Titulo = 'Faturamento'; Largura = 0.12; Alinha = 'D' })
    $wFilial = 0.33 / $Rel.Filiais.Count
    foreach ($f in $Rel.Filiais) { $cols += @{ Titulo = $f.Nome; Largura = $wFilial; Alinha = 'D' } }
    $cols += @{ Titulo = 'Estoque atual'; Largura = 0.08; Alinha = 'D' }
    $pos = 0
    $linhas = foreach ($p in $Rel.TopProdutos) {
        $pos++
        $cel = @($pos, "$($p.IdProduto) $($p.Produto)", $p.Categoria, (Format-RelInt $p.Unidades), (Format-RelMoeda $p.Valor))
        foreach ($f in $Rel.Filiais) { $cel += (Format-RelMoeda $p.PorFilial[$f.Id]) }
        $cel += (Format-RelInt $p.Estoque)
        @{ Estilo = if ($p.Estoque -eq 0) { 'destaque' } else { '' }; Celulas = $cel }
    }
    $pctTop = if ($Rel.FaturamentoAno -gt 0) { $Rel.ValorTop / $Rel.FaturamentoAno } else { $null }
    $total = @('', "Dez produtos: $(Format-RelPct $pctTop) do faturamento", '', '', (Format-RelMoeda $Rel.ValorTop))
    foreach ($f in $Rel.Filiais) { $s = [decimal]0; foreach ($p in $Rel.TopProdutos) { $s += $p.PorFilial[$f.Id] }; $total += (Format-RelMoeda $s) }
    $total += ''
    $linhas = @($linhas) + @(@{ Estilo = 'total'; Celulas = $total })
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas $linhas
    $nota = "Faturamento total do período: $(Format-RelMoeda $Rel.FaturamentoAno)."
    if (@($Rel.TopProdutos | Where-Object { $_.Estoque -eq 0 }).Count) { $nota += ' Em vermelho, produto do top 10 sem estoque nas três filiais (§5).' }
    Add-RelNota $doc $nota

    $doc.Y += 8
    $ve = $Rel.Estoque
    $vazias = @($Rel.ParadosPorFilial | Where-Object { -not $_.Parados.Count }).Count
    Confirm-RelBloco $doc (40 + (Get-RelAlturaTabela ($ve.Parados.Count + $Rel.ParadosPorFilial.Count + $vazias + 1)))
    Add-PdfParagrafo -Doc $doc -Texto "Estoque parado por filial" -Fonte 'F2' -Tamanho 10 -Cor $script:RelCorTitulo -EspacoDepois 1
    Add-PdfParagrafo -Doc $doc -Texto "Parado = linha (produto, filial) com quantidade > 0 e $($ve.DiasParado) dias ou mais sem entrada nem saída, na foto de $(Format-RelData $ve.DataBase) (§5, prazo padrão). Ao lado, as oportunidades abertas do produto em todas as filiais: parado não é o mesmo que sem demanda (§5.4)." -Tamanho 8 -Cor $script:RelCorSuave -EspacoDepois 3
    $cols = @(
        @{ Titulo = 'Produto'; Largura = 0.34 }, @{ Titulo = 'Categoria'; Largura = 0.13 }, @{ Titulo = 'Quantidade'; Largura = 0.07; Alinha = 'D' },
        @{ Titulo = 'Valor imobilizado'; Largura = 0.12; Alinha = 'D' }, @{ Titulo = 'Último movimento'; Largura = 0.12; Alinha = 'C' },
        @{ Titulo = 'Dias parado'; Largura = 0.08; Alinha = 'D' }, @{ Titulo = 'Oportunidades abertas'; Largura = 0.14; Alinha = 'D' }
    )
    $linhas = New-Object Collections.Generic.List[object]
    foreach ($pf in $Rel.ParadosPorFilial) {
        $linhas.Add(@{ Estilo = 'grupo'; Celulas = @("$($pf.Filial.Nome): $($pf.Parados.Count) linha(s) parada(s)", '', '', (Format-RelMoeda $pf.Valor), '', '', '') })
        if (-not $pf.Parados.Count) { $linhas.Add(@{ Celulas = @('Nenhuma linha parada.', '', '', '', '', '', '') }); continue }
        foreach ($x in $pf.Parados) {
            $l = $x.Linha
            $mov = if ($l.UltimoMovimento) { Format-RelData $l.UltimoMovimento } else { 'sem movimento registrado' }
            $ops = if ($null -eq $x.Abertas) { 'CRM indisponível' } elseif ($x.Abertas.Qtd -eq 0) { 'nenhuma' } else { "$($x.Abertas.Qtd) · $(Format-RelMoeda $x.Abertas.Valor)" }
            $linhas.Add(@{ Celulas = @("$($l.Produto.Id) $($l.Produto.Nome)$(if ($l.Produto.Descontinuado) { ' (descontinuado)' })", $l.Produto.Categoria, (Format-RelInt $l.Quantidade), (Format-RelMoeda $l.Valor), $mov, $(if ($null -ne $l.DiasSemMovimento) { Format-RelInt $l.DiasSemMovimento } else { '—' }), $ops) })
        }
    }
    $linhas.Add(@{ Estilo = 'total'; Celulas = @("Total: $($ve.Parados.Count) linha(s) parada(s)", '', '', (Format-RelMoeda $ve.Total.Parado), '', '', '') })
    Add-PdfTabela -Doc $doc -Colunas $cols -Linhas $linhas.ToArray()
    Add-RelNota $doc "Valor imobilizado de todo o estoque na foto: $(Format-RelMoeda $ve.Total.Valor); a parte parada é $(Format-RelPct $(if ($ve.Total.Valor -gt 0) { $ve.Total.Parado / $ve.Total.Valor })) dele (§5)."

    # ---------------------------------------------------------------- item 6
    Add-RelSecao $doc '6. Três riscos da semana (análise do Claude Code, não é cálculo)' 'Os itens 1 a 5 são cálculo e saem iguais para os mesmos dados. Este item é leitura do agente sobre os dados desta semana. Cada risco traz o número que o sustenta: quando vem do catálogo de indicadores do calculo.json, o valor impresso é o do cálculo, e não o digitado na análise.' $(if ($Riscos) { ($Riscos | ForEach-Object { (Get-RelBlocoRisco $_ $util).Altura + 8 } | Measure-Object -Sum).Sum } else { 20 })
    if (-not $Riscos) {
        Add-PdfParagrafo -Doc $doc -Texto 'Análise não incluída nesta emissão.' -Fonte 'F2' -Tamanho 10 -Cor @(0.7, 0.15, 0.1)
    } else {
        $n = 0
        foreach ($r in $Riscos) {
            $n++
            $b = Get-RelBlocoRisco $r $util
            $wNum = $b.WNum; $wTxt = $b.WTxt; $lLeitura = $b.Leitura; $lFonte = $b.Fonte; $lNumero = $b.Numero; $altura = $b.Altura
            [void](Confirm-PdfEspaco $doc ($altura + 8))
            $y0 = $doc.Y
            Add-PdfRetangulo -Doc $doc -X $doc.Margem -Y $y0 -Largura $util -Altura $altura -Cor @(0.97, 0.96, 0.93)
            Add-PdfRetangulo -Doc $doc -X $doc.Margem -Y $y0 -Largura 4 -Altura $altura -Cor @(0.78, 0.45, 0.1)
            Add-PdfTexto -Doc $doc -X ($doc.Margem + 12) -Y ($y0 + 14) -Texto "Risco $n" -Tamanho 7.5 -Cor $script:RelCorSuave
            $yy = $y0 + 32
            foreach ($l in $lNumero) { Add-PdfTexto -Doc $doc -X ($doc.Margem + 12) -Y $yy -Texto $l -Fonte 'F2' -Tamanho 15 -Cor @(0.6, 0.25, 0.05); $yy += 18 }
            $xT = $doc.Margem + $wNum + 12
            Add-PdfTexto -Doc $doc -X $xT -Y ($y0 + 14) -Texto (Limit-PdfTexto $r.Titulo $wTxt 'F2' 10.5) -Fonte 'F2' -Tamanho 10.5 -Cor $script:RelCorTitulo
            $yy = $y0 + 28
            foreach ($l in $lLeitura) { Add-PdfTexto -Doc $doc -X $xT -Y $yy -Texto $l -Tamanho 9 -Cor @(0.15, 0.15, 0.15); $yy += 12.2 }
            $yy += 2
            foreach ($l in $lFonte) { Add-PdfTexto -Doc $doc -X $xT -Y $yy -Texto $l -Tamanho 7.5 -Cor $script:RelCorSuave; $yy += 10 }
            $doc.Y = $y0 + $altura + 8
        }
    }

    # ---------------------------------------------------------------- rodapé de todas as páginas
    $total = $doc.Paginas.Count
    $yRod = $doc.Altura - $doc.Margem + 4
    for ($i = 0; $i -lt $total; $i++) {
        Add-PdfLinha -Doc $doc -X1 $doc.Margem -Y1 ($yRod - 11) -X2 ($doc.Largura - $doc.Margem) -Y2 ($yRod - 11) -Cor @(0.8, 0.8, 0.8) -Pagina $i
        Add-PdfTexto -Doc $doc -X $doc.Margem -Y $yRod -Texto "Impressão digital dos dados: $($Rel.ImpressaoDigital) · itens 1 a 5 calculados por app/relatorio_diretoria.ps1 com as regras do REGRAS_NEGOCIO.md" -Tamanho 7 -Cor $script:RelCorSuave -Pagina $i
        Add-PdfTextoAlinhado -Doc $doc -X $doc.Margem -Largura $util -Y $yRod -Texto "Página $($i + 1) de $total" -Alinha 'D' -Tamanho 7 -Cor $script:RelCorSuave -Pagina $i
    }
    Save-PdfDocumento -Doc $doc -Caminho $Caminho -Titulo "Relatório semanal da diretoria — data-base $(Format-RelData $Rel.DataBaseVendas)" -Produtor 'Horizonte Máquinas — app/relatorio_diretoria.ps1'
}
