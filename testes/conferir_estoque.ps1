# Confere a visão de estoque (/estoque) contra somas diretas nas planilhas brutas.
# Não usa RegrasEstoque.ps1 para recalcular: lê as linhas e soma de novo, por fora.
# Uso: .\conferir_estoque.cmd   (ou: powershell -ExecutionPolicy Bypass -File conferir_estoque.ps1)

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\app\lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Regras.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasPipeline.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasEstoque.ps1')
if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -eq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado $Esperado, painel $Obtido") }
}
function Din([string]$t) { [math]::Round([decimal]::Parse($t, [Globalization.NumberStyles]::Float, $inv), 2, [MidpointRounding]::AwayFromZero) }
function Txt($d) { ([decimal]$d).ToString('0.00', $inv) }

# --- Painel ---
$est = Import-BaseEstoque -DirDados $DirDados
if ($est.Erros.Count) { Write-Host 'Estoque com erros:' -ForegroundColor Red; $est.Erros; exit 1 }
$base = Import-BaseComercial -DirDados $DirDados
$pipe = Import-BasePipeline -DirDados $DirDados -Comercial $base
$abertasPorProduto = Get-AbertasPorProduto $pipe

# --- Fonte bruta ---
$arquivo = 'estoque_2026-08-31.xlsx'
$dataFoto = [datetime]'2026-08-31'   # REGRAS_NEGOCIO.md §0.3: data do nome do arquivo
Confere 'foto usada' $arquivo $est.Arquivo.Name
Confere 'data-base da foto' '2026-08-31' $est.DataBase.ToString('yyyy-MM-dd')
$bruto = Read-XlsxSheet (Join-Path $DirDados $arquivo) 'Estoque'
$catPorId = @{}; foreach ($p in (Read-XlsxSheet (Join-Path $DirDados 'produtos.xlsx') 'Produtos')) { $catPorId[$p.'ID Produto'] = $p.Categoria }
$ops = @((Read-XlsxSheet (Join-Path $DirDados 'crm_oportunidades.xlsx') 'Oportunidades') | Where-Object { $_.Etapa -ne 'Fechada Ganha' -and $_.Etapa -ne 'Fechada Perdida' })

function Get-Dias($r) {
    $datas = @(@($r.'Data da Última Entrada', $r.'Data da Última Saída') | Where-Object { $_ } | Sort-Object)
    if (-not $datas.Count) { return $null }
    return ($dataFoto - [datetime]$datas[-1]).Days
}

foreach ($prazo in @(180, 90, 365, 30)) {
    $vis = Get-VisaoEstoque -Estoque $est -DiasParado $prazo -AbertasPorProduto $abertasPorProduto
    $rot = "prazo $prazo"

    # Valor imobilizado por filial × categoria (categoria de produtos.xlsx pelo ID)
    $soma = @{}; $somaParado = @{}; $total = [decimal]0; $totalParado = [decimal]0
    $idsParados = New-Object Collections.Generic.List[string]
    foreach ($r in $bruto) {
        $q = [int]$r.Quantidade
        $v = $q * (Din $r.'Custo Médio')
        $k = "$($catPorId[$r.'ID Produto'])|$($r.Filial)"
        $soma[$k] = [decimal]$soma[$k] + $v; $total += $v
        $dias = Get-Dias $r
        if ($q -gt 0 -and ($null -eq $dias -or $dias -ge $prazo)) {
            $somaParado[$k] = [decimal]$somaParado[$k] + $v; $totalParado += $v
            $idsParados.Add("$($r.'ID Produto')/$($r.Filial)")
        }
    }
    foreach ($f in $vis.Filiais) {
        foreach ($c in $vis.Categorias) {
            $cel = $vis.Celulas["$c|$($f.Id)"]
            Confere "$rot imobilizado $c/$($f.Nome)" (Txt ([decimal]$soma["$c|$($f.Nome)"])) (Txt $cel.Valor)
            Confere "$rot parado $c/$($f.Nome)" (Txt ([decimal]$somaParado["$c|$($f.Nome)"])) (Txt $cel.Parado)
        }
    }
    Confere "$rot imobilizado total" (Txt $total) (Txt $vis.Total.Valor)
    Confere "$rot parado total" (Txt $totalParado) (Txt $vis.Total.Parado)
    Confere "$rot linhas paradas" (($idsParados | Sort-Object) -join ',') ((@($vis.Parados | ForEach-Object { "$($_.Linha.Produto.Id)/$($_.Linha.Filial)" }) | Sort-Object) -join ',')

    if ($prazo -eq 180) {
        # REGRAS_NEGOCIO.md §6
        Confere 'Valor imobilizado (REGRAS_NEGOCIO.md §6)' '18420192.72' (Txt $vis.Total.Valor)
        Confere 'Linhas paradas (REGRAS_NEGOCIO.md §6)' 7 $vis.Parados.Count
        Confere 'Valor parado (REGRAS_NEGOCIO.md §6)' '1693610.48' (Txt $vis.Total.Parado)
        Confere 'Linhas abaixo do mínimo (REGRAS_NEGOCIO.md §6)' 28 $vis.Abaixo.Count
        Confere 'Produtos sem estoque (REGRAS_NEGOCIO.md §6)' 8 $vis.SemEstoque.Count
        Confere 'P008 Cascavel parado com 2 oportunidades (REGRAS_NEGOCIO.md §5.4)' 2 (@($vis.Parados | Where-Object { $_.Linha.Produto.Id -eq 'P008' -and $_.Linha.Filial -eq 'Cascavel' })[0].Abertas.Qtd)

        # Abaixo do mínimo, linha a linha
        $brutoAbaixo = @($bruto | Where-Object { [int]$_.Quantidade -lt [int]$_.'Estoque Mínimo' } | ForEach-Object { "$($_.'ID Produto')/$($_.Filial)/$([int]$_.'Estoque Mínimo' - [int]$_.Quantidade)" } | Sort-Object)
        Confere 'abaixo do mínimo: linhas e faltas' ($brutoAbaixo -join ',') ((@($vis.Abaixo | ForEach-Object { "$($_.Linha.Produto.Id)/$($_.Linha.Filial)/$($_.Falta)" }) | Sort-Object) -join ',')

        # Sem estoque, produto a produto, com oportunidades abertas
        $brutoSem = @($bruto | Group-Object 'ID Produto' | Where-Object { ($_.Group | ForEach-Object { [int]$_.Quantidade } | Measure-Object -Sum).Sum -eq 0 } | ForEach-Object Name | Sort-Object)
        Confere 'sem estoque: produtos' ($brutoSem -join ',') ((@($vis.SemEstoque | ForEach-Object { $_.Produto.Id }) | Sort-Object) -join ',')
        foreach ($s in $vis.SemEstoque) {
            Confere "sem estoque $($s.Produto.Id): oportunidades abertas" (@($ops | Where-Object { $_.'ID Produto' -eq $s.Produto.Id }).Count) $s.Abertas.Qtd
        }
        foreach ($p in $vis.Parados) {
            Confere "parado $($p.Linha.Produto.Id)/$($p.Linha.Filial): oportunidades abertas" (@($ops | Where-Object { $_.'ID Produto' -eq $p.Linha.Produto.Id }).Count) $p.Abertas.Qtd
        }
    }
}

$total = $script:ok + $script:falhas.Count
if ($script:falhas.Count) {
    Write-Host "$($script:falhas.Count) de $total conferências falharam:" -ForegroundColor Red
    $script:falhas | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
$vis = Get-VisaoEstoque -Estoque $est -DiasParado 180 -AbertasPorProduto $abertasPorProduto
Write-Host "Todas as $total conferências do estoque bateram." -ForegroundColor Green
Write-Host ("Foto de 31/08/2026: imobilizado R$ {0:N2}; {1} linhas paradas (180 dias), R$ {2:N2}; {3} abaixo do mínimo; {4} produtos sem estoque." -f $vis.Total.Valor, $vis.Parados.Count, $vis.Total.Parado, $vis.Abaixo.Count, $vis.SemEstoque.Count)
