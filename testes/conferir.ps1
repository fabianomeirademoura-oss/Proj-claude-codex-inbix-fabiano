# Conferência: recalcula os números lendo as planilhas por um caminho independente
# de Regras.ps1 (somas diretas sobre as linhas brutas) e compara com o que o painel
# exibe. Também confere contra os números fixados no REGRAS_NEGOCIO.md §6.
# Uso: conferir.cmd     — sai com código 0 se tudo bater, 1 se algo divergir.

param([string]$DirDados)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
. (Join-Path $PSScriptRoot '..\app\lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Regras.ps1')

if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$script:falhas = 0
$script:total = 0

function Confere([string]$Descricao, $Obtido, $Esperado) {
    $script:total++
    $numerico = { param($x) $x -is [decimal] -or $x -is [double] -or $x -is [int] -or $x -is [long] }
    $igual = if ((& $numerico $Obtido) -and (& $numerico $Esperado)) { [decimal]$Obtido -eq [decimal]$Esperado } else { "$Obtido" -eq "$Esperado" }
    if ($igual) { Write-Host ("  OK    {0}: {1}" -f $Descricao, $Obtido) -ForegroundColor Green }
    else { $script:falhas++; Write-Host ("  FALHA {0}: painel={1} esperado={2}" -f $Descricao, $Obtido, $Esperado) -ForegroundColor Red }
}
function Dec([string]$t) { [decimal]::Parse($t, [Globalization.NumberStyles]::Float, $inv) }
function R2([decimal]$x) { [math]::Round($x, 2, [MidpointRounding]::AwayFromZero) }

$base = Import-BaseComercial -DirDados $DirDados
if ($base.Erros.Count) { $base.Erros | ForEach-Object { Write-Host "  ERRO DE IMPORTAÇÃO: $_" -ForegroundColor Red }; exit 1 }

# Linhas brutas, sem passar por Regras.ps1
$vendasBrutas = Read-XlsxSheet (Join-Path $DirDados 'vendas_2026_jan-ago.xlsx') 'Vendas'
$metasBrutas = Read-XlsxSheet (Join-Path $DirDados 'metas_2026.xlsx') 'Metas'
$vendedoresBrutos = Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Vendedores'

function Confere-Periodo([int]$De, [int]$Ate) {
    Write-Host "`n== Período $De a $Ate/2026: planilha bruta x painel ==" -ForegroundColor Cyan
    $d = Get-Desempenho -Base $base -MesInicial $De -MesFinal $Ate
    $somaReal = [decimal]0; $somaMeta = [decimal]0; $somaQtd = 0; $canc = [decimal]0; $qtdCanc = 0
    foreach ($vb in $vendedoresBrutos) {
        $id = $vb.'ID Vendedor'
        $meta = [decimal]0; $mesesMeta = @()
        foreach ($m in $metasBrutas) {
            $mes = [int]$m.'Mês'
            if ($m.'ID Vendedor' -eq $id -and [int]$m.Ano -eq 2026 -and $mes -ge $De -and $mes -le $Ate) { $meta += Dec $m.'Meta (R$)'; $mesesMeta += $mes }
        }
        $real = [decimal]0; $qtd = 0
        foreach ($s in $vendasBrutas) {
            $mes = [datetime]::ParseExact($s.Data, 'yyyy-MM-dd', $inv).Month
            if ($s.'ID Vendedor' -ne $id -or $mes -lt $De -or $mes -gt $Ate) { continue }
            if ($s.Status -eq 'Faturada') { if ($mesesMeta -contains $mes) { $real += Dec $s.'Valor Total' }; $qtd++ }
        }
        $real = R2 $real
        $linha = $d.Linhas | Where-Object { $_.Vendedor.Id -eq $id }
        Confere "$id meta" $linha.Meta (R2 $meta)
        Confere "$id realizado" $linha.RealizadoMesesMeta $real
        Confere "$id vendas faturadas" $linha.QtdVendas $qtd
        $at = if ($meta -gt 0) { [math]::Round($real / $meta * 100, 1) } else { 'sem meta' }
        $atPainel = if ($null -ne $linha.Atingimento) { [math]::Round($linha.Atingimento * 100, 1) } else { 'sem meta' }
        Confere "$id atingimento %" $atPainel $at
        $det = Get-DetalheVendedor -Base $base -IdVendedor $id -MesInicial $De -MesFinal $Ate
        $somaMeses = [decimal]0; foreach ($mm in $det.Meses) { $somaMeses += $mm.Realizado }
        Confere "$id soma do mês a mês = total da linha" $somaMeses $linha.Realizado
        $somaMeta += $meta; $somaQtd += $qtd
    }
    foreach ($s in $vendasBrutas) {
        $mes = [datetime]::ParseExact($s.Data, 'yyyy-MM-dd', $inv).Month
        if ($mes -lt $De -or $mes -gt $Ate) { continue }
        if ($s.Status -eq 'Faturada') { $somaReal += Dec $s.'Valor Total' } else { $canc += Dec $s.'Valor Total'; $qtdCanc++ }
    }
    Confere 'empresa realizado (soma de todas as faturadas)' $d.Empresa.Realizado (R2 $somaReal)
    Confere 'empresa meta' $d.Empresa.Meta (R2 $somaMeta)
    Confere 'empresa vendas faturadas' $d.Empresa.QtdVendas $somaQtd
    Confere 'canceladas (valor)' $d.ValorCanceladas (R2 $canc)
    Confere 'canceladas (quantidade)' $d.QtdCanceladas $qtdCanc
    $somaLinhas = [decimal]0; foreach ($l in $d.Linhas) { $somaLinhas += $l.Realizado }
    Confere 'soma das linhas do quadro = total da empresa' $somaLinhas $d.Empresa.Realizado
    return $d
}

$d = Confere-Periodo 1 8
[void](Confere-Periodo 3 5)
[void](Confere-Periodo 7 8)

Write-Host "`n== Jan–ago x números fixados no REGRAS_NEGOCIO.md §6 ==" -ForegroundColor Cyan
Confere 'vendas realizadas' $d.Empresa.QtdVendas 723
Confere 'realizado' $d.Empresa.Realizado 45295119.16
Confere 'canceladas' "$($d.QtdCanceladas) / $($d.ValorCanceladas)" '12 / 407327.39'
Confere 'meta jan–ago' $d.Empresa.Meta 50720000.00
Confere 'atingimento empresa %' ([math]::Round($d.Empresa.Atingimento * 100, 1)) 89.3
$porId = @{}; foreach ($l in $d.Linhas) { $porId[$l.Vendedor.Id] = $l }
Confere 'V003 atingimento %' ([math]::Round($porId['V003'].Atingimento * 100, 1)) 114.9
Confere 'V011 atingimento % (jan–abr)' ([math]::Round($porId['V011'].Atingimento * 100, 1)) 62.8
Confere 'V012 atingimento % (jul–ago)' ([math]::Round($porId['V012'].Atingimento * 100, 1)) 80.1

Write-Host "`n== Rankings (regras §2.3 e §3.3) ==" -ForegroundColor Cyan
$rf = Get-RankingFaturamento $d
$ra = Get-RankingAtingimento $d
Confere '1º em faturamento' $rf[0].Vendedor.Id 'V003'
Confere '1º em atingimento' $ra[0].Vendedor.Id 'V003'
Confere 'desligado V011 fora dos rankings padrão' (@($rf + $ra | Where-Object { $_.Vendedor.Id -eq 'V011' }).Count) 0
Confere 'admitida V012 dentro do ranking de atingimento' (@($ra | Where-Object { $_.Vendedor.Id -eq 'V012' }).Count) 1
Confere 'ranking padrão tem só os 11 ativos' $rf.Count 11
Confere 'com desligados, ranking tem 12' (Get-RankingFaturamento $d -IncluirDesligados).Count 12

Write-Host ""
if ($script:falhas) { Write-Host "$script:falhas de $script:total conferências FALHARAM." -ForegroundColor Red; exit 1 }
Write-Host "Todas as $script:total conferências bateram." -ForegroundColor Green
exit 0
