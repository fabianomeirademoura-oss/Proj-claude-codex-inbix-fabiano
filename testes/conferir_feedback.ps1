# Conferência do feedback semanal (REGRAS_NEGOCIO.md §11): recalcula cada indicador
# somando as linhas brutas das planilhas, sem passar por RegrasFeedback.ps1, e compara.
# Dois cenários: a base (data-base 31/08, agosto fechado) e a base + setembro da pasta
# atualizacoes/ (data-base 19/09, mês em andamento). Sai com 0 se tudo bater, 1 se não.

param([string]$DirDados)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
foreach ($mod in @('Xlsx', 'Regras', 'RegrasPipeline', 'RegrasFeedback')) { . (Join-Path $PSScriptRoot "..\app\lib\$mod.ps1") }

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
    else { $script:falhas++; Write-Host ("  FALHA {0}: obtido={1} esperado={2}" -f $Descricao, $Obtido, $Esperado) -ForegroundColor Red }
}
function Dec([string]$t) { [math]::Round([decimal]::Parse($t, [Globalization.NumberStyles]::Float, $inv), 2, [MidpointRounding]::AwayFromZero) }
function R2([decimal]$x) { [math]::Round($x, 2, [MidpointRounding]::AwayFromZero) }
function Pct1($f) { if ($null -eq $f) { 'sem' } else { [math]::Round([decimal]$f * 100, 1, [MidpointRounding]::AwayFromZero) } }
function ParaData([string]$t) { [datetime]::ParseExact($t, [string[]]@('yyyy-MM-dd', 'dd/MM/yyyy'), $inv, [Globalization.DateTimeStyles]::None) }

# Linhas brutas
$metasBrutas = Read-XlsxSheet (Join-Path $DirDados 'metas_2026.xlsx') 'Metas'
$vendedoresBrutos = Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Vendedores'
$crmBruto = Read-XlsxSheet (Join-Path $DirDados 'crm_oportunidades.xlsx') 'Oportunidades'
$setembro = Join-Path $DirDados 'atualizacoes\vendas_2026_setembro.xlsx'
$ativos = @($vendedoresBrutos | Where-Object Status -eq 'Ativo' | ForEach-Object 'ID Vendedor' | Sort-Object)

function Confere-Cenario([string]$Nome, [string[]]$ArquivosVendas, [string[]]$Adicionais) {
    Write-Host "`n== $Nome ==" -ForegroundColor Cyan
    $base = Import-BaseComercial -DirDados $DirDados -VendasAdicionais $Adicionais
    if ($base.Erros.Count) { $base.Erros | ForEach-Object { Write-Host "  ERRO DE IMPORTAÇÃO: $_" -ForegroundColor Red }; exit 1 }
    $crm = Import-BasePipeline -DirDados $DirDados -Comercial $base
    if ($crm.Erros.Count) { $crm.Erros | ForEach-Object { Write-Host "  ERRO DE IMPORTAÇÃO: $_" -ForegroundColor Red }; exit 1 }
    $fb = Get-FeedbackSemanal -Comercial $base -Pipeline $crm

    $vendas = New-Object Collections.Generic.List[object]
    foreach ($a in $ArquivosVendas) { foreach ($r in (Read-XlsxSheet $a 'Vendas')) { $vendas.Add($r) } }
    $faturadas = @($vendas | Where-Object Status -eq 'Faturada')
    $db = ($vendas | ForEach-Object { ParaData $_.Data } | Measure-Object -Maximum).Maximum
    $emAndamento = $db.AddDays(1).Month -eq $db.Month
    $ultimoFechado = if ($emAndamento) { $db.Month - 1 } else { $db.Month }
    Confere 'data-base de vendas' $fb.DataBaseVendas.ToString('dd/MM/yyyy') $db.ToString('dd/MM/yyyy')
    Confere 'mês parcial (§4.5)' $fb.MesParcial $emAndamento
    Confere 'só ativos recebem feedback (§11.1)' (($fb.Vendedores | ForEach-Object { $_.Vendedor.Id }) -join ',') ($ativos -join ',')

    foreach ($x in $fb.Vendedores) {
        $id = $x.Vendedor.Id
        $metas = @{}; foreach ($m in ($metasBrutas | Where-Object { $_.'ID Vendedor' -eq $id -and $_.Ano -eq '2026' })) { $metas[[int]$m.'Mês'] = Dec $m.'Meta (R$)' }
        $real = @{}; foreach ($v in ($faturadas | Where-Object 'ID Vendedor' -eq $id)) { $mes = (ParaData $v.Data).Month; $real[$mes] = [decimal]$real[$mes] + (Dec $v.'Valor Total') }

        # §11.2 semana
        $ini = $db.AddDays(-6)
        $sem = [decimal]0; $qtd = 0
        foreach ($v in ($faturadas | Where-Object 'ID Vendedor' -eq $id)) { $d = ParaData $v.Data; if ($d -ge $ini -and $d -le $db) { $sem += Dec $v.'Valor Total'; $qtd++ } }
        $ant = [decimal]0
        foreach ($v in ($faturadas | Where-Object 'ID Vendedor' -eq $id)) { $d = ParaData $v.Data; if ($d -ge $db.AddDays(-13) -and $d -le $db.AddDays(-7)) { $ant += Dec $v.'Valor Total' } }
        Confere "$id semana (R$)" $x.Semana.Realizado $sem
        Confere "$id semana (vendas)" $x.Semana.QtdVendas $qtd
        Confere "$id semana anterior (R$)" $x.Semana.RealizadoAnterior $ant

        # §11.3 mês da data-base
        $metaMes = $metas[$db.Month]
        Confere "$id mês: realizado" $x.MesAtual.Realizado ([decimal]$real[$db.Month])
        Confere "$id mês: atingimento %" (Pct1 $x.MesAtual.Atingimento) (Pct1 $(if ($metaMes) { [decimal]$real[$db.Month] / $metaMes }))
        Confere "$id mês: no ritmo" $x.MesAtual.NoRitmo $(if ($metaMes) { ([decimal]$real[$db.Month] / $metaMes) -ge ([decimal]$db.Day / [datetime]::DaysInMonth($db.Year, $db.Month)) })

        # §11.4 ano e projeção
        $mesesAno = @($metas.Keys | Sort-Object)
        $fech = @($mesesAno | Where-Object { $_ -le $ultimoFechado })
        $metaAno = [decimal]0; foreach ($k in $mesesAno) { $metaAno += $metas[$k] }
        $realFech = [decimal]0; foreach ($k in $fech) { $realFech += [decimal]$real[$k] }
        $realAno = [decimal]0; foreach ($k in $mesesAno) { if ($k -le $db.Month) { $realAno += [decimal]$real[$k] } }
        $media = if ($fech.Count) { R2 ($realFech / $fech.Count) }
        $rest = $mesesAno.Count - $fech.Count
        Confere "$id meta do ano" $x.Ano.Meta $metaAno
        Confere "$id realizado no ano" $x.Ano.Realizado $realAno
        Confere "$id média mensal" $x.Fechados.MediaMensal $media
        Confere "$id projeção" $x.Ano.Projecao $(if ($null -ne $media) { $media * $mesesAno.Count })
        Confere "$id necessário por mês" $x.Ano.NecessarioPorMes $(if ($rest) { R2 ([math]::Max([decimal]0, $metaAno - $realFech) / $rest) })

        # §11.5 pipeline do vendedor, direto do CRM bruto
        $abertas = @($crmBruto | Where-Object { $_.'ID Vendedor' -eq $id -and $_.Etapa -notin @('Fechada Ganha', 'Fechada Perdida') })
        $pv = [decimal]0; $pp = [decimal]0; $atr = 0
        foreach ($o in $abertas) {
            $pv += Dec $o.'Valor Estimado'
            $pp += R2 ((Dec $o.'Valor Estimado') * [decimal]::Parse($o.Probabilidade, [Globalization.NumberStyles]::Float, $inv))
            if ((ParaData $o.'Previsão de Fechamento') -lt [datetime]'2026-08-31') { $atr++ }
        }
        Confere "$id pipeline" $x.Pipeline.Valor $pv
        Confere "$id ponderado" $x.Pipeline.Ponderado $pp
        Confere "$id atrasadas" $x.Pipeline.QtdAtrasadas $atr
        Confere "$id oportunidades listadas" $x.Pipeline.Oportunidades.Count $abertas.Count
    }
    return $fb
}

$base = Confere-Cenario 'Base: jan–ago, data-base 31/08' @((Join-Path $DirDados 'vendas_2026_jan-ago.xlsx')) @()

Write-Host "`n== Números fixados (REGRAS_NEGOCIO.md §6) ==" -ForegroundColor Cyan
$v003 = $base.Vendedores | Where-Object { $_.Vendedor.Id -eq 'V003' }
$v012 = $base.Vendedores | Where-Object { $_.Vendedor.Id -eq 'V012' }
Confere 'V003 atingimento jan–ago' (Pct1 $v003.Fechados.Atingimento) 114.9
Confere 'V003 realizado jan–ago' $v003.Fechados.Realizado 6474285.51
Confere 'V012 atingimento jul–ago' (Pct1 $v012.Fechados.Atingimento) 80.1
Confere 'V012 meses com meta no ano (§3.1)' $v012.Ano.MesesComMeta.Count 6
Confere 'V012 marcada como admitida no ano' $v012.Admitida $true
$pipe = [decimal]0; $pond = [decimal]0; $qtd = 0; $atr = 0
foreach ($x in $base.Vendedores) { $pipe += $x.Pipeline.Valor; $pond += $x.Pipeline.Ponderado; $qtd += $x.Pipeline.QtdAbertas; $atr += $x.Pipeline.QtdAtrasadas }
Confere 'ativos + órfãs = pipeline aberto (§6)' ($pipe + 255300) 13870500
Confere 'ativos + 2 órfãs = abertas (§6)' ($qtd + 2) 68
Confere 'agosto fechado: sem mês parcial' $base.MesParcial $false
Confere 'CRM não defasado na base' $base.CrmDefasado $false
$erro = $null; try { Get-FeedbackSemanal -Comercial (Import-BaseComercial -DirDados $DirDados) -Pipeline (Import-BasePipeline -DirDados $DirDados -Comercial (Import-BaseComercial -DirDados $DirDados)) -IdVendedor 'V011' | Out-Null } catch { $erro = $_.Exception.Message }
Confere 'V011 desligado não recebe feedback (§11.1)' ([bool]$erro) $true

if (Test-Path -LiteralPath $setembro) {
    $set = Confere-Cenario 'Base + setembro (atualizacoes/), data-base 19/09' @((Join-Path $DirDados 'vendas_2026_jan-ago.xlsx'), $setembro) @($setembro)
    Confere 'setembro em andamento: mês parcial' $set.MesParcial $true
    Confere 'CRM de 31/08 defasado em relação a 19/09 (§11.5)' $set.CrmDefasado $true
}
else { Write-Host "  (sem ${setembro}: cenário de setembro não conferido)" -ForegroundColor Yellow }

Write-Host ''
if ($script:falhas) { Write-Host "$($script:falhas) de $($script:total) verificações FALHARAM." -ForegroundColor Red; exit 1 }
Write-Host "Todas as $($script:total) verificações bateram." -ForegroundColor Green
exit 0
