# Confere a visão de pipeline (/pipeline) e o funil (/funil) contra somas diretas nas planilhas brutas.
# Não usa RegrasPipeline.ps1 para recalcular: lê as linhas e soma de novo, por fora.
# Uso: .\conferir_pipeline.cmd   (ou: powershell -ExecutionPolicy Bypass -File conferir_pipeline.ps1)

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\app\lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Regras.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasPipeline.ps1')
if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -eq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado $Esperado, painel $Obtido") }
}

# --- Painel ---
$base = Import-BaseComercial -DirDados $DirDados
if ($base.Erros.Count) { Write-Host 'Base comercial com erros:' -ForegroundColor Red; $base.Erros; exit 1 }
$pipe = Import-BasePipeline -DirDados $DirDados -Comercial $base
if ($pipe.Erros.Count) { Write-Host 'CRM com erros:' -ForegroundColor Red; $pipe.Erros; exit 1 }

# --- Fonte bruta ---
$ops = Read-XlsxSheet (Join-Path $DirDados 'crm_oportunidades.xlsx') 'Oportunidades'
$vendas = Read-XlsxSheet (Join-Path $DirDados 'vendas_2026_jan-ago.xlsx') 'Vendas'
$metas = Read-XlsxSheet (Join-Path $DirDados 'metas_2026.xlsx') 'Metas'
$dataBaseCrm = '2026-08-31'   # REGRAS_NEGOCIO.md §0.3
$brutoAbertas = @($ops | Where-Object { $_.Etapa -ne 'Fechada Ganha' -and $_.Etapa -ne 'Fechada Perdida' })

# REGRAS_NEGOCIO.md §6
foreach ($periodo in @(@(1, 8), @(3, 5), @(7, 8))) {
    $de, $ate = $periodo
    $rot = "$de-$ate"
    $desemp = Get-Desempenho -Base $base -MesInicial $de -MesFinal $ate
    $vis = Get-VisaoPipeline -Comercial $base -Desempenho $desemp -Pipeline $pipe
    if ($rot -eq '1-8') {
        Confere 'Abertas (REGRAS_NEGOCIO.md §6)' 68 $vis.Total.QtdAbertas
        Confere 'Atrasadas (REGRAS_NEGOCIO.md §6)' 29 $vis.Total.QtdAtrasadas
        Confere 'Meta jan–ago (REGRAS_NEGOCIO.md §6)' '50720000.00' $vis.Total.Meta.ToString('0.00', $inv)
        Confere 'Realizado jan–ago (REGRAS_NEGOCIO.md §6)' '45295119.16' $vis.Total.RealizadoMesesMeta.ToString('0.00', $inv)
    }
    $somaPipe = [decimal]0; $somaPond = [decimal]0
    foreach ($id in ($base.Vendedores | ForEach-Object Id)) {
        $minhas = @($brutoAbertas | Where-Object { $_.'ID Vendedor' -eq $id })
        $valor = [decimal]0; $pond = [decimal]0
        foreach ($o in $minhas) {
            $v = [decimal]::Parse($o.'Valor Estimado', $inv); $p = [decimal]::Parse($o.Probabilidade, [Globalization.NumberStyles]::Float, $inv)
            $valor += $v; $pond += [math]::Round($v * $p, 2)
        }
        $atr = @($minhas | Where-Object { $_.'Previsão de Fechamento' -lt $dataBaseCrm }).Count
        $meta = [decimal]0; $mesesMeta = @{}
        foreach ($m in $metas) { if ($m.'ID Vendedor' -eq $id -and [int]$m.'Mês' -ge $de -and [int]$m.'Mês' -le $ate) { $meta += [decimal]::Parse($m.'Meta (R$)', [Globalization.NumberStyles]::Float, $inv); $mesesMeta[[int]$m.'Mês'] = $true } }
        $real = [decimal]0
        foreach ($s in $vendas) {
            if ($s.'ID Vendedor' -ne $id -or $s.Status -ne 'Faturada') { continue }
            $mes = [int]$s.Data.Substring(5, 2)
            if ($mesesMeta.ContainsKey($mes)) { $real += [math]::Round([decimal]::Parse($s.'Valor Total', [Globalization.NumberStyles]::Float, $inv), 2, [MidpointRounding]::AwayFromZero) }
        }
        $l = $vis.Linhas | Where-Object { $_.Vendedor.Id -eq $id }
        Confere "$rot $id abertas" $minhas.Count $l.QtdAbertas
        Confere "$rot $id atrasadas" $atr $l.QtdAtrasadas
        Confere "$rot $id pipeline" $valor.ToString('0.00', $inv) $l.Pipeline.ToString('0.00', $inv)
        Confere "$rot $id ponderado" $pond.ToString('0.00', $inv) $l.Ponderado.ToString('0.00', $inv)
        Confere "$rot $id meta" $meta.ToString('0.00', $inv) $l.Meta.ToString('0.00', $inv)
        Confere "$rot $id realizado" $real.ToString('0.00', $inv) $l.RealizadoMesesMeta.ToString('0.00', $inv)
        $somaPipe += $valor; $somaPond += $pond
    }
    Confere "$rot total pipeline" $somaPipe.ToString('0.00', $inv) $vis.Total.Pipeline.ToString('0.00', $inv)
    Confere "$rot total ponderado" $somaPond.ToString('0.00', $inv) $vis.Total.Ponderado.ToString('0.00', $inv)
    Confere "$rot órfãs (vendedor inativo)" 2 $vis.QtdOrfas
}

# REGRAS_NEGOCIO.md §8: funil, recalculado por fora para vários recortes (filial = filial do vendedor, categoria pelo catálogo).
$vendBruto = Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Vendedores'
$filBruto = Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Filiais'
$prodBruto = Read-XlsxSheet (Join-Path $DirDados 'produtos.xlsx') 'Produtos'
$filialDoVendedor = @{}; foreach ($v in $vendBruto) { $filialDoVendedor[$v.'ID Vendedor'] = ($filBruto | Where-Object Filial -eq $v.Filial).'ID Filial' }
$categoriaDoProduto = @{}; foreach ($p in $prodBruto) { $categoriaDoProduto[$p.'ID Produto'] = $p.Categoria }
$recortes = @(
    @{ IdVendedor = ''; IdFilial = ''; Categoria = '' }, @{ IdVendedor = 'V011'; IdFilial = ''; Categoria = '' },
    @{ IdVendedor = ''; IdFilial = 'F02'; Categoria = '' }, @{ IdVendedor = ''; IdFilial = ''; Categoria = 'Tratores' },
    @{ IdVendedor = ''; IdFilial = 'F03'; Categoria = 'Implementos de Solo' }, @{ IdVendedor = 'V002'; IdFilial = 'F01'; Categoria = '' }
)
foreach ($f in $vendBruto) { $recortes += @{ IdVendedor = $f.'ID Vendedor'; IdFilial = ''; Categoria = '' } }
foreach ($c in ($prodBruto | ForEach-Object Categoria | Sort-Object -Unique)) { $recortes += @{ IdVendedor = ''; IdFilial = 'F01'; Categoria = $c } }
foreach ($r in $recortes) {
    $rot = "funil [$($r.IdVendedor)|$($r.IdFilial)|$($r.Categoria)]"
    $fun = Get-VisaoFunil -Comercial $base -Pipeline $pipe -IdVendedor $r.IdVendedor -IdFilial $r.IdFilial -Categoria $r.Categoria
    $sel = @($ops | Where-Object {
            (-not $r.IdVendedor -or $_.'ID Vendedor' -eq $r.IdVendedor) -and
            (-not $r.IdFilial -or $filialDoVendedor[$_.'ID Vendedor'] -eq $r.IdFilial) -and
            (-not $r.Categoria -or $categoriaDoProduto[$_.'ID Produto'] -eq $r.Categoria) })
    foreach ($etapa in @('Prospecção', 'Qualificação', 'Proposta Enviada', 'Negociação')) {
        $na = @($sel | Where-Object Etapa -eq $etapa)
        $valor = [decimal]0; foreach ($o in $na) { $valor += [decimal]::Parse($o.'Valor Estimado', $inv) }
        $e = $fun.Etapas | Where-Object Etapa -eq $etapa
        Confere "$rot $etapa abertas" $na.Count $e.Qtd
        Confere "$rot $etapa atrasadas" (@($na | Where-Object { $_.'Previsão de Fechamento' -lt $dataBaseCrm }).Count) $e.QtdAtrasadas
        Confere "$rot $etapa valor" $valor.ToString('0.00', $inv) $e.Valor.ToString('0.00', $inv)
    }
    $g = @($sel | Where-Object Etapa -eq 'Fechada Ganha').Count; $p = @($sel | Where-Object Etapa -eq 'Fechada Perdida').Count
    Confere "$rot ganhas" $g $fun.Ganhas.Qtd
    Confere "$rot perdidas" $p $fun.Perdidas.Qtd
    Confere "$rot conversão" $(if ($g + $p) { ([decimal]$g / ($g + $p)).ToString('0.0000', $inv) } else { '' }) $(if ($null -ne $fun.Conversao) { $fun.Conversao.ToString('0.0000', $inv) } else { '' })
}
$fun = Get-VisaoFunil -Comercial $base -Pipeline $pipe
Confere 'Funil: conversão 49,7% (REGRAS_NEGOCIO.md §6)' '49.7' ($fun.Conversao * 100).ToString('0.0', $inv)
Confere 'Funil: etapas 16/13/22/17 (REGRAS_NEGOCIO.md §8.5)' '16/13/22/17' (($fun.Etapas | ForEach-Object Qtd) -join '/')
Confere 'Funil: 95 ganhas / 96 perdidas (REGRAS_NEGOCIO.md §6)' '95/96' "$($fun.Ganhas.Qtd)/$($fun.Perdidas.Qtd)"
Confere 'Funil: total aberto = pipeline (REGRAS_NEGOCIO.md §6)' '13870500.00/5605495.00/6813700.00' ('{0}/{1}/{2}' -f $fun.Total.Valor.ToString('0.00', $inv), $fun.Total.Ponderado.ToString('0.00', $inv), $fun.Total.ValorAtrasado.ToString('0.00', $inv))

$total = $script:ok + $script:falhas.Count
if ($script:falhas.Count) {
    Write-Host "$($script:falhas.Count) de $total conferências falharam:" -ForegroundColor Red
    $script:falhas | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
$vis = Get-VisaoPipeline -Comercial $base -Desempenho (Get-Desempenho -Base $base -MesInicial 1 -MesFinal 8) -Pipeline $pipe
Write-Host "Todas as $total conferências do pipeline bateram." -ForegroundColor Green
Write-Host ("CRM de 31/08/2026: {0} abertas, {1} atrasadas, pipeline R$ {2:N2}, ponderado R$ {3:N2}." -f $vis.Total.QtdAbertas, $vis.Total.QtdAtrasadas, $vis.Total.Pipeline, $vis.Total.Ponderado)
