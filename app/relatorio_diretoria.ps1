# Relatório semanal da diretoria em PDF — Horizonte Máquinas.
#
# Dois passos (o comando /relatorio-diretoria do Claude Code faz os dois):
#   1. pwsh -NoProfile -File app/relatorio_diretoria.ps1 -SoCalculo
#      Calcula os itens 1 a 5 e grava relatorios/<data-base>/calculo.json (com o catálogo de indicadores).
#   2. O agente lê o calculo.json e escreve relatorios/<data-base>/riscos.json (item 6, três riscos).
#   3. pwsh -NoProfile -File app/relatorio_diretoria.ps1
#      Recalcula os itens 1 a 5, valida o riscos.json e gera relatorios/<data-base>/relatorio_diretoria_<data-base>.pdf.
#
# Os dados são os mesmos que o painel usa: a base de dados/ e as importações confirmadas (§9.8).
# Mesmos dados => mesmos itens 1 a 5, byte a byte (calculo.json e PDF). Nada aqui lê o relógio.

param(
    [string]$DirDados,
    [string]$Saida,            # pasta de saída; padrão: relatorios/<data-base de vendas>
    [string]$Riscos,           # arquivo do item 6; padrão: <Saida>/riscos.json
    [switch]$SoCalculo,        # passo 1: só grava o calculo.json
    [switch]$SemAnalise,       # gera o PDF sem o item 6 (a seção diz que a análise não foi incluída)
    [switch]$SemImportacoes    # só a base, sem as importações (é o que as conferências usam, §9.9)
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot 'lib\Regras.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasPipeline.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasEstoque.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasImportacao.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasRelatorioDiretoria.ps1')
. (Join-Path $PSScriptRoot 'lib\Pdf.ps1')
. (Join-Path $PSScriptRoot 'lib\RelatorioDiretoriaPdf.ps1')

$raiz = Split-Path $PSScriptRoot -Parent
if (-not $DirDados) { $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path $raiz 'dados' } }
$DirDados = (Resolve-Path -LiteralPath $DirDados).Path
$dirImportacoes = if ($SemImportacoes) { $null } else { Get-DirImportacoes $DirDados }

function Stop-Relatorio([string]$Titulo, [string[]]$Erros, [int]$Codigo = 1) {
    Write-Host $Titulo -ForegroundColor Red
    foreach ($e in $Erros) { Write-Host "  - $e" }
    exit $Codigo
}

# --- Carga, com as mesmas validações do painel (§0.2, §2.6, §5, §7.5).
$comercial = Import-BaseComercial -DirDados $DirDados -DirImportacoes $dirImportacoes
if ($comercial.Erros.Count) { Stop-Relatorio 'Vendas/metas com erros de importação; relatório não gerado:' $comercial.Erros }
$pipeline = Import-BasePipeline -DirDados $DirDados -Comercial $comercial -DirImportacoes $dirImportacoes
if ($pipeline.Erros.Count) { Stop-Relatorio 'CRM com erros de importação; relatório não gerado:' $pipeline.Erros }
$estoque = Import-BaseEstoque -DirDados $DirDados -DirImportacoes $dirImportacoes
if ($estoque.Erros.Count) { Stop-Relatorio 'Estoque com erros de importação; relatório não gerado:' $estoque.Erros }

$rel = Get-RelatorioDiretoria -Comercial $comercial -Pipeline $pipeline -Estoque $estoque -DirDados $DirDados
$indicadores = Get-IndicadoresRelatorio $rel
$rotulo = $rel.DataBaseVendas.ToString('yyyy-MM-dd')

if (-not $Saida) { $Saida = Join-Path (Join-Path $raiz 'relatorios') $rotulo }
if (-not (Test-Path -LiteralPath $Saida)) { [void](New-Item -ItemType Directory -Path $Saida) }
$Saida = (Resolve-Path -LiteralPath $Saida).Path
$utf8 = New-Object Text.UTF8Encoding $false
$arqCalculo = Join-Path $Saida 'calculo.json'
[IO.File]::WriteAllText($arqCalculo, (ConvertTo-RelatorioJson $rel $indicadores) + "`n", $utf8)

Write-Host "Data-base: vendas $($rel.DataBaseVendas.ToString('dd/MM/yyyy')), CRM $($rel.DataBaseCrm.ToString('dd/MM/yyyy')), estoque $($rel.DataBaseEstoque.ToString('dd/MM/yyyy'))."
Write-Host "Impressão digital dos dados: $($rel.ImpressaoDigital)"
Write-Host "Itens 1 a 5 calculados: $arqCalculo"

if ($SoCalculo) {
    $arqRiscos = if ($Riscos) { $Riscos } else { Join-Path $Saida 'riscos.json' }
    Write-Host ''
    Write-Host "Próximo passo: escrever $arqRiscos com 3 riscos e ""impressao_digital"": ""$($rel.ImpressaoDigital)""." -ForegroundColor Yellow
    Write-Host 'Depois, rode este comando sem -SoCalculo para gerar o PDF.'
    exit 0
}

$listaRiscos = $null
if (-not $SemAnalise) {
    if (-not $Riscos) { $Riscos = Join-Path $Saida 'riscos.json' }
    $lidos = Read-RiscosRelatorio -Caminho $Riscos -Indicadores $indicadores -ImpressaoDigital $rel.ImpressaoDigital
    if ($lidos.Erros.Count) { Stop-Relatorio 'Item 6 (riscos) inválido; PDF não gerado. Corrija o arquivo ou use -SemAnalise:' $lidos.Erros 2 }
    $listaRiscos = $lidos.Riscos
}

$arqPdf = Join-Path $Saida "relatorio_diretoria_$rotulo.pdf"
New-PdfRelatorioDiretoria -Rel $rel -Riscos $listaRiscos -Caminho $arqPdf
Write-Host "PDF gerado: $arqPdf" -ForegroundColor Green
