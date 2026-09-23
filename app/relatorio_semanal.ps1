# Gera cálculos/evidências; finaliza só com três riscos redigidos pelo agente.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PastaSaida,
    [string]$DirDados,
    [string]$AnaliseJson
)
$ErrorActionPreference='Stop'
foreach ($mod in @('Xlsx','Regras','RegrasPipeline','RegrasEstoque','RegrasRelatorio','RelatorioMarkdown')) {
    . (Join-Path $PSScriptRoot "lib/$mod.ps1")
}
if (-not $DirDados) { $DirDados=if ($env:HORIZONTE_DADOS) {$env:HORIZONTE_DADOS} else {Join-Path (Split-Path $PSScriptRoot -Parent) 'dados'} }
$DirDados=(Resolve-Path -LiteralPath $DirDados).Path
# Snapshot de assinatura antes/depois: não publicar um relatório com cargas misturadas.
function Assinatura {
    return "$(Get-AssinaturaBase $DirDados (Join-Path $DirDados 'importacoes'))|$(Get-AssinaturaPipeline $DirDados)|$(Get-AssinaturaEstoque $DirDados (Join-Path $DirDados 'importacoes'))"
}
$antes=Assinatura
$b=Import-BaseComercial -DirDados $DirDados -DirImportacoes (Join-Path $DirDados 'importacoes')
if ($b.Erros.Count) { throw ($b.Erros -join "`n") }
$c=Import-BasePipeline -DirDados $DirDados -Comercial $b
$e=Import-BaseEstoque -DirDados $DirDados -DirImportacoes (Join-Path $DirDados 'importacoes')
$d=Get-DadosRelatorio $b $c $e
$md=Get-RelatorioCalculos $d
$ev=@(Get-EvidenciasRelatorio $d)
$pacote=[ordered]@{versao=1;semanaInicio=$d.InicioSemana.ToString('yyyy-MM-dd');semanaFim=$d.FimSemana.ToString('yyyy-MM-dd');mesParcial=$d.Parcial;evidencias=$ev}
$pacote.baseSha256=Get-RelatorioHash ($md+($pacote | ConvertTo-Json -Depth 8 -Compress))
if ((Assinatura) -cne $antes) { throw 'Os dados mudaram durante a geração. Execute novamente.' }
$final=$null
if ($AnaliseJson) {
    $analise=Get-Content -LiteralPath $AnaliseJson -Raw -Encoding UTF8 | ConvertFrom-Json
    $final=Complete-Relatorio $md $pacote $analise
}
$saida=[IO.Path]::GetFullPath($PastaSaida)
if (-not (Test-Path -LiteralPath $saida)) { [void](New-Item -ItemType Directory -Path $saida) }
$prefixo=Join-Path $saida ('relatorio-semanal-'+$d.FimSemana.ToString('yyyy-MM-dd'))
$utf8=New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText("$prefixo.calculos.md",$md,$utf8)
[IO.File]::WriteAllText("$prefixo.evidencias.json",($pacote | ConvertTo-Json -Depth 8),$utf8)
if ($final) {
    [IO.File]::WriteAllText("$prefixo.md",$final,$utf8)
    Write-Output "Relatório completo: $prefixo.md"
} else {
    Write-Output "Cálculos (itens 1–5): $prefixo.calculos.md"
    Write-Output "Evidências para análise: $prefixo.evidencias.json"
    Write-Output 'Etapa 6 pendente: o agente deve redigir três riscos e finalizar com -AnaliseJson. Nenhum relatório completo foi gerado nesta execução.'
}
