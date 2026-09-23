# Testes independentes das somas, recortes, análise e repetibilidade.
$ErrorActionPreference='Stop'
foreach ($m in @('Xlsx','Regras','RegrasPipeline','RegrasEstoque','RegrasRelatorio','RelatorioMarkdown')) { . (Join-Path $PSScriptRoot "../app/lib/$m.ps1") }
$dir=Join-Path $PSScriptRoot '../dados'
$script:n=0
function Ok($Nome,[bool]$Condicao) { $script:n++; if (-not $Condicao) { throw "FALHA: $Nome" } }
function Falha($Nome,[scriptblock]$Acao) { $falhou=$false;try { & $Acao | Out-Null } catch { $falhou=$true };Ok $Nome $falhou }
function Din($v) { return [math]::Round([decimal]::Parse([string]$v,[Globalization.CultureInfo]::InvariantCulture),2,[MidpointRounding]::AwayFromZero) }
$b=Import-BaseComercial $dir;$c=Import-BasePipeline $dir $b;$e=Import-BaseEstoque $dir
$d=Get-DadosRelatorio $b $c $e
$raw=Read-XlsxSheet (Join-Path $dir 'vendas_2026_jan-ago.xlsx') 'Vendas'
$metas=Read-XlsxSheet (Join-Path $dir 'metas_2026.xlsx') 'Metas'
Ok 'semana inclusiva de sete dias' ($d.InicioSemana -eq [datetime]'2026-08-25' -and $d.FimSemana -eq [datetime]'2026-08-31')
Ok 'ranking só ativos' ($d.Ranking.Count -eq 11 -and @($d.Ranking | Where-Object {$_.Vendedor.Id -eq 'V011'}).Count -eq 0)
$mesFat=@($raw | Where-Object { $_.Status -eq 'Faturada' -and ([datetime]$_.Data).Month -eq 8 })
foreach ($l in $d.Ranking) {
    $fat=[decimal]0;foreach ($r in $mesFat) {if ($r.'ID Vendedor' -eq $l.Vendedor.Id) {$fat+=Din $r.'Valor Total'}}
    $meta=[decimal]0;foreach ($m in $metas) {if ($m.'ID Vendedor' -eq $l.Vendedor.Id -and [int]$m.'Mês' -eq 8) {$meta+=Din $m.'Meta (R$)'}}
    Ok 'faturamento mensal independente' ($l.Realizado -eq $fat)
    Ok 'meta mensal independente' ($l.Meta -eq $meta)
    Ok 'atingimento independente' ($l.Atingimento -eq $fat/$meta)
    Ok 'filtro abaixo80' ((@($d.Abaixo | Where-Object {$_.Vendedor.Id -eq $l.Vendedor.Id}).Count -eq 1) -eq ($fat/$meta -lt [decimal]0.8))
}
Ok 'valores mensais não acumulados no ano' ($d.Empresa.Realizado -eq [decimal]'6129704.89' -and $d.Empresa.Meta -eq 7190000)
Ok 'gap mensal e limite' ($d.Abaixo.Count -eq 7 -and $d.GapAbaixo -eq [decimal]'1996662.10')
$ops=Read-XlsxSheet (Join-Path $dir 'crm_oportunidades.xlsx') 'Oportunidades'
foreach ($et in $d.Funil.Etapas) {
    $valor=[decimal]0;$pond=[decimal]0;$q=0
    foreach ($o in $ops) {if ($o.Etapa -eq $et.Etapa) {$q++;$valor+=Din $o.'Valor Estimado';$pond+=[math]::Round((Din $o.'Valor Estimado')*[decimal]::Parse($o.Probabilidade,[Globalization.CultureInfo]::InvariantCulture),2,[MidpointRounding]::AwayFromZero)}}
    Ok 'pipeline bruto independente' ($et.Valor -eq $valor -and $et.Qtd -eq $q)
    Ok 'pipeline ponderado independente' ($et.Ponderado -eq $pond)
}
$ids=@($ops | Where-Object {$_.Etapa -notin @('Fechada Ganha','Fechada Perdida') -and [datetime]$_.'Previsão de Fechamento' -lt [datetime]'2026-08-31'} | ForEach-Object {$_.'ID Oportunidade'} | Sort-Object)
Ok 'atrasadas exatamente iguais à fonte' (($ids -join ',') -eq (($d.Atrasadas | ForEach-Object {$_.Oportunidade.Id} | Sort-Object) -join ','))
$semana=@($mesFat | Where-Object {[datetime]$_.Data -ge [datetime]'2026-08-25'})
foreach ($f in $d.Saidas) {
    $somas=@{}
    foreach ($r in $semana) {if ($r.Filial -eq $f.Filial.Nome) {$id=$r.'ID Produto';$somas[$id]=[long]$somas[$id]+[long]$r.Quantidade}}
    $esperado=@($somas.Keys | Sort-Object -Property @{Expression={$somas[$_]};Descending=$true}, @{Expression={$_}} | Select-Object -First 10)
    Ok 'top10 independente por filial' (($esperado -join ',') -eq (($f.Produtos | ForEach-Object Id) -join ','))
    foreach ($prod in $f.Produtos) {Ok 'quantidade independente' ($prod.Quantidade -eq $somas[$prod.Id])}
}
$est=Read-XlsxSheet (Join-Path $dir 'estoque_2026-08-31.xlsx') 'Estoque'
foreach ($f in $d.Filiais) {
    $valor=[decimal]0
    foreach ($l in $est) {
        if ($l.Filial -ne $f.Nome -or [int]$l.Quantidade -le 0) {continue}
        $datas=@(@($l.'Data da Última Entrada',$l.'Data da Última Saída') | Where-Object {$_} | Sort-Object)
        if (-not $datas.Count -or ([datetime]'2026-08-31'-[datetime]$datas[-1]).Days -ge 180) {$valor += [int]$l.Quantidade*(Din $l.'Custo Médio')}
    }
    Ok 'estoque parado independente por filial' ($d.Estoque.TotFilial[$f.Id].Parado -eq $valor)
}
$md=Get-RelatorioCalculos $d
Ok 'cinco seções fixas' ([regex]::Matches($md,'(?m)^## [1-5]\.').Count -eq 5)
Ok 'unicode legível' ($md.Contains('João Pedro Almeida'))
$original=$b.Vendas;$b.Vendas=@($b.Vendas | Sort-Object Id -Descending)
Ok 'ordem da entrada não muda Markdown' ((Get-RelatorioCalculos (Get-DadosRelatorio $b $c $e)) -ceq $md)
$b.Vendas=$original
$cultura=[Threading.Thread]::CurrentThread.CurrentCulture
try {
    [Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('en-US')
    Ok 'mesmo texto em locale en-US' ((Get-RelatorioCalculos (Get-DadosRelatorio $b $c $e)) -ceq $md)
} finally {[Threading.Thread]::CurrentThread.CurrentCulture=$cultura}
# Limites: igualdade a 80%, mês sem meta e nenhuma venda na semana.
$originalMetas=$b.Metas
$b.Vendas=@([pscustomobject]@{Id='T1';Data=[datetime]'2026-08-01';IdVendedor='V001';IdFilial='F01';IdProduto='P001';Quantidade=1;Valor=[decimal]80;Status='Faturada'})
$b.Metas=@([pscustomobject]@{IdVendedor='V001';Ano=2026;Mes=8;Valor=[decimal]100})
$lim=Get-DadosRelatorio $b $c $e
Ok '80% exatos fora de abaixo80' ($lim.Abaixo.Count -eq 0)
Ok 'ausência de meta fica nula' ($null -eq ($lim.Ranking | Where-Object {$_.Vendedor.Id -eq 'V012'}).Atingimento)
Ok 'sem saídas preserva filiais vazias' (@($lim.Saidas | Where-Object {$_.Produtos.Count}).Count -eq 0 -and $lim.Saidas.Count -eq 3)
$b.Vendas[0].Valor=[decimal]'79.999'
Ok 'limite antes do arredondamento de percentual' ((Get-DadosRelatorio $b $c $e).Abaixo.Count -eq 1)
$b.Vendas=$original;$b.Metas=$originalMetas
# Análise: referências válidas, exatos três riscos e recusa de versão antiga.
$pac=[pscustomobject]@{baseSha256='atual';evidencias=@(Get-EvidenciasRelatorio $d)}
$risco=[pscustomobject]@{titulo='Teste';analise='Interpretação';acao='Verificar';evidencias=@('crm-atrasadas')}
$an=[pscustomobject]@{autor='Teste';baseSha256='atual';riscos=@($risco,$risco,$risco)}
$final=Complete-Relatorio $md $pac $an
Ok 'preserva os cálculos ao finalizar' ($final.StartsWith($md))
Ok 'exatamente seis seções' ([regex]::Matches($final,'(?m)^## [1-6]\.').Count -eq 6)
Ok 'insere número da evidência e fonte' ($final.Contains('R$ 6.813.700,00') -and $final.Contains('Evidência [crm-atrasadas]'))
$an.baseSha256='antigo';Falha 'recusa análise antiga' {Complete-Relatorio $md $pac $an};$an.baseSha256='atual'
$an.riscos=@($risco,$risco);Falha 'recusa só dois riscos' {Complete-Relatorio $md $pac $an};$an.riscos=@($risco,$risco,$risco)
$risco.evidencias=@('inventada');Falha 'recusa evidência inventada' {Complete-Relatorio $md $pac $an};$risco.evidencias=@('crm-atrasadas')
# Integração: duas execuções, importações e finalização em diretório descartável.
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('horizonte_rel_'+[guid]::NewGuid().ToString('N'))
$cmd=Join-Path $PSScriptRoot '../app/relatorio_semanal.ps1'
try {
    $a=Join-Path $tmp 'a';$z=Join-Path $tmp 'z';$cop=Join-Path $tmp 'dados'
    [void](New-Item -ItemType Directory -Path $cop -Force)
    Get-ChildItem -LiteralPath $dir -Filter '*.xlsx' -File | Copy-Item -Destination $cop
    & $cmd -PastaSaida $a -DirDados $cop | Out-Null
    & $cmd -PastaSaida $z -DirDados $cop | Out-Null
    $nome='relatorio-semanal-2026-08-31'
    Ok 'dois cálculos idênticos byte a byte' ((Get-FileHash (Join-Path $a "$nome.calculos.md")).Hash -eq (Get-FileHash (Join-Path $z "$nome.calculos.md")).Hash)
    Ok 'evidências idênticas byte a byte' ((Get-FileHash (Join-Path $a "$nome.evidencias.json")).Hash -eq (Get-FileHash (Join-Path $z "$nome.evidencias.json")).Hash)
    Ok 'sem análise não há relatório final' (-not (Test-Path (Join-Path $a "$nome.md")))
    $pac=Get-Content (Join-Path $a "$nome.evidencias.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $an.baseSha256=$pac.baseSha256
    $analise=Join-Path $tmp 'analise.json';$an | ConvertTo-Json -Depth 8 | Set-Content $analise -Encoding UTF8
    & $cmd -PastaSaida $a -DirDados $cop -AnaliseJson $analise | Out-Null
    Ok 'com análise gera final' (Test-Path (Join-Path $a "$nome.md"))
    $imp=Join-Path $cop 'importacoes';[void](New-Item -ItemType Directory -Path (Join-Path $imp 'vendas') -Force);[void](New-Item -ItemType Directory -Path (Join-Path $imp 'estoque') -Force)
    Copy-Item (Join-Path $dir 'atualizacoes/vendas_2026_setembro.xlsx') (Join-Path $imp 'vendas/20260923-120000__setembro.xlsx')
    Copy-Item (Join-Path $dir 'atualizacoes/estoque_2026-09-19.xlsx') (Join-Path $imp 'estoque/estoque_2026-09-19__20260923-120000.xlsx')
    $bi=Import-BaseComercial $cop $imp;$ci=Import-BasePipeline $cop $bi;$ei=Import-BaseEstoque $cop $imp
    $di=Get-DadosRelatorio $bi $ci $ei
    Ok 'importação muda a semana' ($di.InicioSemana -eq [datetime]'2026-09-13' -and $di.FimSemana -eq [datetime]'2026-09-19')
    Ok 'importação usa realizado setembro' ($di.Empresa.Realizado -eq [decimal]'3850268.28')
    Ok 'datas independentes e parcial' ($di.Parcial -and $di.DataCrm -eq [datetime]'2026-08-31' -and $di.DataEstoque -eq [datetime]'2026-09-19')
    Ok 'estoque atualizado' ($di.Estoque.Total.Parado -eq [decimal]'1802599.08')
    Falha 'CLI recusa análise de agosto sobre setembro' { & $cmd -PastaSaida $z -DirDados $cop -AnaliseJson $analise }
} finally { if (Test-Path $tmp) {Remove-Item -LiteralPath $tmp -Recurse -Force} }
Write-Host "Todas as $script:n conferências do relatório passaram."
