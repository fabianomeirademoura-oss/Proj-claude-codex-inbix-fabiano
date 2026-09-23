# Confere o relatório semanal da diretoria (app/relatorio_diretoria.ps1).
# - Itens 1 a 5 contra somas diretas nas planilhas brutas, por fora das regras, e contra a §6.
# - Repetibilidade: mesmos dados => calculo.json e PDF idênticos, byte a byte.
# - Item 6: o arquivo de riscos é recusado quando não traz 3 riscos com o número que os sustenta.
# - Setembro importado (numa cópia temporária de dados/): mês parcial e data-base 19/09.
# Uso: pwsh -NoProfile -File testes/conferir_relatorio_diretoria.ps1

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\app\lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Regras.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasPipeline.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasEstoque.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasImportacao.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasRelatorioDiretoria.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Pdf.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RelatorioDiretoriaPdf.ps1')
if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -eq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado '$Esperado', obtido '$Obtido'") }
}
function Din([string]$t) { [math]::Round([decimal]::Parse($t, [Globalization.NumberStyles]::Float, $inv), 2, [MidpointRounding]::AwayFromZero) }
function Txt($d) { if ($null -eq $d) { return '' }; ([decimal]$d).ToString('0.00', $inv) }
function Fr($d) { if ($null -eq $d) { return '' }; ([decimal]$d).ToString('0.000000', $inv) }
function Get-TextoPdf([string]$Caminho) {
    # Bytes WinAnsi -> texto: os acentos do português são os mesmos do Latin-1.
    $b = [IO.File]::ReadAllBytes($Caminho); $sb = New-Object Text.StringBuilder
    foreach ($x in $b) { [void]$sb.Append([char]$x) }
    return $sb.ToString()
}
$pwsh = (Get-Process -Id $PID).Path
$app = Join-Path (Split-Path $PSScriptRoot -Parent) 'app\relatorio_diretoria.ps1'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte_rel_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null

try {
    # ================================================================ relatório sobre a base (§9.9)
    $c = Import-BaseComercial -DirDados $DirDados
    $p = Import-BasePipeline -DirDados $DirDados -Comercial $c
    $e = Import-BaseEstoque -DirDados $DirDados
    $rel = Get-RelatorioDiretoria -Comercial $c -Pipeline $p -Estoque $e -DirDados $DirDados
    $ind = Get-IndicadoresRelatorio $rel

    # --- REGRAS_NEGOCIO.md §6
    Confere 'data-base de vendas' '2026-08-31' $rel.DataBaseVendas.ToString('yyyy-MM-dd')
    Confere 'período acumulado' 'jan–ago/2026' $rel.RotuloAno
    Confere 'mês da data-base' 'agosto/2026' $rel.RotuloMes
    Confere 'agosto não é parcial (§4.5)' 'False' $rel.MesParcial
    Confere 'Meta jan–ago (§6)' '50720000.00' (Txt $rel.Ano.Empresa.Meta)
    Confere 'Realizado jan–ago (§6)' '45295119.16' (Txt $rel.Ano.Empresa.RealizadoMesesMeta)
    Confere 'Atingimento da empresa (§6)' '89,3%' (Format-RelPct $rel.Ano.Empresa.Atingimento)
    Confere 'V011 fora do ranking, atingimento 62,8% (§6, §2.3)' 'V011 62,8%' "$($rel.Desligados[0].Vendedor.Id) $(Format-RelPct $rel.Desligados[0].AtingimentoAno)"
    Confere 'V003 lidera com 114,9% (§6)' 'V003 114,9%' "$($rel.Ranking[0].Vendedor.Id) $(Format-RelPct $rel.Ranking[0].AtingimentoAno)"
    Confere 'Pipeline aberto: quantidade (§6)' 68 $rel.Funil.Total.Qtd
    Confere 'Pipeline aberto: bruto (§6)' '13870500.00' (Txt $rel.Funil.Total.Valor)
    Confere 'Pipeline aberto: ponderado (§6)' '5605495.00' (Txt $rel.Funil.Total.Ponderado)
    Confere 'Atrasadas: quantidade (§6)' 29 $rel.Atrasadas.Count
    Confere 'Atrasadas: valor (§6)' '6813700.00' (Txt $rel.Funil.Total.ValorAtrasado)
    Confere 'Órfãs (§6)' '2 255300.00' "$($rel.QtdOrfas) $(Txt $rel.ValorOrfas)"
    Confere 'Etapas do funil (§8.5)' 'Prospecção 16,Qualificação 13,Proposta Enviada 22,Negociação 17' (($rel.Funil.Etapas | ForEach-Object { "$($_.Etapa) $($_.Qtd)" }) -join ',')
    Confere 'Linhas paradas (§6)' 7 $rel.Estoque.Parados.Count
    Confere 'Valor parado (§6)' '1693610.48' (Txt $rel.Estoque.Total.Parado)

    # --- Fonte bruta
    $vendBruto = Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Vendedores'
    $filialId = @{}; foreach ($f in (Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Filiais')) { $filialId[$f.Filial] = $f.'ID Filial' }
    $metasBruto = Read-XlsxSheet (Join-Path $DirDados 'metas_2026.xlsx') 'Metas'
    $linhasVendas = Read-XlsxSheet (Join-Path $DirDados 'vendas_2026_jan-ago.xlsx') 'Vendas'
    $vendasBruto = @($linhasVendas | Where-Object { $_.Status -eq 'Faturada' })
    function Get-MesBruto($v) { ([datetime]::ParseExact($v.Data, 'yyyy-MM-dd', $inv)).Month }

    function Get-DesempenhoBruto([int[]]$Meses) {
        $r = @{}
        foreach ($v in $vendBruto) {
            $mm = @($metasBruto | Where-Object { $_.'ID Vendedor' -eq $v.'ID Vendedor' -and [int]$_.Ano -eq 2026 -and [int]$_.'Mês' -in $Meses })
            $meta = [decimal]0; foreach ($m in $mm) { $meta += Din $m.'Meta (R$)' }
            $mesesMeta = @($mm | ForEach-Object { [int]$_.'Mês' })
            $real = [decimal]0
            foreach ($x in $vendasBruto) { if ($x.'ID Vendedor' -eq $v.'ID Vendedor' -and (Get-MesBruto $x) -in $mesesMeta) { $real += Din $x.'Valor Total' } }
            $r[$v.'ID Vendedor'] = [pscustomobject]@{ Id = $v.'ID Vendedor'; Ativo = ($v.Status -eq 'Ativo'); Meta = $meta; Real = $real; Ating = $(if ($meta -gt 0) { $real / $meta }); Meses = $mesesMeta.Count }
        }
        return $r
    }
    $ano = Get-DesempenhoBruto @(1..8)
    $mes = Get-DesempenhoBruto @(8)

    # --- Item 1: ranking dos ativos pelo atingimento do ano, com o mês ao lado
    $esperado = @($ano.Values | Where-Object Ativo | Sort-Object -Property @{ Expression = { $_.Ating }; Descending = $true }, @{ Expression = { $_.Id } })
    Confere 'item 1: ordem do ranking' (($esperado | ForEach-Object Id) -join ',') (($rel.Ranking | ForEach-Object { $_.Vendedor.Id }) -join ',')
    foreach ($r in $rel.Ranking) {
        $id = $r.Vendedor.Id; $a = $ano[$id]; $m = $mes[$id]
        Confere "item 1 $id meta do ano" (Txt $a.Meta) (Txt $r.MetaAno)
        Confere "item 1 $id realizado do ano" (Txt $a.Real) (Txt $r.RealizadoAno)
        Confere "item 1 $id atingimento do ano" (Fr $a.Ating) (Fr $r.AtingimentoAno)
        Confere "item 1 $id meses com meta" $a.Meses $r.MesesComMeta.Count
        Confere "item 1 $id meta do mês" (Txt $(if ($m.Meta -gt 0) { $m.Meta })) (Txt $r.MetaMes)
        Confere "item 1 $id realizado do mês" (Txt $m.Real) (Txt $r.RealizadoMes)
        Confere "item 1 $id atingimento do mês" (Fr $m.Ating) (Fr $r.AtingimentoMes)
    }
    $somaMetaMes = [decimal]0; $somaRealMes = [decimal]0; foreach ($x in $mes.Values) { $somaMetaMes += $x.Meta; $somaRealMes += $x.Real }
    Confere 'item 1: meta da empresa no mês (com desligados)' (Txt $somaMetaMes) (Txt $rel.Mes.Empresa.Meta)
    Confere 'item 1: realizado da empresa no mês' (Txt $somaRealMes) (Txt $rel.Mes.Empresa.RealizadoMesesMeta)

    # --- Item 2: abaixo de 80% no acumulado (só ativos), com o que falta
    $abaixo = @($esperado | Where-Object { $null -ne $_.Ating -and $_.Ating -lt 0.8 } | Sort-Object -Property @{ Expression = { $_.Ating } }, @{ Expression = { $_.Id } })
    Confere 'item 2: vendedores abaixo de 80%' (($abaixo | ForEach-Object Id) -join ',') (($rel.Abaixo | ForEach-Object { $_.Linha.Vendedor.Id }) -join ',')
    foreach ($x in $abaixo) {
        $a = @($rel.Abaixo | Where-Object { $_.Linha.Vendedor.Id -eq $x.Id })[0]
        Confere "item 2 $($x.Id) falta para a meta" (Txt ($x.Meta - $x.Real)) (Txt $a.Gap)
        Confere "item 2 $($x.Id) falta para 80%" (Txt ([math]::Round($x.Meta * 0.8, 2) - $x.Real)) (Txt $a.FaltaCorte)
    }
    Confere 'item 2: V012 com 80,1% fica fora (corte estrito)' 0 @($rel.Abaixo | Where-Object { $_.Linha.Vendedor.Id -eq 'V012' }).Count
    Confere 'item 2: V011 desligado fica fora (§2.3)' 0 @($rel.Abaixo | Where-Object { $_.Linha.Vendedor.Id -eq 'V011' }).Count

    # --- Itens 3 e 4: CRM bruto
    $linhasCrm = Read-XlsxSheet (Join-Path $DirDados 'crm_oportunidades.xlsx') 'Oportunidades'
    $crm = @($linhasCrm | Where-Object { $_.Etapa -notin @('Fechada Ganha', 'Fechada Perdida') })
    $dbCrm = [datetime]'2026-08-31'
    foreach ($et in $rel.Funil.Etapas) {
        $ls = @($crm | Where-Object { $_.Etapa -eq $et.Etapa })
        $b = [decimal]0; $pd = [decimal]0; $at = [decimal]0; $nat = 0
        foreach ($o in $ls) {
            $v = Din $o.'Valor Estimado'; $b += $v
            $pd += [math]::Round($v * [decimal]::Parse($o.Probabilidade, $inv), 2, [MidpointRounding]::AwayFromZero)
            if ([datetime]::ParseExact($o.'Previsão de Fechamento', 'yyyy-MM-dd', $inv) -lt $dbCrm) { $at += $v; $nat++ }
        }
        Confere "item 3 $($et.Etapa): quantidade" $ls.Count $et.Qtd
        Confere "item 3 $($et.Etapa): bruto" (Txt $b) (Txt $et.Valor)
        Confere "item 3 $($et.Etapa): ponderado" (Txt $pd) (Txt $et.Ponderado)
        Confere "item 3 $($et.Etapa): vencidas" "$nat $(Txt $at)" "$($et.QtdAtrasadas) $(Txt $et.ValorAtrasado)"
    }
    $atrBruto = @($crm | ForEach-Object { [pscustomobject]@{ Id = $_.'ID Oportunidade'; Dias = ($dbCrm - [datetime]::ParseExact($_.'Previsão de Fechamento', 'yyyy-MM-dd', $inv)).Days } } | Where-Object { $_.Dias -gt 0 } |
        Sort-Object -Property @{ Expression = { $_.Dias }; Descending = $true }, @{ Expression = { $_.Id } })
    Confere 'item 4: oportunidades vencidas, na ordem' (($atrBruto | ForEach-Object { "$($_.Id)/$($_.Dias)" }) -join ',') (($rel.Atrasadas | ForEach-Object { "$($_.Oportunidade.Id)/$($_.DiasAtraso)" }) -join ',')
    Confere 'item 4: órfãs marcadas (§2.4)' 'OP-0130,OP-0140' ((@($rel.Atrasadas | Where-Object Orfa | ForEach-Object { $_.Oportunidade.Id }) | Sort-Object) -join ',')

    # --- Item 5a: dez produtos por faturamento no acumulado, com a filial da venda
    $porProd = @{}
    foreach ($x in $vendasBruto) {
        $id = $x.'ID Produto'
        if (-not $porProd.ContainsKey($id)) { $porProd[$id] = [pscustomobject]@{ Id = $id; Valor = [decimal]0; Un = 0; F = @{} } }
        $pp = $porProd[$id]; $v = Din $x.'Valor Total'
        $pp.Valor += $v; $pp.Un += [int]$x.Quantidade
        $fid = $filialId[$x.Filial]; $pp.F[$fid] = [decimal]$pp.F[$fid] + $v
    }
    $top = @($porProd.Values | Sort-Object -Property @{ Expression = { $_.Valor }; Descending = $true }, @{ Expression = { $_.Id } } | Select-Object -First 10)
    Confere 'item 5: dez produtos, na ordem' (($top | ForEach-Object Id) -join ',') (($rel.TopProdutos | ForEach-Object IdProduto) -join ',')
    $estBruto = Read-XlsxSheet (Join-Path $DirDados 'estoque_2026-08-31.xlsx') 'Estoque'
    for ($i = 0; $i -lt $top.Count; $i++) {
        $t = $top[$i]; $r = $rel.TopProdutos[$i]
        Confere "item 5 $($t.Id): faturamento" (Txt $t.Valor) (Txt $r.Valor)
        Confere "item 5 $($t.Id): unidades" $t.Un $r.Unidades
        foreach ($f in $rel.Filiais) { Confere "item 5 $($t.Id): faturamento em $($f.Nome)" (Txt ([decimal]$t.F[$f.Id])) (Txt $r.PorFilial[$f.Id]) }
        $q = 0; foreach ($l in $estBruto) { if ($l.'ID Produto' -eq $t.Id) { $q += [int]$l.Quantidade } }
        Confere "item 5 $($t.Id): estoque atual" $q $r.Estoque
    }

    # --- Item 5b: parados por filial (§5, 180 dias, foto de 31/08)
    foreach ($pf in $rel.ParadosPorFilial) {
        $ls = @($estBruto | Where-Object { $filialId[$_.Filial] -eq $pf.Filial.Id -and [int]$_.Quantidade -gt 0 } | Where-Object {
                $d = @(@($_.'Data da Última Entrada', $_.'Data da Última Saída') | Where-Object { $_ } | Sort-Object)
                (-not $d.Count) -or (([datetime]'2026-08-31' - [datetime]$d[-1]).Days -ge 180)
            })
        $v = [decimal]0; foreach ($l in $ls) { $v += [int]$l.Quantidade * (Din $l.'Custo Médio') }
        Confere "item 5 parados em $($pf.Filial.Nome)" ((($ls | ForEach-Object { $_.'ID Produto' }) | Sort-Object) -join ',') ((($pf.Parados | ForEach-Object { $_.Linha.Produto.Id }) | Sort-Object) -join ',')
        Confere "item 5 valor parado em $($pf.Filial.Nome)" (Txt $v) (Txt $pf.Valor)
    }

    # --- Catálogo do item 6: os números vêm do mesmo cálculo
    Confere 'catálogo: pipeline vencido' (Txt $rel.Funil.Total.ValorAtrasado) (Txt $ind['pipeline.atrasadas.valor'].Valor)
    Confere 'catálogo: atingimento da empresa' (Fr $rel.Ano.Empresa.Atingimento) (Fr $ind['empresa.ano.atingimento'].Valor)
    Confere 'catálogo: conversão (§6)' '49,7%' (Format-RelPct $ind['pipeline.conversao_qtd'].Valor)
    Confere 'catálogo: abaixo do mínimo (§6)' 28 $ind['estoque.abaixo_minimo.qtd'].Valor

    # ================================================================ item 6: validação do arquivo de riscos
    $imp = $rel.ImpressaoDigital
    $gravaRiscos = {
        param([string]$Nome, $Obj)
        $caminho = Join-Path $tmp $Nome
        [IO.File]::WriteAllText($caminho, ($Obj | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
        return $caminho
    }
    $r1 = @{ titulo = 'Risco A'; indicador = 'pipeline.atrasadas.valor'; leitura = 'Leitura A.' }
    $r2 = @{ titulo = 'Risco B'; numero = '12 dias'; calculo = 'Conta feita à mão.'; leitura = 'Leitura B.' }
    $r3 = @{ titulo = 'Risco C'; indicador = 'estoque.parado.valor'; leitura = 'Leitura C.' }
    $valido = & $gravaRiscos 'riscos_ok.json' @{ impressao_digital = $imp; riscos = @($r1, $r2, $r3) }
    $lido = Read-RiscosRelatorio -Caminho $valido -Indicadores $ind -ImpressaoDigital $imp
    Confere 'riscos: arquivo válido não tem erros' 0 $lido.Erros.Count
    Confere 'riscos: número do indicador vem do cálculo' 'R$ 6.813.700,00' (Format-RelIndicador $lido.Riscos[0].Indicador)
    $casos = @(
        @{ Nome = 'ausente'; Caminho = (Join-Path $tmp 'nao_existe.json'); Trecho = 'não encontrado' },
        @{ Nome = 'dois riscos'; Caminho = (& $gravaRiscos 'r2.json' @{ impressao_digital = $imp; riscos = @($r1, $r3) }); Trecho = 'exatamente 3' },
        @{ Nome = 'outros dados'; Caminho = (& $gravaRiscos 'r3.json' @{ impressao_digital = '0000000000000000'; riscos = @($r1, $r2, $r3) }); Trecho = 'outros dados' },
        @{ Nome = 'indicador inexistente'; Caminho = (& $gravaRiscos 'r4.json' @{ impressao_digital = $imp; riscos = @($r1, $r2, @{ titulo = 'X'; indicador = 'nao.existe'; leitura = 'L.' }) }); Trecho = 'não existe no catálogo' },
        @{ Nome = 'sem número'; Caminho = (& $gravaRiscos 'r5.json' @{ impressao_digital = $imp; riscos = @($r1, $r2, @{ titulo = 'X'; numero = '5%'; leitura = 'L.' }) }); Trecho = 'sem o número' },
        @{ Nome = 'sem leitura'; Caminho = (& $gravaRiscos 'r6.json' @{ impressao_digital = $imp; riscos = @($r1, $r2, @{ titulo = 'X'; indicador = 'estoque.parado.valor' }) }); Trecho = 'falta a leitura' }
    )
    foreach ($caso in $casos) {
        $l = Read-RiscosRelatorio -Caminho $caso.Caminho -Indicadores $ind -ImpressaoDigital $imp
        Confere "riscos recusados: $($caso.Nome)" 'True' ([bool](@($l.Erros | Where-Object { $_ -like "*$($caso.Trecho)*" }).Count))
    }

    # ================================================================ comando: repetibilidade e PDF
    $out1 = Join-Path $tmp 'out1'; $out2 = Join-Path $tmp 'out2'
    & $pwsh -NoProfile -File $app -DirDados $DirDados -SemImportacoes -Saida $out1 -Riscos $valido | Out-Null
    Confere 'comando gera o PDF (execução 1)' 0 $LASTEXITCODE
    & $pwsh -NoProfile -File $app -DirDados $DirDados -SemImportacoes -Saida $out2 -Riscos $valido | Out-Null
    Confere 'comando gera o PDF (execução 2)' 0 $LASTEXITCODE
    $pdf1 = Join-Path $out1 'relatorio_diretoria_2026-08-31.pdf'; $pdf2 = Join-Path $out2 'relatorio_diretoria_2026-08-31.pdf'
    Confere 'calculo.json idêntico nas duas execuções' (Get-FileHash (Join-Path $out1 'calculo.json')).Hash (Get-FileHash (Join-Path $out2 'calculo.json')).Hash
    Confere 'PDF idêntico nas duas execuções' (Get-FileHash $pdf1).Hash (Get-FileHash $pdf2).Hash
    $json = [IO.File]::ReadAllText((Join-Path $out1 'calculo.json')) | ConvertFrom-Json
    Confere 'calculo.json: impressão digital' $imp $json.impressao_digital
    Confere 'calculo.json: item 2' 'V009,V007,V002' (($json.item2_abaixo_80 | ForEach-Object id) -join ',')
    Confere 'calculo.json: pipeline ponderado' '5605495.00' (Txt $json.item3_pipeline.total.ponderado)

    # Estrutura do PDF: cabeçalho, fim, e cada entrada do xref apontando para o objeto certo.
    $texto = Get-TextoPdf $pdf1
    Confere 'PDF: cabeçalho' '%PDF-1.4' $texto.Substring(0, 8)
    Confere 'PDF: termina em %%EOF' 'True' ($texto.TrimEnd().EndsWith('%%EOF'))
    $inicioXref = [int]([regex]::Match($texto, 'startxref\n(\d+)').Groups[1].Value)
    Confere 'PDF: startxref aponta para a tabela' 'xref' $texto.Substring($inicioXref, 4)
    $entradas = [regex]::Matches($texto.Substring($inicioXref), '(\d{10}) 00000 n ')
    $okXref = $true; $n = 1
    foreach ($m in $entradas) { if (-not $texto.Substring([int]$m.Groups[1].Value).StartsWith("$n 0 obj")) { $okXref = $false }; $n++ }
    Confere 'PDF: deslocamentos do xref' 'True' $okXref
    foreach ($trecho in @('89,3%', 'R$ 13.870.500,00', 'R$ 5.605.495,00', 'R$ 6.813.700,00', 'R$ 1.693.610,48', 'Risco A', 'R$ 6.813.700,00', 'Número calculado na análise: Conta feita à mão.', "Impressão digital dos dados: $imp")) {
        Confere "PDF contém '$trecho'" 'True' ($texto.Contains((ConvertTo-PdfString $trecho)))
    }
    # Riscos de outros dados: não gera PDF (código 2).
    $out3 = Join-Path $tmp 'out3'
    & $pwsh -NoProfile -File $app -DirDados $DirDados -SemImportacoes -Saida $out3 -Riscos (Join-Path $tmp 'r3.json') | Out-Null
    Confere 'comando recusa riscos de outros dados (código 2)' 2 $LASTEXITCODE
    Confere 'comando recusado não grava PDF' 'False' (Test-Path (Join-Path $out3 'relatorio_diretoria_2026-08-31.pdf'))
    # Sem análise: o PDF sai e diz que o item 6 não foi incluído.
    $out4 = Join-Path $tmp 'out4'
    & $pwsh -NoProfile -File $app -DirDados $DirDados -SemImportacoes -Saida $out4 -SemAnalise | Out-Null
    Confere 'comando -SemAnalise gera o PDF' 0 $LASTEXITCODE
    Confere 'PDF -SemAnalise avisa que o item 6 ficou de fora' 'True' ((Get-TextoPdf (Join-Path $out4 'relatorio_diretoria_2026-08-31.pdf')).Contains((ConvertTo-PdfString 'Análise não incluída nesta emissão.')))

    # ================================================================ setembro importado (cópia temporária)
    $dados = Join-Path $tmp 'dados'
    New-Item -ItemType Directory -Path (Join-Path $dados 'importacoes\vendas'), (Join-Path $dados 'importacoes\estoque') -Force | Out-Null
    Get-ChildItem -LiteralPath $DirDados -File -Filter '*.xlsx' | Copy-Item -Destination $dados
    Copy-Item -LiteralPath (Join-Path $DirDados 'atualizacoes\vendas_2026_setembro.xlsx') -Destination (Join-Path $dados 'importacoes\vendas\20260923-120000__vendas_2026_setembro.xlsx')
    Copy-Item -LiteralPath (Join-Path $DirDados 'atualizacoes\estoque_2026-09-19.xlsx') -Destination (Join-Path $dados 'importacoes\estoque\estoque_2026-09-19__20260923-120000.xlsx')
    $dirImp = Get-DirImportacoes $dados
    $c2 = Import-BaseComercial -DirDados $dados -DirImportacoes $dirImp
    $p2 = Import-BasePipeline -DirDados $dados -Comercial $c2
    $e2 = Import-BaseEstoque -DirDados $dados -DirImportacoes $dirImp
    $rel2 = Get-RelatorioDiretoria -Comercial $c2 -Pipeline $p2 -Estoque $e2 -DirDados $dados
    Confere 'setembro: data-base de vendas (§0.3)' '2026-09-19' $rel2.DataBaseVendas.ToString('yyyy-MM-dd')
    Confere 'setembro: foto de estoque em vigor (§9.8)' '2026-09-19' $rel2.DataBaseEstoque.ToString('yyyy-MM-dd')
    Confere 'setembro: acumulado jan–set' 'jan–set/2026' $rel2.RotuloAno
    Confere 'setembro: mês parcial (§4.5)' 'True' $rel2.MesParcial
    Confere 'setembro: faturado no mês (importação documentada)' '3850268.28' (Txt $rel2.Mes.Empresa.RealizadoMesesMeta)
    Confere 'setembro: meta cheia do mês (§4.5)' '7190000.00' (Txt $rel2.Mes.Empresa.Meta)
    Confere 'setembro: impressão digital muda com os dados' 'True' ($rel2.ImpressaoDigital -ne $imp)
    $pdfSet = Join-Path $tmp 'set.pdf'
    New-PdfRelatorioDiretoria -Rel $rel2 -Riscos $null -Caminho $pdfSet
    $txtSet = Get-TextoPdf $pdfSet
    Confere 'setembro: PDF marca o mês como parcial' 'True' ($txtSet.Contains((ConvertTo-PdfString 'setembro/2026 (parcial até 19/09)')))
    Confere 'setembro: PDF avisa que o CRM é de outra data' 'True' ($txtSet.Contains((ConvertTo-PdfString 'as fotos não são da mesma data')))
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

$total = $script:ok + $script:falhas.Count
if ($script:falhas.Count) {
    Write-Host "$($script:falhas.Count) de $total conferências falharam:" -ForegroundColor Red
    $script:falhas | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
Write-Host "Todas as $total conferências do relatório da diretoria bateram." -ForegroundColor Green
