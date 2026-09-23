# Confere a importação de arquivos (REGRAS_NEGOCIO.md §9) numa CÓPIA temporária de dados/.
# As planilhas reais e a pasta dados/importacoes/ real nunca são tocadas.
# Os números esperados são contados direto nas planilhas brutas, por fora das regras.
# Uso: .\conferir_importacao.cmd   (ou: pwsh -NoProfile -File testes/conferir_importacao.ps1)

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\app\lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\Regras.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasEstoque.ps1')
. (Join-Path $PSScriptRoot '..\app\lib\RegrasImportacao.ps1')
. (Join-Path $PSScriptRoot 'XlsxTeste.ps1')
if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -eq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado '$Esperado', obtido '$Obtido'") }
}
function Din([string]$t) { [math]::Round([decimal]::Parse($t, [Globalization.NumberStyles]::Float, $inv), 2, [MidpointRounding]::AwayFromZero) }

# --- Cópia de trabalho: só as planilhas, sem importações anteriores ---
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte_imp_" + [Guid]::NewGuid().ToString('N'))
$dados = Join-Path $tmp 'dados'
New-Item -ItemType Directory -Path (Join-Path $dados 'atualizacoes') -Force | Out-Null
Get-ChildItem -LiteralPath $DirDados -File -Filter '*.xlsx' | Copy-Item -Destination $dados
Get-ChildItem -LiteralPath (Join-Path $DirDados 'atualizacoes') -File -Filter '*.xlsx' | Copy-Item -Destination (Join-Path $dados 'atualizacoes')
$dirImp = Get-DirImportacoes $dados
$arqSet = Join-Path $dados 'atualizacoes\vendas_2026_setembro.xlsx'
$arqEst19 = Join-Path $dados 'atualizacoes\estoque_2026-09-19.xlsx'
$arqEst31 = Join-Path $dados 'estoque_2026-08-31.xlsx'

function Get-Resumo([string]$Caminho, [string]$Tipo, [string]$Nome, [string]$Login = 'diretoria') {
    $token = New-ImportacaoPendente -DirDados $dados -Bytes ([IO.File]::ReadAllBytes($Caminho)) -NomeOriginal $Nome -Tipo $Tipo -Login $Login
    $p = Get-ImportacaoPendente $dados $token $Login
    $r = Get-ResumoImportacao -DirDados $dados -Caminho $p.Caminho -Tipo $Tipo -NomeOriginal $Nome
    return [pscustomobject]@{ Resumo = $r; Pendente = $p; Token = $token }
}
function Get-Gravados([string]$Tipo) { $p = Join-Path $dirImp $Tipo; if (Test-Path -LiteralPath $p) { @(Get-ChildItem -LiteralPath $p -File).Count } else { 0 } }
function Get-Sha([string]$Caminho) { (Get-FileHash -LiteralPath $Caminho -Algorithm SHA256).Hash }

try {
    # --- Fonte bruta ---
    $brutoSet = Read-XlsxSheet $arqSet 'Vendas'
    $cabVendas = @($brutoSet[0].PSObject.Properties.Name | Where-Object { $_ -ne '_Linha' })
    $fatSet = [decimal]0; foreach ($r in $brutoSet) { if ($r.Status -eq 'Faturada') { $fatSet += Din $r.'Valor Total' } }
    $brutoBase = Read-XlsxSheet (Join-Path $dados 'vendas_2026_jan-ago.xlsx') 'Vendas'

    # A. Sem importações, a carga é só a base
    $b0 = Import-BaseComercial -DirDados $dados -DirImportacoes $dirImp
    Confere 'antes: vendas em vigor = linhas da base' $brutoBase.Count $b0.Vendas.Count
    Confere 'antes: data-base 31/08' '2026-08-31' $b0.DataBase.ToString('yyyy-MM-dd')

    # B. Setembro: resumo, sem gravar
    $x = Get-Resumo $arqSet 'vendas' 'vendas_2026_setembro.xlsx'
    $r = $x.Resumo
    Confere 'set: aceito' $false $r.Recusado
    Confere 'set: novas = linhas do arquivo' $brutoSet.Count $r.Novas
    Confere 'set: atualizadas' 0 $r.Atualizadas.Count
    Confere 'set: ignoradas' 0 $r.Ignoradas
    Confere 'set: faturado nas novas' $fatSet.ToString('0.00', $inv) $r.FaturadoNovas.ToString('0.00', $inv)
    Confere 'set: faturado = R$ 3.850.268,28 (README da base)' '3850268.28' $r.FaturadoNovas.ToString('0.00', $inv)
    Confere 'set: canceladas novas' (@($brutoSet | Where-Object Status -eq 'Cancelada').Count) $r.NovasCanceladas
    Confere 'set: data-base depois 19/09' '2026-09-19' $r.DataBaseDepois.ToString('yyyy-MM-dd')
    Confere 'set: nada gravado antes de confirmar' 0 (Get-Gravados 'vendas')
    Confere 'set: vendas em vigor não mudaram antes de confirmar' $brutoBase.Count (Import-BaseComercial -DirDados $dados -DirImportacoes $dirImp).Vendas.Count

    # C. Confirmação grava o arquivo sem alteração e registra
    $c = Confirm-Importacao -DirDados $dados -Pendente $x.Pendente
    $gravado = Join-Path $dirImp "vendas\$($c.ArquivoGravado)"
    Confere 'set: gravado 1 arquivo' 1 (Get-Gravados 'vendas')
    Confere 'set: arquivo gravado idêntico ao enviado' (Get-Sha $arqSet) (Get-Sha $gravado)
    Confere 'set: pendência removida' 0 @(Get-ChildItem -LiteralPath (Join-Path $dirImp 'pendentes') -File).Count
    $b1 = Import-BaseComercial -DirDados $dados -DirImportacoes $dirImp
    Confere 'depois: vendas em vigor' ($brutoBase.Count + $brutoSet.Count) $b1.Vendas.Count
    Confere 'depois: data-base 19/09' '2026-09-19' $b1.DataBase.ToString('yyyy-MM-dd')
    $dSet = Get-Desempenho -Base $b1 -MesInicial 9 -MesFinal 9
    Confere 'depois: realizado de setembro' $fatSet.ToString('0.00', $inv) $dSet.Empresa.Realizado.ToString('0.00', $inv)
    Confere 'depois: setembro marcado como parcial (§4.5)' $true $dSet.Parcial
    Confere 'depois: jan–ago não muda (§6)' '45295119.16' (Get-Desempenho -Base $b1 -MesInicial 1 -MesFinal 8).Empresa.Realizado.ToString('0.00', $inv)
    Confere 'depois: carga sem importações continua só a base' $brutoBase.Count (Import-BaseComercial -DirDados $dados).Vendas.Count

    # D. Reimportar o mesmo arquivo não duplica
    $r = (Get-Resumo $arqSet 'vendas' 'vendas_2026_setembro.xlsx').Resumo
    Confere 'reimportação: novas' 0 $r.Novas
    Confere 'reimportação: ignoradas = todas' $brutoSet.Count $r.Ignoradas
    Confere 'reimportação: atualizadas' 0 $r.Atualizadas.Count
    Confere 'reimportação: total em vigor não muda' $b1.Vendas.Count $r.VendasDepois

    # E. Arquivo com 1 venda alterada (Faturada -> Cancelada) e 1 venda nova
    $alvo = @($brutoSet | Where-Object Status -eq 'Faturada')[0]
    $linhasE = foreach ($l in $brutoSet) {
        $h = [ordered]@{}; foreach ($col in $cabVendas) { $h[$col] = $l.$col }
        if ($l.'ID Venda' -eq $alvo.'ID Venda') { $h.Status = 'Cancelada' }
        [pscustomobject]$h
    }
    $nova = [ordered]@{}; foreach ($col in $cabVendas) { $nova[$col] = $alvo.$col }; $nova.'ID Venda' = 'VD-2026-09999'
    $arqE = Join-Path $tmp 'correcao_setembro.xlsx'
    New-XlsxTeste -Caminho $arqE -Aba 'Vendas' -Cabecalho $cabVendas -Linhas (@($linhasE) + [pscustomobject]$nova)
    $x = Get-Resumo $arqE 'vendas' 'correção setembro.xlsx'
    $r = $x.Resumo
    Confere 'correção: novas' 1 $r.Novas
    Confere 'correção: atualizadas' 1 $r.Atualizadas.Count
    Confere 'correção: ignoradas' ($brutoSet.Count - 1) $r.Ignoradas
    Confere 'correção: o que mudou' "$($alvo.'ID Venda') Status Faturada->Cancelada" ("{0} {1} {2}->{3}" -f $r.Atualizadas[0].Id, $r.Atualizadas[0].Mudancas[0].Coluna, $r.Atualizadas[0].Mudancas[0].De, $r.Atualizadas[0].Mudancas[0].Para)
    $c = Confirm-Importacao -DirDados $dados -Pendente $x.Pendente
    Confere 'correção: nome gravado sem acento nem espaço' $true ($c.ArquivoGravado -match '^\d{8}-\d{6}__correcao_setembro\.xlsx$')
    $b2 = Import-BaseComercial -DirDados $dados -DirImportacoes $dirImp
    $v = $b2.Vendas | Where-Object Id -eq $alvo.'ID Venda'
    Confere 'correção: a versão mais nova vale (upsert)' 'Cancelada' $v.Status
    Confere 'correção: a venda aponta o arquivo de onde veio' $c.ArquivoGravado $v.Arquivo
    Confere 'correção: total em vigor' ($b1.Vendas.Count + 1) $b2.Vendas.Count
    $esperado = $fatSet - (Din $alvo.'Valor Total') + (Din $alvo.'Valor Total')   # saiu a alterada, entrou a nova (mesmo valor)
    Confere 'correção: realizado de setembro' $esperado.ToString('0.00', $inv) (Get-Desempenho -Base $b2 -MesInicial 9 -MesFinal 9).Empresa.Realizado.ToString('0.00', $inv)

    # F. Recusas de vendas: nada é gravado
    $antesRecusas = Get-Gravados 'vendas'
    $semDuas = @($cabVendas | Where-Object { $_ -notin @('Status', 'Forma de Pagamento') })
    $arqF = Join-Path $tmp 'sem_colunas.xlsx'; New-XlsxTeste -Caminho $arqF -Aba 'Vendas' -Cabecalho $semDuas -Linhas @($brutoSet | Select-Object -First 3)
    $r = (Get-Resumo $arqF 'vendas' 'sem_colunas.xlsx').Resumo
    Confere 'falta 2 colunas: recusado' $true $r.Recusado
    Confere 'falta 2 colunas: diz quais' $true ($r.Erros[0] -match "Faltam 2 colunas obrigatórias: 'Forma de Pagamento', 'Status'")
    $arqF1 = Join-Path $tmp 'sem_id.xlsx'; New-XlsxTeste -Caminho $arqF1 -Aba 'Vendas' -Cabecalho @($cabVendas | Where-Object { $_ -ne 'ID Venda' }) -Linhas @($brutoSet | Select-Object -First 3)
    $r = (Get-Resumo $arqF1 'vendas' 'sem_id.xlsx').Resumo
    Confere 'falta 1 coluna: diz qual' $true ($r.Recusado -and $r.Erros[0] -match "Falta a coluna obrigatória: 'ID Venda'")
    $arqG = Join-Path $tmp 'duplicado.xlsx'; New-XlsxTeste -Caminho $arqG -Aba 'Vendas' -Cabecalho $cabVendas -Linhas @($brutoSet[0], $brutoSet[1], $brutoSet[0])
    $r = (Get-Resumo $arqG 'vendas' 'duplicado.xlsx').Resumo
    Confere 'ID repetido no arquivo: recusado com o nome do arquivo' $true ($r.Recusado -and ($r.Erros -join ' ') -match "duplicado\.xlsx, linha 4: ID Venda $($brutoSet[0].'ID Venda') duplicado")
    $desl = [ordered]@{}; foreach ($col in $cabVendas) { $desl[$col] = $brutoSet[0].$col }
    $desl.'ID Venda' = 'VD-2026-09998'; $desl.'ID Vendedor' = 'V011'; $desl.Vendedor = 'Ricardo Alves Pereira'; $desl.Filial = 'Chapecó'
    $arqH = Join-Path $tmp 'desligado.xlsx'; New-XlsxTeste -Caminho $arqH -Aba 'Vendas' -Cabecalho $cabVendas -Linhas @([pscustomobject]$desl)
    $r = (Get-Resumo $arqH 'vendas' 'desligado.xlsx').Resumo
    Confere 'venda de desligado depois do desligamento: recusado (§2.6)' $true ($r.Recusado -and ($r.Erros -join ' ') -match 'vendedor desligado em 30/04/2026')
    $arqI = Join-Path $tmp 'aba_errada.xlsx'; New-XlsxTeste -Caminho $arqI -Aba 'Planilha1' -Cabecalho $cabVendas -Linhas @($brutoSet[0])
    $r = (Get-Resumo $arqI 'vendas' 'aba_errada.xlsx').Resumo
    Confere 'aba errada: recusado' $true ($r.Recusado -and $r.Erros[0] -match "não tem a aba 'Vendas'")
    $arqJ = Join-Path $tmp 'nao_e_excel.xlsx'; [IO.File]::WriteAllText($arqJ, 'isto não é uma planilha')
    $r = (Get-Resumo $arqJ 'vendas' 'nao_e_excel.xlsx').Resumo
    Confere 'arquivo que não é Excel: recusado' $true ($r.Recusado -and $r.Erros[0] -match 'não pôde ser lido')
    $r = (Get-Resumo $arqEst19 'vendas' 'estoque_2026-09-19.xlsx').Resumo
    Confere 'estoque enviado como vendas: recusado' $true $r.Recusado
    Confere 'recusas de vendas não gravam nada' $antesRecusas (Get-Gravados 'vendas')

    # G. Estoque 19/09: foto mais nova substitui a inteira
    $bruto31 = Read-XlsxSheet $arqEst31 'Estoque'
    $bruto19 = Read-XlsxSheet $arqEst19 'Estoque'
    $q31 = @{}; foreach ($l in $bruto31) { $q31["$($l.'ID Produto')|$($l.Filial)"] = [int]$l.Quantidade }
    $mudam = @($bruto19 | Where-Object { $q31["$($_.'ID Produto')|$($_.Filial)"] -ne [int]$_.Quantidade }).Count
    $semEst19 = @($bruto19 | Group-Object 'ID Produto' | Where-Object { ($_.Group | ForEach-Object { [int]$_.Quantidade } | Measure-Object -Sum).Sum -eq 0 } | ForEach-Object Name | Sort-Object)
    $x = Get-Resumo $arqEst19 'estoque' 'estoque_2026-09-19.xlsx'
    $r = $x.Resumo
    Confere 'est 19/09: aceito' $false $r.Recusado
    Confere 'est 19/09: foto em vigor 31/08' '2026-08-31' $r.DataAtual.ToString('yyyy-MM-dd')
    Confere 'est 19/09: é mais nova (não é a mesma data)' $false $r.MesmaData
    Confere 'est 19/09: linhas substituídas' $bruto31.Count $r.LinhasSubstituidas
    Confere 'est 19/09: linhas da foto nova' $bruto19.Count $r.LinhasNova
    Confere 'est 19/09: quantidades que mudam' $mudam $r.QtdMudou
    Confere 'est 19/09: quantidades que mudam = 61 (PERFIL.md)' 61 $r.QtdMudou
    Confere 'est 19/09: sem estoque depois' ($semEst19 -join ',') (($r.SemEstoqueDepois | Sort-Object) -join ',')
    Confere 'est 19/09: P003 sai e P030/P034/P042 entram (PERFIL.md)' 'True True True True' ("{0} {1} {2} {3}" -f ('P003' -notin $r.SemEstoqueDepois), ('P030' -in $r.SemEstoqueDepois), ('P034' -in $r.SemEstoqueDepois), ('P042' -in $r.SemEstoqueDepois))
    Confere 'est 19/09: nada gravado antes de confirmar' 0 (Get-Gravados 'estoque')
    $c = Confirm-Importacao -DirDados $dados -Pendente $x.Pendente
    Confere 'est 19/09: nome gravado' $true ($c.ArquivoGravado -match '^estoque_2026-09-19__\d{8}-\d{6}\.xlsx$')
    Confere 'est 19/09: arquivo gravado idêntico ao enviado' (Get-Sha $arqEst19) (Get-Sha (Join-Path $dirImp "estoque\$($c.ArquivoGravado)"))
    Confere 'est 19/09: foto em vigor passa a 19/09' '2026-09-19' (Import-BaseEstoque -DirDados $dados -DirImportacoes $dirImp).DataBase.ToString('yyyy-MM-dd')
    Confere 'est 19/09: carga sem importações continua em 31/08 (§6)' '2026-08-31' (Import-BaseEstoque -DirDados $dados).DataBase.ToString('yyyy-MM-dd')

    # H. Foto mais antiga é recusada; mesma data substitui a daquela data
    $r = (Get-Resumo $arqEst31 'estoque' 'estoque_2026-08-31.xlsx').Resumo
    Confere 'est 31/08 depois de 19/09: recusada por ser mais antiga' $true ($r.Recusado -and $r.Erros[0] -match 'mais antiga que a foto em vigor, de 19/09/2026')
    $x = Get-Resumo $arqEst19 'estoque' 'estoque_2026-09-19.xlsx'
    Confere 'est 19/09 de novo: mesma data substitui' $true ((-not $x.Resumo.Recusado) -and $x.Resumo.MesmaData)
    Confere 'est 19/09 de novo: nenhuma quantidade muda' 0 $x.Resumo.QtdMudou
    $c2 = Confirm-Importacao -DirDados $dados -Pendente $x.Pendente
    Confere 'est 19/09 de novo: a importação mais recente é a foto em vigor' $c2.ArquivoGravado (Get-ArquivoEstoque $dados $dirImp).Arquivo.Name

    # I. Recusas de estoque
    $antesEst = Get-Gravados 'estoque'
    $r = (Get-Resumo $arqEst19 'estoque' 'posicao_setembro.xlsx').Resumo
    Confere 'est com nome fora do padrão: recusado' $true ($r.Recusado -and $r.Erros[0] -match 'estoque_AAAA-MM-DD\.xlsx')
    $cabEst = @($bruto19[0].PSObject.Properties.Name | Where-Object { $_ -ne '_Linha' })
    $arqK = Join-Path $tmp 'estoque_2026-09-30.xlsx'; New-XlsxTeste -Caminho $arqK -Aba 'Estoque' -Cabecalho @($cabEst | Where-Object { $_ -ne 'Custo Médio' }) -Linhas @($bruto19 | Select-Object -First 5)
    $r = (Get-Resumo $arqK 'estoque' 'estoque_2026-09-30.xlsx').Resumo
    Confere 'est sem Custo Médio: diz qual coluna' $true ($r.Recusado -and $r.Erros[0] -match "Falta a coluna obrigatória: 'Custo Médio'")
    $ruim = [ordered]@{}; foreach ($col in $cabEst) { $ruim[$col] = $bruto19[0].$col }; $ruim.Filial = 'Curitiba'
    $arqL = Join-Path $tmp 'estoque_2026-09-30b.xlsx'; New-XlsxTeste -Caminho $arqL -Aba 'Estoque' -Cabecalho $cabEst -Linhas @([pscustomobject]$ruim)
    $r = (Get-Resumo $arqL 'estoque' 'estoque_2026-09-30.xlsx').Resumo
    Confere 'est com filial desconhecida: recusado (§0.2)' $true ($r.Recusado -and ($r.Erros -join ' ') -match "estoque_2026-09-30\.xlsx, linha 2: filial desconhecida 'Curitiba'")
    Confere 'recusas de estoque não gravam nada' $antesEst (Get-Gravados 'estoque')

    # J. Pendências: só quem enviou confirma; token inválido não vale
    $x = Get-Resumo $arqSet 'vendas' 'vendas_2026_setembro.xlsx' 'diretoria'
    Confere 'pendência de outro login não vale' $true ($null -eq (Get-ImportacaoPendente $dados $x.Token 'outra.pessoa'))
    Confere 'token com caminho não vale' $true ($null -eq (Get-ImportacaoPendente $dados '..\..\vendas' 'diretoria'))
    Remove-ImportacaoPendente $x.Pendente

    # K. Registro
    $hist = Get-HistoricoImportacoes $dados
    Confere 'registro: 4 importações confirmadas' 4 $hist.Count
    Confere 'registro: a mais recente primeiro' 'estoque' $hist[0].Tipo
    Confere 'registro: contagens da correção' '1/1/81' ('{0}/{1}/{2}' -f $hist[2].Novas, $hist[2].Atualizadas, $hist[2].Ignoradas)
} finally {
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}

$total = $script:ok + $script:falhas.Count
if ($script:falhas.Count) {
    Write-Host "$($script:falhas.Count) de $total conferências da importação falharam:" -ForegroundColor Red
    $script:falhas | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
Write-Host "Todas as $total conferências da importação bateram." -ForegroundColor Green
Write-Host ("Setembro: {0} vendas novas, R$ {1:N2} faturados; reimportação sem duplicar; estoque 19/09 substitui a foto de 31/08." -f $brutoSet.Count, $fatSet)
