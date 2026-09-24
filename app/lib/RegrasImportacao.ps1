# Regras da importação de arquivos — implementação do REGRAS_NEGOCIO.md §9.
# Único lugar que decide se um arquivo é aceito, o que ele muda (resumo) e como é gravado.
# Depende de Xlsx.ps1, Regras.ps1 (carga de vendas) e RegrasEstoque.ps1 (carga de estoque).
# Nada aqui altera as planilhas originais de dados/ (§9.7).

# §9.3: colunas obrigatórias, com o nome exato do cabeçalho.
$script:ColunasImportacao = @{
    vendas  = @('ID Venda', 'Data', 'ID Vendedor', 'Vendedor', 'Filial', 'ID Cliente', 'Cliente', 'Cidade', 'UF',
        'ID Produto', 'Produto', 'Categoria', 'Quantidade', 'Valor Unitário', 'Valor Total', 'Forma de Pagamento',
        'ID Oportunidade', 'Status')
    estoque = @('ID Produto', 'Produto', 'Categoria', 'Filial', 'Quantidade', 'Estoque Mínimo', 'Custo Médio',
        'Data da Última Entrada', 'Data da Última Saída')
}
$script:AbaImportacao = @{ vendas = 'Vendas'; estoque = 'Estoque' }
$script:ColunasDinheiroVenda = @('Valor Unitário', 'Valor Total')
$script:PadraoToken = '^[0-9a-f]{32}$'
$script:ValidadePendente = [timespan]::FromHours(2)

function Get-DirImportacoes([string]$DirDados) { return (Join-Path $DirDados 'importacoes') }

function Get-NomeSeguro([string]$Nome) {
    # Nome original vira parte do nome gravado: só letras, números, ponto, hífen e sublinhado.
    $n = [IO.Path]::GetFileNameWithoutExtension($Nome).Normalize([Text.NormalizationForm]::FormD)
    $n = ($n.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' }) -join ''
    $n = ($n -replace '[^A-Za-z0-9._-]', '_').Trim('_', '.')
    if (-not $n) { $n = 'arquivo' }
    if ($n.Length -gt 80) { $n = $n.Substring(0, 80) }
    return "$n.xlsx"
}

function Test-ArquivoImportacao {
    # §9.2 e §9.3: a aba certa existe, tem linhas e tem todas as colunas obrigatórias.
    param([Parameter(Mandatory)][string]$Caminho, [Parameter(Mandatory)][string]$Tipo)
    $aba = $script:AbaImportacao[$Tipo]
    try {
        $linhas = Read-XlsxSheet $Caminho $aba
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'não encontrada') { return @("A planilha não tem a aba '$aba'. Um arquivo de $Tipo precisa ter a aba '$aba'.") }
        return @("O arquivo não pôde ser lido como planilha Excel (.xlsx): $msg")
    }
    if (-not $linhas.Count) { return @("A aba '$aba' não tem nenhuma linha de dados.") }
    $cabecalho = @($linhas[0].PSObject.Properties.Name | Where-Object { $_ -ne '_Linha' })
    $faltam = @($script:ColunasImportacao[$Tipo] | Where-Object { $_ -notin $cabecalho })
    if ($faltam.Count) {
        $lista = ($faltam | ForEach-Object { "'$_'" }) -join ', '
        $rotulo = if ($faltam.Count -eq 1) { 'Falta a coluna obrigatória' } else { "Faltam $($faltam.Count) colunas obrigatórias" }
        return @("${rotulo}: $lista. O nome do cabeçalho precisa ser exatamente este (REGRAS_NEGOCIO.md §9.3).")
    }
    return @()
}

function Get-ValorComparavel([string]$Coluna, $Valor) {
    # §9.5: comparação de "ignorada": dinheiro a 2 casas, datas como data, números como número.
    $t = if ($null -eq $Valor) { '' } else { ([string]$Valor).Trim() }
    if ($t -eq '') { return '' }
    if ($Coluna -in $script:ColunasDinheiroVenda) { $d = ConvertTo-Dinheiro $t; if ($null -ne $d) { return $d.ToString('0.00', $script:Inv) } }
    if ($Coluna -eq 'Data') { $d = ConvertTo-Data $t; if ($d) { return $d.ToString('yyyy-MM-dd') } }
    if ($Coluna -eq 'Quantidade') { $d = [decimal]0; if ([decimal]::TryParse($t, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d.ToString($script:Inv) } }
    return $t
}

function Get-VendasBrutasVigentes([string]$DirDados, [string]$DirImportacoes) {
    # Linhas brutas em vigor por ID Venda: a base e depois as importações, na ordem (§9.8).
    $porId = @{}
    $fontes = @(Join-Path $DirDados $script:ArquivosBase.Vendas) + @(Get-ArquivosVendasImportados $DirImportacoes | ForEach-Object FullName)
    foreach ($f in $fontes) { foreach ($r in (Read-XlsxSheet $f 'Vendas')) { if ($r.'ID Venda') { $porId[$r.'ID Venda'] = $r } } }
    return $porId
}

function New-Recusa([string]$Tipo, [string]$NomeOriginal, [string[]]$Erros) {
    return [pscustomobject]@{ Tipo = $Tipo; NomeOriginal = $NomeOriginal; Recusado = $true; Erros = @($Erros) }
}

function Get-ResumoImportacao {
    # §9.4 e §9.5: valida tudo e diz o que vai acontecer. Não grava nada.
    param(
        [Parameter(Mandatory)][string]$DirDados,
        [Parameter(Mandatory)][string]$Caminho,
        [Parameter(Mandatory)][ValidateSet('vendas', 'estoque')][string]$Tipo,
        [Parameter(Mandatory)][string]$NomeOriginal
    )
    $dirImp = Get-DirImportacoes $DirDados
    $nomeTemp = Split-Path $Caminho -Leaf
    $trocaNome = { param($erros) @($erros | ForEach-Object { ([string]$_).Replace($nomeTemp, $NomeOriginal) }) }

    if ($Tipo -eq 'estoque' -and $NomeOriginal -notmatch $script:PadraoArquivoEstoque) {
        return New-Recusa $Tipo $NomeOriginal @("O nome do arquivo de estoque precisa ser estoque_AAAA-MM-DD.xlsx (ex.: estoque_2026-09-19.xlsx): a data da foto vem do nome (REGRAS_NEGOCIO.md §5 e §9.6). Nome recebido: '$NomeOriginal'.")
    }
    $problemas = Test-ArquivoImportacao -Caminho $Caminho -Tipo $Tipo
    if ($problemas.Count) { return New-Recusa $Tipo $NomeOriginal $problemas }

    if ($Tipo -eq 'vendas') {
        $depois = Import-BaseComercial -DirDados $DirDados -DirImportacoes $dirImp -VendasAdicionais $Caminho
        if ($depois.Erros.Count) { return New-Recusa $Tipo $NomeOriginal (& $trocaNome $depois.Erros) }
        $antes = Import-BaseComercial -DirDados $DirDados -DirImportacoes $dirImp
        $vigentes = Get-VendasBrutasVigentes $DirDados $dirImp
        $novas = New-Object Collections.Generic.List[object]
        $atualizadas = New-Object Collections.Generic.List[object]
        $ignoradas = New-Object Collections.Generic.List[object]
        $linhas = Read-XlsxSheet $Caminho 'Vendas'
        foreach ($r in $linhas) {
            $id = $r.'ID Venda'
            if (-not $vigentes.ContainsKey($id)) { $novas.Add($r); continue }
            $atual = $vigentes[$id]
            $mudancas = foreach ($c in $script:ColunasImportacao.vendas) {
                $a = Get-ValorComparavel $c $atual.$c; $b = Get-ValorComparavel $c $r.$c
                if ($a -ne $b) { [pscustomobject]@{ Coluna = $c; De = $atual.$c; Para = $r.$c } }
            }
            if ($mudancas) { $atualizadas.Add([pscustomobject]@{ Id = $id; Linha = $r._Linha; Mudancas = @($mudancas) }) } else { $ignoradas.Add($r) }
        }
        $datas = @($linhas | ForEach-Object { ConvertTo-Data $_.Data } | Where-Object { $_ } | Sort-Object)
        $faturadoNovas = [decimal]0
        foreach ($r in $novas) { if ($r.Status -eq 'Faturada') { $faturadoNovas += ConvertTo-Dinheiro $r.'Valor Total' } }
        return [pscustomobject]@{
            Tipo            = 'vendas'
            NomeOriginal    = $NomeOriginal
            Recusado        = $false
            Erros           = @()
            LinhasArquivo   = $linhas.Count
            Novas           = $novas.Count
            Atualizadas     = $atualizadas.ToArray()
            Ignoradas       = $ignoradas.Count
            PrimeiraData    = $datas | Select-Object -First 1
            UltimaData      = $datas | Select-Object -Last 1
            FaturadoNovas   = $faturadoNovas
            NovasCanceladas = @($novas.ToArray() | Where-Object { $_.Status -eq 'Cancelada' }).Count
            VendasAntes     = $antes.Vendas.Count
            VendasDepois    = $depois.Vendas.Count
            DataBaseAntes   = $antes.DataBase
            DataBaseDepois  = $depois.DataBase
        }
    }

    # Estoque (§5, §9.6)
    [void]($NomeOriginal -match $script:PadraoArquivoEstoque)
    $dataNova = ConvertTo-Data $Matches[1]
    if (-not $dataNova) { return New-Recusa $Tipo $NomeOriginal @("A data no nome '$NomeOriginal' não é uma data válida.") }
    $emVigor = Get-ArquivoEstoque $DirDados $dirImp
    if ($emVigor -and $dataNova -lt $emVigor.Data) {
        return New-Recusa $Tipo $NomeOriginal @("A foto de $($dataNova.ToString('dd/MM/yyyy')) é mais antiga que a foto em vigor, de $($emVigor.Data.ToString('dd/MM/yyyy')). Uma foto mais antiga não pode substituir uma mais nova (REGRAS_NEGOCIO.md §5 e §9.6).")
    }
    $nova = Import-BaseEstoque -DirDados $DirDados -DirImportacoes $dirImp -ArquivoFoto $Caminho -DataFoto $dataNova   # catálogo em vigor (§11.4)
    if ($nova.Erros.Count) { return New-Recusa $Tipo $NomeOriginal (& $trocaNome $nova.Erros) }
    $atual = Import-BaseEstoque -DirDados $DirDados -DirImportacoes $dirImp
    $chave = { param($l) "$($l.Produto.Id)|$($l.IdFilial)" }
    $mapaAtual = @{}; foreach ($l in $atual.Linhas) { $mapaAtual[(& $chave $l)] = $l }
    $mapaNova = @{}; foreach ($l in $nova.Linhas) { $mapaNova[(& $chave $l)] = $l }
    $qtdMudou = 0; $soNaNova = 0
    foreach ($k in $mapaNova.Keys) { if (-not $mapaAtual.ContainsKey($k)) { $soNaNova++ } elseif ($mapaAtual[$k].Quantidade -ne $mapaNova[$k].Quantidade) { $qtdMudou++ } }
    $somem = @($mapaAtual.Keys | Where-Object { -not $mapaNova.ContainsKey($_) }).Count
    $visAtual = Get-VisaoEstoque -Estoque $atual -DiasParado $script:DiasParadoPadrao -AbertasPorProduto $null
    $visNova = Get-VisaoEstoque -Estoque $nova -DiasParado $script:DiasParadoPadrao -AbertasPorProduto $null
    return [pscustomobject]@{
        Tipo              = 'estoque'
        NomeOriginal      = $NomeOriginal
        Recusado          = $false
        Erros             = @()
        DataNova          = $dataNova
        DataAtual         = $atual.DataBase
        ArquivoAtual      = $atual.Arquivo.Name
        MesmaData         = ($dataNova -eq $atual.DataBase)
        LinhasSubstituidas = $atual.Linhas.Count
        LinhasNova        = $nova.Linhas.Count
        QtdMudou          = $qtdMudou
        LinhasNovasChave  = $soNaNova
        LinhasSomem       = $somem
        ImobilizadoAntes  = $visAtual.Total.Valor
        ImobilizadoDepois = $visNova.Total.Valor
        SemEstoqueAntes   = @($visAtual.SemEstoque | ForEach-Object { $_.Produto.Id })
        SemEstoqueDepois  = @($visNova.SemEstoque | ForEach-Object { $_.Produto.Id })
        AbaixoAntes       = $visAtual.Abaixo.Count
        AbaixoDepois      = $visNova.Abaixo.Count
    }
}

function New-ImportacaoPendente {
    # Guarda o arquivo enviado até a confirmação (§9.5). Devolve o token da pendência.
    param([Parameter(Mandatory)][string]$DirDados, [Parameter(Mandatory)][byte[]]$Bytes, [Parameter(Mandatory)][string]$NomeOriginal, [Parameter(Mandatory)][string]$Tipo, [Parameter(Mandatory)][string]$Login)
    $pasta = Join-Path (Get-DirImportacoes $DirDados) 'pendentes'
    [void](New-Item -ItemType Directory -Force -Path $pasta)
    # Pendências abandonadas (não confirmadas nem canceladas) saem depois de 24 h.
    Get-ChildItem -LiteralPath $pasta -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-24) } | Remove-Item -Force -ErrorAction SilentlyContinue
    $rng =[Security.Cryptography.RandomNumberGenerator]::Create(); $b = New-Object byte[] 16; $rng.GetBytes($b); $rng.Dispose()
    $token = -join ($b | ForEach-Object { $_.ToString('x2') })
    [IO.File]::WriteAllBytes((Join-Path $pasta "$token.xlsx"), $Bytes)
    $meta = [ordered]@{ tipo = $Tipo; nome = $NomeOriginal; login = $Login; criadoEm = (Get-Date).ToString('o') }
    [IO.File]::WriteAllText((Join-Path $pasta "$token.json"), (ConvertTo-Json $meta), (New-Object Text.UTF8Encoding($false)))
    return $token
}

function Get-ImportacaoPendente([string]$DirDados, [string]$Token, [string]$Login) {
    # Só quem enviou confirma, e só dentro da validade.
    if ($Token -notmatch $script:PadraoToken) { return $null }
    $pasta = Join-Path (Get-DirImportacoes $DirDados) 'pendentes'
    $arqMeta = Join-Path $pasta "$Token.json"; $arq = Join-Path $pasta "$Token.xlsx"
    if (-not (Test-Path -LiteralPath $arqMeta) -or -not (Test-Path -LiteralPath $arq)) { return $null }
    $meta = Get-Content -LiteralPath $arqMeta -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($meta.login -ne $Login) { return $null }
    # O PowerShell 7 já devolve a data convertida pelo ConvertFrom-Json; o 5.1 devolve o texto ISO.
    $criado = if ($meta.criadoEm -is [datetime]) { $meta.criadoEm } else { [datetime]::Parse([string]$meta.criadoEm, $script:Inv, [Globalization.DateTimeStyles]::RoundtripKind) }
    if (((Get-Date) - $criado) -gt $script:ValidadePendente) { return $null }
    return [pscustomobject]@{ Token = $Token; Tipo = $meta.tipo; NomeOriginal = $meta.nome; Login = $meta.login; Caminho = $arq; Meta = $arqMeta }
}

function Remove-ImportacaoPendente($Pendente) {
    if (-not $Pendente) { return }
    Remove-Item -LiteralPath $Pendente.Caminho, $Pendente.Meta -Force -ErrorAction SilentlyContinue
}

function Confirm-Importacao {
    # §9.5 e §9.7: refaz a validação e, se passar, grava o arquivo sem alteração e registra.
    param([Parameter(Mandatory)][string]$DirDados, [Parameter(Mandatory)]$Pendente)
    $resumo = Get-ResumoImportacao -DirDados $DirDados -Caminho $Pendente.Caminho -Tipo $Pendente.Tipo -NomeOriginal $Pendente.NomeOriginal
    if ($resumo.Recusado) { return $resumo }
    $dirImp = Get-DirImportacoes $DirDados
    $pasta = Join-Path $dirImp $Pendente.Tipo
    [void](New-Item -ItemType Directory -Force -Path $pasta)
    $agora = Get-Date
    do {
        $carimbo = $agora.ToString('yyyyMMdd-HHmmss')
        $nome = if ($Pendente.Tipo -eq 'vendas') { "${carimbo}__$(Get-NomeSeguro $Pendente.NomeOriginal)" } else { "estoque_$($resumo.DataNova.ToString('yyyy-MM-dd'))__$carimbo.xlsx" }
        $destino = Join-Path $pasta $nome
        $agora = $agora.AddSeconds(1)
    } while (Test-Path -LiteralPath $destino)
    Move-Item -LiteralPath $Pendente.Caminho -Destination $destino
    Remove-Item -LiteralPath $Pendente.Meta -Force -ErrorAction SilentlyContinue

    $registro = Join-Path $dirImp 'registro.csv'
    if (-not (Test-Path -LiteralPath $registro)) {
        [IO.File]::WriteAllText($registro, "Quando;Login;Tipo;Arquivo enviado;Arquivo gravado;Novas;Atualizadas;Ignoradas;Data da foto;Linhas substituídas`r`n", (New-Object Text.UTF8Encoding($true)))
    }
    $limpa = { param($t) ([string]$t) -replace '[;\r\n]', ' ' }
    $campos = if ($resumo.Tipo -eq 'vendas') { @($resumo.Novas, $resumo.Atualizadas.Count, $resumo.Ignoradas, '', '') } else { @('', '', '', $resumo.DataNova.ToString('dd/MM/yyyy'), $resumo.LinhasSubstituidas) }
    $linha = (@((Get-Date).ToString('dd/MM/yyyy HH:mm:ss'), (& $limpa $Pendente.Login), $resumo.Tipo, (& $limpa $Pendente.NomeOriginal), $nome) + $campos) -join ';'
    [IO.File]::AppendAllText($registro, "$linha`r`n", (New-Object Text.UTF8Encoding($false)))
    $resumo | Add-Member ArquivoGravado $nome
    return $resumo
}

function Get-HistoricoImportacoes([string]$DirDados) {
    $registro = Join-Path (Get-DirImportacoes $DirDados) 'registro.csv'
    if (-not (Test-Path -LiteralPath $registro)) { return @() }
    $linhas = @(Get-Content -LiteralPath $registro -Encoding UTF8 | Select-Object -Skip 1 | Where-Object { $_ })
    [array]::Reverse($linhas)
    return @($linhas | ForEach-Object {
            $c = $_ -split ';'
            [pscustomobject]@{ Quando = $c[0]; Login = $c[1]; Tipo = $c[2]; Enviado = $c[3]; Gravado = $c[4]; Novas = $c[5]; Atualizadas = $c[6]; Ignoradas = $c[7]; DataFoto = $c[8]; Substituidas = $c[9] }
        })
}
