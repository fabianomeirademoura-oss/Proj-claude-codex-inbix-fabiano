# Regras de estoque — implementação do REGRAS_NEGOCIO.md §5.
# Único lugar que decide valor imobilizado, parado, sem estoque e abaixo do mínimo.
# Estoque é sempre por linha (produto, filial); o produto no agregado só entra em "sem estoque".
# Depende de Regras.ps1 (conversões).

$script:DiasParadoPadrao = 180   # REGRAS_NEGOCIO.md §5: prazo padrão; a tela permite outro valor.
$script:DiasParadoMaximo = 3650
$script:PadraoArquivoEstoque = '^estoque_(\d{4}-\d{2}-\d{2})\.xlsx$'
$script:PadraoEstoqueImportado = '^estoque_(\d{4}-\d{2}-\d{2})__(\d{8}-\d{6})\.xlsx$'   # REGRAS_NEGOCIO.md §9.7
$script:ArquivosEstoqueFixos = @('produtos.xlsx', 'vendedores.xlsx')

function Get-FotosEstoque([string]$DirDados, [string]$DirImportacoes) {
    # §5: a foto vale na data do nome do arquivo (estoque_AAAA-MM-DD.xlsx).
    # §9.7: fotos importadas ficam em importacoes/estoque/ como estoque_AAAA-MM-DD__AAAAMMDD-HHMMSS.xlsx.
    $fotos = foreach ($f in (Get-ChildItem -LiteralPath $DirDados -File -Filter 'estoque_*.xlsx')) {
        if ($f.Name -notmatch $script:PadraoArquivoEstoque) { continue }
        $data = ConvertTo-Data $Matches[1]
        if ($data) { [pscustomobject]@{ Arquivo = $f; Data = $data; Importada = $false; Carimbo = '' } }
    }
    $pasta = if ($DirImportacoes) { Join-Path $DirImportacoes 'estoque' }
    if ($pasta -and (Test-Path -LiteralPath $pasta)) {
        $fotos = @($fotos) + @(foreach ($f in (Get-ChildItem -LiteralPath $pasta -File -Filter 'estoque_*.xlsx')) {
                if ($f.Name -notmatch $script:PadraoEstoqueImportado) { continue }
                $data = ConvertTo-Data $Matches[1]
                if ($data) { [pscustomobject]@{ Arquivo = $f; Data = $data; Importada = $true; Carimbo = $Matches[2] } }
            })
    }
    # §9.8: vale a data mais recente; no empate, a importação mais recente vence a planilha base.
    return @($fotos | Where-Object { $_ } | Sort-Object -Property @{ Expression = { $_.Data }; Descending = $true }, @{ Expression = { $_.Importada }; Descending = $true }, @{ Expression = { $_.Carimbo }; Descending = $true })
}

function Get-ArquivoEstoque([string]$DirDados, [string]$DirImportacoes) {
    return @(Get-FotosEstoque $DirDados $DirImportacoes)[0]
}

function Get-AssinaturaEstoque([string]$DirDados, [string]$DirImportacoes) {
    $partes = foreach ($nome in $script:ArquivosEstoqueFixos) {
        $p = Join-Path $DirDados $nome
        if (Test-Path -LiteralPath $p) { $f = Get-Item -LiteralPath $p; "$nome|$($f.LastWriteTimeUtc.Ticks)|$($f.Length)" } else { "$nome|ausente" }
    }
    $partes = @($partes) + @(Get-FotosEstoque $DirDados $DirImportacoes | ForEach-Object { "$($_.Arquivo.FullName)|$($_.Arquivo.LastWriteTimeUtc.Ticks)|$($_.Arquivo.Length)" })
    return (@($partes) + @(Get-AssinaturaAlteracoes $DirImportacoes) -join ';')
}

function ConvertTo-Quantidade([string]$Texto) {
    $n = ConvertTo-Inteiro $Texto
    if ($null -eq $n -or $n -lt 0) { return $null }
    return $n
}

function Import-BaseEstoque {
    param(
        [Parameter(Mandatory)][string]$DirDados,
        # §9.8: quando informado, fotos importadas também concorrem a foto em vigor. Sem ele, só a base (conferências, §6).
        [string]$DirImportacoes,
        # §9.5: valida uma foto candidata específica (o resumo da importação), em vez da foto em vigor.
        [string]$ArquivoFoto,
        [Nullable[datetime]]$DataFoto
    )

    $erros = New-Object Collections.Generic.List[string]
    foreach ($nome in $script:ArquivosEstoqueFixos) {
        if (-not (Test-Path -LiteralPath (Join-Path $DirDados $nome))) { $erros.Add("Arquivo não encontrado: $(Join-Path $DirDados $nome)") }
    }
    $foto = if ($ArquivoFoto) { [pscustomobject]@{ Arquivo = (Get-Item -LiteralPath $ArquivoFoto); Data = $DataFoto; Importada = $true; Carimbo = '' } } else { Get-ArquivoEstoque $DirDados $DirImportacoes }
    if (-not $foto) { $erros.Add("Nenhuma foto de estoque (estoque_AAAA-MM-DD.xlsx) em $DirDados") }
    if ($erros.Count) { return [pscustomobject]@{ Erros = $erros.ToArray() } }

    # Filiais (§0.2): nome vira ID; nome desconhecido é erro.
    $filiais = New-Object Collections.Generic.List[object]
    $idFilialPorNome = @{}
    foreach ($f in (Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Filiais')) {
        $idFilialPorNome[$f.Filial.Normalize()] = $f.'ID Filial'
        $filiais.Add([pscustomobject]@{ Id = $f.'ID Filial'; Nome = $f.Filial })
    }

    # Catálogo
    $produtos = New-Object Collections.Generic.List[object]
    $produtoPorId = @{}
    # §11.4: produtos cadastrados ou editados por alteração confirmada valem por cima do catálogo.
    $linhasProdutos = Merge-Alteracoes (Read-XlsxSheet (Join-Path $DirDados 'produtos.xlsx') 'Produtos') (Read-Alteracoes $DirImportacoes) 'produto' { param($r) $r.'ID Produto' } $script:ColunasProdutos $erros
    foreach ($r in $linhasProdutos) {
        $onde = "produtos.xlsx, linha $($r._Linha)"
        $id = $r.'ID Produto'
        if (-not $id) { $erros.Add("${onde}: ID Produto vazio"); continue }
        if ($produtoPorId.ContainsKey($id)) { $erros.Add("${onde}: ID Produto $id duplicado"); continue }
        if (-not $r.Categoria) { $erros.Add("${onde} ($id): Categoria vazia"); continue }
        if ($r.Status -notin @('Ativo', 'Descontinuado')) { $erros.Add("${onde} ($id): Status inválido '$($r.Status)'"); continue }
        $p = [pscustomobject]@{
            Id            = $id
            Nome          = $r.Produto
            Categoria     = $r.Categoria
            Unidade       = $r.Unidade
            Status        = $r.Status
            Descontinuado = ($r.Status -eq 'Descontinuado')
        }
        $produtos.Add($p)
        $produtoPorId[$id] = $p
    }

    # Foto do estoque: uma linha por (produto, filial).
    $nomeArq = $foto.Arquivo.Name
    $linhas = New-Object Collections.Generic.List[object]
    $chaves = @{}
    foreach ($r in (Read-XlsxSheet $foto.Arquivo.FullName 'Estoque')) {
        $onde = "$nomeArq, linha $($r._Linha)"
        $prod = $produtoPorId[$r.'ID Produto']
        $idFilial = if ($r.Filial) { $idFilialPorNome[$r.Filial.Normalize()] }
        $qtd = ConvertTo-Quantidade $r.Quantidade
        $min = ConvertTo-Quantidade $r.'Estoque Mínimo'
        $custo = ConvertTo-Dinheiro $r.'Custo Médio'
        $entrada = if ($r.'Data da Última Entrada') { ConvertTo-Data $r.'Data da Última Entrada' }
        $saida = if ($r.'Data da Última Saída') { ConvertTo-Data $r.'Data da Última Saída' }
        $falhou = $false
        if (-not $prod) { $erros.Add("${onde}: ID Produto '$($r.'ID Produto')' não existe em produtos.xlsx"); $falhou = $true }
        if (-not $idFilial) { $erros.Add("${onde}: filial desconhecida '$($r.Filial)'"); $falhou = $true }
        if ($null -eq $qtd) { $erros.Add("${onde}: Quantidade inválida '$($r.Quantidade)'"); $falhou = $true }
        if ($null -eq $min) { $erros.Add("${onde}: Estoque Mínimo inválido '$($r.'Estoque Mínimo')'"); $falhou = $true }
        if ($null -eq $custo -or $custo -lt 0) { $erros.Add("${onde}: Custo Médio inválido '$($r.'Custo Médio')'"); $falhou = $true }
        if ($r.'Data da Última Entrada' -and -not $entrada) { $erros.Add("${onde}: Data da Última Entrada inválida '$($r.'Data da Última Entrada')'"); $falhou = $true }
        if ($r.'Data da Última Saída' -and -not $saida) { $erros.Add("${onde}: Data da Última Saída inválida '$($r.'Data da Última Saída')'"); $falhou = $true }
        if ($falhou) { continue }
        $chave = "$($prod.Id)|$idFilial"
        if ($chaves.ContainsKey($chave)) { $erros.Add("${onde}: produto $($prod.Id) repetido na filial $($r.Filial)"); continue }
        $chaves[$chave] = $true

        # §5: último movimento = a mais recente entre entrada e saída; em branco nas duas = sem movimento registrado.
        $ultimo = if ($entrada -and $saida) { if ($entrada -gt $saida) { $entrada } else { $saida } } elseif ($entrada) { $entrada } else { $saida }
        $linhas.Add([pscustomobject]@{
                Produto          = $prod
                IdFilial         = $idFilial
                Filial           = $r.Filial
                Quantidade       = $qtd
                Minimo           = $min
                CustoMedio       = $custo
                Valor            = $qtd * $custo   # §5: valor imobilizado da linha
                UltimaEntrada    = $entrada
                UltimaSaida      = $saida
                UltimoMovimento  = $ultimo
                DiasSemMovimento = if ($ultimo) { ($foto.Data - $ultimo).Days } else { $null }
                Linha            = $r._Linha
            })
    }

    return [pscustomobject]@{
        Erros        = $erros.ToArray()
        DataBase     = $foto.Data
        Arquivo      = $foto.Arquivo
        Filiais      = $filiais.ToArray()
        Produtos     = $produtos.ToArray()
        ProdutoPorId = $produtoPorId
        Linhas       = $linhas.ToArray()
        CarregadoEm  = Get-Date
    }
}

function Test-LinhaParada($Linha, [int]$DiasParado) {
    # §5: PARADO ⇔ Quantidade > 0 E (sem movimento registrado OU dias sem movimento ≥ prazo).
    if ($Linha.Quantidade -le 0) { return $false }
    return ($null -eq $Linha.DiasSemMovimento -or $Linha.DiasSemMovimento -ge $DiasParado)
}

function Test-AbaixoMinimo($Linha) { return $Linha.Quantidade -lt $Linha.Minimo }   # §5, por linha

function Get-AbertasPorProduto($Pipeline) {
    # §5.4: oportunidades abertas do produto (todas as filiais), para mostrar ao lado dos parados.
    $mapa = @{}
    if (-not $Pipeline -or $Pipeline.Erros.Count) { return $null }
    foreach ($o in $Pipeline.Oportunidades) {
        if (-not $o.Aberta) { continue }
        if (-not $mapa.ContainsKey($o.IdProduto)) { $mapa[$o.IdProduto] = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0 } }
        $mapa[$o.IdProduto].Qtd++
        $mapa[$o.IdProduto].Valor += $o.Valor
    }
    return $mapa
}

function Get-VisaoEstoque {
    param(
        [Parameter(Mandatory)]$Estoque,
        [Parameter(Mandatory)][int]$DiasParado,
        $AbertasPorProduto   # $null = CRM indisponível
    )
    $semOp = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0 }
    $opsDe = { param($id) if ($null -eq $AbertasPorProduto) { $null } elseif ($AbertasPorProduto.ContainsKey($id)) { $AbertasPorProduto[$id] } else { $semOp } }

    $linhasPorProduto = @{}
    foreach ($l in $Estoque.Linhas) {
        if (-not $linhasPorProduto.ContainsKey($l.Produto.Id)) { $linhasPorProduto[$l.Produto.Id] = New-Object Collections.Generic.List[object] }
        $linhasPorProduto[$l.Produto.Id].Add($l)
    }
    $outras = { param($l) @($linhasPorProduto[$l.Produto.Id] | Where-Object { $_.IdFilial -ne $l.IdFilial } | Sort-Object IdFilial) }

    # 1. Dinheiro no pátio: valor imobilizado por categoria × filial, com a parte parada.
    $celulas = @{}
    $categorias = @($Estoque.Produtos | ForEach-Object Categoria | Sort-Object -Unique)
    foreach ($c in $categorias) { foreach ($f in $Estoque.Filiais) { $celulas["$c|$($f.Id)"] = [pscustomobject]@{ Valor = [decimal]0; Parado = [decimal]0; Unidades = 0 } } }
    $totFilial = @{}; foreach ($f in $Estoque.Filiais) { $totFilial[$f.Id] = [pscustomobject]@{ Valor = [decimal]0; Parado = [decimal]0; Unidades = 0 } }
    $totCategoria = @{}; foreach ($c in $categorias) { $totCategoria[$c] = [pscustomobject]@{ Valor = [decimal]0; Parado = [decimal]0; Unidades = 0 } }
    $total = [pscustomobject]@{ Valor = [decimal]0; Parado = [decimal]0; Unidades = 0 }

    $parados = New-Object Collections.Generic.List[object]
    $abaixo = New-Object Collections.Generic.List[object]
    foreach ($l in $Estoque.Linhas) {
        $parado = Test-LinhaParada $l $DiasParado
        foreach ($acum in @($celulas["$($l.Produto.Categoria)|$($l.IdFilial)"], $totFilial[$l.IdFilial], $totCategoria[$l.Produto.Categoria], $total)) {
            $acum.Valor += $l.Valor; $acum.Unidades += $l.Quantidade
            if ($parado) { $acum.Parado += $l.Valor }
        }
        if ($parado) { $parados.Add([pscustomobject]@{ Linha = $l; Abertas = (& $opsDe $l.Produto.Id); Outras = (& $outras $l) }) }
        if (Test-AbaixoMinimo $l) { $abaixo.Add([pscustomobject]@{ Linha = $l; Falta = $l.Minimo - $l.Quantidade; Abertas = (& $opsDe $l.Produto.Id); Outras = (& $outras $l) }) }
    }
    $categorias = @($categorias | Sort-Object -Property @{ Expression = { $totCategoria[$_].Valor }; Descending = $true }, @{ Expression = { $_ } })

    # 2. Sem estoque: soma da quantidade do produto nas filiais = 0 (produto sem linha na foto conta como zerado).
    $semEstoque = foreach ($p in $Estoque.Produtos) {
        $ls = if ($linhasPorProduto.ContainsKey($p.Id)) { $linhasPorProduto[$p.Id].ToArray() } else { @() }
        $soma = 0; foreach ($l in $ls) { $soma += $l.Quantidade }
        if ($soma -ne 0) { continue }
        $ultSaida = $null; foreach ($l in $ls) { if ($l.UltimaSaida -and (-not $ultSaida -or $l.UltimaSaida -gt $ultSaida)) { $ultSaida = $l.UltimaSaida } }
        [pscustomobject]@{ Produto = $p; Linhas = @($ls | Sort-Object IdFilial); UltimaSaida = $ultSaida; Abertas = (& $opsDe $p.Id) }
    }
    $semEstoque = @($semEstoque | Sort-Object -Property @{ Expression = { if ($_.Abertas) { $_.Abertas.Qtd } else { 0 } }; Descending = $true }, @{ Expression = { $_.Produto.Id } })

    $parados = @($parados | Sort-Object -Property @{ Expression = { if ($null -eq $_.Linha.DiasSemMovimento) { [int]::MaxValue } else { $_.Linha.DiasSemMovimento } }; Descending = $true }, @{ Expression = { $_.Linha.Produto.Id } }, @{ Expression = { $_.Linha.IdFilial } })
    $abaixo = @($abaixo | Sort-Object -Property @{ Expression = { $_.Linha.IdFilial } }, @{ Expression = { $_.Falta }; Descending = $true }, @{ Expression = { $_.Linha.Produto.Id } })

    return [pscustomobject]@{
        DataBase     = $Estoque.DataBase
        Arquivo      = $Estoque.Arquivo
        DiasParado   = $DiasParado
        Filiais      = $Estoque.Filiais
        Categorias   = $categorias
        Celulas      = $celulas
        TotFilial    = $totFilial
        TotCategoria = $totCategoria
        Total        = $total
        Parados      = $parados
        SemEstoque   = $semEstoque
        Abaixo       = $abaixo
        ComCrm       = ($null -ne $AbertasPorProduto)
        Linhas       = $Estoque.Linhas
    }
}
