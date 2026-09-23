# Regras de negócio da Horizonte Máquinas — implementação do REGRAS_NEGOCIO.md.
# Este é o ÚNICO lugar que decide o que é venda realizada, meta do período e
# atingimento. Páginas e exportações só formatam o que sai daqui.

$script:Inv = [Globalization.CultureInfo]::InvariantCulture
$script:AnoMetas = 2026
$script:ArquivosBase = [ordered]@{
    Vendedores = 'vendedores.xlsx'
    Metas      = 'metas_2026.xlsx'
    Vendas     = 'vendas_2026_jan-ago.xlsx'
}

function ConvertTo-Dinheiro([string]$Texto) {
    # REGRAS_NEGOCIO.md §0.4: decimal com 2 casas, arredondado na importação.
    $d = [decimal]0
    if (-not [decimal]::TryParse($Texto, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $null }
    return [math]::Round($d, 2, [MidpointRounding]::AwayFromZero)
}

function ConvertTo-Data([string]$Texto) {
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Texto, [string[]]@('yyyy-MM-dd', 'dd/MM/yyyy'), $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function ConvertTo-Inteiro([string]$Texto) {
    $d = [decimal]0
    if ([decimal]::TryParse($Texto, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$d) -and $d -eq [math]::Truncate($d)) { return [int]$d }
    return $null
}

$script:PadraoVendaImportada = '^\d{8}-\d{6}__.+\.xlsx$'   # REGRAS_NEGOCIO.md §9.7: carimbo da importação + nome original

function Get-ArquivosVendasImportados([string]$DirImportacoes) {
    # §9.8: importações de vendas em ordem cronológica (o nome começa pelo carimbo AAAAMMDD-HHMMSS).
    if (-not $DirImportacoes) { return @() }
    $pasta = Join-Path $DirImportacoes 'vendas'
    if (-not (Test-Path -LiteralPath $pasta)) { return @() }
    return @(Get-ChildItem -LiteralPath $pasta -File -Filter '*.xlsx' | Where-Object { $_.Name -match $script:PadraoVendaImportada } | Sort-Object Name)
}

function Get-AssinaturaBase([string]$DirDados, [string]$DirImportacoes) {
    $partes = foreach ($nome in $script:ArquivosBase.Values) {
        $p = Join-Path $DirDados $nome
        if (Test-Path -LiteralPath $p) { $f = Get-Item -LiteralPath $p; "$nome|$($f.LastWriteTimeUtc.Ticks)|$($f.Length)" } else { "$nome|ausente" }
    }
    $partes = @($partes) + @(Get-ArquivosVendasImportados $DirImportacoes | ForEach-Object { "$($_.Name)|$($_.LastWriteTimeUtc.Ticks)|$($_.Length)" })
    return ($partes -join ';')
}

function Import-BaseComercial {
    param(
        [Parameter(Mandatory)][string]$DirDados,
        # §9.8: quando informado, as vendas importadas são aplicadas por cima da base (upsert por ID Venda).
        # Sem ele, a carga é só a base — é o que as conferências (§6) usam.
        [string]$DirImportacoes,
        # §9.5: arquivo candidato, aplicado por último; só para calcular o resumo antes de gravar.
        [string[]]$VendasAdicionais
    )

    $erros = New-Object Collections.Generic.List[string]
    $caminhos = @{}
    foreach ($chave in $script:ArquivosBase.Keys) {
        $p = Join-Path $DirDados $script:ArquivosBase[$chave]
        if (-not (Test-Path -LiteralPath $p)) { $erros.Add("Arquivo não encontrado: $p") }
        $caminhos[$chave] = $p
    }
    if ($erros.Count) { return [pscustomobject]@{ Erros = $erros.ToArray() } }

    # Filiais: REGRAS_NEGOCIO.md §0.2 — nome vira ID Filial; nome desconhecido é erro.
    $idFilialPorNome = @{}
    $nomeFilial = @{}
    foreach ($f in (Read-XlsxSheet $caminhos.Vendedores 'Filiais')) {
        $idFilialPorNome[$f.Filial.Normalize()] = $f.'ID Filial'
        $nomeFilial[$f.'ID Filial'] = $f.Filial
    }

    # Vendedores
    $vendedores = New-Object Collections.Generic.List[object]
    $vendedorPorId = @{}
    foreach ($r in (Read-XlsxSheet $caminhos.Vendedores 'Vendedores')) {
        $onde = "vendedores.xlsx, linha $($r._Linha)"
        $id = $r.'ID Vendedor'
        if (-not $id) { $erros.Add("${onde}: ID Vendedor vazio"); continue }
        if ($vendedorPorId.ContainsKey($id)) { $erros.Add("${onde}: ID Vendedor $id duplicado"); continue }
        $idFilial = if ($r.Filial) { $idFilialPorNome[$r.Filial.Normalize()] }
        if (-not $idFilial) { $erros.Add("${onde}: filial desconhecida '$($r.Filial)'") }
        if ($r.Status -notin @('Ativo', 'Inativo')) { $erros.Add("${onde}: Status inválido '$($r.Status)'") }
        $admissao = ConvertTo-Data $r.'Data de Admissão'
        $desligamento = if ($r.'Data de Desligamento') { ConvertTo-Data $r.'Data de Desligamento' }
        if ($r.Status -eq 'Inativo' -and -not $desligamento) { $erros.Add("${onde}: vendedor Inativo sem Data de Desligamento válida") }
        $v = [pscustomobject]@{
            Id           = $id
            Nome         = $r.Nome
            Email        = $r.'E-mail'
            IdFilial     = $idFilial
            Filial       = $nomeFilial[$idFilial]
            Status       = $r.Status
            Ativo        = ($r.Status -eq 'Ativo')
            Admissao     = $admissao
            Desligamento = $desligamento
        }
        $vendedores.Add($v)
        $vendedorPorId[$id] = $v
    }

    # Metas: uma linha por (vendedor, ano, mês). Ausência de linha = sem meta (§3.1).
    $metas = New-Object Collections.Generic.List[object]
    $chavesMeta = @{}
    foreach ($r in (Read-XlsxSheet $caminhos.Metas 'Metas')) {
        $onde = "metas_2026.xlsx, linha $($r._Linha)"
        $id = $r.'ID Vendedor'
        $ano = ConvertTo-Inteiro $r.Ano
        $mes = ConvertTo-Inteiro $r.'Mês'
        $valor = ConvertTo-Dinheiro $r.'Meta (R$)'
        if (-not $vendedorPorId.ContainsKey($id)) { $erros.Add("${onde}: ID Vendedor '$id' não existe no cadastro"); continue }
        if ($null -eq $ano -or $null -eq $mes -or $mes -lt 1 -or $mes -gt 12) { $erros.Add("${onde}: Ano/Mês inválido"); continue }
        if ($null -eq $valor -or $valor -lt 0) { $erros.Add("${onde}: Meta inválida '$($r.'Meta (R$)')'"); continue }
        $chave = "$id|$ano|$mes"
        if ($chavesMeta.ContainsKey($chave)) { $erros.Add("${onde}: meta duplicada para $id em $mes/$ano"); continue }
        $chavesMeta[$chave] = $true
        $metas.Add([pscustomobject]@{ IdVendedor = $id; Ano = $ano; Mes = $mes; Valor = $valor })
    }

    # Vendas: a base e, na ordem, cada arquivo importado. §1 e §9.8: upsert por ID Venda — a última versão vale.
    $fontesVendas = @([pscustomobject]@{ Caminho = $caminhos.Vendas; Nome = $script:ArquivosBase.Vendas })
    $fontesVendas += @(Get-ArquivosVendasImportados $DirImportacoes | ForEach-Object { [pscustomobject]@{ Caminho = $_.FullName; Nome = $_.Name } })
    $fontesVendas += @($VendasAdicionais | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ Caminho = $_; Nome = Split-Path $_ -Leaf } })
    $vendaPorId = [ordered]@{}
    foreach ($fonte in $fontesVendas) {
        $idsVenda = @{}   # ID repetido dentro do MESMO arquivo é erro (§9.4)
        foreach ($r in (Read-XlsxSheet $fonte.Caminho 'Vendas')) {
            $onde = "$($fonte.Nome), linha $($r._Linha)"
            $id = $r.'ID Venda'
            if (-not $id) { $erros.Add("${onde}: ID Venda vazio"); continue }
            if ($idsVenda.ContainsKey($id)) { $erros.Add("${onde}: ID Venda $id duplicado"); continue }
            $idsVenda[$id] = $true
            $data = ConvertTo-Data $r.Data
            $valor = ConvertTo-Dinheiro $r.'Valor Total'
            $vend = $vendedorPorId[$r.'ID Vendedor']
            $idFilial = if ($r.Filial) { $idFilialPorNome[$r.Filial.Normalize()] }
            $falhou = $false
            if (-not $data) { $erros.Add("${onde} ($id): Data inválida '$($r.Data)'"); $falhou = $true }
            if ($null -eq $valor) { $erros.Add("${onde} ($id): Valor Total inválido '$($r.'Valor Total')'"); $falhou = $true }
            if (-not $vend) { $erros.Add("${onde} ($id): ID Vendedor '$($r.'ID Vendedor')' não existe no cadastro"); $falhou = $true }
            if (-not $idFilial) { $erros.Add("${onde} ($id): filial desconhecida '$($r.Filial)'"); $falhou = $true }
            if ($r.Status -notin @('Faturada', 'Cancelada')) { $erros.Add("${onde} ($id): Status inválido '$($r.Status)'"); $falhou = $true }
            if ($falhou) { continue }
            # REGRAS_NEGOCIO.md §2.6: venda posterior ao desligamento é erro de importação.
            if ($vend.Desligamento -and $data -gt $vend.Desligamento) {
                $erros.Add("${onde} ($id): venda em $($data.ToString('dd/MM/yyyy')) de vendedor desligado em $($vend.Desligamento.ToString('dd/MM/yyyy'))")
                continue
            }
            $vendaPorId[$id] = [pscustomobject]@{
                Id             = $id
                Data           = $data
                IdVendedor     = $vend.Id
                IdFilial       = $idFilial
                Cliente        = $r.Cliente
                IdCliente      = $r.'ID Cliente'
                Produto        = $r.Produto
                IdProduto      = $r.'ID Produto'
                Categoria      = $r.Categoria
                Quantidade     = $r.Quantidade
                Valor          = $valor
                Status         = $r.Status
                IdOportunidade = $r.'ID Oportunidade'
                Linha          = $r._Linha
                Arquivo        = $fonte.Nome   # de onde veio a versão em vigor (a base ou uma importação)
            }
        }
    }
    $vendas = @($vendaPorId.Values)

    $dataBase = $null
    foreach ($v in $vendas) { if ($null -eq $dataBase -or $v.Data -gt $dataBase) { $dataBase = $v.Data } }
    $arquivos = foreach ($chave in $script:ArquivosBase.Keys) {
        $f = Get-Item -LiteralPath $caminhos[$chave]
        [pscustomobject]@{ Nome = $f.Name; Caminho = $f.FullName; Modificado = $f.LastWriteTime }
    }
    $arquivos = @($arquivos) + @(Get-ArquivosVendasImportados $DirImportacoes | ForEach-Object { [pscustomobject]@{ Nome = $_.Name; Caminho = $_.FullName; Modificado = $_.LastWriteTime } })

    return [pscustomobject]@{
        Erros         = $erros.ToArray()
        Vendedores    = $vendedores.ToArray()
        VendedorPorId = $vendedorPorId
        Metas         = $metas.ToArray()
        Vendas        = $vendas
        DataBase      = $dataBase
        Arquivos      = @($arquivos)
        CarregadoEm   = Get-Date
    }
}

function Test-VendaRealizada($Venda) {
    # REGRAS_NEGOCIO.md §1: venda realizada é Status = 'Faturada'. Nada mais.
    return $Venda.Status -eq 'Faturada'
}

function Get-MesesDisponiveis($Base) {
    # Metas são mensais (§4): o período vai de janeiro até o mês da data-base de vendas.
    if (-not $Base.DataBase -or $Base.DataBase.Year -ne $script:AnoMetas) { return @() }
    return @(1..$Base.DataBase.Month)
}

function Get-Desempenho {
    param(
        [Parameter(Mandatory)]$Base,
        [Parameter(Mandatory)][int]$MesInicial,
        [Parameter(Mandatory)][int]$MesFinal
    )
    $mesesPeriodo = @{}
    foreach ($m in $MesInicial..$MesFinal) { $mesesPeriodo[$m] = $true }

    $metasPorVendedor = @{}
    foreach ($m in $Base.Metas) {
        if ($m.Ano -ne $script:AnoMetas -or -not $mesesPeriodo.ContainsKey($m.Mes)) { continue }
        if (-not $metasPorVendedor.ContainsKey($m.IdVendedor)) { $metasPorVendedor[$m.IdVendedor] = @{} }
        $metasPorVendedor[$m.IdVendedor][$m.Mes] = $m.Valor
    }

    $realizadasPorVendedor = @{}
    $qtdCanceladas = 0
    $valorCanceladas = [decimal]0
    foreach ($v in $Base.Vendas) {
        if ($v.Data.Year -ne $script:AnoMetas -or -not $mesesPeriodo.ContainsKey($v.Data.Month)) { continue }
        if (-not (Test-VendaRealizada $v)) { $qtdCanceladas++; $valorCanceladas += $v.Valor; continue }
        if (-not $realizadasPorVendedor.ContainsKey($v.IdVendedor)) { $realizadasPorVendedor[$v.IdVendedor] = New-Object Collections.Generic.List[object] }
        $realizadasPorVendedor[$v.IdVendedor].Add($v)
    }

    $linhas = foreach ($vend in $Base.Vendedores) {
        $metasMes = if ($metasPorVendedor.ContainsKey($vend.Id)) { $metasPorVendedor[$vend.Id] } else { @{} }
        $meta = [decimal]0
        foreach ($valor in $metasMes.Values) { $meta += $valor }
        $realizado = [decimal]0
        $realizadoMesesMeta = [decimal]0
        $qtd = 0
        if ($realizadasPorVendedor.ContainsKey($vend.Id)) {
            foreach ($venda in $realizadasPorVendedor[$vend.Id]) {
                $realizado += $venda.Valor
                $qtd++
                if ($metasMes.ContainsKey($venda.Data.Month)) { $realizadoMesesMeta += $venda.Valor }
            }
        }
        [pscustomobject]@{
            Vendedor           = $vend
            MesesComMeta       = @($metasMes.Keys | Sort-Object)
            Meta               = $meta
            Realizado          = $realizado            # faturamento no período (ranking por faturamento)
            RealizadoMesesMeta = $realizadoMesesMeta   # numerador do atingimento (§4.1)
            ForaMesesMeta      = $realizado - $realizadoMesesMeta
            QtdVendas          = $qtd
            # §4.1: sem meta no período => $null ("sem meta"), nunca 0% nem erro.
            Atingimento        = if ($meta -gt 0) { $realizadoMesesMeta / $meta } else { $null }
        }
    }
    $linhas = @($linhas)

    # §4.2: empresa = soma ÷ soma, com inativos, nunca média de percentuais.
    $totalRealizado = [decimal]0; $totalRealMeta = [decimal]0; $totalMeta = [decimal]0; $totalQtd = 0
    foreach ($l in $linhas) {
        $totalRealizado += $l.Realizado; $totalRealMeta += $l.RealizadoMesesMeta
        $totalMeta += $l.Meta; $totalQtd += $l.QtdVendas
    }

    # §4.5: mês da data-base ainda não terminado => resultado parcial.
    $db = $Base.DataBase
    $parcial = ($MesFinal -eq $db.Month -and $db.AddDays(1).Month -eq $db.Month)

    return [pscustomobject]@{
        MesInicial      = $MesInicial
        MesFinal        = $MesFinal
        DataBase        = $db
        Parcial         = $parcial
        Linhas          = $linhas
        Empresa         = [pscustomobject]@{
            Realizado          = $totalRealizado
            RealizadoMesesMeta = $totalRealMeta
            ForaMesesMeta      = $totalRealizado - $totalRealMeta
            Meta               = $totalMeta
            QtdVendas          = $totalQtd
            TicketMedio        = if ($totalQtd) { [math]::Round($totalRealizado / $totalQtd, 2, [MidpointRounding]::AwayFromZero) } else { $null }
            Atingimento        = if ($totalMeta -gt 0) { $totalRealMeta / $totalMeta } else { $null }
        }
        QtdCanceladas   = $qtdCanceladas
        ValorCanceladas = $valorCanceladas
    }
}

function Get-DetalheVendedor {
    # Abre o número de um vendedor mês a mês e venda a venda, para conferência na fonte.
    param(
        [Parameter(Mandatory)]$Base,
        [Parameter(Mandatory)][string]$IdVendedor,
        [Parameter(Mandatory)][int]$MesInicial,
        [Parameter(Mandatory)][int]$MesFinal
    )
    $metaMes = @{}
    foreach ($m in $Base.Metas) {
        if ($m.IdVendedor -eq $IdVendedor -and $m.Ano -eq $script:AnoMetas) { $metaMes[$m.Mes] = $m.Valor }
    }
    $vendas = @($Base.Vendas | Where-Object {
            $_.IdVendedor -eq $IdVendedor -and $_.Data.Year -eq $script:AnoMetas -and
            $_.Data.Month -ge $MesInicial -and $_.Data.Month -le $MesFinal
        } | Sort-Object Data, Id)
    $meses = foreach ($mes in $MesInicial..$MesFinal) {
        $realizado = [decimal]0; $qtd = 0; $qtdCanc = 0; $valorCanc = [decimal]0
        foreach ($v in $vendas) {
            if ($v.Data.Month -ne $mes) { continue }
            if (Test-VendaRealizada $v) { $realizado += $v.Valor; $qtd++ } else { $qtdCanc++; $valorCanc += $v.Valor }
        }
        $meta = if ($metaMes.ContainsKey($mes)) { $metaMes[$mes] } else { $null }
        [pscustomobject]@{
            Mes             = $mes
            Meta            = $meta
            Realizado       = $realizado
            QtdVendas       = $qtd
            QtdCanceladas   = $qtdCanc
            ValorCanceladas = $valorCanc
            Atingimento     = if ($meta -gt 0) { $realizado / $meta } else { $null }
        }
    }
    return [pscustomobject]@{ Meses = @($meses); Vendas = $vendas }
}

function Get-RankingFaturamento($Desempenho, [switch]$IncluirDesligados) {
    # §2.3: rankings mostram só ativos, salvo pedido explícito.
    return @($Desempenho.Linhas |
        Where-Object { $IncluirDesligados -or $_.Vendedor.Ativo } |
        Sort-Object -Property @{ Expression = { $_.Realizado }; Descending = $true }, @{ Expression = { $_.Vendedor.Id } })
}

function Get-RankingAtingimento($Desempenho, [switch]$IncluirDesligados) {
    return @($Desempenho.Linhas |
        Where-Object { ($IncluirDesligados -or $_.Vendedor.Ativo) -and $null -ne $_.Atingimento } |
        Sort-Object -Property @{ Expression = { $_.Atingimento }; Descending = $true }, @{ Expression = { $_.Vendedor.Id } })
}
