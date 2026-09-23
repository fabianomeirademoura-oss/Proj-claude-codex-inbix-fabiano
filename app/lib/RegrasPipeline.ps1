# Regras do pipeline do CRM — implementação do REGRAS_NEGOCIO.md §7.
# Único lugar que decide o que é oportunidade aberta, pipeline, ponderado e atrasada.
# Depende de Regras.ps1 (conversões e base comercial).

# REGRAS_NEGOCIO.md §0.3: data-base do CRM. O arquivo não traz a data da foto, então ela vem do contrato.
$script:DataBaseCrm = [datetime]'2026-08-31'
$script:EtapasFechadas = @('Fechada Ganha', 'Fechada Perdida')
$script:ArquivosPipeline = @('crm_oportunidades.xlsx', 'clientes.xlsx', 'vendedores.xlsx', 'produtos.xlsx')
# REGRAS_NEGOCIO.md §8: ordem do funil, de cima para baixo. Etapa aberta fora desta lista entra depois, pela probabilidade.
$script:EtapasFunil = @('Prospecção', 'Qualificação', 'Proposta Enviada', 'Negociação')

function Get-AssinaturaPipeline([string]$DirDados) {
    $partes = foreach ($nome in $script:ArquivosPipeline) {
        $p = Join-Path $DirDados $nome
        if (Test-Path -LiteralPath $p) { $f = Get-Item -LiteralPath $p; "$nome|$($f.LastWriteTimeUtc.Ticks)|$($f.Length)" } else { "$nome|ausente" }
    }
    return ($partes -join ';')
}

function Test-OportunidadeAberta([string]$Etapa) {
    # REGRAS_NEGOCIO.md §7: aberta = etapa não é Fechada Ganha nem Fechada Perdida.
    return $Etapa -notin $script:EtapasFechadas
}

function Import-BasePipeline {
    param(
        [Parameter(Mandatory)][string]$DirDados,
        [Parameter(Mandatory)]$Comercial   # base de Regras.ps1, já sem erros: fornece o cadastro de vendedores
    )
    $erros = New-Object Collections.Generic.List[string]
    $caminhos = @{}
    foreach ($nome in $script:ArquivosPipeline) {
        $p = Join-Path $DirDados $nome
        if (-not (Test-Path -LiteralPath $p)) { $erros.Add("Arquivo não encontrado: $p") }
        $caminhos[$nome] = $p
    }
    if ($erros.Count) { return [pscustomobject]@{ Erros = $erros.ToArray() } }

    # Filiais (§0.2)
    $idFilialPorNome = @{}
    $nomeFilial = @{}
    $filiais = New-Object Collections.Generic.List[object]
    foreach ($f in (Read-XlsxSheet $caminhos['vendedores.xlsx'] 'Filiais')) {
        $idFilialPorNome[$f.Filial.Normalize()] = $f.'ID Filial'
        $nomeFilial[$f.'ID Filial'] = $f.Filial
        $filiais.Add([pscustomobject]@{ Id = $f.'ID Filial'; Nome = $f.Filial })
    }

    # Produtos: só o necessário para a categoria do funil (§8.4). A chave é o ID.
    $produtoPorId = @{}
    foreach ($r in (Read-XlsxSheet $caminhos['produtos.xlsx'] 'Produtos')) {
        if ($r.'ID Produto') { $produtoPorId[$r.'ID Produto'] = [pscustomobject]@{ Id = $r.'ID Produto'; Nome = $r.Produto; Categoria = $r.Categoria } }
    }

    # Clientes
    $clientePorId = @{}
    foreach ($r in (Read-XlsxSheet $caminhos['clientes.xlsx'] 'Clientes')) {
        $onde = "clientes.xlsx, linha $($r._Linha)"
        $id = $r.'ID Cliente'
        if (-not $id) { $erros.Add("${onde}: ID Cliente vazio"); continue }
        if ($clientePorId.ContainsKey($id)) { $erros.Add("${onde}: ID Cliente $id duplicado"); continue }
        $idFilial = if ($r.'Filial de Atendimento') { $idFilialPorNome[$r.'Filial de Atendimento'.Normalize()] }
        if (-not $idFilial) { $erros.Add("${onde} ($id): filial de atendimento desconhecida '$($r.'Filial de Atendimento')'"); continue }
        $clientePorId[$id] = [pscustomobject]@{
            Id       = $id
            Nome     = $r.Cliente
            Tipo     = $r.Tipo
            Cidade   = $r.Cidade
            UF       = $r.UF
            IdFilial = $idFilial
            Filial   = $nomeFilial[$idFilial]
        }
    }

    # Oportunidades
    $oportunidades = New-Object Collections.Generic.List[object]
    $ids = @{}
    foreach ($r in (Read-XlsxSheet $caminhos['crm_oportunidades.xlsx'] 'Oportunidades')) {
        $onde = "crm_oportunidades.xlsx, linha $($r._Linha)"
        $id = $r.'ID Oportunidade'
        if (-not $id) { $erros.Add("${onde}: ID Oportunidade vazio"); continue }
        if ($ids.ContainsKey($id)) { $erros.Add("${onde}: ID Oportunidade $id duplicado"); continue }
        $ids[$id] = $true
        $vend = $Comercial.VendedorPorId[$r.'ID Vendedor']
        $cli = $clientePorId[$r.'ID Cliente']
        $valor = ConvertTo-Dinheiro $r.'Valor Estimado'
        $prob = $null
        $d = [decimal]0
        if ([decimal]::TryParse([string]$r.Probabilidade, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { $prob = $d }
        $criacao = ConvertTo-Data $r.'Data de Criação'
        $previsao = if ($r.'Previsão de Fechamento') { ConvertTo-Data $r.'Previsão de Fechamento' }
        $aberta = Test-OportunidadeAberta $r.Etapa
        $falhou = $false
        if (-not $vend) { $erros.Add("${onde} ($id): ID Vendedor '$($r.'ID Vendedor')' não existe no cadastro"); $falhou = $true }
        if (-not $cli) { $erros.Add("${onde} ($id): ID Cliente '$($r.'ID Cliente')' não existe em clientes.xlsx"); $falhou = $true }
        if (-not $produtoPorId.ContainsKey([string]$r.'ID Produto')) { $erros.Add("${onde} ($id): ID Produto '$($r.'ID Produto')' não existe em produtos.xlsx"); $falhou = $true }
        if (-not $r.Etapa) { $erros.Add("${onde} ($id): Etapa vazia"); $falhou = $true }
        if ($null -eq $valor -or $valor -lt 0) { $erros.Add("${onde} ($id): Valor Estimado inválido '$($r.'Valor Estimado')'"); $falhou = $true }
        if ($null -eq $prob -or $prob -lt 0 -or $prob -gt 1) { $erros.Add("${onde} ($id): Probabilidade inválida '$($r.Probabilidade)' (esperado fração de 0 a 1)"); $falhou = $true }
        if (-not $criacao) { $erros.Add("${onde} ($id): Data de Criação inválida '$($r.'Data de Criação')'"); $falhou = $true }
        if ($r.'Previsão de Fechamento' -and -not $previsao) { $erros.Add("${onde} ($id): Previsão de Fechamento inválida '$($r.'Previsão de Fechamento')'"); $falhou = $true }
        if ($aberta -and -not $r.'Previsão de Fechamento') { $erros.Add("${onde} ($id): oportunidade aberta sem Previsão de Fechamento"); $falhou = $true }
        if ($falhou) { continue }
        # §2.6: oportunidade nova depois do desligamento do dono é erro. §3.4: não validar criação ≥ admissão.
        if ($vend.Desligamento -and $criacao -gt $vend.Desligamento) {
            $erros.Add("${onde} ($id): criada em $($criacao.ToString('dd/MM/yyyy')) para vendedor desligado em $($vend.Desligamento.ToString('dd/MM/yyyy'))")
            continue
        }
        $oportunidades.Add([pscustomobject]@{
                Id            = $id
                Criacao       = $criacao
                IdVendedor    = $vend.Id
                IdCliente     = $cli.Id
                IdProduto     = $r.'ID Produto'
                Produto       = $r.Produto
                Quantidade    = $r.Quantidade
                Valor         = $valor
                Etapa         = $r.Etapa
                Probabilidade = $prob
                Previsao      = $previsao
                Origem        = $r.Origem
                MotivoPerda   = $r.'Motivo da Perda'
                Aberta        = $aberta
                Linha         = $r._Linha
            })
    }

    return [pscustomobject]@{
        Erros         = $erros.ToArray()
        DataBase      = $script:DataBaseCrm
        Oportunidades = $oportunidades.ToArray()
        ClientePorId  = $clientePorId
        ProdutoPorId  = $produtoPorId
        Filiais       = $filiais.ToArray()
        Arquivo       = Get-Item -LiteralPath $caminhos['crm_oportunidades.xlsx']
        CarregadoEm   = Get-Date
    }
}

function Get-VisaoPipeline {
    # Uma linha por vendedor: meta e realizado do período (§4) + pipeline da foto do CRM (§7).
    param(
        [Parameter(Mandatory)]$Comercial,
        [Parameter(Mandatory)]$Desempenho,
        [Parameter(Mandatory)]$Pipeline
    )
    $db = $Pipeline.DataBase
    $abertas = foreach ($o in $Pipeline.Oportunidades) {
        if (-not $o.Aberta) { continue }
        $vend = $Comercial.VendedorPorId[$o.IdVendedor]
        $atrasada = ($o.Previsao -lt $db)
        [pscustomobject]@{
            Oportunidade = $o
            Vendedor     = $vend
            Cliente      = $Pipeline.ClientePorId[$o.IdCliente]
            Ponderado    = [math]::Round($o.Valor * $o.Probabilidade, 2, [MidpointRounding]::AwayFromZero)
            Atrasada     = $atrasada
            DiasAtraso   = if ($atrasada) { ($db - $o.Previsao).Days } else { 0 }
            Orfa         = (-not $vend.Ativo)   # §2.4
        }
    }
    # Atrasadas primeiro (mais antigas no topo); depois por previsão de fechamento.
    $abertas = @($abertas | Sort-Object -Property @{ Expression = { $_.Atrasada }; Descending = $true }, @{ Expression = { $_.Oportunidade.Previsao } }, @{ Expression = { $_.Oportunidade.Id } })

    $porVendedor = @{}
    foreach ($a in $abertas) {
        $id = $a.Vendedor.Id
        if (-not $porVendedor.ContainsKey($id)) { $porVendedor[$id] = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0; Ponderado = [decimal]0; QtdAtrasadas = 0; ValorAtrasado = [decimal]0 } }
        $p = $porVendedor[$id]
        $p.Qtd++; $p.Valor += $a.Oportunidade.Valor; $p.Ponderado += $a.Ponderado
        if ($a.Atrasada) { $p.QtdAtrasadas++; $p.ValorAtrasado += $a.Oportunidade.Valor }
    }

    # §7.4: todos os vendedores (órfãs de inativos entram nos totais).
    $linhas = foreach ($l in ($Desempenho.Linhas | Sort-Object { $_.Vendedor.Id })) {
        $p = $porVendedor[$l.Vendedor.Id]
        [pscustomobject]@{
            Vendedor           = $l.Vendedor
            MesesComMeta       = $l.MesesComMeta
            Meta               = $l.Meta
            RealizadoMesesMeta = $l.RealizadoMesesMeta
            ForaMesesMeta      = $l.ForaMesesMeta
            Atingimento        = $l.Atingimento
            Gap                = if ($l.Meta -gt 0) { $l.Meta - $l.RealizadoMesesMeta } else { $null }   # §4.3
            QtdAbertas         = if ($p) { $p.Qtd } else { 0 }
            Pipeline           = if ($p) { $p.Valor } else { [decimal]0 }
            Ponderado          = if ($p) { $p.Ponderado } else { [decimal]0 }
            QtdAtrasadas       = if ($p) { $p.QtdAtrasadas } else { 0 }
            ValorAtrasado      = if ($p) { $p.ValorAtrasado } else { [decimal]0 }
        }
    }
    $linhas = @($linhas)

    $tot = [pscustomobject]@{ Meta = [decimal]0; RealizadoMesesMeta = [decimal]0; ForaMesesMeta = [decimal]0; QtdAbertas = 0; Pipeline = [decimal]0; Ponderado = [decimal]0; QtdAtrasadas = 0; ValorAtrasado = [decimal]0 }
    foreach ($l in $linhas) {
        $tot.Meta += $l.Meta; $tot.RealizadoMesesMeta += $l.RealizadoMesesMeta; $tot.ForaMesesMeta += $l.ForaMesesMeta
        $tot.QtdAbertas += $l.QtdAbertas; $tot.Pipeline += $l.Pipeline; $tot.Ponderado += $l.Ponderado
        $tot.QtdAtrasadas += $l.QtdAtrasadas; $tot.ValorAtrasado += $l.ValorAtrasado
    }
    $tot | Add-Member Atingimento $Desempenho.Empresa.Atingimento   # §4.2: soma ÷ soma
    $tot | Add-Member Gap $(if ($tot.Meta -gt 0) { $tot.Meta - $tot.RealizadoMesesMeta } else { $null })

    $orfas = @($abertas | Where-Object Orfa)
    $valorOrfas = [decimal]0   # soma em decimal: Measure-Object devolveria double (§0.4)
    foreach ($a in $orfas) { $valorOrfas += $a.Oportunidade.Valor }
    return [pscustomobject]@{
        DataBaseCrm    = $db
        Desempenho     = $Desempenho
        Linhas         = $linhas
        Total          = $tot
        Abertas        = $abertas
        QtdOrfas       = $orfas.Count
        ValorOrfas     = $valorOrfas
        ArquivoCrm     = $Pipeline.Arquivo
    }
}


function Get-VisaoFunil {
    # REGRAS_NEGOCIO.md §8: abertas por etapa ATUAL (o CRM não guarda histórico) + fechadas à parte.
    # Filtros combinam em E; filial = filial do vendedor (§8.4). Vazio = sem filtro.
    param(
        [Parameter(Mandatory)]$Comercial,
        [Parameter(Mandatory)]$Pipeline,
        [string]$IdVendedor,
        [string]$IdFilial,
        [string]$Categoria
    )
    $db = $Pipeline.DataBase

    # Etapas: as da §8 sempre (mesmo vazias), mais qualquer outra etapa aberta existente na foto.
    $probEtapa = @{}
    foreach ($o in $Pipeline.Oportunidades) { if ($o.Aberta -and -not $probEtapa.ContainsKey($o.Etapa)) { $probEtapa[$o.Etapa] = $o.Probabilidade } }
    $extras = @($probEtapa.Keys | Where-Object { $_ -notin $script:EtapasFunil } | Sort-Object { $probEtapa[$_] }, { $_ })
    $etapas = [ordered]@{}
    foreach ($nome in @($script:EtapasFunil) + $extras) {
        $etapas[$nome] = [pscustomobject]@{ Etapa = $nome; Probabilidade = $probEtapa[$nome]; Qtd = 0; Valor = [decimal]0; Ponderado = [decimal]0; QtdAtrasadas = 0; ValorAtrasado = [decimal]0 }
    }

    $abertas = New-Object Collections.Generic.List[object]
    $ganhas = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0 }
    $perdidas = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0 }
    $motivos = @{}
    foreach ($o in $Pipeline.Oportunidades) {
        $vend = $Comercial.VendedorPorId[$o.IdVendedor]
        $prod = $Pipeline.ProdutoPorId[$o.IdProduto]
        if ($IdVendedor -and $o.IdVendedor -ne $IdVendedor) { continue }
        if ($IdFilial -and $vend.IdFilial -ne $IdFilial) { continue }
        if ($Categoria -and $prod.Categoria -ne $Categoria) { continue }
        if ($o.Etapa -eq 'Fechada Ganha') { $ganhas.Qtd++; $ganhas.Valor += $o.Valor; continue }
        if ($o.Etapa -eq 'Fechada Perdida') {
            $perdidas.Qtd++; $perdidas.Valor += $o.Valor
            $m = if ($o.MotivoPerda) { $o.MotivoPerda } else { '(sem motivo informado)' }
            if (-not $motivos.ContainsKey($m)) { $motivos[$m] = [pscustomobject]@{ Motivo = $m; Qtd = 0; Valor = [decimal]0 } }
            $motivos[$m].Qtd++; $motivos[$m].Valor += $o.Valor
            continue
        }
        $atrasada = ($o.Previsao -lt $db)   # §7
        $a = [pscustomobject]@{
            Oportunidade = $o
            Vendedor     = $vend
            Cliente      = $Pipeline.ClientePorId[$o.IdCliente]
            Produto      = $prod
            Ponderado    = [math]::Round($o.Valor * $o.Probabilidade, 2, [MidpointRounding]::AwayFromZero)
            Atrasada     = $atrasada
            DiasAtraso   = if ($atrasada) { ($db - $o.Previsao).Days } else { 0 }
            Orfa         = (-not $vend.Ativo)
        }
        $abertas.Add($a)
        $e = $etapas[$o.Etapa]
        $e.Qtd++; $e.Valor += $o.Valor; $e.Ponderado += $a.Ponderado
        if ($atrasada) { $e.QtdAtrasadas++; $e.ValorAtrasado += $o.Valor }
    }

    $ordem = @{}; $i = 0; foreach ($k in $etapas.Keys) { $ordem[$k] = $i++ }
    $lista = @($abertas | Sort-Object -Property @{ Expression = { $ordem[$_.Oportunidade.Etapa] }; Descending = $true }, @{ Expression = { $_.Atrasada }; Descending = $true }, @{ Expression = { $_.Oportunidade.Previsao } }, @{ Expression = { $_.Oportunidade.Id } })

    $tot = [pscustomobject]@{ Qtd = 0; Valor = [decimal]0; Ponderado = [decimal]0; QtdAtrasadas = 0; ValorAtrasado = [decimal]0 }
    foreach ($e in $etapas.Values) { $tot.Qtd += $e.Qtd; $tot.Valor += $e.Valor; $tot.Ponderado += $e.Ponderado; $tot.QtdAtrasadas += $e.QtdAtrasadas; $tot.ValorAtrasado += $e.ValorAtrasado }
    $fechadas = $ganhas.Qtd + $perdidas.Qtd
    $valorFechadas = $ganhas.Valor + $perdidas.Valor

    # Categorias que existem no CRM (peças não passam por ele).
    $categorias = @($Pipeline.Oportunidades | ForEach-Object { $Pipeline.ProdutoPorId[$_.IdProduto].Categoria } | Sort-Object -Unique)

    return [pscustomobject]@{
        DataBaseCrm     = $db
        ArquivoCrm      = $Pipeline.Arquivo
        Etapas          = @($etapas.Values)
        Abertas         = $lista
        Total           = $tot
        Ganhas          = $ganhas
        Perdidas        = $perdidas
        Conversao       = if ($fechadas) { [decimal]$ganhas.Qtd / $fechadas } else { $null }
        ConversaoValor  = if ($valorFechadas -gt 0) { $ganhas.Valor / $valorFechadas } else { $null }
        Motivos         = @($motivos.Values | Sort-Object -Property @{ Expression = { $_.Qtd }; Descending = $true }, @{ Expression = { $_.Motivo } })
        Filtro          = [pscustomobject]@{ IdVendedor = $IdVendedor; IdFilial = $IdFilial; Categoria = $Categoria }
        Vendedores      = @($Comercial.Vendedores | Sort-Object Nome)
        Filiais         = $Pipeline.Filiais
        Categorias      = $categorias
    }
}
