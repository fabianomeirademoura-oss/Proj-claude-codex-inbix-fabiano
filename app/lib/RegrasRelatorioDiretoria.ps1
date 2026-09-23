# Relatório semanal da diretoria — cálculo dos itens 1 a 5.
# Não cria regra: monta o relatório com as funções de Regras.ps1, RegrasPipeline.ps1 e
# RegrasEstoque.ps1, que implementam o REGRAS_NEGOCIO.md. O item 6 (riscos) é análise do
# agente e fica fora deste arquivo; aqui só se valida o arquivo de riscos.
# Mesmos dados => mesmo resultado: nada aqui lê o relógio, e toda ordenação tem desempate por ID.
# Depende de Regras.ps1, RegrasPipeline.ps1 e RegrasEstoque.ps1.

# Decisões do coordenador para este relatório (23/09/2026, PR #3):
$script:RelCorteAtingimento = [decimal]0.80   # item 2: abaixo de 80% da meta, pelo acumulado do ano
$script:RelTopProdutos = 10                   # item 5: dez produtos da empresa, por faturamento no acumulado do ano
$script:RelNomesMes = @('', 'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro')
$script:RelMesesCurtos = @('', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez')

function Get-ImpressaoDigitalDados {
    # SHA-256 do conteúdo de cada planilha usada (nome + hash), na ordem fixa da carga.
    # Mesma impressão digital => mesmos arquivos => mesmos itens 1 a 5.
    param([Parameter(Mandatory)]$Comercial, [Parameter(Mandatory)]$Estoque, [Parameter(Mandatory)][string]$DirDados)
    $caminhos = @($Comercial.Arquivos | ForEach-Object Caminho)
    $caminhos += @('crm_oportunidades.xlsx', 'clientes.xlsx', 'produtos.xlsx' | ForEach-Object { Join-Path $DirDados $_ })
    $caminhos += $Estoque.Arquivo.FullName
    $linhas = foreach ($c in $caminhos) { "$(Split-Path $c -Leaf)|$((Get-FileHash -LiteralPath $c -Algorithm SHA256).Hash.ToLowerInvariant())" }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($linhas -join "`n"))) } finally { $sha.Dispose() }
    return (-join ($bytes | ForEach-Object { $_.ToString('x2') })).Substring(0, 16)
}

function Get-RelLinhaVendedor($LinhaAno, $LinhaMes) {
    $v = $LinhaAno.Vendedor
    return [pscustomobject]@{
        Vendedor       = $v
        MesesComMeta   = $LinhaAno.MesesComMeta
        MetaAno        = $LinhaAno.Meta
        RealizadoAno   = $LinhaAno.RealizadoMesesMeta   # numerador do atingimento (§4.1)
        AtingimentoAno = $LinhaAno.Atingimento
        GapAno         = if ($LinhaAno.Meta -gt 0) { $LinhaAno.Meta - $LinhaAno.RealizadoMesesMeta } else { $null }   # §4.3
        MetaMes        = if ($LinhaMes.Meta -gt 0) { $LinhaMes.Meta } else { $null }   # §3.1: sem linha de meta = sem meta
        RealizadoMes   = $LinhaMes.RealizadoMesesMeta
        AtingimentoMes = $LinhaMes.Atingimento
    }
}

function Get-RelatorioDiretoria {
    param(
        [Parameter(Mandatory)]$Comercial,
        [Parameter(Mandatory)]$Pipeline,
        [Parameter(Mandatory)]$Estoque,
        [Parameter(Mandatory)][string]$DirDados
    )
    $meses = Get-MesesDisponiveis $Comercial
    if (-not $meses.Count) { throw "Data-base de vendas fora de $($script:AnoMetas): não há período de metas." }
    $mesFinal = $meses[-1]
    $ano = Get-Desempenho -Base $Comercial -MesInicial 1 -MesFinal $mesFinal   # acumulado do ano (§4)
    $mes = Get-Desempenho -Base $Comercial -MesInicial $mesFinal -MesFinal $mesFinal   # mês da data-base (§4.5: meta cheia)

    # --- Item 1: ranking de ativos (§2.3) pelo atingimento no acumulado do ano; mês da data-base ao lado.
    $linhaMesPorId = @{}
    foreach ($l in $mes.Linhas) { $linhaMesPorId[$l.Vendedor.Id] = $l }
    $vendedores = foreach ($l in $ano.Linhas) { Get-RelLinhaVendedor $l $linhaMesPorId[$l.Vendedor.Id] }
    $ranking = @($vendedores | Where-Object { $_.Vendedor.Ativo } | Sort-Object -Property `
            @{ Expression = { if ($null -eq $_.AtingimentoAno) { 1 } else { 0 } } },
            @{ Expression = { if ($null -eq $_.AtingimentoAno) { [decimal]0 } else { $_.AtingimentoAno } }; Descending = $true },
            @{ Expression = { $_.Vendedor.Id } })
    $desligados = @($vendedores | Where-Object { -not $_.Vendedor.Ativo } | Sort-Object { $_.Vendedor.Id })

    # --- Item 2: ativos com atingimento do ano abaixo do corte. Gap = meta − realizado (§4.3).
    $abaixo = foreach ($r in $ranking) {
        if ($null -eq $r.AtingimentoAno -or $r.AtingimentoAno -ge $script:RelCorteAtingimento) { continue }
        $alvoCorte = [math]::Round($r.MetaAno * $script:RelCorteAtingimento, 2, [MidpointRounding]::AwayFromZero)
        [pscustomobject]@{ Linha = $r; Gap = $r.GapAno; FaltaCorte = $alvoCorte - $r.RealizadoAno }
    }
    $abaixo = @($abaixo | Sort-Object -Property @{ Expression = { $_.Linha.AtingimentoAno } }, @{ Expression = { $_.Linha.Vendedor.Id } })

    # --- Itens 3 e 4: foto do CRM, sem filtro de período (§7.3, §8).
    $funil = Get-VisaoFunil -Comercial $Comercial -Pipeline $Pipeline
    $atrasadas = @($funil.Abertas | Where-Object Atrasada | Sort-Object -Property @{ Expression = { $_.DiasAtraso }; Descending = $true }, @{ Expression = { $_.Oportunidade.Id } })
    $orfas = @($funil.Abertas | Where-Object Orfa)
    $valorOrfas = [decimal]0; foreach ($a in $orfas) { $valorOrfas += $a.Oportunidade.Valor }

    # --- Item 5a: faturamento por produto no acumulado do ano (§1: só faturadas; filial = filial da venda).
    $filiais = @($Estoque.Filiais | Sort-Object Id)
    $porProduto = @{}
    $faturamentoAno = [decimal]0
    foreach ($v in $Comercial.Vendas) {
        if (-not (Test-VendaRealizada $v) -or $v.Data.Year -ne $script:AnoMetas -or $v.Data.Month -gt $mesFinal) { continue }
        $faturamentoAno += $v.Valor
        if (-not $porProduto.ContainsKey($v.IdProduto)) {
            $pf = @{}; foreach ($f in $filiais) { $pf[$f.Id] = [decimal]0 }
            $porProduto[$v.IdProduto] = [pscustomobject]@{ IdProduto = $v.IdProduto; NomeVenda = $v.Produto; CategoriaVenda = $v.Categoria; Valor = [decimal]0; Unidades = 0; QtdVendas = 0; PorFilial = $pf }
        }
        $p = $porProduto[$v.IdProduto]
        $p.Valor += $v.Valor; $p.QtdVendas++
        $un = ConvertTo-Inteiro $v.Quantidade
        if ($null -ne $un) { $p.Unidades += $un }
        $p.PorFilial[$v.IdFilial] += $v.Valor
    }
    $estoquePorProduto = @{}
    foreach ($l in $Estoque.Linhas) {
        if (-not $estoquePorProduto.ContainsKey($l.Produto.Id)) { $estoquePorProduto[$l.Produto.Id] = 0 }
        $estoquePorProduto[$l.Produto.Id] += $l.Quantidade
    }
    $topProdutos = @($porProduto.Values | Sort-Object -Property @{ Expression = { $_.Valor }; Descending = $true }, @{ Expression = { $_.IdProduto } } | Select-Object -First $script:RelTopProdutos)
    $topProdutos = @(foreach ($p in $topProdutos) {
            $cat = $Estoque.ProdutoPorId[$p.IdProduto]   # §0.1: nome e categoria do catálogo, pelo ID
            [pscustomobject]@{
                IdProduto = $p.IdProduto
                Produto   = if ($cat) { $cat.Nome } else { $p.NomeVenda }
                Categoria = if ($cat) { $cat.Categoria } else { $p.CategoriaVenda }
                Valor     = $p.Valor
                Unidades  = $p.Unidades
                QtdVendas = $p.QtdVendas
                PorFilial = $p.PorFilial
                Estoque   = if ($estoquePorProduto.ContainsKey($p.IdProduto)) { $estoquePorProduto[$p.IdProduto] } else { 0 }
            }
        })
    $valorTop = [decimal]0; foreach ($p in $topProdutos) { $valorTop += $p.Valor }

    # --- Item 5b: estoque parado por filial (§5, prazo padrão, foto em vigor), com a demanda ao lado (§5.4).
    $visaoEstoque = Get-VisaoEstoque -Estoque $Estoque -DiasParado $script:DiasParadoPadrao -AbertasPorProduto (Get-AbertasPorProduto $Pipeline)
    $paradosPorFilial = foreach ($f in $filiais) {
        $ls = @($visaoEstoque.Parados | Where-Object { $_.Linha.IdFilial -eq $f.Id })
        $valor = [decimal]0; foreach ($p in $ls) { $valor += $p.Linha.Valor }
        [pscustomobject]@{ Filial = $f; Parados = $ls; Valor = $valor }
    }

    # Contexto para a análise (item 6): atingimento da empresa mês a mês (§4.2) e vínculo das vendas do mês com o CRM.
    $mensal = foreach ($m in $meses) { [pscustomobject]@{ Mes = $m; Atingimento = (Get-Desempenho -Base $Comercial -MesInicial $m -MesFinal $m).Empresa.Atingimento } }
    $mesComOp = 0
    foreach ($v in $Comercial.Vendas) { if ((Test-VendaRealizada $v) -and $v.Data.Year -eq $script:AnoMetas -and $v.Data.Month -eq $mesFinal -and $v.IdOportunidade) { $mesComOp++ } }

    $dbVendas = $Comercial.DataBase
    return [pscustomobject]@{
        ImpressaoDigital = Get-ImpressaoDigitalDados -Comercial $Comercial -Estoque $Estoque -DirDados $DirDados
        DataBaseVendas   = $dbVendas
        DataBaseCrm      = $Pipeline.DataBase
        DataBaseEstoque  = $Estoque.DataBase
        MesFinal         = $mesFinal
        RotuloAno        = if ($mesFinal -eq 1) { "jan/$($script:AnoMetas)" } else { "jan–$($script:RelMesesCurtos[$mesFinal])/$($script:AnoMetas)" }
        RotuloMes        = "$($script:RelNomesMes[$mesFinal])/$($script:AnoMetas)"
        MesParcial       = $mes.Parcial   # §4.5: "parcial até dd/mm"
        Ano              = $ano
        Mes              = $mes
        Ranking          = $ranking
        Desligados       = $desligados
        Corte            = $script:RelCorteAtingimento
        Abaixo           = $abaixo
        Funil            = $funil
        Atrasadas        = $atrasadas
        QtdOrfas         = $orfas.Count
        ValorOrfas       = $valorOrfas
        Filiais          = $filiais
        FaturamentoAno   = $faturamentoAno
        TopProdutos      = $topProdutos
        ValorTop         = $valorTop
        Estoque          = $visaoEstoque
        ParadosPorFilial = @($paradosPorFilial)
        Mensal           = @($mensal)
        MesComOportunidade = $mesComOp
    }
}

function Get-IndicadoresRelatorio($Rel) {
    # Catálogo de números que a análise (item 6) pode citar pela chave. Tudo sai do mesmo cálculo
    # dos itens 1 a 5; o PDF mostra o valor daqui, e não o que foi digitado no arquivo de riscos.
    $ind = [ordered]@{}
    $add = { param($Chave, $Tipo, $Valor, $Descricao) $ind[$Chave] = [pscustomobject]@{ Tipo = $Tipo; Valor = $Valor; Descricao = $Descricao } }
    $ano = $Rel.Ano.Empresa; $mes = $Rel.Mes.Empresa
    & $add 'empresa.ano.meta' 'dinheiro' $ano.Meta "Meta da empresa, $($Rel.RotuloAno) (§4.2)"
    & $add 'empresa.ano.realizado' 'dinheiro' $ano.RealizadoMesesMeta "Realizado da empresa nos meses com meta, $($Rel.RotuloAno)"
    & $add 'empresa.ano.atingimento' 'pct' $ano.Atingimento "Atingimento da empresa, $($Rel.RotuloAno) (§4.2)"
    & $add 'empresa.ano.gap' 'dinheiro' ($ano.Meta - $ano.RealizadoMesesMeta) "Quanto falta para a meta da empresa, $($Rel.RotuloAno) (§4.3)"
    & $add 'empresa.mes.meta' 'dinheiro' $mes.Meta "Meta da empresa em $($Rel.RotuloMes)"
    & $add 'empresa.mes.realizado' 'dinheiro' $mes.RealizadoMesesMeta "Realizado da empresa em $($Rel.RotuloMes)"
    & $add 'empresa.mes.gap' 'dinheiro' ($mes.Meta - $mes.RealizadoMesesMeta) "Quanto falta para a meta cheia da empresa em $($Rel.RotuloMes) (§4.3, §4.5)"
    & $add 'empresa.mes.atingimento' 'pct' $mes.Atingimento "Atingimento da empresa em $($Rel.RotuloMes)$(if ($Rel.MesParcial) { ", parcial até $($Rel.DataBaseVendas.ToString('dd/MM'))" })"
    & $add 'vendas.ano.canceladas.qtd' 'inteiro' $Rel.Ano.QtdCanceladas "Vendas canceladas, $($Rel.RotuloAno) (fora do realizado, §1)"
    & $add 'vendas.ano.canceladas.valor' 'dinheiro' $Rel.Ano.ValorCanceladas "Valor das vendas canceladas, $($Rel.RotuloAno)"
    & $add 'vendas.ano.ticket_medio' 'dinheiro' $ano.TicketMedio "Ticket médio, $($Rel.RotuloAno) (§1)"
    & $add 'vendas.mes.qtd' 'inteiro' $mes.QtdVendas "Vendas faturadas em $($Rel.RotuloMes)"
    & $add 'vendas.mes.canceladas.valor' 'dinheiro' $Rel.Mes.ValorCanceladas "Valor das vendas canceladas em $($Rel.RotuloMes)"
    & $add 'vendas.mes.com_oportunidade.qtd' 'inteiro' $Rel.MesComOportunidade "Vendas faturadas em $($Rel.RotuloMes) com ID Oportunidade preenchido (§0.1)"
    foreach ($x in $Rel.Mensal) {
        & $add ("empresa.mes_{0:00}.atingimento" -f $x.Mes) 'pct' $x.Atingimento "Atingimento da empresa em $($script:RelNomesMes[$x.Mes])/$($script:AnoMetas)"
    }
    foreach ($f in $Rel.Filiais) {
        foreach ($p in @(@{ Nome = 'ano'; D = $Rel.Ano; R = $Rel.RotuloAno }, @{ Nome = 'mes'; D = $Rel.Mes; R = $Rel.RotuloMes })) {
            $meta = [decimal]0; $real = [decimal]0
            foreach ($l in $p.D.Linhas) { if ($l.Vendedor.IdFilial -eq $f.Id) { $meta += $l.Meta; $real += $l.RealizadoMesesMeta } }
            & $add "filial.$($f.Id).$($p.Nome).atingimento" 'pct' $(if ($meta -gt 0) { $real / $meta } else { $null }) "Atingimento da filial $($f.Nome), $($p.R) (§4.2, com desligados)"
        }
    }
    $gapAbaixo = [decimal]0; foreach ($a in $Rel.Abaixo) { $gapAbaixo += $a.Gap }
    & $add 'item2.qtd' 'inteiro' $Rel.Abaixo.Count "Vendedores ativos abaixo de 80% da meta, $($Rel.RotuloAno)"
    & $add 'item2.gap_total' 'dinheiro' $gapAbaixo "Soma do que falta para a meta dos vendedores abaixo de 80%"
    foreach ($r in @($Rel.Ranking) + @($Rel.Desligados)) {
        $id = $r.Vendedor.Id
        & $add "vendedor.$id.ano.atingimento" 'pct' $r.AtingimentoAno "$($r.Vendedor.Nome): atingimento, $($Rel.RotuloAno)"
        & $add "vendedor.$id.ano.gap" 'dinheiro' $r.GapAno "$($r.Vendedor.Nome): meta − realizado, $($Rel.RotuloAno) (§4.3)"
        & $add "vendedor.$id.mes.atingimento" 'pct' $r.AtingimentoMes "$($r.Vendedor.Nome): atingimento em $($Rel.RotuloMes)"
        $ab = @($Rel.Funil.Abertas | Where-Object { $_.Vendedor.Id -eq $id })
        $v = [decimal]0; $pond = [decimal]0; $atr = [decimal]0
        foreach ($a in $ab) { $v += $a.Oportunidade.Valor; $pond += $a.Ponderado; if ($a.Atrasada) { $atr += $a.Oportunidade.Valor } }
        & $add "vendedor.$id.pipeline.valor" 'dinheiro' $v "$($r.Vendedor.Nome): pipeline aberto (§7)"
        & $add "vendedor.$id.pipeline.ponderado" 'dinheiro' $pond "$($r.Vendedor.Nome): pipeline ponderado (§7)"
        & $add "vendedor.$id.pipeline.atrasado" 'dinheiro' $atr "$($r.Vendedor.Nome): pipeline atrasado (§7)"
    }
    $t = $Rel.Funil.Total
    & $add 'pipeline.abertas.qtd' 'inteiro' $t.Qtd 'Oportunidades abertas (§7)'
    & $add 'pipeline.abertas.valor' 'dinheiro' $t.Valor 'Pipeline aberto, valor bruto (§7)'
    & $add 'pipeline.abertas.ponderado' 'dinheiro' $t.Ponderado 'Pipeline ponderado (§7)'
    & $add 'pipeline.atrasadas.qtd' 'inteiro' $t.QtdAtrasadas 'Oportunidades abertas com previsão vencida (§7)'
    & $add 'pipeline.atrasadas.valor' 'dinheiro' $t.ValorAtrasado 'Valor bruto das oportunidades com previsão vencida (§7)'
    & $add 'pipeline.atrasadas.pct_valor' 'pct' $(if ($t.Valor -gt 0) { $t.ValorAtrasado / $t.Valor } else { $null }) 'Parte do pipeline bruto com previsão vencida'
    $maiorAtraso = if ($Rel.Atrasadas.Count) { $Rel.Atrasadas[0].DiasAtraso } else { 0 }
    & $add 'pipeline.atrasadas.maior_atraso_dias' 'inteiro' $maiorAtraso 'Maior atraso entre as oportunidades abertas, em dias'
    $mais90 = @($Rel.Atrasadas | Where-Object { $_.DiasAtraso -gt 90 })
    $v90 = [decimal]0; foreach ($a in $mais90) { $v90 += $a.Oportunidade.Valor }
    & $add 'pipeline.atrasadas.mais_90_dias.qtd' 'inteiro' $mais90.Count 'Oportunidades com previsão vencida há mais de 90 dias'
    & $add 'pipeline.atrasadas.mais_90_dias.valor' 'dinheiro' $v90 'Valor das oportunidades com previsão vencida há mais de 90 dias'
    & $add 'pipeline.orfas.qtd' 'inteiro' $Rel.QtdOrfas 'Oportunidades abertas de vendedor desligado (órfãs, §2.4)'
    & $add 'pipeline.orfas.valor' 'dinheiro' $Rel.ValorOrfas 'Valor das oportunidades órfãs (§2.4)'
    & $add 'pipeline.conversao_qtd' 'pct' $Rel.Funil.Conversao 'Conversão em quantidade: ganhas ÷ (ganhas + perdidas) (§8)'
    & $add 'pipeline.conversao_valor' 'pct' $Rel.Funil.ConversaoValor 'Conversão em valor (§8.2)'
    foreach ($e in $Rel.Funil.Etapas) {
        $k = ($e.Etapa.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^A-Za-z]+', '_').ToLowerInvariant().Trim('_')
        & $add "pipeline.etapa.$k.valor" 'dinheiro' $e.Valor "Pipeline bruto na etapa $($e.Etapa)"
        & $add "pipeline.etapa.$k.ponderado" 'dinheiro' $e.Ponderado "Pipeline ponderado na etapa $($e.Etapa)"
        & $add "pipeline.etapa.$k.atrasado" 'dinheiro' $e.ValorAtrasado "Pipeline com previsão vencida na etapa $($e.Etapa)"
    }
    # Concentração do pipeline aberto por cliente (pelo ID, §0.1).
    $porCliente = @{}
    foreach ($a in $Rel.Funil.Abertas) {
        $id = $a.Cliente.Id
        if (-not $porCliente.ContainsKey($id)) { $porCliente[$id] = [pscustomobject]@{ Cliente = $a.Cliente; Valor = [decimal]0 } }
        $porCliente[$id].Valor += $a.Oportunidade.Valor
    }
    $clientes = @($porCliente.Values | Sort-Object -Property @{ Expression = { $_.Valor }; Descending = $true }, @{ Expression = { $_.Cliente.Id } })
    if ($clientes.Count) {
        $top3 = [decimal]0; foreach ($c in ($clientes | Select-Object -First 3)) { $top3 += $c.Valor }
        & $add 'pipeline.maior_cliente.nome' 'texto' "$($clientes[0].Cliente.Nome) ($($clientes[0].Cliente.Id))" 'Cliente com o maior pipeline aberto'
        & $add 'pipeline.maior_cliente.valor' 'dinheiro' $clientes[0].Valor 'Pipeline aberto do maior cliente'
        & $add 'pipeline.maior_cliente.pct' 'pct' $(if ($t.Valor -gt 0) { $clientes[0].Valor / $t.Valor } else { $null }) 'Parte do pipeline aberto no maior cliente'
        & $add 'pipeline.top3_clientes.pct' 'pct' $(if ($t.Valor -gt 0) { $top3 / $t.Valor } else { $null }) 'Parte do pipeline aberto nos três maiores clientes'
    }
    & $add 'produtos.ano.faturamento' 'dinheiro' $Rel.FaturamentoAno "Faturamento total, $($Rel.RotuloAno) (§1)"
    & $add 'produtos.top10.pct' 'pct' $(if ($Rel.FaturamentoAno -gt 0) { $Rel.ValorTop / $Rel.FaturamentoAno } else { $null }) 'Parte do faturamento nos dez produtos de maior saída'
    $semEstoqueTop = @($Rel.TopProdutos | Where-Object { $_.Estoque -eq 0 })
    & $add 'produtos.top10.sem_estoque.qtd' 'inteiro' $semEstoqueTop.Count 'Produtos do top 10 sem estoque nas três filiais (§5)'
    $e = $Rel.Estoque
    & $add 'estoque.imobilizado' 'dinheiro' $e.Total.Valor 'Valor imobilizado em estoque (§5)'
    & $add 'estoque.parado.qtd_linhas' 'inteiro' $e.Parados.Count "Linhas de estoque paradas, prazo de $($e.DiasParado) dias (§5)"
    & $add 'estoque.parado.valor' 'dinheiro' $e.Total.Parado 'Valor imobilizado nas linhas paradas (§5)'
    & $add 'estoque.parado.pct' 'pct' $(if ($e.Total.Valor -gt 0) { $e.Total.Parado / $e.Total.Valor } else { $null }) 'Parte do valor imobilizado que está parada'
    foreach ($pf in $Rel.ParadosPorFilial) { & $add "estoque.parado.$($pf.Filial.Id).valor" 'dinheiro' $pf.Valor "Estoque parado em $($pf.Filial.Nome)" }
    $paradoSemOp = @($e.Parados | Where-Object { $_.Abertas -and $_.Abertas.Qtd -eq 0 })
    $vSemOp = [decimal]0; foreach ($p in $paradoSemOp) { $vSemOp += $p.Linha.Valor }
    & $add 'estoque.parado_sem_oportunidade.valor' 'dinheiro' $vSemOp 'Estoque parado de produtos sem nenhuma oportunidade aberta (§5.4)'
    & $add 'estoque.abaixo_minimo.qtd' 'inteiro' $e.Abaixo.Count 'Linhas de estoque abaixo do mínimo (§5)'
    & $add 'estoque.sem_estoque.qtd' 'inteiro' $e.SemEstoque.Count 'Produtos sem estoque nas três filiais (§5)'
    $semComDemanda = @($e.SemEstoque | Where-Object { $_.Abertas -and $_.Abertas.Qtd -gt 0 })
    $vDem = [decimal]0; foreach ($s in $semComDemanda) { $vDem += $s.Abertas.Valor }
    & $add 'estoque.sem_estoque_com_oportunidade.qtd' 'inteiro' $semComDemanda.Count 'Produtos sem estoque que têm oportunidade aberta'
    & $add 'estoque.sem_estoque_com_oportunidade.valor' 'dinheiro' $vDem 'Pipeline aberto de produtos sem estoque'
    & $add 'datas_base.defasagem_crm_dias' 'inteiro' ($Rel.DataBaseVendas - $Rel.DataBaseCrm).Days 'Dias entre a data-base de vendas e a foto do CRM'
    & $add 'datas_base.defasagem_estoque_dias' 'inteiro' ($Rel.DataBaseVendas - $Rel.DataBaseEstoque).Days 'Dias entre a data-base de vendas e a foto de estoque'
    return $ind
}

function ConvertTo-RelValorJson($Tipo, $Valor) {
    if ($null -eq $Valor) { return $null }
    switch ($Tipo) {
        'dinheiro' { return [decimal]::Round([decimal]$Valor, 2, [MidpointRounding]::AwayFromZero) }
        'pct' { return [decimal]::Round([decimal]$Valor, 6, [MidpointRounding]::AwayFromZero) }
        default { return $Valor }
    }
}

function ConvertTo-RelatorioJson($Rel, $Indicadores) {
    # Os itens 1 a 5 e o catálogo, em ordem fixa: é o que a análise lê e o que o teste compara.
    $d = { param($x) if ($x) { $x.ToString('yyyy-MM-dd') } }
    $m = { param($x) ConvertTo-RelValorJson 'dinheiro' $x }
    $p = { param($x) ConvertTo-RelValorJson 'pct' $x }
    $obj = [ordered]@{
        relatorio         = 'Relatório semanal da diretoria — Horizonte Máquinas'
        impressao_digital = $Rel.ImpressaoDigital
        datas_base        = [ordered]@{ vendas = & $d $Rel.DataBaseVendas; crm = & $d $Rel.DataBaseCrm; estoque = & $d $Rel.DataBaseEstoque }
        periodo           = [ordered]@{ acumulado = $Rel.RotuloAno; mes = $Rel.RotuloMes; mes_parcial = [bool]$Rel.MesParcial }
        item1_ranking     = [ordered]@{
            vendedores = @(foreach ($r in $Rel.Ranking) {
                    [ordered]@{
                        id = $r.Vendedor.Id; nome = $r.Vendedor.Nome; filial = $r.Vendedor.Filial
                        admissao = & $d $r.Vendedor.Admissao; meses_com_meta = @($r.MesesComMeta)
                        ano = [ordered]@{ meta = & $m $r.MetaAno; realizado = & $m $r.RealizadoAno; atingimento = & $p $r.AtingimentoAno }
                        mes = [ordered]@{ meta = & $m $r.MetaMes; realizado = & $m $r.RealizadoMes; atingimento = & $p $r.AtingimentoMes }
                    }
                })
            empresa    = [ordered]@{
                ano = [ordered]@{ meta = & $m $Rel.Ano.Empresa.Meta; realizado = & $m $Rel.Ano.Empresa.RealizadoMesesMeta; atingimento = & $p $Rel.Ano.Empresa.Atingimento }
                mes = [ordered]@{ meta = & $m $Rel.Mes.Empresa.Meta; realizado = & $m $Rel.Mes.Empresa.RealizadoMesesMeta; atingimento = & $p $Rel.Mes.Empresa.Atingimento }
            }
            fora_do_ranking = @(foreach ($r in $Rel.Desligados) { [ordered]@{ id = $r.Vendedor.Id; nome = $r.Vendedor.Nome; desligamento = & $d $r.Vendedor.Desligamento } })
        }
        item2_abaixo_80   = @(foreach ($a in $Rel.Abaixo) {
                [ordered]@{ id = $a.Linha.Vendedor.Id; nome = $a.Linha.Vendedor.Nome; atingimento = & $p $a.Linha.AtingimentoAno; falta_meta = & $m $a.Gap; falta_80 = & $m $a.FaltaCorte }
            })
        item3_pipeline    = [ordered]@{
            etapas = @(foreach ($e in $Rel.Funil.Etapas) {
                    [ordered]@{ etapa = $e.Etapa; probabilidade = $e.Probabilidade; qtd = $e.Qtd; bruto = & $m $e.Valor; ponderado = & $m $e.Ponderado; atrasadas = $e.QtdAtrasadas; atrasado = & $m $e.ValorAtrasado }
                })
            total  = [ordered]@{ qtd = $Rel.Funil.Total.Qtd; bruto = & $m $Rel.Funil.Total.Valor; ponderado = & $m $Rel.Funil.Total.Ponderado; atrasadas = $Rel.Funil.Total.QtdAtrasadas; atrasado = & $m $Rel.Funil.Total.ValorAtrasado }
            orfas  = [ordered]@{ qtd = $Rel.QtdOrfas; valor = & $m $Rel.ValorOrfas }
        }
        item4_atrasadas   = @(foreach ($a in $Rel.Atrasadas) {
                [ordered]@{
                    id = $a.Oportunidade.Id; vendedor = $a.Vendedor.Id; orfa = [bool]$a.Orfa; cliente = $a.Cliente.Id; produto = $a.Oportunidade.IdProduto
                    etapa = $a.Oportunidade.Etapa; valor = & $m $a.Oportunidade.Valor; ponderado = & $m $a.Ponderado; previsao = & $d $a.Oportunidade.Previsao; dias_atraso = $a.DiasAtraso
                }
            })
        item5_produtos    = [ordered]@{
            faturamento_periodo = & $m $Rel.FaturamentoAno
            top10               = @(foreach ($t in $Rel.TopProdutos) {
                    $pf = [ordered]@{}; foreach ($f in $Rel.Filiais) { $pf[$f.Id] = & $m $t.PorFilial[$f.Id] }
                    [ordered]@{ id = $t.IdProduto; produto = $t.Produto; categoria = $t.Categoria; faturamento = & $m $t.Valor; unidades = $t.Unidades; por_filial = $pf; estoque_atual = $t.Estoque }
                })
        }
        item5_parados     = [ordered]@{
            prazo_dias = $Rel.Estoque.DiasParado
            filiais    = @(foreach ($pf in $Rel.ParadosPorFilial) {
                    [ordered]@{
                        filial = $pf.Filial.Id; nome = $pf.Filial.Nome; valor = & $m $pf.Valor
                        linhas = @(foreach ($x in $pf.Parados) {
                                [ordered]@{ produto = $x.Linha.Produto.Id; quantidade = $x.Linha.Quantidade; valor = & $m $x.Linha.Valor; dias_sem_movimento = $x.Linha.DiasSemMovimento; oportunidades_abertas = $(if ($x.Abertas) { $x.Abertas.Qtd }) }
                            })
                    }
                })
            total      = & $m $Rel.Estoque.Total.Parado
        }
        indicadores       = [ordered]@{}
    }
    foreach ($k in $Indicadores.Keys) {
        $i = $Indicadores[$k]
        $obj.indicadores[$k] = [ordered]@{ tipo = $i.Tipo; valor = ConvertTo-RelValorJson $i.Tipo $i.Valor; descricao = $i.Descricao }
    }
    $json = $obj | ConvertTo-Json -Depth 12
    return ($json -replace "`r`n", "`n")
}

function Read-RiscosRelatorio {
    # Item 6: lê e valida o arquivo de riscos escrito pelo agente. Devolve Erros e Riscos.
    # Regras do formato: exatamente 3 riscos; cada um com título, leitura e um número que o sustenta,
    # dado pela chave de um indicador do catálogo OU por número + como foi calculado.
    param([Parameter(Mandatory)][string]$Caminho, [Parameter(Mandatory)]$Indicadores, [Parameter(Mandatory)][string]$ImpressaoDigital)
    $erros = New-Object Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $Caminho)) { return [pscustomobject]@{ Erros = @("Arquivo de riscos não encontrado: $Caminho"); Riscos = @() } }
    try { $dados = [IO.File]::ReadAllText($Caminho, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { return [pscustomobject]@{ Erros = @("Arquivo de riscos não é um JSON válido: $($_.Exception.Message)"); Riscos = @() } }
    if ([string]$dados.impressao_digital -ne $ImpressaoDigital) {
        $erros.Add("Os riscos foram escritos para outros dados (impressão digital '$($dados.impressao_digital)'; a atual é '$ImpressaoDigital'). Refaça a análise sobre o calculo.json desta semana.")
    }
    $lista = @($dados.riscos | Where-Object { $null -ne $_ })
    if ($lista.Count -ne 3) { $erros.Add("São exigidos exatamente 3 riscos; o arquivo tem $($lista.Count).") }
    $riscos = New-Object Collections.Generic.List[object]
    $n = 0
    foreach ($r in $lista) {
        $n++
        $titulo = ([string]$r.titulo).Trim(); $leitura = ([string]$r.leitura).Trim()
        $chave = ([string]$r.indicador).Trim(); $numero = ([string]$r.numero).Trim(); $calculo = ([string]$r.calculo).Trim()
        if (-not $titulo) { $erros.Add("Risco ${n}: falta o título.") } elseif ($titulo.Length -gt 90) { $erros.Add("Risco ${n}: título com mais de 90 caracteres.") }
        if (-not $leitura) { $erros.Add("Risco ${n}: falta a leitura (o que o número significa e por que é risco).") } elseif ($leitura.Length -gt 700) { $erros.Add("Risco ${n}: leitura com mais de 700 caracteres.") }
        $ind = $null
        if ($chave) {
            if (-not $Indicadores.Contains($chave)) { $erros.Add("Risco ${n}: indicador '$chave' não existe no catálogo do calculo.json.") }
            else { $ind = $Indicadores[$chave] }
        } elseif (-not $numero -or -not $calculo) {
            $erros.Add("Risco ${n}: sem o número que o sustenta. Informe 'indicador' (chave do catálogo) ou 'numero' e 'calculo'.")
        }
        $riscos.Add([pscustomobject]@{ Titulo = $titulo; Leitura = $leitura; Chave = $chave; Indicador = $ind; Numero = $numero; Calculo = $calculo })
    }
    return [pscustomobject]@{ Erros = $erros.ToArray(); Riscos = $riscos.ToArray() }
}
