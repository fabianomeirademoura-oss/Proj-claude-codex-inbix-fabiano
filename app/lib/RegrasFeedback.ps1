# Regras do feedback semanal ao vendedor — implementação do REGRAS_NEGOCIO.md §11.
# Único lugar que decide ritmo do mês, média mensal, projeção de fim de ano, necessário
# por mês e cobertura do ponderado. Meta, realizado e atingimento vêm de Regras.ps1 (§4);
# pipeline e conversão vêm de RegrasPipeline.ps1 (§7, §8). Quem redige a mensagem só formata.

# §11.6: faixas do sinal, pela projeção de fim de ano ÷ meta do ano.
$script:FaixaNoRitmo = [decimal]1.00
$script:FaixaPoucoAbaixo = [decimal]0.90

function Get-FeedbackSemanal {
    param(
        [Parameter(Mandatory)]$Comercial,   # Import-BaseComercial, já sem erros
        [Parameter(Mandatory)]$Pipeline,    # Import-BasePipeline, já sem erros
        [string]$IdVendedor                 # vazio = todos os ativos
    )
    $db = $Comercial.DataBase
    if (-not $db -or $db.Year -ne $script:AnoMetas) { throw "Data-base de vendas fora de $($script:AnoMetas)." }

    # §4.5: mês da data-base em andamento => parcial até a data-base.
    $emAndamento = ($db.AddDays(1).Month -eq $db.Month)
    $mesAtual = $db.Month
    $ultimoFechado = if ($emAndamento) { $mesAtual - 1 } else { $mesAtual }
    $diasNoMes = [datetime]::DaysInMonth($db.Year, $mesAtual)

    # §11.2: semana = 7 dias corridos terminando na data-base; semana anterior = os 7 dias antes dela.
    $inicioSemana = $db.AddDays(-6)
    $inicioAnterior = $db.AddDays(-13)
    $fimAnterior = $db.AddDays(-7)

    $metaMes = @{}   # IdVendedor -> @{ mês -> meta }
    foreach ($m in $Comercial.Metas) {
        if ($m.Ano -ne $script:AnoMetas) { continue }
        if (-not $metaMes.ContainsKey($m.IdVendedor)) { $metaMes[$m.IdVendedor] = @{} }
        $metaMes[$m.IdVendedor][$m.Mes] = $m.Valor
    }
    $realMes = @{}   # IdVendedor -> @{ mês -> realizado }  (§1: só faturadas)
    $vendasSemana = @{}
    $realAnterior = @{}
    foreach ($v in $Comercial.Vendas) {
        if (-not (Test-VendaRealizada $v) -or $v.Data.Year -ne $script:AnoMetas) { continue }
        if (-not $realMes.ContainsKey($v.IdVendedor)) { $realMes[$v.IdVendedor] = @{} }
        $realMes[$v.IdVendedor][$v.Data.Month] = [decimal]$realMes[$v.IdVendedor][$v.Data.Month] + $v.Valor
        if ($v.Data -ge $inicioSemana -and $v.Data -le $db) {
            if (-not $vendasSemana.ContainsKey($v.IdVendedor)) { $vendasSemana[$v.IdVendedor] = New-Object Collections.Generic.List[object] }
            $vendasSemana[$v.IdVendedor].Add($v)
        }
        elseif ($v.Data -ge $inicioAnterior -and $v.Data -le $fimAnterior) {
            $realAnterior[$v.IdVendedor] = [decimal]$realAnterior[$v.IdVendedor] + $v.Valor
        }
    }

    # §4.1 nos meses fechados: reaproveita o cálculo auditado do painel.
    $fechados = @{}
    if ($ultimoFechado -ge 1) {
        foreach ($l in (Get-Desempenho -Base $Comercial -MesInicial 1 -MesFinal $ultimoFechado).Linhas) { $fechados[$l.Vendedor.Id] = $l }
    }

    # §11.1: só vendedores ativos recebem feedback (§2.3).
    $alvos = @($Comercial.Vendedores | Where-Object { $_.Ativo -and (-not $IdVendedor -or $_.Id -eq $IdVendedor) } | Sort-Object Id)
    if ($IdVendedor -and -not $alvos.Count) { throw "Vendedor $IdVendedor não existe ou não está ativo." }

    $lista = foreach ($vend in $alvos) {
        $metas = if ($metaMes.ContainsKey($vend.Id)) { $metaMes[$vend.Id] } else { @{} }
        $reais = if ($realMes.ContainsKey($vend.Id)) { $realMes[$vend.Id] } else { @{} }

        # Mês da data-base (§4.5: meta cheia, sem pró-rata; o ritmo é indicador separado, §11.3).
        $metaAtual = if ($metas.ContainsKey($mesAtual)) { $metas[$mesAtual] } else { $null }
        $realAtual = [decimal]$reais[$mesAtual]
        $mes = [pscustomobject]@{
            Mes              = $mesAtual
            Parcial          = $emAndamento
            DiasDecorridos   = $db.Day
            DiasNoMes        = $diasNoMes
            FracaoDoMes      = [decimal]$db.Day / $diasNoMes
            Meta             = $metaAtual
            Realizado        = $realAtual
            Atingimento      = if ($metaAtual -gt 0) { $realAtual / $metaAtual } else { $null }
            Falta            = if ($metaAtual -gt 0) { [math]::Max([decimal]0, $metaAtual - $realAtual) } else { $null }
            # §11.3: atingimento do mês comparado com a fração de dias corridos; não é pró-rata da meta.
            NoRitmo          = if ($metaAtual -gt 0) { ($realAtual / $metaAtual) -ge ([decimal]$db.Day / $diasNoMes) } else { $null }
        }

        # Semana (§11.2).
        $semana = @(if ($vendasSemana.ContainsKey($vend.Id)) { $vendasSemana[$vend.Id] | Sort-Object Data, Id })
        $realSemana = [decimal]0
        foreach ($v in $semana) { $realSemana += $v.Valor }

        # Meses fechados com meta (§4.1) e média mensal (§11.4).
        $f = $fechados[$vend.Id]
        $nFechados = if ($f) { @($f.MesesComMeta).Count } else { 0 }
        $realFechados = if ($f) { $f.RealizadoMesesMeta } else { [decimal]0 }
        $media = if ($nFechados) { [math]::Round($realFechados / $nFechados, 2, [MidpointRounding]::AwayFromZero) } else { $null }

        # Ano (§11.4): só meses com meta, como no atingimento.
        $mesesAno = @($metas.Keys | Sort-Object)
        $metaAno = [decimal]0
        foreach ($k in $mesesAno) { $metaAno += $metas[$k] }
        $realAno = [decimal]0
        foreach ($k in $mesesAno) { if ($k -le $mesAtual) { $realAno += [decimal]$reais[$k] } }
        $restantes = @($mesesAno | Where-Object { $_ -gt $ultimoFechado })
        $metaRestante = [decimal]0
        foreach ($k in $restantes) { $metaRestante += $metas[$k] }
        $projecao = if ($null -ne $media) { $media * $mesesAno.Count } else { $null }
        $necessario = if ($restantes.Count) { [math]::Round([math]::Max([decimal]0, $metaAno - $realFechados) / $restantes.Count, 2, [MidpointRounding]::AwayFromZero) } else { $null }
        $faltaAno = [math]::Max([decimal]0, $metaAno - $realAno)
        $pctProj = if ($metaAno -gt 0 -and $null -ne $projecao) { $projecao / $metaAno } else { $null }
        $sinal = if ($null -eq $pctProj) { 'sem histórico' }
        elseif ($pctProj -ge $script:FaixaNoRitmo) { 'no ritmo' }
        elseif ($pctProj -ge $script:FaixaPoucoAbaixo) { 'um pouco abaixo' }
        else { 'abaixo' }

        # Pipeline e conversão do vendedor (§7, §8), foto do CRM.
        $funil = Get-VisaoFunil -Comercial $Comercial -Pipeline $Pipeline -IdVendedor $vend.Id
        $ops = foreach ($a in ($funil.Abertas | Sort-Object -Property @{ Expression = { $_.Ponderado }; Descending = $true }, @{ Expression = { $_.Oportunidade.Id } })) {
            [pscustomobject]@{
                Id            = $a.Oportunidade.Id
                Cliente       = $a.Cliente.Nome
                Produto       = $a.Produto.Nome
                Etapa         = $a.Oportunidade.Etapa
                Probabilidade = $a.Oportunidade.Probabilidade
                Valor         = $a.Oportunidade.Valor
                Ponderado     = $a.Ponderado
                Previsao      = $a.Oportunidade.Previsao
                Atrasada      = $a.Atrasada
                DiasAtraso    = $a.DiasAtraso
            }
        }
        $ops = @($ops)

        [pscustomobject]@{
            Vendedor   = $vend
            Admitida   = ($vend.Admissao -and $vend.Admissao.Year -eq $script:AnoMetas)   # §3: ramp-up
            Semana     = [pscustomobject]@{
                Inicio            = $inicioSemana
                Fim               = $db
                Realizado         = $realSemana
                QtdVendas         = $semana.Count
                Vendas            = $semana
                RealizadoAnterior = [decimal]$realAnterior[$vend.Id]
            }
            MesAtual   = $mes
            Fechados   = [pscustomobject]@{
                Meses       = @(if ($f) { $f.MesesComMeta })
                Meta        = if ($f) { $f.Meta } else { [decimal]0 }
                Realizado   = $realFechados
                Atingimento = if ($f) { $f.Atingimento } else { $null }
                MediaMensal = $media
            }
            Ano        = [pscustomobject]@{
                MesesComMeta        = $mesesAno
                Meta                = $metaAno
                Realizado           = $realAno
                Atingimento         = if ($metaAno -gt 0) { $realAno / $metaAno } else { $null }
                Falta               = $faltaAno
                Projecao            = $projecao
                ProjecaoPct         = $pctProj
                ProjecaoFaltaria    = if ($null -ne $projecao -and $projecao -lt $metaAno) { $metaAno - $projecao } else { $null }
                ProjecaoSobra       = if ($null -ne $projecao -and $projecao -ge $metaAno) { $projecao - $metaAno } else { $null }
                MesesRestantes      = $restantes
                MetaMediaRestante   = if ($restantes.Count) { [math]::Round($metaRestante / $restantes.Count, 2, [MidpointRounding]::AwayFromZero) } else { $null }
                NecessarioPorMes    = $necessario
                Sinal               = $sinal
            }
            Pipeline   = [pscustomobject]@{
                DataBaseCrm      = $Pipeline.DataBase
                QtdAbertas       = $funil.Total.Qtd
                Valor            = $funil.Total.Valor
                Ponderado        = $funil.Total.Ponderado
                QtdAtrasadas     = $funil.Total.QtdAtrasadas
                ValorAtrasado    = $funil.Total.ValorAtrasado
                CoberturaFalta   = if ($faltaAno -gt 0) { $funil.Total.Ponderado / $faltaAno } else { $null }
                CenarioPonderado = $realAno + $funil.Total.Ponderado
                CenarioPct       = if ($metaAno -gt 0) { ($realAno + $funil.Total.Ponderado) / $metaAno } else { $null }
                Ganhas           = $funil.Ganhas.Qtd
                Perdidas         = $funil.Perdidas.Qtd
                Conversao        = $funil.Conversao
                PrincipalMotivo  = if ($funil.Motivos.Count) { $funil.Motivos[0].Motivo } else { $null }
                Oportunidades    = $ops
            }
        }
    }

    return [pscustomobject]@{
        DataBaseVendas = $db
        DataBaseCrm    = $Pipeline.DataBase
        CrmDefasado    = ($Pipeline.DataBase -lt $db)   # §11.5
        MesParcial     = $emAndamento
        Vendedores     = @($lista)
    }
}
