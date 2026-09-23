# HTML e CSV da visão de pipeline. Só formata: os números vêm de RegrasPipeline.ps1.
# Usa os auxiliares de Paginas.ps1 (Esc, Format-*, Get-Marcas, New-Layout).

function Format-Prob($Fracao) { return ([decimal]$Fracao * 100).ToString('N0', $script:PtBr) + '%' }

function New-PaginaPipeline($Base, $Visao, $Usuario, [string]$FiltroVendedor) {
    $vis = $Visao
    $d = $vis.Desempenho
    $t = $vis.Total
    $q = Get-QueryPeriodo $d $false
    $meses = Get-MesesDisponiveis $Base
    $opcoes = {
        param($selecionado)
        ($meses | ForEach-Object { $s = if ($_ -eq $selecionado) { ' selected' } else { '' }; "<option value=""$_""$s>$($script:NomesMesLongo[$_])</option>" }) -join ''
    }
    $periodo = "$($script:NomesMesLongo[$d.MesInicial]) a $($script:NomesMesLongo[$d.MesFinal]) de 2026"
    $dbCrm = Format-Data $vis.DataBaseCrm
    $euId = $Usuario.idVendedor

    $parcial = if ($d.Parcial) {
        "<p class=""parcial"">Meta e realizado parciais até $(Format-Data $d.DataBase): $($script:NomesMesLongo[$d.MesFinal]) ainda não terminou e é comparado com a meta cheia do mês.</p>"
    } else { '' }

    $orfas = @($vis.Abertas | Where-Object Orfa)
    $avisoOrfas = if ($orfas.Count) {
        $donos = ($orfas | Group-Object { $_.Vendedor.Id } | ForEach-Object {
                $v = $_.Group[0].Vendedor
                "$(Esc $v.Nome) (desligado em $(Format-Data $v.Desligamento)): $(($_.Group | ForEach-Object { Esc $_.Oportunidade.Id }) -join ', ')"
            }) -join ' · '
        "<p class=""aviso-orfa"" role=""alert""><strong>$($orfas.Count) oportunidade(s) órfã(s), $(Format-Moeda $vis.ValorOrfas).</strong> $donos. Continuam no pipeline e nos totais; a reatribuição é uma ação do gestor (REGRAS_NEGOCIO.md §2.4).</p>"
    } else { '' }

    $temFora = $t.ForaMesesMeta -ne 0
    $notaFora = if ($temFora) { "<p class=""nota"">$(Format-Moeda $t.ForaMesesMeta) foram vendidos em meses sem meta do vendedor: ficam fora do realizado desta tabela (REGRAS_NEGOCIO.md §4.1).</p>" } else { '' }

    # Uma linha por vendedor
    $linhasVend = foreach ($l in $vis.Linhas) {
        $v = $l.Vendedor
        $cls = @()
        if ($v.Id -eq $euId) { $cls += 'voce' }
        if ($FiltroVendedor -eq $v.Id) { $cls += 'selecionado' }
        $attr = if ($cls) { " class=""$($cls -join ' ')""" } else { '' }
        $atr = if ($l.QtdAtrasadas) { "<span class=""marca atraso"">$($l.QtdAtrasadas)</span>" } else { '0' }
        $gap = if ($null -eq $l.Gap) { '—' } elseif ($l.Gap -le 0) { "<span class=""nota"">meta batida</span>" } else { Format-Moeda $l.Gap }
        $orfa = if (-not $v.Ativo -and $l.QtdAbertas) { ' <span class="marca atraso">carteira órfã</span>' } else { '' }
        "<tr$attr><td><a href=""/pipeline?$q&amp;vendedor=$(Esc $v.Id)#oportunidades"">$(Esc $v.Nome)</a> $(Get-Marcas $v)$orfa</td><td>$(Esc $v.Filial)</td><td class=""num"">$(Format-Moeda $l.Meta)</td><td class=""num"">$(Format-Moeda $l.RealizadoMesesMeta)</td><td class=""num"">$(Format-Pct $l.Atingimento)</td><td class=""num"">$gap</td><td class=""num"">$($l.QtdAbertas)</td><td class=""num"">$atr</td><td class=""num"">$(Format-Moeda $l.Pipeline)</td><td class=""num forte"">$(Format-Moeda $l.Ponderado)</td></tr>"
    }
    $gapTotal = if ($null -eq $t.Gap) { '—' } elseif ($t.Gap -le 0) { 'meta batida' } else { Format-Moeda $t.Gap }

    # Lista de oportunidades abertas (atrasadas em destaque)
    $lista = @($vis.Abertas | Where-Object { -not $FiltroVendedor -or $_.Vendedor.Id -eq $FiltroVendedor })
    $nomeFiltro = if ($FiltroVendedor) { ($vis.Linhas | Where-Object { $_.Vendedor.Id -eq $FiltroVendedor } | Select-Object -First 1).Vendedor.Nome }
    $tituloLista = if ($nomeFiltro) { "Oportunidades abertas de $(Esc $nomeFiltro) <a class=""nota"" href=""/pipeline?$q#oportunidades"">ver todas</a>" } else { 'Oportunidades abertas' }
    $qtdAtrLista = @($lista | Where-Object Atrasada).Count
    $linhasOp = foreach ($a in $lista) {
        $o = $a.Oportunidade; $v = $a.Vendedor; $c = $a.Cliente
        $cls = @()
        if ($a.Atrasada) { $cls += 'atrasada' }
        if ($v.Id -eq $euId) { $cls += 'voce' }
        $attr = if ($cls) { " class=""$($cls -join ' ')""" } else { '' }
        $sit = if ($a.Atrasada) { "<span class=""marca atraso"">atrasada há $($a.DiasAtraso) dias</span>" } else { '<span class="nota">no prazo</span>' }
        $orfa = if ($a.Orfa) { ' <span class="marca desligado">órfã</span>' } else { '' }
        "<tr$attr><td class=""nw"">$(Esc $o.Id)</td><td>$(Esc $v.Nome)$orfa</td><td>$(Esc $c.Nome) <span class=""id"">$(Esc $c.Cidade)/$(Esc $c.UF)</span></td><td>$(Esc $o.Produto)</td><td class=""nw"">$(Esc $o.Etapa)</td><td class=""num"">$(Format-Prob $o.Probabilidade)</td><td class=""num"">$(Format-Moeda $o.Valor)</td><td class=""num"">$(Format-Moeda $a.Ponderado)</td><td class=""nw"">$(Format-Data $o.Previsao)</td><td class=""nw"">$sit</td></tr>"
    }

    $corpo = @"
<section class="filtros">
  <h1>Pipeline por vendedor</h1>
  <form method="get" action="/pipeline">
    <label>Meta e realizado de <select name="de">$(& $opcoes $d.MesInicial)</select></label>
    <label>até <select name="ate">$(& $opcoes $d.MesFinal)</select></label>
    $(if ($FiltroVendedor) { "<input type=""hidden"" name=""vendedor"" value=""$(Esc $FiltroVendedor)"">" })
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Meta e realizado: <strong>$periodo</strong>, data-base das vendas $(Format-Data $d.DataBase). Pipeline: <strong>foto do CRM de $dbCrm</strong> ($(Esc $vis.ArquivoCrm.Name)), não filtrada pelo período. Atraso contado até $dbCrm, não até hoje.</p>
</section>
$parcial
$avisoOrfas
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">$(Format-Moeda $t.Meta)</span><span class="det">realizado $(Format-Moeda $t.RealizadoMesesMeta) · $(Format-Pct $t.Atingimento)</span></div>
  <div class="cartao"><span class="rotulo">Pipeline aberto</span><span class="valor">$(Format-Moeda $t.Pipeline)</span><span class="det">$($t.QtdAbertas) oportunidades · soma do valor estimado</span></div>
  <div class="cartao destaque"><span class="rotulo">Pipeline ponderado</span><span class="valor">$(Format-Moeda $t.Ponderado)</span><span class="det">valor estimado × probabilidade da etapa</span></div>
  <div class="cartao alerta"><span class="rotulo">Previsão de fechamento vencida</span><span class="valor">$($t.QtdAtrasadas) oportunidades</span><span class="det">$(Format-Moeda $t.ValorAtrasado) em pipeline atrasado</span></div>
</section>

<section class="bloco">
  <h2>Uma linha por vendedor</h2>
  <p class="nota">Todos os vendedores, inclusive desligados, para que o pipeline órfão apareça e o total feche (REGRAS_NEGOCIO.md §7.4). Não é ranking: a ordem é pelo ID. Clique num nome para filtrar as oportunidades.</p>
  <table>
    <thead><tr><th>Vendedor</th><th>Filial</th><th class="num">Meta do período</th><th class="num">Realizado</th><th class="num">Atingimento</th><th class="num">Falta p/ meta</th><th class="num">Abertas</th><th class="num">Atrasadas</th><th class="num">Pipeline aberto</th><th class="num">Ponderado</th></tr></thead>
    <tbody>$($linhasVend -join "`n")</tbody>
    <tfoot><tr><td>Total da empresa</td><td></td><td class="num">$(Format-Moeda $t.Meta)</td><td class="num">$(Format-Moeda $t.RealizadoMesesMeta)</td><td class="num">$(Format-Pct $t.Atingimento)</td><td class="num">$gapTotal</td><td class="num">$($t.QtdAbertas)</td><td class="num">$($t.QtdAtrasadas)</td><td class="num">$(Format-Moeda $t.Pipeline)</td><td class="num">$(Format-Moeda $t.Ponderado)</td></tr></tfoot>
  </table>
  $notaFora
</section>

<section class="bloco" id="oportunidades">
  <div class="cab-bloco"><h2>$tituloLista</h2><a class="botao" href="/pipeline.csv">Baixar CSV</a></div>
  <p class="nota">$($lista.Count) abertas, $qtdAtrLista com previsão de fechamento anterior a $dbCrm (em destaque, as mais atrasadas primeiro). Aberta = etapa diferente de Fechada Ganha e Fechada Perdida.</p>
  <table>
    <thead><tr><th>Oportunidade</th><th>Vendedor</th><th>Cliente</th><th>Produto</th><th>Etapa</th><th class="num">Prob.</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th>Previsão</th><th>Situação</th></tr></thead>
    <tbody>$($linhasOp -join "`n")</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  aberta = etapa ≠ Fechada Ganha e ≠ Fechada Perdida (§7) ·
  pipeline = soma do valor estimado; ponderado = valor × probabilidade (§7) ·
  atrasada = previsão anterior à data-base do CRM (§0.3, §7) ·
  meta e realizado só nos meses com meta, só vendas faturadas (§1, §4.1) ·
  empresa = soma ÷ soma (§4.2) ·
  órfãs de desligados ficam no pipeline (§2.4).
</footer>
"@
    return New-Layout "Pipeline · CRM de $dbCrm" $corpo $Usuario 'pipeline'
}

function New-CsvPipeline($Visao) {
    $f = { param($x) if ($null -eq $x) { '' } else { ([decimal]$x).ToString('0.00', $script:PtBr) } }
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine('ID Oportunidade;ID Vendedor;Vendedor;Status do vendedor;ID Cliente;Cliente;ID Produto;Produto;Etapa;Probabilidade;Valor estimado;Ponderado;Previsão de fechamento;Atrasada;Dias de atraso;Órfã;Linha na planilha')
    foreach ($a in $Visao.Abertas) {
        $o = $a.Oportunidade
        [void]$sb.AppendLine((@(
                    $o.Id, $a.Vendedor.Id, $a.Vendedor.Nome, $a.Vendedor.Status, $a.Cliente.Id, $a.Cliente.Nome, $o.IdProduto, $o.Produto, $o.Etapa,
                    ([decimal]$o.Probabilidade).ToString('0.00', $script:PtBr), (& $f $o.Valor), (& $f $a.Ponderado), (Format-Data $o.Previsao),
                    $(if ($a.Atrasada) { 'sim' } else { 'não' }), $a.DiasAtraso, $(if ($a.Orfa) { 'sim' } else { 'não' }), $o.Linha
                ) -join ';'))
    }
    $t = $Visao.Total
    [void]$sb.AppendLine(("TOTAL ({0} abertas, {1} atrasadas);;;;;;;;;;{2};{3};;;;;" -f $t.QtdAbertas, $t.QtdAtrasadas, (& $f $t.Pipeline), (& $f $t.Ponderado)))
    return $sb.ToString()
}
