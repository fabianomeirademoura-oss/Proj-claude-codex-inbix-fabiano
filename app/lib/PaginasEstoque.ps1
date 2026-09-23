# HTML e CSV da visão de estoque. Só formata: os números vêm de RegrasEstoque.ps1.
# Usa os auxiliares de Paginas.ps1 (Esc, Format-*, New-Layout).

function Get-FrasePrazoParado($Visao) {
    # A definição de "parado" que a tela usa, numa frase, com o prazo escolhido (REGRAS_NEGOCIO.md §5).
    return "Parado é a linha de estoque (um produto numa filial) que tem pelo menos 1 unidade e está há $($Visao.DiasParado) dias ou mais sem nenhuma entrada nem saída, contados até a data da foto ($(Format-Data $Visao.DataBase)); linha sem nenhuma data de movimento também conta como parada."
}

function Format-Ops($Abertas) {
    if ($null -eq $Abertas) { return '<span class="nota">CRM indisponível</span>' }
    if ($Abertas.Qtd -eq 0) { return '<span class="nota">nenhuma</span>' }
    return "<strong>$($Abertas.Qtd)</strong> <span class=""id"">$(Format-Moeda $Abertas.Valor)</span>"
}

function Format-Outras($Outras, [switch]$ComSaida) {
    # "Estoque é por produto E por filial": mostra o mesmo produto nas outras filiais.
    $itens = foreach ($o in $Outras) {
        $saida = if ($ComSaida) { if ($o.UltimaSaida) { ", última saída $(Format-Data $o.UltimaSaida)" } else { ', sem saída em 2026' } } else { '' }
        $cls = if ($o.Quantidade -gt $o.Minimo) { 'item quebra forte' } else { 'item quebra' }
        "<span class=""$cls"">$(Esc $o.Filial): $($o.Quantidade) (mín. $($o.Minimo))$saida</span>"
    }
    return ($itens -join '')
}

function Format-Movimento($Linha) {
    if ($null -eq $Linha.DiasSemMovimento) { return '<span class="marca atraso">sem movimento registrado</span>' }
    return "$($Linha.DiasSemMovimento) dias"
}

function Get-QueryEstoque($Visao, [string]$IdFilial) {
    $q = "dias=$($Visao.DiasParado)"
    if ($IdFilial) { $q += "&amp;filial=$(Esc $IdFilial)" }
    return $q
}

function New-PaginaEstoque($Visao, $Usuario, [string]$FiltroFilial, [string]$AvisoPrazo) {
    $v = $Visao
    $db = Format-Data $v.DataBase
    $t = $v.Total
    $q = Get-QueryEstoque $v $FiltroFilial
    $nomeFiltro = if ($FiltroFilial) { ($v.Filiais | Where-Object Id -eq $FiltroFilial).Nome }
    $opcoesFilial = '<option value="">todas</option>' + (($v.Filiais | ForEach-Object { $s = if ($_.Id -eq $FiltroFilial) { ' selected' } else { '' }; "<option value=""$(Esc $_.Id)""$s>$(Esc $_.Nome)</option>" }) -join '')
    $aviso = if ($AvisoPrazo) { "<p class=""parcial"" role=""alert"">$(Esc $AvisoPrazo)</p>" } else { '' }
    $semCrm = if (-not $v.ComCrm) { '<p class="parcial">O CRM não pôde ser carregado: a coluna de oportunidades abertas está indisponível. Veja a página Pipeline para os erros.</p>' } else { '' }
    $noFiltro = { param($l) -not $FiltroFilial -or $l.IdFilial -eq $FiltroFilial }

    # 1. Matriz categoria × filial
    $cab = ($v.Filiais | ForEach-Object { "<th class=""num""><a href=""/estoque?dias=$($v.DiasParado)&amp;filial=$(Esc $_.Id)#parados"">$(Esc $_.Nome)</a></th>" }) -join ''
    $cel = {
        param($c)
        $parado = if ($c.Parado -gt 0) { "<span class=""item alerta-txt"">parado $(Format-Moeda $c.Parado)</span>" } else { '' }
        "<td class=""num"">$(Format-Moeda $c.Valor)$parado</td>"
    }
    $linhasMatriz = foreach ($cat in $v.Categorias) {
        $celulas = ($v.Filiais | ForEach-Object { & $cel $v.Celulas["$cat|$($_.Id)"] }) -join ''
        $tc = $v.TotCategoria[$cat]
        $pct = if ($t.Valor -gt 0) { Format-Pct ($tc.Valor / $t.Valor) } else { '—' }
        "<tr><td>$(Esc $cat)</td>$celulas$(& $cel $tc)<td class=""num"">$pct</td></tr>"
    }
    $rodape = ($v.Filiais | ForEach-Object { & $cel $v.TotFilial[$_.Id] }) -join ''

    # 2. Parados (lista pedida; prazo configurável na tela)
    $parados = @($v.Parados | Where-Object { & $noFiltro $_.Linha })
    $valorParadosLista = [decimal]0; foreach ($p in $parados) { $valorParadosLista += $p.Linha.Valor }
    $linhasParados = foreach ($p in $parados) {
        $l = $p.Linha
        $desc = if ($l.Produto.Descontinuado) { ' <span class="marca desligado">descontinuado</span>' } else { '' }
        "<tr><td class=""nw"">$(Esc $l.Produto.Id)</td><td>$(Esc $l.Produto.Nome)$desc<span class=""item"">$(Esc $l.Produto.Categoria)</span></td><td>$(Esc $l.Filial)</td><td class=""num"">$($l.Quantidade)</td><td class=""num"">$(Format-Moeda $l.Valor)</td><td class=""nw"">$(Format-Data $l.UltimaEntrada)</td><td class=""nw"">$(Format-Data $l.UltimaSaida)</td><td class=""num"">$(Format-Movimento $l)</td><td class=""num"">$(Format-Ops $p.Abertas)</td><td>$(Format-Outras $p.Outras -ComSaida)</td></tr>"
    }
    $vazioParados = if (-not $parados.Count) { '<tr><td colspan="10" class="nota">Nenhuma linha parada com este prazo.</td></tr>' } else { '' }

    # 3. Sem estoque em nenhuma filial
    $linhasSem = foreach ($s in $v.SemEstoque) {
        $desc = if ($s.Produto.Descontinuado) { ' <span class="marca desligado">descontinuado</span>' } else { '' }
        $mins = ($s.Linhas | ForEach-Object { "<span class=""item"">$(Esc $_.Filial): mín. $($_.Minimo)</span>" }) -join ''
        $cls = if ($s.Abertas -and $s.Abertas.Qtd -gt 0) { ' class="atrasada"' } else { '' }
        "<tr$cls><td class=""nw"">$(Esc $s.Produto.Id)</td><td>$(Esc $s.Produto.Nome)$desc</td><td>$(Esc $s.Produto.Categoria)</td><td>$mins</td><td class=""nw"">$(if ($s.UltimaSaida) { Format-Data $s.UltimaSaida } else { '<span class=""nota"">sem saída em 2026</span>' })</td><td class=""num"">$(Format-Ops $s.Abertas)</td></tr>"
    }

    # 4. Abaixo do mínimo
    $abaixo = @($v.Abaixo | Where-Object { & $noFiltro $_.Linha })
    $linhasAbaixo = foreach ($a in $abaixo) {
        $l = $a.Linha
        $zerada = if ($l.Quantidade -eq 0) { ' <span class="marca atraso">zerado</span>' } else { '' }
        "<tr><td>$(Esc $l.Filial)</td><td class=""nw"">$(Esc $l.Produto.Id)</td><td>$(Esc $l.Produto.Nome)<span class=""item"">$(Esc $l.Produto.Categoria)</span></td><td class=""num"">$($l.Quantidade)$zerada</td><td class=""num"">$($l.Minimo)</td><td class=""num forte"">$($a.Falta)</td><td>$(Format-Outras $a.Outras)</td><td class=""num"">$(Format-Ops $a.Abertas)</td></tr>"
    }
    $sufixoFiltro = if ($nomeFiltro) { " em $(Esc $nomeFiltro) <a class=""nota"" href=""/estoque?dias=$($v.DiasParado)"">ver todas as filiais</a>" } else { '' }

    $corpo = @"
<section class="filtros">
  <h1>Estoque</h1>
  <form method="get" action="/estoque">
    <label>Parado a partir de <input type="number" name="dias" min="1" max="$($script:DiasParadoMaximo)" step="1" value="$($v.DiasParado)" class="curto"> dias sem movimento</label>
    <label>Filial nas listas <select name="filial">$opcoesFilial</select></label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Foto do estoque de <strong>$db</strong> ($(Esc $v.Arquivo.Name)). Dias contados até $db, não até hoje. Estoque é por produto <em>e</em> por filial: cada linha abaixo é um produto numa filial. Padrão do prazo: $($script:DiasParadoPadrao) dias (REGRAS_NEGOCIO.md §5).</p>
</section>
$aviso
$semCrm
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Dinheiro no pátio (valor imobilizado)</span><span class="valor">$(Format-Moeda $t.Valor)</span><span class="det">$($t.Unidades) unidades · quantidade × custo médio</span></div>
  <div class="cartao alerta"><span class="rotulo">Parado há $($v.DiasParado)+ dias</span><span class="valor">$(Format-Moeda $t.Parado)</span><span class="det">$($v.Parados.Count) linhas (produto, filial)</span></div>
  <div class="cartao"><span class="rotulo">Sem estoque em nenhuma filial</span><span class="valor">$($v.SemEstoque.Count) produtos</span><span class="det">$(@($v.SemEstoque | Where-Object { $_.Abertas -and $_.Abertas.Qtd }).Count) com oportunidade aberta no CRM</span></div>
  <div class="cartao"><span class="rotulo">Abaixo do estoque mínimo</span><span class="valor">$($v.Abaixo.Count) linhas</span><span class="det">quantidade &lt; mínimo na filial</span></div>
</section>

<section class="bloco">
  <h2>1. Quanto dinheiro está parado no pátio, por filial e por categoria</h2>
  <p class="nota">Valor imobilizado = quantidade × custo médio de cada linha. Em vermelho, a parte que está parada há $($v.DiasParado) dias ou mais. Clique numa filial para ver os parados e os itens a repor dela.</p>
  <table>
    <thead><tr><th>Categoria</th>$cab<th class="num">Total</th><th class="num">% do total</th></tr></thead>
    <tbody>$($linhasMatriz -join "`n")</tbody>
    <tfoot><tr><td>Total</td>$rodape$(& $cel $t)<td class="num">100,0%</td></tr></tfoot>
  </table>
</section>

<section class="bloco" id="parados">
  <h2>2. Produtos parados$sufixoFiltro</h2>
  <p class="definicao"><strong>Definição usada:</strong> $(Esc (Get-FrasePrazoParado $v)) Para usar outro prazo, mude o campo no topo da página.</p>
  <p class="nota">$($parados.Count) linhas, $(Format-Moeda $valorParadosLista). Mais tempo parado primeiro. "Oportunidades abertas" é do produto no CRM, em qualquer filial: parado não quer dizer sem demanda. A última coluna mostra o mesmo produto nas outras filiais; em negrito, onde há mais que o mínimo.</p>
  <table>
    <thead><tr><th>ID</th><th>Produto</th><th>Filial</th><th class="num">Qtd.</th><th class="num">Valor parado</th><th>Última entrada</th><th>Última saída</th><th class="num">Sem movimento</th><th class="num">Oport. abertas</th><th>Nas outras filiais</th></tr></thead>
    <tbody>$($linhasParados -join "`n")$vazioParados</tbody>
  </table>
</section>

<section class="bloco">
  <h2>3. O que está sem estoque em lugar nenhum</h2>
  <p class="nota">Produtos com quantidade zero nas $($v.Filiais.Count) filiais somadas. Em destaque, os que têm oportunidade aberta no CRM: há cliente esperando e não há máquina no pátio.</p>
  <table>
    <thead><tr><th>ID</th><th>Produto</th><th>Categoria</th><th>Estoque mínimo</th><th>Última saída</th><th class="num">Oport. abertas</th></tr></thead>
    <tbody>$($linhasSem -join "`n")</tbody>
  </table>
</section>

<section class="bloco">
  <div class="cab-bloco"><h2>4. Abaixo do estoque mínimo: precisa repor$sufixoFiltro</h2><a class="botao" href="/estoque.csv?$q">Baixar CSV</a></div>
  <p class="nota">$($abaixo.Count) linhas em que a quantidade da filial é menor que o mínimo da filial. "Faltam" = mínimo − quantidade. Antes de comprar, veja se outra filial tem sobra do mesmo produto (em negrito).</p>
  <table>
    <thead><tr><th>Filial</th><th>ID</th><th>Produto</th><th class="num">Qtd.</th><th class="num">Mínimo</th><th class="num">Faltam</th><th>Nas outras filiais</th><th class="num">Oport. abertas</th></tr></thead>
    <tbody>$($linhasAbaixo -join "`n")</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md §5):</strong>
  tudo por linha (produto, filial) ·
  valor imobilizado = quantidade × custo médio ·
  parado = quantidade &gt; 0 e dias sem entrada nem saída ≥ prazo ·
  sem estoque = soma das filiais = 0 ·
  abaixo do mínimo = quantidade &lt; mínimo na linha ·
  dias contados até a data da foto, nunca até hoje.
</footer>
"@
    return New-Layout "Estoque · foto de $db" $corpo $Usuario 'estoque'
}

function New-CsvEstoque($Visao) {
    # Uma linha por (produto, filial), com as marcas usadas na tela, para conferir no Excel.
    $f = { param($x) if ($null -eq $x) { '' } else { ([decimal]$x).ToString('0.00', $script:PtBr) } }
    $semEstoque = @{}; foreach ($s in $Visao.SemEstoque) { $semEstoque[$s.Produto.Id] = $true }
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine("ID Produto;Produto;Categoria;Status do produto;Filial;Quantidade;Estoque mínimo;Custo médio;Valor imobilizado;Última entrada;Última saída;Dias sem movimento;Parado (prazo $($Visao.DiasParado) dias);Abaixo do mínimo;Produto sem estoque em nenhuma filial;Linha na planilha")
    foreach ($l in ($Visao.Linhas | Sort-Object { $_.Produto.Id }, IdFilial)) {
        [void]$sb.AppendLine((@(
                    $l.Produto.Id, $l.Produto.Nome, $l.Produto.Categoria, $l.Produto.Status, $l.Filial, $l.Quantidade, $l.Minimo, (& $f $l.CustoMedio), (& $f $l.Valor),
                    (Format-Data $l.UltimaEntrada), (Format-Data $l.UltimaSaida), $(if ($null -eq $l.DiasSemMovimento) { 'sem movimento' } else { $l.DiasSemMovimento }),
                    $(if (Test-LinhaParada $l $Visao.DiasParado) { 'sim' } else { 'não' }), $(if (Test-AbaixoMinimo $l) { 'sim' } else { 'não' }),
                    $(if ($semEstoque.ContainsKey($l.Produto.Id)) { 'sim' } else { 'não' }), $l.Linha
                ) -join ';'))
    }
    $t = $Visao.Total
    [void]$sb.AppendLine(("TOTAL;;;;;{0};;;{1};;;;{2} linhas, {3};{4} linhas;{5} produtos;" -f $t.Unidades, (& $f $t.Valor), $Visao.Parados.Count, (& $f $t.Parado), $Visao.Abaixo.Count, $Visao.SemEstoque.Count))
    return $sb.ToString()
}
