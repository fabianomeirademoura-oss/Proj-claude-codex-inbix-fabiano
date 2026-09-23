# HTML da página do funil de vendas. Só formata: os números vêm de Get-VisaoFunil (RegrasPipeline.ps1).
# O desenho é SVG inline com classes do estilo.css (a CSP não permite script nem style inline).

function Format-Qtd([int]$N, [string]$Singular, [string]$Plural) { if ($N -eq 1) { "1 $Singular" } else { "$N $Plural" } }

function Get-QueryFunil($Filtro, [string]$Medida, [hashtable]$Troca = @{}) {
    $v = @{ vendedor = $Filtro.IdVendedor; filial = $Filtro.IdFilial; categoria = $Filtro.Categoria; medida = $Medida }
    foreach ($k in $Troca.Keys) { $v[$k] = $Troca[$k] }
    $partes = foreach ($k in @('vendedor', 'filial', 'categoria', 'medida')) { if ($v[$k]) { "$k=$([Uri]::EscapeDataString([string]$v[$k]))" } }
    return (Esc ($partes -join '&'))
}

function New-SvgFunil($Etapas, [string]$Medida) {
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $n = { param($x) ([double]$x).ToString('0.#', $inv) }
    $medir = { param($e) if ($Medida -eq 'valor') { [double]$e.Valor } else { [double]$e.Qtd } }
    $medirAtraso = { param($e) if ($Medida -eq 'valor') { [double]$e.ValorAtrasado } else { [double]$e.QtdAtrasadas } }
    $max = 0.0; foreach ($e in $Etapas) { $m = & $medir $e; if ($m -gt $max) { $max = $m } }
    $larguraMax = 400.0; $centro = 395.0; $alt = 50; $passo = 64
    $partes = New-Object Collections.Generic.List[string]
    $y = 8
    foreach ($e in $Etapas) {
        $m = & $medir $e
        $w = if ($max -gt 0) { $larguraMax * $m / $max } else { 0 }
        $x = $centro - $w / 2
        $ym = $y + $alt / 2
        $prob = if ($null -ne $e.Probabilidade) { " · $(Format-Prob $e.Probabilidade)" } else { '' }
        $partes.Add("<text class=""f-rotulo"" x=""0"" y=""$(& $n ($ym - 3))"">$(Esc $e.Etapa)</text>")
        $partes.Add("<text class=""f-sub"" x=""0"" y=""$(& $n ($ym + 15))"">probabilidade$prob</text>")
        if ($m -gt 0) {
            $partes.Add("<rect class=""f-barra"" x=""$(& $n $x)"" y=""$y"" width=""$(& $n $w)"" height=""$alt"" rx=""4""><title>$(Esc $e.Etapa): $($e.Qtd) abertas, $(Format-Moeda $e.Valor)</title></rect>")
            $wa = $larguraMax * (& $medirAtraso $e) / $max
            if ($wa -gt 0) { $partes.Add("<rect class=""f-atraso"" x=""$(& $n $x)"" y=""$y"" width=""$(& $n $wa)"" height=""$alt"" rx=""4""><title>$(Esc $e.Etapa): $($e.QtdAtrasadas) atrasadas, $(Format-Moeda $e.ValorAtrasado)</title></rect>") }
        } else {
            $partes.Add("<line class=""f-vazio"" x1=""$(& $n ($centro - 20))"" x2=""$(& $n ($centro + 20))"" y1=""$ym"" y2=""$ym""></line>")
        }
        $principal = if ($Medida -eq 'valor') { Format-Moeda $e.Valor } else { Format-Qtd $e.Qtd 'aberta' 'abertas' }
        $secundario = if ($Medida -eq 'valor') { "$(Format-Qtd $e.Qtd 'aberta' 'abertas') · $(Format-Qtd $e.QtdAtrasadas 'atrasada' 'atrasadas')" } else { "$(Format-Moeda $e.Valor) · $(Format-Qtd $e.QtdAtrasadas 'atrasada' 'atrasadas')" }
        $partes.Add("<text class=""f-rotulo"" x=""800"" y=""$(& $n ($ym - 3))"" text-anchor=""end"">$(Esc $principal)</text>")
        $partes.Add("<text class=""f-sub"" x=""800"" y=""$(& $n ($ym + 15))"" text-anchor=""end"">$(Esc $secundario)</text>")
        $y += $passo
    }
    $altura = $y - $passo + $alt + 8
    $resumo = ($Etapas | ForEach-Object { "$($_.Etapa): $($_.Qtd) abertas, $($_.QtdAtrasadas) atrasadas, $(Format-Moeda $_.Valor)" }) -join '; '
    return "<svg class=""funil"" viewBox=""0 0 800 $altura"" role=""img"" aria-label=""Funil de oportunidades abertas por etapa. $(Esc $resumo)"">$($partes -join '')</svg>"
}

function New-PaginaFunil($Visao, $Usuario, [string]$Medida) {
    $v = $Visao
    $f = $v.Filtro
    $t = $v.Total
    $dbCrm = Format-Data $v.DataBaseCrm
    $euId = $Usuario.idVendedor
    $sel = { param($a, $b) if ($a -eq $b) { ' selected' } else { '' } }

    $opVend = '<option value="">todos</option>' + (($v.Vendedores | ForEach-Object {
                $marca = if (-not $_.Ativo) { ' (desligado)' } else { '' }
                "<option value=""$(Esc $_.Id)""$(& $sel $_.Id $f.IdVendedor)>$(Esc $_.Nome)$marca</option>"
            }) -join '')
    $opFilial = '<option value="">todas</option>' + (($v.Filiais | ForEach-Object { "<option value=""$(Esc $_.Id)""$(& $sel $_.Id $f.IdFilial)>$(Esc $_.Nome)</option>" }) -join '')
    $opCat = '<option value="">todas</option>' + (($v.Categorias | ForEach-Object { "<option value=""$(Esc $_)""$(& $sel $_ $f.Categoria)>$(Esc $_)</option>" }) -join '')
    $chkQtd = if ($Medida -ne 'valor') { ' checked' } else { '' }
    $chkVal = if ($Medida -eq 'valor') { ' checked' } else { '' }

    # Descrição do recorte, em texto
    $recorte = @()
    if ($f.IdVendedor) { $vd = $v.Vendedores | Where-Object Id -eq $f.IdVendedor; $recorte += "vendedor <strong>$(Esc $vd.Nome)</strong> $(Get-Marcas $vd)" }
    if ($f.IdFilial) { $recorte += "filial do vendedor <strong>$(Esc ($v.Filiais | Where-Object Id -eq $f.IdFilial).Nome)</strong>" }
    if ($f.Categoria) { $recorte += "categoria <strong>$(Esc $f.Categoria)</strong>" }
    $textoRecorte = if ($recorte) { "Mostrando: $($recorte -join ' · ') · <a href=""/funil?$(Get-QueryFunil ([pscustomobject]@{}) $Medida)"">limpar filtros</a>" } else { 'Mostrando: todo o CRM, sem filtro.' }

    # Tabela por etapa
    $linhasEtapa = foreach ($e in $v.Etapas) {
        $pct = if ($t.Valor -gt 0) { Format-Pct ($e.Valor / $t.Valor) } else { '—' }
        $atr = if ($e.QtdAtrasadas) { "<span class=""marca atraso"">$($e.QtdAtrasadas)</span> <span class=""id"">$(Format-Moeda $e.ValorAtrasado)</span>" } else { '0' }
        $prob = if ($null -ne $e.Probabilidade) { Format-Prob $e.Probabilidade } else { '—' }
        "<tr><td>$(Esc $e.Etapa)</td><td class=""num"">$prob</td><td class=""num"">$($e.Qtd)</td><td class=""num"">$atr</td><td class=""num"">$(Format-Moeda $e.Valor)</td><td class=""num forte"">$(Format-Moeda $e.Ponderado)</td><td class=""num"">$pct</td></tr>"
    }

    # Fechadas e motivos de perda
    $maxMotivo = 0; foreach ($m in $v.Motivos) { if ($m.Qtd -gt $maxMotivo) { $maxMotivo = $m.Qtd } }
    $linhasMotivo = foreach ($m in $v.Motivos) {
        "<tr><td>$(Esc $m.Motivo)</td><td><meter min=""0"" max=""$maxMotivo"" value=""$($m.Qtd)""></meter></td><td class=""num"">$($m.Qtd)</td><td class=""num"">$(Format-Moeda $m.Valor)</td></tr>"
    }
    $conv = if ($null -ne $v.Conversao) { Format-Pct $v.Conversao } else { '—' }
    $convValor = if ($null -ne $v.ConversaoValor) { Format-Pct $v.ConversaoValor } else { '—' }
    $tabelaMotivos = if ($v.Motivos.Count) {
        "<table><thead><tr><th>Motivo da perda</th><th></th><th class=""num"">Perdidas</th><th class=""num"">Valor estimado</th></tr></thead><tbody>$($linhasMotivo -join "`n")</tbody></table>"
    } else { '<p class="nota">Nenhuma oportunidade perdida neste recorte.</p>' }

    # Lista das abertas, atrasadas em destaque
    $linhasOp = foreach ($a in $v.Abertas) {
        $o = $a.Oportunidade; $vd = $a.Vendedor; $c = $a.Cliente
        $cls = @(); if ($a.Atrasada) { $cls += 'atrasada' }; if ($vd.Id -eq $euId) { $cls += 'voce' }
        $attr = if ($cls) { " class=""$($cls -join ' ')""" } else { '' }
        $sit = if ($a.Atrasada) { "<span class=""marca atraso"">atrasada há $($a.DiasAtraso) dias</span>" } else { '<span class="nota">no prazo</span>' }
        $orfa = if ($a.Orfa) { ' <span class="marca desligado">órfã</span>' } else { '' }
        "<tr$attr><td class=""nw"">$(Esc $o.Etapa)</td><td class=""nw"">$(Esc $o.Id)</td><td><a href=""/funil?$(Get-QueryFunil $f $Medida @{ vendedor = $vd.Id })"">$(Esc $vd.Nome)</a>$orfa <span class=""id"">$(Esc $vd.Filial)</span></td><td>$(Esc $c.Nome)</td><td>$(Esc $o.Produto)<span class=""item"">$(Esc $a.Produto.Categoria)</span></td><td class=""num"">$(Format-Moeda $o.Valor)</td><td class=""num"">$(Format-Moeda $a.Ponderado)</td><td class=""nw"">$(Format-Data $o.Previsao)</td><td class=""nw"">$sit</td></tr>"
    }
    $vazio = if (-not $v.Abertas.Count) { '<tr><td colspan="9" class="nota">Nenhuma oportunidade aberta neste recorte.</td></tr>' } else { '' }
    $rotuloMedida = if ($Medida -eq 'valor') { 'valor estimado' } else { 'quantidade de oportunidades' }

    $corpo = @"
<section class="filtros">
  <h1>Funil de vendas</h1>
  <form method="get" action="/funil">
    <label>Vendedor <select name="vendedor">$opVend</select></label>
    <label>Filial <select name="filial">$opFilial</select></label>
    <label>Categoria <select name="categoria">$opCat</select></label>
    <span class="check">Largura das barras:
      <label class="check"><input type="radio" name="medida" value="qtd"$chkQtd> quantidade</label>
      <label class="check"><input type="radio" name="medida" value="valor"$chkVal> valor</label>
    </span>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">$textoRecorte</p>
  <p class="fonte">Foto do CRM de <strong>$dbCrm</strong> ($(Esc $v.ArquivoCrm.Name)), não filtrada por período. Atraso contado até $dbCrm. Filial = filial do vendedor; categoria = categoria do produto no catálogo.</p>
</section>
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Oportunidades abertas</span><span class="valor">$($t.Qtd)</span><span class="det">$(Format-Moeda $t.Valor) em valor estimado</span></div>
  <div class="cartao destaque"><span class="rotulo">Pipeline ponderado</span><span class="valor">$(Format-Moeda $t.Ponderado)</span><span class="det">valor estimado × probabilidade da etapa</span></div>
  <div class="cartao alerta"><span class="rotulo">Previsão de fechamento vencida</span><span class="valor">$(Format-Qtd $t.QtdAtrasadas 'aberta' 'abertas')</span><span class="det">$(Format-Moeda $t.ValorAtrasado) em pipeline atrasado</span></div>
  <div class="cartao"><span class="rotulo">Conversão das fechadas</span><span class="valor">$conv</span><span class="det">$($v.Ganhas.Qtd) ganhas ÷ $($v.Ganhas.Qtd + $v.Perdidas.Qtd) fechadas</span></div>
</section>

<section class="bloco">
  <h2>Onde estão as oportunidades abertas hoje</h2>
  <p class="definicao">Cada barra é o total de oportunidades <strong>abertas</strong> que estão hoje naquela etapa, medido por $rotuloMedida. O CRM guarda só a etapa atual: o funil não mostra quantas oportunidades passaram por cada etapa (REGRAS_NEGOCIO.md §8).</p>
  <p class="legenda"><span class="cor cor-dia"></span> no prazo <span class="cor cor-atraso"></span> previsão de fechamento já passou</p>
  $(New-SvgFunil $v.Etapas $Medida)
  <table>
    <thead><tr><th>Etapa</th><th class="num">Prob.</th><th class="num">Abertas</th><th class="num">Atrasadas</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th class="num">% do valor aberto</th></tr></thead>
    <tbody>$($linhasEtapa -join "`n")</tbody>
    <tfoot><tr><td>Total aberto</td><td></td><td class="num">$($t.Qtd)</td><td class="num">$($t.QtdAtrasadas)</td><td class="num">$(Format-Moeda $t.Valor)</td><td class="num">$(Format-Moeda $t.Ponderado)</td><td class="num">$(if ($t.Valor -gt 0) { '100,0%' } else { '—' })</td></tr></tfoot>
  </table>
</section>

<section class="rankings">
  <div class="bloco">
    <h2>Fechadas: ganhas × perdidas</h2>
    <table>
      <thead><tr><th></th><th class="num">Oportunidades</th><th class="num">Valor estimado</th></tr></thead>
      <tbody>
        <tr><td>Fechadas Ganhas</td><td class="num">$($v.Ganhas.Qtd)</td><td class="num">$(Format-Moeda $v.Ganhas.Valor)</td></tr>
        <tr><td>Fechadas Perdidas</td><td class="num">$($v.Perdidas.Qtd)</td><td class="num">$(Format-Moeda $v.Perdidas.Valor)</td></tr>
      </tbody>
      <tfoot><tr><td>Conversão</td><td class="num">$conv</td><td class="num">$convValor</td></tr></tfoot>
    </table>
    <p class="nota">Conversão = ganhas ÷ (ganhas + perdidas), em quantidade e, ao lado, em valor estimado. Abertas não entram.</p>
  </div>
  <div class="bloco">
    <h2>Por que perdemos</h2>
    $tabelaMotivos
  </div>
</section>

<section class="bloco" id="abertas">
  <h2>Oportunidades abertas deste recorte</h2>
  <p class="nota">$($v.Abertas.Count) abertas, $($t.QtdAtrasadas) com previsão de fechamento anterior a $dbCrm (em destaque). Da etapa mais avançada para a menos avançada; dentro da etapa, atrasadas primeiro. Clique num vendedor para filtrar o funil por ele.</p>
  <table>
    <thead><tr><th>Etapa</th><th>Oportunidade</th><th>Vendedor</th><th>Cliente</th><th>Produto</th><th class="num">Valor estimado</th><th class="num">Ponderado</th><th>Previsão</th><th>Situação</th></tr></thead>
    <tbody>$($linhasOp -join "`n")$vazio</tbody>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  aberta = etapa ≠ Fechada Ganha e ≠ Fechada Perdida (§7) ·
  barra = abertas na etapa atual, sem histórico de etapas (§8) ·
  atrasada = previsão anterior à data-base do CRM (§7) ·
  conversão = ganhas ÷ (ganhas + perdidas) (§8) ·
  filial = filial do vendedor (§8.4) ·
  órfãs de desligados ficam no funil (§2.4).
</footer>
"@
    return New-Layout "Funil de vendas · CRM de $dbCrm" $corpo $Usuario 'funil'
}
