# Apresentação da origem do faturamento. Os cálculos ficam em RegrasOrigem.ps1.
function New-PaginaOrigem($Base, $Visao, $Usuario) {
    $v = $Visao
    $opcoes = { param($selecionado)
        (Get-MesesDisponiveis $Base | ForEach-Object {
            $sel = if ($_ -eq $selecionado) { ' selected' } else { '' }
            "<option value=""$_""$sel>$($script:NomesMesLongo[$_])</option>"
        }) -join ''
    }
    $pct = if ($null -ne $v.FracaoCrm) { Format-Pct $v.FracaoCrm } else { '—' }
    $pctSem = if ($null -ne $v.FracaoSem) { Format-Pct $v.FracaoSem } else { '—' }
    $resumo = if ($null -ne $v.FracaoCrm) {
        "<strong>$pct do faturamento nasceu em oportunidades registradas no CRM.</strong> São $($v.QtdComCrm) vendas faturadas, somando $(Format-Moeda $v.ComCrm)."
    } else { 'Sem faturamento no período. A participação do CRM não se aplica.' }
    $parcial = if ($v.Parcial) { "<p class=""parcial"">Faturamento parcial até $(Format-Data $v.DataBase).</p>" } else { '' }
    $categorias = foreach ($c in $v.CategoriasSem) {
        "<tr><td>$(Esc $c.Categoria)</td><td class=""num"">$($c.Quantidade)</td><td class=""num"">$(Format-Moeda $c.Valor)</td></tr>"
    }
    $vazioCat = if (-not $v.CategoriasSem.Count) { '<tr><td colspan="3">Nenhuma venda faturada sem oportunidade neste período.</td></tr>' } else { '' }
    $linhas = foreach ($l in $v.Linhas) {
        $s = $l.Venda; $o = $l.Oportunidade
        $origem = if ($o) { "$(Esc $o.Id)<span class=""item"">$(Esc $o.Etapa)</span><span class=""item"">Origem: $(Esc $o.Origem) · linha $($o.Linha) do CRM</span>" } else { 'Sem oportunidade registrada' }
        $classe = if (-not $l.Faturada) { ' class="cancelada"' } else { '' }
        $status = if ($l.Faturada) { 'Faturada' } else { 'Cancelada · fora do faturamento' }
        "<tr$classe><td class=""nw"">$(Esc $s.Id)<span class=""item"">$(Format-Data $s.Data) · linha $($s.Linha)</span></td><td>$(Esc $s.Cliente)<span class=""item quebra"">$(Esc $s.Produto)</span></td><td>$origem</td><td class=""num"">$(Format-Moeda $s.Valor)</td><td>$status</td></tr>"
    }
    $vazio = if (-not $v.Linhas.Count) { '<tr><td colspan="5">Nenhuma venda no período.</td></tr>' } else { '' }
    $corpo = @"
<section class="filtros">
  <h1>Origem do faturamento</h1>
  <form method="get" action="/origem-vendas">
    <label>De <select name="de">$(& $opcoes $v.MesInicial)</select></label>
    <label>até <select name="ate">$(& $opcoes $v.MesFinal)</select></label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Período: $($script:NomesMesLongo[$v.MesInicial]) a $($script:NomesMesLongo[$v.MesFinal]) de 2026 · vendas até $(Format-Data $v.DataBase) · foto do CRM: $(Format-Data $v.DataBaseCrm).</p>
</section>
$parcial
<section class="bloco"><p>$resumo</p><p class="nota">Participação = valor das vendas faturadas com oportunidade vinculada ÷ valor de todas as vendas faturadas do período. Usamos o valor vendido, já com desconto, e não o valor estimado da oportunidade.</p></section>
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Faturamento total</span><span class="valor">$(Format-Moeda $v.Total)</span><span class="det">$($v.QtdFaturadas) vendas faturadas</span></div>
  <div class="cartao destaque"><span class="rotulo">Com origem no CRM · $pct</span><span class="valor">$(Format-Moeda $v.ComCrm)</span><span class="det">$($v.QtdComCrm) vendas vinculadas por ID Oportunidade</span></div>
  <div class="cartao"><span class="rotulo">Sem oportunidade · $pctSem</span><span class="valor">$(Format-Moeda $v.SemOportunidade)</span><span class="det">$($v.QtdSemOportunidade) vendas sem vínculo registrado</span></div>
</section>
<section class="bloco">
  <h2>O que são as vendas sem oportunidade?</h2>
  <p>São vendas cuja coluna <strong>ID Oportunidade está vazia</strong>. Nesta base didática, são descritas como <strong>vendas de balcão</strong>: faturamentos sem uma oportunidade registrada no CRM.</p>
  <p>Isso não significa ausência de trabalho comercial, nem que sejam somente peças. A composição abaixo mostra as categorias efetivamente vendidas sem esse vínculo.</p>
  <p class="nota">Não inferimos vínculos por cliente, produto ou nome. Uma venda sem ID continua sem oportunidade registrada, mesmo que exista uma oportunidade parecida. Este indicador mede a origem documentada no CRM; não é taxa de conversão do funil nem mede, sozinho, a influência de todo o trabalho comercial.</p>
  <table><thead><tr><th>Categoria sem oportunidade</th><th class="num">Vendas faturadas</th><th class="num">Faturamento</th></tr></thead><tbody>$($categorias -join "`n")$vazioCat</tbody><tfoot><tr><td>Total sem oportunidade</td><td class="num">$($v.QtdSemOportunidade)</td><td class="num">$(Format-Moeda $v.SemOportunidade)</td></tr></tfoot></table>
</section>
<section class="bloco">
  <h2>Conferência: cada venda e sua oportunidade</h2>
  <p class="nota">$($v.Linhas.Count) registros no período. As $($v.QtdCanceladas) canceladas, no valor de $(Format-Moeda $v.ValorCancelado), aparecem riscadas e não entram nos indicadores. Todos os vendedores estão incluídos, inclusive o histórico dos desligados.</p>
  <p class="fonte">Fontes: vendas_2026_jan-ago.xlsx, aba Vendas; crm_oportunidades.xlsx, aba Oportunidades. Os números de linha permitem conferir o vínculo nas planilhas.</p>
  <details><summary>Ver vendas e vínculos ($($v.Linhas.Count) registros)</summary>
  <table><thead><tr><th>Venda / data / linha</th><th>Cliente / produto</th><th>Oportunidade / etapa / origem</th><th class="num">Valor da venda</th><th>Status</th></tr></thead><tbody>$($linhas -join "`n")$vazio</tbody></table>
  </details>
</section>
<footer class="regras">REGRAS_NEGOCIO.md §0.1: vínculo pelo ID · §1: somente faturadas no realizado; vendas com e sem oportunidade contam no faturamento · §2.1: histórico dos desligados preservado. O filtro usa a data da venda, não a data de criação da oportunidade.</footer>
"@
    return New-Layout 'Origem do faturamento' $corpo $Usuario 'origem'
}
