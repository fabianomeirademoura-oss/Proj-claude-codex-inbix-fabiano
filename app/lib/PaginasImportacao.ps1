# HTML da tela de importação. Só formata: as decisões vêm de RegrasImportacao.ps1 (REGRAS_NEGOCIO.md §9).
# Usa os auxiliares de Paginas.ps1 (Esc, Format-*, New-Layout). Sem JavaScript (CSP).

function Format-ColunasObrigatorias([string]$Tipo) {
    return (($script:ColunasImportacao[$Tipo] | ForEach-Object { "<code>$(Esc $_)</code>" }) -join ', ')
}

function New-PaginaImportacao($Estado, $Usuario, [string]$Aviso) {
    $aviso = if ($Aviso) { "<p class=""parcial"" role=""alert"">$(Esc $Aviso)</p>" } else { '' }
    $importadosVendas = @($Estado.ArquivosVendas)
    $listaVendas = if ($importadosVendas.Count) { ($importadosVendas | ForEach-Object { "<span class=""item"">$(Esc $_)</span>" }) -join '' } else { '<span class="nota">nenhuma importação de vendas ainda</span>' }
    $linhasHist = foreach ($h in $Estado.Historico) {
        $contagem = if ($h.Tipo -eq 'vendas') { "$($h.Novas) novas · $($h.Atualizadas) atualizadas · $($h.Ignoradas) ignoradas" } else { "foto de $(Esc $h.DataFoto) · $($h.Substituidas) linhas substituídas" }
        "<tr><td class=""nw"">$(Esc $h.Quando)</td><td>$(Esc $h.Tipo)</td><td>$(Esc $h.Enviado)<span class=""item id"">gravado como $(Esc $h.Gravado)</span></td><td>$contagem</td><td>$(Esc $h.Login)</td></tr>"
    }
    $tabelaHist = if ($Estado.Historico.Count) {
        "<table><thead><tr><th>Quando</th><th>Tipo</th><th>Arquivo</th><th>Resultado</th><th>Quem</th></tr></thead><tbody>$($linhasHist -join "`n")</tbody></table>"
    } else { '<p class="nota">Nenhuma importação confirmada até agora.</p>' }
    $erroBase = if ($Estado.ErroBase) { "<p class=""parcial"">$(Esc $Estado.ErroBase)</p>" } else { '' }

    $corpo = @"
<section class="filtros">
  <h1>Importar arquivos</h1>
  <p class="fonte">Dois comportamentos diferentes (REGRAS_NEGOCIO.md §9): <strong>vendas</strong> se somam às existentes sem duplicar, pela chave <code>ID Venda</code>; <strong>estoque</strong> é uma fotografia, e o arquivo novo substitui a posição inteira daquela data.</p>
</section>
$aviso
$erroBase
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Vendas em vigor</span><span class="valor">$($Estado.QtdVendas)</span><span class="det">data-base $(Format-Data $Estado.DataBaseVendas) · base + importações</span></div>
  <div class="cartao"><span class="rotulo">Foto de estoque em vigor</span><span class="valor">$(Format-Data $Estado.DataFoto)</span><span class="det">$(Esc $Estado.ArquivoFoto)</span></div>
</section>
<section class="bloco">
  <h2>1. Escolha o arquivo</h2>
  <form method="post" action="/importar" enctype="multipart/form-data" class="form-importar">
    <fieldset>
      <legend>O arquivo é de</legend>
      <label class="check"><input type="radio" name="tipo" value="vendas" required> <span><strong>Vendas</strong>: acrescenta vendas novas e atualiza as que mudaram (aba <code>Vendas</code>)</span></label>
      <label class="check"><input type="radio" name="tipo" value="estoque" required> <span><strong>Estoque</strong>: substitui a foto inteira (aba <code>Estoque</code>, nome <code>estoque_AAAA-MM-DD.xlsx</code>)</span></label>
    </fieldset>
    <label>Arquivo Excel (.xlsx) <input type="file" name="arquivo" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required></label>
    <button type="submit">Analisar arquivo</button>
    <p class="nota">Nesta etapa nada é gravado. Você verá o resumo do que vai acontecer e decide se confirma.</p>
  </form>
  <details>
    <summary>Colunas obrigatórias</summary>
    <p><strong>Vendas ($($script:ColunasImportacao.vendas.Count)):</strong> $(Format-ColunasObrigatorias 'vendas').</p>
    <p><strong>Estoque ($($script:ColunasImportacao.estoque.Count)):</strong> $(Format-ColunasObrigatorias 'estoque').</p>
    <p class="nota">O nome do cabeçalho precisa ser exatamente este. Colunas a mais são ignoradas. Se faltar alguma, o arquivo é recusado e a tela diz qual faltou.</p>
  </details>
</section>
<section class="bloco">
  <h2>Importações de vendas aplicadas</h2>
  <p>$listaVendas</p>
</section>
<section class="bloco">
  <h2>Histórico</h2>
  $tabelaHist
  <p class="nota">Os arquivos confirmados ficam guardados sem alteração em <code>dados/importacoes/</code>. As planilhas originais nunca são alteradas. Para desfazer uma importação, retire o arquivo daquela pasta.</p>
</section>
"@
    return New-Layout 'Importar arquivos' $corpo $Usuario 'importar'
}

function New-PaginaRecusaImportacao($Resumo, $Usuario) {
    $itens = ($Resumo.Erros | Select-Object -First 50 | ForEach-Object { "<li>$(Esc $_)</li>" }) -join "`n"
    $mais = if ($Resumo.Erros.Count -gt 50) { "<p class=""nota"">… e mais $($Resumo.Erros.Count - 50) problema(s).</p>" } else { '' }
    $corpo = @"
<section class="aviso-erro" role="alert">
  <h1>Arquivo recusado</h1>
  <p><strong>$(Esc $Resumo.NomeOriginal)</strong>, enviado como <strong>$(Esc $Resumo.Tipo)</strong>. <strong>Nada foi gravado.</strong></p>
  <ul>$itens</ul>
  $mais
  <p><a class="botao" href="/importar">Escolher outro arquivo</a></p>
</section>
"@
    return New-Layout 'Arquivo recusado · Importar' $corpo $Usuario 'importar'
}

function Get-BotoesConfirmacao([string]$Token, [string]$Rotulo) {
    return @"
<div class="acoes">
  <form method="post" action="/importar/confirmar"><input type="hidden" name="token" value="$(Esc $Token)"><button type="submit">$(Esc $Rotulo)</button></form>
  <form method="post" action="/importar/cancelar"><input type="hidden" name="token" value="$(Esc $Token)"><button type="submit" class="secundario">Cancelar</button></form>
</div>
"@
}

function New-PaginaResumoImportacao($Resumo, [string]$Token, $Usuario) {
    $r = $Resumo
    if ($r.Tipo -eq 'vendas') {
        $periodo = if ($r.PrimeiraData) { "de $(Format-Data $r.PrimeiraData) a $(Format-Data $r.UltimaData)" } else { '' }
        $linhasAt = foreach ($a in ($r.Atualizadas | Select-Object -First 100)) {
            $valor = { param($col, $v) if ($col -in $script:ColunasDinheiroVenda -and $null -ne (ConvertTo-Dinheiro $v)) { Format-Moeda (ConvertTo-Dinheiro $v) } elseif ([string]$v -eq '') { '(vazio)' } else { [string]$v } }
            $mud = ($a.Mudancas | ForEach-Object { "<span class=""item""><strong>$(Esc $_.Coluna)</strong>: $(Esc (& $valor $_.Coluna $_.De)) → $(Esc (& $valor $_.Coluna $_.Para))</span>" }) -join ''
            "<tr><td class=""nw"">$(Esc $a.Id)</td><td class=""num"">$($a.Linha)</td><td>$mud</td></tr>"
        }
        $tabelaAt = if ($r.Atualizadas.Count) {
            "<h3>Vendas que serão atualizadas</h3><p class=""nota"">Mesmo ID Venda com algum valor diferente: a linha nova substitui a anterior (REGRAS_NEGOCIO.md §1).</p><table><thead><tr><th>ID Venda</th><th class=""num"">Linha no arquivo</th><th>O que muda</th></tr></thead><tbody>$($linhasAt -join "`n")</tbody></table>"
        } else { '' }
        $nadaMuda = if (-not $r.Novas -and -not $r.Atualizadas.Count) { '<p class="parcial">Todas as vendas deste arquivo já existem iguais. Confirmar não muda nenhum número; só fica registrado no histórico.</p>' } else { '' }
        $corpo = @"
<section class="filtros">
  <h1>Resumo da importação de vendas</h1>
  <p class="fonte"><strong>$(Esc $r.NomeOriginal)</strong>: $($r.LinhasArquivo) linhas $periodo. Nada foi gravado ainda.</p>
</section>
<section class="cartoes">
  <div class="cartao destaque"><span class="rotulo">Linhas novas</span><span class="valor">$($r.Novas)</span><span class="det">$(Format-Moeda $r.FaturadoNovas) faturados · $($r.NovasCanceladas) canceladas</span></div>
  <div class="cartao"><span class="rotulo">Atualizadas (substituem a anterior)</span><span class="valor">$($r.Atualizadas.Count)</span><span class="det">ID já existente com dados diferentes</span></div>
  <div class="cartao neutro"><span class="rotulo">Ignoradas</span><span class="valor">$($r.Ignoradas)</span><span class="det">ID já existente, tudo igual</span></div>
</section>
$nadaMuda
<section class="bloco">
  <h2>O que muda no painel</h2>
  <table>
    <tbody>
      <tr><td>Vendas em vigor</td><td class="num">$($r.VendasAntes) → <strong>$($r.VendasDepois)</strong></td></tr>
      <tr><td>Data-base das vendas</td><td class="num">$(Format-Data $r.DataBaseAntes) → <strong>$(Format-Data $r.DataBaseDepois)</strong></td></tr>
    </tbody>
  </table>
  $tabelaAt
  $(Get-BotoesConfirmacao $Token 'Confirmar e gravar')
  <p class="nota">Ao confirmar, o arquivo é guardado sem alteração em <code>dados/importacoes/vendas/</code> e o painel passa a considerá-lo. A validação é refeita na hora da confirmação.</p>
</section>
"@
        return New-Layout 'Resumo · Importar vendas' $corpo $Usuario 'importar'
    }

    $saem = @($r.SemEstoqueAntes | Where-Object { $_ -notin $r.SemEstoqueDepois })
    $entram = @($r.SemEstoqueDepois | Where-Object { $_ -notin $r.SemEstoqueAntes })
    $modo = if ($r.MesmaData) { "A foto em vigor já é de $(Format-Data $r.DataNova): este arquivo <strong>substitui a posição inteira daquela data</strong>." } else { "A foto em vigor é de $(Format-Data $r.DataAtual). Este arquivo, de $(Format-Data $r.DataNova), <strong>passa a ser a foto em vigor</strong> e substitui a anterior inteira." }
    $corpo = @"
<section class="filtros">
  <h1>Resumo da importação de estoque</h1>
  <p class="fonte"><strong>$(Esc $r.NomeOriginal)</strong>. Nada foi gravado ainda.</p>
</section>
<p class="definicao">$modo</p>
<section class="cartoes">
  <div class="cartao destaque"><span class="rotulo">Linhas substituídas</span><span class="valor">$($r.LinhasSubstituidas)</span><span class="det">a foto de $(Format-Data $r.DataAtual) inteira ($(Esc $r.ArquivoAtual))</span></div>
  <div class="cartao"><span class="rotulo">Linhas na foto nova</span><span class="valor">$($r.LinhasNova)</span><span class="det">$($r.QtdMudou) com quantidade diferente</span></div>
  <div class="cartao neutro"><span class="rotulo">Linhas novas / que deixam de existir</span><span class="valor">$($r.LinhasNovasChave) / $($r.LinhasSomem)</span><span class="det">as que deixam de existir contam como zero</span></div>
</section>
<section class="bloco">
  <h2>O que muda no painel</h2>
  <table>
    <thead><tr><th>Indicador</th><th class="num">Hoje ($(Format-Data $r.DataAtual))</th><th class="num">Depois ($(Format-Data $r.DataNova))</th></tr></thead>
    <tbody>
      <tr><td>Valor imobilizado</td><td class="num">$(Format-Moeda $r.ImobilizadoAntes)</td><td class="num">$(Format-Moeda $r.ImobilizadoDepois)</td></tr>
      <tr><td>Produtos sem estoque</td><td class="num">$($r.SemEstoqueAntes.Count)</td><td class="num">$($r.SemEstoqueDepois.Count)</td></tr>
      <tr><td>Linhas abaixo do mínimo</td><td class="num">$($r.AbaixoAntes)</td><td class="num">$($r.AbaixoDepois)</td></tr>
    </tbody>
  </table>
  <p class="nota">Saem da lista de sem estoque: $(if ($saem.Count) { Esc ($saem -join ', ') } else { 'nenhum' }). Entram: $(if ($entram.Count) { Esc ($entram -join ', ') } else { 'nenhum' }).</p>
  $(Get-BotoesConfirmacao $Token 'Confirmar e substituir a foto')
  <p class="nota">Ao confirmar, o arquivo é guardado sem alteração em <code>dados/importacoes/estoque/</code>. A foto anterior não é apagada, só deixa de ser a foto em vigor.</p>
</section>
"@
    return New-Layout 'Resumo · Importar estoque' $corpo $Usuario 'importar'
}

function New-PaginaImportacaoGravada($Resumo, $Usuario) {
    $r = $Resumo
    $texto = if ($r.Tipo -eq 'vendas') {
        "$($r.Novas) vendas novas, $($r.Atualizadas.Count) atualizadas e $($r.Ignoradas) ignoradas. Data-base das vendas: $(Format-Data $r.DataBaseDepois)."
    } else {
        "A foto de estoque de $(Format-Data $r.DataNova) está em vigor. $($r.LinhasSubstituidas) linhas da foto anterior foram substituídas."
    }
    $destino = if ($r.Tipo -eq 'vendas') { '<a class="botao" href="/">Ver vendas e metas</a>' } else { '<a class="botao" href="/estoque">Ver estoque</a>' }
    $corpo = @"
<section class="bloco">
  <h1>Importação gravada</h1>
  <p><strong>$(Esc $r.NomeOriginal)</strong> foi guardado como <code>$(Esc $r.ArquivoGravado)</code>.</p>
  <p>$texto</p>
  <div class="acoes">$destino <a href="/importar">Importar outro arquivo</a></div>
</section>
"@
    return New-Layout 'Importação gravada' $corpo $Usuario 'importar'
}
