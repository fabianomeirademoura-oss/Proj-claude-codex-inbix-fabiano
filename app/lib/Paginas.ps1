# Geração de HTML e CSV. Só formata: todo número vem de Regras.ps1.

$script:PtBr = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
$script:NomesMes = @('', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez')
$script:NomesMesLongo = @('', 'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro')

function Esc($Texto) { return [Net.WebUtility]::HtmlEncode([string]$Texto) }

function Format-Moeda($Valor) {
    if ($null -eq $Valor) { return '—' }
    return 'R$ ' + ([decimal]$Valor).ToString('N2', $script:PtBr)
}

function Format-Pct($Fracao) {
    if ($null -eq $Fracao) { return 'sem meta' }
    return ([decimal]$Fracao * 100).ToString('N1', $script:PtBr) + '%'
}

function Format-Data($Data) { if ($Data) { return $Data.ToString('dd/MM/yyyy') } }

function Format-Meses([int[]]$Meses) {
    if (-not $Meses -or $Meses.Count -eq 0) { return 'nenhum' }
    $partes = @()
    $inicio = $Meses[0]; $anterior = $Meses[0]
    for ($i = 1; $i -le $Meses.Count; $i++) {
        $atual = if ($i -lt $Meses.Count) { $Meses[$i] } else { -1 }
        if ($atual -ne $anterior + 1) {
            $partes += if ($inicio -eq $anterior) { $script:NomesMes[$inicio] } else { "$($script:NomesMes[$inicio])–$($script:NomesMes[$anterior])" }
            $inicio = $atual
        }
        $anterior = $atual
    }
    return ($partes -join ', ')
}

function Get-Marcas($Vendedor) {
    # REGRAS_NEGOCIO.md §2.3 e §3.3: desligado e admitido no ano aparecem marcados.
    if (-not $Vendedor.Ativo) { return "<span class=""marca desligado"">desligado em $(Format-Data $Vendedor.Desligamento)</span>" }
    if ($Vendedor.Admissao -and $Vendedor.Admissao.Year -eq 2026) { return "<span class=""marca admitido"">admissão em $($Vendedor.Admissao.ToString('MM/yyyy'))</span>" }
    return ''
}

function Get-Medidor($Fracao) {
    if ($null -eq $Fracao) { return '' }
    $v = [math]::Min([double]$Fracao, 1.5).ToString('0.0000', [Globalization.CultureInfo]::InvariantCulture)
    return "<meter min=""0"" max=""1.5"" low=""0.8"" high=""0.9999"" optimum=""1.2"" value=""$v""></meter>"
}

function Get-QueryPeriodo($Desempenho, [bool]$IncluirDesligados) {
    $q = "de=$($Desempenho.MesInicial)&amp;ate=$($Desempenho.MesFinal)"
    if ($IncluirDesligados) { $q += '&amp;desligados=1' }
    return $q
}

function New-Layout([string]$Titulo, [string]$Corpo, $Usuario, [string]$Secao = 'vendas') {
    $ativo = { param($s) if ($s -eq $Secao) { ' class="ativo" aria-current="page"' } else { '' } }
    $topo = if ($Usuario) {
        @"
<header class="topo">
  <span class="marca-app">Horizonte Máquinas</span>
  <nav class="nav" aria-label="Seções do painel">
    <a href="/"$(& $ativo 'vendas')>Vendas e metas</a>
    <a href="/origem-vendas"$(& $ativo 'origem')>Origem do faturamento</a>
    <a href="/pipeline"$(& $ativo 'pipeline')>Pipeline</a>
    <a href="/funil"$(& $ativo 'funil')>Funil</a>
    <a href="/estoque"$(& $ativo 'estoque')>Estoque</a>
    $(if ($Usuario.perfil -eq 'diretoria') { "<a href=""/importar""$(& $ativo 'importar')>Importar</a>" })
  </nav>
  <div class="usuario"><span class="nome-usuario">$(Esc $Usuario.nome)</span>
    <form method="post" action="/sair"><button type="submit" class="link">Sair</button></form>
  </div>
</header>
"@
    } else { '' }
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#ffffff" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#252523" media="(prefers-color-scheme: dark)">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Horizonte">
<title>$(Esc $Titulo)</title>
<link rel="manifest" href="/manifest.webmanifest">
<link rel="icon" type="image/png" href="/icones/favicone-32.png">
<link rel="apple-touch-icon" href="/icones/apple-touch-icon.png">
<link rel="stylesheet" href="/estilo.css">
<script src="/painel.js" defer></script>
</head>
<body>
$topo
<main>
$Corpo
</main>
</body>
</html>
"@
}

function New-PaginaLogin([string]$Mensagem, [string]$LoginDigitado) {
    $aviso = if ($Mensagem) { "<p class=""erro"" role=""alert"">$(Esc $Mensagem)</p>" } else { '' }
    $corpo = @"
<section class="login">
  <h1>Horizonte Máquinas</h1>
  <p class="sub">Desempenho comercial 2026</p>
  $aviso
  <form method="post" action="/entrar">
    <label>Usuário (e-mail)<input name="login" type="text" autocomplete="username" required value="$(Esc $LoginDigitado)"></label>
    <label>Senha<input name="senha" type="password" autocomplete="current-password" required></label>
    <button type="submit">Entrar</button>
  </form>
</section>
"@
    return New-Layout 'Entrar · Horizonte Máquinas' $corpo $null
}

function New-PaginaErroImportacao([string[]]$Erros, $Usuario, [string]$Secao = 'vendas') {
    $itens = ($Erros | Select-Object -First 50 | ForEach-Object { "<li>$(Esc $_)</li>" }) -join "`n"
    $corpo = @"
<section class="aviso-erro">
  <h1>Os dados não puderam ser carregados</h1>
  <p>A importação encontrou $($Erros.Count) problema(s). Nenhum número é exibido até que as planilhas sejam corrigidas (regras do REGRAS_NEGOCIO.md).</p>
  <ul>$itens</ul>
</section>
"@
    return New-Layout 'Erro de importação' $corpo $Usuario $Secao
}

function New-PaginaPainel($Base, $Desempenho, $Usuario, [bool]$IncluirDesligados) {
    $d = $Desempenho
    $e = $d.Empresa
    $q = Get-QueryPeriodo $d $IncluirDesligados
    $meses = Get-MesesDisponiveis $Base
    $opcoes = {
        param($selecionado)
        ($meses | ForEach-Object { $s = if ($_ -eq $selecionado) { ' selected' } else { '' }; "<option value=""$_""$s>$($script:NomesMesLongo[$_])</option>" }) -join ''
    }
    $marcado = if ($IncluirDesligados) { ' checked' } else { '' }
    $periodo = "$($script:NomesMesLongo[$d.MesInicial]) a $($script:NomesMesLongo[$d.MesFinal]) de 2026"
    $fontes = ($Base.Arquivos | ForEach-Object { "$(Esc $_.Nome) ($($_.Modificado.ToString('dd/MM/yyyy HH:mm')))" }) -join ' · '
    $euId = $Usuario.idVendedor

    $parcial = if ($d.Parcial) {
        "<p class=""parcial"">Resultado parcial até $(Format-Data $d.DataBase): $($script:NomesMesLongo[$d.MesFinal]) ainda não terminou e é comparado com a meta cheia do mês.</p>"
    } else { '' }

    $foraMeta = if ($e.ForaMesesMeta -ne 0) {
        "<p class=""nota"">$(Format-Moeda $e.ForaMesesMeta) foram vendidos em meses em que o vendedor não tinha meta: entram no faturamento, não no atingimento (REGRAS_NEGOCIO.md §4.1).</p>"
    } else { '' }

    # Ranking por faturamento
    $rf = Get-RankingFaturamento $d -IncluirDesligados:$IncluirDesligados
    $pos = 0
    $linhasFat = foreach ($l in $rf) {
        $pos++
        $v = $l.Vendedor
        $cls = if ($v.Id -eq $euId) { ' class="voce"' } else { '' }
        $part = if ($e.Realizado -gt 0) { Format-Pct ($l.Realizado / $e.Realizado) } else { '—' }
        "<tr$cls><td class=""pos"">$pos</td><td><a href=""/vendedor?id=$(Esc $v.Id)&amp;$q"">$(Esc $v.Nome)</a> $(Get-Marcas $v)</td><td>$(Esc $v.Filial)</td><td class=""num"">$($l.MesesComMeta.Count)</td><td class=""num"">$(Format-Moeda $l.Realizado)</td><td class=""num"">$part</td></tr>"
    }

    # Ranking por atingimento
    $ra = Get-RankingAtingimento $d -IncluirDesligados:$IncluirDesligados
    $pos = 0
    $linhasAt = foreach ($l in $ra) {
        $pos++
        $v = $l.Vendedor
        $cls = if ($v.Id -eq $euId) { ' class="voce"' } else { '' }
        "<tr$cls><td class=""pos"">$pos</td><td><a href=""/vendedor?id=$(Esc $v.Id)&amp;$q"">$(Esc $v.Nome)</a> $(Get-Marcas $v)</td><td class=""num"">$(Format-Moeda $l.Meta)</td><td class=""num"">$(Format-Moeda $l.RealizadoMesesMeta)</td><td class=""num""><span class=""at"">$(Get-Medidor $l.Atingimento) $(Format-Pct $l.Atingimento)</span></td></tr>"
    }
    $semMeta = @($d.Linhas | Where-Object { ($IncluirDesligados -or $_.Vendedor.Ativo) -and $null -eq $_.Atingimento })
    $notaSemMeta = if ($semMeta.Count) {
        "<p class=""nota"">Sem meta no período (fora do ranking de atingimento): $((($semMeta | ForEach-Object { Esc $_.Vendedor.Nome }) -join ', ')).</p>"
    } else { '' }
    $foraRanking = @($d.Linhas | Where-Object { -not $_.Vendedor.Ativo })
    $notaDesligados = if (-not $IncluirDesligados -and $foraRanking.Count) {
        "<p class=""nota"">Desligados não entram nos rankings (REGRAS_NEGOCIO.md §2.3): $((($foraRanking | ForEach-Object { Esc $_.Vendedor.Nome }) -join ', ')). Os números deles estão no quadro de conferência e nos totais.</p>"
    } else { '' }

    # Quadro de conferência: todos os vendedores, total = empresa
    $temFora = @($d.Linhas | Where-Object { $_.ForaMesesMeta -ne 0 }).Count -gt 0
    $cabFora = if ($temFora) { '<th class="num">Realizado fora dos meses c/ meta</th>' } else { '' }
    $rotuloReal = if ($temFora) { 'Realizado nos meses c/ meta' } else { 'Realizado' }
    $linhasQuadro = foreach ($l in ($d.Linhas | Sort-Object { $_.Vendedor.Id })) {
        $v = $l.Vendedor
        $cls = if ($v.Id -eq $euId) { ' class="voce"' } else { '' }
        $celFora = if ($temFora) { "<td class=""num"">$(Format-Moeda $l.ForaMesesMeta)</td>" } else { '' }
        "<tr$cls><td>$(Esc $v.Id)</td><td><a href=""/vendedor?id=$(Esc $v.Id)&amp;$q"">$(Esc $v.Nome)</a> $(Get-Marcas $v)</td><td>$(Esc $v.Filial)</td><td class=""nw"">$(Format-Meses $l.MesesComMeta)</td><td class=""num"">$(Format-Moeda $l.Meta)</td><td class=""num"">$(Format-Moeda $l.RealizadoMesesMeta)</td>$celFora<td class=""num"">$(Format-Pct $l.Atingimento)</td><td class=""num"">$($l.QtdVendas)</td></tr>"
    }
    $totFora = if ($temFora) { "<td class=""num"">$(Format-Moeda $e.ForaMesesMeta)</td>" } else { '' }

    $corpo = @"
<section class="filtros">
  <form method="get" action="/">
    <label>De <select name="de">$(& $opcoes $d.MesInicial)</select></label>
    <label>até <select name="ate">$(& $opcoes $d.MesFinal)</select></label>
    <label class="check"><input type="checkbox" name="desligados" value="1"$marcado> Incluir desligados nos rankings</label>
    <button type="submit">Aplicar</button>
  </form>
  <p class="fonte">Período: <strong>$periodo</strong> · data-base das vendas: $(Format-Data $d.DataBase) · fontes: $fontes</p>
</section>
$parcial
<section class="cartoes">
  <div class="cartao"><span class="rotulo">Realizado (vendas faturadas)</span><span class="valor">$(Format-Moeda $e.Realizado)</span><span class="det">$($e.QtdVendas) vendas · ticket médio $(Format-Moeda $e.TicketMedio)</span></div>
  <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">$(Format-Moeda $e.Meta)</span><span class="det">soma só dos meses em que cada vendedor tinha meta</span></div>
  <div class="cartao destaque"><span class="rotulo">Atingimento da empresa</span><span class="valor">$(Format-Pct $e.Atingimento)</span><span class="det">realizado ÷ meta (soma ÷ soma)</span></div>
  <div class="cartao neutro"><span class="rotulo">Canceladas (fora do realizado)</span><span class="valor">$(Format-Moeda $d.ValorCanceladas)</span><span class="det">$($d.QtdCanceladas) vendas com Status = Cancelada</span></div>
</section>
$foraMeta
<section class="rankings">
  <div class="bloco">
    <h2>Ranking por faturamento</h2>
    <table>
      <thead><tr><th>#</th><th>Vendedor</th><th>Filial</th><th class="num">Meses c/ meta</th><th class="num">Realizado</th><th class="num">% do total</th></tr></thead>
      <tbody>$($linhasFat -join "`n")</tbody>
    </table>
  </div>
  <div class="bloco">
    <h2>Ranking por atingimento</h2>
    <table>
      <thead><tr><th>#</th><th>Vendedor</th><th class="num">Meta</th><th class="num">Realizado</th><th class="num">Atingimento</th></tr></thead>
      <tbody>$($linhasAt -join "`n")</tbody>
    </table>
    $notaSemMeta
  </div>
</section>
$notaDesligados
<section class="bloco">
  <div class="cab-bloco">
    <h2>Quadro de conferência por vendedor</h2>
    <a class="botao" href="/conferencia.csv?$q">Baixar CSV</a>
  </div>
  <p class="nota">Todos os vendedores, inclusive desligados. A linha de total é a da empresa e deve bater com a soma das planilhas. Clique num nome para ver mês a mês e venda a venda.</p>
  <table>
    <thead><tr><th>ID</th><th>Vendedor</th><th>Filial</th><th>Meses c/ meta</th><th class="num">Meta</th><th class="num">$rotuloReal</th>$cabFora<th class="num">Atingimento</th><th class="num">Vendas</th></tr></thead>
    <tbody>$($linhasQuadro -join "`n")</tbody>
    <tfoot><tr><td></td><td>Total da empresa</td><td></td><td></td><td class="num">$(Format-Moeda $e.Meta)</td><td class="num">$(Format-Moeda $e.RealizadoMesesMeta)</td>$totFora<td class="num">$(Format-Pct $e.Atingimento)</td><td class="num">$($e.QtdVendas)</td></tr></tfoot>
  </table>
</section>
<footer class="regras">
  <strong>Regras aplicadas (REGRAS_NEGOCIO.md):</strong>
  realizado = só vendas com Status = Faturada (§1) ·
  meta e realizado somados só nos meses em que o vendedor tem meta (§4.1) ·
  empresa = soma ÷ soma, com desligados (§4.2) ·
  rankings só com ativos, salvo filtro (§2.3) ·
  filial = filial do vendedor (§1).
</footer>
"@
    return New-Layout "Desempenho comercial · $periodo" $corpo $Usuario
}

function New-PaginaVendedor($Base, $Desempenho, $Detalhe, $Linha, $Usuario, [bool]$IncluirDesligados) {
    $v = $Linha.Vendedor
    $q = Get-QueryPeriodo $Desempenho $IncluirDesligados
    $situacao = if ($v.Ativo) { "Ativo desde $(Format-Data $v.Admissao)" } else { "Desligado em $(Format-Data $v.Desligamento) (admitido em $(Format-Data $v.Admissao))" }

    $linhasMes = foreach ($m in $Detalhe.Meses) {
        $meta = if ($null -eq $m.Meta) { '<span class="nota">sem meta</span>' } else { Format-Moeda $m.Meta }
        $canc = if ($m.QtdCanceladas) { "$($m.QtdCanceladas) · $(Format-Moeda $m.ValorCanceladas)" } else { '—' }
        "<tr><td>$($script:NomesMesLongo[$m.Mes])</td><td class=""num"">$meta</td><td class=""num"">$(Format-Moeda $m.Realizado)</td><td class=""num"">$(if ($null -ne $m.Meta) { Format-Pct $m.Atingimento } else { '—' })</td><td class=""num"">$($m.QtdVendas)</td><td class=""num"">$canc</td></tr>"
    }
    $linhasVenda = foreach ($s in $Detalhe.Vendas) {
        $real = Test-VendaRealizada $s
        $cls = if ($real) { '' } else { ' class="cancelada"' }
        $status = if ($real) { 'Faturada' } else { 'Cancelada <span class="marca desligado">fora do realizado</span>' }
        "<tr$cls><td>$(Esc $s.Id)</td><td class=""num"">$($s.Linha)</td><td>$(Format-Data $s.Data)</td><td>$(Esc $s.Cliente)</td><td>$(Esc $s.Produto)</td><td>$status</td><td class=""num"">$(Format-Moeda $s.Valor)</td></tr>"
    }

    $corpo = @"
<p><a href="/?$q">← Voltar ao painel</a></p>
<section class="bloco">
  <h1>$(Esc $v.Nome) <span class="id">$(Esc $v.Id)</span> $(Get-Marcas $v)</h1>
  <p class="fonte">$(Esc $v.Filial) · $situacao · período: $($script:NomesMesLongo[$Desempenho.MesInicial]) a $($script:NomesMesLongo[$Desempenho.MesFinal]) de 2026</p>
  <section class="cartoes">
    <div class="cartao"><span class="rotulo">Meta do período</span><span class="valor">$(Format-Moeda $Linha.Meta)</span><span class="det">meses com meta: $(Format-Meses $Linha.MesesComMeta)</span></div>
    <div class="cartao"><span class="rotulo">Realizado</span><span class="valor">$(Format-Moeda $Linha.RealizadoMesesMeta)</span><span class="det">$($Linha.QtdVendas) vendas faturadas</span></div>
    <div class="cartao destaque"><span class="rotulo">Atingimento</span><span class="valor">$(Format-Pct $Linha.Atingimento)</span><span class="det">realizado ÷ meta dos mesmos meses</span></div>
  </section>
</section>
<section class="bloco">
  <h2>Mês a mês</h2>
  <table>
    <thead><tr><th>Mês</th><th class="num">Meta</th><th class="num">Realizado</th><th class="num">Atingimento</th><th class="num">Vendas</th><th class="num">Canceladas</th></tr></thead>
    <tbody>$($linhasMes -join "`n")</tbody>
    <tfoot><tr><td>Período</td><td class="num">$(Format-Moeda $Linha.Meta)</td><td class="num">$(Format-Moeda $Linha.Realizado)</td><td class="num">$(Format-Pct $Linha.Atingimento)</td><td class="num">$($Linha.QtdVendas)</td><td></td></tr></tfoot>
  </table>
</section>
<section class="bloco">
  <h2>Vendas do período</h2>
  <p class="nota">"Linha" é o número da linha em vendas_2026_jan-ago.xlsx, para achar a venda direto na planilha.</p>
  <table>
    <thead><tr><th>ID Venda</th><th class="num">Linha</th><th>Data</th><th>Cliente</th><th>Produto</th><th>Status</th><th class="num">Valor total</th></tr></thead>
    <tbody>$($linhasVenda -join "`n")</tbody>
  </table>
</section>
"@
    return New-Layout "$($v.Nome) · Desempenho" $corpo $Usuario
}

function New-CsvConferencia($Desempenho) {
    $f = { param($x) if ($null -eq $x) { '' } else { ([decimal]$x).ToString('0.00', $script:PtBr) } }
    $p = { param($x) if ($null -eq $x) { 'sem meta' } else { ([decimal]$x * 100).ToString('0.00', $script:PtBr) } }
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine('ID Vendedor;Vendedor;Filial;Status;Meses com meta;Meta;Realizado nos meses com meta;Realizado fora dos meses com meta;Realizado total;Atingimento (%);Vendas faturadas')
    foreach ($l in ($Desempenho.Linhas | Sort-Object { $_.Vendedor.Id })) {
        $v = $l.Vendedor
        [void]$sb.AppendLine(("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $v.Id, $v.Nome, $v.Filial, $v.Status, (Format-Meses $l.MesesComMeta), (& $f $l.Meta), (& $f $l.RealizadoMesesMeta), (& $f $l.ForaMesesMeta), (& $f $l.Realizado), (& $p $l.Atingimento), $l.QtdVendas))
    }
    $e = $Desempenho.Empresa
    [void]$sb.AppendLine(("TOTAL;Empresa;;;;{0};{1};{2};{3};{4};{5}" -f (& $f $e.Meta), (& $f $e.RealizadoMesesMeta), (& $f $e.ForaMesesMeta), (& $f $e.Realizado), (& $p $e.Atingimento), $e.QtdVendas))
    [void]$sb.AppendLine(("CANCELADAS (fora do realizado);;;;;;;;{0};;{1}" -f (& $f $Desempenho.ValorCanceladas), $Desempenho.QtdCanceladas))
    return $sb.ToString()
}

function Get-Css {
    return @'
:root{--fundo:#f6f5f1;--papel:#fff;--texto:#1f1e1b;--suave:#6b6a64;--linha:#e3e1da;--acento:#0f6e56;--acento-claro:#e1f5ee;--alerta:#993c1d;--alerta-claro:#faece7;--aviso:#854f0b;--aviso-claro:#faeeda}
@media (prefers-color-scheme:dark){:root{--fundo:#1b1b19;--papel:#252523;--texto:#ecebe6;--suave:#a3a29b;--linha:#3a3936;--acento:#5dcaa5;--acento-claro:#0e3b30;--alerta:#f0997b;--alerta-claro:#4a1b0c;--aviso:#fac775;--aviso-claro:#412402}}
*{box-sizing:border-box}
body{margin:0;background:var(--fundo);color:var(--texto);font:15px/1.5 "Segoe UI",system-ui,sans-serif}
main{max-width:1180px;margin:0 auto;padding:16px}
a{color:var(--acento)}
h1{font-size:22px;font-weight:600;margin:0 0 4px}h2{font-size:17px;font-weight:600;margin:0 0 10px}
.topo{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 16px;background:var(--papel);border-bottom:1px solid var(--linha)}
.marca-app{font-weight:600;text-decoration:none;color:var(--texto)}
.nav{display:flex;gap:18px;align-items:center;flex-wrap:wrap;flex:1}
.nav a{color:var(--suave);text-decoration:none;padding:4px 0;border-bottom:2px solid transparent}
.nav a.ativo{color:var(--texto);border-bottom-color:var(--acento)}
h3{font-size:15px;font-weight:600;margin:20px 0 6px}
.forte{font-weight:600}
.item{display:block;white-space:nowrap;font-size:13px}
.cartao.alerta{background:var(--alerta-claro);border-color:var(--alerta)}
.cartao.alerta .valor{color:var(--alerta)}
.usuario{display:flex;gap:12px;align-items:center;color:var(--suave)}
.usuario form{margin:0}
button,.botao{font:inherit;background:var(--acento);color:var(--papel);border:0;border-radius:6px;padding:7px 14px;cursor:pointer;text-decoration:none;display:inline-block}
button.link{background:none;color:var(--acento);padding:0;text-decoration:underline}
.filtros form{display:flex;flex-wrap:wrap;gap:12px;align-items:center}
select,input[type=text],input[type=password]{font:inherit;padding:6px 8px;border:1px solid var(--linha);border-radius:6px;background:var(--papel);color:var(--texto)}
.check{display:flex;gap:6px;align-items:center}
.fonte,.nota{color:var(--suave);font-size:13px;margin:8px 0}
.cartoes{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin:16px 0}
.cartao{background:var(--papel);border:1px solid var(--linha);border-radius:10px;padding:14px 16px;display:flex;flex-direction:column;gap:2px}
.cartao .rotulo{font-size:13px;color:var(--suave)}.cartao .valor{font-size:24px;font-weight:600;font-variant-numeric:tabular-nums}.cartao .det{font-size:12px;color:var(--suave)}
.cartao.destaque{background:var(--acento-claro);border-color:var(--acento)}
.cartao.neutro .valor{color:var(--suave)}
.rankings{display:grid;grid-template-columns:repeat(auto-fit,minmax(460px,1fr));gap:16px}
.bloco{background:var(--papel);border:1px solid var(--linha);border-radius:10px;padding:16px;margin-bottom:16px;overflow-x:auto}
.cab-bloco{display:flex;justify-content:space-between;align-items:center;gap:12px}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;font-weight:600;color:var(--suave);font-size:12px;border-bottom:1px solid var(--linha);padding:6px 8px}
td{padding:6px 8px;border-bottom:1px solid var(--linha);vertical-align:middle}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.pos{color:var(--suave);width:2em}
tfoot td{font-weight:600;border-top:2px solid var(--texto);border-bottom:0}
tr.voce td{background:var(--acento-claro)}
tr.cancelada td{color:var(--suave);text-decoration:line-through}
tr.cancelada td .marca{text-decoration:none}
.at{display:inline-flex;gap:8px;align-items:center}
.nw,th{white-space:nowrap}
meter{width:70px;height:10px}
.marca{font-size:11px;padding:1px 7px;border-radius:999px;white-space:nowrap;display:inline-block}
.marca.desligado{background:var(--alerta-claro);color:var(--alerta)}
.marca.admitido{background:var(--aviso-claro);color:var(--aviso)}
.id{font-size:14px;color:var(--suave);font-weight:400}
.parcial{background:var(--aviso-claro);color:var(--aviso);padding:10px 14px;border-radius:8px}
.regras{font-size:12px;color:var(--suave);padding:8px 0 24px}
.login{max-width:360px;margin:12vh auto;background:var(--papel);border:1px solid var(--linha);border-radius:12px;padding:28px}
.login .sub{color:var(--suave);margin:0 0 16px}
.login form{display:flex;flex-direction:column;gap:12px}
.login label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--suave)}
.erro{background:var(--alerta-claro);color:var(--alerta);padding:8px 12px;border-radius:6px}
.aviso-erro{background:var(--papel);border:1px solid var(--alerta);border-radius:10px;padding:20px}
tr.atrasada td{background:var(--alerta-claro)}
tr.atrasada td:first-child{box-shadow:inset 3px 0 0 var(--alerta)}
tr.selecionado td{font-weight:600}
.marca.atraso{background:var(--alerta);color:var(--papel)}
.definicao{background:var(--aviso-claro);color:var(--texto);border-left:3px solid var(--aviso);padding:10px 14px;border-radius:6px;margin:8px 0}
.alerta-txt{color:var(--alerta)}
.item.quebra{white-space:normal;min-width:14em}
input[type=number]{font:inherit;padding:6px 8px;border:1px solid var(--linha);border-radius:6px;background:var(--papel);color:var(--texto)}
input.curto{width:6em}
svg.funil{width:100%;height:auto;max-width:980px;display:block;margin:6px 0 16px;font-family:inherit}
.f-barra{fill:var(--acento)}
.f-atraso{fill:var(--alerta)}
.f-vazio{stroke:var(--linha);stroke-width:2}
.f-rotulo{fill:var(--texto);font-size:15px;font-weight:600}
.f-sub{fill:var(--suave);font-size:12px}
.legenda{display:flex;gap:8px;align-items:center;font-size:13px;color:var(--suave);margin:6px 0}
.cor{display:inline-block;width:12px;height:12px;border-radius:3px}
.cor-dia{background:var(--acento)}
.cor-atraso{background:var(--alerta);margin-left:10px}
.form-importar{display:flex;flex-direction:column;gap:12px;align-items:flex-start}
.form-importar fieldset{border:1px solid var(--linha);border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:8px;margin:0}
.form-importar legend{font-size:13px;color:var(--suave);padding:0 4px}
.form-importar label{display:flex;flex-direction:column;gap:4px;font-size:14px}
.form-importar label.check{flex-direction:row;align-items:flex-start;gap:8px}
input[type=file]{font:inherit}
.acoes{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin:14px 0 6px}
.acoes form{margin:0}
button.secundario{background:var(--papel);color:var(--texto);border:1px solid var(--linha)}
details{margin-top:12px}summary{cursor:pointer;color:var(--acento)}
code{font-size:13px;background:var(--fundo);padding:0 4px;border-radius:4px}
.aviso-orfa{background:var(--alerta-claro);color:var(--alerta);padding:10px 14px;border-radius:8px;border:1px solid var(--alerta)}
/* Celular: menu em abas roláveis, filtros empilhados, toques de 44 px, campos de 16 px (o iPhone não dá zoom), áreas seguras do notch */
@media (max-width:720px){
body{font-size:14px}
main{padding:12px max(12px,env(safe-area-inset-right)) max(24px,env(safe-area-inset-bottom)) max(12px,env(safe-area-inset-left))}
h1{font-size:19px}h2{font-size:16px}
.topo{position:sticky;top:0;z-index:10;flex-wrap:wrap;gap:2px 12px;padding:max(8px,env(safe-area-inset-top)) max(12px,env(safe-area-inset-right)) 0 max(12px,env(safe-area-inset-left))}
.usuario{font-size:13px;gap:8px;min-width:0}
.nome-usuario{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:45vw}
.nav{order:3;flex:0 0 100%;flex-wrap:nowrap;overflow-x:auto;gap:6px;padding:6px 0 8px;scrollbar-width:none;-webkit-overflow-scrolling:touch}
.nav::-webkit-scrollbar{display:none}
.nav a{flex:0 0 auto;padding:8px 14px;border:1px solid var(--linha);border-radius:999px;white-space:nowrap}
.nav a.ativo{background:var(--acento);border-color:var(--acento);color:var(--papel)}
button,.botao{min-height:44px;padding:10px 16px}
button.link{min-height:44px;padding:0 4px}
select,input[type=text],input[type=password],input[type=number]{font-size:16px;min-height:44px}
.filtros form{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));align-items:end;gap:10px}
.filtros form label:not(.check){display:flex;flex-direction:column;gap:4px;min-width:0}
.filtros form select{width:100%}
.filtros form>.check,.filtros form>button{grid-column:1/-1}
.check{flex-wrap:wrap;min-height:44px}
input[type=checkbox],input[type=radio]{width:20px;height:20px}
.cartoes{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin:12px 0}
.cartao{padding:10px 12px}
.cartao .valor{font-size:min(18px,4.2vw);white-space:nowrap}
.rankings{grid-template-columns:1fr;gap:0}
.bloco{padding:12px;margin-bottom:12px}
.cab-bloco{flex-wrap:wrap}
table{font-size:13px}
th,td{padding:8px 6px}
svg.funil{min-width:560px}
.login{margin:6vh auto;padding:22px}
}
@media (max-width:360px){.cartoes{grid-template-columns:1fr}}
'@
}
