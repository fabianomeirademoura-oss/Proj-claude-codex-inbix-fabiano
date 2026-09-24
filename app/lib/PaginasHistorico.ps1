# Tela do histórico de alterações (REGRAS_NEGOCIO.md §11.3). Só formata: as linhas vêm de Read-Alteracoes (Regras.ps1).
# Usa os auxiliares de Paginas.ps1 (Esc, New-Layout). Sem JavaScript nem estilo novo (CSP).

$script:TiposAlteracao = [ordered]@{ vendedor = 'Vendedor'; meta = 'Meta'; produto = 'Produto'; oportunidade = 'Oportunidade' }

function Test-PodeVerHistorico($Usuario) { return $Usuario.perfil -in @('diretoria', 'gerente') }   # §11.3

function Get-HistoricoFiltrado($Alteracoes, [string]$Tipo, [string]$Registro) {
    # Da alteração mais recente para a mais antiga (o arquivo só recebe acréscimos, em ordem).
    $lista = @($Alteracoes | Where-Object { $_ })
    [array]::Reverse($lista)
    return @($lista | Where-Object {
            (-not $Tipo -or $_.Tipo -ceq $Tipo) -and (-not $Registro -or $_.Registro.IndexOf($Registro, [StringComparison]::OrdinalIgnoreCase) -ge 0)
        })
}

function Format-ValorHistorico([string]$Valor) { if ($Valor -eq '') { return '—' } return (Esc $Valor) }

function New-PaginaHistorico($Alteracoes, [string]$Tipo, [string]$Registro, $Usuario) {
    $Alteracoes = @($Alteracoes | Where-Object { $_ })
    $linhas = @(Get-HistoricoFiltrado $Alteracoes $Tipo $Registro)
    $opcoes = '<option value="">todos</option>' + (($script:TiposAlteracao.Keys | ForEach-Object {
                $sel = if ($_ -ceq $Tipo) { ' selected' } else { '' }
                "<option value=""$_""$sel>$($script:TiposAlteracao[$_])</option>"
            }) -join '')
    $q = "tipo=$([Uri]::EscapeDataString($Tipo))&amp;registro=$([Uri]::EscapeDataString($Registro))"
    $corpoTabela = foreach ($a in $linhas) {
        $tipo = if ($a.Tipo -cin @($script:TiposAlteracao.Keys)) { $script:TiposAlteracao[$a.Tipo] } else { $a.Tipo }
        "<tr><td class=""nw"">$(Esc $a.Quando)</td><td>$(Esc $a.Nome)<span class=""item id"">$(Esc $a.Login) · $(Esc $a.Perfil)</span></td><td>$(Esc $a.'Ação')</td><td class=""nw"">$(Esc $tipo) <code>$(Esc $a.Registro)</code></td><td>$(Esc $a.Campo)</td><td>$(Format-ValorHistorico $a.Anterior)</td><td>$(Format-ValorHistorico $a.Novo)</td><td class=""nw""><code>$(Esc $a.Lote)</code></td></tr>"
    }
    $tabela = if ($linhas.Count) {
        "<table><thead><tr><th>Quando</th><th>Quem</th><th>Ação</th><th>Registro</th><th>Campo</th><th>Anterior</th><th>Novo</th><th>Lote</th></tr></thead><tbody>$($corpoTabela -join "`n")</tbody></table>"
    } elseif (@($Alteracoes).Count) { '<p class="nota">Nenhuma alteração com esse filtro.</p>' } else { '<p class="nota">Nenhuma alteração registrada até agora.</p>' }
    $corpo = @"
<section class="filtros">
  <h1>Histórico de alterações</h1>
  <p class="fonte">Quem alterou o quê pelas ferramentas de escrita (REGRAS_NEGOCIO.md §11): uma linha por campo, com o valor anterior e o novo. Nada é apagado; desfazer é uma nova alteração, que também aparece aqui. As alterações do mesmo lote foram confirmadas juntas.</p>
  <form method="get" action="/historico">
    <label>Tipo <select name="tipo">$opcoes</select></label>
    <label>Registro <input name="registro" type="text" value="$(Esc $Registro)" placeholder="ex.: V011, P066, OP-0130"></label>
    <button type="submit">Filtrar</button>
    <a href="/historico.csv?$q">Baixar CSV</a>
  </form>
</section>
<section class="bloco">
  <p class="nota">$($linhas.Count) de $(@($Alteracoes).Count) alterações registradas, da mais recente para a mais antiga.</p>
  $tabela
</section>
<footer class="regras">REGRAS_NEGOCIO.md §11.1: nada é gravado sem confirmação · §11.2: só gerentes (na própria filial) e a diretoria alteram · §11.3: registro por campo, só com acréscimos · §11.4: as planilhas originais não mudam; o painel aplica o registro por cima delas.</footer>
"@
    return New-Layout 'Histórico · Horizonte Máquinas' $corpo $Usuario 'historico'
}

function New-CsvHistorico($Alteracoes, [string]$Tipo, [string]$Registro) {
    $aspas = { param($t) '"' + ([string]$t).Replace('"', '""') + '"' }
    $saida = New-Object Text.StringBuilder
    [void]$saida.Append((($script:ColunasAlteracoes | ForEach-Object { & $aspas $_ }) -join ';') + "`r`n")
    foreach ($a in (Get-HistoricoFiltrado $Alteracoes $Tipo $Registro)) {
        [void]$saida.Append((($script:ColunasAlteracoes | ForEach-Object { & $aspas $a.$_ }) -join ';') + "`r`n")
    }
    return $saida.ToString()
}

function New-PaginaSemAcessoHistorico($Usuario) {
    return New-Layout 'Sem acesso' '<section class="aviso-erro"><h1>Sem acesso</h1><p>Só a diretoria e os gerentes veem o histórico de alterações (REGRAS_NEGOCIO.md §11.3).</p></section>' $Usuario 'historico'
}
