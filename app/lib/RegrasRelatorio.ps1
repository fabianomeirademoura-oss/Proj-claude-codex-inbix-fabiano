# Projeção do relatório: reaproveita REGRAS_NEGOCIO.md §§1–5, 7–10.
function Get-DadosRelatorio($Base, $Pipeline, $Estoque) {
    foreach ($fonte in @($Base,$Pipeline,$Estoque)) {
        if ($fonte.Erros.Count) { throw ($fonte.Erros -join "`n") }
    }
    if (-not $Base.DataBase) { throw 'Sem data-base de vendas: não é possível definir a semana.' }
    $fim = $Base.DataBase.Date
    if ($fim.Year -ne $script:AnoMetas) { throw 'Ano da base incompatível com as metas.' }
    $inicio = $fim.AddDays(-6)
    $des = Get-Desempenho $Base $fim.Month $fim.Month
    $ranking = @(Get-RankingFaturamento $des)
    $abaixo = @($ranking | Where-Object { $null -ne $_.Atingimento -and $_.Atingimento -lt [decimal]0.8 } |
        Sort-Object -Property @{Expression={ $_.Meta - $_.RealizadoMesesMeta };Descending=$true}, @{Expression={$_.Vendedor.Id}})
    $funil = Get-VisaoFunil -Comercial $Base -Pipeline $Pipeline
    $atrasadas = @($funil.Abertas | Where-Object Atrasada | Sort-Object -Property @{Expression={$_.Oportunidade.Previsao}}, @{Expression={$_.Oportunidade.Id}})
    $visEst = Get-VisaoEstoque $Estoque 180 (Get-AbertasPorProduto $Pipeline)
    $filiais = @($Estoque.Filiais | Sort-Object Id)
    $porFilial = @{}
    foreach ($f in $filiais) { $porFilial[$f.Id] = @{} }
    foreach ($v in $Base.Vendas) {
        if (-not (Test-VendaRealizada $v) -or $v.Data -lt $inicio -or $v.Data -gt $fim) { continue }
        $prod = $Estoque.ProdutoPorId[$v.IdProduto]
        $q = ConvertTo-Inteiro ([string]$v.Quantidade)
        if (-not $prod -or $null -eq $q -or $q -lt 0 -or -not $porFilial.ContainsKey($v.IdFilial)) {
            throw "Venda $($v.Id): produto, filial ou quantidade inválidos para o ranking de saída."
        }
        $mapa = $porFilial[$v.IdFilial]
        if (-not $mapa.ContainsKey($v.IdProduto)) { $mapa[$v.IdProduto]=[pscustomobject]@{Id=$prod.Id;Nome=$prod.Nome;Unidade=$prod.Unidade;Quantidade=[long]0;Valor=[decimal]0} }
        $mapa[$v.IdProduto].Quantidade += $q; $mapa[$v.IdProduto].Valor += $v.Valor
    }
    $saidas = foreach ($f in $filiais) {
        [pscustomobject]@{ Filial=$f; Produtos=@($porFilial[$f.Id].Values | Where-Object { $_.Quantidade -gt 0 } |
            Sort-Object -Property @{Expression={$_.Quantidade};Descending=$true}, Id | Select-Object -First 10) }
    }
    $gap = [decimal]0
    foreach ($l in $abaixo) { $gap += $l.Meta - $l.RealizadoMesesMeta }
    $orfas = @($funil.Abertas | Where-Object Orfa)
    $valorOrfas = [decimal]0; foreach ($a in $orfas) { $valorOrfas += $a.Oportunidade.Valor }
    return [pscustomobject]@{
        InicioSemana=$inicio; FimSemana=$fim; InicioMes=[datetime]::new($fim.Year,$fim.Month,1)
        DataCrm=$Pipeline.DataBase; DataEstoque=$Estoque.DataBase; Parcial=$des.Parcial
        Ranking=$ranking; Abaixo=$abaixo; GapAbaixo=$gap; Empresa=$des.Empresa
        Funil=$funil; Atrasadas=$atrasadas; Estoque=$visEst; Filiais=$filiais; Saidas=@($saidas)
        Orfas=$orfas.Count; ValorOrfas=$valorOrfas
    }
}
