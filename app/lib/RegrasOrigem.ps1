# Origem do faturamento: REGRAS_NEGOCIO.md §1 e §0.1.
# Uma linha por venda, ligada exclusivamente pelo ID Oportunidade.
function Get-VisaoOrigem {
    param($Base, $Pipeline, [int]$MesInicial, [int]$MesFinal)
    $erros = New-Object Collections.Generic.List[string]
    if ($Base.Erros.Count -or $Pipeline.Erros.Count) {
        return [pscustomobject]@{ Erros = @($Base.Erros) + @($Pipeline.Erros) }
    }
    $porId = @{}
    foreach ($o in $Pipeline.Oportunidades) {
        if ($porId.ContainsKey($o.Id)) { $erros.Add("ID Oportunidade duplicado no CRM: $($o.Id)"); continue }
        $porId[$o.Id] = $o
    }
    $com = [decimal]0; $sem = [decimal]0; $cancelado = [decimal]0
    $qCom = 0; $qSem = 0; $qCancelado = 0
    $linhas = New-Object Collections.Generic.List[object]
    $categorias = @{}
    foreach ($v in ($Base.Vendas | Sort-Object Data, Id)) {
        if ($v.Data.Year -ne $script:AnoMetas -or $v.Data.Month -lt $MesInicial -or $v.Data.Month -gt $MesFinal) { continue }
        $o = $null
        if ($v.IdOportunidade) {
            if (-not $porId.ContainsKey($v.IdOportunidade)) {
                $erros.Add("Venda $($v.Id), linha $($v.Linha): ID Oportunidade '$($v.IdOportunidade)' não encontrado no CRM.")
            } else { $o = $porId[$v.IdOportunidade] }
        }
        $faturada = Test-VendaRealizada $v
        $linhas.Add([pscustomobject]@{ Venda = $v; Oportunidade = $o; Faturada = $faturada })
        if (-not $faturada) { $qCancelado++; $cancelado += $v.Valor; continue }
        if ($o) { $qCom++; $com += $v.Valor }
        elseif (-not $v.IdOportunidade) {
            $qSem++; $sem += $v.Valor
            $c = [string]$v.Categoria
            if (-not $categorias.ContainsKey($c)) { $categorias[$c] = [pscustomobject]@{ Categoria = $c; Quantidade = 0; Valor = [decimal]0 } }
            $categorias[$c].Quantidade++; $categorias[$c].Valor += $v.Valor
        }
    }
    # Não publicar totais incompletos nem tratar ID inválido como venda de balcão.
    if ($erros.Count) { return [pscustomobject]@{ Erros = $erros.ToArray() } }
    $total = $com + $sem
    return [pscustomobject]@{
        Erros = @(); MesInicial = $MesInicial; MesFinal = $MesFinal
        DataBase = $Base.DataBase; DataBaseCrm = $Pipeline.DataBase
        Total = $total; ComCrm = $com; SemOportunidade = $sem
        QtdComCrm = $qCom; QtdSemOportunidade = $qSem; QtdFaturadas = $qCom + $qSem
        FracaoCrm = if ($total -gt 0) { $com / $total } else { $null }
        FracaoSem = if ($total -gt 0) { $sem / $total } else { $null }
        QtdCanceladas = $qCancelado; ValorCancelado = $cancelado
        Linhas = $linhas.ToArray()
        CategoriasSem = @($categorias.Values | Sort-Object -Property @{ Expression = { $_.Valor }; Descending = $true }, Categoria)
        Parcial = ($MesFinal -eq $Base.DataBase.Month -and $Base.DataBase.AddDays(1).Month -eq $Base.DataBase.Month)
    }
}
