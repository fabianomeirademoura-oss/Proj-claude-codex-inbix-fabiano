# Conferência independente e casos de borda da origem do faturamento.
$ErrorActionPreference = 'Stop'
foreach ($mod in @('Xlsx','Regras','RegrasPipeline','RegrasOrigem','Paginas','PaginasOrigem')) { . (Join-Path $PSScriptRoot "../app/lib/$mod.ps1") }
$dir = Join-Path $PSScriptRoot '../dados'
$base = Import-BaseComercial $dir
$pipe = Import-BasePipeline $dir $base
$total = 0
function Confere($Nome, $Ok) { $script:total++; if (-not $Ok) { throw "FALHA: $Nome" } }
$raw = Read-XlsxSheet (Join-Path $dir 'vendas_2026_jan-ago.xlsx') 'Vendas'
$ops = Read-XlsxSheet (Join-Path $dir 'crm_oportunidades.xlsx') 'Oportunidades'
foreach ($periodo in @(@(1,8),@(7,8),@(3,5),@(1,1))) {
    $de,$ate = $periodo
    $v = Get-VisaoOrigem $base $pipe $de $ate
    Confere 'sem erros' ($v.Erros.Count -eq 0)
    $com = [decimal]0; $sem = [decimal]0; $nCom = 0; $nSem = 0
    foreach ($r in $raw) {
        $data = [datetime]$r.Data
        if ($r.Status -ne 'Faturada' -or $data.Month -lt $de -or $data.Month -gt $ate) { continue }
        $valor = [math]::Round([decimal]::Parse($r.'Valor Total', [Globalization.CultureInfo]::InvariantCulture),2,[MidpointRounding]::AwayFromZero)
        if ($r.'ID Oportunidade') {
            $match = @($ops | Where-Object { $_.'ID Oportunidade' -eq $r.'ID Oportunidade' })
            Confere 'vínculo único na fonte' ($match.Count -eq 1)
            $link = @($v.Linhas | Where-Object { $_.Venda.Id -eq $r.'ID Venda' })
            Confere 'mesma oportunidade na visão' ($link.Count -eq 1 -and $link[0].Oportunidade.Id -eq $match[0].'ID Oportunidade')
            $com += $valor; $nCom++
        } else { $sem += $valor; $nSem++ }
    }
    Confere 'valor CRM' ($v.ComCrm -eq $com)
    Confere 'valor sem oportunidade' ($v.SemOportunidade -eq $sem)
    Confere 'quantidades' ($v.QtdComCrm -eq $nCom -and $v.QtdSemOportunidade -eq $nSem)
    Confere 'participação por faturamento' ($v.FracaoCrm -eq $com/($com+$sem))
    Confere 'total reconcilia com painel' ($v.Total -eq (Get-Desempenho $base $de $ate).Empresa.Realizado)
    $soma = [decimal]0; foreach ($c in $v.CategoriasSem) { $soma += $c.Valor }
    Confere 'categorias sem oportunidade reconciliam' ($soma -eq $sem)
}
# Fixtures em memória: nenhuma alteração nas planilhas.
function Venda($Id,$Op,$Valor,$Status,$Data='2026-01-15') {
    [pscustomobject]@{ Id=$Id; IdOportunidade=$Op; Valor=[decimal]$Valor; Status=$Status; Data=[datetime]$Data; Categoria='Máquinas'; Linha=2; Cliente='<script>cliente</script>'; Produto='Produto' }
}
$b = [pscustomobject]@{ Erros=@(); DataBase=[datetime]'2026-08-31'; Vendas=@((Venda 'A' 'OP-TESTE' 100 'Faturada'),(Venda 'B' '' 300 'Faturada'),(Venda 'C' 'OP-TESTE' 900 'Cancelada'),(Venda 'D' '' 500 'Faturada' '2026-09-01')) }
$o = [pscustomobject]@{Id='OP-TESTE'; Valor=[decimal]9999; Etapa='Fechada Ganha'; Origem='Teste'; Linha=2}
$c = [pscustomobject]@{Erros=@(); DataBase=[datetime]'2026-08-31'; Oportunidades=@($o)}
$v=Get-VisaoOrigem $b $c 1 8
Confere 'usa valor vendido e exclui canceladas e setembro' ($v.Total -eq 400 -and $v.FracaoCrm -eq [decimal]0.25 -and $v.QtdCanceladas -eq 1 -and $v.Linhas.Count -eq 3)
Confere 'cancelada mantém vínculo sem entrar no faturamento' ($v.Linhas[2].Oportunidade.Id -eq 'OP-TESTE' -and -not $v.Linhas[2].Faturada)
$html=New-PaginaOrigem $b $v ([pscustomobject]@{nome='Teste'})
Confere 'escapa conteúdo das planilhas' ($html -notmatch '<script>cliente' -and $html -match '&lt;script&gt;cliente')
Confere 'sem script ou estilo inline' ($html -notmatch '<script|<style| style=')
$v=Get-VisaoOrigem $b $c 2 2
Confere 'período vazio não divide por zero' ($v.Total -eq 0 -and $null -eq $v.FracaoCrm -and $v.Linhas.Count -eq 0)
Confere 'mensagem sem faturamento' ((New-PaginaOrigem $b $v ([pscustomobject]@{nome='Teste'})) -match 'Sem faturamento no período')
$b.Vendas=@((Venda 'Z' '' 0 'Faturada'))
$v=Get-VisaoOrigem $b $c 1 8
Confere 'venda zerada não vira percentual zero' ($null -eq $v.FracaoCrm -and $v.QtdSemOportunidade -eq 1)
$b.Vendas=@((Venda 'Z' 'INEXISTENTE' 100 'Faturada'))
$v=Get-VisaoOrigem $b $c 1 8
Confere 'ID inválido não vira balcão nem total incompleto' ($v.Erros.Count -eq 1 -and $null -eq $v.Total)
$c.Oportunidades=@($o,$o)
Confere 'duplicata não multiplica faturamento' ((Get-VisaoOrigem $b $c 1 8).Erros.Count -gt 0)
Write-Host "Todas as $total conferências da origem bateram."
