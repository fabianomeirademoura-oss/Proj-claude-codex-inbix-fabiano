# Números do feedback semanal ao vendedor (REGRAS_NEGOCIO.md §11), em JSON.
# Só calcula e formata: quem redige a mensagem de WhatsApp é a skill feedback-semanal.
# Uso: pwsh -NoProfile -File app/feedback_semanal.ps1 [-Vendedor V003|Mariana] [-DirDados <pasta>]
[CmdletBinding()]
param(
    [string]$Vendedor,   # ID (V003) ou parte do nome; vazio = todos os ativos
    [string]$DirDados
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
foreach ($mod in @('Xlsx', 'Regras', 'RegrasPipeline', 'RegrasImportacao', 'RegrasFeedback', 'Paginas', 'PaginasPipeline')) {
    . (Join-Path $PSScriptRoot "lib/$mod.ps1")
}
if (-not $DirDados) { $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' } }
$DirDados = (Resolve-Path -LiteralPath $DirDados).Path
$dirImp = Get-DirImportacoes $DirDados   # §9.8: vendas em vigor = base + importações

$base = Import-BaseComercial -DirDados $DirDados -DirImportacoes $dirImp
if ($base.Erros.Count) { throw ("Erros na carga de vendas:`n" + ($base.Erros -join "`n")) }
$crm = Import-BasePipeline -DirDados $DirDados -Comercial $base
if ($crm.Erros.Count) { throw ("Erros na carga do CRM:`n" + ($crm.Erros -join "`n")) }

$id = $null
if ($Vendedor) {
    $achados = @($base.Vendedores | Where-Object { $_.Id -eq $Vendedor -or $_.Nome -like "*$Vendedor*" })
    if ($achados.Count -ne 1) {
        $nomes = ($base.Vendedores | ForEach-Object { "$($_.Id) $($_.Nome)" }) -join '; '
        throw "'$Vendedor' deve identificar exatamente um vendedor ($($achados.Count) encontrados). Cadastro: $nomes"
    }
    $id = $achados[0].Id
}
$fb = Get-FeedbackSemanal -Comercial $base -Pipeline $crm -IdVendedor $id

$gerente = @{}
foreach ($f in (Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Filiais')) { $gerente[$f.'ID Filial'] = $f.'Gerente Comercial' }

function M($v) { if ($null -eq $v) { $null } else { Format-Moeda $v } }
function P($v) { if ($null -eq $v) { $null } else { Format-Pct $v } }
function D($d) { if ($d) { $d.ToString('dd/MM/yyyy') } }

$vendedores = foreach ($x in $fb.Vendedores) {
    $v = $x.Vendedor
    [ordered]@{
        id              = $v.Id
        nome            = $v.Nome
        primeiroNome    = ($v.Nome -split ' ')[0]
        filial          = $v.Filial
        gerenteFilial   = $gerente[$v.IdFilial]
        admissao        = D $v.Admissao
        admitidoNoAno   = $x.Admitida
        semana          = [ordered]@{
            periodo           = "$(D $x.Semana.Inicio) a $(D $x.Semana.Fim)"
            realizado         = M $x.Semana.Realizado
            qtdVendas         = $x.Semana.QtdVendas
            semanaAnterior    = M $x.Semana.RealizadoAnterior
            vendas            = @($x.Semana.Vendas | ForEach-Object { [ordered]@{ data = D $_.Data; cliente = $_.Cliente; produto = $_.Produto; valor = M $_.Valor } })
        }
        mesAtual        = [ordered]@{
            mes            = $script:NomesMesLongo[$x.MesAtual.Mes]
            parcialAte     = if ($x.MesAtual.Parcial) { D $fb.DataBaseVendas }
            diasDecorridos = "$($x.MesAtual.DiasDecorridos) de $($x.MesAtual.DiasNoMes) dias ($(P $x.MesAtual.FracaoDoMes) do mês)"
            meta           = M $x.MesAtual.Meta
            realizado      = M $x.MesAtual.Realizado
            atingimento    = P $x.MesAtual.Atingimento
            falta          = M $x.MesAtual.Falta
            noRitmo        = $x.MesAtual.NoRitmo
        }
        mesesFechados   = [ordered]@{
            meses       = Format-Meses $x.Fechados.Meses
            meta        = M $x.Fechados.Meta
            realizado   = M $x.Fechados.Realizado
            atingimento = P $x.Fechados.Atingimento
            mediaMensal = M $x.Fechados.MediaMensal
        }
        ano             = [ordered]@{
            mesesComMeta       = Format-Meses $x.Ano.MesesComMeta
            meta               = M $x.Ano.Meta
            realizado          = M $x.Ano.Realizado
            atingimento        = P $x.Ano.Atingimento
            falta              = M $x.Ano.Falta
            projecaoPelaMedia  = M $x.Ano.Projecao
            projecaoPct        = P $x.Ano.ProjecaoPct
            projecaoFaltaria   = M $x.Ano.ProjecaoFaltaria
            projecaoSobra      = M $x.Ano.ProjecaoSobra
            mesesRestantes     = Format-Meses $x.Ano.MesesRestantes
            metaMediaRestante  = M $x.Ano.MetaMediaRestante
            necessarioPorMes   = M $x.Ano.NecessarioPorMes
            sinal              = $x.Ano.Sinal
        }
        pipeline        = [ordered]@{
            fotoCrm          = D $x.Pipeline.DataBaseCrm
            abertas          = $x.Pipeline.QtdAbertas
            valor            = M $x.Pipeline.Valor
            ponderado        = M $x.Pipeline.Ponderado
            atrasadas        = $x.Pipeline.QtdAtrasadas
            valorAtrasado    = M $x.Pipeline.ValorAtrasado
            coberturaDaFalta = P $x.Pipeline.CoberturaFalta
            cenarioPonderado = M $x.Pipeline.CenarioPonderado
            cenarioPct       = P $x.Pipeline.CenarioPct
            ganhas           = $x.Pipeline.Ganhas
            perdidas         = $x.Pipeline.Perdidas
            conversao        = P $x.Pipeline.Conversao
            principalMotivoPerda = $x.Pipeline.PrincipalMotivo
            oportunidades    = @($x.Pipeline.Oportunidades | ForEach-Object {
                    [ordered]@{
                        id = $_.Id; cliente = $_.Cliente; produto = $_.Produto; etapa = $_.Etapa
                        probabilidade = Format-Prob $_.Probabilidade; valor = M $_.Valor; ponderado = M $_.Ponderado
                        previsao = D $_.Previsao; atrasada = $_.Atrasada; diasAtraso = $_.DiasAtraso
                    } })
        }
    }
}

[ordered]@{
    geradoEm        = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    dataBaseVendas  = D $fb.DataBaseVendas
    dataBaseCrm     = D $fb.DataBaseCrm
    crmDefasado     = $fb.CrmDefasado
    mesParcial      = $fb.MesParcial
    importacoes     = @(Get-ArquivosVendasImportados $dirImp | ForEach-Object Name)
    vendedores      = @($vendedores)
} | ConvertTo-Json -Depth 8
