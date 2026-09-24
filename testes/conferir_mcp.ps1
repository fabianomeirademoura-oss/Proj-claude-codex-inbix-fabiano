# Confere o servidor MCP (mcp/servidor.js, só leitura) falando com ele pelo protocolo, como o Claude Desktop faz.
# Compara as respostas com as regras da versão PowerShell (a referência) e com os números da REGRAS_NEGOCIO.md §6.
# Roda sobre uma cópia temporária das planilhas de dados/, sem importações (§9.9). Precisa do node.
# Uso: pwsh -NoProfile -File testes/conferir_mcp.ps1

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$raiz = Split-Path $PSScriptRoot -Parent
foreach ($mod in @('Xlsx', 'Regras', 'RegrasPipeline', 'RegrasEstoque')) { . (Join-Path $raiz "app/lib/$mod.ps1") }
if (-not $DirDados) { $DirDados = Join-Path $raiz 'dados' }
$inv = [Globalization.CultureInfo]::InvariantCulture

$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -ceq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado '$Esperado', MCP '$Obtido'") }
}
function Txt($d) { if ($null -eq $d) { return '' } ([decimal]$d).ToString('0.00', $inv) }
function Pct($n, $d) { ([math]::Round([decimal]$n / [decimal]$d * 100, 1, [MidpointRounding]::AwayFromZero)).ToString('0.0', [Globalization.CultureInfo]'pt-BR') + '%' }

# --- cópia temporária só com as planilhas da base ---
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte-mcp-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
Get-ChildItem $DirDados -Filter '*.xlsx' | Copy-Item -Destination $tmp

$psi = [Diagnostics.ProcessStartInfo]::new('node', (Join-Path $raiz 'mcp/servidor.js'))
$psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.StandardInputEncoding = [Text.UTF8Encoding]::new($false)   # sem BOM: a primeira linha tem de ser JSON puro
$psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
$psi.Environment['HORIZONTE_DADOS'] = $tmp
$proc = [Diagnostics.Process]::Start($psi)
$script:id = 0

function Enviar([string]$Metodo, $Params, [switch]$Notificacao) {
    $msg = [ordered]@{ jsonrpc = '2.0'; method = $Metodo }
    if (-not $Notificacao) { $script:id++; $msg.id = $script:id }
    if ($null -ne $Params) { $msg.params = $Params }
    $proc.StandardInput.WriteLine(($msg | ConvertTo-Json -Depth 10 -Compress))
    $proc.StandardInput.Flush()
    if ($Notificacao) { return }
    $linha = $proc.StandardOutput.ReadLine()
    if ($null -eq $linha) { throw "O servidor fechou a saída. Erro: $($proc.StandardError.ReadToEnd())" }
    $resp = $linha | ConvertFrom-Json -Depth 50
    Confere "$Metodo#$($script:id): id da resposta" $script:id $resp.id
    return $resp
}
function Chamar([string]$Nome, [hashtable]$Argumentos = @{}) {
    $r = Enviar 'tools/call' @{ name = $Nome; arguments = $Argumentos }
    return [pscustomobject]@{ Erro = [bool]$r.result.isError; Corpo = ($r.result.content[0].text | ConvertFrom-Json -Depth 50) }
}

try {
    # --- protocolo ---
    $r = Enviar 'initialize' @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 'conferir_mcp'; version = '1' } }
    Confere 'versão do protocolo' '2025-06-18' $r.result.protocolVersion
    Confere 'nome do servidor' 'horizonte-maquinas' $r.result.serverInfo.name
    Enviar 'notifications/initialized' $null -Notificacao
    $r = Enviar 'tools/list' @{}
    Confere 'ferramentas' 'consultar_vendedor,buscar_produto,ver_estoque,ver_meta,listar_oportunidades' (($r.result.tools | ForEach-Object name) -join ',')
    foreach ($t in $r.result.tools) {
        Confere "$($t.name): só leitura" 'True|False' "$($t.annotations.readOnlyHint)|$($t.annotations.destructiveHint)"
        Confere "$($t.name): parâmetros em objeto" 'object' $t.inputSchema.type
    }
    $r = Enviar 'metodo/inexistente' @{}
    Confere 'método desconhecido' '-32601' $r.error.code
    $r = Enviar 'tools/call' @{ name = 'apagar_tudo'; arguments = @{} }
    Confere 'ferramenta desconhecida' '-32602' $r.error.code

    # --- referência: versão PowerShell sobre a mesma base ---
    $base = Import-BaseComercial -DirDados $tmp
    $pipe = Import-BasePipeline -DirDados $tmp -Comercial $base
    $est = Import-BaseEstoque -DirDados $tmp
    if ($base.Erros.Count -or $pipe.Erros.Count -or $est.Erros.Count) { throw 'A base tem erros de validação.' }
    $mesDb = $base.DataBase.Month

    # --- 1. consultar_vendedor: todos os vendedores, em três períodos ---
    foreach ($periodo in @(@(1, $mesDb), @(7, 8), @(3, 5))) {
        $de, $ate = $periodo
        $des = Get-Desempenho -Base $base -MesInicial $de -MesFinal $ate
        foreach ($l in $des.Linhas) {
            $v = $l.Vendedor
            $c = (Chamar 'consultar_vendedor' @{ vendedor = $v.Id; mes_inicial = $de; mes_final = $ate }).Corpo
            $q = "consultar_vendedor $($v.Id) $de-$ate"
            Confere "$q meta" (Txt $l.Meta) (Txt $c.meta)
            Confere "$q realizado" (Txt $l.RealizadoMesesMeta) (Txt $c.realizado)
            Confere "$q faturamento" (Txt $l.Realizado) (Txt $c.faturamento_no_periodo)
            Confere "$q vendas" $l.QtdVendas $c.vendas_realizadas
            Confere "$q meses com meta" (@($l.MesesComMeta) -join ',') (@($c.meses_com_meta) -join ',')
            Confere "$q atingimento" $(if ($l.Meta -gt 0) { Pct $l.RealizadoMesesMeta $l.Meta } else { 'sem meta' }) $c.atingimento.texto
            Confere "$q ativo" $v.Ativo $c.vendedor.ativo
            Confere "$q filial" $v.IdFilial $c.vendedor.filial.id
        }
    }
    # §6: os três casos da conferência, chamados pelo nome.
    $c = (Chamar 'consultar_vendedor' @{ vendedor = 'mariana' }).Corpo
    Confere '§6 V003 pelo nome' 'V003|114,9%|6474285.51|5635000.00' "$($c.vendedor.id)|$($c.atingimento.texto)|$(Txt $c.realizado)|$(Txt $c.meta)"
    $c = (Chamar 'consultar_vendedor' @{ vendedor = 'Ricardo Alves' }).Corpo
    Confere '§6 V011 desligado' 'V011|62,8%|desligado em 30/04/2026|1,2,3,4' "$($c.vendedor.id)|$($c.atingimento.texto)|$($c.vendedor.marcacao)|$($c.meses_com_meta -join ',')"
    $c = (Chamar 'consultar_vendedor' @{ vendedor = 'CAMILA' }).Corpo
    Confere '§6 V012 admitida' 'V012|80,1%|admissão em 07/2026|629152.07|785000.00' "$($c.vendedor.id)|$($c.atingimento.texto)|$($c.vendedor.marcacao)|$(Txt $c.realizado)|$(Txt $c.meta)"
    Confere '§6 V012 jan–jun sem meta' 'sem meta' $c.por_mes[0].atingimento.texto
    # Entradas ruins viram erro com as opções, nunca um palpite.
    $c = Chamar 'consultar_vendedor' @{ vendedor = 'Fulano de Tal' }
    Confere 'vendedor inexistente' 'True|12' "$($c.Erro)|$(@($c.Corpo.vendedores).Count)"
    $c = Chamar 'consultar_vendedor' @{ vendedor = 'a' }
    Confere 'nome ambíguo' 'True' $c.Erro
    $c = Chamar 'consultar_vendedor' @{ vendedor = 'V003'; mes_final = $mesDb + 1 }
    Confere 'mês depois da data-base' 'True' $c.Erro
    $c = Chamar 'consultar_vendedor' @{}
    Confere 'sem vendedor' 'True' $c.Erro

    # --- 4. ver_meta ---
    foreach ($mes in 1..$mesDb) {
        $des = Get-Desempenho -Base $base -MesInicial $mes -MesFinal $mes
        $c = (Chamar 'ver_meta' @{ mes = $mes }).Corpo
        Confere "ver_meta $mes empresa meta" (Txt $des.Empresa.Meta) (Txt $c.empresa.meta)
        Confere "ver_meta $mes empresa realizado" (Txt $des.Empresa.RealizadoMesesMeta) (Txt $c.empresa.realizado)
        Confere "ver_meta $mes empresa atingimento" (Pct $des.Empresa.RealizadoMesesMeta $des.Empresa.Meta) $c.empresa.atingimento.texto
        Confere "ver_meta $mes vendedores" 12 @($c.vendedores).Count
        foreach ($l in $des.Linhas) {
            $x = @($c.vendedores | Where-Object { $_.vendedor.id -eq $l.Vendedor.Id })[0]
            Confere "ver_meta $mes $($l.Vendedor.Id) meta" $(if ($l.MesesComMeta.Count) { Txt $l.Meta } else { '' }) (Txt $x.meta)
            Confere "ver_meta $mes $($l.Vendedor.Id) realizado" (Txt $l.RealizadoMesesMeta) (Txt $x.realizado)
        }
    }
    $c = (Chamar 'ver_meta' @{ vendedor = 'V012' }).Corpo
    Confere 'ver_meta V012: 12 meses' 12 @($c.meses).Count
    Confere 'ver_meta V012: sem meta antes da admissão' '|sem meta' "$(Txt $c.meses[5].meta)|$($c.meses[5].atingimento.texto)"
    Confere 'ver_meta V012: mês futuro sem atingimento' '' "$($c.meses[11].atingimento.fracao)"
    $c = (Chamar 'ver_meta' @{ vendedor = 'V003'; mes = 8 }).Corpo
    Confere 'ver_meta vendedor e mês' '1|8|mês fechado' "$(@($c.meses).Count)|$($c.meses[0].mes)|$($c.meses[0].situacao)"
    $c = Chamar 'ver_meta' @{ ano = 2025 }
    Confere 'ver_meta outro ano' 'True' $c.Erro

    # --- 5. listar_oportunidades ---
    $c = (Chamar 'listar_oportunidades' @{ limite = 300 }).Corpo
    Confere '§6 abertas' '68|13870500.00|5605495.00' "$($c.resumo.quantidade)|$(Txt $c.resumo.valor)|$(Txt $c.resumo.ponderado_das_abertas)"
    Confere '§6 atrasadas' '29|6813700.00' "$($c.resumo.vencidas.quantidade)|$(Txt $c.resumo.vencidas.valor)"
    Confere 'vencidas primeiro, maior atraso antes' 'True' "$(($c.oportunidades[0].dias_de_atraso -ge $c.oportunidades[1].dias_de_atraso) -and $c.oportunidades[0].vencida)"
    $c = (Chamar 'listar_oportunidades' @{ vencidas = $true; limite = 1 }).Corpo
    Confere 'filtro vencidas' '29|1' "$($c.resumo.quantidade)|$(@($c.oportunidades).Count)"
    $c = (Chamar 'listar_oportunidades' @{ vencidas = $false }).Corpo
    Confere 'filtro não vencidas' 39 $c.resumo.quantidade
    $c = (Chamar 'listar_oportunidades' @{ vendedor = 'V011' }).Corpo
    Confere '§6 órfãs' '2|255300.00|True' "$($c.resumo.quantidade)|$(Txt $c.resumo.valor)|$(@($c.oportunidades | Where-Object orfa).Count -eq 2)"
    $c = (Chamar 'listar_oportunidades' @{ etapa = 'todas'; limite = 1 }).Corpo
    Confere 'todas as etapas' $pipe.Oportunidades.Count $c.resumo.quantidade
    $c = (Chamar 'listar_oportunidades' @{ etapa = 'Fechada Ganha' }).Corpo
    Confere '§6 ganhas' 95 $c.resumo.quantidade
    $c = (Chamar 'listar_oportunidades' @{ etapa = 'fechada perdida' }).Corpo
    Confere '§6 perdidas' 96 $c.resumo.quantidade
    foreach ($etapa in @(@('Prospecção', 16), @('qualificacao', 13), @('Proposta Enviada', 22), @('NEGOCIAÇÃO', 17))) {
        $c = (Chamar 'listar_oportunidades' @{ etapa = $etapa[0] }).Corpo
        Confere "§8.5 etapa $($etapa[0])" $etapa[1] $c.resumo.quantidade
    }
    $c = Chamar 'listar_oportunidades' @{ etapa = 'ganhou' }
    Confere 'etapa desconhecida' 'True' $c.Erro
    $vis = Get-VisaoPipeline -Comercial $base -Desempenho (Get-Desempenho -Base $base -MesInicial 1 -MesFinal $mesDb) -Pipeline $pipe
    foreach ($l in $vis.Linhas) {
        $c = (Chamar 'listar_oportunidades' @{ vendedor = $l.Vendedor.Id; limite = 1 }).Corpo
        $q = "pipeline $($l.Vendedor.Id)"
        Confere "$q abertas" $l.QtdAbertas $c.resumo.quantidade
        Confere "$q valor" (Txt $l.Pipeline) (Txt $c.resumo.valor)
        Confere "$q ponderado" (Txt $l.Ponderado) (Txt $c.resumo.ponderado_das_abertas)
        Confere "$q atrasadas" "$($l.QtdAtrasadas)|$(Txt $l.ValorAtrasado)" "$($c.resumo.vencidas.quantidade)|$(Txt $c.resumo.vencidas.valor)"
    }

    # --- 3. ver_estoque ---
    $c = (Chamar 'ver_estoque' @{}).Corpo
    Confere 'estoque: data da foto' $est.DataBase.ToString('yyyy-MM-dd') $c.foto.data
    Confere 'estoque: linhas' $est.Linhas.Count $c.resumo.linhas
    Confere '§6 valor imobilizado' '18420192.72' (Txt $c.resumo.valor_imobilizado)
    Confere '§6 parados' '7|1693610.48' "$($c.resumo.parados.linhas)|$(Txt $c.resumo.parados.valor)"
    Confere '§6 abaixo do mínimo' 28 $c.resumo.abaixo_do_minimo
    foreach ($l in $est.Linhas) {
        $x = @($c.linhas | Where-Object { $_.produto.id -eq $l.Produto.Id -and $_.filial.id -eq $l.IdFilial })
        $q = "estoque $($l.Produto.Id)/$($l.IdFilial)"
        Confere "$q única" 1 $x.Count
        Confere "$q quantidade e valor" "$($l.Quantidade)|$($l.Minimo)|$(Txt $l.Valor)" "$($x[0].quantidade)|$($x[0].estoque_minimo)|$(Txt $x[0].valor_imobilizado)"
        Confere "$q parado" (Test-LinhaParada $l 180) $x[0].parado
        Confere "$q dias" "$($l.DiasSemMovimento)" "$($x[0].dias_parado)"
    }
    foreach ($dias in @(90, 365)) {
        $esperado = @($est.Linhas | Where-Object { Test-LinhaParada $_ $dias }).Count
        $c = (Chamar 'ver_estoque' @{ somente_parados = $true; dias_parado = $dias }).Corpo
        Confere "parados com prazo $dias" $esperado $c.resumo.linhas
    }
    $c = (Chamar 'ver_estoque' @{ produto = 'P008'; filial = 'cascavel' }).Corpo
    Confere '§5.4 P008 Cascavel parado com 2 oportunidades' '1|True|2' "$($c.resumo.linhas)|$($c.linhas[0].parado)|$($c.linhas[0].oportunidades_abertas_do_produto.quantidade)"
    $c = (Chamar 'ver_estoque' @{ filial = 'Chapeco' }).Corpo
    Confere 'filial pelo nome sem acento' "F02|$(@($est.Linhas | Where-Object IdFilial -eq 'F02').Count)" "$($c.filtro.filial)|$($c.resumo.linhas)"
    $c = Chamar 'ver_estoque' @{ filial = 'Curitiba' }
    Confere 'filial desconhecida' 'True|3' "$($c.Erro)|$(@($c.Corpo.filiais).Count)"
    $c = Chamar 'ver_estoque' @{ dias_parado = 0 }
    Confere 'prazo inválido' 'True' $c.Erro

    # --- 2. buscar_produto ---
    $c = (Chamar 'buscar_produto' @{ busca = 'p008' }).Corpo
    Confere 'produto pelo código' 'P008|2|830000.00' "$($c.produtos[0].id)|$($c.produtos[0].estoque_unidades)|$(Txt $c.produtos[0].preco_tabela)"
    $cat = @($est.Produtos | Where-Object Categoria -eq 'Tratores')
    $c = (Chamar 'buscar_produto' @{ categoria = 'tratores'; limite = 300 }).Corpo
    Confere 'categoria inteira' "$($cat.Count)|Tratores" "$($c.encontrados)|$($c.categoria)"
    $c = (Chamar 'buscar_produto' @{ busca = 'P002' }).Corpo
    Confere '§5.5 descontinuado aparece marcado' 'True' $c.produtos[0].descontinuado
    $sem = @($est.Produtos | Where-Object { $p = $_.Id; $s = 0; foreach ($l in $est.Linhas) { if ($l.Produto.Id -eq $p) { $s += $l.Quantidade } }; $s -eq 0 })
    $todos = foreach ($cat in ($est.Produtos | ForEach-Object Categoria | Sort-Object -Unique)) { (Chamar 'buscar_produto' @{ categoria = $cat; limite = 300 }).Corpo.produtos }
    Confere 'catálogo inteiro pelas categorias' $est.Produtos.Count @($todos).Count
    Confere '§6 sem estoque' 8 @($todos | Where-Object { $_.estoque_unidades -eq 0 }).Count
    Confere 'sem estoque confere com a foto' (($sem | ForEach-Object Id | Sort-Object) -join ',') ((@($todos | Where-Object { $_.estoque_unidades -eq 0 }) | ForEach-Object id | Sort-Object) -join ',')
    $c = (Chamar 'buscar_produto' @{ busca = 'plantadeira'; limite = 2 }).Corpo
    Confere 'limite respeitado' 2 @($c.produtos).Count
    $c = Chamar 'buscar_produto' @{}
    Confere 'busca vazia' 'True' $c.Erro

    # --- só leitura: nada mudou na pasta de dados ---
    Confere 'nenhum arquivo criado' ((Get-ChildItem $DirDados -Filter '*.xlsx').Count) (Get-ChildItem $tmp).Count
} finally {
    $proc.StandardInput.Close()
    if (-not $proc.WaitForExit(5000)) { $proc.Kill() }
    Remove-Item $tmp -Recurse -Force
}

if ($script:falhas.Count) {
    $script:falhas | ForEach-Object { Write-Host "  FALHA $_" -ForegroundColor Red }
    Write-Host "$($script:falhas.Count) falhas em $($script:ok + $script:falhas.Count) conferências do MCP." -ForegroundColor Red
    exit 1
}
Write-Host "Todas as $($script:ok) conferências do MCP bateram." -ForegroundColor Green
