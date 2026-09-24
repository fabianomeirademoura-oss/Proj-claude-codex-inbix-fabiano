# Confere as ferramentas de escrita do servidor MCP (REGRAS_NEGOCIO.md §11) falando com ele pelo protocolo:
# confirmação (nada gravado antes do "sim", código de uso único, confirmação que refaz as validações),
# permissão (sem conta, vendedor, gerente fora da filial, diretoria) e registro (uma linha por campo).
# Depois confere que a versão PowerShell (a referência) aplica o registro e que a base pura continua a da §6.
# Roda sobre uma cópia temporária das planilhas de dados/, com contas temporárias. Precisa do node.
# Uso: pwsh -NoProfile -File testes/conferir_escrita_mcp.ps1

param([string]$DirDados)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$raiz = Split-Path $PSScriptRoot -Parent
foreach ($mod in @('Xlsx', 'Regras', 'RegrasPipeline', 'RegrasEstoque', 'Paginas', 'PaginasHistorico')) { . (Join-Path $raiz "app/lib/$mod.ps1") }
if (-not $DirDados) { $DirDados = Join-Path $raiz 'dados' }

$script:ok = 0; $script:falhas = New-Object Collections.Generic.List[string]
function Confere([string]$Oque, $Esperado, $Obtido) {
    if ("$Esperado" -ceq "$Obtido") { $script:ok++ } else { $script:falhas.Add("$Oque`: esperado '$Esperado', obtido '$Obtido'") }
}

# --- cópia temporária e contas ---
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte-escrita-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
Get-ChildItem $DirDados -Filter '*.xlsx' | Copy-Item -Destination $tmp
$imp = Join-Path $tmp 'importacoes'
$arqAlt = Join-Path $imp 'alteracoes.csv'
$arqContas = Join-Path $tmp 'usuarios.json'
$contas = @(
    [ordered]@{ login = 'diretoria'; nome = 'Diretoria comercial'; perfil = 'diretoria'; idVendedor = $null; idFilial = $null }
    [ordered]@{ login = 'gerente.cascavel'; nome = 'Gerente Cascavel'; perfil = 'gerente'; idVendedor = $null; idFilial = 'F01' }
    [ordered]@{ login = 'gerente.chapeco'; nome = 'Gerente Chapecó'; perfil = 'gerente'; idVendedor = $null; idFilial = 'F02' }
    [ordered]@{ login = 'joao.almeida@horizontemaquinas.com.br'; nome = 'João Pedro Almeida'; perfil = 'vendedor'; idVendedor = 'V001'; idFilial = $null }
)
[IO.File]::WriteAllText($arqContas, (ConvertTo-Json -InputObject $contas -Depth 3), [Text.UTF8Encoding]::new($false))

$script:processos = New-Object Collections.Generic.List[object]
function New-Mcp([string]$Login, [switch]$ComElicitation) {
    $psi = [Diagnostics.ProcessStartInfo]::new('node', (Join-Path $raiz 'mcp/servidor.js'))
    $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.StandardInputEncoding = [Text.UTF8Encoding]::new($false)
    $psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $psi.Environment['HORIZONTE_DADOS'] = $tmp
    $psi.Environment['HORIZONTE_ARQUIVO_USUARIOS'] = $arqContas
    $psi.Environment['HORIZONTE_USUARIO'] = $Login
    $m = [pscustomobject]@{ Proc = [Diagnostics.Process]::Start($psi); Id = 0; Login = $Login }
    $script:processos.Add($m)
    $cap = if ($ComElicitation) { @{ elicitation = @{} } } else { @{} }
    [void](Enviar $m 'initialize' @{ protocolVersion = '2025-06-18'; capabilities = $cap; clientInfo = @{ name = 'conferir_escrita'; version = '1' } })
    Enviar $m 'notifications/initialized' $null -Notificacao
    return $m
}
# $Responder: o que dizer se o servidor pedir confirmação à pessoa (elicitation); $null = não esperar pedido.
function Enviar($M, [string]$Metodo, $Params, [switch]$Notificacao, [scriptblock]$Responder) {
    $msg = [ordered]@{ jsonrpc = '2.0'; method = $Metodo }
    if (-not $Notificacao) { $M.Id++; $msg.id = $M.Id }
    if ($null -ne $Params) { $msg.params = $Params }
    $M.Proc.StandardInput.WriteLine(($msg | ConvertTo-Json -Depth 10 -Compress)); $M.Proc.StandardInput.Flush()
    if ($Notificacao) { return }
    while ($true) {
        $linha = $M.Proc.StandardOutput.ReadLine()
        if ($null -eq $linha) { throw "O servidor fechou a saída. Erro: $($M.Proc.StandardError.ReadToEnd())" }
        $r = $linha | ConvertFrom-Json -Depth 50
        if ($r.method) {   # pedido do servidor ao cliente
            $script:pedidosElicitation++
            $resp = if ($Responder) { & $Responder $r } else { @{ action = 'cancel' } }
            $M.Proc.StandardInput.WriteLine((@{ jsonrpc = '2.0'; id = $r.id; result = $resp } | ConvertTo-Json -Depth 10 -Compress)); $M.Proc.StandardInput.Flush()
            continue
        }
        return $r
    }
}
function Chamar($M, [string]$Nome, [hashtable]$Argumentos = @{}, [scriptblock]$Responder) {
    $r = Enviar $M 'tools/call' @{ name = $Nome; arguments = $Argumentos } -Responder $Responder
    return [pscustomobject]@{ Erro = [bool]$r.result.isError; Corpo = ($r.result.content[0].text | ConvertFrom-Json -Depth 50) }
}
function LinhasRegistro { if (Test-Path -LiteralPath $arqAlt) { @(Read-Alteracoes $imp | Where-Object { $_ }).Count } else { 0 } }
# Prepara e confirma; devolve a resposta da confirmação.
function Gravar($M, [string]$Nome, [hashtable]$Argumentos, [string]$Rotulo) {
    $p = Chamar $M $Nome $Argumentos
    Confere "$Rotulo`: preparou" 'False' $p.Erro
    if ($p.Erro) { $script:falhas.Add("$Rotulo`: $($p.Corpo.erro) $($p.Corpo | ConvertTo-Json -Depth 5 -Compress)"); return $null }
    Confere "$Rotulo`: nada gravado ao preparar" 'AGUARDANDO CONFIRMAÇÃO — nada foi gravado' $p.Corpo.situacao
    $antes = LinhasRegistro
    $c = Chamar $M 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao }
    Confere "$Rotulo`: gravou" 'GRAVADO' $c.Corpo.situacao
    Confere "$Rotulo`: uma linha por campo no registro" ($antes + @($p.Corpo.mudancas).Count) (LinhasRegistro)
    return $c
}
function Recusa($M, [string]$Nome, [hashtable]$Argumentos, [string]$Rotulo, [string]$Trecho) {
    $antes = LinhasRegistro
    $r = Chamar $M $Nome $Argumentos
    Confere "$Rotulo`: recusado" 'True' $r.Erro
    if ($Trecho) { Confere "$Rotulo`: motivo" 'True' ([string]$r.Corpo.erro).Contains($Trecho) }
    Confere "$Rotulo`: nada gravado" $antes (LinhasRegistro)
    return $r
}

try {
    $dir = New-Mcp 'diretoria'
    $gCas = New-Mcp 'gerente.cascavel'
    $gCha = New-Mcp 'gerente.chapeco'
    $vend = New-Mcp 'joao.almeida@horizontemaquinas.com.br'
    $sem = New-Mcp ''

    # --- catálogo das ferramentas ---
    $lista = (Enviar $dir 'tools/list' @{}).result.tools
    Confere 'ferramentas' 'consultar_vendedor,buscar_produto,ver_estoque,ver_meta,listar_oportunidades,inativar_vendedor,reativar_vendedor,alterar_meta,cadastrar_ou_editar_produto,transferir_oportunidades,confirmar_alteracao,ver_historico' (($lista | ForEach-Object name) -join ',')
    foreach ($t in $lista) {
        $esperado = if ($t.name -eq 'confirmar_alteracao') { 'False|True' } else { 'True|False' }
        Confere "$($t.name): só a confirmação grava" $esperado "$($t.annotations.readOnlyHint)|$($t.annotations.destructiveHint)"
    }

    # --- permissão (§11.2) ---
    $pedidos = @(
        @('inativar_vendedor', @{ vendedor = 'V001'; data_desligamento = '2026-08-31' }),
        @('reativar_vendedor', @{ vendedor = 'V011' }),
        @('alterar_meta', @{ vendedor = 'V001'; mes = 10; valor = 1000 }),
        @('cadastrar_ou_editar_produto', @{ produto_id = 'P001'; marca = 'X' }),
        @('transferir_oportunidades', @{ de_vendedor = 'V011'; para_vendedor = 'V001' })
    )
    foreach ($p in $pedidos) {
        [void](Recusa $sem $p[0] $p[1] "sem conta: $($p[0])" 'Nenhuma conta configurada')
        [void](Recusa $vend $p[0] $p[1] "vendedor: $($p[0])" "perfil 'vendedor'")
    }
    Confere 'vendedor lê normalmente' 'False' (Chamar $vend 'consultar_vendedor' @{ vendedor = 'V001' }).Erro
    [void](Recusa $gCas 'inativar_vendedor' @{ vendedor = 'V004'; data_desligamento = '2026-09-30' } 'gerente F01 inativa vendedor de F02' 'só altera a própria filial')
    [void](Recusa $gCas 'alterar_meta' @{ vendedor = 'V002'; mes = 10; valor = 1000 } 'gerente F01 altera meta de F03' 'só altera a própria filial')
    [void](Recusa $gCas 'transferir_oportunidades' @{ de_vendedor = 'V001'; para_vendedor = 'V004' } 'gerente F01 transfere para F02' 'só altera a própria filial')
    [void](Recusa $gCha 'reativar_vendedor' @{ vendedor = 'V005' } 'gerente F02 mexe em vendedor de F01' 'só altera a própria filial')
    Confere 'nenhum arquivo de registro antes da primeira confirmação' 'False' (Test-Path -LiteralPath $arqAlt)

    # --- inativar (§11.5) ---
    [void](Recusa $gCas 'inativar_vendedor' @{ vendedor = 'V005'; data_desligamento = '2026-08-27' } 'desligamento antes da última venda' 'última venda')
    [void](Recusa $gCas 'inativar_vendedor' @{ vendedor = 'V005'; data_desligamento = '2026-02-30' } 'data inválida' 'inválida')
    [void](Recusa $gCas 'inativar_vendedor' @{ vendedor = 'V005' } 'sem data' 'obrigatória')
    [void](Recusa $gCha 'inativar_vendedor' @{ vendedor = 'V011'; data_desligamento = '2026-09-30' } 'já inativo' 'já está inativo')
    $p = Chamar $gCas 'inativar_vendedor' @{ vendedor = 'rafael'; data_desligamento = '31/08/2026' }
    Confere 'inativar V005: campos' 'vendedor/Status/Ativo/Inativo;vendedor/Data de Desligamento//2026-08-31;meta/Meta (R$)/690000.00/;meta/Meta (R$)/630000.00/;meta/Meta (R$)/540000.00/;meta/Meta (R$)/510000.00/' (($p.Corpo.mudancas | ForEach-Object { "$($_.tipo)/$($_.campo)/$($_.atual)/$($_.novo)" }) -join ';')
    Confere 'inativar V005: avisa as 7 órfãs' 'True' (($p.Corpo.efeitos -join ' ').Contains('7 oportunidade(s) aberta(s) viram órfãs'))
    [void](Recusa $dir 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao } 'código de outra conta (outro processo)' 'desconhecido')
    [void](Recusa $gCas 'confirmar_alteracao' @{ codigo = 'ZZZZ0000' } 'código inventado' 'desconhecido')
    $c = Chamar $gCas 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao }
    Confere 'inativar V005: gravado' 'GRAVADO|6|gerente.cascavel' "$($c.Corpo.situacao)|$($c.Corpo.campos_gravados)|$($c.Corpo.quem.login)"
    [void](Recusa $gCas 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao } 'código reusado' 'já usado')
    $cv = (Chamar $dir 'consultar_vendedor' @{ vendedor = 'V005' }).Corpo
    Confere 'MCP lê a própria gravação' 'Inativo|2026-08-31' "$($cv.vendedor.status)|$($cv.vendedor.desligamento)"

    # --- alterar meta (§11.7); com a base, a data-base de vendas é 31/08: agosto a dezembro estão abertos ---
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 7; valor = 1000 } 'mês fechado' 'já fechou')
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 10; valor = 0 } 'meta zero' 'maior que zero')
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 10; valor = '455.000,505' } 'três casas' 'mais de 2 casas')
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 10; valor = 10; sem_meta = $true } 'valor e sem_meta juntos' 'não os dois')
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 10 } 'sem valor' "Informe 'valor'")
    [void](Recusa $gCas 'alterar_meta' @{ vendedor = 'V005'; mes = 10; valor = 1000 } 'meta depois do desligamento' 'desligado')
    $metaV003Out = (Get-Desempenho -Base (Import-BaseComercial -DirDados $tmp) -MesInicial 10 -MesFinal 10).Linhas | Where-Object { $_.Vendedor.Id -eq 'V003' }
    [void](Recusa $dir 'alterar_meta' @{ vendedor = 'V003'; mes = 10; valor = [string]$metaV003Out.Meta } 'mesmo valor' 'Nada a alterar')
    $c = Gravar $dir 'alterar_meta' @{ vendedor = 'mariana'; mes = 10; valor = '700.000,50' } 'meta V003 out'
    Confere 'meta V003 out: anterior → novo' "$(([decimal]$metaV003Out.Meta).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture))|700000.50" "$($c.Corpo.mudancas[0].anterior)|$($c.Corpo.mudancas[0].novo)"
    [void](Gravar $dir 'alterar_meta' @{ vendedor = 'V012'; mes = 9; sem_meta = $true } 'meta V012 set → sem meta')
    [void](Gravar $gCas 'alterar_meta' @{ vendedor = 'V001'; mes = 8; valor = 600000 } 'meta V001 ago (mês da data-base)')

    # --- confirmação refaz tudo (§11.1): o resumo ficou velho ---
    $velho = Chamar $dir 'alterar_meta' @{ vendedor = 'V001'; mes = 11; valor = 1000 }
    [void](Gravar $gCas 'alterar_meta' @{ vendedor = 'V001'; mes = 11; valor = 2000 } 'meta V001 nov por outro gerente')
    $antes = LinhasRegistro
    $r = Chamar $dir 'confirmar_alteracao' @{ codigo = $velho.Corpo.codigo_confirmacao }
    Confere 'resumo velho: não grava' "True|$antes" "$($r.Erro)|$(LinhasRegistro)"
    Confere 'resumo velho: devolve o novo (2000.00 → 1000.00)' '2000.00|1000.00' "$($r.Corpo.resumo_novo.mudancas[0].atual)|$($r.Corpo.resumo_novo.mudancas[0].novo)"
    $c = Chamar $dir 'confirmar_alteracao' @{ codigo = $r.Corpo.resumo_novo.codigo_confirmacao }
    Confere 'resumo novo confirmado' 'GRAVADO' $c.Corpo.situacao

    # --- produto (§11.8) ---
    [void](Recusa $gCha 'cadastrar_ou_editar_produto' @{ nome = 'Sem unidade'; categoria = 'Tratores'; preco_tabela = 1 } 'produto novo sem unidade' 'Unidade')
    [void](Recusa $gCha 'cadastrar_ou_editar_produto' @{ nome = 'X'; categoria = 'Drones'; preco_tabela = 1; unidade = 'UN' } 'categoria nova' 'não existe no catálogo')
    [void](Recusa $gCha 'cadastrar_ou_editar_produto' @{ produto_id = 'P999'; nome = 'X' } 'código inexistente' 'não existe')
    [void](Recusa $gCha 'cadastrar_ou_editar_produto' @{ produto_id = 'P008'; categoria = 'tratores'; preco_tabela = '830000' } 'edição sem diferença' 'Nada a alterar')
    [void](Recusa $gCha 'cadastrar_ou_editar_produto' @{ produto_id = 'P008'; status = 'Suspenso' } 'status inválido' 'inválido')
    $nomeNovo = 'Plantadeira Teste "PT-9"; linha 2 <b>'
    $c = Gravar $gCha 'cadastrar_ou_editar_produto' @{ nome = $nomeNovo; categoria = 'plantadeiras'; marca = 'Terramax'; preco_tabela = '98.500,00'; unidade = 'UN' } 'produto novo'
    Confere 'produto novo: próximo código livre e status padrão' 'P066|Ativo' "$($c.Corpo.mudancas[0].registro)|$(@($c.Corpo.mudancas | Where-Object campo -eq 'Status').novo)"
    [void](Gravar $gCas 'cadastrar_ou_editar_produto' @{ produto_id = 'p008'; status = 'descontinuado'; preco_tabela = 845000 } 'P008 descontinuado e preço')
    $bp = (Chamar $dir 'buscar_produto' @{ busca = 'P066' }).Corpo
    Confere 'buscar_produto vê o produto novo' "P066|$nomeNovo|Plantadeiras|98500|0" "$($bp.produtos[0].id)|$($bp.produtos[0].produto)|$($bp.produtos[0].categoria)|$($bp.produtos[0].preco_tabela)|$($bp.produtos[0].estoque_unidades)"

    # --- transferir (§11.9) ---
    [void](Recusa $dir 'transferir_oportunidades' @{ de_vendedor = 'V011'; para_vendedor = 'V005' } 'destino desligado' 'desligado')
    [void](Recusa $dir 'transferir_oportunidades' @{ de_vendedor = 'V011'; para_vendedor = 'V011' } 'mesmo dono' 'diferente')
    $r = Recusa $dir 'transferir_oportunidades' @{ de_vendedor = 'V011'; para_vendedor = 'V002'; oportunidades = 'OP-0002, OP-0001, OP-9999' } 'fechada, de outro dono e inexistente' 'nada foi preparado'
    Confere 'transferência: lista os três problemas' 3 @($r.Corpo.problemas).Count
    $c = Gravar $dir 'transferir_oportunidades' @{ de_vendedor = 'V011'; para_vendedor = 'V002'; oportunidades = 'op-0130' } 'transferir OP-0130'
    Confere 'transferir OP-0130: avisa a mudança de filial' 'True' ((@($c.Corpo.efeitos) -join ' ').Contains('mudam de filial'))
    [void](Gravar $gCas 'transferir_oportunidades' @{ de_vendedor = 'V005'; para_vendedor = 'V001' } 'órfãs de V005 para V001 (todas)')

    # --- reativar (§11.6) ---
    $c = Gravar $gCha 'reativar_vendedor' @{ vendedor = 'Ricardo' } 'reativar V011'
    Confere 'reativar V011: data sai e fica no registro' '2026-04-30|' "$($c.Corpo.mudancas[1].anterior)|$($c.Corpo.mudancas[1].novo)"
    [void](Recusa $gCha 'reativar_vendedor' @{ vendedor = 'V011' } 'já ativo' 'já está ativo')

    # --- confirmação direto à pessoa (elicitation), quando o cliente sabe perguntar ---
    $eli = New-Mcp 'diretoria' -ComElicitation
    $script:pedidosElicitation = 0
    $p = Chamar $eli 'alterar_meta' @{ vendedor = 'V007'; mes = 12; valor = 123456 }
    $antes = LinhasRegistro
    $c = Chamar $eli 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao } -Responder { param($req) @{ action = 'decline' } }
    Confere 'elicitation recusada: não grava' "1|NÃO GRAVADO — a pessoa não confirmou|$antes" "$($script:pedidosElicitation)|$($c.Corpo.situacao)|$(LinhasRegistro)"
    $p = Chamar $eli 'alterar_meta' @{ vendedor = 'V007'; mes = 12; valor = 123456 }
    $c = Chamar $eli 'confirmar_alteracao' @{ codigo = $p.Corpo.codigo_confirmacao } -Responder { param($req) $script:msgElicitation = $req.params.message; @{ action = 'accept'; content = @{ confirmar = $true } } }
    Confere 'elicitation aceita: grava' "2|GRAVADO|$($antes + 1)" "$($script:pedidosElicitation)|$($c.Corpo.situacao)|$(LinhasRegistro)"
    Confere 'elicitation mostra o que muda' 'True' ([string]$script:msgElicitation).Contains('R$ 123.456,00')

    # --- registro (§11.3) ---
    $alts = @(Read-Alteracoes $imp | Where-Object { $_ })
    $h = (Chamar $vend 'ver_historico' @{ limite = 300 }).Corpo
    Confere 'ver_historico: tudo, do mais recente ao mais antigo' "$($alts.Count)|V007/2026-12" "$($h.total_no_registro)|$($h.alteracoes[0].registro)"
    Confere 'ver_historico: filtro (as 4 metas tiradas de V005)' 4 (Chamar $dir 'ver_historico' @{ registro = 'V005/'; tipo = 'meta' }).Corpo.encontradas
    Confere 'registro: arquivo com BOM e cabeçalho' 'EF-BB-BF' ([BitConverter]::ToString([IO.File]::ReadAllBytes($arqAlt), 0, 3))
    foreach ($a in $alts) {
        $okLinha = ($a.Quando -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') -and $a.Lote -and $a.Login -in @('diretoria', 'gerente.cascavel', 'gerente.chapeco') -and $a.Nome -and $a.Perfil -in @('diretoria', 'gerente') -and $a.'Ação' -and $a.Tipo -in @('vendedor', 'meta', 'produto', 'oportunidade') -and $a.Registro -and $a.Campo
        Confere "registro: linha completa ($($a.Registro) $($a.Campo))" 'True' $okLinha
    }
    Confere 'registro: produto com aspas e ponto e vírgula intacto' $nomeNovo (@($alts | Where-Object { $_.Registro -eq 'P066' -and $_.Campo -eq 'Produto' })[0].Novo)

    # --- a versão PowerShell aplica o registro (§11.4) ---
    $b = Import-BaseComercial -DirDados $tmp -DirImportacoes $imp
    Confere 'PowerShell: base em vigor sem erros' '' ($b.Erros -join ' | ')
    Confere 'PowerShell: V005 inativo em 31/08' 'False|2026-08-31' "$($b.VendedorPorId['V005'].Ativo)|$($b.VendedorPorId['V005'].Desligamento.ToString('yyyy-MM-dd'))"
    Confere 'PowerShell: V011 reativado' 'True|' "$($b.VendedorPorId['V011'].Ativo)|$($b.VendedorPorId['V011'].Desligamento)"
    $meta = { param($id, $mes) $x = @($b.Metas | Where-Object { $_.IdVendedor -eq $id -and $_.Mes -eq $mes }); if ($x.Count) { ([decimal]$x[0].Valor).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) } else { 'sem meta' } }
    Confere 'PowerShell: metas' '700000.50|sem meta|600000.00|1000.00|123456.00|sem meta|sem meta' "$(& $meta 'V003' 10)|$(& $meta 'V012' 9)|$(& $meta 'V001' 8)|$(& $meta 'V001' 11)|$(& $meta 'V007' 12)|$(& $meta 'V005' 9)|$(& $meta 'V005' 12)"
    $pp = Import-BasePipeline -DirDados $tmp -Comercial $b -DirImportacoes $imp
    Confere 'PowerShell: pipeline sem erros' '' ($pp.Erros -join ' | ')
    $dono = { param($op) ($pp.Oportunidades | Where-Object Id -eq $op).IdVendedor }
    Confere 'PowerShell: donos transferidos' 'V002|V011|V001' "$(& $dono 'OP-0130')|$(& $dono 'OP-0140')|$(& $dono 'OP-0165')"
    $vp = Get-VisaoPipeline -Comercial $b -Desempenho (Get-Desempenho -Base $b -MesInicial 1 -MesFinal 8) -Pipeline $pp
    Confere 'PowerShell: nenhuma órfã (V011 reativado, órfãs de V005 transferidas)' 0 $vp.QtdOrfas
    $e = Import-BaseEstoque -DirDados $tmp -DirImportacoes $imp
    Confere 'PowerShell: estoque sem erros' '' ($e.Erros -join ' | ')
    Confere 'PowerShell: P066 e P008' "Plantadeiras|Ativo|Descontinuado" "$($e.ProdutoPorId['P066'].Categoria)|$($e.ProdutoPorId['P066'].Status)|$($e.ProdutoPorId['P008'].Status)"
    $ve = Get-VisaoEstoque -Estoque $e -DiasParado 180 -AbertasPorProduto (Get-AbertasPorProduto $pp)
    Confere 'PowerShell: produto novo aparece em sem estoque (§11.8)' 9 @($ve.SemEstoque).Count
    $html = New-PaginaHistorico $alts '' '' ([pscustomobject]@{ nome = 'Teste'; perfil = 'diretoria' })
    Confere 'tela Histórico: conta todas' 'True' $html.Contains("$($alts.Count) de $($alts.Count) alterações registradas")
    Confere 'tela Histórico: escapa o nome do produto' 'True' ($html.Contains('&lt;b&gt;') -and -not $html.Contains('<b>'))

    # --- a base pura continua a da §6 (§11.4 e §9.9) ---
    $b0 = Import-BaseComercial -DirDados $tmp
    $d0 = Get-Desempenho -Base $b0 -MesInicial 1 -MesFinal 8
    Confere '§6 com a base pura' '45295119.16|50720000.00|False' "$(([decimal]$d0.Empresa.Realizado).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture))|$(([decimal]$d0.Empresa.Meta).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture))|$(@($b0.Vendedores | Where-Object Id -eq 'V005')[0].Ativo -eq $false)"
    Confere 'planilhas originais intactas' 'True' ((Get-FileHash (Join-Path $tmp 'metas_2026.xlsx')).Hash -eq (Get-FileHash (Join-Path $DirDados 'metas_2026.xlsx')).Hash -and (Get-FileHash (Join-Path $tmp 'vendedores.xlsx')).Hash -eq (Get-FileHash (Join-Path $DirDados 'vendedores.xlsx')).Hash)
} finally {
    foreach ($m in $script:processos) { try { $m.Proc.StandardInput.Close() } catch { }; if (-not $m.Proc.WaitForExit(5000)) { $m.Proc.Kill() } }
    Remove-Item $tmp -Recurse -Force
}

if ($script:falhas.Count) {
    $script:falhas | ForEach-Object { Write-Host "  FALHA $_" -ForegroundColor Red }
    Write-Host "$($script:falhas.Count) falhas em $($script:ok + $script:falhas.Count) conferências da escrita." -ForegroundColor Red
    exit 1
}
Write-Host "Todas as $($script:ok) conferências da escrita bateram." -ForegroundColor Green
