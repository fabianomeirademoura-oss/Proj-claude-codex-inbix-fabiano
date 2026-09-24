# Servidor do painel de desempenho comercial — Horizonte Máquinas.
# Uso: .\iniciar.cmd   (ou: powershell -ExecutionPolicy Bypass -File servidor.ps1 [-Porta 8080])

param(
    [int]$Porta = 8080,
    [string]$DirDados,
    [string]$ArquivoUsuarios
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot 'lib\Regras.ps1')
. (Join-Path $PSScriptRoot 'lib\Auth.ps1')
. (Join-Path $PSScriptRoot 'lib\Paginas.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasPipeline.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasPipeline.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasFunil.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasEstoque.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasOrigem.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasOrigem.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasEstoque.ps1')
. (Join-Path $PSScriptRoot 'lib\RegrasImportacao.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasImportacao.ps1')
. (Join-Path $PSScriptRoot 'lib\PaginasHistorico.ps1')

if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
if (-not $ArquivoUsuarios) { $ArquivoUsuarios = Join-Path $PSScriptRoot 'usuarios.json' }
if (-not (Test-Path -LiteralPath $ArquivoUsuarios)) {
    Write-Host "Arquivo de usuários não encontrado: $ArquivoUsuarios" -ForegroundColor Red
    Write-Host 'Rode primeiro: criar_usuarios.cmd' -ForegroundColor Yellow
    exit 1
}
# REGRAS_NEGOCIO.md §9.8: o painel aplica as importações confirmadas por cima da base.
$script:DirImportacoes = Get-DirImportacoes $DirDados
$script:LimiteUpload = 10MB

$script:Base = $null
$script:AssinaturaBase = $null

function Update-Base {
    # Recarrega as planilhas sempre que alguma delas mudar no disco.
    $assinatura = Get-AssinaturaBase $DirDados $script:DirImportacoes
    if ($assinatura -eq $script:AssinaturaBase) { return }
    try {
        $script:Base = Import-BaseComercial -DirDados $DirDados -DirImportacoes $script:DirImportacoes
    } catch {
        $script:Base = [pscustomobject]@{ Erros = @("Falha ao ler as planilhas: $($_.Exception.Message)") }
    }
    $script:AssinaturaBase = $assinatura
    if ($script:Base.Erros.Count) { Write-Host "Importação com $($script:Base.Erros.Count) erro(s)." -ForegroundColor Red }
    else { Write-Host "Planilhas carregadas: $($script:Base.Vendas.Count) vendas, $($script:Base.Metas.Count) metas, $($script:Base.Vendedores.Count) vendedores." -ForegroundColor Green }
}

$script:Pipeline = $null
$script:AssinaturaPipeline = $null

function Update-Pipeline {
    # CRM e clientes; depende do cadastro de vendedores, então recarrega também quando a base comercial muda.
    $assinatura = "$(Get-AssinaturaPipeline $DirDados $script:DirImportacoes)#$($script:AssinaturaBase)"
    if ($assinatura -eq $script:AssinaturaPipeline) { return }
    try {
        $script:Pipeline = Import-BasePipeline -DirDados $DirDados -Comercial $script:Base -DirImportacoes $script:DirImportacoes
    } catch {
        $script:Pipeline = [pscustomobject]@{ Erros = @("Falha ao ler o CRM: $($_.Exception.Message)") }
    }
    $script:AssinaturaPipeline = $assinatura
    if ($script:Pipeline.Erros.Count) { Write-Host "CRM com $($script:Pipeline.Erros.Count) erro(s) de importação." -ForegroundColor Red }
    else { Write-Host "CRM carregado: $($script:Pipeline.Oportunidades.Count) oportunidades." -ForegroundColor Green }
}

$script:Estoque = $null
$script:AssinaturaEstoque = $null

function Update-Estoque {
    # Foto de estoque e catálogo; recarrega quando qualquer um deles muda (§5: foto substitutiva).
    $assinatura = Get-AssinaturaEstoque $DirDados $script:DirImportacoes
    if ($assinatura -eq $script:AssinaturaEstoque) { return }
    try {
        $script:Estoque = Import-BaseEstoque -DirDados $DirDados -DirImportacoes $script:DirImportacoes
    } catch {
        $script:Estoque = [pscustomobject]@{ Erros = @("Falha ao ler o estoque: $($_.Exception.Message)") }
    }
    $script:AssinaturaEstoque = $assinatura
    if ($script:Estoque.Erros.Count) { Write-Host "Estoque com $($script:Estoque.Erros.Count) erro(s) de importação." -ForegroundColor Red }
    else { Write-Host "Estoque carregado: foto de $($script:Estoque.DataBase.ToString('dd/MM/yyyy')), $($script:Estoque.Linhas.Count) linhas, $($script:Estoque.Produtos.Count) produtos." -ForegroundColor Green }
}

function Get-DiasParado($Ctx) {
    # REGRAS_NEGOCIO.md §5.0: prazo escolhido na tela; vazio usa o padrão; inválido usa o padrão e avisa.
    $texto = [string]$Ctx.Request.QueryString['dias']
    if (-not $texto) { return @($script:DiasParadoPadrao, '') }
    $n = 0
    if ([int]::TryParse($texto, [ref]$n) -and $n -ge 1 -and $n -le $script:DiasParadoMaximo) { return @($n, '') }
    return @($script:DiasParadoPadrao, "Prazo inválido: use um número inteiro de 1 a $($script:DiasParadoMaximo) dias. Mostrando o padrão de $($script:DiasParadoPadrao) dias.")
}

function Send-Resposta($Ctx, [int]$Status, [string]$Tipo, [string]$Corpo, [switch]$ComBom) {
    $r = $Ctx.Response
    $r.StatusCode = $Status
    $r.ContentType = $Tipo
    $r.Headers['Content-Security-Policy'] = "default-src 'none'; style-src 'self'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'"
    $r.Headers['X-Content-Type-Options'] = 'nosniff'
    $r.Headers['Referrer-Policy'] = 'no-referrer'
    $r.Headers['Cache-Control'] = 'no-store'
    $bytes = [Text.Encoding]::UTF8.GetBytes($Corpo)
    if ($ComBom) { $bytes = [byte[]](0xEF, 0xBB, 0xBF) + $bytes }
    $r.ContentLength64 = $bytes.Length
    $r.OutputStream.Write($bytes, 0, $bytes.Length)
    $r.OutputStream.Close()
}

function Send-Redirecionamento($Ctx, [string]$Destino, [string]$Cookie) {
    $r = $Ctx.Response
    $r.StatusCode = 303
    $r.RedirectLocation = $Destino
    if ($Cookie) { $r.AppendHeader('Set-Cookie', $Cookie) }
    $r.Headers['Cache-Control'] = 'no-store'
    $r.ContentLength64 = 0
    $r.OutputStream.Close()
}

function Read-Formulario($Ctx) {
    $campos = @{}
    if ($Ctx.Request.ContentLength64 -gt 8192) { return $campos }
    $leitor = New-Object IO.StreamReader($Ctx.Request.InputStream, [Text.Encoding]::UTF8)
    try { $texto = $leitor.ReadToEnd() } finally { $leitor.Dispose() }
    foreach ($par in ($texto -split '&')) {
        if (-not $par) { continue }
        $k, $v = $par -split '=', 2
        $campos[[Net.WebUtility]::UrlDecode($k)] = [Net.WebUtility]::UrlDecode([string]$v)
    }
    return $campos
}

function Read-Multipart($Ctx, [long]$Limite) {
    # multipart/form-data do formulário de importação. Devolve $null se passar do limite.
    # Os bytes são lidos como Latin-1 (1 byte = 1 caractere) para separar as partes sem estragar o arquivo.
    $req = $Ctx.Request
    if ($req.ContentType -notmatch 'multipart/form-data;\s*boundary=("?)([^";]+)\1') { return @{ Campos = @{}; Arquivos = @{} } }
    $fronteira = '--' + $Matches[2]
    if ($req.ContentLength64 -gt $Limite) { return $null }
    $ms = New-Object IO.MemoryStream
    $buf = New-Object byte[] 65536
    while (($n = $req.InputStream.Read($buf, 0, $buf.Length)) -gt 0) {
        $ms.Write($buf, 0, $n)
        if ($ms.Length -gt $Limite) { return $null }
    }
    $latin = [Text.Encoding]::GetEncoding(28591)
    $texto = "`r`n" + $latin.GetString($ms.ToArray())
    $campos = @{}; $arquivos = @{}
    foreach ($parte in ($texto -split [regex]::Escape("`r`n$fronteira"))) {
        if (-not $parte -or $parte.StartsWith('--')) { continue }
        $parte = $parte.TrimStart("`r", "`n")
        $fim = $parte.IndexOf("`r`n`r`n")
        if ($fim -lt 0) { continue }
        $cab = [Text.Encoding]::UTF8.GetString($latin.GetBytes($parte.Substring(0, $fim)))
        $corpo = $parte.Substring($fim + 4)
        if ($cab -notmatch 'name="([^"]*)"') { continue }
        $nome = $Matches[1]
        if ($cab -match 'filename="([^"]*)"') {
            $arquivo = ($Matches[1] -split '[\\/]')[-1]   # navegadores antigos mandam o caminho inteiro
            $arquivos[$nome] = [pscustomobject]@{ Nome = $arquivo; Bytes = [byte[]]$latin.GetBytes($corpo) }
        } else {
            $campos[$nome] = [Text.Encoding]::UTF8.GetString($latin.GetBytes($corpo))
        }
    }
    return @{ Campos = $campos; Arquivos = $arquivos }
}

function Get-EstadoImportacao {
    # O que está em vigor hoje, para a tela de importação.
    Update-Base; Update-Estoque
    $b = $script:Base; $e = $script:Estoque
    $erro = if ($b.Erros.Count) { "As planilhas em vigor têm $($b.Erros.Count) erro(s). Veja a tela Vendas e metas." } elseif ($e.Erros.Count) { "A foto de estoque em vigor tem $($e.Erros.Count) erro(s). Veja a tela Estoque." } else { '' }
    return [pscustomobject]@{
        QtdVendas      = if ($b.Vendas) { $b.Vendas.Count } else { 0 }
        DataBaseVendas = $b.DataBase
        ArquivosVendas = @(Get-ArquivosVendasImportados $script:DirImportacoes | ForEach-Object Name)
        DataFoto       = $e.DataBase
        ArquivoFoto    = if ($e.Arquivo) { $e.Arquivo.Name } else { '' }
        Historico      = @(Get-HistoricoImportacoes $DirDados)
        ErroBase       = $erro
    }
}

function Invoke-Importacao($Ctx, $Usuario, [string]$Caminho, [string]$Metodo) {
    # REGRAS_NEGOCIO.md §9.1: só a diretoria importa.
    if ($Usuario.perfil -ne 'diretoria') {
        Send-Resposta $Ctx 403 'text/html; charset=utf-8' (New-Layout 'Sem acesso' '<section class="aviso-erro"><h1>Sem acesso</h1><p>Só a diretoria importa arquivos (REGRAS_NEGOCIO.md §9.1).</p></section>' $Usuario 'importar'); return
    }
    $html = 'text/html; charset=utf-8'
    if ($Caminho -eq '/importar' -and $Metodo -eq 'GET') { Send-Resposta $Ctx 200 $html (New-PaginaImportacao (Get-EstadoImportacao) $Usuario ''); return }

    if ($Caminho -eq '/importar' -and $Metodo -eq 'POST') {
        $form = Read-Multipart $Ctx $script:LimiteUpload
        if ($null -eq $form) { Send-Resposta $Ctx 413 $html (New-PaginaImportacao (Get-EstadoImportacao) $Usuario "O arquivo passa do limite de $($script:LimiteUpload / 1MB) MB."); return }
        $tipo = [string]$form.Campos['tipo']
        $arq = $form.Arquivos['arquivo']
        if ($tipo -notin @('vendas', 'estoque')) { Send-Resposta $Ctx 400 $html (New-PaginaImportacao (Get-EstadoImportacao) $Usuario 'Escolha se o arquivo é de vendas ou de estoque.'); return }
        if (-not $arq -or -not $arq.Nome -or -not $arq.Bytes.Length) { Send-Resposta $Ctx 400 $html (New-PaginaImportacao (Get-EstadoImportacao) $Usuario 'Escolha um arquivo .xlsx para enviar.'); return }
        $token = New-ImportacaoPendente -DirDados $DirDados -Bytes $arq.Bytes -NomeOriginal $arq.Nome -Tipo $tipo -Login $Usuario.login
        $pend = Get-ImportacaoPendente $DirDados $token $Usuario.login
        try {
            $resumo = Get-ResumoImportacao -DirDados $DirDados -Caminho $pend.Caminho -Tipo $tipo -NomeOriginal $arq.Nome
        } catch {
            $resumo = New-Recusa $tipo $arq.Nome @("O arquivo não pôde ser analisado: $($_.Exception.Message)")
        }
        if ($resumo.Recusado) {
            Remove-ImportacaoPendente $pend
            Write-Host "Importação recusada ($tipo): $($arq.Nome)" -ForegroundColor Yellow
            Send-Resposta $Ctx 422 $html (New-PaginaRecusaImportacao $resumo $Usuario); return
        }
        Send-Resposta $Ctx 200 $html (New-PaginaResumoImportacao $resumo $token $Usuario); return
    }

    if ($Caminho -in @('/importar/confirmar', '/importar/cancelar') -and $Metodo -eq 'POST') {
        $pend = Get-ImportacaoPendente $DirDados ([string](Read-Formulario $Ctx)['token']) $Usuario.login
        if ($Caminho -eq '/importar/cancelar') { Remove-ImportacaoPendente $pend; Send-Redirecionamento $Ctx '/importar'; return }
        if (-not $pend) { Send-Resposta $Ctx 409 $html (New-PaginaImportacao (Get-EstadoImportacao) $Usuario 'Esta importação expirou ou já foi tratada. Envie o arquivo de novo.'); return }
        try {
            $resumo = Confirm-Importacao -DirDados $DirDados -Pendente $pend
        } catch {
            $resumo = New-Recusa $pend.Tipo $pend.NomeOriginal @("A gravação falhou: $($_.Exception.Message)")
        }
        if ($resumo.Recusado) {
            Remove-ImportacaoPendente $pend
            Send-Resposta $Ctx 422 $html (New-PaginaRecusaImportacao $resumo $Usuario); return
        }
        Write-Host "Importação gravada ($($resumo.Tipo)) por $($Usuario.login): $($resumo.ArquivoGravado)" -ForegroundColor Cyan
        Update-Base; Update-Estoque
        Send-Resposta $Ctx 200 $html (New-PaginaImportacaoGravada $resumo $Usuario); return
    }
    Send-Resposta $Ctx 404 'text/plain; charset=utf-8' 'Página não encontrada'
}

function Get-TokenCookie($Ctx) {
    $c = $Ctx.Request.Cookies['sessao']
    if ($c) { return $c.Value }
    return $null
}

function Get-Periodo($Ctx, $Base) {
    $meses = Get-MesesDisponiveis $Base
    $max = if ($meses.Count) { $meses[-1] } else { 1 }
    $de = 1; $ate = $max
    [void][int]::TryParse($Ctx.Request.QueryString['de'], [ref]$de)
    [void][int]::TryParse($Ctx.Request.QueryString['ate'], [ref]$ate)
    if (-not $Ctx.Request.QueryString['de']) { $de = 1 }
    if (-not $Ctx.Request.QueryString['ate']) { $ate = $max }
    $de = [math]::Max(1, [math]::Min($de, $max))
    $ate = [math]::Max(1, [math]::Min($ate, $max))
    if ($de -gt $ate) { $de, $ate = $ate, $de }
    return @($de, $ate)
}

function Invoke-Entrar($Ctx) {
    $form = Read-Formulario $Ctx
    $login = ([string]$form['login']).Trim().ToLowerInvariant()
    $senha = [string]$form['senha']
    $generica = 'Usuário ou senha inválidos.'
    if (-not $login -or -not $senha) { Send-Resposta $Ctx 400 'text/html; charset=utf-8' (New-PaginaLogin $generica $login); return }
    if (Test-Bloqueado $login) {
        Send-Resposta $Ctx 429 'text/html; charset=utf-8' (New-PaginaLogin 'Muitas tentativas. Aguarde alguns minutos e tente de novo.' $login); return
    }
    $usuario = Read-Usuarios $ArquivoUsuarios | Where-Object { $_.login -eq $login } | Select-Object -First 1
    if (-not $usuario -or -not (Test-Senha $usuario $senha)) {
        Register-Falha $login
        Write-Host "Login recusado: $login" -ForegroundColor Yellow
        Send-Resposta $Ctx 401 'text/html; charset=utf-8' (New-PaginaLogin $generica $login); return
    }
    # REGRAS_NEGOCIO.md §2.2: vendedor inativo não faz login. Confere no cadastro atual.
    if ($usuario.idVendedor) {
        Update-Base
        $vend = if (-not $script:Base.Erros.Count) { $script:Base.VendedorPorId[$usuario.idVendedor] }
        if (-not $vend -or -not $vend.Ativo) {
            Write-Host "Login recusado (vendedor inativo ou não encontrado): $login" -ForegroundColor Yellow
            Send-Resposta $Ctx 403 'text/html; charset=utf-8' (New-PaginaLogin 'Acesso desativado. Procure a diretoria comercial.' $login); return
        }
    }
    $token = New-Sessao $usuario
    Write-Host "Login: $login" -ForegroundColor Cyan
    Send-Redirecionamento $Ctx '/' "sessao=$token; Path=/; HttpOnly; SameSite=Strict"
}

function Invoke-Requisicao($Ctx) {
    $req = $Ctx.Request
    $caminho = $req.Url.AbsolutePath
    $metodo = $req.HttpMethod

    if ($caminho -eq '/estilo.css' -and $metodo -eq 'GET') { Send-Resposta $Ctx 200 'text/css; charset=utf-8' (Get-Css); return }

    $token = Get-TokenCookie $Ctx
    $sessao = Get-Sessao $token

    if ($caminho -eq '/entrar') {
        if ($metodo -eq 'POST') { Invoke-Entrar $Ctx; return }
        if ($sessao) { Send-Redirecionamento $Ctx '/'; return }
        Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaLogin '' ''); return
    }
    if ($caminho -eq '/sair' -and $metodo -eq 'POST') {
        Remove-Sessao $token
        Send-Redirecionamento $Ctx '/entrar' 'sessao=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0'; return
    }

    if (-not $sessao) { Send-Redirecionamento $Ctx '/entrar'; return }
    if ($caminho -eq '/importar' -or $caminho.StartsWith('/importar/')) { Invoke-Importacao $Ctx $sessao.Usuario $caminho $metodo; return }
    if ($metodo -ne 'GET') { Send-Resposta $Ctx 405 'text/plain; charset=utf-8' 'Método não permitido'; return }
    $usuario = $sessao.Usuario

    if ($caminho -eq '/historico' -or $caminho -eq '/historico.csv') {
        # REGRAS_NEGOCIO.md §11.3: diretoria e gerentes. O registro é lido a cada pedido (só acréscimos).
        if (-not (Test-PodeVerHistorico $usuario)) { Send-Resposta $Ctx 403 'text/html; charset=utf-8' (New-PaginaSemAcessoHistorico $usuario); return }
        $tipo = [string]$req.QueryString['tipo']
        if ($tipo -cnotin @($script:TiposAlteracao.Keys)) { $tipo = '' }
        $registro = ([string]$req.QueryString['registro']).Trim()
        $alteracoes = @(Read-Alteracoes $script:DirImportacoes | Where-Object { $_ })
        if ($caminho -eq '/historico') { Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaHistorico $alteracoes $tipo $registro $usuario) }
        else {
            $Ctx.Response.AddHeader('Content-Disposition', 'attachment; filename="historico_alteracoes.csv"')
            Send-Resposta $Ctx 200 'text/csv; charset=utf-8' (New-CsvHistorico $alteracoes $tipo $registro) -ComBom
        }
        return
    }

    Update-Base
    $base = $script:Base

    if ($base.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $base.Erros $usuario); return }

    $de, $ate = Get-Periodo $Ctx $base
    $incluirDesligados = ($req.QueryString['desligados'] -eq '1')
    $desempenho = Get-Desempenho -Base $base -MesInicial $de -MesFinal $ate

    if ($caminho -eq '/estoque' -or $caminho -eq '/estoque.csv') {
        Update-Estoque
        if ($script:Estoque.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $script:Estoque.Erros $usuario 'estoque'); return }
        Update-Pipeline   # só para contar oportunidades abertas por produto (§5.4); se falhar, a tela avisa
        $dias, $avisoPrazo = Get-DiasParado $Ctx
        $visao = Get-VisaoEstoque -Estoque $script:Estoque -DiasParado $dias -AbertasPorProduto (Get-AbertasPorProduto $script:Pipeline)
        if ($caminho -eq '/estoque') {
            $filial = [string]$req.QueryString['filial']
            if ($filial -and -not ($visao.Filiais | Where-Object Id -eq $filial)) { $filial = '' }
            Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaEstoque $visao $usuario $filial $avisoPrazo)
        } else {
            $Ctx.Response.AddHeader('Content-Disposition', ('attachment; filename="estoque_{0:yyyy-MM-dd}_parado{1}d.csv"' -f $visao.DataBase, $dias))
            Send-Resposta $Ctx 200 'text/csv; charset=utf-8' (New-CsvEstoque $visao) -ComBom
        }
        return
    }

    if ($caminho -eq '/funil') {
        Update-Pipeline
        if ($script:Pipeline.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $script:Pipeline.Erros $usuario 'funil'); return }
        # Filtros (REGRAS_NEGOCIO.md §8.4): só valores conhecidos; qualquer outro vira "sem filtro".
        $qs = $req.QueryString
        $fVend = [string]$qs['vendedor'];   if ($fVend -and -not $base.VendedorPorId.ContainsKey($fVend)) { $fVend = '' }
        $fFilial = [string]$qs['filial'];   if ($fFilial -and -not ($script:Pipeline.Filiais | Where-Object Id -eq $fFilial)) { $fFilial = '' }
        $fCat = [string]$qs['categoria']
        $categorias = @($script:Pipeline.ProdutoPorId.Values | ForEach-Object Categoria)
        if ($fCat -and $fCat -notin $categorias) { $fCat = '' }
        $medida = if ($qs['medida'] -eq 'valor') { 'valor' } else { 'qtd' }
        $visao = Get-VisaoFunil -Comercial $base -Pipeline $script:Pipeline -IdVendedor $fVend -IdFilial $fFilial -Categoria $fCat
        Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaFunil $visao $usuario $medida)
        return
    }

    if ($caminho -eq '/pipeline' -or $caminho -eq '/pipeline.csv') {
        Update-Pipeline
        if ($script:Pipeline.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $script:Pipeline.Erros $usuario 'pipeline'); return }
        $visao = Get-VisaoPipeline -Comercial $base -Desempenho $desempenho -Pipeline $script:Pipeline
        if ($caminho -eq '/pipeline') {
            $filtro = [string]$req.QueryString['vendedor']
            if ($filtro -and -not $base.VendedorPorId.ContainsKey($filtro)) { $filtro = '' }
            Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaPipeline $base $visao $usuario $filtro)
        } else {
            $Ctx.Response.AddHeader('Content-Disposition', ('attachment; filename="pipeline_crm_{0:yyyy-MM-dd}.csv"' -f $visao.DataBaseCrm))
            Send-Resposta $Ctx 200 'text/csv; charset=utf-8' (New-CsvPipeline $visao) -ComBom
        }
        return
    }

    if ($caminho -eq '/origem-vendas') {
        Update-Pipeline
        if ($script:Pipeline.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $script:Pipeline.Erros $usuario 'origem'); return }
        $visao = Get-VisaoOrigem -Base $base -Pipeline $script:Pipeline -MesInicial $de -MesFinal $ate
        if ($visao.Erros.Count) { Send-Resposta $Ctx 500 'text/html; charset=utf-8' (New-PaginaErroImportacao $visao.Erros $usuario 'origem'); return }
        Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaOrigem $base $visao $usuario)
        return
    }

    switch ($caminho) {
        '/' {
            Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaPainel $base $desempenho $usuario $incluirDesligados)
        }
        '/vendedor' {
            $id = [string]$req.QueryString['id']
            $linha = $desempenho.Linhas | Where-Object { $_.Vendedor.Id -eq $id } | Select-Object -First 1
            if (-not $linha) { Send-Resposta $Ctx 404 'text/plain; charset=utf-8' 'Vendedor não encontrado'; return }
            $detalhe = Get-DetalheVendedor -Base $base -IdVendedor $id -MesInicial $de -MesFinal $ate
            Send-Resposta $Ctx 200 'text/html; charset=utf-8' (New-PaginaVendedor $base $desempenho $detalhe $linha $usuario $incluirDesligados)
        }
        '/conferencia.csv' {
            $disposicao = 'attachment; filename="conferencia_2026_{0:00}-{1:00}.csv"' -f $de, $ate
            $Ctx.Response.AddHeader('Content-Disposition', $disposicao)
            Send-Resposta $Ctx 200 'text/csv; charset=utf-8' (New-CsvConferencia $desempenho) -ComBom
        }
        default { Send-Resposta $Ctx 404 'text/plain; charset=utf-8' 'Página não encontrada' }
    }
}

Update-Base
$ouvinte = New-Object Net.HttpListener
# Só a própria máquina, mas pelos três nomes: sem 127.0.0.1 e [::1] o http.sys
# responde "400 Invalid Hostname" a quem digitar o IP no navegador.
# Fora do Windows (pwsh no macOS/Linux) o HttpListener gerenciado não aceita [::1].
$nomesLocais = if ($PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows) { @('localhost', '127.0.0.1', '[::1]') } else { @('localhost', '127.0.0.1') }
foreach ($host_ in $nomesLocais) { $ouvinte.Prefixes.Add("http://${host_}:$Porta/") }
try {
    $ouvinte.Start()
} catch {
    Write-Host "Não foi possível abrir a porta ${Porta}: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    Write-Host "Provavelmente o painel (ou outro programa) já está rodando nessa porta." -ForegroundColor Yellow
    Write-Host "Abra http://localhost:$Porta/ no navegador, ou suba em outra porta: -Porta 9000" -ForegroundColor Yellow
    exit 1
}
Write-Host "Painel no ar: http://localhost:$Porta/   (Ctrl+C para parar)" -ForegroundColor Green
Write-Host "Dados: $DirDados"
try {
    while ($ouvinte.IsListening) {
        $tarefa = $ouvinte.GetContextAsync()
        while (-not $tarefa.AsyncWaitHandle.WaitOne(500)) { }
        $ctx = $tarefa.GetAwaiter().GetResult()
        try {
            Invoke-Requisicao $ctx
        } catch {
            Write-Host "Erro em $($ctx.Request.HttpMethod) $($ctx.Request.Url.AbsolutePath): $($_.Exception.Message)" -ForegroundColor Red
            try { Send-Resposta $ctx 500 'text/plain; charset=utf-8' 'Erro interno. Veja o console do servidor.' } catch { }
        }
    }
} finally {
    $ouvinte.Stop()
    $ouvinte.Close()
}
