# Paridade entre as duas versões do painel: sobe a PowerShell (app/servidor.ps1, a referência)
# e a Node (web/servidor.js, a que roda na Vercel) com as mesmas contas temporárias e os mesmos
# dados, faz as mesmas requisições nas duas e confere status, tipo e corpo byte a byte.
# Como as conferências (testes/conferir*.ps1) validam a versão PowerShell contra as planilhas,
# corpo idêntico quer dizer que a versão Node mostra exatamente os mesmos números.
# Uso: pwsh -NoProfile -File testes/paridade_node.ps1 [-DirDados caminho] [-SalvarDiferencasEm pasta]

param([string]$DirDados, [int]$PortaPs = 8795, [int]$PortaNode = 8796, [string]$SalvarDiferencasEm)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$raiz = Split-Path $PSScriptRoot -Parent
if (-not $DirDados) { $DirDados = Join-Path $raiz 'dados' }
$DirDados = (Resolve-Path $DirDados).Path
. (Join-Path $raiz 'app/lib/Xlsx.ps1')
. (Join-Path $raiz 'app/lib/Regras.ps1')
. (Join-Path $raiz 'app/lib/RegrasPipeline.ps1')
. (Join-Path $raiz 'app/lib/RegrasImportacao.ps1')
. (Join-Path $raiz 'app/lib/Auth.ps1')

$falhas = 0; $total = 0
function Confere([string]$Descricao, [bool]$Ok, [string]$Detalhe) {
    $script:total++
    if ($Ok) { return }
    $script:falhas++
    Write-Host "  FALHA $Descricao" -ForegroundColor Red
    if ($Detalhe) { Write-Host "        $Detalhe" -ForegroundColor DarkYellow }
}

function Invoke-Http([int]$Porta, [string]$Metodo, [string]$Caminho, [string]$Corpo, [string]$Cookie) {
    $req = [Net.HttpWebRequest]::Create("http://localhost:$Porta$Caminho")
    $req.Method = $Metodo
    $req.AllowAutoRedirect = $false
    if ($Cookie) { $req.Headers.Add('Cookie', $Cookie) }
    if ($Corpo) {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Corpo)
        $req.ContentType = 'application/x-www-form-urlencoded'
        $req.ContentLength = $bytes.Length
        $s = $req.GetRequestStream(); $s.Write($bytes, 0, $bytes.Length); $s.Close()
    } elseif ($Metodo -eq 'POST') { $req.ContentLength = 0 }
    try { $resp = $req.GetResponse() } catch [Net.WebException] { $resp = $_.Exception.Response; if (-not $resp) { throw } }
    $ms = New-Object IO.MemoryStream
    $resp.GetResponseStream().CopyTo($ms)
    $r = [pscustomobject]@{
        Status = [int]$resp.StatusCode; Tipo = $resp.ContentType; Location = $resp.Headers['Location']
        Disposicao = $resp.Headers['Content-Disposition']; SetCookie = $resp.Headers['Set-Cookie']; Bytes = $ms.ToArray()
    }
    $resp.Close()
    return $r
}
function Form([hashtable]$h) { ($h.GetEnumerator() | ForEach-Object { "$($_.Key)=$([Uri]::EscapeDataString($_.Value))" }) -join '&' }

function Compare-Resposta([string]$Rotulo, $A, $B) {
    Confere "$Rotulo — status ($($A.Status) × $($B.Status))" ($A.Status -eq $B.Status)
    Confere "$Rotulo — Content-Type ($($A.Tipo) × $($B.Tipo))" ($A.Tipo -eq $B.Tipo)
    Confere "$Rotulo — Content-Disposition" ($A.Disposicao -eq $B.Disposicao) "$($A.Disposicao) × $($B.Disposicao)"
    if ($A.Status -eq 303) { Confere "$Rotulo — Location" ($A.Location -eq $B.Location -or $A.Location -eq "http://localhost:$PortaPs$($B.Location)") "$($A.Location) × $($B.Location)" }
    $igual = [Linq.Enumerable]::SequenceEqual([byte[]]$A.Bytes, [byte[]]$B.Bytes)
    $detalhe = ''
    if (-not $igual) {
        $ta = [Text.Encoding]::UTF8.GetString($A.Bytes); $tb = [Text.Encoding]::UTF8.GetString($B.Bytes)
        $i = 0; while ($i -lt $ta.Length -and $i -lt $tb.Length -and $ta[$i] -eq $tb[$i]) { $i++ }
        $ini = [math]::Max(0, $i - 60)
        $detalhe = "1ª diferença no caractere $i. PowerShell: «$($ta.Substring($ini, [math]::Min(140, $ta.Length - $ini)))» | Node: «$($tb.Substring($ini, [math]::Min(140, $tb.Length - $ini)))»"
        if ($SalvarDiferencasEm) {
            New-Item -ItemType Directory -Force -Path $SalvarDiferencasEm | Out-Null
            $nome = ($Rotulo -replace '[^\w.-]+', '_').Trim('_')
            [IO.File]::WriteAllBytes((Join-Path $SalvarDiferencasEm "$nome.ps.txt"), $A.Bytes)
            [IO.File]::WriteAllBytes((Join-Path $SalvarDiferencasEm "$nome.node.txt"), $B.Bytes)
        }
    }
    Confere "$Rotulo — corpo idêntico ($($A.Bytes.Length) × $($B.Bytes.Length) bytes)" $igual $detalhe
}

# ---------- contas temporárias ----------
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte_paridade_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$base = Import-BaseComercial -DirDados $DirDados -DirImportacoes (Get-DirImportacoes $DirDados)
if ($base.Erros.Count) { throw "A base tem erros: $($base.Erros[0])" }
$vendedorAtivo = $base.Vendedores | Where-Object Ativo | Select-Object -First 1
$senhaDir = New-SenhaAleatoria; $senhaVend = New-SenhaAleatoria
$publica = (Get-Content -Raw (Join-Path $raiz 'web/contas_publicas.json') | ConvertFrom-Json)[0]
$contas = @(
    [ordered]@{ login = 'teste.diretoria'; nome = 'Teste Diretoria'; perfil = 'diretoria'; idVendedor = $null; senha = $senhaDir }
    [ordered]@{ login = $vendedorAtivo.Email.Trim().ToLowerInvariant(); nome = $vendedorAtivo.Nome; perfil = 'vendedor'; idVendedor = $vendedorAtivo.Id; senha = $senhaVend }
)
$registros = foreach ($c in $contas) { $r = New-RegistroSenha $c.senha; [ordered]@{ login = $c.login; nome = $c.nome; perfil = $c.perfil; idVendedor = $c.idVendedor; iteracoes = $r.iteracoes; sal = $r.sal; hash = $r.hash } }
$arqNode = Join-Path $tmp 'usuarios_node.json'   # a Node já tem a conta teste em web/contas_publicas.json
[IO.File]::WriteAllText($arqNode, (ConvertTo-Json -InputObject @($registros) -Depth 3), (New-Object Text.UTF8Encoding($false)))
$arqPs = Join-Path $tmp 'usuarios_ps.json'       # a PowerShell recebe a mesma conta teste (mesmo hash)
$todas = @($registros) + @([ordered]@{ login = $publica.login; nome = $publica.nome; perfil = $publica.perfil; idVendedor = $null; iteracoes = $publica.iteracoes; sal = $publica.sal; hash = $publica.hash })
[IO.File]::WriteAllText($arqPs, (ConvertTo-Json -InputObject $todas -Depth 3), (New-Object Text.UTF8Encoding($false)))

# ---------- sobe as duas versões ----------
$logPs = Join-Path $tmp 'ps.log'; $logNode = Join-Path $tmp 'node.log'
$procPs = Start-Process pwsh -ArgumentList @('-NoProfile', '-File', (Join-Path $raiz 'app/servidor.ps1'), '-Porta', $PortaPs, '-DirDados', $DirDados, '-ArquivoUsuarios', $arqPs) -PassThru -RedirectStandardOutput $logPs -RedirectStandardError "$logPs.err"
$env:HORIZONTE_DADOS = $DirDados; $env:HORIZONTE_USUARIOS_ARQUIVO = $arqNode; $env:HORIZONTE_SEGREDO = [Guid]::NewGuid().ToString('N')
$procNode = Start-Process node -ArgumentList @((Join-Path $raiz 'web/servidor.js'), $PortaNode) -PassThru -RedirectStandardOutput $logNode -RedirectStandardError "$logNode.err"
Remove-Item Env:HORIZONTE_USUARIOS_ARQUIVO, Env:HORIZONTE_SEGREDO, Env:HORIZONTE_DADOS

try {
    foreach ($porta in @($PortaPs, $PortaNode)) {
        $ok = $false
        for ($i = 0; $i -lt 120 -and -not $ok; $i++) {
            try { [void](Invoke-Http $porta 'GET' '/entrar'); $ok = $true } catch { Start-Sleep -Milliseconds 250 }
        }
        if (-not $ok) { throw "Servidor na porta $porta não respondeu. Logs em $tmp" }
    }
    Write-Host "Dados: $DirDados" -ForegroundColor Cyan

    # ---------- sem login ----------
    Compare-Resposta 'sem login GET /' (Invoke-Http $PortaPs 'GET' '/') (Invoke-Http $PortaNode 'GET' '/')
    Compare-Resposta 'sem login GET /entrar' (Invoke-Http $PortaPs 'GET' '/entrar') (Invoke-Http $PortaNode 'GET' '/entrar')
    Compare-Resposta 'sem login GET /estilo.css' (Invoke-Http $PortaPs 'GET' '/estilo.css') (Invoke-Http $PortaNode 'GET' '/estilo.css')
    $f = Form @{ login = 'teste'; senha = 'errada' }
    Compare-Resposta 'login com senha errada' (Invoke-Http $PortaPs 'POST' '/entrar' $f) (Invoke-Http $PortaNode 'POST' '/entrar' $f)
    $f = Form @{ login = ' Alguém <x> '; senha = '' }
    Compare-Resposta 'login sem senha' (Invoke-Http $PortaPs 'POST' '/entrar' $f) (Invoke-Http $PortaNode 'POST' '/entrar' $f)

    # ---------- recortes ----------
    $pipe = Import-BasePipeline -DirDados $DirDados -Comercial $base
    $ids = @($base.Vendedores | ForEach-Object Id)
    $filiais = @($pipe.Filiais | ForEach-Object Id)
    $categorias = @($pipe.Oportunidades | ForEach-Object { $pipe.ProdutoPorId[$_.IdProduto].Categoria } | Sort-Object -Unique)
    $max = @(Get-MesesDisponiveis $base)[-1]
    $urls = New-Object Collections.Generic.List[string]
    foreach ($u in @('/', '/?desligados=1', "/?de=2&ate=5", "/?de=$max&ate=$max", '/?de=7&ate=3&desligados=1', '/?de=abc&ate=99', '/?de=0',
            '/conferencia.csv', '/conferencia.csv?de=3&ate=3', "/conferencia.csv?de=1&ate=$max",
            '/pipeline', '/pipeline?de=1&ate=4', "/pipeline?de=$max&ate=$max", '/pipeline?vendedor=ZZZ', '/pipeline.csv',
            '/funil', '/funil?medida=valor', '/funil?medida=VALOR', '/funil?vendedor=ZZZ&filial=ZZZ&categoria=ZZZ',
            '/estoque', '/estoque?dias=1', '/estoque?dias=30', '/estoque?dias=90', '/estoque?dias=365', '/estoque?dias=3650', '/estoque?dias=0', '/estoque?dias=abc', '/estoque?dias=3651', '/estoque?filial=ZZZ',
            '/estoque.csv', '/estoque.csv?dias=60', '/origem-vendas', '/origem-vendas?de=3&ate=5', "/origem-vendas?de=$max&ate=$max",
            '/vendedor?id=ZZZ', '/naoexiste', '/importar')) { $urls.Add($u) }
    foreach ($id in $ids) {
        $urls.Add("/vendedor?id=$id"); $urls.Add("/vendedor?id=$id&de=2&ate=$max&desligados=1")
        $urls.Add("/pipeline?vendedor=$id"); $urls.Add("/funil?vendedor=$id"); $urls.Add("/funil?vendedor=$id&medida=valor")
    }
    foreach ($fi in $filiais) {
        $urls.Add("/funil?filial=$fi"); $urls.Add("/estoque?filial=$fi"); $urls.Add("/estoque?dias=90&filial=$fi")
        foreach ($c in $categorias) { $urls.Add("/funil?filial=$fi&categoria=$([Uri]::EscapeDataString($c))&medida=valor") }
    }
    foreach ($c in $categorias) { $urls.Add("/funil?categoria=$([Uri]::EscapeDataString($c))") }

    # ---------- com login: conta teste (alunos), vendedor e diretoria ----------
    $logins = @(
        @{ Login = 'teste'; Senha = 'teste'; Rotulo = 'teste' }
        @{ Login = $contas[1].login; Senha = $senhaVend; Rotulo = $contas[1].idVendedor }
        @{ Login = 'teste.diretoria'; Senha = $senhaDir; Rotulo = 'diretoria' }
    )
    foreach ($l in $logins) {
        $f = Form @{ login = $l.Login; senha = $l.Senha }
        $ra = Invoke-Http $PortaPs 'POST' '/entrar' $f; $rb = Invoke-Http $PortaNode 'POST' '/entrar' $f
        Confere "login $($l.Rotulo) na PowerShell" ($ra.Status -eq 303 -and $ra.SetCookie -match 'sessao=([^;]+)')
        $cookiePs = "sessao=$($Matches[1])"
        Confere "login $($l.Rotulo) na Node" ($rb.Status -eq 303 -and $rb.SetCookie -match 'sessao=([^;]+)')
        $cookieNode = "sessao=$($Matches[1])"
        $n0 = $falhas
        foreach ($u in $urls) {
            if ($l.Rotulo -eq 'diretoria' -and $u -eq '/importar') { continue }   # importação só existe na versão local (§9)
            Compare-Resposta "[$($l.Rotulo)] GET $u" (Invoke-Http $PortaPs 'GET' $u '' $cookiePs) (Invoke-Http $PortaNode 'GET' $u '' $cookieNode)
        }
        Compare-Resposta "[$($l.Rotulo)] POST /" (Invoke-Http $PortaPs 'POST' '/' '' $cookiePs) (Invoke-Http $PortaNode 'POST' '/' '' $cookieNode)
        Write-Host ("  {0,-10} {1} URLs, {2}" -f $l.Rotulo, $urls.Count, $(if ($falhas -eq $n0) { 'todas idênticas' } else { "$($falhas - $n0) diferença(s)" })) -ForegroundColor $(if ($falhas -eq $n0) { 'Green' } else { 'Red' })
    }
    $rp = Invoke-Http $PortaNode 'GET' '/importar' '' $cookieNode
    Confere 'Node: diretoria vê aviso de importação indisponível (501)' ($rp.Status -eq 501)
    $rs = Invoke-Http $PortaNode 'POST' '/sair' '' $cookieNode
    Confere 'Node: sair apaga o cookie' ($rs.Status -eq 303 -and $rs.SetCookie -match 'Max-Age=0')
    Confere 'Node: cookie adulterado não entra' ((Invoke-Http $PortaNode 'GET' '/' '' ($cookieNode + 'x')).Status -eq 303)
} finally {
    foreach ($p in @($procPs, $procNode)) { if ($p -and -not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
    if ($falhas -eq 0) { Remove-Item -Recurse -Force $tmp } else { Write-Host "Logs dos servidores em $tmp" -ForegroundColor Yellow }
}

if ($falhas) { Write-Host "$falhas de $total verificações falharam." -ForegroundColor Red; exit 1 }
Write-Host "Todas as $total verificações de paridade bateram." -ForegroundColor Green
