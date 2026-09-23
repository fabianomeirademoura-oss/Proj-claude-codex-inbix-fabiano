# Autenticação: senhas com PBKDF2-SHA256, sessões em memória, bloqueio por tentativas.

$script:IteracoesSenha = 120000
$script:DuracaoSessao = [TimeSpan]::FromHours(8)
$script:MaxFalhas = 5
$script:DuracaoBloqueio = [TimeSpan]::FromMinutes(5)
$script:Sessoes = New-Object Collections.Hashtable ([StringComparer]::Ordinal)   # token diferencia maiúsculas
$script:Falhas = @{}

function Get-BytesAleatorios([int]$Quantidade) {
    $bytes = New-Object byte[] $Quantidade
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return , $bytes
}

function New-SenhaAleatoria([int]$Tamanho = 12) {
    # Sem caracteres ambíguos (0/O, 1/l/I). Rejeição evita viés de módulo.
    $alfabeto = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
    $limite = 256 - (256 % $alfabeto.Length)
    $saida = New-Object Text.StringBuilder
    while ($saida.Length -lt $Tamanho) {
        foreach ($b in (Get-BytesAleatorios 32)) {
            if ($b -lt $limite -and $saida.Length -lt $Tamanho) { [void]$saida.Append($alfabeto[$b % $alfabeto.Length]) }
        }
    }
    return $saida.ToString()
}

function Get-HashPbkdf2([string]$Senha, [byte[]]$Sal, [int]$Iteracoes) {
    $kdf = New-Object Security.Cryptography.Rfc2898DeriveBytes($Senha, $Sal, $Iteracoes, [Security.Cryptography.HashAlgorithmName]::SHA256)
    try { return , $kdf.GetBytes(32) } finally { $kdf.Dispose() }
}

function New-RegistroSenha([string]$Senha) {
    $sal = Get-BytesAleatorios 16
    return [ordered]@{
        iteracoes = $script:IteracoesSenha
        sal       = [Convert]::ToBase64String($sal)
        hash      = [Convert]::ToBase64String((Get-HashPbkdf2 $Senha $sal $script:IteracoesSenha))
    }
}

function Test-IgualTempoConstante([byte[]]$A, [byte[]]$B) {
    if ($A.Length -ne $B.Length) { return $false }
    $dif = 0
    for ($i = 0; $i -lt $A.Length; $i++) { $dif = $dif -bor ($A[$i] -bxor $B[$i]) }
    return $dif -eq 0
}

function Test-Senha($Usuario, [string]$Senha) {
    $sal = [Convert]::FromBase64String($Usuario.sal)
    $esperado = [Convert]::FromBase64String($Usuario.hash)
    return Test-IgualTempoConstante (Get-HashPbkdf2 $Senha $sal ([int]$Usuario.iteracoes)) $esperado
}

function Read-Usuarios([string]$Arquivo) {
    $texto = [IO.File]::ReadAllText($Arquivo, [Text.Encoding]::UTF8)
    # No PowerShell 5.1 o ConvertFrom-Json devolve o array inteiro como um item só.
    $lista = ConvertFrom-Json $texto
    foreach ($u in $lista) { $u }
}

function Test-Bloqueado([string]$Login) {
    $f = $script:Falhas[$Login]
    return ($f -and $f.BloqueadoAte -and $f.BloqueadoAte -gt (Get-Date))
}

function Register-Falha([string]$Login) {
    $f = $script:Falhas[$Login]
    if (-not $f -or ($f.BloqueadoAte -and $f.BloqueadoAte -le (Get-Date))) { $f = @{ Quantidade = 0; BloqueadoAte = $null } }
    $f.Quantidade++
    if ($f.Quantidade -ge $script:MaxFalhas) { $f.BloqueadoAte = (Get-Date).Add($script:DuracaoBloqueio) }
    $script:Falhas[$Login] = $f
}

function New-Sessao($Usuario) {
    $token = [Convert]::ToBase64String((Get-BytesAleatorios 32)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $script:Sessoes[$token] = @{ Usuario = $Usuario; ExpiraEm = (Get-Date).Add($script:DuracaoSessao) }
    $script:Falhas.Remove($Usuario.login)
    return $token
}

function Get-Sessao([string]$Token) {
    if (-not $Token) { return $null }
    $s = $script:Sessoes[$Token]
    if (-not $s) { return $null }
    if ($s.ExpiraEm -le (Get-Date)) { $script:Sessoes.Remove($Token); return $null }
    return $s
}

function Remove-Sessao([string]$Token) {
    if ($Token) { $script:Sessoes.Remove($Token) }
}
