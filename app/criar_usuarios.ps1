# Cria usuarios.json com uma conta por vendedor ATIVO (login = e-mail) e uma conta
# "diretoria". Gera senhas aleatórias; o servidor só conhece o hash (usuarios.json).
# Os logins e senhas em texto vão para o arquivo .env na raiz do projeto — nunca
# para a tela — para não aparecerem em terminal, logs ou histórico.
#
# Uso: criar_usuarios.cmd            (não sobrescreve usuarios.json existente)
#      criar_usuarios.cmd -Recriar   (gera tudo de novo e invalida as senhas antigas)

param(
    [string]$DirDados,
    [string]$ArquivoEnv,
    [switch]$Recriar
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\Xlsx.ps1')
. (Join-Path $PSScriptRoot 'lib\Regras.ps1')
. (Join-Path $PSScriptRoot 'lib\Auth.ps1')

if (-not $DirDados) {
    $DirDados = if ($env:HORIZONTE_DADOS) { $env:HORIZONTE_DADOS } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'dados' }
}
$arqUsuarios = Join-Path $PSScriptRoot 'usuarios.json'
if (-not $ArquivoEnv) { $ArquivoEnv = Join-Path (Split-Path $PSScriptRoot -Parent) '.env' }

if ((Test-Path -LiteralPath $arqUsuarios) -and -not $Recriar) {
    Write-Host "usuarios.json já existe. Use -Recriar para gerar novas senhas para todos." -ForegroundColor Yellow
    exit 1
}

$base = Import-BaseComercial -DirDados $DirDados
if ($base.Erros.Count) { $base.Erros | ForEach-Object { Write-Host $_ -ForegroundColor Red }; exit 1 }

$usuarios = New-Object Collections.Generic.List[object]
$linhasEnv = New-Object Collections.Generic.List[string]
$linhasEnv.Add('# Acessos do painel Horizonte Máquinas — NÃO versionar, NÃO enviar por chat/e-mail.')
$linhasEnv.Add("# Gerado em $((Get-Date).ToString('dd/MM/yyyy HH:mm')). Para trocar todas as senhas: app\criar_usuarios.cmd -Recriar")
$linhasEnv.Add('# O servidor não lê este arquivo: ele guarda só o hash das senhas em app\usuarios.json.')
$linhasEnv.Add('PAINEL_URL=http://localhost:8080')

$contas = @([pscustomobject]@{ chave = 'DIRETORIA'; login = 'diretoria'; nome = 'Diretoria comercial'; perfil = 'diretoria'; idVendedor = $null })
$contas += foreach ($v in ($base.Vendedores | Where-Object Ativo)) {
    [pscustomobject]@{ chave = $v.Id; login = $v.Email.Trim().ToLowerInvariant(); nome = $v.Nome; perfil = 'vendedor'; idVendedor = $v.Id }
}

foreach ($c in $contas) {
    $senha = New-SenhaAleatoria 12
    $registro = New-RegistroSenha $senha
    $usuarios.Add([ordered]@{
            login      = $c.login
            nome       = $c.nome
            perfil     = $c.perfil
            idVendedor = $c.idVendedor
            iteracoes  = $registro.iteracoes
            sal        = $registro.sal
            hash       = $registro.hash
        })
    $linhasEnv.Add('')
    $linhasEnv.Add("# $($c.nome)")
    $linhasEnv.Add("$($c.chave)_LOGIN=$($c.login)")
    $linhasEnv.Add("$($c.chave)_SENHA=$senha")
}

$json = ConvertTo-Json -InputObject $usuarios.ToArray() -Depth 3
[IO.File]::WriteAllText($arqUsuarios, $json, (New-Object Text.UTF8Encoding($false)))
[IO.File]::WriteAllLines($ArquivoEnv, [string[]]$linhasEnv.ToArray(), (New-Object Text.UTF8Encoding($false)))

# O arquivo antigo de senhas (versões anteriores deste script) ficou inválido: remove.
$antigo = Join-Path $PSScriptRoot 'senhas_iniciais.txt'
if (Test-Path -LiteralPath $antigo) { Remove-Item -LiteralPath $antigo -Force; Write-Host 'senhas_iniciais.txt antigo removido (senhas nele não valem mais).' }

Write-Host "$($usuarios.Count) usuários criados. Hashes em $arqUsuarios" -ForegroundColor Green
Write-Host "Logins e senhas em $ArquivoEnv (as senhas não são exibidas na tela)." -ForegroundColor Yellow
Write-Host "Vendedores inativos não recebem conta (REGRAS_NEGOCIO.md §2.2)."
