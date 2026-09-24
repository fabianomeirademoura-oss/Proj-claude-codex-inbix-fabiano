# Cria usuarios.json com uma conta por vendedor ATIVO (login = e-mail), uma por gerente
# comercial de filial (perfil "gerente", REGRAS_NEGOCIO.md §11.2) e uma conta "diretoria".
# Gera senhas aleatórias; o servidor só conhece o hash (usuarios.json).
# Os logins e senhas em texto vão para o arquivo .env na raiz do projeto — nunca
# para a tela — para não aparecerem em terminal, logs ou histórico.
#
# Uso: criar_usuarios.cmd             (não sobrescreve usuarios.json existente)
#      criar_usuarios.cmd -Completar  (só acrescenta as contas que faltam; as existentes e suas senhas não mudam)
#      criar_usuarios.cmd -Recriar    (gera tudo de novo e invalida as senhas antigas)

param(
    [string]$DirDados,
    [string]$ArquivoEnv,
    [switch]$Recriar,
    [switch]$Completar
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

if ($Recriar -and $Completar) { Write-Host 'Use -Recriar ou -Completar, não os dois.' -ForegroundColor Red; exit 1 }
$existentes = @()
if (Test-Path -LiteralPath $arqUsuarios) {
    if ($Completar) { $existentes = @(Read-Usuarios $arqUsuarios) }
    elseif (-not $Recriar) {
        Write-Host "usuarios.json já existe. Use -Completar para acrescentar só as contas que faltam, ou -Recriar para gerar novas senhas para todos." -ForegroundColor Yellow
        exit 1
    }
} elseif ($Completar) { Write-Host 'usuarios.json ainda não existe: rode sem -Completar.' -ForegroundColor Yellow; exit 1 }

# Cadastro em vigor: com as alterações confirmadas (§11.4), para que um vendedor reativado ganhe conta.
$base = Import-BaseComercial -DirDados $DirDados -DirImportacoes (Join-Path $DirDados 'importacoes')
if ($base.Erros.Count) { $base.Erros | ForEach-Object { Write-Host $_ -ForegroundColor Red }; exit 1 }

$usuarios = New-Object Collections.Generic.List[object]
foreach ($u in $existentes) { $usuarios.Add($u) }
$linhasEnv = New-Object Collections.Generic.List[string]
if ($Completar) {
    $linhasEnv.Add('')
    $linhasEnv.Add("# Contas acrescentadas em $((Get-Date).ToString('dd/MM/yyyy HH:mm')) (criar_usuarios -Completar)")
} else {
    $linhasEnv.Add('# Acessos do painel Horizonte Máquinas — NÃO versionar, NÃO enviar por chat/e-mail.')
    $linhasEnv.Add("# Gerado em $((Get-Date).ToString('dd/MM/yyyy HH:mm')). Para trocar todas as senhas: app\criar_usuarios.cmd -Recriar")
    $linhasEnv.Add('# O servidor não lê este arquivo: ele guarda só o hash das senhas em app\usuarios.json.')
    $linhasEnv.Add('PAINEL_URL=http://localhost:8080')
}

$contas = @([pscustomobject]@{ chave = 'DIRETORIA'; login = 'diretoria'; nome = 'Diretoria comercial'; perfil = 'diretoria'; idVendedor = $null; idFilial = $null })
# §11.2: um gerente por filial, com o nome da aba Filiais; o login vem do nome da filial (gerente.cascavel).
foreach ($f in (Read-XlsxSheet (Join-Path $DirDados 'vendedores.xlsx') 'Filiais')) {
    $semAcento = -join ($f.Filial.Normalize([Text.NormalizationForm]::FormD).ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
    $contas += [pscustomobject]@{ chave = "GERENTE_$($f.'ID Filial')"; login = 'gerente.' + ($semAcento.ToLowerInvariant() -replace '[^a-z0-9]', ''); nome = $f.'Gerente Comercial'; perfil = 'gerente'; idVendedor = $null; idFilial = $f.'ID Filial' }
}
$contas += foreach ($v in ($base.Vendedores | Where-Object Ativo)) {
    [pscustomobject]@{ chave = $v.Id; login = $v.Email.Trim().ToLowerInvariant(); nome = $v.Nome; perfil = 'vendedor'; idVendedor = $v.Id; idFilial = $null }
}
$loginsExistentes = @($existentes | ForEach-Object { ([string]$_.login).ToLowerInvariant() })
$novas = 0

foreach ($c in $contas) {
    if ($c.login -in $loginsExistentes) { continue }
    $novas++
    $senha = New-SenhaAleatoria 12
    $registro = New-RegistroSenha $senha
    $usuarios.Add([ordered]@{
            login      = $c.login
            nome       = $c.nome
            perfil     = $c.perfil
            idVendedor = $c.idVendedor
            idFilial   = $c.idFilial
            iteracoes  = $registro.iteracoes
            sal        = $registro.sal
            hash       = $registro.hash
        })
    $linhasEnv.Add('')
    $linhasEnv.Add("# $($c.nome)")
    $linhasEnv.Add("$($c.chave)_LOGIN=$($c.login)")
    $linhasEnv.Add("$($c.chave)_SENHA=$senha")
}

if ($Completar -and -not $novas) { Write-Host 'Nenhuma conta faltando: usuarios.json e .env não foram alterados.' -ForegroundColor Green; exit 0 }
$json = ConvertTo-Json -InputObject $usuarios.ToArray() -Depth 3
[IO.File]::WriteAllText($arqUsuarios, $json, (New-Object Text.UTF8Encoding($false)))
if ($Completar) { [IO.File]::AppendAllLines($ArquivoEnv, [string[]]$linhasEnv.ToArray(), (New-Object Text.UTF8Encoding($false))) }
else { [IO.File]::WriteAllLines($ArquivoEnv, [string[]]$linhasEnv.ToArray(), (New-Object Text.UTF8Encoding($false))) }
if ($Completar) { Write-Host "$novas conta(s) acrescentada(s); as existentes não mudaram." -ForegroundColor Green }

# O arquivo antigo de senhas (versões anteriores deste script) ficou inválido: remove.
$antigo = Join-Path $PSScriptRoot 'senhas_iniciais.txt'
if (Test-Path -LiteralPath $antigo) { Remove-Item -LiteralPath $antigo -Force; Write-Host 'senhas_iniciais.txt antigo removido (senhas nele não valem mais).' }

Write-Host "$($usuarios.Count) $(if ($Completar) { 'contas no total' } else { 'usuários criados' }). Hashes em $arqUsuarios" -ForegroundColor Green
Write-Host "Logins e senhas em $ArquivoEnv (as senhas não são exibidas na tela)." -ForegroundColor Yellow
Write-Host "Vendedores inativos não recebem conta (REGRAS_NEGOCIO.md §2.2)."
