# Teste de ponta a ponta: sobe o servidor numa porta de teste com contas temporárias
# (não usa usuarios.json real), faz as requisições HTTP e confere telas e permissões.
# Uso: powershell -ExecutionPolicy Bypass -File teste_e2e.ps1 [-SalvarPainelEm caminho.html]

param([int]$Porta = 8791, [string]$SalvarPainelEm)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
. (Join-Path $PSScriptRoot '..\app\lib\Auth.ps1')
. (Join-Path $PSScriptRoot 'XlsxTeste.ps1')

$falhas = 0; $total = 0
function Confere([string]$Descricao, [bool]$Ok) {
    $script:total++
    if ($Ok) { Write-Host "  OK    $Descricao" -ForegroundColor Green } else { $script:falhas++; Write-Host "  FALHA $Descricao" -ForegroundColor Red }
}

function Invoke-Http([string]$Metodo, [string]$Caminho, [string]$Corpo, [string]$Cookie) {
    $req = [Net.HttpWebRequest]::Create("http://localhost:$Porta$Caminho")
    $req.Method = $Metodo
    $req.AllowAutoRedirect = $false
    if ($Cookie) { $req.Headers.Add('Cookie', $Cookie) }
    if ($Corpo) {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Corpo)
        $req.ContentType = 'application/x-www-form-urlencoded'
        $req.ContentLength = $bytes.Length
        $s = $req.GetRequestStream(); $s.Write($bytes, 0, $bytes.Length); $s.Close()
    } elseif ($Metodo -eq 'POST') {
        $req.ContentLength = 0   # como o navegador faz num formulário sem campos
    }
    try { $resp = $req.GetResponse() } catch [Net.WebException] { $resp = $_.Exception.Response; if (-not $resp) { throw } }
    $leitor = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
    $texto = $leitor.ReadToEnd(); $leitor.Dispose()
    # Corpo: HTML com entidades decodificadas (o .NET 4 codifica acentos como &#227;). Bruto: como veio.
    $r = [pscustomobject]@{ Status = [int]$resp.StatusCode; Location = $resp.Headers['Location']; SetCookie = $resp.Headers['Set-Cookie']; Corpo = [Net.WebUtility]::HtmlDecode($texto); Bruto = $texto; Csp = $resp.Headers['Content-Security-Policy'] }
    $resp.Close()
    return $r
}
function Form([hashtable]$h) { ($h.GetEnumerator() | ForEach-Object { "$($_.Key)=$([Uri]::EscapeDataString($_.Value))" }) -join '&' }
function Invoke-Upload([string]$Caminho, [string]$Tipo, [string]$Arquivo, [string]$NomeArquivo, [string]$Cookie) {
    # multipart/form-data como o navegador envia o formulário de importação
    $fronteira = '----teste' + [Guid]::NewGuid().ToString('N')
    $ms = New-Object IO.MemoryStream
    $escreve = { param($t) $b = [Text.Encoding]::UTF8.GetBytes($t); $ms.Write($b, 0, $b.Length) }
    if ($Tipo) { & $escreve "--$fronteira`r`nContent-Disposition: form-data; name=""tipo""`r`n`r`n$Tipo`r`n" }
    if ($Arquivo) {
        & $escreve "--$fronteira`r`nContent-Disposition: form-data; name=""arquivo""; filename=""$NomeArquivo""`r`nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n"
        $b = [IO.File]::ReadAllBytes($Arquivo); $ms.Write($b, 0, $b.Length); & $escreve "`r`n"
    }
    & $escreve "--$fronteira--`r`n"
    $req = [Net.HttpWebRequest]::Create("http://localhost:$Porta$Caminho")
    $req.Method = 'POST'; $req.AllowAutoRedirect = $false; $req.Headers.Add('Cookie', $Cookie)
    $req.ContentType = "multipart/form-data; boundary=$fronteira"
    $corpo = $ms.ToArray(); $req.ContentLength = $corpo.Length
    $st = $req.GetRequestStream(); $st.Write($corpo, 0, $corpo.Length); $st.Close()
    try { $resp = $req.GetResponse() } catch [Net.WebException] { $resp = $_.Exception.Response; if (-not $resp) { throw } }
    $leitor = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
    $texto = $leitor.ReadToEnd(); $leitor.Dispose()
    $r = [pscustomobject]@{ Status = [int]$resp.StatusCode; Location = $resp.Headers['Location']; Corpo = [Net.WebUtility]::HtmlDecode($texto); Bruto = $texto }
    $resp.Close()
    return $r
}
function Get-TokenImportacao($Resposta) { if ($Resposta.Bruto -match 'name="token" value="([0-9a-f]{32})"') { return $Matches[1] } }
function Entrar([string]$Login, [string]$Senha) { Invoke-Http 'POST' '/entrar' (Form @{ login = $Login; senha = $Senha }) }

$inicioTeste = Get-Date
# Contas temporárias só para o teste
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("horizonte_e2e_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$senhaDir = New-SenhaAleatoria; $senhaMar = New-SenhaAleatoria; $senhaRic = New-SenhaAleatoria
$contas = @(
    @{ login = 'teste.diretoria'; nome = 'Teste Diretoria'; perfil = 'diretoria'; idVendedor = $null; senha = $senhaDir }
    @{ login = 'mariana.ribeiro@horizontemaquinas.com.br'; nome = 'Mariana Costa Ribeiro'; perfil = 'vendedor'; idVendedor = 'V003'; senha = $senhaMar }
    @{ login = 'ricardo.pereira@horizontemaquinas.com.br'; nome = 'Ricardo Alves Pereira'; perfil = 'vendedor'; idVendedor = 'V011'; senha = $senhaRic }
)
$usuarios = foreach ($c in $contas) { $r = New-RegistroSenha $c.senha; [ordered]@{ login = $c.login; nome = $c.nome; perfil = $c.perfil; idVendedor = $c.idVendedor; iteracoes = $r.iteracoes; sal = $r.sal; hash = $r.hash } }
$arqUsuarios = Join-Path $tmp 'usuarios.json'
[IO.File]::WriteAllText($arqUsuarios, (ConvertTo-Json -InputObject @($usuarios) -Depth 3), (New-Object Text.UTF8Encoding($false)))

# Cópia dos dados só para o teste: a importação grava em dados/importacoes/ e não pode tocar nos dados reais.
$dadosReais = Join-Path (Split-Path $PSScriptRoot -Parent) 'dados'
$dadosTeste = Join-Path $tmp 'dados'
New-Item -ItemType Directory -Path (Join-Path $dadosTeste 'atualizacoes') -Force | Out-Null
Get-ChildItem -LiteralPath $dadosReais -File -Filter '*.xlsx' | Copy-Item -Destination $dadosTeste
Get-ChildItem -LiteralPath (Join-Path $dadosReais 'atualizacoes') -File -Filter '*.xlsx' | Copy-Item -Destination (Join-Path $dadosTeste 'atualizacoes')

# Windows PowerShell no Windows; pwsh onde ele não existe (macOS/Linux).
$exe = if (Get-Command powershell -ErrorAction SilentlyContinue) { 'powershell' } else { 'pwsh' }
$janela = if ($exe -eq 'powershell') { @{ WindowStyle = 'Hidden' } } else { @{} }   # -WindowStyle só existe no Windows
$proc = Start-Process $exe -PassThru @janela -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $PSScriptRoot '..\app\servidor.ps1')`"", '-Porta', $Porta, '-ArquivoUsuarios', "`"$arqUsuarios`"", '-DirDados', "`"$dadosTeste`"")
try {
    $pronto = $false
    for ($i = 0; $i -lt 60 -and -not $pronto; $i++) {
        Start-Sleep -Milliseconds 500
        try { [void](Invoke-Http 'GET' '/estilo.css'); $pronto = $true } catch { }
    }
    if (-not $pronto) { throw 'Servidor não subiu' }

    Write-Host "`n== Acesso ==" -ForegroundColor Cyan
    $r = Invoke-Http 'GET' '/'
    Confere 'sem login, / redireciona para /entrar' ($r.Status -eq 303 -and $r.Location -eq '/entrar')
    Confere 'sem login, CSV também redireciona' ((Invoke-Http 'GET' '/conferencia.csv').Status -eq 303)
    $r = Entrar 'teste.diretoria' 'senha-errada'
    Confere 'senha errada: 401 e mensagem genérica' ($r.Status -eq 401 -and $r.Corpo -match 'Usuário ou senha inválidos')
    $r = Entrar 'ninguem@x.com' 'qualquer'
    Confere 'usuário inexistente: mesma mensagem genérica' ($r.Status -eq 401 -and $r.Corpo -match 'Usuário ou senha inválidos')
    $r = Entrar 'ricardo.pereira@horizontemaquinas.com.br' $senhaRic
    Confere 'vendedor desligado (V011) com senha certa: 403, sem sessão' ($r.Status -eq 403 -and -not $r.SetCookie)
    $r = Entrar 'teste.diretoria' $senhaDir
    Confere 'diretoria: login ok, cookie HttpOnly e SameSite' ($r.Status -eq 303 -and $r.SetCookie -match 'HttpOnly' -and $r.SetCookie -match 'SameSite=Strict')
    $cookie = ($r.SetCookie -split ';')[0]

    Write-Host "`n== Painel jan–ago ==" -ForegroundColor Cyan
    $p = Invoke-Http 'GET' '/' '' $cookie
    Confere 'painel 200 com CSP' ($p.Status -eq 200 -and $p.Csp -match "default-src 'none'")
    Confere 'realizado R$ 45.295.119,16' ($p.Corpo -match 'R\$ 45\.295\.119,16')
    Confere 'meta R$ 50.720.000,00' ($p.Corpo -match 'R\$ 50\.720\.000,00')
    Confere 'atingimento 89,3%' ($p.Corpo -match '89,3%')
    Confere 'canceladas R$ 407.327,39 / 12' ($p.Corpo -match 'R\$ 407\.327,39' -and $p.Corpo -match '12 vendas com Status = Cancelada')
    Confere '723 vendas no cartão' ($p.Corpo -match '723 vendas')
    $rankings = ($p.Corpo -split 'Quadro de conferência')[0]
    Confere 'Ricardo fora dos rankings padrão' ($rankings -notmatch 'Ricardo Alves Pereira<')
    Confere 'Ricardo no quadro de conferência, marcado' ($p.Corpo -match 'Ricardo Alves Pereira</a> <span class="marca desligado">desligado em 30/04/2026')
    Confere 'Camila marcada como admissão em 07/2026' ($p.Corpo -match 'admissão em 07/2026')
    $fat = ($rankings -split 'Ranking por atingimento')[0]
    Confere 'faturamento: Mariana antes de João' ($fat.IndexOf('Mariana Costa Ribeiro') -lt $fat.IndexOf('João Pedro Almeida') -and $fat.IndexOf('Mariana Costa Ribeiro') -gt 0)
    Confere 'atingimento: Mariana 114,9%' ($rankings -match '114,9%')
    if ($SalvarPainelEm) { [IO.File]::WriteAllText($SalvarPainelEm, $p.Bruto, (New-Object Text.UTF8Encoding($false))) }

    Write-Host "`n== Filtros, detalhe e CSV ==" -ForegroundColor Cyan
    $p2 = Invoke-Http 'GET' '/?desligados=1' '' $cookie
    Confere 'com desligados: Ricardo entra no ranking' ((($p2.Corpo -split 'Quadro de conferência')[0]) -match 'Ricardo Alves Pereira<')
    $p3 = Invoke-Http 'GET' '/?de=7&ate=8' '' $cookie
    Confere 'jul–ago: Camila com 80,1%' ($p3.Corpo -match '80,1%')
    $p4 = Invoke-Http 'GET' '/?de=3&ate=5' '' $cookie
    Confere 'mar–mai: Camila sem meta no período' ($p4.Corpo -match 'Sem meta no período[^<]*Camila Rodrigues Teixeira')
    $v = Invoke-Http 'GET' '/vendedor?id=V011' '' $cookie
    Confere 'detalhe V011: 62,8% e canceladas riscadas' ($v.Status -eq 200 -and $v.Corpo -match '62,8%')
    $x = Invoke-Http 'GET' '/vendedor?id=%3Cscript%3E' '' $cookie
    Confere 'id inválido: 404 sem refletir a entrada' ($x.Status -eq 404 -and $x.Bruto -notmatch '<script>')
    $csv = Invoke-Http 'GET' '/conferencia.csv' '' $cookie
    Confere 'CSV: total 45295119,16 e 13 linhas + canceladas' ($csv.Corpo -match 'TOTAL;Empresa;;;;50720000,00;45295119,16;0,00;45295119,16;89,30;723' -and ($csv.Corpo.Trim() -split "`n").Count -eq 15)

    Write-Host "`n== Origem do faturamento ==" -ForegroundColor Cyan
    Confere 'origem exige login' ((Invoke-Http 'GET' '/origem-vendas').Status -eq 303)
    $origem = Invoke-Http 'GET' '/origem-vendas' '' $cookie
    Confere 'origem acessível pelo menu e com CSP' ($origem.Status -eq 200 -and $origem.Corpo -match 'href="/origem-vendas" class="ativo"' -and $origem.Csp -match "default-src 'none'")
    Confere 'CRM representa 33,8% e R$ 15.301.000,00' ($origem.Corpo -match '33,8% do faturamento' -and $origem.Corpo -match 'R\$ 15\.301\.000,00' -and $origem.Corpo -match '95 vendas vinculadas')
    Confere 'sem oportunidade: 628 vendas, 66,2%, R$ 29.994.119,16' ($origem.Corpo -match '628 vendas sem vínculo' -and $origem.Corpo -match '66,2%' -and $origem.Corpo -match 'R\$ 29\.994\.119,16')
    Confere 'todas as vendas ligadas ou explicitamente sem oportunidade' ([regex]::Matches($origem.Corpo,'VD-2026-\d{5}').Count -eq 735 -and $origem.Corpo -match 'VD-2026-00011[^\n]*OP-0011')
    Confere 'explica balcão e canceladas' ($origem.Corpo -match 'vendas de balcão' -and $origem.Corpo -match 'nem que sejam somente peças' -and $origem.Corpo -match '12 canceladas')
    $origemParcial = Invoke-Http 'GET' '/origem-vendas?de=7&ate=8' '' $cookie
    Confere 'origem acompanha período jul-ago' ($origemParcial.Status -eq 200 -and $origemParcial.Corpo -match '23,6% do faturamento' -and $origemParcial.Corpo -match 'R\$ 2\.871\.500,00' -and [regex]::Matches($origemParcial.Corpo,'VD-2026-\d{5}').Count -lt 735)

    Write-Host "`n== Estoque ==" -ForegroundColor Cyan
    Confere 'sem login, /estoque redireciona' ((Invoke-Http 'GET' '/estoque').Status -eq 303)
    $e = Invoke-Http 'GET' '/estoque' '' $cookie
    Confere 'estoque 200, link no menu marcado' ($e.Status -eq 200 -and $e.Corpo -match 'href="/estoque" class="ativo"')
    Confere 'mostra a data da foto 31/08/2026 e o arquivo' ($e.Corpo -match 'Foto do estoque de <strong>31/08/2026</strong> \(estoque_2026-08-31\.xlsx\)')
    Confere 'imobilizado R$ 18.420.192,72' ($e.Corpo -match 'R\$ 18\.420\.192,72')
    Confere 'parado padrão 180 dias: 7 linhas, R$ 1.693.610,48' ($e.Corpo -match 'Parado há 180\+ dias' -and $e.Corpo -match '7 linhas \(produto, filial\)' -and $e.Corpo -match 'R\$ 1\.693\.610,48')
    Confere 'definição de parado escrita na tela, com o prazo' ($e.Corpo -match 'Definição usada:</strong> Parado é a linha de estoque[^<]*há 180 dias ou mais[^<]*31/08/2026')
    Confere '8 produtos sem estoque, 28 abaixo do mínimo' ($e.Corpo -match '8 produtos' -and $e.Corpo -match '28 linhas')
    Confere 'P008 Cascavel parado com 2 oportunidades' ($e.Corpo -match 'Trator Vulcan T-260[^\n]*<td>Cascavel</td>[^\n]*<strong>2</strong>')
    $e2 = Invoke-Http 'GET' '/estoque?dias=90' '' $cookie
    Confere 'prazo 90 dias muda a lista e a frase, sem mexer no código' ($e2.Corpo -match 'há 90 dias ou mais' -and $e2.Corpo -match 'Parado há 90\+ dias' -and $e2.Corpo -match '21 linhas \(produto, filial\)')
    $e3 = Invoke-Http 'GET' '/estoque?dias=abc' '' $cookie
    Confere 'prazo inválido: avisa e usa 180' ($e3.Status -eq 200 -and $e3.Corpo -match 'Prazo inválido' -and $e3.Corpo -match 'há 180 dias ou mais')
    $e4 = Invoke-Http 'GET' '/estoque?filial=F01' '' $cookie
    Confere 'filtro Cascavel: 3 parados na lista' ($e4.Corpo -match '<p class="nota">3 linhas, ')
    $e5 = Invoke-Http 'GET' '/estoque?filial=%3Cscript%3E' '' $cookie
    Confere 'filial inválida é ignorada, sem refletir a entrada' ($e5.Status -eq 200 -and $e5.Bruto -notmatch '<script>')
    $ec = Invoke-Http 'GET' '/estoque.csv?dias=180' '' $cookie
    Confere 'CSV de estoque: 195 linhas + cabeçalho + total' (($ec.Corpo.Trim() -split "`n").Count -eq 197 -and $ec.Corpo -match 'TOTAL;;;;;[0-9]+;;;18420192,72;;;;7 linhas, 1693610,48;28 linhas;8 produtos;')

    Write-Host "`n== Funil ==" -ForegroundColor Cyan
    Confere 'sem login, /funil redireciona' ((Invoke-Http 'GET' '/funil').Status -eq 303)
    $fu = Invoke-Http 'GET' '/funil' '' $cookie
    Confere 'funil 200, link no menu marcado, desenho SVG' ($fu.Status -eq 200 -and $fu.Corpo -match 'href="/funil" class="ativo"' -and $fu.Bruto -match '<svg class="funil"')
    Confere 'funil sem filtro: 68 abertas, 29 atrasadas, 49,7%' ($fu.Corpo -match '<span class="valor">68</span>' -and $fu.Corpo -match '29 abertas' -and $fu.Corpo -match '49,7%')
    Confere 'funil: etapas 16/13/22/17 e ponderado R$ 5.605.495,00' ($fu.Corpo -match '16 abertas' -and $fu.Corpo -match '13 abertas' -and $fu.Corpo -match '22 abertas' -and $fu.Corpo -match '17 abertas' -and $fu.Corpo -match 'R\$ 5\.605\.495,00')
    Confere 'funil diz que o CRM guarda só a etapa atual' ($fu.Corpo -match 'O CRM guarda só a etapa atual')
    $fv = Invoke-Http 'GET' '/funil?vendedor=V011' '' $cookie
    Confere 'filtro vendedor desligado: 2 órfãs' ($fv.Corpo -match 'Ricardo Alves Pereira</strong>' -and $fv.Corpo -match '<span class="valor">2</span>' -and $fv.Corpo -match 'órfã')
    $fc = Invoke-Http 'GET' '/funil?filial=F03&categoria=Tratores&medida=valor' '' $cookie
    Confere 'filtro filial + categoria + medida por valor' ($fc.Status -eq 200 -and $fc.Corpo -match 'filial do vendedor <strong>Passo Fundo' -and $fc.Corpo -match 'categoria <strong>Tratores' -and $fc.Bruto -match 'value="valor" checked')
    $fx = Invoke-Http 'GET' '/funil?vendedor=%3Cscript%3E&categoria=%3Cb%3E' '' $cookie
    Confere 'filtros inválidos viram "sem filtro", sem refletir a entrada' ($fx.Status -eq 200 -and $fx.Corpo -match 'todo o CRM, sem filtro' -and $fx.Bruto -notmatch '<script>|<b>')

    Write-Host "`n== Importação (REGRAS_NEGOCIO.md §9) ==" -ForegroundColor Cyan
    # Por último antes da sessão: importar setembro muda os números das seções acima.
    $arqSet = Join-Path $dadosTeste 'atualizacoes\vendas_2026_setembro.xlsx'
    $arqEst19 = Join-Path $dadosTeste 'atualizacoes\estoque_2026-09-19.xlsx'
    Confere 'sem login, /importar redireciona' ((Invoke-Http 'GET' '/importar').Status -eq 303)
    $im = Invoke-Http 'GET' '/importar' '' $cookie
    Confere 'diretoria: tela de importação e link no menu' ($im.Status -eq 200 -and $im.Corpo -match 'href="/importar" class="ativo"' -and $im.Bruto -match 'enctype="multipart/form-data"')
    $rm = Entrar 'mariana.ribeiro@horizontemaquinas.com.br' $senhaMar
    $cookieVend = ($rm.SetCookie -split ';')[0]
    Confere 'vendedora: sem link Importar no menu' ((Invoke-Http 'GET' '/' '' $cookieVend).Bruto -notmatch 'href="/importar"')
    Confere 'vendedora: GET e POST de importação recusados (403)' ((Invoke-Http 'GET' '/importar' '' $cookieVend).Status -eq 403 -and (Invoke-Upload '/importar' 'vendas' $arqSet 'vendas_2026_setembro.xlsx' $cookieVend).Status -eq 403)
    $arqSemStatus = Join-Path $tmp 'sem_status.xlsx'
    New-XlsxTeste -Caminho $arqSemStatus -Aba 'Vendas' -Cabecalho @('ID Venda', 'Data', 'ID Vendedor', 'Vendedor', 'Filial', 'ID Cliente', 'Cliente', 'Cidade', 'UF', 'ID Produto', 'Produto', 'Categoria', 'Quantidade', 'Valor Unitário', 'Valor Total', 'Forma de Pagamento', 'ID Oportunidade') -Linhas @([pscustomobject]@{ 'ID Venda' = 'VD-2026-09000'; Data = '2026-09-02' })
    $u = Invoke-Upload '/importar' 'vendas' $arqSemStatus 'sem_status.xlsx' $cookie
    Confere 'coluna faltando: recusado (422) dizendo qual' ($u.Status -eq 422 -and $u.Corpo -match "Falta a coluna obrigatória: 'Status'" -and $u.Corpo -match 'Nada foi gravado')
    $u = Invoke-Upload '/importar' '' $arqSet 'vendas_2026_setembro.xlsx' $cookie
    Confere 'sem escolher o tipo: 400' ($u.Status -eq 400 -and $u.Corpo -match 'Escolha se o arquivo é de vendas ou de estoque')
    $u = Invoke-Upload '/importar' 'vendas' $arqSet 'vendas_2026_setembro.xlsx' $cookie
    $tok = Get-TokenImportacao $u
    Confere 'setembro: resumo com 82 novas antes de gravar' ($u.Status -eq 200 -and $u.Corpo -match 'Linhas novas</span><span class="valor">82<' -and $u.Corpo -match 'R\$ 3\.850\.268,28 faturados' -and $tok)
    Confere 'setembro: nada muda antes de confirmar' ((Invoke-Http 'GET' '/' '' $cookie).Corpo -match '723 vendas')
    $c = Invoke-Http 'POST' '/importar/confirmar' (Form @{ token = $tok }) $cookie
    Confere 'setembro: confirmado e gravado' ($c.Status -eq 200 -and $c.Corpo -match 'Importação gravada' -and $c.Corpo -match '82 vendas novas')
    $ps = Invoke-Http 'GET' '/?de=9&ate=9' '' $cookie
    Confere 'painel: setembro R$ 3.850.268,28, parcial até 19/09 (§4.5)' ($ps.Corpo -match 'R\$ 3\.850\.268,28' -and $ps.Corpo -match 'parcial até 19/09/2026')
    Confere 'confirmar de novo o mesmo token: 409' ((Invoke-Http 'POST' '/importar/confirmar' (Form @{ token = $tok }) $cookie).Status -eq 409)
    $u = Invoke-Upload '/importar' 'vendas' $arqSet 'vendas_2026_setembro.xlsx' $cookie
    Confere 'reimportar setembro: 0 novas e 82 ignoradas' ($u.Corpo -match 'Linhas novas</span><span class="valor">0<' -and $u.Corpo -match 'Ignoradas</span><span class="valor">82<')
    $x = Invoke-Http 'POST' '/importar/cancelar' (Form @{ token = (Get-TokenImportacao $u) }) $cookie
    Confere 'cancelar volta para /importar sem gravar' ($x.Status -eq 303 -and $x.Location -eq '/importar' -and ([regex]::Matches((Invoke-Http 'GET' '/importar' '' $cookie).Corpo, '82 novas · 0 atualizadas')).Count -eq 1)
    $u = Invoke-Upload '/importar' 'estoque' $arqEst19 'estoque_2026-09-19.xlsx' $cookie
    Confere 'estoque 19/09: resumo com 195 linhas substituídas' ($u.Status -eq 200 -and $u.Corpo -match 'Linhas substituídas</span><span class="valor">195<' -and $u.Corpo -match 'passa a ser a foto em vigor')
    [void](Invoke-Http 'POST' '/importar/confirmar' (Form @{ token = (Get-TokenImportacao $u) }) $cookie)
    Confere 'estoque: tela passa a usar a foto de 19/09' ((Invoke-Http 'GET' '/estoque' '' $cookie).Corpo -match 'Foto do estoque de <strong>19/09/2026</strong>')
    $u = Invoke-Upload '/importar' 'estoque' (Join-Path $dadosTeste 'estoque_2026-08-31.xlsx') 'estoque_2026-08-31.xlsx' $cookie
    Confere 'estoque 31/08 depois de 19/09: recusado por ser mais antigo' ($u.Status -eq 422 -and $u.Corpo -match 'mais antiga que a foto em vigor')
    Confere 'dados reais intocados' (-not (Test-Path -LiteralPath (Join-Path $dadosReais 'importacoes\vendas')) -or @(Get-ChildItem -LiteralPath (Join-Path $dadosReais 'importacoes\vendas') -File | Where-Object { $_.LastWriteTime -gt $inicioTeste }).Count -eq 0)

    Write-Host "`n== Sessão ==" -ForegroundColor Cyan
    $r = Entrar 'mariana.ribeiro@horizontemaquinas.com.br' $senhaMar
    $cookieMar = ($r.SetCookie -split ';')[0]
    $pm = Invoke-Http 'GET' '/' '' $cookieMar
    Confere 'vendedora ativa entra e vê a própria linha destacada' ($pm.Status -eq 200 -and $pm.Corpo -match '<tr class="voce"><td class="pos">1</td><td><a href="/vendedor\?id=V003')
    $s = Invoke-Http 'POST' '/sair' '' $cookie
    Confere 'sair: redireciona para /entrar' ($s.Status -eq 303 -and $s.Location -eq '/entrar')
    Confere 'após sair, cookie antigo não vale' ((Invoke-Http 'GET' '/' '' $cookie).Status -eq 303)
    Confere 'cookie inventado não vale' ((Invoke-Http 'GET' '/' '' 'sessao=abc').Status -eq 303)
    for ($i = 0; $i -lt 5; $i++) { [void](Entrar 'teste.diretoria' 'errada') }
    $r = Entrar 'teste.diretoria' $senhaDir
    Confere '5 erros seguidos bloqueiam até a senha certa (429)' ($r.Status -eq 429)
} finally {
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}

Write-Host ""
if ($falhas) { Write-Host "$falhas de $total verificações FALHARAM." -ForegroundColor Red; exit 1 }
Write-Host "Todas as $total verificações passaram." -ForegroundColor Green
exit 0
