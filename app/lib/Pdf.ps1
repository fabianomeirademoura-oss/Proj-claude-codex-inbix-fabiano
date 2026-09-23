# Gerador mínimo de PDF, sem dependências: texto, linhas, retângulos e tabelas com quebra de página.
# Usa as fontes padrão Helvetica e Helvetica-Bold (não são embutidas) com WinAnsiEncoding, que cobre
# os acentos do português. Não grava data nem identificador aleatório: o mesmo conteúdo gera o mesmo
# arquivo, byte a byte. As coordenadas das funções públicas são em pontos, medidas a partir do topo.

# Larguras (1/1000 do corpo) dos caracteres 32 a 126, das métricas AFM padrão.
$script:PdfLarguras = @{
    F1 = @(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)
    F2 = @(278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584)
}
# Caracteres fora do ASCII: código WinAnsi (0x80–0x9F) e largura [F1, F2].
$script:PdfEspeciais = @{
    ([char]0x20AC) = @(0x80, 556, 556); ([char]0x2026) = @(0x85, 1000, 1000); ([char]0x2018) = @(0x91, 222, 278)
    ([char]0x2019) = @(0x92, 222, 278); ([char]0x201C) = @(0x93, 333, 500); ([char]0x201D) = @(0x94, 333, 500)
    ([char]0x2022) = @(0x95, 350, 350); ([char]0x2013) = @(0x96, 556, 556); ([char]0x2014) = @(0x97, 1000, 1000)
}
# Latin-1 (0xA0–0xFF): largura pela letra-base; os demais símbolos com a métrica própria.
$script:PdfBaseLatin1 = @{ ([char]0xA0) = ' '; ([char]0xC7) = 'C'; ([char]0xE7) = 'c'; ([char]0xD1) = 'N'; ([char]0xF1) = 'n' }
$script:PdfLargurasLatin1 = @{ ([char]0xB0) = @(400, 400); ([char]0xBA) = @(365, 365); ([char]0xAA) = @(370, 370); ([char]0xA7) = @(556, 556); ([char]0xD7) = @(584, 584); ([char]0xB7) = @(278, 278) }
$script:PdfTroca = @{ ([char]0x2265) = '>='; ([char]0x2264) = '<='; ([char]0x2212) = '-'; ([char]0x2192) = '->'; ([char]0x2248) = '~' }

function ConvertTo-PdfTexto([string]$Texto) {
    # Normaliza para acentos compostos e troca o que o WinAnsi não tem.
    if (-not $Texto) { return '' }
    $t = $Texto.Normalize([Text.NormalizationForm]::FormC)
    foreach ($k in $script:PdfTroca.Keys) { $t = $t.Replace([string]$k, $script:PdfTroca[$k]) }
    return $t
}

function Get-PdfLarguraChar([char]$C, [string]$Fonte) {
    $i = if ($Fonte -eq 'F2') { 1 } else { 0 }
    $n = [int]$C
    if ($n -ge 32 -and $n -le 126) { return $script:PdfLarguras[$Fonte][$n - 32] }
    if ($script:PdfEspeciais.ContainsKey($C)) { return $script:PdfEspeciais[$C][1 + $i] }
    if ($script:PdfLargurasLatin1.ContainsKey($C)) { return $script:PdfLargurasLatin1[$C][$i] }
    if ($n -ge 0xA0 -and $n -le 0xFF) {
        if ($script:PdfBaseLatin1.ContainsKey($C)) { return (Get-PdfLarguraChar $script:PdfBaseLatin1[$C] $Fonte) }
        $base = ([string]$C).Normalize([Text.NormalizationForm]::FormD)[0]
        if ([int]$base -lt 127) {
            if ($base -eq 'i') { return 278 }   # í, ì, î, ï usam o i sem pingo
            return (Get-PdfLarguraChar $base $Fonte)
        }
        return 556
    }
    return 556
}

function Get-PdfLargura([string]$Texto, [string]$Fonte = 'F1', [double]$Tamanho = 9) {
    $soma = 0
    foreach ($c in (ConvertTo-PdfTexto $Texto).ToCharArray()) { $soma += Get-PdfLarguraChar $c $Fonte }
    return $soma * $Tamanho / 1000
}

function Limit-PdfTexto([string]$Texto, [double]$Largura, [string]$Fonte = 'F1', [double]$Tamanho = 9) {
    # Corta com reticências o que não cabe na largura.
    $t = ConvertTo-PdfTexto $Texto
    if ((Get-PdfLargura $t $Fonte $Tamanho) -le $Largura) { return $t }
    $ret = [string][char]0x2026
    while ($t.Length -gt 0 -and (Get-PdfLargura ($t + $ret) $Fonte $Tamanho) -gt $Largura) { $t = $t.Substring(0, $t.Length - 1) }
    return ($t.TrimEnd() + $ret)
}

function Split-PdfTexto([string]$Texto, [double]$Largura, [string]$Fonte = 'F1', [double]$Tamanho = 9) {
    # Quebra em linhas pelo espaço; palavra maior que a linha é cortada.
    $linhas = New-Object Collections.Generic.List[string]
    foreach ($paragrafo in ((ConvertTo-PdfTexto $Texto) -split "`n")) {
        $atual = ''
        foreach ($palavra in ($paragrafo -split ' ' | Where-Object { $_ -ne '' })) {
            $teste = if ($atual) { "$atual $palavra" } else { $palavra }
            if ((Get-PdfLargura $teste $Fonte $Tamanho) -le $Largura) { $atual = $teste; continue }
            if ($atual) { $linhas.Add($atual) }
            $atual = if ((Get-PdfLargura $palavra $Fonte $Tamanho) -le $Largura) { $palavra } else { Limit-PdfTexto $palavra $Largura $Fonte $Tamanho }
        }
        $linhas.Add($atual)
    }
    return $linhas.ToArray()
}

function ConvertTo-PdfString([string]$Texto) {
    # Literal PDF: cada caractere vira um byte WinAnsi (guardado como char 0–255); escapa \ ( ).
    $sb = New-Object Text.StringBuilder
    foreach ($c in (ConvertTo-PdfTexto $Texto).ToCharArray()) {
        $n = [int]$c
        if ($c -eq '\' -or $c -eq '(' -or $c -eq ')') { [void]$sb.Append('\').Append($c) }
        elseif ($n -ge 32 -and $n -le 126) { [void]$sb.Append($c) }
        elseif ($script:PdfEspeciais.ContainsKey($c)) { [void]$sb.Append([char]$script:PdfEspeciais[$c][0]) }
        elseif ($n -ge 0xA0 -and $n -le 0xFF) { [void]$sb.Append($c) }
        else { [void]$sb.Append('?') }
    }
    return $sb.ToString()
}

function Format-PdfNum([double]$N) {
    # Número com até 2 casas, ponto decimal, sem notação científica.
    return ([math]::Round($N, 2)).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-PdfCor($Cor) { return ($Cor | ForEach-Object { Format-PdfNum $_ }) -join ' ' }

# ---------------------------------------------------------------- documento

function New-PdfDocumento {
    param([double]$Largura = 841.89, [double]$Altura = 595.28, [double]$Margem = 36, [double]$Rodape = 20)
    return [pscustomobject]@{
        Largura = $Largura; Altura = $Altura; Margem = $Margem; Rodape = $Rodape
        Paginas = New-Object Collections.Generic.List[Text.StringBuilder]
        Y       = 0
        AoAbrirPagina = $null   # scriptblock chamado a cada página nova (cabeçalho)
    }
}

function Add-PdfPagina($Doc) {
    $Doc.Paginas.Add((New-Object Text.StringBuilder))
    $Doc.Y = $Doc.Margem
    if ($Doc.AoAbrirPagina) { & $Doc.AoAbrirPagina $Doc }
}

function Get-PdfFundo($Doc) { return $Doc.Altura - $Doc.Margem - $Doc.Rodape }

function Confirm-PdfEspaco($Doc, [double]$Altura) {
    # Abre página nova se a altura pedida não cabe. Devolve $true quando abriu.
    if ($Doc.Paginas.Count -eq 0 -or $Doc.Y + $Altura -gt (Get-PdfFundo $Doc)) { Add-PdfPagina $Doc; return $true }
    return $false
}

function Add-PdfOp($Doc, [string]$Op, [int]$Pagina = -1) {
    $p = if ($Pagina -ge 0) { $Doc.Paginas[$Pagina] } else { $Doc.Paginas[$Doc.Paginas.Count - 1] }
    [void]$p.Append($Op).Append("`n")
}

function Add-PdfTexto {
    param($Doc, [double]$X, [double]$Y, [string]$Texto, [string]$Fonte = 'F1', [double]$Tamanho = 9, $Cor = @(0.1, 0.1, 0.1), [int]$Pagina = -1)
    # Y = linha de base, medida do topo.
    if (-not $Texto) { return }
    $op = "BT /$Fonte $(Format-PdfNum $Tamanho) Tf $(Format-PdfCor $Cor) rg $(Format-PdfNum $X) $(Format-PdfNum ($Doc.Altura - $Y)) Td ($(ConvertTo-PdfString $Texto)) Tj ET"
    Add-PdfOp $Doc $op $Pagina
}

function Add-PdfTextoAlinhado {
    param($Doc, [double]$X, [double]$Largura, [double]$Y, [string]$Texto, [string]$Alinha = 'E', [string]$Fonte = 'F1', [double]$Tamanho = 9, $Cor = @(0.1, 0.1, 0.1), [int]$Pagina = -1)
    $t = Limit-PdfTexto $Texto $Largura $Fonte $Tamanho
    $w = Get-PdfLargura $t $Fonte $Tamanho
    $x0 = switch ($Alinha) { 'D' { $X + $Largura - $w } 'C' { $X + ($Largura - $w) / 2 } default { $X } }
    Add-PdfTexto -Doc $Doc -X $x0 -Y $Y -Texto $t -Fonte $Fonte -Tamanho $Tamanho -Cor $Cor -Pagina $Pagina
}

function Add-PdfRetangulo {
    param($Doc, [double]$X, [double]$Y, [double]$Largura, [double]$Altura, $Cor = @(0.9, 0.9, 0.9))
    Add-PdfOp $Doc "$(Format-PdfCor $Cor) rg $(Format-PdfNum $X) $(Format-PdfNum ($Doc.Altura - $Y - $Altura)) $(Format-PdfNum $Largura) $(Format-PdfNum $Altura) re f"
}

function Add-PdfLinha {
    param($Doc, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, $Cor = @(0.6, 0.6, 0.6), [double]$Espessura = 0.5, [int]$Pagina = -1)
    Add-PdfOp $Doc "$(Format-PdfCor $Cor) RG $(Format-PdfNum $Espessura) w $(Format-PdfNum $X1) $(Format-PdfNum ($Doc.Altura - $Y1)) m $(Format-PdfNum $X2) $(Format-PdfNum ($Doc.Altura - $Y2)) l S" $Pagina
}

function Add-PdfParagrafo {
    param($Doc, [string]$Texto, [string]$Fonte = 'F1', [double]$Tamanho = 9, $Cor = @(0.15, 0.15, 0.15), [double]$Recuo = 0, [double]$Entrelinha = 1.35, [double]$EspacoDepois = 4)
    $largura = $Doc.Largura - 2 * $Doc.Margem - $Recuo
    $altLinha = $Tamanho * $Entrelinha
    foreach ($l in (Split-PdfTexto $Texto $largura $Fonte $Tamanho)) {
        [void](Confirm-PdfEspaco $Doc $altLinha)
        Add-PdfTexto -Doc $Doc -X ($Doc.Margem + $Recuo) -Y ($Doc.Y + $Tamanho) -Texto $l -Fonte $Fonte -Tamanho $Tamanho -Cor $Cor
        $Doc.Y += $altLinha
    }
    $Doc.Y += $EspacoDepois
}

function Add-PdfTabela {
    # Colunas: @{ Titulo; Largura (fração da largura útil); Alinha = 'E'|'D'|'C' }.
    # Linhas: @{ Celulas = @(...); Estilo = ''|'total'|'destaque'|'grupo' }. O cabeçalho se repete a cada página.
    param($Doc, $Colunas, $Linhas, [double]$Tamanho = 8, [double]$AlturaLinha = 14, [string]$Vazio = 'Nenhuma linha.')
    $util = $Doc.Largura - 2 * $Doc.Margem
    $xs = @(); $ws = @(); $x = $Doc.Margem
    foreach ($c in $Colunas) { $w = $c.Largura * $util; $xs += $x; $ws += $w; $x += $w }
    $pad = 3
    $cabecalho = {
        Add-PdfRetangulo -Doc $Doc -X $Doc.Margem -Y $Doc.Y -Largura $util -Altura $AlturaLinha -Cor @(0.87, 0.9, 0.94)
        for ($i = 0; $i -lt $Colunas.Count; $i++) {
            $al = if ($Colunas[$i].Alinha) { $Colunas[$i].Alinha } else { 'E' }
            Add-PdfTextoAlinhado -Doc $Doc -X ($xs[$i] + $pad) -Largura ($ws[$i] - 2 * $pad) -Y ($Doc.Y + $AlturaLinha - 4) -Texto $Colunas[$i].Titulo -Alinha $al -Fonte 'F2' -Tamanho $Tamanho -Cor @(0.12, 0.2, 0.32)
        }
        $Doc.Y += $AlturaLinha
    }
    [void](Confirm-PdfEspaco $Doc (2 * $AlturaLinha))
    & $cabecalho
    if (-not $Linhas -or @($Linhas).Count -eq 0) {
        Add-PdfTexto -Doc $Doc -X ($Doc.Margem + $pad) -Y ($Doc.Y + $AlturaLinha - 4) -Texto $Vazio -Tamanho $Tamanho -Cor @(0.4, 0.4, 0.4)
        $Doc.Y += $AlturaLinha
        return
    }
    $n = 0
    $lista = @($Linhas)
    for ($k = 0; $k -lt $lista.Count; $k++) {
        $linha = $lista[$k]
        # Sem viúva: as últimas linhas (e o total) não vão sozinhas para a página seguinte.
        $resto = $lista.Count - $k
        $precisa = if ($resto -le 3) { $resto * $AlturaLinha } else { $AlturaLinha }
        if (Confirm-PdfEspaco $Doc $precisa) { & $cabecalho }
        $estilo = [string]$linha.Estilo
        $fonte = if ($estilo -eq 'total' -or $estilo -eq 'grupo') { 'F2' } else { 'F1' }
        switch ($estilo) {
            'destaque' { Add-PdfRetangulo -Doc $Doc -X $Doc.Margem -Y $Doc.Y -Largura $util -Altura $AlturaLinha -Cor @(1, 0.92, 0.9) }
            'grupo' { Add-PdfRetangulo -Doc $Doc -X $Doc.Margem -Y $Doc.Y -Largura $util -Altura $AlturaLinha -Cor @(0.94, 0.94, 0.94) }
            'total' { Add-PdfLinha -Doc $Doc -X1 $Doc.Margem -Y1 $Doc.Y -X2 ($Doc.Margem + $util) -Y2 $Doc.Y -Cor @(0.3, 0.3, 0.3) -Espessura 0.8 }
            default { if ($n % 2 -eq 1) { Add-PdfRetangulo -Doc $Doc -X $Doc.Margem -Y $Doc.Y -Largura $util -Altura $AlturaLinha -Cor @(0.97, 0.97, 0.98) } }
        }
        for ($i = 0; $i -lt $Colunas.Count; $i++) {
            $al = if ($Colunas[$i].Alinha) { $Colunas[$i].Alinha } else { 'E' }
            Add-PdfTextoAlinhado -Doc $Doc -X ($xs[$i] + $pad) -Largura ($ws[$i] - 2 * $pad) -Y ($Doc.Y + $AlturaLinha - 4) -Texto ([string]$linha.Celulas[$i]) -Alinha $al -Fonte $fonte -Tamanho $Tamanho
        }
        $Doc.Y += $AlturaLinha
        if ($estilo -ne 'grupo') { $n++ } else { $n = 0 }
    }
    Add-PdfLinha -Doc $Doc -X1 $Doc.Margem -Y1 $Doc.Y -X2 ($Doc.Margem + $util) -Y2 $Doc.Y -Cor @(0.75, 0.75, 0.75)
}

function Save-PdfDocumento {
    param($Doc, [Parameter(Mandatory)][string]$Caminho, [string]$Titulo = '', [string]$Produtor = '')
    $objs = New-Object Collections.Generic.List[string]
    $nPag = $Doc.Paginas.Count
    $kids = (0..($nPag - 1) | ForEach-Object { "$(6 + 2 * $_) 0 R" }) -join ' '
    $objs.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objs.Add("<< /Type /Pages /Kids [$kids] /Count $nPag >>")
    $objs.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>')
    $objs.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>')
    $objs.Add("<< /Title ($(ConvertTo-PdfString $Titulo)) /Producer ($(ConvertTo-PdfString $Produtor)) >>")
    for ($i = 0; $i -lt $nPag; $i++) {
        $conteudo = $Doc.Paginas[$i].ToString()
        $objs.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $(Format-PdfNum $Doc.Largura) $(Format-PdfNum $Doc.Altura)] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents $(7 + 2 * $i) 0 R >>")
        $objs.Add("<< /Length $($conteudo.Length) >>`nstream`n$conteudo`nendstream")
    }
    # Cada char é um byte (0–255): o tamanho em chars é o deslocamento em bytes.
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n%" + [char]0xE2 + [char]0xE3 + [char]0xCF + [char]0xD3 + "`n")
    $offsets = New-Object Collections.Generic.List[int]
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offsets.Add($sb.Length)
        [void]$sb.Append("$($i + 1) 0 obj`n").Append($objs[$i]).Append("`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offsets) { [void]$sb.Append($o.ToString('0000000000')).Append(" 00000 n `n") }
    [void]$sb.Append("trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R /Info 5 0 R >>`nstartxref`n$xref`n%%EOF`n")
    $texto = $sb.ToString()
    $bytes = New-Object byte[] $texto.Length
    for ($i = 0; $i -lt $texto.Length; $i++) { $bytes[$i] = [byte][int]$texto[$i] }
    [IO.File]::WriteAllBytes($Caminho, $bytes)
}
