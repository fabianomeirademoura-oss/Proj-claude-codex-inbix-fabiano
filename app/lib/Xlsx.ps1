# Leitor mínimo de .xlsx (Office Open XML) sem dependências externas.
# Devolve as linhas de uma aba como objetos, usando a primeira linha como cabeçalho.
# Números vêm como texto invariante ("8891.280000000001"); células com formato
# de data vêm como "yyyy-MM-dd". A conversão de tipos fica com quem chama.

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:NsPlanilha = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsRelacao = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:NsPacote = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Get-ZipXml {
    param([IO.Compression.ZipArchive]$Zip, [string]$Nome)
    $entrada = $Zip.GetEntry($Nome)
    if ($null -eq $entrada) { return $null }
    $leitor = New-Object IO.StreamReader($entrada.Open(), [Text.Encoding]::UTF8)
    try {
        $xml = New-Object Xml.XmlDocument
        $xml.LoadXml($leitor.ReadToEnd())
        return , $xml
    } finally { $leitor.Dispose() }
}

function Get-IndiceColuna([string]$Referencia) {
    $n = 0
    foreach ($c in ($Referencia -replace '\d', '').ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
    return $n - 1
}

function Test-FormatoData([int]$NumFmtId, [hashtable]$Personalizados) {
    if (($NumFmtId -ge 14 -and $NumFmtId -le 22) -or ($NumFmtId -ge 45 -and $NumFmtId -le 47)) { return $true }
    if ($Personalizados.ContainsKey($NumFmtId)) {
        $codigo = $Personalizados[$NumFmtId] -replace '"[^"]*"', '' -replace '\[[^\]]*\]', '' -replace '\\.', ''
        return $codigo -match '[dmyhs]'
    }
    return $false
}

function Read-XlsxSheet {
    param(
        [Parameter(Mandatory)][string]$Caminho,
        [Parameter(Mandatory)][string]$Aba
    )
    # FileShare.ReadWrite: funciona mesmo com a planilha aberta no Excel.
    $arquivo = [IO.File]::Open($Caminho, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $zip = New-Object IO.Compression.ZipArchive($arquivo, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $ns = $null

        $compartilhadas = New-Object Collections.Generic.List[string]
        $xmlStrings = Get-ZipXml $zip 'xl/sharedStrings.xml'
        if ($xmlStrings) {
            $ns = New-Object Xml.XmlNamespaceManager($xmlStrings.NameTable)
            $ns.AddNamespace('m', $script:NsPlanilha)
            foreach ($si in $xmlStrings.SelectNodes('/m:sst/m:si', $ns)) {
                $partes = foreach ($t in $si.SelectNodes('m:t | m:r/m:t', $ns)) { $t.InnerText }
                $compartilhadas.Add(($partes -join ''))
            }
        }

        $estilosData = @{}
        $xmlEstilos = Get-ZipXml $zip 'xl/styles.xml'
        if ($xmlEstilos) {
            $ns = New-Object Xml.XmlNamespaceManager($xmlEstilos.NameTable)
            $ns.AddNamespace('m', $script:NsPlanilha)
            $personalizados = @{}
            foreach ($nf in $xmlEstilos.SelectNodes('/m:styleSheet/m:numFmts/m:numFmt', $ns)) {
                $personalizados[[int]$nf.GetAttribute('numFmtId')] = $nf.GetAttribute('formatCode')
            }
            $i = 0
            foreach ($xf in $xmlEstilos.SelectNodes('/m:styleSheet/m:cellXfs/m:xf', $ns)) {
                if (Test-FormatoData ([int]$xf.GetAttribute('numFmtId')) $personalizados) { $estilosData[$i] = $true }
                $i++
            }
        }

        $xmlLivro = Get-ZipXml $zip 'xl/workbook.xml'
        $ns = New-Object Xml.XmlNamespaceManager($xmlLivro.NameTable)
        $ns.AddNamespace('m', $script:NsPlanilha)
        $noAba = $null
        foreach ($s in $xmlLivro.SelectNodes('/m:workbook/m:sheets/m:sheet', $ns)) {
            if ($s.GetAttribute('name') -eq $Aba) { $noAba = $s }
        }
        if ($null -eq $noAba) { throw "Aba '$Aba' não encontrada em $Caminho" }
        $idRel = $noAba.GetAttribute('id', $script:NsRelacao)

        $xmlRel = Get-ZipXml $zip 'xl/_rels/workbook.xml.rels'
        $nsRel = New-Object Xml.XmlNamespaceManager($xmlRel.NameTable)
        $nsRel.AddNamespace('r', $script:NsPacote)
        $alvo = $xmlRel.SelectSingleNode("/r:Relationships/r:Relationship[@Id='$idRel']", $nsRel).GetAttribute('Target')
        $alvo = if ($alvo.StartsWith('/')) { $alvo.TrimStart('/') } else { 'xl/' + $alvo }

        $xmlAba = Get-ZipXml $zip $alvo
        $ns = New-Object Xml.XmlNamespaceManager($xmlAba.NameTable)
        $ns.AddNamespace('m', $script:NsPlanilha)

        $cabecalho = $null
        $linhas = New-Object Collections.Generic.List[object]
        foreach ($row in $xmlAba.SelectNodes('/m:worksheet/m:sheetData/m:row', $ns)) {
            $valores = @{}
            $maior = -1
            foreach ($c in $row.SelectNodes('m:c', $ns)) {
                $col = Get-IndiceColuna $c.GetAttribute('r')
                $tipo = $c.GetAttribute('t')
                $noV = $c.SelectSingleNode('m:v', $ns)
                $bruto = if ($noV) { $noV.InnerText } else { $null }
                $valor = switch ($tipo) {
                    's' { $compartilhadas[[int]$bruto] }
                    'inlineStr' { $c.SelectSingleNode('m:is', $ns).InnerText }
                    'e' { throw "Célula com erro ($bruto) em $Aba!$($c.GetAttribute('r')) de $Caminho" }
                    default {
                        $estilo = $c.GetAttribute('s')
                        if ($null -ne $bruto -and $tipo -ne 'str' -and $estilo -and $estilosData.ContainsKey([int]$estilo)) {
                            [DateTime]::FromOADate([double]::Parse($bruto, [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
                        } else { $bruto }
                    }
                }
                $valores[$col] = $valor
                if ($col -gt $maior) { $maior = $col }
            }
            if ($null -eq $cabecalho) {
                $cabecalho = for ($k = 0; $k -le $maior; $k++) { ([string]$valores[$k]).Trim() }
                continue
            }
            $vazia = $true
            $obj = [ordered]@{}
            for ($k = 0; $k -lt $cabecalho.Count; $k++) {
                $v = $valores[$k]
                if ($null -ne $v) { $v = ([string]$v).Trim(); if ($v -ne '') { $vazia = $false } else { $v = $null } }
                $obj[$cabecalho[$k]] = $v
            }
            if (-not $vazia) { $obj['_Linha'] = [int]$row.GetAttribute('r'); $linhas.Add([pscustomobject]$obj) }
        }
        return , $linhas.ToArray()
    } finally {
        $zip.Dispose()
        $arquivo.Dispose()
    }
}
