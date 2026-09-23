# Só para os testes: grava um .xlsx mínimo (uma aba, texto e número) para montar
# arquivos de importação com defeitos de propósito (coluna faltando, ID repetido…).
# Não é usado pela aplicação.

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-LetraColuna([int]$Indice) {
    $s = ''; $n = $Indice + 1
    while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char](65 + $r) + $s; $n = [math]::Floor(($n - 1) / 26) }
    return $s
}

function New-XlsxTeste {
    param(
        [Parameter(Mandatory)][string]$Caminho,
        [Parameter(Mandatory)][string]$Aba,
        [Parameter(Mandatory)][string[]]$Cabecalho,
        [object[]]$Linhas = @()   # cada linha: hashtable/objeto com as colunas do cabeçalho
    )
    $esc = { param($t) [Security.SecurityElement]::Escape([string]$t) }
    $sb = New-Object Text.StringBuilder
    $todas = @(, $null) + @($Linhas)
    for ($i = 0; $i -lt $todas.Count; $i++) {
        [void]$sb.Append("<row r=""$($i + 1)"">")
        for ($c = 0; $c -lt $Cabecalho.Count; $c++) {
            $ref = "$(Get-LetraColuna $c)$($i + 1)"
            $v = if ($i -eq 0) { $Cabecalho[$c] } else { $todas[$i].($Cabecalho[$c]) }
            if ($null -eq $v -or [string]$v -eq '') { continue }
            if ($i -gt 0 -and [string]$v -match '^-?\d+(\.\d+)?$') { [void]$sb.Append("<c r=""$ref""><v>$v</v></c>") }
            else { [void]$sb.Append("<c r=""$ref"" t=""inlineStr""><is><t>$(& $esc $v)</t></is></c>") }
        }
        [void]$sb.Append('</row>')
    }
    $partes = [ordered]@{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'            = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets><sheet name=""$(& $esc $Aba)"" sheetId=""1"" r:id=""rId1""/></sheets></workbook>"
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml'   = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData>$($sb.ToString())</sheetData></worksheet>"
    }
    if (Test-Path -LiteralPath $Caminho) { Remove-Item -LiteralPath $Caminho -Force }
    $arq = [IO.File]::Open($Caminho, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($arq, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($nome in $partes.Keys) {
            $w = New-Object IO.StreamWriter($zip.CreateEntry($nome).Open(), (New-Object Text.UTF8Encoding($false)))
            $w.Write($partes[$nome]); $w.Dispose()
        }
    } finally { $zip.Dispose(); $arq.Dispose() }
}
