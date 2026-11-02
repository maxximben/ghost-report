param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Pdf = "",
    [switch]$Paragraphs
)
# Открывает .docx в Microsoft Word (COM), считает страницы, положения закладок _Toc*, строк таблиц и (по флагу)
# заголовков; по желанию экспортирует PDF. Результат – JSON в файле $Out. Исходный файл не изменяется.
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($Docx, $false, $true, $false)
    $doc.Repaginate()
    $result = [ordered]@{}
    $result.pages = [int]$doc.ComputeStatistics(2)

    $bm = [ordered]@{}
    $doc.Bookmarks.ShowHidden = $true
    foreach ($b in $doc.Bookmarks) {
        if ($b.Name -like '_Toc*') {
            $r = $b.Range
            $r.Collapse(1)
            $bm[$b.Name] = [int]$r.Information(3)
        }
    }
    $result.bookmarks = $bm

    $tables = @()
    for ($i = 1; $i -le $doc.Tables.Count; $i++) {
        $t = $doc.Tables.Item($i)
        $pg = @()
        for ($j = 1; $j -le $t.Rows.Count; $j++) {
            $r = $t.Rows.Item($j).Range
            $r.Collapse(1)
            $pg += [int]$r.Information(3)
        }
        $tables += [ordered]@{ title = [string]$t.Title; rows = $pg }
    }
    $result.tables = $tables

    if ($Paragraphs) {
        $hs = @()
        foreach ($p in $doc.Paragraphs) {
            $sn = [string]$p.Style.NameLocal
            if ($sn -like 'Heading*' -or $sn -like 'Заголовок*' -or $sn -like 'Структурный*') {
                $r = $p.Range
                $r.Collapse(1)
                $txt = $p.Range.Text -replace '[\r\n\a]', ' '
                $hs += [ordered]@{ text = $txt.Trim(); style = $sn; page = [int]$r.Information(3) }
            }
        }
        $result.headings = $hs
    }

    if ($Pdf -ne "") {
        $doc.ExportAsFixedFormat($Pdf, 17)
    }

    $json = $result | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding($false)))
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }
    if ($word -ne $null) { $word.Quit() }
}
