<#
  OfficeDiff.ps1  ―  2つの Word / PowerPoint / Excel ファイルの変更点を .html（ブラウザ表示用）と .diff（unified diff 形式）に出力する

  使い方:
    OfficeDiff.bat に2つのファイルをまとめてドラッグ＆ドロップ
    1つだけドロップした場合や、ダブルクリックした場合は、足りない分をファイル選択画面で選ぶ
    （更新日時が古い方を「変更前」、新しい方を「変更後」として比較）

  直接実行する場合:
    powershell -ExecutionPolicy Bypass -File OfficeDiff.ps1 <ファイル1> <ファイル2> [-Swap] [-NoOpen] [-OutDir <フォルダ>]
      -Swap    変更前／変更後を入れ替える
      -NoOpen  出力した .html をブラウザで開かない
      -OutDir  出力先フォルダ（省略時は このツールと同じフォルダの output。無ければ作成）

  ・Office のインストールは不要（.docx / .pptx / .xlsx の中身を直接読む）
  ・Windows PowerShell 5.1（Windows 10/11 標準）で動作
  ・テキスト化と差分のルールは GitHub の PR コメント版（scripts/office_diff.py）と同じ
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths,
    [switch]$Swap,
    [switch]$NoOpen,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'
# 出力をファイル等にリダイレクトしたときも日本語が化けないようにする（画面表示はそのまま）
if ([Console]::IsOutputRedirected) { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 注意: PowerShell の変数名は大文字小文字を区別しないため、ローカルの $ns と衝突しない名前にする
$OOXML = @{
    w   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    p   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
}

# ===================================================================== ZIP / XML

function Open-OfficeZip([string]$Path) {
    # Office で開いたままのファイルも読めるよう、共有モードで開く
    $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        return New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
    } catch {
        $fs.Dispose()
        throw "ファイルを読み込めません（パスワード付き、または旧形式 .doc/.ppt の可能性があります）: $Path"
    }
}

function Get-ZipXml($Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.Load($stream)
        return ,$xml
    } finally { $stream.Dispose() }
}

function New-Ns($Xml) {
    $m = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    foreach ($k in $OOXML.Keys) { $m.AddNamespace($k, $OOXML[$k]) }
    return ,$m   # IEnumerable なので , を付けて展開を防ぐ
}

function Resolve-PartPath([string]$BaseDir, [string]$Target) {
    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($s in ($BaseDir + '/' + $Target).Split('/')) {
        if ($s -eq '' -or $s -eq '.') { continue }
        if ($s -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) } }
        else { $parts.Add($s) }
    }
    return ($parts -join '/')
}

function Get-Rels($Zip, [string]$PartPath) {
    $i = $PartPath.LastIndexOf('/')
    $dir = $PartPath.Substring(0, [Math]::Max($i, 0))
    $name = $PartPath.Substring($i + 1)
    $relsPath = $(if ($dir) { "$dir/_rels/$name.rels" } else { "_rels/$name.rels" })
    $result = @{}
    $x = Get-ZipXml $Zip $relsPath
    if ($null -eq $x) { return $result }
    foreach ($r in $x.DocumentElement.ChildNodes) {
        if ($r.LocalName -ne 'Relationship') { continue }
        if ($r.GetAttribute('TargetMode') -eq 'External') { continue }
        $result[$r.GetAttribute('Id')] = [pscustomobject]@{
            Type   = $r.GetAttribute('Type')
            Target = (Resolve-PartPath $dir $r.GetAttribute('Target'))
        }
    }
    return $result
}

# ========================================================================= Word

function Get-WordParaText($P, $Ns) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in $P.SelectNodes('.//w:r/w:t | .//w:r/w:tab | .//w:r/w:br | .//w:r/w:cr', $Ns)) {
        switch ($n.LocalName) {
            't'   { [void]$sb.Append($n.InnerText) }
            'tab' { [void]$sb.Append("`t") }
            default { [void]$sb.Append(' ') }
        }
    }
    return $sb.ToString().Trim()
}

function Add-WordBlocks($Nodes, $Ns, $Styles, $Lines) {
    foreach ($node in $Nodes) {
        switch ($node.LocalName) {
            'p' {
                $text = Get-WordParaText $node $Ns
                if (-not $text) { continue }
                $styleName = ''
                $ps = $node.SelectSingleNode('w:pPr/w:pStyle', $Ns)
                if ($ps) {
                    $id = $ps.GetAttribute('val', $OOXML.w)
                    $styleName = $(if ($Styles.ContainsKey($id)) { $Styles[$id] } else { $id })
                }
                if ($styleName -match '^(?i:heading|見出し)\s*(\d)') {
                    $Lines.Add('')
                    $Lines.Add(('#' * [int]$Matches[1]) + ' ' + $text)
                } elseif ($styleName -match '(?i)list' -or $node.SelectSingleNode('w:pPr/w:numPr', $Ns)) {
                    $Lines.Add('・' + $text)
                } else {
                    $Lines.Add($text)
                }
            }
            'tbl' {
                $Lines.Add('')
                $first = $true
                foreach ($tr in $node.SelectNodes('w:tr', $Ns)) {
                    $cells = @()
                    foreach ($tc in $tr.SelectNodes('w:tc', $Ns)) {
                        $ptexts = @()
                        foreach ($p in $tc.SelectNodes('.//w:p', $Ns)) { $ptexts += (Get-WordParaText $p $Ns) }
                        $cells += (($ptexts -join ' ').Trim())
                    }
                    $Lines.Add('| ' + ($cells -join ' | ') + ' |')
                    if ($first) { $Lines.Add('|' + ('---|' * $cells.Count)); $first = $false }
                }
                $Lines.Add('')
            }
            'sdt' {
                # 目次などのコンテンツコントロールの中身
                $content = $node.SelectSingleNode('w:sdtContent', $Ns)
                if ($content) { Add-WordBlocks $content.ChildNodes $Ns $Styles $Lines }
            }
        }
    }
}

function Get-WordLines($Zip) {
    $doc = Get-ZipXml $Zip 'word/document.xml'
    if ($null -eq $doc) { throw 'word/document.xml が見つかりません' }
    $ns = New-Ns $doc
    $styles = @{}
    $st = Get-ZipXml $Zip 'word/styles.xml'
    if ($st) {
        $sns = New-Ns $st
        foreach ($s in $st.SelectNodes('//w:style', $sns)) {
            $n = $s.SelectSingleNode('w:name', $sns)
            if ($n) { $styles[$s.GetAttribute('styleId', $OOXML.w)] = $n.GetAttribute('val', $OOXML.w) }
        }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    $body = $doc.SelectSingleNode('/w:document/w:body', $ns)
    Add-WordBlocks $body.ChildNodes $ns $styles $lines
    # 先頭・末尾の空行と、連続する空行を整理
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($l -eq '' -and ($out.Count -eq 0 -or $out[$out.Count - 1] -eq '')) { continue }
        $out.Add($l)
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    return ,$out.ToArray()
}

# =================================================================== PowerPoint

function Get-ATextOfPara($P, $Ns) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $P.SelectNodes('a:r/a:t | a:fld/a:t', $Ns)) { [void]$sb.Append($t.InnerText) }
    return $sb.ToString().Trim()
}

function Add-ShapeLines($Container, $Ns, $Lines, $State) {
    foreach ($sh in $Container.ChildNodes) {
        switch ($sh.LocalName) {
            'sp' {
                $tx = $sh.SelectSingleNode('p:txBody', $Ns)
                if ($null -eq $tx) { continue }
                $ph = $sh.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $Ns)
                $isTitle = $ph -and @('title', 'ctrTitle') -contains $ph.GetAttribute('type')
                $all = @()
                foreach ($p in $tx.SelectNodes('a:p', $Ns)) {
                    $text = Get-ATextOfPara $p $Ns
                    if (-not $text) { continue }
                    $all += $text
                    $lvl = 0
                    $pPr = $p.SelectSingleNode('a:pPr', $Ns)
                    if ($pPr -and $pPr.GetAttribute('lvl')) { $lvl = [int]$pPr.GetAttribute('lvl') }
                    $Lines.Add(('  ' * $lvl) + '・' + $text)
                }
                if ($isTitle -and -not $State.Title) { $State.Title = ($all -join ' ') }
            }
            'graphicFrame' {
                $tbl = $sh.SelectSingleNode('a:graphic/a:graphicData/a:tbl', $Ns)
                if ($null -eq $tbl) { continue }
                foreach ($tr in $tbl.SelectNodes('a:tr', $Ns)) {
                    $cells = @()
                    foreach ($tc in $tr.SelectNodes('a:tc', $Ns)) {
                        $pt = @()
                        foreach ($p in $tc.SelectNodes('.//a:p', $Ns)) { $pt += (Get-ATextOfPara $p $Ns) }
                        $cells += (($pt -join ' ').Trim())
                    }
                    $Lines.Add('| ' + ($cells -join ' | ') + ' |')
                }
            }
            'pic' {
                $c = $sh.SelectSingleNode('p:nvPicPr/p:cNvPr', $Ns)
                if ($c) { $Lines.Add('[画像: ' + $c.GetAttribute('name') + ']') }
            }
            'grpSp' { Add-ShapeLines $sh $Ns $Lines $State }
        }
    }
}

function Get-Slides($Zip) {
    $pres = Get-ZipXml $Zip 'ppt/presentation.xml'
    if ($null -eq $pres) { throw 'ppt/presentation.xml が見つかりません' }
    $ns = New-Ns $pres
    $rels = Get-Rels $Zip 'ppt/presentation.xml'
    $slides = New-Object System.Collections.Generic.List[object]
    foreach ($sid in $pres.SelectNodes('/p:presentation/p:sldIdLst/p:sldId', $ns)) {
        $rid = $sid.GetAttribute('id', $OOXML.r)
        if (-not $rels.ContainsKey($rid)) { continue }
        $part = $rels[$rid].Target
        $sx = Get-ZipXml $Zip $part
        if ($null -eq $sx) { continue }
        $sns = New-Ns $sx
        $lines = New-Object System.Collections.Generic.List[string]
        $state = @{ Title = '' }
        $tree = $sx.SelectSingleNode('/p:sld/p:cSld/p:spTree', $sns)
        if ($tree) { Add-ShapeLines $tree $sns $lines $state }
        # ノート
        foreach ($r in (Get-Rels $Zip $part).Values) {
            if ($r.Type -notlike '*/notesSlide') { continue }
            $nx = Get-ZipXml $Zip $r.Target
            if ($null -eq $nx) { continue }
            $nns = New-Ns $nx
            foreach ($sp in $nx.SelectNodes('//p:sp', $nns)) {
                $ph = $sp.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $nns)
                if (-not $ph -or $ph.GetAttribute('type') -ne 'body') { continue }
                $pt = @()
                foreach ($p in $sp.SelectNodes('p:txBody/a:p', $nns)) {
                    $pt += (($p.SelectNodes('.//a:t', $nns) | ForEach-Object { $_.InnerText }) -join '')
                }
                $notes = ($pt -join "`n").Trim()
                if ($notes) { $lines.Add('> ノート: ' + ($notes -replace "\r?\n", ' / ')) }
            }
        }
        $title = $state.Title
        if (-not $title -and $lines.Count -gt 0) { $title = $lines[0].TrimStart(' ', '・') }
        $slides.Add([pscustomobject]@{ Id = $sid.GetAttribute('id'); Title = $title; Lines = $lines.ToArray() })
    }
    return ,$slides.ToArray()
}

# ======================================================================== Excel

$XLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function New-XlNs($Xml) {
    $m = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $m.AddNamespace('x', $XLNS)
    $m.AddNamespace('r', $OOXML.r)
    return ,$m
}

function Get-XlsxDateStyles($Zip) {
    # セルの書式番号（s 属性）ごとに「日付・時刻の書式か」を返す
    $result = New-Object System.Collections.Generic.List[bool]
    $st = Get-ZipXml $Zip 'xl/styles.xml'
    if ($null -eq $st) { return ,$result.ToArray() }
    $ns = New-XlNs $st
    $custom = @{}
    foreach ($f in $st.SelectNodes('/x:styleSheet/x:numFmts/x:numFmt', $ns)) {
        $code = $f.GetAttribute('formatCode')
        # "文字列" [色/条件] \x _x *x を除いてから、日付・時刻の記号があるかを見る
        $plain = ($code -replace '"[^"]*"', '' -replace '\[[^\]]*\]', '' -replace '\\.', '' -replace '[_*].', '')
        $custom[[int]$f.GetAttribute('numFmtId')] = ($plain -match '[ymdhsYMDHSgegE]' -and $plain -notmatch '^(?i)general$')
    }
    # Excel 組み込みの日付・時刻書式（日本語版の和暦・年月日表記を含む）
    $builtin = @(14..22) + @(27..36) + @(45..47) + @(50..58)
    foreach ($xf in $st.SelectNodes('/x:styleSheet/x:cellXfs/x:xf', $ns)) {
        $id = [int]$xf.GetAttribute('numFmtId')
        $result.Add($(if ($custom.ContainsKey($id)) { $custom[$id] } else { $builtin -contains $id }))
    }
    return ,$result.ToArray()
}

function Format-XlsxNumber([string]$V, [bool]$IsDate) {
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $d = 0.0
    if (-not [double]::TryParse($V, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { return $V }
    if ($IsDate -and $d -ge 0 -and $d -lt 2958466) {
        $dt = [DateTime]::FromOADate($d)
        if ($d -eq [Math]::Floor($d)) { return $dt.ToString('yyyy-MM-dd') }
        if ($d -lt 1) { return $dt.ToString('HH:mm:ss') }
        return $dt.ToString('yyyy-MM-dd HH:mm:ss')
    }
    # 0.1+0.2=0.30000000000000004 のような誤差を見せないよう 15 桁に丸める
    return $d.ToString('G15', $inv)
}

function Read-XlsxRows($Stream, $Sst, $DateStyles) {
    # シートの XML を先頭から順に読む。大きなシート（数万セル）でも速いよう、
    # XPath や関数呼び出しを使わず、よくある「文字列」「数値」のセルはこの中で直接処理する
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $rd = [System.Xml.XmlReader]::Create($Stream, $settings)
    $ELEM = [System.Xml.XmlNodeType]::Element; $END = [System.Xml.XmlNodeType]::EndElement
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $nstyle = [System.Globalization.NumberStyles]::Float
    $nDate = $DateStyles.Count
    $rows = New-Object System.Collections.Generic.List[object]
    $sb = New-Object System.Text.StringBuilder
    $inRow = $false; $inCell = $false; $rowNo = 0; $colNo = 0
    $col = ''; $t = ''; $st = ''; $v = ''; $f = $null; $inl = ''
    $d = 0.0
    $advance = $true   # ReadElementContentAsString / Skip は読み位置を進めるので、そのときは Read() しない
    try {
        while ($true) {
            if ($advance -and -not $rd.Read()) { break }
            if ($rd.EOF) { break }
            $advance = $true
            $nt = $rd.NodeType
            if ($nt -eq $ELEM) {
                $ln = $rd.LocalName
                if ($ln -ceq 'c') {
                    if (-not $inRow) { continue }
                    $colNo++
                    $ref = $rd.GetAttribute('r')
                    $col = $(if ($ref) { $ref.TrimEnd('0123456789'.ToCharArray()) } else { [string][char](64 + $colNo) })
                    $t = $rd.GetAttribute('t'); $st = $rd.GetAttribute('s'); $v = ''; $f = $null; $inl = ''
                    $inCell = -not $rd.IsEmptyElement   # 書式だけの空セルは読まない
                } elseif ($ln -ceq 'v') {
                    if ($inCell) { $v = $rd.ReadElementContentAsString(); $advance = $false }
                } elseif ($ln -ceq 'row') {
                    $r = $rd.GetAttribute('r')
                    $rowNo = $(if ($r) { [int]$r } else { $rowNo + 1 }); $colNo = 0
                    [void]$sb.Clear()
                    $inRow = -not $rd.IsEmptyElement
                } elseif ($ln -ceq 'f') {
                    if ($inCell) {
                        if ($rd.IsEmptyElement) { $f = '' } else { $f = $rd.ReadElementContentAsString(); $advance = $false }
                    }
                } elseif ($ln -ceq 't') {
                    if ($inCell -and $t -ceq 'inlineStr') { $inl += $rd.ReadElementContentAsString(); $advance = $false }
                } elseif ($ln -ceq 'rPh') {
                    $rd.Skip(); $advance = $false   # ふりがなは読まない
                }
            } elseif ($nt -eq $END) {
                $ln = $rd.LocalName
                if ($ln -ceq 'c' -and $inCell) {
                    $inCell = $false
                    if ($t -ceq 's') {
                        $i = 0
                        $val = $(if ([int]::TryParse($v, [ref]$i) -and $i -lt $Sst.Count) { $Sst[$i] } else { '' })
                    } elseif (-not $t -or $t -ceq 'n') {
                        if ($v -eq '') { $val = '' }
                        else {
                            $si = $(if ($st) { [int]$st } else { 0 })
                            if ($si -lt $nDate -and $DateStyles[$si]) { $val = Format-XlsxNumber $v $true }
                            elseif ([double]::TryParse($v, $nstyle, $inv, [ref]$d)) { $val = $d.ToString('G15', $inv) }
                            else { $val = $v }
                        }
                    } elseif ($t -ceq 'inlineStr') { $val = $inl }
                    elseif ($t -ceq 'b') { $val = $(if ($v -eq '1') { 'TRUE' } elseif ($v -eq '0') { 'FALSE' } else { '' }) }
                    else { $val = $v }   # str（数式の文字列結果）・e（エラー）
                    if ($val.IndexOf("`n") -ge 0) { $val = ($val -replace "\r?\n", ' / ') }
                    if ($null -ne $f) {
                        $fx = $(if ($f) { '=' + $f } else { '=(共有数式)' })
                        $val = $(if ($val -ne '') { "$fx → $val" } else { $fx })
                    }
                    if ($val -ne '') {
                        if ($sb.Length) { [void]$sb.Append(' | ') }
                        [void]$sb.Append($col).Append(':').Append($val)
                    }
                } elseif ($ln -ceq 'row') {
                    if ($inRow -and $sb.Length) { $rows.Add([pscustomobject]@{ Num = $rowNo; Text = $sb.ToString() }) }
                    $inRow = $false
                }
            }
        }
    } finally { $rd.Dispose() }
    return ,$rows.ToArray()
}

function Get-XlsxSheets($Zip) {
    $wb = Get-ZipXml $Zip 'xl/workbook.xml'
    if ($null -eq $wb) { throw 'xl/workbook.xml が見つかりません' }
    $ns = New-XlNs $wb
    $rels = Get-Rels $Zip 'xl/workbook.xml'

    # 共有文字列（ふりがな rPh は除く）
    $sst = New-Object System.Collections.Generic.List[string]
    $sstPath = 'xl/sharedStrings.xml'
    foreach ($r in $rels.Values) { if ($r.Type -like '*/sharedStrings') { $sstPath = $r.Target } }
    $sx = Get-ZipXml $Zip $sstPath
    if ($sx) {
        $sns = New-XlNs $sx
        foreach ($si in $sx.SelectNodes('/x:sst/x:si', $sns)) {
            $sst.Add((($si.SelectNodes('x:t | x:r/x:t', $sns) | ForEach-Object { $_.InnerText }) -join ''))
        }
    }
    $dateStyles = Get-XlsxDateStyles $Zip

    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($sh in $wb.SelectNodes('/x:workbook/x:sheets/x:sheet', $ns)) {
        $name = $sh.GetAttribute('name')
        $rid = $sh.GetAttribute('id', $OOXML.r)
        $rows = @()
        if ($rels.ContainsKey($rid) -and $rels[$rid].Type -like '*/worksheet') {
            $entry = $Zip.GetEntry($rels[$rid].Target)
            if ($entry) {
                $stream = $entry.Open()
                try { $rows = Read-XlsxRows $stream $sst $dateStyles } finally { $stream.Dispose() }
            }
        }
        $sheets.Add([pscustomobject]@{ Name = $name; Rows = @($rows) })
    }
    return ,$sheets.ToArray()
}

function Format-XlsxRow($Row) { return ('{0}行目 | {1}' -f $Row.Num, $Row.Text) }

function Get-RowDiffLines($A, $B, [int]$Context = 2) {
    # 行単位で比べる（行を挿入・削除しても、それ以降の行がすべて変更扱いにならない）
    $ta = [string[]]@($A | ForEach-Object { $_.Text }); $tb = [string[]]@($B | ForEach-Object { $_.Text })
    # 計算表がこれより大きくなる（巨大なシートで変更範囲も広い）ときは、行番号どうしの突き合わせに切り替える
    $max = $(if ($env:OFFICEDIFF_XLSX_MAXCELLS) { [long]$env:OFFICEDIFF_XLSX_MAXCELLS } else { 4000000 })
    $ops = Get-EditOps $ta $tb $max
    $disp = New-Object System.Collections.Generic.List[object]
    if ($null -ne $ops) {
        foreach ($o in $ops) {
            $row = $(if ($o.Tag -eq '-') { $A[$o.AI] } else { $B[$o.BI] })
            $disp.Add(@{ Tag = $o.Tag; Line = (Format-XlsxRow $row); Num = $row.Num })
        }
    } else {
        # 非常に大きなシートで変更範囲も広いときは、行番号どうしで突き合わせる
        $ma = @{}; foreach ($r in $A) { $ma[$r.Num] = $r }
        $mb = @{}; foreach ($r in $B) { $mb[$r.Num] = $r }
        foreach ($num in (@($ma.Keys) + @($mb.Keys) | Sort-Object -Unique)) {
            $ra = $ma[$num]; $rb = $mb[$num]
            if ($ra -and $rb -and $ra.Text -ceq $rb.Text) { $disp.Add(@{ Tag = ' '; Line = (Format-XlsxRow $rb); Num = $num }); continue }
            if ($ra) { $disp.Add(@{ Tag = '-'; Line = (Format-XlsxRow $ra); Num = $num }) }
            if ($rb) { $disp.Add(@{ Tag = '+'; Line = (Format-XlsxRow $rb); Num = $num }) }
        }
    }
    $out = New-Object System.Collections.Generic.List[string]
    $changed = @(); for ($k = 0; $k -lt $disp.Count; $k++) { if ($disp[$k].Tag -ne ' ') { $changed += $k } }
    if ($changed.Count -eq 0) { return ,$out.ToArray() }
    $groups = @(); $s = $changed[0]; $e = $changed[0]
    for ($c = 1; $c -lt $changed.Count; $c++) {
        $k = $changed[$c]
        if ($k - $e -le 2 * $Context) { $e = $k } else { $groups += , @($s, $e); $s = $k; $e = $k }
    }
    $groups += , @($s, $e)
    foreach ($g in $groups) {
        $from = [Math]::Max(0, $g[0] - $Context); $to = [Math]::Min($disp.Count - 1, $g[1] + $Context)
        $out.Add("@@ $($disp[$g[0]].Num)行目付近 @@")
        for ($k = $from; $k -le $to; $k++) { $out.Add($disp[$k].Tag + $disp[$k].Line) }
    }
    return ,$out.ToArray()
}

function Get-SheetDiffLines($Old, $New, [int]$MaxRows = 200) {
    # シートは名前で対応付ける。名前が見つからないシート同士は、並び順と中身が似ていれば「名前変更」とみなす
    $newPos = @{}; for ($j = 0; $j -lt $New.Count; $j++) { $newPos[$New[$j].Name] = $j }
    $oldNames = @{}; foreach ($s in $Old) { $oldNames[$s.Name] = $true }
    $unOld = @(for ($i = 0; $i -lt $Old.Count; $i++) { if (-not $newPos.ContainsKey($Old[$i].Name)) { $i } })
    $unNew = @(for ($j = 0; $j -lt $New.Count; $j++) { if (-not $oldNames.ContainsKey($New[$j].Name)) { $j } })
    $renamed = @{}   # 旧index → 新index
    $u = 0
    foreach ($i in $unOld) {
        if ($u -ge $unNew.Count) { break }
        $j = $unNew[$u]
        $sa = [string[]]@($Old[$i].Rows | Select-Object -First 200 | ForEach-Object { $_.Text })
        $sb = [string[]]@($New[$j].Rows | Select-Object -First 200 | ForEach-Object { $_.Text })
        if ((Get-LcsRatio $sa $sb) -ge 0.5) { $renamed[$i] = $j; $u++ }
    }
    $matchedNew = @{}; foreach ($j in $renamed.Values) { $matchedNew[$j] = $true }

    $keyed = @(); $lastNew = -1; $seq = 0
    for ($i = 0; $i -lt $Old.Count; $i++) {
        $j = $(if ($newPos.ContainsKey($Old[$i].Name)) { $newPos[$Old[$i].Name] } elseif ($renamed.ContainsKey($i)) { $renamed[$i] } else { $null })
        if ($null -ne $j) { $lastNew = $j; $keyed += [pscustomobject]@{ K = [double]$j; Seq = $seq++; I = $i; J = $j } }
        else { $keyed += [pscustomobject]@{ K = $lastNew + 0.5; Seq = $seq++; I = $i; J = $null } }
    }
    for ($j = 0; $j -lt $New.Count; $j++) {
        if (-not $oldNames.ContainsKey($New[$j].Name) -and -not $matchedNew.ContainsKey($j)) {
            $keyed += [pscustomobject]@{ K = [double]$j; Seq = $seq++; I = $null; J = $j }
        }
    }

    $out = New-Object System.Collections.Generic.List[string]
    $addAll = {
        param($Tag, $Rows)
        $n = 0
        foreach ($r in $Rows) {
            if ($n -ge $MaxRows) { $out.Add(" …ほか $($Rows.Count - $MaxRows) 行"); break }
            $out.Add($Tag + (Format-XlsxRow $r)); $n++
        }
        if ($Rows.Count -eq 0) { $out.Add(' （空のシート）') }
    }
    foreach ($p in ($keyed | Sort-Object -Property K, Seq)) {
        if ($null -ne $p.I -and $null -ne $p.J) {
            $lines = Get-RowDiffLines $Old[$p.I].Rows $New[$p.J].Rows
            $rename = ($Old[$p.I].Name -cne $New[$p.J].Name)
            if ($lines.Count -eq 0 -and -not $rename) { continue }
            $title = $(if ($rename) { "$($Old[$p.I].Name) → $($New[$p.J].Name)" } else { $New[$p.J].Name })
            $out.Add(" ## シート: $title")
            if ($lines.Count -eq 0) { $out.Add(' （シート名の変更のみ）') }
            foreach ($l in $lines) { $out.Add($l) }
        } elseif ($null -ne $p.I) {
            $out.Add("-## シート: $($Old[$p.I].Name)（削除）")
            & $addAll '-' $Old[$p.I].Rows
        } else {
            $out.Add("+## シート: $($New[$p.J].Name)（追加）")
            & $addAll '+' $New[$p.J].Rows
        }
        $out.Add(' ')
    }
    if ($out.Count -gt 0) { $out.RemoveAt($out.Count - 1) }
    return ,$out.ToArray()
}

# ========================================================================= 差分

function Get-EditOps([string[]]$A, [string[]]$B, [long]$MaxCells = 0) {
    # LCS による編集列。戻り値: @{Tag=' '|'-'|'+'; Text; AI; BI} の配列（AI/BI は0始まりの行番号）
    # MaxCells を指定すると、計算表がそれより大きくなる場合は計算せず $null を返す（巨大な Excel シート用）
    if ($null -eq $A) { $A = @() }
    if ($null -eq $B) { $B = @() }
    $ops = New-Object System.Collections.Generic.List[object]
    $n = $A.Count; $m = $B.Count
    $pre = 0
    while ($pre -lt $n -and $pre -lt $m -and $A[$pre] -ceq $B[$pre]) { $pre++ }
    $suf = 0
    while ($suf -lt ($n - $pre) -and $suf -lt ($m - $pre) -and $A[$n - 1 - $suf] -ceq $B[$m - 1 - $suf]) { $suf++ }
    for ($k = 0; $k -lt $pre; $k++) { $ops.Add(@{ Tag = ' '; Text = $A[$k]; AI = $k; BI = $k }) }

    $an = $n - $pre - $suf; $bm = $m - $pre - $suf
    if ($MaxCells -gt 0 -and [long]$an * $bm -gt $MaxCells) { return $null }
    $dp = New-Object 'int[][]' ($an + 1)
    for ($i = 0; $i -le $an; $i++) { $dp[$i] = New-Object 'int[]' ($bm + 1) }
    for ($i = $an - 1; $i -ge 0; $i--) {
        $ai = $A[$pre + $i]; $row = $dp[$i]; $next = $dp[$i + 1]
        for ($j = $bm - 1; $j -ge 0; $j--) {
            if ($ai -ceq $B[$pre + $j]) { $row[$j] = $next[$j + 1] + 1 }
            elseif ($next[$j] -ge $row[$j + 1]) { $row[$j] = $next[$j] }
            else { $row[$j] = $row[$j + 1] }
        }
    }
    $i = 0; $j = 0
    while ($i -lt $an -or $j -lt $bm) {
        if ($i -lt $an -and $j -lt $bm -and $A[$pre + $i] -ceq $B[$pre + $j]) {
            $ops.Add(@{ Tag = ' '; Text = $A[$pre + $i]; AI = $pre + $i; BI = $pre + $j }); $i++; $j++
        } elseif ($j -ge $bm -or ($i -lt $an -and $dp[$i + 1][$j] -ge $dp[$i][$j + 1])) {
            $ops.Add(@{ Tag = '-'; Text = $A[$pre + $i]; AI = $pre + $i; BI = $pre + $j }); $i++
        } else {
            $ops.Add(@{ Tag = '+'; Text = $B[$pre + $j]; AI = $pre + $i; BI = $pre + $j }); $j++
        }
    }
    for ($k = 0; $k -lt $suf; $k++) {
        $ops.Add(@{ Tag = ' '; Text = $A[$n - $suf + $k]; AI = $n - $suf + $k; BI = $m - $suf + $k })
    }
    return ,$ops.ToArray()
}

function Get-LcsRatio([string[]]$A, [string[]]$B) {
    $t = $A.Count + $B.Count
    if ($t -eq 0) { return 1.0 }
    # 関数の戻り値は配列1個として渡ってくるので、いったん変数に受けてから数える
    $ops = Get-EditOps $A $B
    $same = @($ops | Where-Object { $_.Tag -eq ' ' }).Count
    return (2.0 * $same / $t)
}

function Get-UnifiedLines([string[]]$A, [string[]]$B, [int]$Context = 2) {
    # Word 用: 変更箇所の前後 $Context 行だけを @@ 見出し付きで出す
    $ops = Get-EditOps $A $B
    $out = New-Object System.Collections.Generic.List[string]
    $changed = @(); for ($k = 0; $k -lt $ops.Count; $k++) { if ($ops[$k].Tag -ne ' ') { $changed += $k } }
    if ($changed.Count -eq 0) { return ,$out.ToArray() }
    # 近い変更（間が 2×Context 行以内）は1つのまとまりにする
    $groups = @(); $s = $changed[0]; $e = $changed[0]
    for ($c = 1; $c -lt $changed.Count; $c++) {
        $k = $changed[$c]
        if ($k - $e -le 2 * $Context) { $e = $k } else { $groups += , @($s, $e); $s = $k; $e = $k }
    }
    $groups += , @($s, $e)
    foreach ($g in $groups) {
        $from = [Math]::Max(0, $g[0] - $Context); $to = [Math]::Min($ops.Count - 1, $g[1] + $Context)
        $slice = $ops[$from..$to]
        $aCount = @($slice | Where-Object { $_.Tag -ne '+' }).Count
        $bCount = @($slice | Where-Object { $_.Tag -ne '-' }).Count
        $out.Add(('@@ -{0},{1} +{2},{3} @@' -f ($slice[0].AI + 1), $aCount, ($slice[0].BI + 1), $bCount))
        foreach ($o in $slice) { $out.Add($o.Tag + $o.Text) }
    }
    return ,$out.ToArray()
}

function Test-Similar($S1, $S2) {
    if ($S1.Title -and $S1.Title -ceq $S2.Title) { return $true }
    return ((Get-LcsRatio $S1.Lines $S2.Lines) -ge 0.5)
}

function Get-SlidePairs($Old, $New) {
    # スライドIDで対応付け（PowerPoint では並べ替え・削除しても変わらない）
    $newPos = @{}; for ($j = 0; $j -lt $New.Count; $j++) { $newPos[$New[$j].Id] = $j }
    $oldIds = @{}; foreach ($s in $Old) { $oldIds[$s.Id] = $true }
    $keyed = @(); $lastNew = -1; $seq = 0
    for ($i = 0; $i -lt $Old.Count; $i++) {
        if ($newPos.ContainsKey($Old[$i].Id)) {
            $j = $newPos[$Old[$i].Id]; $lastNew = $j
            $keyed += [pscustomobject]@{ K = [double]$j; Seq = $seq++; I = $i; J = $j }
        } else {
            $keyed += [pscustomobject]@{ K = $lastNew + 0.5; Seq = $seq++; I = $i; J = $null }   # 削除
        }
    }
    for ($j = 0; $j -lt $New.Count; $j++) {
        if (-not $oldIds.ContainsKey($New[$j].Id)) { $keyed += [pscustomobject]@{ K = [double]$j; Seq = $seq++; I = $null; J = $j } }
    }
    # Windows PowerShell 5.1 には -Stable が無いため、登録順（Seq）を第2キーにする
    $pairs = @($keyed | Sort-Object -Property K, Seq)

    # ID が振り直されている（別ツールで保存し直した等）場合はタイトルの並びで対応付ける
    $both = @($pairs | Where-Object { $null -ne $_.I -and $null -ne $_.J })
    if ($Old.Count -gt 0 -and $New.Count -gt 0) {
        $ok = @($both | Where-Object { Test-Similar $Old[$_.I] $New[$_.J] }).Count
        if ($both.Count -eq 0 -or $ok * 2 -lt $both.Count) { return ,(Get-PairsByTitle $Old $New) }
    }
    return ,$pairs
}

function Get-PairsByTitle($Old, $New) {
    $ops = Get-EditOps @($Old | ForEach-Object { $_.Title }) @($New | ForEach-Object { $_.Title })
    $pairs = New-Object System.Collections.Generic.List[object]
    $dels = New-Object System.Collections.Generic.List[int]
    $adds = New-Object System.Collections.Generic.List[int]
    # 番兵として末尾に「一致」を足し、置換ブロック（削除＋追加の連続）を処理しきる
    foreach ($o in (@($ops) + @(@{ Tag = 'END' }))) {
        if ($o.Tag -eq '-') { $dels.Add($o.AI); continue }
        if ($o.Tag -eq '+') { $adds.Add($o.BI); continue }
        $n = [Math]::Min($dels.Count, $adds.Count)
        for ($k = 0; $k -lt $n; $k++) { $pairs.Add([pscustomobject]@{ I = $dels[$k]; J = $adds[$k] }) }
        for ($k = $n; $k -lt $dels.Count; $k++) { $pairs.Add([pscustomobject]@{ I = $dels[$k]; J = $null }) }
        for ($k = $n; $k -lt $adds.Count; $k++) { $pairs.Add([pscustomobject]@{ I = $null; J = $adds[$k] }) }
        $dels.Clear(); $adds.Clear()
        if ($o.Tag -eq ' ') { $pairs.Add([pscustomobject]@{ I = $o.AI; J = $o.BI }) }
    }
    return ,$pairs.ToArray()
}

function Get-SlideDiffLines($Old, $New) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($pr in (Get-SlidePairs $Old $New)) {
        $i = $pr.I; $j = $pr.J
        if ($null -ne $i -and $null -ne $j) {
            $a = $Old[$i].Lines; $b = $New[$j].Lines
            if (($a -join "`n") -ceq ($b -join "`n")) { continue }   # 変更なし（番号ずれのみも出さない）
            $num = $(if ($i -ne $j) { "$($i + 1) → $($j + 1)" } else { "$($j + 1)" })
            $out.Add(" ## スライド $num")
            foreach ($o in (Get-EditOps $a $b)) { $out.Add($o.Tag + $o.Text) }
        } elseif ($null -ne $i) {
            $out.Add("-## スライド $($i + 1)（削除）")
            foreach ($l in $Old[$i].Lines) { $out.Add('-' + $l) }
        } else {
            $out.Add("+## スライド $($j + 1)（追加）")
            foreach ($l in $New[$j].Lines) { $out.Add('+' + $l) }
        }
        $out.Add(' ')
    }
    if ($out.Count -gt 0) { $out.RemoveAt($out.Count - 1) }
    return ,$out.ToArray()
}

# ===================================================================== HTML出力

function ConvertTo-Html([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }

function Get-CharDiffHtml([string]$A, [string]$B) {
    # 1行の中で変わった文字だけを <mark> で囲む（「3営業日」→「5営業日」の「3」「5」など）
    $ea = ConvertTo-Html $A; $eb = ConvertTo-Html $B
    if ($A.Length * $B.Length -gt 250000) { return @($ea, $eb) }   # 長すぎる行は行単位の色分けのみ
    $ca = $A.ToCharArray(); $cb = $B.ToCharArray()
    $n = $ca.Length; $m = $cb.Length
    $dp = New-Object 'int[][]' ($n + 1)
    for ($i = 0; $i -le $n; $i++) { $dp[$i] = New-Object 'int[]' ($m + 1) }
    for ($i = $n - 1; $i -ge 0; $i--) {
        $row = $dp[$i]; $next = $dp[$i + 1]; $c = $ca[$i]
        for ($j = $m - 1; $j -ge 0; $j--) {
            if ($c -ceq $cb[$j]) { $row[$j] = $next[$j + 1] + 1 }
            elseif ($next[$j] -ge $row[$j + 1]) { $row[$j] = $next[$j] }
            else { $row[$j] = $row[$j + 1] }
        }
    }
    # 似ていない行同士（共通部分が少ない）は、部分ハイライトせず行全体の色分けのみ
    if ($n + $m -gt 0 -and (2.0 * $dp[0][0] / ($n + $m)) -lt 0.4) { return @($ea, $eb) }
    $sa = New-Object System.Text.StringBuilder; $sb = New-Object System.Text.StringBuilder
    $i = 0; $j = 0; $inA = $false; $inB = $false
    while ($i -lt $n -or $j -lt $m) {
        if ($i -lt $n -and $j -lt $m -and $ca[$i] -ceq $cb[$j]) {
            if ($inA) { [void]$sa.Append('</mark>'); $inA = $false }
            if ($inB) { [void]$sb.Append('</mark>'); $inB = $false }
            $t = ConvertTo-Html ([string]$ca[$i]); [void]$sa.Append($t); [void]$sb.Append($t); $i++; $j++
        } elseif ($j -ge $m -or ($i -lt $n -and $dp[$i + 1][$j] -ge $dp[$i][$j + 1])) {
            if (-not $inA) { [void]$sa.Append('<mark>'); $inA = $true }
            [void]$sa.Append((ConvertTo-Html ([string]$ca[$i]))); $i++
        } else {
            if (-not $inB) { [void]$sb.Append('<mark>'); $inB = $true }
            [void]$sb.Append((ConvertTo-Html ([string]$cb[$j]))); $j++
        }
    }
    if ($inA) { [void]$sa.Append('</mark>') }
    if ($inB) { [void]$sb.Append('</mark>') }
    return @($sa.ToString(), $sb.ToString())
}

function Get-LineRatio([string]$A, [string]$B) {
    # 2行がどのくらい似ているか（0〜1）。行内ハイライトする組み合わせを選ぶのに使う
    $n = $A.Length; $m = $B.Length
    if ($n + $m -eq 0) { return 1.0 }
    if ($n * $m -gt 250000) { return 0.0 }
    # 長さが大きく違えば似ていない（計算を省く）
    if ((2.0 * [Math]::Min($n, $m) / ($n + $m)) -lt 0.4) { return 0.0 }
    # 文字単位の最長共通部分列（LCS）の長さを、1行分の配列だけで計算する
    $prev = New-Object 'int[]' ($m + 1); $cur = New-Object 'int[]' ($m + 1)
    for ($i = 1; $i -le $n; $i++) {
        $c = $A[$i - 1]
        for ($j = 1; $j -le $m; $j++) {
            if ($c -ceq $B[$j - 1]) { $cur[$j] = $prev[$j - 1] + 1 }
            elseif ($prev[$j] -ge $cur[$j - 1]) { $cur[$j] = $prev[$j] }
            else { $cur[$j] = $cur[$j - 1] }
        }
        $t = $prev; $prev = $cur; $cur = $t
    }
    return (2.0 * $prev[$m] / ($n + $m))
}

function Get-ChangeBlockHtml($Dels, $Adds) {
    # 連続する削除行・追加行を、似ている行同士で組にして行内の違いをハイライトする
    $htmlD = @($Dels | ForEach-Object { ConvertTo-Html $_ }); $htmlA = @($Adds | ForEach-Object { ConvertTo-Html $_ })
    $next = 0
    for ($x = 0; $x -lt $Dels.Count; $x++) {
        for ($y = $next; $y -lt $Adds.Count; $y++) {
            if ((Get-LineRatio $Dels[$x] $Adds[$y]) -ge 0.4) {
                $pair = Get-CharDiffHtml $Dels[$x] $Adds[$y]
                $htmlD[$x] = $pair[0]; $htmlA[$y] = $pair[1]; $next = $y + 1; break
            }
        }
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($h in $htmlD) { [void]$sb.Append("<div class=`"ln del`"><span class=`"sg`">-</span><span class=`"tx`">$h</span></div>`n") }
    foreach ($h in $htmlA) { [void]$sb.Append("<div class=`"ln add`"><span class=`"sg`">+</span><span class=`"tx`">$h</span></div>`n") }
    return $sb.ToString()
}

function Get-DiffHtmlBody([string[]]$Lines, [bool]$SkipBlank) {
    # 1) 差分行を見出し（スライド／シート／Wordの変更箇所）ごとのまとまりに分ける
    $secs = New-Object System.Collections.Generic.List[object]
    $cur = $null
    foreach ($l in $Lines) {
        $hdr = $null
        if ($l -match '^ ## ((?:スライド|シート:) .+)$') { $hdr = @{ Title = $Matches[1]; Kind = 'chg' } }
        elseif ($l -match '^-## ((?:スライド|シート:) .+?)（削除）$') { $hdr = @{ Title = $Matches[1] + '（削除）'; Kind = 'del' } }
        elseif ($l -match '^\+## ((?:スライド|シート:) .+?)（追加）$') { $hdr = @{ Title = $Matches[1] + '（追加）'; Kind = 'add' } }
        # Word の @@ は新しい変更箇所。Excel のシート内の @@ はシートの中の区切り
        elseif ($l.StartsWith('@@') -and -not ($cur -and $cur.Kind -ne 'hunk')) { $hdr = @{ Title = ''; Kind = 'hunk' } }
        if ($hdr) { $cur = $hdr; $cur.Lines = New-Object System.Collections.Generic.List[string]; $secs.Add($cur); continue }
        if ($null -eq $cur) { $cur = @{ Title = ''; Kind = 'hunk'; Lines = (New-Object System.Collections.Generic.List[string]) }; $secs.Add($cur) }
        $cur.Lines.Add($l)
    }

    # 2) Word の変更箇所には、近くの見出し（# 〜）を名前として付ける
    $lastHeading = ''; $no = 0
    foreach ($sec in $secs) {
        if ($sec.Kind -ne 'hunk') { continue }
        $no++
        $title = ''
        foreach ($l in $sec.Lines) {
            if ($l.StartsWith('-') -or $l.StartsWith('+')) { break }
            if ($l -match '^ (#+) (.+)$') { $title = $Matches[2] }
        }
        if (-not $title) { $title = $lastHeading }
        $sec.Title = $(if ($title) { "$title 付近" } else { "変更箇所 $no" })
        foreach ($l in $sec.Lines) { if ($l -match '^[ +](#+) (.+)$') { $lastHeading = $Matches[2] } }
    }

    # 3) HTML にする
    $body = New-Object System.Text.StringBuilder
    $toc = New-Object System.Collections.Generic.List[string]
    $label = @{ chg = '変更'; del = '削除'; add = '追加'; hunk = '変更' }
    $i = 0
    foreach ($sec in $secs) {
        $i++
        $t = ConvertTo-Html $sec.Title; $k = $sec.Kind
        [void]$body.Append("<section id=`"s$i`" class=`"sec $k`"><h2><span class=`"badge $k`">$($label[$k])</span>$t</h2><div class=`"code`">`n")
        $toc.Add("<li><a href=`"#s$i`"><span class=`"badge $k`">$($label[$k])</span>$t</a></li>")
        $dels = New-Object System.Collections.Generic.List[string]; $adds = New-Object System.Collections.Generic.List[string]
        foreach ($l in (@($sec.Lines) + @('<END>'))) {
            $tag = $(if ($l -eq '<END>') { 'END' } elseif ($l.Length) { $l.Substring(0, 1) } else { ' ' })
            $text = $(if ($l.Length -gt 1) { $l.Substring(1) } else { '' })
            if ($tag -eq '-' -and $adds.Count -eq 0) { $dels.Add($text); continue }
            if ($tag -eq '+') { $adds.Add($text); continue }
            if ($dels.Count -or $adds.Count) {
                [void]$body.Append((Get-ChangeBlockHtml $dels.ToArray() $adds.ToArray())); $dels.Clear(); $adds.Clear()
            }
            if ($tag -eq '-') { $dels.Add($text); continue }   # 「+ の後の -」は新しいまとまりの始まり
            if ($tag -eq 'END') { break }
            if ($SkipBlank -and $text -eq '') { continue }      # スライド・シート間の区切りの空行
            if ($l -match '^@@ (.+?) @@$') {                    # シート内の区切り（○行目付近）
                [void]$body.Append("<div class=`"ln hunk`"><span class=`"sg`">⋯</span><span class=`"tx`">$(ConvertTo-Html $Matches[1])</span></div>`n"); continue
            }
            [void]$body.Append("<div class=`"ln ctx`"><span class=`"sg`"> </span><span class=`"tx`">$(ConvertTo-Html $text)</span></div>`n")
        }
        [void]$body.Append("</div></section>`n")
    }
    return @($body.ToString(), ($toc -join "`n"), $secs.Count)
}

function New-DiffHtml($Old, $New, [string[]]$Diff, [string]$Ext, [int]$Added, [int]$Removed) {
    $fmt = 'yyyy-MM-dd HH:mm'
    $parts = Get-DiffHtmlBody $Diff ($Ext -ne '.docx')
    $sections = $parts[0]; $toc = $parts[1]; $count = $parts[2]
    $kind = @{ '.pptx' = '変更のあったスライド'; '.xlsx' = '変更のあったシート'; '.docx' = '変更箇所' }[$Ext]
    if ($Diff.Count -eq 0) {
        $main = '<p class="nochange">テキストの変更はありません（書式・画像のみの変更の可能性があります）。</p>'
        $nav = ''
    } else {
        $main = $sections
        $nav = "<nav><h3>$kind（$count）</h3><ol>$toc</ol></nav>"
    }
    $title = ConvertTo-Html "Office差分: $($New.Name)"
    return @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$title</title>
<style>
:root { --bg:#ffffff; --fg:#1f2328; --muted:#59636e; --line:#d1d9e0; --panel:#f6f8fa;
        --add:#dafbe1; --addm:#aceebb; --addfg:#116329; --del:#ffebe9; --delm:#ffcecb; --delfg:#a40e26;
        --chg:#0969da; --hunk:#ddf4ff; }
@media (prefers-color-scheme: dark) {
  :root { --bg:#0d1117; --fg:#e6edf3; --muted:#9198a1; --line:#3d444d; --panel:#151b23;
          --add:#12261e; --addm:#1f6f3b; --addfg:#7ee787; --del:#25171c; --delm:#8e1519; --delfg:#ff9492;
          --chg:#4493f8; --hunk:#121d2f; }
}
* { box-sizing: border-box; }
body { margin:0; background:var(--bg); color:var(--fg);
       font-family: "Yu Gothic UI","Meiryo UI","Hiragino Sans",system-ui,sans-serif; font-size:15px; line-height:1.6; }
header { padding:20px 24px 12px; border-bottom:1px solid var(--line); }
h1 { margin:0 0 10px; font-size:20px; }
table.meta { border-collapse:collapse; font-size:13px; }
table.meta th { text-align:left; color:var(--muted); font-weight:normal; padding:2px 12px 2px 0; white-space:nowrap; vertical-align:top; }
table.meta td { padding:2px 12px 2px 0; word-break:break-all; }
.stat { margin-top:10px; font-size:14px; }
.stat .a { color:var(--addfg); font-weight:bold; } .stat .d { color:var(--delfg); font-weight:bold; }
.wrap { display:flex; gap:24px; padding:16px 24px 40px; align-items:flex-start; }
nav { position:sticky; top:12px; flex:0 0 240px; max-height:calc(100vh - 24px); overflow:auto;
      background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:10px 12px; font-size:13px; }
nav h3 { margin:0 0 6px; font-size:13px; color:var(--muted); }
nav ol { margin:0; padding-left:0; list-style:none; }
nav li { margin:2px 0; }
nav a { color:var(--fg); text-decoration:none; display:block; padding:3px 4px; border-radius:4px; }
nav a:hover { background:var(--line); }
main { flex:1; min-width:0; }
section.sec { border:1px solid var(--line); border-radius:8px; margin-bottom:16px; overflow:hidden; scroll-margin-top:12px; }
section.sec h2 { margin:0; font-size:15px; padding:8px 12px; background:var(--panel); border-bottom:1px solid var(--line); }
.badge { display:inline-block; font-size:11px; font-weight:bold; padding:1px 7px; border-radius:10px; margin-right:8px;
         vertical-align:1px; color:#fff; background:var(--chg); }
.badge.del { background:var(--delfg); } .badge.add { background:var(--addfg); } .badge.hunk { background:var(--muted); }
@media (prefers-color-scheme: dark) { .badge { color:#0d1117; } }
.code { font-family: "BIZ UDGothic","MS Gothic",Consolas,monospace; font-size:14px; }
.ln { display:flex; }
.ln .sg { flex:0 0 26px; text-align:center; color:var(--muted); user-select:none; }
.ln .tx { flex:1; white-space:pre-wrap; word-break:break-word; padding-right:12px; }
.ln.add { background:var(--add); } .ln.add .sg { color:var(--addfg); }
.ln.del { background:var(--del); } .ln.del .sg { color:var(--delfg); }
.ln.add mark { background:var(--addm); color:inherit; border-radius:2px; }
.ln.del mark { background:var(--delm); color:inherit; border-radius:2px; text-decoration:line-through; }
.ln.hunk { background:var(--hunk); color:var(--muted); font-size:12px; }
.nochange { padding:16px; background:var(--panel); border-radius:8px; }
footer { color:var(--muted); font-size:12px; padding:0 24px 24px; }
@media (max-width: 760px) { .wrap { display:block; padding:12px; } nav { position:static; margin-bottom:12px; } }
@media print { nav { display:none; } .wrap { display:block; } section.sec { break-inside:avoid-page; } }
</style>
</head>
<body>
<header>
<h1>$title</h1>
<table class="meta">
<tr><th>変更前</th><td>$(ConvertTo-Html $Old.FullName)</td><td>$($Old.LastWriteTime.ToString($fmt))</td></tr>
<tr><th>変更後</th><td>$(ConvertTo-Html $New.FullName)</td><td>$($New.LastWriteTime.ToString($fmt))</td></tr>
</table>
<div class="stat">追加 <span class="a">+$Added</span> 行 ／ 削除 <span class="d">-$Removed</span> 行
$(if ($Diff.Count) { "　／　$kind <b>$count</b> 件" })</div>
</header>
<div class="wrap">
$nav
<main>
$main
</main>
</div>
<footer>作成: $((Get-Date).ToString($fmt))　Office差分ツール（デスクトップ版）　｜　<span style="background:var(--add)">緑＝追加</span>　<span style="background:var(--del)">赤＝削除</span>　濃い色＝行の中で変わった文字</footer>
</body>
</html>
"@
}

# ======================================================================= メイン

function Test-FileAssociation([string]$Ext) {
    # 「プログラムから開く」の選択画面が出ないよう、関連付けがあるかを事前に確認する
    try {
        $user = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Ext\UserChoice")
        if ($user -and $user.GetValue('ProgId')) { return $true }
        $cls = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($Ext)
        if ($cls -and $cls.GetValue('')) { return $true }
    } catch { }
    return $false
}

function Show-Usage {
    Write-Host ''
    Write-Host '使い方: 比較したい Word / PowerPoint / Excel ファイルを OfficeDiff.bat にドラッグ＆ドロップしてください。'
    Write-Host '        ・2つまとめてドロップ → すぐに比較'
    Write-Host '        ・1つだけドロップ     → もう1つをファイル選択画面で選ぶ（別フォルダの同名ファイルも可）'
    Write-Host '        ・ダブルクリック       → 2つともファイル選択画面で選ぶ'
    Write-Host '        （更新日時が古い方が「変更前」、新しい方が「変更後」になります）'
}

function Select-OfficeFile([string]$Title, [string]$StartDir, [string]$Ext) {
    # テスト用: 環境変数 OFFICEDIFF_PICK にパスを ; 区切りで入れておくと、画面を出さずに順に使う
    if ($env:OFFICEDIFF_PICK) {
        $list = @($env:OFFICEDIFF_PICK.Split(';') | Where-Object { $_ })
        $env:OFFICEDIFF_PICK = (@($list | Select-Object -Skip 1) -join ';')
        return $list[0]
    }
    if ($env:OFFICEDIFF_NODIALOG) { return $null }
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $Title
    if ($Ext -eq '.pptx') { $dlg.Filter = 'PowerPoint (*.pptx)|*.pptx' }
    elseif ($Ext -eq '.docx') { $dlg.Filter = 'Word (*.docx)|*.docx' }
    elseif ($Ext -eq '.xlsx' -or $Ext -eq '.xlsm') { $dlg.Filter = 'Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm' }
    else { $dlg.Filter = 'Word / PowerPoint / Excel (*.docx;*.pptx;*.xlsx;*.xlsm)|*.docx;*.pptx;*.xlsx;*.xlsm' }
    if ($StartDir -and (Test-Path -LiteralPath $StartDir)) { $dlg.InitialDirectory = $StartDir }
    # 黒い画面の裏に隠れないよう、最前面のダミーウィンドウを親にする
    $owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true; ShowInTaskbar = $false }
    try {
        if ($dlg.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
        return $null
    } finally { $owner.Dispose(); $dlg.Dispose() }
}

# テスト用: . (ドット) で読み込んだときは関数定義だけ行い、処理は実行しない
if ($MyInvocation.InvocationName -eq '.') { return }

try {
    $files = @($Paths | Where-Object { $_ } | ForEach-Object { $_.Trim('"') })
    if ($files.Count -gt 2) {
        Write-Host "ファイルは2つまでにしてください（今回: $($files.Count) 個）" -ForegroundColor Yellow
        Show-Usage; exit 1
    }
    # 足りない分はファイル選択画面で選ぶ（別フォルダにある同じ名前のファイル同士を比べられるように）
    if ($files.Count -eq 0) {
        Write-Host '1つ目のファイルを選んでください...'
        $f = Select-OfficeFile '比較する1つ目のファイルを選んでください' ([Environment]::GetFolderPath('Desktop')) ''
        if ($f) { $files += $f }
    }
    if ($files.Count -eq 1) {
        $first = Get-Item -LiteralPath $files[0]
        Write-Host "1つ目: $($first.FullName)"
        Write-Host '比較するもう1つのファイルを選んでください...'
        $f = Select-OfficeFile "「$($first.Name)」と比較するファイルを選んでください" $first.DirectoryName $first.Extension.ToLower()
        if ($f) { $files += $f }
    }
    if ($files.Count -ne 2) {
        Write-Host 'ファイルの選択がキャンセルされました。' -ForegroundColor Yellow
        Show-Usage; exit 1
    }
    if ((Get-Item -LiteralPath $files[0]).FullName -eq (Get-Item -LiteralPath $files[1]).FullName) {
        throw '同じファイルが2回選ばれています。別のファイルを選んでください'
    }
    $items = @($files | ForEach-Object { Get-Item -LiteralPath $_ })
    # .xlsm（マクロ付き）は .xlsx と同じ扱い
    $kindOf = { param($e) $e = $e.ToLower(); if ($e -eq '.xlsm') { '.xlsx' } else { $e } }
    $ext = & $kindOf $items[0].Extension
    if ($ext -ne (& $kindOf $items[1].Extension)) { throw "同じ種類のファイルを指定してください: $($items[0].Name) / $($items[1].Name)" }
    if (@('.docx', '.pptx', '.xlsx') -notcontains $ext) {
        throw "対応しているのは .docx / .pptx / .xlsx だけです（$($items[0].Extension) は未対応。.doc/.ppt/.xls は Office で新しい形式に保存し直してください）"
    }

    # 更新日時の古い方を「変更前」に
    $items = @($items | Sort-Object LastWriteTime)
    if ($Swap) { $items = @($items[1], $items[0]) }
    $old = $items[0]; $new = $items[1]

    $zo = Open-OfficeZip $old.FullName
    $zn = Open-OfficeZip $new.FullName
    try {
        if ($ext -eq '.pptx') {
            $diff = Get-SlideDiffLines (Get-Slides $zo) (Get-Slides $zn)
        } elseif ($ext -eq '.xlsx') {
            $diff = Get-SheetDiffLines (Get-XlsxSheets $zo) (Get-XlsxSheets $zn)
        } else {
            $diff = Get-UnifiedLines (Get-WordLines $zo) (Get-WordLines $zn)
        }
    } finally { $zo.Dispose(); $zn.Dispose() }

    $added = @($diff | Where-Object { $_.StartsWith('+') -and -not $_.StartsWith('+## ') }).Count
    $removed = @($diff | Where-Object { $_.StartsWith('-') -and -not $_.StartsWith('-## ') }).Count

    # 出力は .diff（unified diff 形式）。先頭の説明行は diff の書式上「ヘッダーのコメント」として扱われる
    $stamp = { param($f) $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm') }
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("Office差分: $($new.Name)")
    $out.Add("変更前: $($old.FullName)  ($(& $stamp $old))")
    $out.Add("変更後: $($new.FullName)  ($(& $stamp $new))")
    $out.Add("作成: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))  追加 +$added 行 / 削除 -$removed 行")
    $out.Add('')
    if ($diff.Count -eq 0) {
        $out.Add('テキストの変更はありません（書式・画像のみの変更の可能性があります）。')
    } else {
        $out.Add("--- $($old.FullName)`t$(& $stamp $old)")
        $out.Add("+++ $($new.FullName)`t$(& $stamp $new)")
        foreach ($l in $diff) { $out.Add($l) }
    }

    # 出力先: 指定が無ければ、このツールと同じフォルダの output フォルダ（無ければ作る）
    $dir = $(if ($OutDir) { $OutDir } else { Join-Path $PSScriptRoot 'output' })
    try {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        # 書き込めるか確認（Program Files など書き込み禁止の場所に置かれた場合に備える）
        $probe = Join-Path $dir ('.write-test-' + [Guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, ''); Remove-Item -LiteralPath $probe -Force
    } catch {
        Write-Host "出力先フォルダに書き込めないため、変更後のファイルと同じフォルダに出力します: $dir" -ForegroundColor Yellow
        $dir = $new.DirectoryName
    }
    $base = Join-Path $dir ('{0}_差分_{1}' -f [IO.Path]::GetFileNameWithoutExtension($new.Name), (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $outPath = "$base.diff"
    [System.IO.File]::WriteAllText($outPath, (($out -join "`r`n") + "`r`n"), $utf8)
    # ブラウザで見る用の HTML（外部ファイルやネット接続を使わない1ファイル完結）
    $htmlPath = "$base.html"
    [System.IO.File]::WriteAllText($htmlPath, (New-DiffHtml $old $new $diff $ext $added $removed), $utf8)

    Write-Host ''
    Write-Host "変更前: $($old.Name)  ($($old.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))"
    Write-Host "変更後: $($new.Name)  ($($new.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))"
    Write-Host "追加 +$added 行 / 削除 -$removed 行"
    Write-Host ''
    Write-Host "出力しました: $htmlPath  （ブラウザで表示）" -ForegroundColor Green
    Write-Host "              $outPath  （テキスト版）" -ForegroundColor Green
    Write-Host '（前後が逆なら OfficeDiff_入れ替え.bat にドロップし直してください）'
    if (-not $NoOpen -and -not $env:OFFICEDIFF_NOOPEN) {
        # HTML を既定のブラウザで開く（.html はどのPCでも必ずブラウザに関連付けられている）
        try { Invoke-Item -LiteralPath $htmlPath } catch { Write-Host "ブラウザで開けませんでした。上の .html ファイルをダブルクリックしてください" -ForegroundColor Yellow }
    }
    exit 0
} catch {
    Write-Host ''
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
