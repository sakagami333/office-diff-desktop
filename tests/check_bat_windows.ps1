# CI（Windows）用: OfficeDiff.bat を日本語・空白入りのパスで実行し、.diff が出力されるか確認する
# Windows PowerShell 5.1 で日本語を正しく読むため、このファイルは UTF-8（BOM付き）で保存すること
$ErrorActionPreference = 'Stop'

function Fail([string]$msg) {
    # ジョブログが読めない環境でも原因がわかるよう、注釈（annotation）として出す
    Write-Host "::error::$msg"
    exit 1
}

Write-Host "PowerShell $($PSVersionTable.PSVersion)"
$root = Split-Path -Parent $PSScriptRoot
$work = Join-Path $env:RUNNER_TEMP 'サンプル 比較'
$before = Join-Path $work '変更前'
$after = Join-Path $work '変更後'
python (Join-Path $root 'tests\make_samples.py') $before v1 | Out-Null
python (Join-Path $root 'tests\make_samples.py') $after v2 | Out-Null
$oldFile = Join-Path $before '研修案内.pptx'
$newFile = Join-Path $after '研修案内.pptx'
if (-not (Test-Path -LiteralPath $oldFile)) { Fail "サンプルが作成されていません: $oldFile" }
(Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-7)

# ドラッグ＆ドロップと同じく、エクスプローラーが渡す形（"パス" "パス"）で bat を起動。
# pause には空行を流し込む。新しい方を先に渡し、日時で並べ替わることも確認する
$bat = Join-Path $root 'OfficeDiff.bat'
$outDir = Join-Path $root 'output'
if (Test-Path -LiteralPath $outDir) { Remove-Item -LiteralPath $outDir -Recurse -Force }
$cmdLine = "echo.| `"$bat`" `"$newFile`" `"$oldFile`""
$env:OFFICEDIFF_NOOPEN = '1'
$console = cmd /c $cmdLine 2>&1 | Out-String
Write-Host $console

if (-not (Test-Path -LiteralPath $outDir)) { Fail "出力先 output フォルダが作られていません: $outDir" }
if (Get-ChildItem -LiteralPath $after -Filter '*_差分_*') { Fail '元のファイルのフォルダに出力されています' }
$md = Get-ChildItem -LiteralPath $outDir -Filter '*.diff' | Select-Object -First 1
if (-not $md) {
    $flat = ($console -replace "`r?`n", ' | ')
    Fail ".diff が出力されていません。bat の出力: $flat"
}
$text = [System.IO.File]::ReadAllText($md.FullName, [System.Text.Encoding]::UTF8)
Write-Host $text
foreach ($expect in @('-## スライド 4（削除）', ' ## スライド 5 → 4', '+  ・Webスクレイピング入門')) {
    if (-not $text.Contains($expect)) { Fail "期待した行がありません: [$expect]  出力ファイル: $($md.Name)" }
}
$html = Get-ChildItem -LiteralPath $outDir -Filter '*.html' | Select-Object -First 1
if (-not $html) { Fail '.html が出力されていません' }
$h = [System.IO.File]::ReadAllText($html.FullName, [System.Text.Encoding]::UTF8)
if (-not $h.Contains('スライド 4（削除）') -or -not $h.Contains('<mark>')) { Fail ".html の内容が期待と違います: $($html.Name)" }
Remove-Item -LiteralPath $html.FullName
Write-Host "::notice::.bat 経由の実行 OK（$($md.Name) / .html）"

# 1つだけドロップ → もう1つは選択画面（CIでは OFFICEDIFF_PICK で代用）。別フォルダの同名ファイル
Remove-Item -LiteralPath $md.FullName
$env:OFFICEDIFF_PICK = $oldFile
$console = cmd /c "echo.| `"$bat`" `"$newFile`"" 2>&1 | Out-String
$env:OFFICEDIFF_PICK = ''
Write-Host $console
$md = Get-ChildItem -LiteralPath $outDir -Filter '*.diff' | Select-Object -First 1
if (-not $md) { Fail ("1つだけドロップした場合に .diff が出力されていません: " + ($console -replace "`r?`n", ' | ')) }
$text = [System.IO.File]::ReadAllText($md.FullName, [System.Text.Encoding]::UTF8)
if (-not $text.Contains('-## スライド 4（削除）')) { Fail '1つだけドロップした場合の差分が期待と違います' }
Write-Host "::notice::1つだけドロップ → 選択 の流れも OK"
