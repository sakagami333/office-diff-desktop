# CLAUDE.md

Windows デスクトップ用の Office 差分ツール。2つの .docx / .pptx を `OfficeDiff.bat` にドロップすると、
変更点を `output/` フォルダに .html（ブラウザ表示）と .diff で出力する。

## 構成

```
OfficeDiff.ps1              本体（Windows PowerShell 5.1、標準機能のみ・Office不要。.docx/.pptx の XML を直接読む）
OfficeDiff.bat              ドロップ先（-STA で起動。ファイル選択画面に必要）
OfficeDiff_入れ替え.bat     変更前／変更後を逆にして比較（-Swap）
README.md / images/         利用者向け説明（画像付き）
使い方.txt                  同じ内容のテキスト版（UTF-8 BOM・CRLF。メモ帳用）
scripts/make_zip.py         配布用 OfficeDiff.zip を作る
tests/make_samples.py       テスト用サンプル資料（v1/v2）を生成
tests/test_*.py             pytest（OFFICEDIFF_PS=powershell で 5.1 を強制）
tests/check_bat_windows.ps1 Windows CI 用: .bat 経由・日本語パスの確認
```

## ルール

- 言語: コード内コメント・メッセージ・ドキュメントは日本語
- `OfficeDiff.ps1` と `tests/check_bat_windows.ps1` は **UTF-8（BOM付き）** で保存する（PS 5.1 が日本語を読むため）
- `.bat` は ASCII のみ・CRLF（コンソールの文字コードで化けないように）
- Windows PowerShell 5.1 で動くこと（`Sort-Object -Stable` など 7 専用の機能は使わない）。CI の Windows ジョブで確認する
- PowerShell の変数名は大文字小文字を区別しない（`$NS` と `$ns` は同じ変数）点に注意
- 使い方を変えたら README.md と 使い方.txt の両方をそろえる
- 変更したら `pytest -q` を通し、push 後に GitHub Actions の Windows ジョブが成功することを確認する
- テキスト化・差分のルールは office-diff-sample（PR コメント版）の spec/SPEC.md「F3 テキスト化ルール」「F4」に合わせる
