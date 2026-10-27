#!/usr/bin/env python3
"""配布用 zip（OfficeDiff.zip）を作る。

使い方: python scripts/make_zip.py [出力先.zip]
中身は zip 内の OfficeDiff/ フォルダにまとめる。日本語ファイル名は UTF-8 フラグ付きで格納し、
Windows のエクスプローラーで展開しても文字化けしないようにする。
"""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT
FILES = [
    'OfficeDiff.bat',
    'OfficeDiff_入れ替え.bat',
    'OfficeDiff.ps1',
    'README.md',
    '使い方.txt',
    'images/sample-pptx.png',
    'images/sample-xlsx.png',
]


def main(out):
    out = Path(out)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in FILES:
            path = SRC / name
            if not path.exists():
                raise SystemExit(f'ファイルがありません: {path}')
            z.write(path, f'OfficeDiff/{name}')
    print(f'作成しました: {out}')
    for info in zipfile.ZipFile(out).infolist():
        print(f'  {info.file_size:>8}  {info.filename}')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else ROOT / 'OfficeDiff.zip')
