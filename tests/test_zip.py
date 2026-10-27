"""配布用 zip（scripts/make_zip.py）のテスト。"""
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_配布zipにツール一式とREADMEと画像が入る(tmp_path):
    out = tmp_path / 'OfficeDiff.zip'
    subprocess.run([sys.executable, str(ROOT / 'scripts' / 'make_zip.py'), str(out)], check=True, capture_output=True)
    infos = zipfile.ZipFile(out).infolist()
    names = {i.filename for i in infos}
    for n in ('OfficeDiff.bat', 'OfficeDiff_入れ替え.bat', 'OfficeDiff.ps1', 'README.md', '使い方.txt', 'images/sample-pptx.png'):
        assert f'OfficeDiff/{n}' in names
    # 開発用のファイルは入れない
    assert not any(n.startswith(('OfficeDiff/tests', 'OfficeDiff/scripts', 'OfficeDiff/.github')) for n in names)
    # 日本語ファイル名は UTF-8 フラグ付き（Windows で文字化けしない）
    assert all(i.flag_bits & 0x800 for i in infos if not i.filename.isascii())
    # README から参照している画像が zip に入っていること
    readme = (ROOT / 'README.md').read_text(encoding='utf-8')
    for img in re.findall(r'!\[[^\]]*\]\(([^)]+)\)', readme):
        assert f'OfficeDiff/{img}' in names, img
