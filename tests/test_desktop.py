"""OfficeDiff.ps1 のテスト。PowerShell が無い環境ではスキップ。"""
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
# OFFICEDIFF_PS=powershell で Windows PowerShell 5.1 を強制できる（CIのWindowsジョブで使用）
PS = os.environ.get('OFFICEDIFF_PS') or shutil.which('pwsh') or shutil.which('powershell')
pytestmark = pytest.mark.skipif(PS is None, reason='PowerShell が無い環境')


@pytest.fixture
def samples(tmp_path):
    for v in ('v1', 'v2'):
        subprocess.run([sys.executable, str(ROOT / 'tests' / 'make_samples.py'), str(tmp_path / v), v],
                       check=True, capture_output=True)
    # v1 を古く、v2 を新しくする（ツールは更新日時で前後を決める）
    for f in (tmp_path / 'v1').iterdir():
        os.utime(f, (1_700_000_000, 1_700_000_000))
    for f in (tmp_path / 'v2').iterdir():
        os.utime(f, (1_800_000_000, 1_800_000_000))
    return tmp_path


OUT = 'out'  # テストでは -OutDir <tmp>/out に出力する（リポジトリ内の output を汚さない）


def run_tool(*args, pick=None, outdir=None, script=None):
    # pick: ファイル選択画面の代わりに使うパスのリスト。None なら画面を出さない（キャンセル扱い）
    # outdir: None なら -OutDir を付けない（既定の output フォルダに出る）
    env = dict(os.environ, OFFICEDIFF_NODIALOG='1')
    if pick:
        env['OFFICEDIFF_PICK'] = ';'.join(map(str, pick))
    extra = ['-OutDir', str(outdir)] if outdir else []
    res = subprocess.run([PS, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                          str(script or ROOT / 'OfficeDiff.ps1'), *map(str, args), '-NoOpen', *extra],
                         capture_output=True, text=True, encoding='utf-8', env=env)
    return res


def output_diff(folder):
    files = list(Path(folder).glob('*_差分_*.diff'))
    assert len(files) == 1, files
    return files[0].read_text(encoding='utf-8')


def test_pptx_新しい方を先に渡しても日時で前後が決まる(samples):
    res = run_tool(samples / 'v2' / '研修案内.pptx', samples / 'v1' / '研修案内.pptx', outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    md = output_diff(samples / OUT)
    assert md.startswith('Office差分: 研修案内.pptx')
    # unified diff のファイル見出し（変更前 → 変更後）
    assert re.search(r'^--- .*v1.研修案内\.pptx\t', md, re.M) and re.search(r'^\+\+\+ .*v2.研修案内\.pptx\t', md, re.M)
    assert '-## スライド 4（削除）' in md
    assert ' ## スライド 5 → 4' in md
    assert '+  ・Webスクレイピング入門' in md


def test_docx(samples):
    res = run_tool(samples / 'v1' / '研修運営マニュアル.docx', samples / 'v2' / '研修運営マニュアル.docx', outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    md = output_diff(samples / OUT)
    assert '+研修開始の5営業日前までに、以下を準備します。' in md
    assert '+| 9:45 | オリエンテーション |' in md
    assert '+・受講者への事前案内メール送信' in md


def test_Swapで前後を入れ替え(samples):
    res = run_tool('-Swap', samples / 'v1' / '研修案内.pptx', samples / 'v2' / '研修案内.pptx', outdir=samples / OUT)
    assert res.returncode == 0
    md = output_diff(samples / OUT)
    assert '+## スライド 4（追加）' in md


def test_同じ内容なら変更なし(samples, tmp_path):
    import shutil as sh
    copy = tmp_path / 'copy' / '研修案内.pptx'
    copy.parent.mkdir()
    sh.copy(samples / 'v1' / '研修案内.pptx', copy)
    res = run_tool(samples / 'v1' / '研修案内.pptx', copy, outdir=samples / OUT)
    assert res.returncode == 0
    md = list((samples / OUT).glob('*_差分_*.diff'))
    assert len(md) == 1 and 'テキストの変更はありません' in md[0].read_text(encoding='utf-8')


def test_エラー_種類違いと個数(samples):
    res = run_tool(samples / 'v1' / '研修案内.pptx', samples / 'v2' / '研修運営マニュアル.docx', outdir=samples / OUT)
    assert res.returncode == 1 and '同じ種類' in res.stdout
    res = run_tool(samples / 'v1' / '研修案内.pptx', outdir=samples / OUT)  # 1つだけで選択画面をキャンセル
    assert res.returncode == 1 and 'キャンセル' in res.stdout


def test_1つだけドロップしてもう1つを選択画面で選ぶ_別フォルダの同名ファイル(samples):
    res = run_tool(samples / 'v2' / '研修案内.pptx', pick=[samples / 'v1' / '研修案内.pptx'], outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    assert '-## スライド 4（削除）' in output_diff(samples / OUT)


def test_ダブルクリックで2つとも選択画面で選ぶ(samples):
    res = run_tool(pick=[samples / 'v1' / '研修運営マニュアル.docx', samples / 'v2' / '研修運営マニュアル.docx'], outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    assert '+・受講者への事前案内メール送信' in output_diff(samples / OUT)


def test_同じファイルを2回選ぶとエラー(samples):
    f = samples / 'v1' / '研修案内.pptx'
    res = run_tool(f, pick=[f], outdir=samples / OUT)
    assert res.returncode == 1 and '同じファイルが2回' in res.stdout


def output_html(folder):
    files = list(Path(folder).glob('*_差分_*.html'))
    assert len(files) == 1, files
    return files[0].read_text(encoding='utf-8')


def test_HTMLレポート_pptx(samples):
    res = run_tool(samples / 'v1' / '研修案内.pptx', samples / 'v2' / '研修案内.pptx', outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    html = output_html(samples / OUT)
    assert html.startswith('<!DOCTYPE html>') and '<meta charset="utf-8">' in html
    # 外部ファイル・ネット接続を使わない（社内PCでもそのまま開ける）
    assert not re.search(r'(src|href)="https?://', html)
    assert html.count('<section ') == 5                  # 変更のあったスライド5件
    assert '<span class="badge del">削除</span>スライド 4（削除）' in html
    assert 'スライド 5 → 4' in html
    # 行内で変わった文字だけ <mark>（似た行同士を組にする: ノート行どうし）
    assert '<mark>／オンライン併用</mark>' in html
    assert '<mark>（3年目まで拡大）</mark>' in html


def test_HTMLレポート_docx_見出しで変更箇所に名前(samples):
    res = run_tool(samples / 'v1' / '研修運営マニュアル.docx', samples / 'v2' / '研修運営マニュアル.docx', outdir=samples / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    html = output_html(samples / OUT)
    for t in ('2. 事前準備 付近', '3. 当日の流れ 付近', '4. 事後対応 付近'):
        assert t in html
    assert '研修開始の<mark>3</mark>営業日前' in html and '研修開始の<mark>5</mark>営業日前' in html


def test_HTMLはエスケープされる(samples, tmp_path):
    # 資料中の < > & がタグとして解釈されないこと
    import docx
    for name, text in (('a', '比較 A < B'), ('b', '比較 A < B & <script>x</script>')):
        d = docx.Document()
        d.add_paragraph(text)
        (tmp_path / name).mkdir()
        d.save(tmp_path / name / 'e.docx')
    os.utime(tmp_path / 'a' / 'e.docx', (1_700_000_000, 1_700_000_000))
    res = run_tool(tmp_path / 'a' / 'e.docx', tmp_path / 'b' / 'e.docx', outdir=tmp_path / OUT)
    assert res.returncode == 0, res.stdout + res.stderr
    html = output_html(tmp_path / OUT)
    assert '<script>' not in html and '&lt;script&gt;' in html


def test_既定の出力先はツールと同じフォルダのoutput_無ければ作る(samples, tmp_path):
    tool = tmp_path / 'OfficeDiff'
    tool.mkdir()
    shutil.copy(ROOT / 'OfficeDiff.ps1', tool / 'OfficeDiff.ps1')
    assert not (tool / 'output').exists()
    res = run_tool(samples / 'v1' / '研修案内.pptx', samples / 'v2' / '研修案内.pptx', script=tool / 'OfficeDiff.ps1')
    assert res.returncode == 0, res.stdout + res.stderr
    assert '-## スライド 4（削除）' in output_diff(tool / 'output')
    output_html(tool / 'output')
    # 元のファイルのフォルダには何も出力しない
    assert not list((samples / 'v2').glob('*_差分_*'))
    # 2回目（output が既にある）も同じ場所に追加される
    res = run_tool(samples / 'v1' / '研修運営マニュアル.docx', samples / 'v2' / '研修運営マニュアル.docx', script=tool / 'OfficeDiff.ps1')
    assert res.returncode == 0
    assert len(list((tool / 'output').glob('*.html'))) == 2
