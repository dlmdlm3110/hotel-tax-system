#!/usr/bin/env python3
"""配布 ZIP ビルドスクリプト（max / ねっぱんvr / Airbnbvr）"""
import ast, re, zipfile
from pathlib import Path

SRC  = Path('/Users/macoro/Desktop/tax')
DIST = Path('/Users/macoro/Desktop')

# ── ZIP に含めるファイル ──────────────────────────────────────
ALWAYS = [
    'requirements.txt',
    'launch.command', 'launch.bat',
    'setup.command',  'setup.bat',
    '_make_icon.py',          # setup.bat が参照するため必須
    'create_shortcut.py',     # Windowsデスクトップショートカット作成（OneDrive対応）
    'README.txt',
    'セットアップ手順_Mac.txt',
    'セットアップ手順_Windows.txt',
    '宿泊税スマートロゴ.png',
    'applogo.png', 'appicon.ico', 'logo.svg',
    '宿泊税月計表テンプレ.xlsx',
    '宿泊税納入書テンプレ.xlsx',
    '納入申告書テンプレ.xlsx',
    '領収書タイプA税込.xlsx',
    '領収書タイプB税別途.xlsx',
    '領収書タイプC税全コミ.xlsx',
    '領収書タイプD宿泊税（税のみ）.xlsx',
]

# ── アップロードループの置換文字列 ───────────────────────────
LOOP_ORIG = (
    "                _has_airbnb     = False\n"
    "                _airbnb_raw_csv = None\n"
    "                for _uf in _ufs:\n"
    "                    _raw = _uf.read()\n"
    "                    _fmt = detect_csv_format(_raw)\n"
    "                    if _fmt == 'airbnb':\n"
    "                        _all_items  += parse_airbnb_csv(_raw, y, m)\n"
    "                        _has_airbnb = True\n"
    "                        if _airbnb_raw_csv is None:\n"
    "                            _airbnb_raw_csv = _raw\n"
    "                        _fmt_labels.append('Airbnb')\n"
    "                    elif _fmt == 'hotel_total':\n"
    "                        _all_items  += parse_hotel_total_csv(_raw, y, m)\n"
    "                        if _raw_csv is None:\n"
    "                            _raw_csv    = _raw\n"
    "                            _csv_format = 'hotel_total'\n"
    "                        _fmt_labels.append('ホテルシステム(集計)')\n"
    "                    else:\n"
    "                        _all_items  += parse_hotel_csv(_raw, y, m)\n"
    "                        if _raw_csv is None:\n"
    "                            _raw_csv    = _raw\n"
    "                            _csv_format = 'hotel'\n"
    "                        _fmt_labels.append('ホテルシステム')"
)

LOOP_NB = (
    "                _has_airbnb     = False\n"
    "                _airbnb_raw_csv = None\n"
    "                for _uf in _ufs:\n"
    "                    _raw = _uf.read()\n"
    "                    _fmt = detect_csv_format(_raw)\n"
    "                    if _fmt == 'hotel_total':\n"
    "                        _all_items  += parse_hotel_total_csv(_raw, y, m)\n"
    "                        if _raw_csv is None:\n"
    "                            _raw_csv    = _raw\n"
    "                            _csv_format = 'hotel_total'\n"
    "                        _fmt_labels.append('ホテルシステム(集計)')\n"
    "                    else:\n"
    "                        _all_items  += parse_hotel_csv(_raw, y, m)\n"
    "                        if _raw_csv is None:\n"
    "                            _raw_csv    = _raw\n"
    "                            _csv_format = 'hotel'\n"
    "                        _fmt_labels.append('ホテルシステム')"
)

LOOP_AB = (
    "                _has_airbnb     = True\n"
    "                _airbnb_raw_csv = None\n"
    "                for _uf in _ufs:\n"
    "                    _raw = _uf.read()\n"
    "                    _all_items  += parse_airbnb_csv(_raw, y, m)\n"
    "                    if _airbnb_raw_csv is None:\n"
    "                        _airbnb_raw_csv = _raw\n"
    "                    _fmt_labels.append('Airbnb')"
)

# ── ヘルパー関数 ─────────────────────────────────────────────
def remove_functions(src: str, names: set) -> str:
    tree = ast.parse(src)
    lines = src.splitlines(keepends=True)
    to_del = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in names:
            to_del.append((node.lineno - 1, node.end_lineno))
    for start, end in sorted(to_del, reverse=True):
        del lines[start:end]
    return ''.join(lines)


def remove_constant(src: str, name: str) -> str:
    return '\n'.join(
        line for line in src.splitlines()
        if not re.match(rf'^{re.escape(name)}\s*=', line)
    ) + '\n'


def build_app(variant: str) -> str:
    src = (SRC / 'app.py').read_text(encoding='utf-8')
    if variant == 'nb':
        src = remove_functions(src, {'parse_airbnb_csv'})  # detect_csv_formatはhotel/hotel_total判別に必要なので残す
        assert LOOP_ORIG in src, 'ループブロックが見つかりません（ねっぱんvr）'
        src = src.replace(LOOP_ORIG, LOOP_NB, 1)
    elif variant == 'ab':
        src = remove_functions(src, {
            'detect_csv_format', 'parse_hotel_csv', 'parse_hotel_total_csv',
            'parse_hotel_csv_for_verification'
        })
        src = remove_constant(src, 'AGODA_COMMISSION')
        assert LOOP_ORIG in src, 'ループブロックが見つかりません（Airbnbvr）'
        src = src.replace(LOOP_ORIG, LOOP_AB, 1)
    ast.parse(src)  # 構文チェック
    return src


def add_file(zf: zipfile.ZipFile, fpath: Path, arcname: str):
    """
    ファイルをZIPに追加する。
    .bat は UTF-8 → CP932 変換 + LF→CRLF でWindowsのCMDに対応。
    """
    suffix = fpath.suffix.lower()
    if suffix == '.bat':
        text = fpath.read_text(encoding='utf-8')
        text = text.replace('\r\n', '\n').replace('\n', '\r\n')  # CRLF統一
        zf.writestr(arcname, text.encode('cp932', errors='replace'))
    elif suffix in ('.command', '.py', '.txt', '.md'):
        # テキスト系はUTF-8のまま（Mac/Linuxはそのまま動く）
        zf.write(fpath, arcname)
    else:
        # バイナリ（xlsx, png, ico, svg）はそのままコピー
        zf.write(fpath, arcname)


def make_zip(variant: str, folder: str, zip_name: str):
    print(f'\n▶ {zip_name} 作成中...')
    app_code = build_app(variant)

    zip_path = DIST / zip_name
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # app.py は UTF-8 のまま（Streamlit / Python はUTF-8を期待）
        zf.writestr(f'{folder}/app.py', app_code.encode('utf-8'))

        for fname in ALWAYS:
            fpath = SRC / fname
            if fpath.exists():
                add_file(zf, fpath, f'{folder}/{fname}')
                print(f'  + {fname}')
            else:
                print(f'  - {fname} （スキップ・ファイルなし）')

    size_kb = zip_path.stat().st_size // 1024
    print(f'✅ {zip_name}  {size_kb:,} KB')


# ── ビルド実行 ──────────────────────────────────────────────
make_zip('max', 'tax', '宿泊税スマートmax.zip')
make_zip('nb',  'nb',  '宿泊税スマートねっぱんvr.zip')
make_zip('ab',  'ab',  '宿泊税スマートAirbnbvr.zip')

print('\n\n========================================')
print('  全 3 ZIP の作成が完了しました。')
print('  デスクトップに保存されています。')
print('========================================')
