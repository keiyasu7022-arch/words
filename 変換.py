#!/usr/bin/env python3
"""ExcelをJSONに変換してNetlifyにデプロイできる形にする"""
import json, os, sys

# --silent: 起動.commandからのバックグラウンド実行時は出力なし・確認なし
SILENT = '--silent' in sys.argv

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))

# サブフォルダも含めて再帰スキャン、相対パスで収集
files = []
for root, dirs, fs in os.walk(BASE):
    dirs[:] = [d for d in dirs if not d.startswith('.')]
    for f in sorted(fs):
        if f.lower().endswith(('.xlsx', '.xls')):
            files.append(os.path.join(root, f))
files = sorted(files)

if not files:
    if not SILENT:
        print('❌ Excelファイルが見つかりません')
        print('   .xlsx ファイルをこのフォルダ（またはサブフォルダ）に入れてから実行してください')
        try:
            input('\nEnterキーで閉じる')
        except EOFError:
            pass
    sys.exit(1)

all_data = {}
for f in files:
    rel = os.path.relpath(f, BASE)
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb.active
    words = []
    for row in ws.iter_rows(values_only=True):
        ja = str(row[0] if row[0] is not None else '').strip()
        en = str(row[1] if row[1] is not None else '').strip()
        if ja and en:
            words.append({'ja': ja, 'en': en})
    wb.close()
    if words:
        all_data[rel] = words
        if not SILENT:
            print(f'✅ {rel}: {len(words)}語')

out = os.path.join(BASE, 'words.json')
with open(out, 'w', encoding='utf-8') as fp:
    json.dump(all_data, fp, ensure_ascii=False, indent=2)

if SILENT:
    sys.exit(0)

print(f'\n📄 words.json を作成しました（{len(all_data)}ファイル）')
print()
print('━' * 40)
print('次の手順:')
print('  1. https://app.netlify.com/drop をブラウザで開く')
print('  2. 「単語帳アプリ」フォルダをそのままドラッグ')
print('  3. 発行されたURLをスマホで開く')
print('━' * 40)
try:
    input('\nEnterキーで閉じる')
except EOFError:
    pass
