#!/usr/bin/env python3
"""ExcelをJSONに変換してNetlifyにデプロイできる形にする"""
import openpyxl, json, glob, os, sys

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))

files = sorted(
    glob.glob(os.path.join(BASE, '*.xlsx')) +
    glob.glob(os.path.join(BASE, '*.xls'))
)

if not files:
    print('❌ Excelファイルが見つかりません')
    print('   .xlsx ファイルをこのフォルダに入れてから実行してください')
    try:
    input('\nEnterキーで閉じる')
except EOFError:
    pass
    sys.exit(1)

all_data = {}
for f in files:
    name = os.path.basename(f)
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
        all_data[name] = words
        print(f'✅ {name}: {len(words)}語')
    else:
        print(f'⚠️  {name}: 単語が見つかりませんでした')

out = os.path.join(BASE, 'words.json')
with open(out, 'w', encoding='utf-8') as fp:
    json.dump(all_data, fp, ensure_ascii=False, indent=2)

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
