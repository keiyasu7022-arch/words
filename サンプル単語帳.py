"""
サンプルExcelファイル生成スクリプト
実行: python3 サンプル単語帳.py
"""
import subprocess, sys

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

words = [
    ("りんご", "apple"), ("犬", "dog"), ("猫", "cat"), ("学校", "school"),
    ("先生", "teacher"), ("本", "book"), ("水", "water"), ("空", "sky"),
    ("山", "mountain"), ("海", "sea"), ("電車", "train"), ("友達", "friend"),
    ("家族", "family"), ("食べる", "eat"), ("飲む", "drink"), ("走る", "run"),
    ("読む", "read"), ("書く", "write"), ("聞く", "listen"), ("話す", "speak"),
    ("大きい", "big"), ("小さい", "small"), ("速い", "fast"), ("遅い", "slow"),
    ("赤", "red"), ("青", "blue"), ("白", "white"), ("黒", "black"),
    ("今日", "today"), ("明日", "tomorrow"),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "単語帳"
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].width = 20

for row in words:
    ws.append(row)

wb.save("サンプル単語帳.xlsx")
print(f"✅ サンプル単語帳.xlsx を作成しました（{len(words)}語）")
