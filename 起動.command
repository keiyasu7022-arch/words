#!/bin/bash
cd "$(dirname "$0")"
# 起動時にwords.jsonを自動更新してからサーバーを起動
python3 変換.py --silent 2>/dev/null &
python3 server.py
