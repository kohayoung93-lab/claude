#!/bin/bash
# 더블클릭으로 실행하는 macOS용 실행 파일
cd "$(dirname "$0")"
python3 append_ledger.py
echo ""
echo "작업이 끝났습니다. 아무 키나 누르면 창이 닫힙니다."
read -n 1
