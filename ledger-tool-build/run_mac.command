#!/bin/bash
cd "$(dirname "$0")"
python3 build_ledger.py
echo ""
echo "작업이 끝났습니다. 아무 키나 누르면 창이 닫힙니다."
read -n 1
