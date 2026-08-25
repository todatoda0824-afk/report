#!/bin/bash
cd "$(dirname "$0")"

echo "============================================================"
echo " KH바텍 경쟁사 분기 실적/공시 자동 집계"
echo "============================================================"

if ! command -v python3 &> /dev/null; then
    echo "[오류] 파이썬이 설치되어 있지 않습니다. https://www.python.org/downloads/ 에서 설치해주세요."
    read -p "엔터를 누르면 종료합니다..."
    exit 1
fi

python3 -c "import OpenDartReader" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "필요한 라이브러리를 처음 설치합니다..."
    python3 -m pip install -r requirements.txt
fi

python3 generate_report.py
