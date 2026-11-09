#!/bin/bash
# Mac용 실행 파일: 더블클릭하면 터미널에서 서버가 켜지고 브라우저가 열립니다.
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  echo "[오류] Python3가 없습니다. https://www.python.org/downloads/ 에서 설치 후 다시 실행하세요."
  open https://www.python.org/downloads/
  read -p "Enter 키를 누르면 닫힙니다."; exit 1
fi
if [ ! -d .venv ]; then
  echo "처음 실행: 필요한 프로그램 설치 중..."
  python3 -m venv .venv || { read -p "venv 생성 실패. Enter"; exit 1; }
fi
.venv/bin/pip install -q --disable-pip-version-check -r requirements.txt || { read -p "패키지 설치 실패(인터넷 확인). Enter"; exit 1; }
(sleep 3; open http://localhost:8080) &
echo "서버 시작: http://localhost:8080  (이 창을 닫으면 페이지가 꺼집니다)"
.venv/bin/python app.py
