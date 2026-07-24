# RODEM ORDER ONE

고객 주문과 사무실 직원의 로젠 엑셀 생성을 위한 심플 웹앱입니다.

## 접속 주소
- 고객 주문: `/order`
- 직원 관리: `/staff`
- 상태 확인: `/health`

## Render 배포
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`

`render.yaml`을 사용하면 위 설정이 자동 적용됩니다.

## 테스트 주의
무료 Render 인스턴스의 로컬 SQLite 데이터는 재배포 또는 서버 재시작 때 사라질 수 있습니다. 현재 버전은 외부 휴대폰 주문 흐름을 확인하는 무료 테스트용입니다.
