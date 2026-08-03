# RODEM ORDER ONE v1.0

고객 주문과 직원 로젠 송장 생성을 위한 최소 운영 버전입니다.

## 화면
- 고객 주문: `/order`
- 직원 관리: `/staff`
- 상태 확인: `/health`

## Render 설정
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 120`
- Instance Type: Free 또는 Starter

## 데이터 저장
`DATABASE_URL`이 있으면 PostgreSQL을 사용하고, 없으면 로컬 SQLite를 사용합니다.
Render의 영구 디스크나 PostgreSQL 없이 SQLite만 사용하면 재배포/재시작 시 데이터가 사라질 수 있으므로 최초 외부 테스트 용도로만 사용하세요.

## 로젠 송장
공식 열 순서의 `.xls` 파일을 생성하며 상품명은 반드시 다음 형식으로 출력됩니다.

`#전복죽20#호박죽20#닭죽20`


## v1.1 주소 검색 업그레이드
- Kakao 우편번호 검색 버튼 추가
- 우편번호 자동 저장
- 도로명/지번 주소 자동 입력
- 상세주소만 고객이 직접 입력
