# RODEM ORDER ONE v3.0 STABLE

고객 모바일 주문, 고객정보 복구, 최근 주문 재주문, 직원 PC 관리, 거래처 발주서 엑셀 자동변환, 로젠 송장 생성 통합판입니다.

## 주요 변경
- 고객정보 쿠키 + 로컬 토큰 이중 저장
- 저장정보가 사라진 경우 업체명/연락처로 복구
- 최근 주문 그대로 불러오기
- 최근 상품명 자동완성
- 주문번호 YYYYMMDD-0001 형식
- 직원 주문 검색 및 상태 필터
- 신규 주문 알림음/브라우저 알림
- 거래처 엑셀 자동분석 및 로젠 형식 변환
- DATABASE_URL이 없으면 /tmp SQLite로 안전 실행

## Render 설정
Build Command: `pip install -r requirements.txt`
Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120`

무료 Render에서는 서버가 잠들 수 있고 /tmp SQLite 데이터는 재배포 시 초기화될 수 있습니다. 실운영 시 DATABASE_URL을 연결하십시오.
