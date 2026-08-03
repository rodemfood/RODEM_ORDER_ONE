# RODEM ORDER ONE v2.1 SAFE RENDER

Render 무료 Web Service에 그대로 업로드하여 실행하는 안전형 테스트 버전입니다.

## 주소
- 고객 모바일 주문: `/order`
- 직원 PC 관리: `/staff`
- 서버 상태 확인: `/health`

## 저장 방식
- `DATABASE_URL`이 없으면 Render의 쓰기 가능한 `/tmp/rodem_order_one.db`를 사용합니다.
- 따라서 별도의 `data` 폴더가 없어도 배포 오류가 발생하지 않습니다.
- 무료 테스트 중 서버 재시작 또는 재배포 시 주문 데이터가 초기화될 수 있습니다.
- 나중에 Supabase PostgreSQL의 연결 주소를 Render 환경변수 `DATABASE_URL`에 넣으면 코드 변경 없이 영구 저장으로 전환됩니다.

## GitHub 업로드
압축을 푼 폴더 안의 파일과 `templates` 폴더를 모두 GitHub 저장소에 업로드하고 Commit 합니다.

## Render 설정
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120`
- Instance Type: Free

## 정상 확인
배포 후 `/health`를 열었을 때 다음과 비슷하게 나오면 정상입니다.

```json
{"ok":true,"database":"temporary-sqlite","persistent":false}
```
