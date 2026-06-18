# Ark Scheduler Final

Microsoft 365 Graph `getSchedule` 기반 조직 참석자 공통 빈 시간 검색 POC입니다.

## 실행

```bash
npm install
npm run dev
```

## 환경변수

`.env.example`를 `.env`로 복사한 뒤 Entra App 값을 입력합니다.

```bash
VITE_MSAL_CLIENT_ID=...
VITE_MSAL_TENANT_ID=...
```

필요 권한:

- User.Read
- User.ReadBasic.All
- Calendars.Read

## 이번 정비 내용

- 화면을 3단 구조로 재정비
  - 좌측: 조건
  - 중앙: 참석자 검색/선택
  - 우측: 스케줄
- 참석자 목록에서 선택된 사람을 항상 최상단 정렬
- 참석자 카드를 낮게 만들어 10명 이상도 한눈에 보이도록 수정
- 선택된 참석자 placeholder와 선택 chip 영역 분리
- 공통시간 산출 로직 보강
  - 선택 참석자 ID에 매칭되는 busy block만 불가 처리
  - Graph 응답 순서와 scheduleId 매칭을 모두 고려
  - 시간대 파싱을 로컬 기준으로 안정화
  - 잘못된 fallback scheduleId가 `requireAll` 조건을 오염시키는 문제 방지
- build 검증 완료


## 참석자 목록 방식
- 로그인 후 Microsoft Graph `/users`에서 조직 내 사용자를 전체 로딩합니다.
- 별도 서버 검색 없이 목록에서 사람을 클릭해 선택합니다.
- 선택된 참석자는 자동으로 목록 최상단으로 정렬됩니다.
- 숨길 메일은 `.env`의 `VITE_EXCLUDED_EMAILS`에 쉼표로 추가합니다.

```env
VITE_EXCLUDED_EMAILS=room@arknpartners.com,shared@arknpartners.com
```
