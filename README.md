# 오션메이드 도어발주현황

## Vercel 배포 방법 (5분 소요)

### 1단계 — GitHub에 올리기
1. https://github.com 가입 (무료)
2. "New repository" → 이름: `oceanmade-door` → Create
3. 이 폴더 파일들을 모두 업로드

### 2단계 — Vercel 배포
1. https://vercel.com 가입 (구글 로그인 가능)
2. "Add New Project" → GitHub 연결 → `oceanmade-door` 선택
3. **Environment Variables** 항목에 추가:
   - Key: `VITE_ANTHROPIC_API_KEY`
   - Value: (Anthropic API 키 — 이미지 견적서 기능에 필요)
4. "Deploy" 클릭

### 3단계 — 직원 공유
배포 완료되면 `oceanmade-door.vercel.app` 같은 주소 생성
→ 카카오톡으로 공유하면 끝!

### 스마트폰 홈화면 추가 방법
- iPhone: Safari에서 열기 → 공유 → "홈 화면에 추가"
- Android: Chrome에서 열기 → 메뉴 → "홈 화면에 추가"
→ 앱처럼 사용 가능!

## API 키 없이 쓰는 법
이미지 견적서 기능만 안 되고, 엑셀 업로드와 직접 입력은 API 키 없이도 완전 사용 가능합니다.
