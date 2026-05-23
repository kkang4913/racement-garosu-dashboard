# RACEMENT Garosu Dashboard

레이스먼트 가로수길 매장의 주간/일간 운영 데이터를 확인하기 위한  
Google Apps Script 기반 Morning Briefing Dashboard.

---

## 주요 기능

- **홈 Morning Briefing Dashboard**
  - C01 오늘의 매장 상태
  - C02 목표 매출 현황 (일/주/월)
  - C03 상품 흐름 Top 5
  - C04 제품 이슈 등록/처리 상태 변경
  - C05 미구매 사유 분석
  - C06 고객 피드백 요약
  - C07 세이프사이즈 진행/구매 전환
  - C08 SNS 인증 이벤트 참여/구매 전환
  - C09 오늘의 특이사항 + 액션
- **제품이슈 탭** — KPI + 게시판 + 상세/처리 팝업
- **미구매 탭** — KPI + 사유별 분석 + 게시판
- **특이사항 탭** — KPI + 중요도별 분석 + 게시판
- **세이프사이즈 탭** — 전환율 KPI + 게시판
- **SNS 인증 탭** — 전환율 KPI + 게시판
- **점장설정** — 매장 상태 문장, 오늘의 액션, 목표 매출 입력

---

## 파일 구조

```
racement-garosu-dashboard/
├─ README.md
├─ .gitignore
├─ frontend/
│  └─ index_v2.html          # 프론트 화면 (HTML + CSS + JS 단일 파일)
├─ backend/
│  └─ dashboard_final.gs     # Apps Script 백엔드
├─ docs/
│  ├─ 데이터_구조.md
│  ├─ 백엔드_함수_명세.md
│  ├─ 배포_체크리스트.md
│  └─ 수정_히스토리.md
└─ archive/
   └─ old_versions/           # 이전 버전 백업
```

---

## 기술 스택

| 항목 | 내용 |
|---|---|
| 호스팅 | Google Apps Script Web App |
| 프론트 | HTML + CSS + JavaScript (단일 파일) |
| 백엔드 | Google Apps Script |
| 데이터 저장 | Google Sheets |
| 고객 피드백 | Google Form 연동 |
| 판매 데이터 | ERP 엑셀 → Sheets 업로드 |
| 사용 환경 | 매장 내 노트북, 1280px 이상 |

---

## Apps Script 배포 방법

1. [script.google.com](https://script.google.com) → Google Sheets에 연결된 프로젝트 열기
2. `Code.gs` 내용을 `backend/dashboard_final.gs`로 교체
3. `index` 파일 내용을 `frontend/index_v2.html`로 교체
4. 저장 (`Ctrl+S`)
5. **배포** → **배포 관리** → ✏️ → 새 버전 → **배포**
6. 웹앱 URL 접속 확인

---

## 민감 정보 관리

아래 값은 GitHub에 올리지 않습니다.  
실제 운영 값은 Apps Script 프로젝트 내부에서만 관리합니다.

```javascript
// 실제 값은 코드에 하드코딩하지 않습니다
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
const GS_DEPLOY_URL = 'YOUR_APPS_SCRIPT_DEPLOY_URL';
```

---

## 개발/수정 흐름

```
수정본 작성 (Claude / GPT)
    ↓
검증 (Claude Code 2 QA)
    ↓
Git 커밋
    ↓
Apps Script에 반영
    ↓
실제 테스트
    ↓
문제 없으면 운영 반영
```

---

## 커밋 메시지 규칙

| prefix | 용도 |
|---|---|
| `feat:` | 새로운 기능 추가 |
| `fix:` | 오류 수정 |
| `ui:` | 화면/디자인 수정 |
| `perf:` | 속도 개선 |
| `docs:` | 문서 수정 |
| `refactor:` | 구조 개선 |
| `chore:` | 설정, 파일 이동 등 |

---

## 버전

현재: **v0.4**  
자세한 변경 이력은 [docs/수정_히스토리.md](docs/수정_히스토리.md) 참고
