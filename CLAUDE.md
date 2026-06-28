# LF Mall CRM 고객 여정 대시보드 — 개발 가이드

## 프로젝트 개요

Streamlit으로 감싼 순수 HTML/CSS/JS 싱글페이지 대시보드.
`crm_journey.html` 하나에 모든 UI·로직이 있고, `crm_journey.py`가 `st.components.v1.html()`로 임베드한다.

- **배포**: Streamlit Share (GitHub 연동 자동 배포)
- **브랜치**: `claude/compassionate-gauss-JC2F7` → main 머지 후 배포 반영
- **저장소**: `lee7120-creator/lf-dashboard`

---

## 파일 구조

```
crm_journey.html   ← 대시보드 본체 (전부 여기)
crm_journey.py     ← Streamlit 래퍼 (건드릴 일 거의 없음)
requirements.txt
```

`crm_journey.py` 내용 (참고용, 수정 불필요):
```python
import streamlit as st
import streamlit.components.v1 as components
import pathlib

st.set_page_config(
    page_title="LF Mall CRM 자동화 메시지 — 고객 여정 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown("""<style>#MainMenu, header, footer { visibility: hidden; }</style>""", unsafe_allow_html=True)

html_path = pathlib.Path(__file__).parent / "crm_journey.html"
components.html(html_path.read_text(encoding="utf-8"), height=920, scrolling=True)
```

---

## 데이터 모델

카드(메시지) 한 건의 구조:

```js
{
  id: "uuid",           // crypto.randomUUID()
  code: "A-001",        // AF코드
  label: "가입 환영",   // 캠페인명
  ch: "sms",            // 발송채널 (아래 CH 참고)
  push: false,          // PUSH 여부
  timing: "가입 D+1",   // 발송시점 (사람이 읽는 형태)
  trigger: "...",       // 트리거 조건
  offer: "...",         // 오퍼/혜택
  msg: "...",           // 메시지 내용
  status: "new",        // 상태 (아래 ST 참고)
  stageId: "stage-1",   // 소속 스테이지 ID
  view: "asis",         // "asis" | "tobe"
  ord: 0,               // 드래그앤드롭 정렬 순서 (timingNum 기반으로 seed)
}
```

**채널 (CH)**:
```js
const CH = {
  sms:     { label:'SMS',   c:'#B83A3A', bg:'#FEF2F2', icon:'📱' },
  alimtok: { label:'알림톡', c:'#A07010', bg:'#FEF8EA', icon:'💬' },
  friends: { label:'플친',  c:'#367A4C', bg:'#EEF8F1', icon:'💚' },
  email:   { label:'이메일', c:'#2E68B0', bg:'#EEF3FA', icon:'✉️' },
  inapp:   { label:'인앱',  c:'#7B5BC0', bg:'#F3EFFB', icon:'📲' },
  none:    { label:'없음',  c:'#96938C', bg:'#F2F1EE', icon:'∅'  },
};
```

**상태 (ST)**:
```js
const ST = {
  new: { label:'신규추가', c:'#367A4C', bg:'#EBF8F2', b:'#B4DEC8' },
  mod: { label:'수정개선', c:'#2E68B0', bg:'#EEF3FA', b:'#B4CCE8' },
  kep: { label:'유지',    c:'#706E68', bg:'#F2F1EE', b:'#D8D6CE' },
  del: { label:'삭제검토', c:'#B03030', bg:'#FEF1F0', b:'#E8C0BC' },
};
```

---

## 핵심 아키텍처 패턴

### 1. localStorage 영속성
```js
const STORE_KEY = 'lfmall_crm_v3';  // 데이터 구조 바뀔 때 버전 올려야 함

function saveData() { localStorage.setItem(STORE_KEY, JSON.stringify({stages, cards})); }
function loadData() {
  const raw = localStorage.getItem(STORE_KEY);
  if (raw) { const d = JSON.parse(raw); stages = d.stages; cards = d.cards; }
  else { stages = DEFAULT_STAGES; cards = DEFAULT_CARDS; }
}
```
> 데이터 구조(타이밍 형식, 채널 enum 등) 바뀌면 반드시 STORE_KEY 버전 올릴 것.
> 안 올리면 구버전 캐시가 그대로 로드됨.

### 2. XSS 방지 esc() 헬퍼
사용자 입력값을 innerHTML에 넣을 때 반드시 `esc()` 적용:
```js
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, m =>
    ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[m])
  );
}
// 줄바꿈이 필요한 경우: esc(s).replace(/\n/g, '<br>')  ← 반드시 esc 먼저!
```

### 3. 렌더 함수 호출 순서
```js
loadData();
ensureOrder();   // ord 필드가 없는 카드에 timingNum 기반으로 seed
renderSidebar(); // 필터 사이드바
render();        // 메인 콘텐츠
```

### 4. 필터 상태 (FILTERS)
```js
let FILTERS = {
  search: '',
  channels: new Set(),
  statuses: new Set(),
  views: new Set(),
  stages: new Set(),
  pushOnly: false,
};

function cardMatches(c) {
  if (FILTERS.channels.size && !FILTERS.channels.has(c.ch)) return false;
  if (FILTERS.pushOnly && !c.push) return false;
  if (FILTERS.statuses.size && (!c.status || !FILTERS.statuses.has(c.status))) return false;
  if (FILTERS.search) {
    const q = FILTERS.search.toLowerCase();
    const hay = [c.code, c.label, c.timing, c.trigger, c.offer, c.msg].join(' ').toLowerCase();
    if (!hay.includes(q)) return false;
  }
  return true;
}
```

> **주의**: `onSearch()`는 `renderSidebar()`를 호출하지 않는다. 검색창 포커스 유지를 위해서임.
> 필터 토글 함수들은 반드시 `renderSidebar(); render();` 둘 다 호출.

### 5. timingNum — 정렬용 숫자 변환
```js
function timingNum(t) {
  if (!t) return 999;
  const s = t.toLowerCase();
  if (s.includes('실시간')) return -1;
  if (s.includes('방송') && s.includes('전')) return -0.5;
  if (s.includes('즉시')) return 0;
  const hm = s.match(/(\d+)\s*(?:시간|h)/); if (hm) return parseInt(hm[1]) / 24;
  const dm = s.match(/d\s*\+\s*(\d+)/i);    if (dm) return parseFloat(dm[1]);
  if (s.includes('배송완료')) return 0.1;
  const dom = s.match(/(\d+)\s*일/);         if (dom) return parseInt(dom[1]);
  if (s.includes('매일')) return 30;
  return 999;
}
```

### 6. 드래그앤드롭
- 카드에 `draggable="true"` + `ondragstart` / 드롭존에 `ondragover` / `ondrop`
- `ord` 필드로 순서 저장. 드롭 후 `saveData()` 호출.
- 같은 row 내 이동 및 다른 stage/row 간 이동 모두 지원.

### 7. 사이드바 토글
```js
function toggleSidebar() { document.body.classList.toggle('sidebar-collapsed'); }
```
```css
.sidebar { position: fixed; left: 0; top: 48px; width: 216px; transition: transform .22s; }
.content { margin-left: 216px; transition: margin-left .22s; }
body.sidebar-collapsed .sidebar { transform: translateX(-100%); }
body.sidebar-collapsed .content { margin-left: 0; }
```

---

## 레이아웃 구조

```
[Topbar 48px — sticky]
  ☰버튼 | LF 로고 | [sep] | 저장됨 | 내보내기 | 불러오기

[Sidebar 216px — fixed left, top:48px]
  필터 | 초기화
  검색창
  발송채널 chips
  상태 chips
  PUSH chips
  AS-IS/TO-BE chips
  스테이지 chips
  [카드 수 카운터]

[Content — margin-left:216px]
  [Stage Section × N]
    [Stage Header — 색상 편집 가능]
    [Gap Row — 골든타임 등 메모]
    [AS-IS Row]
      [card] [card] ... [+ 버튼]
    [divider]
    [TO-BE Row]
      [card] [card] ... [+ 버튼]
```

---

## 카드 뱃지 순서 (중요)

왼쪽부터: **PUSH 뱃지 → 채널 뱃지 → 상태 뱃지**

```html
<div class="card-badges">
  ${c.push ? '<span class="push-badge">🔔 PUSH</span>' : ''}
  <span class="ch-badge" style="background:${ch.bg};color:${ch.c};border-color:${ch.c}30">
    ${ch.icon} ${esc(ch.label)}
  </span>
  ${stBadge}
</div>
```

뱃지 3개 동시에 떠도 안 깨지려면:
- `flex-wrap: wrap; row-gap: 4px` — card-badges
- `white-space: nowrap; flex-shrink: 0` — 각 뱃지
- `status-badge`에 `margin-left: auto` **절대 쓰지 말 것** (두 번째 줄로 밀림)
- 카드 너비는 최소 210px 필요

---

## 타이밍 문자열 형식

사람이 읽기 쉬운 앵커 형식으로 작성:

| 형식 예시 | timingNum 결과 |
|-----------|---------------|
| 즉시 | 0 |
| 가입 D+1 | 1 |
| 구매 D+7 | 7 |
| 미방문 D+90 | 90 |
| 2시간 후 | 0.083 |
| 실시간 | -1 |
| 방송 1시간 전 | -0.5 |

---

## 흔히 생기는 버그 & 해결법

| 버그 | 원인 | 해결 |
|------|------|------|
| 카드 뱃지 줄바꿈 깨짐 | `margin-left:auto` on status badge | 제거, `flex-shrink:0` 추가 |
| HTML 특수문자 깨짐/주입 | innerHTML에 raw 삽입 | `esc()` 적용 |
| 구버전 데이터 로드 | STORE_KEY 그대로 | 버전 올리기 (v3→v4) |
| 검색 중 포커스 빠짐 | onSearch가 renderSidebar 호출 | onSearch에서 renderSidebar 제거 |
| 새 스테이지 칩 미갱신 | 색상/라벨 변경 후 사이드바 미갱신 | color picker change + label save에 `renderSidebar()` 추가 |
| import 후 칩 미갱신 | importJSON이 renderSidebar 미호출 | `importJSON` 마지막에 `renderSidebar()` 추가 |

---

## 새 기능 추가 시 체크리스트

- [ ] 사용자 입력 → innerHTML 넣는 곳에 `esc()` 적용했는가?
- [ ] 데이터 구조 변경 시 `STORE_KEY` 버전 올렸는가?
- [ ] 필터 관련 변경 시 `renderSidebar()` + `render()` 둘 다 호출하는가?
- [ ] 카드 너비 210px 유지하는가?
- [ ] 타이밍 문자열 형식이 `timingNum()`으로 파싱되는가?

---

## 효과적이었던 프롬프트 패턴

### 채널/상태 추가
```
발송채널에 [채널명]도 추가해줘. 색상은 [색상 설명]으로.
```
→ CH 객체에 항목 추가 + 사이드바 chip 자동 반영됨

### 타이밍 표시 변경
```
발송시점 표시를 [기존 형식] 대신 [새 형식]으로 변경해줘.
예시: 가입 D+1, 구매 D+7 이렇게.
```
→ 기존 카드 데이터 일괄 마이그레이션 + STORE_KEY 버전 업 필요

### 채널 일괄 변경
```
[조건]에 해당하는 카드들의 채널을 [채널]로 일괄 변경해줘.
```
→ DEFAULT_CARDS 배열 직접 수정 + STORE_KEY 버전 업

### 필터 사이드바 확장
```
사이드바 필터에 [항목] 기준으로도 필터링 추가해줘.
```
→ FILTERS 객체에 필드 추가 + renderSidebar() + cardMatches() 수정

### 드래그앤드롭
```
카드를 드래그앤드롭으로 같은 row 안에서, 그리고 다른 stage/row 간에도 이동 가능하게 해줘.
```
→ ord 필드 + draggable 속성 + 이벤트 핸들러

### 뱃지 순서/위치
```
[뱃지명] 뱃지가 [위치]에 오도록 변경해줘.
```
→ cardHTML 내 뱃지 HTML 순서 변경 + flex 속성 확인

---

## Git 워크플로

```bash
# 개발 브랜치 확인
git branch

# 변경 커밋
git add crm_journey.html
git commit -m "기능 설명"
git push -u origin claude/compassionate-gauss-JC2F7

# PR 생성 (GitHub MCP)
# mcp__github__create_pull_request 또는 mcp__github__list_pull_requests로 기존 PR 확인 후
# mcp__github__update_pull_request 로 업데이트

# 머지 전 draft → ready 전환 필수
# mcp__github__update_pull_request { draft: false }
# mcp__github__merge_pull_request { merge_method: "squash" }
```

---

## 구글시트 연동 (발송성과 대시보드)

### 개요
`send_perf_dashboard.py`는 구글시트를 영속 저장소로 사용한다.
`gspread` + `google-auth` 서비스 계정 방식이며, 키는 Streamlit Cloud Secrets에 TOML로 저장한다.

### Streamlit Secrets 설정 위치
Streamlit Cloud → 앱 **Settings** → **Secrets** 탭

### Secrets TOML 형식
```toml
[gcp_service_account]
type = "service_account"
project_id = "quick-doodad-397006"
private_key_id = "키ID"
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEv...base64...==\n-----END PRIVATE KEY-----\n"
client_email = "googlesheet@quick-doodad-397006.iam.gserviceaccount.com"
client_id = "116048104131558945028"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/googlesheet%40quick-doodad-397006.iam.gserviceaccount.com"
universe_domain = "googleapis.com"
```

### private_key 주의사항
- Streamlit TOML이 `\n`을 깨뜨리거나 `-----BEGIN/END-----` 마커를 누락시킬 수 있음
- `_fix_pem()` 함수가 자동 복구: 마커 없이 base64만 있어도 표준 PEM으로 재구성
- `_pem_diag()` 함수가 진단: `BEGIN:있음/없음 END:있음/없음 본문:N자` 형태로 에러 원인 표시
- **키는 절대 코드나 깃에 넣지 말 것** — `.gitignore`에 `.streamlit/secrets.toml` 등록됨

### 코드 구조
```python
# 스프레드시트 제목 매핑
GS_TITLES = {"campaign": "campaign_store", "mtd": "mtd_store", "promo": "promo_store"}

# 키 로드 → PEM 복구 → 인증 → 시트 열기
def gs_open(creds_dict, spreadsheet):
    info["private_key"] = _fix_pem(info.get("private_key"))
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open(spreadsheet)  # URL/키/제목 모두 허용

# 저장/로드 — 구글시트 실패 시 로컬 CSV 폴백
_save_gs(kind, df)   # 시트에 덮어쓰기
_load_gs(kind)       # 시트에서 DataFrame 로드
```

### 구글시트 공유 설정
스프레드시트를 서비스 계정 이메일(`client_email`)에 **편집자** 권한으로 공유해야 함.

### 흔한 에러
| 에러 | 원인 | 해결 |
|------|------|------|
| Unable to load PEM file | `private_key`의 마커/줄바꿈 깨짐 | `_fix_pem()`이 자동 복구 (PR #108) |
| BEGIN:없음 END:없음 | Secrets TOML이 마커를 날림 | 코드가 자동으로 마커 추가 |
| 본문:0자 | `private_key`가 비어있음 | Secrets에 키 값 재입력 |
| 403 Forbidden | 시트 공유 안 됨 | `client_email`에 편집자 권한 공유 |

---

## Streamlit 배포 주의사항

- `height=920` — 대시보드가 잘리면 이 값 키워야 함
- `scrolling=True` — 내부 스크롤 허용
- Streamlit iframe 내부라 `window.localStorage`는 정상 작동함
- 외부 폰트(Google Fonts)는 iframe 내부에서도 로드됨
