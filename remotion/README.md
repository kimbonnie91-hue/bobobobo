# LF Dashboard — Remotion 영상 프로젝트

[Remotion](https://www.remotion.dev/)으로 React 코드를 사용해 영상을 만드는 **독립 프로젝트**입니다.
Streamlit/HTML 대시보드(`crm_journey.html` 등)와는 완전히 분리되어 있으며, 별도의 Node 환경에서 동작합니다.

## 요구 사항

- Node.js 18+ (현재 v22 확인됨)
- 최초 `npm install` 시 Remotion이 헤드리스 크롬을 내려받습니다(수백 MB).

> ⚠️ **원격(Claude Code on the web) 환경 주의:** 이 환경은 네트워크 egress가
> allowlist로 제한되어 있어 `remotion.media`(크롬 헤드리스 배포 호스트)가 차단됩니다.
> 그래서 원격 세션에서는 `studio`/`render`의 크롬 다운로드가 403으로 실패합니다.
> **로컬에서 실행**하거나, 환경 네트워크 설정의 allowlist에 `remotion.media`를 추가하세요.
> (`npm install`·타입체크·CLI 인식은 원격에서도 정상 동작 확인됨)

## 설치

```bash
cd remotion
npm install
```

## 미리보기 (Remotion Studio)

```bash
npm run dev
# = npx remotion studio
```

브라우저에서 스튜디오가 열리고, `HelloWorld` 컴포지션을 실시간으로 편집·미리보기할 수 있습니다.

## 렌더링 (mp4 출력)

```bash
npm run render
# = npx remotion render HelloWorld out/HelloWorld.mp4
```

결과물은 `out/HelloWorld.mp4`에 생성됩니다.

## 구조

```
remotion/
├─ package.json          의존성 · 스크립트
├─ tsconfig.json
├─ remotion.config.ts    렌더 설정
└─ src/
   ├─ index.ts           registerRoot 진입점
   ├─ Root.tsx           Composition 등록
   └─ HelloWorld.tsx     샘플 컴포지션 (LF Mall CRM 타이틀 애니메이션)
```

## 새 영상 추가

1. `src/`에 새 컴포넌트(예: `MyVideo.tsx`)를 만듭니다.
2. `src/Root.tsx`에 `<Composition id="MyVideo" component={MyVideo} ... />`를 추가합니다.
3. `npm run dev`로 확인하고 `npx remotion render MyVideo out/MyVideo.mp4`로 렌더합니다.

> 참고: `node_modules/`와 `out/`은 `.gitignore` 처리되어 커밋되지 않습니다.
