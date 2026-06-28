# Notion 스킬 75개 검토표

> 출처: [claude-프롬프트-스킬-모음](https://exultant-principle-9c5.notion.site/claude-34691cb23c4d806db398fd9fe5e1c364)
> 상태: **검토만** (설치 보류). 설치는 외부 코드를 에이전트 스킬로 통합하는 행위라 별도 권한 필요.

## 요약

- **이미 설치됨**: 14개 (superpowers 계열, `obra/superpowers`에서 설치 완료)
- **신규 후보**: 61개 (약 22개 외부 저장소)
- **신뢰도 분포**: 공식/검증 ~30개 · 커뮤니티 ~13개 · 개인·익명 ~32개

신뢰도 범례: ✅ 공식(Anthropic/Figma/Notion/shadcn) · 🟡 알려진 커뮤니티 · ⚠️ 개인·익명 저장소 · 🟢 이미 설치됨

---

## ✅ 공식 출처 (Anthropic / Figma / Notion / shadcn)

| 스킬 | 설명 | 출처 | 상태 |
|------|------|------|------|
| agent-development | AI 에이전트 구조·시스템 프롬프트·트리거 설계 | anthropics/claude-plugins-official | 신규 |
| docx | Word 문서 생성·편집·분석 (표/목차/이미지) | anthropics/claude-plugins-official | 신규 |
| pdf | PDF 생성·병합·분할·OCR·암호화 | anthropics/claude-plugins-official | 신규 |
| pptx | PowerPoint 생성·편집·파싱·병합 | anthropics/claude-plugins-official | 신규 |
| ppt-creator | 주제/문서를 슬라이드 덱으로 변환 | anthropics/skills | 신규 |
| contract-generator | 클라이언트 정보로 계약서 DOCX 생성 | anthropics/skills (doc-coauthoring) | 신규 |
| visual-ui-tester | Playwright로 반응형/CSS 시각 테스트 | anthropics/skills (webapp-testing) | 신규 |
| mcp-builder | Python/TS로 MCP 서버 개발 가이드 | anthropics/claude-plugins-official | 신규 |
| prompt-engineering-patterns | 프로덕션 LLM 프롬프트 패턴 | anthropics/claude-plugins-official | 신규 |
| ui-ux-pro-max | UI/UX 설계 인텔리전스(스타일·팔레트·폰트) | anthropics/claude-plugins-official | 신규 |
| figma-implement-design | Figma MCP로 디자인→프로덕션 코드 | anthropics/claude-plugins-official | 신규 |
| figma-use | Figma CLI 100+ 명령어 도구 | anthropics/claude-plugins-official | 신규 |
| figma | Figma API 디자인 토큰·코드 생성 | figma/mcp-server-guide | 신규 |
| figma-design | Figma 작업 전반 가이드 | figma/mcp-server-guide | 신규 |
| figma-design-analyzer | Figma→Atomic Design 스펙 문서 | figma/mcp-server-guide | 신규 |
| figma-to-html | Figma→정적 HTML/CSS | figma/mcp-server-guide | 신규 |
| converting-figma-designs | Figma Dev Mode MCP 추출·변환 | figma/mcp-server-guide | 신규 |
| generate-design-system | Figma 디자인 시스템 자동 생성 | figma/mcp-server-guide | 신규 |
| generating-code-connect | Figma Code Connect 매핑 생성 | figma/mcp-server-guide | 신규 |
| notion-knowledge-capture | 대화→Notion 문서 구조화 저장 | makenotion/claude-code-notion-plugin | 신규 |
| notion-meeting-intelligence | Notion 미팅 준비 자료 자동 생성 | makenotion/claude-code-notion-plugin | 신규 |
| notion-research-documentation | Notion 워크스페이스 리서치 보고서 | makenotion/claude-code-notion-plugin | 신규 |
| notion-spec-to-implementation | Notion 스펙→구현 태스크 분해 | makenotion/claude-code-notion-plugin | 신규 |
| shadcn | shadcn/ui 컴포넌트 관리 전담 | shadcn-ui/ui | 신규 |

## 🟡 알려진 커뮤니티 저장소

| 스킬 | 설명 | 출처 | 상태 |
|------|------|------|------|
| design-system-patterns | 디자인 토큰·테마·컴포넌트 아키텍처 | wshobson/agents | 신규 |
| interaction-design | 마이크로인터랙션·모션·트랜지션 | wshobson/agents | 신규 |
| responsive-design | Container Query·CSS Grid 반응형 | wshobson/agents | 신규 |
| visual-design-foundations | 타이포·색상·간격·아이콘 시스템 | wshobson/agents | 신규 |
| content-research-writer | 리서치·인용 기반 콘텐츠 작성 | ComposioHQ/awesome-claude-skills | 신규 |
| scrape-do-automation | Rube MCP 웹 스크래핑 자동화 | ComposioHQ/awesome-claude-skills | 신규 |
| html-to-image-automation | HTML→이미지 변환 자동화 | ComposioHQ/awesome-claude-skills | 신규 |
| content-creator | 글쓰기 요청 라우터 | alirezarezvani/claude-skills | 신규 |
| content-strategy | 콘텐츠 전략 수립 (※ 동명 스킬 이미 설치됨) | alirezarezvani/claude-skills | 중복 |
| image-remove-bg | AI 배경 제거(누끼) | alirezarezvani/claude-skills | 신규 |
| senior-frontend | React·Next·TS·Tailwind 패턴 | alirezarezvani/claude-skills | 신규 |
| SEO Optimizer | SEO 전반 최적화 (※ ai-seo 이미 설치됨) | alirezarezvani/claude-skills | 중복 |
| Email Composer | 전문 이메일 작성 (※ cold-email 이미 설치됨) | alirezarezvani/claude-skills | 중복 |
| youtube | YouTube Data API 조회 | alirezarezvani/claude-skills | 신규 |

## ⚠️ 개인·익명 저장소 (신뢰도 주의 — 코드 검토 권장)

| 스킬 | 설명 | 출처 | 상태 |
|------|------|------|------|
| analyzing-projects | 코드베이스 구조·스택·패턴 분석 | levnikolaevich/claude-code-skills | 신규 |
| Course Creator | 교육 과정·커리큘럼 설계 | GarethManning/claude-education-skills | 신규 |
| curriculum-design | Bloom's Taxonomy 학습목표 설계 | GarethManning/claude-education-skills | 신규 |
| curriculum-iterate-feedback | 피드백 기반 커리큘럼 개선 | GarethManning/claude-education-skills | 신규 |
| curriculum-review-pedagogy | 교수설계 품질 검토 | GarethManning/claude-education-skills | 신규 |
| quiz-maker | 퀴즈 자동 생성·채점 | GarethManning/claude-education-skills | 신규 |
| css-module-generator | CSS Module 파일 생성 | jeremylongshore/claude-code-plugins-plus-skills | 신규 |
| converting-figma-to-html | Figma→HTML/CSS | farmanlab/ai_agent_orchestra | 신규 |
| figma-integration | Figma 디자인-코드 워크플로우 | duongdev/ccpm | 신규 |
| gemini-imagegen | Gemini API 이미지 생성·편집 | everyinc/compound-engineering-plugin | 신규 |
| gog | Google Workspace CLI 도구 | openclaw/openclaw | 신규 |
| mcporter | MCP 서버 CLI 관리 | openclaw/openclaw | 신규 |
| instagram-content | 인스타 릴스·스토리 제작 | petrogurcak/skills | 신규 |
| fact-checker | 문서 사실 검증 | lyndonkl/claude | 신규 |
| prototype-to-production | 프로토타입→프로덕션 컴포넌트 | lyndonkl/claude | 신규 |
| research-claim-map | 주장-출처 대조 신뢰도 평가 | lyndonkl/claude | 신규 |
| skills-discovery | 스킬 레지스트리 검색·설치 메타 스킬 | Kamalnrf/claude-plugins | 신규 |
| threejs-skills | Three.js 3D·WebGL 제작 | sickn33/antigravity-awesome-skills | 신규 |
| ui-analyzer | 스크린샷→React/TS/Tailwind 코드 | nextlevelbuilder/ui-ux-pro-max-skill | 신규 |
| ui-implementer | 디자인 레퍼런스→픽셀 퍼펙트 UI | nextlevelbuilder/ui-ux-pro-max-skill | 신규 |
| viral-hook-creator | SNS 바이럴 훅 생성 | ognjengt/founder-skills | 신규 |
| viral-reel-generator | 숏폼 대본 작성 | aaaronmiller/create-viral-content | 신규 |
| product-design | Figma MCP 디자인 리뷰·토큰 추출 | cuellarfr/design-skills | 신규 |

## 🟢 이미 설치됨 (superpowers — 추가 작업 불필요)

brainstorming · dispatching-parallel-agents · executing-plans · finishing-a-development-branch · receiving-code-review · requesting-code-review · subagent-driven-development · systematic-debugging · test-driven-development · using-git-worktrees · using-superpowers · verification-before-completion · writing-plans · writing-skills

---

## 권장안

이 저장소(LF Mall CRM 대시보드)의 실제 작업과 맞물리는 순서로 우선순위를 매기면:

1. **바로 유용** — 이 프로젝트는 순수 HTML/CSS/JS 대시보드라 디자인·프론트 스킬이 직접적:
   `responsive-design`, `interaction-design`, `visual-design-foundations`, `ui-ux-pro-max`, `senior-frontend`
2. **문서·산출물** — `docx`, `pdf`, `pptx`, `ppt-creator` (공식, 안전)
3. **이 프로젝트와 무관** — 교육 커리큘럼(5종), 바이럴 숏폼, 인스타 콘텐츠, gog/mcporter 등은 lf-dashboard 작업과 거의 무관

⚠️ 익명 개인 저장소 스킬은 설치 전 `SKILL.md`와 동봉 스크립트를 직접 검토하는 것을 권장합니다. 스킬은 에이전트의 모든 도구 권한으로 실행됩니다.
