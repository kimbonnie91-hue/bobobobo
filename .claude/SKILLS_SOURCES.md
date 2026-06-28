# Installed Skills, Agents & Commands — Sources

This directory contains third-party Claude Code skills, subagents, and slash
commands installed into the project. They are version-controlled here so they
persist across the remote/ephemeral Claude Code sessions used for this repo.

All source projects are **MIT licensed**. See each upstream repo for full text.

| Source repo | Version | What was installed | Excluded |
|-------------|---------|--------------------|----------|
| [coreyhaines31/marketingskills](https://github.com/coreyhaines31/marketingskills) | 2.5.1 | 45 skills | — |
| [AgriciDaniel/claude-seo](https://github.com/AgriciDaniel/claude-seo) | 2.2.0 | 25 skills + 18 agents | Python scripts, MCP extensions, and the auto-running `PostToolUse` validation hook |
| [obra/superpowers](https://github.com/obra/superpowers) | — | 14 skills | `SessionStart` hook (not auto-enabled) |
| [luongnv89/claude-howto](https://github.com/luongnv89/claude-howto) | 2.1.x | 6 skills + 9 subagents + 8 slash commands | tutorial modules, `06-hooks/`, `07-plugins/` |
| Notion "claude 프롬프트+스킬 모음" 배치 (아래 표 참고) | — | 41 skills | 중복/대량 번들/dead repo 제외 |

**Totals:** 131 skills · 27 agents · 8 commands

## Notion 스킬 배치 (41개 추가)

[Notion 큐레이션 목록](https://exultant-principle-9c5.notion.site/claude-34691cb23c4d806db398fd9fe5e1c364)의
75개 중, 이미 설치돼 있던 14개(superpowers 계열)와 기존 스킬과 이름이 겹치는 항목,
그리고 dead/renamed 저장소를 제외하고 **41개**를 설치했다.

| 출처 repo | 설치한 스킬 |
|-----------|-------------|
| anthropics/claude-plugins-official | agent-development |
| anthropics/skills | docx, pdf, pptx, doc-coauthoring, webapp-testing, mcp-builder |
| wshobson/agents | design-system-patterns, interaction-design, responsive-design, visual-design-foundations, prompt-engineering-patterns |
| figma/mcp-server-guide | figma-use, figma-generate-design, figma-code-connect |
| shadcn-ui/ui | shadcn |
| makenotion/claude-code-notion-plugin | knowledge-capture, meeting-intelligence, research-documentation, spec-to-implementation |
| jeremylongshore/claude-code-plugins-plus-skills | css-module-generator |
| farmanlab/ai_agent_orchestra | converting-figma-to-html |
| duongdev/ccpm | figma-integration |
| openclaw/openclaw | gog, mcporter |
| lyndonkl/claude | prototyping-pretotyping, research-claim-map, scientific-clarity-checker |
| Kamalnrf/claude-plugins | skills-discovery |
| sickn33/antigravity-awesome-skills | threejs-skills |
| alirezarezvani/claude-skills | content-creator, senior-frontend, senior-computer-vision, video-content-strategist, ui-ux-pro-max |
| ognjengt/founder-skills | viral-hook-creator |
| ComposioHQ/awesome-claude-skills | content-research-writer, scrape-do-automation, html-to-image-automation |
| cuellarfr/design-skills | product-design |

### 설치 경위 (중요)

- **설치 방식:** Claude Code 에이전트가 외부 저장소를 자신의 스킬로 직접 통합하는 것은
  보안 가드레일(자기수정 / untrusted code integration)에 막힌다. 사용자가 설치 명령을
  직접 실행해 신뢰 결정을 내린 뒤, 결과를 버전 관리에 반영했다.
- **대량 번들 주의:** 일부 저장소(`ComposioHQ/awesome-claude-skills`, `duongdev/ccpm` 등)는
  수백~수천 개 스킬을 번들로 포함한다. 저장소 전체를 `find SKILL.md`로 긁으면 3,700개가
  설치되는 사고가 난다. **반드시 의도한 스킬 이름 allowlist로 필터링**할 것.

### 설치되지 않은 항목 (수동 처리 필요)

| 스킬 | 사유 |
|------|------|
| instagram-content | `petrogurcak/skills` 클론 실패 (저장소 이동/삭제 추정) |
| analyzing-projects | `levnikolaevich/claude-code-skills` 내 폴더명 불일치 |
| gemini-imagegen | `everyinc/compound-engineering-plugin`에서 `ce-gemini-imagegen` 명으로 존재 |
| ui-analyzer, ui-implementer | `nextlevelbuilder/ui-ux-pro-max-skill` 단일 스킬(ui-ux-pro-max)만 제공 |
| curriculum-*, Course Creator | `GarethManning/claude-education-skills` 폴더 구조 상이 |
| viral-reel-generator | `aaaronmiller/create-viral-content` 폴더명 불일치 |

필요 시 위 항목은 정확한 업스트림 경로를 확인해 개별 설치할 것.

## Notes

- **`seo-audit` name collision:** both `marketingskills` and `claude-seo` ship a
  skill named `seo-audit`. The dedicated SEO suite keeps the canonical name
  (`seo-audit`); the marketing one was renamed to **`seo-audit-marketing`**
  (folder + frontmatter `name`) so both remain usable.
- **No auto-running hooks were installed.** Hooks from `claude-seo`,
  `superpowers`, and `claude-howto` execute code automatically (on edits or
  session start) and require extra Python/Node dependencies, so they were left
  out. The skills/agents themselves are invoked on demand.
- **`claude-seo` script-dependent steps are limited** because the helper Python
  scripts were not installed. Re-run with the scripts if you need full
  functionality.

## Updating

Re-clone the upstream repo and re-copy the relevant directories into
`.claude/skills`, `.claude/agents`, or `.claude/commands`.
