#!/usr/bin/env bash
#
# lf-dashboard의 .claude/ 스킬·에이전트·슬래시 커맨드를
# 전역(~/.claude)으로 설치해, 모든 프로젝트에서 쓸 수 있게 합니다.
#
# 사용법:
#   git pull          # 최신 .claude/ 받기
#   bash install-skills.sh
#
set -euo pipefail

SRC="$(cd "$(dirname "$0")" && pwd)/.claude"
DEST="${HOME}/.claude"

if [ ! -d "$SRC/skills" ]; then
  echo "❌ .claude/skills 폴더를 찾을 수 없습니다: $SRC"
  echo "   lf-dashboard 레포 루트에서 실행하세요."
  exit 1
fi

mkdir -p "$DEST/skills" "$DEST/agents" "$DEST/commands"

# -n(no-clobber): 이미 쓰고 있던 개인 스킬/에이전트/커맨드는 덮어쓰지 않음
cp -Rn "$SRC/skills/."   "$DEST/skills/"   2>/dev/null || true
cp -Rn "$SRC/agents/."   "$DEST/agents/"   2>/dev/null || true
cp -Rn "$SRC/commands/." "$DEST/commands/" 2>/dev/null || true

count_dirs() { find "$1" -mindepth 1 -maxdepth 1 -type d 2>/dev/null | wc -l | tr -d ' '; }
count_md()   { find "$1" -mindepth 1 -maxdepth 1 -type f -name '*.md' 2>/dev/null | wc -l | tr -d ' '; }

echo "✅ 전역 설치 완료 → $DEST"
echo "   skills:   $(count_dirs "$DEST/skills")"
echo "   agents:   $(count_md   "$DEST/agents")"
echo "   commands: $(count_md   "$DEST/commands")"
echo ""
echo "이제 아무 프로젝트 폴더에서 Claude Code를 켜면 스킬이 인식됩니다."
echo "(개인 설정은 -n 옵션으로 보존됩니다. 기존 것을 새 버전으로 강제 교체하려면"
echo " 스크립트의 cp -Rn 을 cp -Rf 로 바꿔 실행하세요.)"
