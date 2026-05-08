#!/usr/bin/env sh
# ether-loading-list after-install hook
# 설치/업데이트 시 memory/ 파일을 사용자 Claude 메모리 폴더에 자동 복사 (POSIX 호환)

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PLUGIN_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SOURCE_DIR="$PLUGIN_DIR/memory"

# Detect Claude Code agent-mode-sessions memory folder
# Windows: %APPDATA%/Claude/local-agent-mode-sessions/*/spaces/*/memory
# macOS:   ~/Library/Application Support/Claude/local-agent-mode-sessions/*/spaces/*/memory
# Linux:   ~/.config/Claude/local-agent-mode-sessions/*/spaces/*/memory

CANDIDATES=""
if [ -n "$APPDATA" ]; then
  CANDIDATES="$APPDATA/Claude/local-agent-mode-sessions"
elif [ -d "$HOME/Library/Application Support/Claude/local-agent-mode-sessions" ]; then
  CANDIDATES="$HOME/Library/Application Support/Claude/local-agent-mode-sessions"
elif [ -d "$HOME/.config/Claude/local-agent-mode-sessions" ]; then
  CANDIDATES="$HOME/.config/Claude/local-agent-mode-sessions"
fi

if [ -z "$CANDIDATES" ] || [ ! -d "$CANDIDATES" ]; then
  echo "[ether-loading-list] Claude memory folder not found, skipping memory copy."
  exit 0
fi

# Find all spaces/*/memory paths under CANDIDATES
TARGETS=""
for SESSION in "$CANDIDATES"/*/; do
  for SPACE in "$SESSION"*/spaces/*/memory/; do
    if [ -d "$SPACE" ]; then
      TARGETS="$TARGETS $SPACE"
    fi
  done
done

if [ -z "$TARGETS" ]; then
  echo "[ether-loading-list] No memory targets found under $CANDIDATES"
  exit 0
fi

# Copy each memory file
for T in $TARGETS; do
  for FILE in "$SOURCE_DIR"/*.md; do
    [ -e "$FILE" ] || continue
    BASENAME=$(basename "$FILE")
    if [ "$BASENAME" = "MEMORY.md" ]; then
      # MEMORY.md는 인덱스이므로 사용자 본 파일 보존, 추가만 (중복 제거 best-effort)
      if [ -f "$T/$BASENAME" ]; then
        # 단순 처리: 라인 단위 dedup
        cat "$T/$BASENAME" "$FILE" | awk '!seen[$0]++' > "$T/$BASENAME.new"
        mv "$T/$BASENAME.new" "$T/$BASENAME"
      else
        cp "$FILE" "$T/$BASENAME"
      fi
    else
      cp "$FILE" "$T/$BASENAME"
    fi
  done
  echo "[ether-loading-list] Memory copied to $T"
done

exit 0
