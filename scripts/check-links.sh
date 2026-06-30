#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS_DIR="$REPO_ROOT/skills"
CLAUDE_SKILLS="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
CODEX_SKILLS="${CODEX_SKILLS_DIR:-$HOME/.agents/skills}"

SKILLS=(
  arr-to-bookings
  customer-concentration-analysis
  enrich-companies
  excel-to-jsonl
  finance-formatting
  market-landscape
  market-map-template-reader
  pipeline-analysis
  render-market-map
  retention-analysis
  sourcing-setup
  standardize-pnl
  theme-first-pass
)

realpath_py() {
  python3 - "$1" <<'PY'
import pathlib
import sys
print(pathlib.Path(sys.argv[1]).resolve())
PY
}

fail=0

for skill in "${SKILLS[@]}"; do
  source="$SKILLS_DIR/$skill"
  if [[ ! -f "$source/SKILL.md" ]]; then
    echo "FAIL missing source: $source/SKILL.md"
    fail=1
  fi

  for root in "$CLAUDE_SKILLS" "$CODEX_SKILLS"; do
    dest="$root/$skill"
    if [[ ! -L "$dest" ]]; then
      echo "FAIL not a symlink: $dest"
      fail=1
      continue
    fi
    actual="$(realpath_py "$dest")"
    expected="$(realpath_py "$source")"
    if [[ "$actual" != "$expected" ]]; then
      echo "FAIL wrong target: $dest -> $actual (expected $expected)"
      fail=1
    else
      echo "OK $dest -> $actual"
    fi
  done
done

if [[ "$fail" -ne 0 ]]; then
  echo "Skill link check failed."
  exit 1
fi

echo "All investor-toolkit skill links are healthy."
