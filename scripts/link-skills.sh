#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS_DIR="$REPO_ROOT/skills"
CLAUDE_SKILLS="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
CODEX_SKILLS="${CODEX_SKILLS_DIR:-$HOME/.agents/skills}"
ARCHIVE_BASE="$REPO_ROOT/.local/archived-discovery-entries"
STAMP="$(date -u +%Y%m%dT%H%M%SZ)"

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

ensure_dir() {
  mkdir -p "$1"
}

link_one() {
  local dest_root="$1"
  local skill="$2"
  local source="$SKILLS_DIR/$skill"
  local dest="$dest_root/$skill"

  if [[ ! -f "$source/SKILL.md" ]]; then
    echo "Missing source skill: $source/SKILL.md" >&2
    return 1
  fi

  ensure_dir "$dest_root"

  if [[ -L "$dest" ]]; then
    rm "$dest"
  elif [[ -e "$dest" ]]; then
    local archive="$ARCHIVE_BASE/$STAMP$(echo "$dest" | sed 's#/#_#g')"
    mkdir -p "$(dirname "$archive")"
    mv "$dest" "$archive"
    echo "Archived existing real entry: $dest -> $archive"
  fi

  ln -s "$source" "$dest"
  echo "Linked $dest -> $source"
}

remove_legacy_aliases() {
  local dest_root="$1"
  local alias="$dest_root/ARR-to-bookings"
  if [[ -L "$alias" ]]; then
    rm "$alias"
    echo "Removed legacy duplicate alias: $alias"
  fi
}

for root in "$CLAUDE_SKILLS" "$CODEX_SKILLS"; do
  ensure_dir "$root"
  remove_legacy_aliases "$root"
  for skill in "${SKILLS[@]}"; do
    link_one "$root" "$skill"
  done
done

echo "Done. Claude and Codex discovery entries now point at $SKILLS_DIR"
