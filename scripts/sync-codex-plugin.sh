#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PLUGIN_NAME="investor-toolkit"
MARKETPLACE_NAME="investor-toolkit-local"
MARKETPLACE_ROOT="$REPO_ROOT"

python3 -m json.tool "$REPO_ROOT/.codex-plugin/plugin.json" >/dev/null
python3 -m json.tool "$REPO_ROOT/.agents/plugins/marketplace.json" >/dev/null

if python3 - <<'PY'
try:
    import yaml  # noqa: F401
except Exception:
    raise SystemExit(1)
PY
then
  python3 /Users/lyndsay/.codex/skills/.system/plugin-creator/scripts/validate_plugin.py "$REPO_ROOT"
  python3 /Users/lyndsay/.codex/skills/.system/plugin-creator/scripts/update_plugin_cachebuster.py "$REPO_ROOT"
else
  echo "PyYAML is not installed, so skipping Codex plugin helper validation/cachebuster."
  echo "JSON manifests are valid; continuing with marketplace/plugin install."
fi

codex plugin marketplace add "$MARKETPLACE_ROOT" || true
codex plugin add "${PLUGIN_NAME}@${MARKETPLACE_NAME}"

echo "Done. Start a new Codex thread to pick up updated plugin skills."
