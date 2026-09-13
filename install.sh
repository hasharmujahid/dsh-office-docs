#!/usr/bin/env bash
# install.sh - put the office-docs skill into a DSH skills root.
#
# Copies this skill directory into a discovered skills root and creates the
# shared Python environment once. Idempotent: re-running updates the copy in
# place and leaves a working environment alone.
#
#   ./install.sh                    # install to ~/.dsh/skills/office-docs
#   ./install.sh --link             # symlink instead of copy (for development)
#   ./install.sh --root ~/.agents/skills
#   ./install.sh --project .        # install for one project: ./.dsh/skills
#
# Skills are discovered only from these roots: <project>/.dsh/skills,
# <project>/.agents/skills, $DSH_HOME/skills, $DSH_AGENTS_HOME/skills.

set -euo pipefail

SOURCE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DSH_HOME_DIR="${DSH_HOME:-$HOME/.dsh}"
ROOT="$DSH_HOME_DIR/skills"
MODE="copy"
PROJECT=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --link) MODE="link"; shift ;;
    --copy) MODE="copy"; shift ;;
    --root) ROOT="$2"; shift 2 ;;
    --project) PROJECT="$2"; shift 2 ;;
    -h|--help) sed -n '2,20p' "${BASH_SOURCE[0]}"; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

if [[ -n "$PROJECT" ]]; then
  ROOT="$PROJECT/.dsh/skills"
fi

TARGET="$ROOT/office-docs"
mkdir -p "$ROOT"

echo "Source : $SOURCE_DIR"
echo "Target : $TARGET"

if [[ -L "$TARGET" || -e "$TARGET" ]]; then
  if [[ "$MODE" == "link" && -L "$TARGET" ]]; then
    rm -f "$TARGET"
  else
    echo "Removing existing install"
    rm -rf "$TARGET"
  fi
fi

if [[ "$MODE" == "link" ]]; then
  ln -s "$SOURCE_DIR" "$TARGET"
  echo "Linked (edits to the source take effect immediately)"
else
  mkdir -p "$TARGET"
  # Copy the skill payload, not VCS metadata or caches.
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --exclude '.git' --exclude '__pycache__' --exclude '*.pyc' "$SOURCE_DIR/" "$TARGET/"
  else
    tar -C "$SOURCE_DIR" --exclude='.git' --exclude='__pycache__' -cf - . | tar -C "$TARGET" -xf -
  fi
  echo "Copied"
fi

echo
echo "Setting up the shared Python environment (once per machine)..."
python3 "$TARGET/scripts/bootstrap.py"

echo
echo "Verifying the toolchain..."
"$DSH_HOME_DIR/office-docs/venv/bin/python" "$TARGET/scripts/smoke_test.py" || {
  echo
  echo "Smoke test failed. The skill is installed but not healthy; see the failures above." >&2
  exit 1
}

cat <<EOF

Installed.

  skill   : $TARGET/SKILL.md
  python  : $DSH_HOME_DIR/office-docs/venv/bin/python
  recheck : $DSH_HOME_DIR/office-docs/venv/bin/python $TARGET/scripts/smoke_test.py

DSH discovers this skill on its next session start. To confirm it is registered,
ask a session in this workspace to list its skills.
EOF
