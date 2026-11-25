#!/bin/bash
# Rolling backup of every .pptx under ~/My Drive that changed since the last run.
# Keeps the 50 newest snapshots in total (all decks together); older ones are deleted.
# Run every 6 h by the LaunchAgent ro.victorrentea.pptx-backup (launchagent-backup.plist).
set -u
SRC="${PPTX_BACKUP_SRC:-$HOME/My Drive}"
DEST="${PPTX_BACKUP_DEST:-$HOME/Backups/pptx}"
KEEP="${PPTX_BACKUP_KEEP:-50}"
STATE="$DEST/.state"            # <mtime>\t<size>\t<path> as last backed up
mkdir -p "$DEST"; touch "$STATE"
now=$(date +%s); stamp=$(date +%Y-%m-%d_%H%M)
first_run=0; [ -s "$STATE" ] || first_run=1
new_state=$(mktemp); copied=0

while IFS= read -r -d '' f; do
  read -r mtime size < <(stat -f '%m %z' "$f")
  [ $((now - mtime)) -lt 120 ] && { grep -F "	$f" "$STATE" >> "$new_state"; continue; }  # still being saved: next run
  line="$mtime	$size	$f"
  if ! grep -qxF "$line" "$STATE"; then
    # first run: seed only with the decks touched in the last 7 days, not all of them
    if [ $first_run = 0 ] || [ $((now - mtime)) -lt 604800 ]; then
      rel="${f#"$SRC"/}"
      cp -p "$f" "$DEST/${stamp}__${rel//\//__}" && copied=$((copied + 1))
    fi
  fi
  echo "$line" >> "$new_state"
done < <(find "$SRC" -name '*.pptx' ! -name '~$*' ! -name '.~*' -type f -print0 2>/dev/null)
mv "$new_state" "$STATE"

# prune: names start with the timestamp, so lexical order = age
ls -1 "$DEST" | grep '\.pptx$' | sort -r | tail -n +$((KEEP + 1)) | while IFS= read -r old; do rm -f "$DEST/$old"; done
echo "$(date '+%F %T') copied=$copied kept=$(ls -1 "$DEST" | grep -c '\.pptx$')"
