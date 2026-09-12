#!/bin/bash

PROJECT_DIR="/Users/jay/Desktop/wijze-lessen-coach"
LOG_FILE="$PROJECT_DIR/auto_backup.log"

cd "$PROJECT_DIR" || exit 1

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Controle gestart." >> "$LOG_FILE"

# Controleer of er wijzigingen zijn
if [[ -n "$(git status --porcelain)" ]]; then

    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Wijzigingen gevonden. Backup maken..." >> "$LOG_FILE"

    git add .
    git commit -m "Automatische backup $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
    git push >> "$LOG_FILE" 2>&1

    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Backup naar GitHub voltooid." >> "$LOG_FILE"

else

    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Geen wijzigingen gevonden." >> "$LOG_FILE"

fi
