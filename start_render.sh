#!/usr/bin/env bash
set -euo pipefail

PORT="${PORT:-8501}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${SCRIPT_DIR}"

exec streamlit run dashboard_achados_criticos.py \
  --server.address 0.0.0.0 \
  --server.port "${PORT}" \
  --server.headless true \
  --server.enableCORS=false \
  --server.enableXsrfProtection=false \
  --server.maxUploadSize=400 \
  --browser.gatherUsageStats false \
  --theme.base dark \
  --theme.primaryColor "#667eea" \
  --theme.backgroundColor "#0e1117" \
  --theme.secondaryBackgroundColor "#262730"
