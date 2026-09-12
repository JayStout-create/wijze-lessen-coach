#!/bin/bash
set -e
cd "$(dirname "$0")"
source /opt/anaconda3/bin/activate
if conda env list | awk '{print $1}' | grep -qx "WijzeLessen"; then
  echo "Anaconda-omgeving WijzeLessen bestaat al."
else
  conda create -n WijzeLessen python=3.12 -y
fi
source /opt/anaconda3/bin/activate WijzeLessen
python -m pip install -r requirements.txt
python - <<'PY'
from database.db import init_db
init_db()
print("Database geïnitialiseerd.")
PY
echo
echo "Klaar. Start de app met:"
echo "source /opt/anaconda3/bin/activate WijzeLessen"
echo "streamlit run app.py"
