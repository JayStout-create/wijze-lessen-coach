#!/bin/bash
cd "$(dirname "$0")"
source /opt/anaconda3/bin/activate base
python -m streamlit run app.py