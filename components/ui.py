import streamlit as st


def status_badge(status: str, score: int):
    labels = {"green": "🟢 Sterk", "orange": "🟠 Aandachtspunt", "red": "🔴 Prioriteit"}
    st.markdown(f"**{labels.get(status, status)} · {score}/4**")


def metric_card(label: str, value: str, help_text: str | None = None):
    st.metric(label, value, help=help_text)
