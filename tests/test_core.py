import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from database.db import init_db, SessionLocal
from database.models import Principle
from ai.prompt import SYSTEM_PROMPT
from ai.analyzer import _extract_json
from sqlalchemy import select


def test_principles():
    init_db()
    s = SessionLocal()
    principles = s.scalars(select(Principle).order_by(Principle.number)).all()
    assert len(principles) == 12
    assert principles[0].number == 1
    assert principles[-1].number == 12
    s.close()


def test_json_parser():
    result = _extract_json('{"overall_score": 12, "principles": []}')
    assert result["overall_score"] == 12


def test_prompt():
    assert "DE 12 BOUWSTENEN" in SYSTEM_PROMPT
    assert "formatieve evaluatie" in SYSTEM_PROMPT.lower()
