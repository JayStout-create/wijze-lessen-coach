from pathlib import Path
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "wijze_lessencoach.db"

engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)

class Base(DeclarativeBase):
    pass

def init_db():
    from . import models  # noqa: F401
    Base.metadata.create_all(bind=engine)
    seed_principles()

def seed_principles():
    from .models import Principle
    session = SessionLocal()
    try:
        if session.query(Principle).count() == 0:
            principles = [
                (1, "Activeer relevante voorkennis", "Activeer noodzakelijke voorkennis en spoor misconcepties op."),
                (2, "Geef duidelijke, gestructureerde en uitdagende instructie", "Maak doelen, structuur, instructie en succescriteria helder."),
                (3, "Gebruik voorbeelden", "Gebruik concrete voorbeelden, worked examples en modelling."),
                (4, "Combineer woord en beeld", "Gebruik visuele representaties wanneer die het begrip ondersteunen."),
                (5, "Laat leerstof actief verwerken", "Laat leerlingen informatie selecteren, organiseren, verklaren en toepassen."),
                (6, "Achterhaal of de hele klas het begrepen heeft", "Gebruik formatieve checks waarmee je individuele beheersing kunt vaststellen."),
                (7, "Ondersteun bij moeilijke opdrachten", "Gebruik tijdelijke scaffolding en bouw die geleidelijk af."),
                (8, "Spreid oefening met leerstof in de tijd", "Plan herhaling en retrieval over meerdere momenten."),
                (9, "Zorg voor afwisseling in oefentypes", "Gebruik betekenisvolle interleaving en verschillende probleemtypes."),
                (10, "Gebruik toetsing als leer- en oefenstrategie", "Gebruik retrieval practice als leerstrategie, niet alleen als evaluatie."),
                (11, "Geef feedback die leerlingen aan het denken zet", "Werk met feed-up, feedback en feed-forward."),
                (12, "Leer je leerlingen effectief leren", "Maak effectieve leerstrategieën en metacognitie expliciet."),
            ]
            session.add_all([Principle(number=n, name=name, short_description=desc) for n, name, desc in principles])
            session.commit()
    finally:
        session.close()
