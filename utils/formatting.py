from datetime import date, timedelta

PRINCIPLE_NAMES = {
    1: "Activeer relevante voorkennis",
    2: "Geef duidelijke, gestructureerde en uitdagende instructie",
    3: "Gebruik voorbeelden",
    4: "Combineer woord en beeld",
    5: "Laat leerstof actief verwerken",
    6: "Achterhaal of de hele klas het begrepen heeft",
    7: "Ondersteun bij moeilijke opdrachten",
    8: "Spreid oefening met leerstof in de tijd",
    9: "Zorg voor afwisseling in oefentypes",
    10: "Gebruik toetsing als leer- en oefenstrategie",
    11: "Geef feedback die leerlingen aan het denken zet",
    12: "Leer je leerlingen effectief leren",
}

def status_for_score(score: int) -> str:
    return "green" if score >= 3 else "orange" if score == 2 else "red"

def next_spacing_dates(start: date) -> list[date]:
    return [start + timedelta(days=d) for d in (2, 7, 21)]
