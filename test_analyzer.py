from ai.analyzer import extract_lesson_from_text


lesson = """
Les Nederlands – 5DF

Onderwerp:
Protagonist en antagonist in Reynaert de Vos.

Leerdoel:
De leerlingen kunnen na de les de protagonist en antagonist
in een verhaal herkennen en hun antwoord met voorbeelden uit
de tekst onderbouwen.

Lesfase 1 – Voorkennis activeren – 5 minuten
De leerlingen schrijven individueel op wat ze nog weten over
protagonist en antagonist. Daarna vergelijken ze hun antwoord
met een buur. Vervolgens bespreken we enkele antwoorden klassikaal.

Lesfase 2 – Instructie – 10 minuten
Ik leg de begrippen protagonist en antagonist uit op het bord.
Ik gebruik Reynaert en Isegrim als voorbeelden.

Lesfase 3 – Samen oefenen – 15 minuten
We bekijken samen een fragment uit Reynaert de Vos.
Ik denk hardop na over de vraag: wie is hier de protagonist?
Daarna zoeken de leerlingen zelf bewijs in de tekst.

Lesfase 4 – Individuele verwerking – 15 minuten
De leerlingen krijgen drie korte fragmenten.
Per fragment bepalen ze de protagonist en antagonist en schrijven
ze een korte motivatie.

Lesfase 5 – Exit-ticket – 5 minuten
Iedere leerling beantwoordt individueel:

1. Wat is een protagonist?
2. Wat is een antagonist?
3. Wie is de protagonist in Reynaert de Vos en waarom?

Ik verzamel de exit-tickets en bekijk welke leerlingen de begrippen
nog niet beheersen.
"""


print("========================================")
print("WIJZE LESSEN COACH - EXTRACTIE TEST")
print("========================================")
print()
print("Gemma 4 gaat de les formatteren...")
print("Dit kan enkele minuten duren.")
print()


try:

    result = extract_lesson_from_text(
        lesson_text=lesson,
        model="gemma4:e4b"
    )

except Exception as error:

    print()
    print("========================================")
    print("FOUT")
    print("========================================")
    print()
    print(type(error).__name__)
    print()
    print(str(error))
    print()

    raise


print()
print("========================================")
print("RESULTAAT")
print("========================================")
print()

print("TITEL:")
print(result.get("title"))

print()
print("ONDERWERP:")
print(result.get("topic"))

print()
print("LEERDOELEN:")
print(result.get("learning_objectives"))

print()
print("DUUR:")
print(result.get("duration"))

print()
print("========================================")
print("LESFASEN")
print("========================================")
print()

phases = result.get("phases", [])

print(f"Aantal fasen: {len(phases)}")
print()

for number, phase in enumerate(phases, start=1):

    print("----------------------------------------")
    print(f"FASE {number}")
    print("----------------------------------------")

    print("Titel:")
    print(phase.get("title"))

    print()

    print("Duur:")
    print(phase.get("duration"))

    print()

    print("LERAAR:")
    print(phase.get("teacher_activity"))

    print()

    print("LEERLINGEN:")
    print(phase.get("student_activity"))

    print()

    print("WERKVORM:")
    print(phase.get("activity_type"))

    print()

    print("MATERIAAL:")
    print(phase.get("materials"))

    print()


print("========================================")
print("TEST AFGEROND")
print("========================================")
