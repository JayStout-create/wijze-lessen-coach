import ast
import json
import re
import os

try:
    import ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False

try:
    from google import genai
    from google.genai import types
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False


# ============================================================
# GOOGLE GEMINI API-KEY
# ============================================================

GEMINI_API_KEY = None

try:
    import streamlit as st
    GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY")
except Exception:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if GEMINI_API_KEY:
    GEMINI_API_KEY = str(GEMINI_API_KEY).strip()


__all__ = [
    "analyze_lesson",
    "create_didactic_draft",
    "extract_lesson_from_text",
    "assign_curriculum_goals",
    "generate_student_goals",
    "generate_spaced_retrieval",
    "get_models",
    "get_default_model",
    "analyseer_lesvoorbereiding",
    "ollama_available",
    "google_available",
    "ACTIVITY_OPTIONS",
    "MATERIAL_OPTIONS",
    "PRINCIPLES_CONTEXT",
    "DEFAULT_MODEL",
]


FALLBACK_MODELS = [
    "qwen2.5:7b",
    "qwen2.5:14b",
    "qwen2.5:latest",
    "gemma4:e4b"
]


# ============================================================
# GOOGLE GEMINI MODELLEN
# ============================================================

GOOGLE_MODELS = [
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.1-pro"
]


ACTIVITY_OPTIONS = [
    "Instructie", "Retrieval practice", "Individueel werk", "Duo",
    "Groepswerk", "Klassikale bespreking", "Think-Pair-Share", "Quiz",
    "Wisbordjes", "Free recall", "Lezen", "Schrijven",
    "Luisteren", "Spreken", "Anders"
]


MATERIAL_OPTIONS = [
    "Geen specifiek materiaal", "Leerboek", "Handboek / cursus",
    "Werkblad", "Tekst", "Afbeelding", "Presentatie", "Digibord",
    "Video", "Audio", "Kahoot", "Wooclap", "BookWidgets",
    "Quiz", "Wisbordjes", "Flashcards", "Anders"
]


PRINCIPLES_CONTEXT = """
1. 🧩 Activeer relevante voorkennis (retrieval, misconcepties checken)
2. 🎯 Geef duidelijke, gestructureerde en uitdagende instructie (lesdoelen, succescriteria)
3. 🔍 Gebruik voorbeelden (uitgewerkte voorbeelden, modelling, hardop denken)
4. 🖼️ Combineer woord en beeld (dual coding, geen overbodige decoratie)
5. ⚙️ Laat leerstof actief verwerken (selecteren, organiseren, elaboreren, toepassen)
6. 🙋 Achterhaal of de hele klas het begrepen heeft (formatieve evaluatie, wisbordjes, gerichte checks)
7. 🪜 Ondersteun bij moeilijke opdrachten (scaffolding, stappenplannen, afbouw)
8. 🗓️ Spreid oefening met leerstof in de tijd (spacing, herhaling over lessen heen)
9. 🔀 Zorg voor afwisseling in oefentypes (interleaving, gevarieerde opgaven)
10. 📝 Gebruik toetsing als leer- en oefenstrategie (retrieval practice, quizzen, exit tickets)
11. 💬 Geef feedback die leerlingen aan het denken zet (feed-up, feedback, feed-forward)
12. 🧭 Leer je leerlingen effectief leren (metacognitie, studiestrategieën)
"""


def ollama_available() -> bool:
    """Controleert of de Ollama daemon draait en bereikbaar is."""
    if not OLLAMA_AVAILABLE:
        return False

    try:
        ollama.list()
        return True
    except Exception:
        return False


def google_available() -> bool:
    """Controleert of Google Gemini beschikbaar is."""
    return GOOGLE_AVAILABLE and bool(GEMINI_API_KEY)


def get_models(provider: str = "Lokaal (Ollama)") -> list:
    """Geeft de lijst van beschikbare modellen op basis van de gekozen provider."""

    if "google" in provider.lower() or "online" in provider.lower():
        return GOOGLE_MODELS

    found_models = []

    if OLLAMA_AVAILABLE:
        try:
            response = ollama.list()

            if hasattr(response, 'models'):
                for m in response.models:
                    model_name = getattr(
                        m,
                        'model',
                        getattr(m, 'name', None)
                    )

                    if model_name:
                        found_models.append(str(model_name))

            elif isinstance(response, dict) and 'models' in response:
                for m in response['models']:
                    model_name = m.get('model') or m.get('name')

                    if model_name:
                        found_models.append(str(model_name))

        except Exception:
            pass

    if not found_models:
        return FALLBACK_MODELS

    qwen_models = [
        m for m in found_models
        if "qwen2.5" in m.lower()
    ]

    other_models = [
        m for m in found_models
        if "qwen2.5" not in m.lower()
    ]

    return qwen_models + other_models


def get_default_model(provider: str = "Lokaal (Ollama)") -> str:

    models = get_models(provider)

    if "google" in provider.lower() or "online" in provider.lower():
        return "gemini-3.6-flash"

    return models[0] if models else "qwen2.5:7b"


DEFAULT_MODEL = get_default_model()


def _resolve_model(
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> str:

    if model_name and str(model_name).strip():
        return model_name

    return get_default_model(provider)


def _extract_raw_text(response) -> str:

    try:

        if hasattr(response, 'message') and hasattr(
            response.message,
            'content'
        ):

            if (
                response.message.content
                and str(response.message.content).strip()
            ):
                return str(response.message.content).strip()

        if hasattr(response, 'response'):

            if (
                response.response
                and str(response.response).strip()
            ):
                return str(response.response).strip()

    except Exception:
        pass

    parts = []

    try:

        if hasattr(response, 'model_dump'):
            r_dict = response.model_dump()

        elif hasattr(response, 'dict'):
            r_dict = response.dict()

        else:
            r_dict = dict(response)

    except Exception:
        r_dict = {}

    msg = r_dict.get('message')

    if isinstance(msg, dict):

        content = msg.get('content')

        if content and str(content).strip():
            return str(content).strip()

    for key in ['response', 'content']:

        val = r_dict.get(key)

        if val and str(val).strip():
            parts.append(str(val))

    if parts:
        return parts[0].strip()

    text = str(response)

    match = re.search(
        r"content='(.*?)'",
        text,
        re.DOTALL
    )

    if match:

        extracted = match.group(1)

        return (
            extracted
            .replace("\\n", "\n")
            .replace('\\"', '"')
            .replace("\\'", "'")
        )

    return text


def _clean_and_parse_json(raw_text: str) -> dict:

    if not raw_text or not raw_text.strip():

        raise ValueError(
            "Het AI-model gaf een lege respons terug."
        )

    text = re.sub(
        r'<thought>[\s\S]*?</thought>',
        '',
        raw_text,
        flags=re.IGNORECASE
    ).strip()

    text = re.sub(
        r'<think>[\s\S]*?</think>',
        '',
        text,
        flags=re.IGNORECASE
    ).strip()

    text = re.sub(
        r'<\|thought\|>[\s\S]*?<\|thought\|>',
        '',
        text,
        flags=re.IGNORECASE
    ).strip()

    text = re.sub(
        r'^```(?:json)?\s*',
        '',
        text,
        flags=re.MULTILINE
    )

    text = re.sub(
        r'\s*```$',
        '',
        text,
        flags=re.MULTILINE
    ).strip()

    start_idx = text.find('{')
    end_idx = text.rfind('}')

    if start_idx == -1 and text.find('[') != -1:

        start_idx = text.find('[')
        end_idx = text.rfind(']')

    if start_idx != -1:

        if end_idx != -1 and end_idx > start_idx:
            candidate = text[start_idx:end_idx + 1]

        else:
            candidate = text[start_idx:]

    else:

        candidate = text

    candidate = candidate.replace("\\'", "'")

    try:

        return json.loads(candidate)

    except Exception:
        pass

    try:

        result = ast.literal_eval(candidate.strip())

        if isinstance(result, (dict, list)):
            return result

    except Exception:
        pass

    raise ValueError(
        f"Kon de AI-uitvoer niet omzetten naar JSON.\n\n"
        f"Ontvangen tekst:\n{raw_text[:300]}..."
    )


# ============================================================
# LLM WRAPPERS
# LOKAAL / OLLAMA vs ONLINE / GOOGLE GEMINI
# ============================================================

def _ask_llm_text(
    prompt: str,
    provider: str = "Lokaal (Ollama)",
    model_name: str = None,
    custom_options: dict = None,
    system_instruction: str = ""
) -> str:

    is_google = (
        "google" in provider.lower()
        or "online" in provider.lower()
    )

    active_model = _resolve_model(
        model_name,
        provider
    )

    # --------------------------------------------------------
    # GOOGLE GEMINI
    # --------------------------------------------------------

    if is_google:

        if not GOOGLE_AVAILABLE:

            raise RuntimeError(
                "De Python-bibliotheek 'google-genai' "
                "is niet geïnstalleerd."
            )

        if not GEMINI_API_KEY:

            raise RuntimeError(
                "Geen GEMINI_API_KEY gevonden. "
                "Controleer .streamlit/secrets.toml."
            )

        client = genai.Client(
            api_key=GEMINI_API_KEY
        )

        config_args = {}

        if system_instruction:

            config_args['system_instruction'] = (
                system_instruction
            )

        try:

            response = client.models.generate_content(
                model=active_model,
                contents=prompt,
                config=(
                    types.GenerateContentConfig(
                        **config_args
                    )
                    if config_args
                    else None
                )
            )

            return (
                response.text.strip()
                if response.text
                else ""
            )

        except Exception as e:

            raise RuntimeError(
                f"Google Gemini API fout: {str(e)}"
            )

    # --------------------------------------------------------
    # OLLAMA
    # --------------------------------------------------------

    else:

        if not OLLAMA_AVAILABLE:

            raise RuntimeError(
                "De Python-bibliotheek 'ollama' "
                "is niet geïnstalleerd."
            )

        inhoud = ""

        ollama_options = {
            'temperature': 0.1,
            'top_p': 0.90,
            'num_predict': 4096,
            'num_ctx': 8192
        }

        if custom_options:
            ollama_options.update(custom_options)

        if not system_instruction:

            system_instruction = (
                "Je bent een strikte data-extractor. "
                "Volg de instructies tot op de letter "
                "en bedenk zelf GEEN extra informatie."
            )

        try:

            response = ollama.chat(
                model=active_model,
                messages=[
                    {
                        'role': 'system',
                        'content': system_instruction
                    },
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ],
                keep_alive='30m',
                options=ollama_options
            )

            inhoud = _extract_raw_text(response)

        except Exception:
            pass

        return inhoud.strip()


def _ask_llm_json(
    prompt: str,
    provider: str = "Lokaal (Ollama)",
    model_name: str = None,
    custom_options: dict = None
) -> dict | list:

    is_google = (
        "google" in provider.lower()
        or "online" in provider.lower()
    )

    active_model = _resolve_model(
        model_name,
        provider
    )

    system_instruction = (
        "Je bent een strikte automatische JSON-generator. "
        "Antwoord UITSLUITEND met valide JSON code. "
        "Geen inleidende of afsluitende tekst. "
        "Gebruik waar nodig null i.p.v. lege strings of Onbekend."
    )

    # --------------------------------------------------------
    # GOOGLE GEMINI
    # --------------------------------------------------------

    if is_google:

        if not GOOGLE_AVAILABLE:

            raise RuntimeError(
                "De Python-bibliotheek 'google-genai' "
                "is niet geïnstalleerd."
            )

        if not GEMINI_API_KEY:

            raise RuntimeError(
                "Geen GEMINI_API_KEY gevonden. "
                "Controleer .streamlit/secrets.toml."
            )

        client = genai.Client(
            api_key=GEMINI_API_KEY
        )

        try:

            response = client.models.generate_content(
                model=active_model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json",
                    temperature=0.05,
                )
            )

            return _clean_and_parse_json(
                response.text
            )

        except Exception as e:

            return {
                "error": str(e),
                "raw": getattr(
                    e,
                    'message',
                    str(e)
                )
            }

    # --------------------------------------------------------
    # OLLAMA
    # --------------------------------------------------------

    else:

        if not OLLAMA_AVAILABLE:

            raise RuntimeError(
                "De Python-bibliotheek 'ollama' "
                "is niet geïnstalleerd."
            )

        inhoud = ""

        ollama_options = {
            'temperature': 0.05,
            'top_p': 0.90,
            'num_predict': 8192,
            'num_ctx': 8192
        }

        if custom_options:
            ollama_options.update(custom_options)

        try:

            response = ollama.chat(
                model=active_model,
                messages=[
                    {
                        'role': 'system',
                        'content': system_instruction
                    },
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ],
                format='json',
                keep_alive='30m',
                options=ollama_options
            )

            inhoud = _extract_raw_text(response)

        except Exception:
            pass

        if not inhoud:

            try:

                response = ollama.generate(
                    model=active_model,
                    prompt=(
                        f"{system_instruction}\n\n"
                        f"{prompt}"
                    ),
                    format='json',
                    keep_alive='30m',
                    options=ollama_options
                )

                inhoud = _extract_raw_text(response)

            except Exception:
                pass

        try:

            return _clean_and_parse_json(
                inhoud
            )

        except Exception as e:

            return {
                "error": str(e),
                "raw": inhoud
            }


# ============================================================
# STAP 1: DIDACTISCHE EXTRACTIE
# ============================================================

def create_didactic_draft(
    raw_text: str,
    duration: int = 50,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> str:

    """Verwerkt ruwe notities via de expert-prompt naar een strakke didactische tabel."""

    if not raw_text or not str(raw_text).strip():
        return ""

    prompt = f"""
Je bent een ervaren didactisch expert, lerarenopleider en curriculumspecialist. Jouw taak is om op basis van de door mij geüploade of geplakte lesbronnen (werkbundel, presentatie, tekst en/of leerplancontext) een volledige, kant-en-klare lesvoorbereiding uit te schrijven.

### STRUCTUUR EN LAY-OUT VAN HET SJABLOON:
Je levert de lesvoorbereiding op als één overzichtelijke Markdown-tabel met EXACT deze 5 kolommen:
| FASENAAM (Bijv: Lesbegin, Lesfase 1: Uitleg, Lesafsluiting) | LESDOELEN | LEERINHOUD | WERKVORMEN / MEDIA / ORGANISATIE | TIJD |

De tabel volgt een logische chronologie:
- Start met **Lesbegin** (Instap, activeren voorkennis, kaderen van lesdoel).
- Gevolgd door opeenvolgende fasen in het **Lesmidden** (Lesfase 1: Inhoud, Lesfase 2, ... Zoveel als nodig voor elke afzonderlijke leeractiviteit).
- Eindig met **Lesafsluiting** (Synthese, formatieve evaluatie, exit-ticket).

### STRIKTE KWALITEITSEISEN PER KOLOM:

1. **LESDOELEN (Wat de leerling kan na de activiteit)**:
   - Formuleer per fase operationele doelen (observeerbare actiewerkwoorden/Bloom) rechtstreeks afgeleid van de handelingen uit de werkbundel.
   - Verzin NIETS als het niet in de bron staat. Nummer ze (D1, D2...).

2. **LEERINHOUD (Enkel het 'WAT' - Theorie & Verbetersleutel):
   - Bevat UITSLUITEND de vakkennis, theoriekernen, definities én de concrete modelantwoorden/verbetersleutel van oefeningen uit de tekst.
   - ⛔ STRIKT VERBODEN: Geen werkvormen, geen "OLG", geen "duo-werk", geen vragen van de leraar, en geen instructies over wat leerlingen moeten doen. (Kortom: inhoud handboek).

3. **WERKVORMEN / MEDIA / ORGANISATIE (Het 'HOE' - Didactisch Draaiboek):
   - LETTERLIJKE OPDRACHT OMSCHRIJVINGEN: Neem bij elke fase de letterlijke vraag en opdrachttekst of videonaam over! (CRUCIAAL)
   - Formuleer de letterlijke mondelinge vragen die de leerkracht stelt om het gesprek te leiden.
   - Benoem werkvorm, benodigde materialen (PPT, werkbundel pagina's, videofragmenten) en organisatie (rol leerkracht, klasopstelling).

4. **TIJD**:
   - Geef een inschatting van realistische tijdsduur per fase (als getal in minuten).
   - De som moet exact optellen tot de voorziene lestijd (ca. {duration} minuten). Staat er geen tijd? Schrijf "N/A".

---

### INVOERGEGEVENS VOOR DEZE LES:
- Totale verwachte lestijd: {duration} minuten
- Bronmateriaal (RUWE LESNOTITIES OM TE VERWERKEN):
{raw_text}

GENEREER UITSLUITEND DE MARKDOWN TABEL ZONDER ANDERE INLEIDENDE OF AFSLUITENDE TEKST.
"""

    return _ask_llm_text(
        prompt,
        provider=provider,
        model_name=model_name,
        system_instruction=(
            "Je bent een onderwijskundige expert die uitsluitend "
            "datagestuurde didactische Markdown-tabellen genereert."
        )
    )


# ============================================================
# STAP 2: PARSING VAN DE TABEL NAAR JSON
# ============================================================

def _run_json_extraction(
    markdown_table: str,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> dict:

    """Zet de door de AI gemaakte expert-tabel foutloos om in het formulier (JSON) van de web-app."""

    prompt = f"""
Zet de onderstaande opgemaakte Markdown-tabel foutloos om in een JSON-object.
VERZIN NIETS ZELF, VOEG NIETS TOE EN VAT NIETS SAM. Neem de inhoud van de kolommen 1-op-1 over in de juiste velden.

JSON-VELDEN BINNEN DE "phases" LIJST:
- "section": Bepaal of de rij logischerwijs behoort tot "LESINTRO", "LESMIDDEN", of "LESAFSLUITING" op basis van de titel.
- "title": (Neem Kolom 1 'FASENAAM' letterlijk over)
- "duration": (Neem het getal uit Kolom 5 'TIJD'. Als er "N/A" of "-" staat, geef dan de waarde 0 door)
- "lesson_objectives": (Neem Kolom 2 'LESDOELEN' over)
- "learning_content": (Neem Kolom 3 'LEERINHOUD' over)
- "organization": (Neem Kolom 4 'WERKVORMEN / MEDIA / ORGANISATIE' over)
- "activity_type": Maak zelf een inschatting op basis van het draaiboek en kies uit: {json.dumps(ACTIVITY_OPTIONS)}
- "materials": Maak zelf een inschatting op basis van het draaiboek en kies uit: {json.dumps(MATERIAL_OPTIONS)}

JSON FORMAT STRUCTUUR (Maak voor ELKE rij in de tabel een apart item aan in de "phases" lijst):
{{
  "phases": [
    {{
      "section": "...",
      "title": "...",
      "duration": 0,
      "lesson_objectives": "...",
      "learning_content": "...",
      "organization": "...",
      "activity_type": ["..."],
      "materials": ["..."]
    }}
  ],
  "title": "Nieuwe Les",
  "topic": "Algemeen",
  "duration": 50,
  "learning_objectives": "Plak hier alle 'lesson_objectives' uit alle rijen onder elkaar als één lange tekst.",
  "curriculum_goals": []
}}

UIT TE LEZEN TABEL:
{markdown_table}
"""

    custom_opts = {
        'temperature': 0.05,
        'num_predict': 8192
    }

    result = _ask_llm_json(
        prompt,
        provider=provider,
        model_name=model_name,
        custom_options=custom_opts
    )

    if isinstance(result, dict) and "error" not in result:

        raw_list = None

        for key in [
            "phases",
            "lesfasen",
            "fases",
            "fasen",
            "lesfases",
            "stappen"
        ]:

            if key in result and result[key]:
                raw_list = result[key]
                break

        if isinstance(raw_list, dict):
            result["phases"] = list(
                raw_list.values()
            )

        elif isinstance(raw_list, list):
            result["phases"] = raw_list

        if isinstance(result.get("phases"), list):

            for phase in result["phases"]:

                if not isinstance(phase, dict):
                    continue

                for key in [
                    "lesson_objectives",
                    "learning_content",
                    "organization"
                ]:

                    if phase.get(key) is None:
                        phase[key] = ""

                    else:
                        phase[key] = str(
                            phase[key]
                        )

                for key in [
                    "activity_type",
                    "materials"
                ]:

                    val = phase.get(key)

                    if isinstance(val, str):
                        phase[key] = [val]

                    elif not isinstance(val, list):
                        phase[key] = []

                try:

                    phase["duration"] = int(
                        phase.get("duration") or 0
                    )

                except (ValueError, TypeError):

                    phase["duration"] = 0

    return result


# ============================================================
# STAP 3: LEERPLANDOELEN KOPPELEN
# ============================================================

def assign_curriculum_goals(
    lesson_json: dict,
    available_goals: list,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> dict:

    """Laat de expert specifieke doelen selecteren uit de opgebouwde JSON."""

    if (
        not available_goals
        or not isinstance(lesson_json, dict)
        or not lesson_json.get("phases")
    ):
        return lesson_json

    compact_goals = [
        {
            "code": g["code"],
            "titel": g["omschrijving"][:90]
        }
        for g in available_goals[:25]
    ]

    lesson_summary = (
        str(
            lesson_json.get(
                'learning_objectives',
                ''
            )
        )
        + "\n"
        + "\n".join([
            p.get('learning_content', '')
            + " "
            + p.get('organization', '')
            for p in lesson_json.get(
                'phases',
                []
            )
        ])
    )

    prompt = f"""
Hieronder staat een overzicht van een uitgewerkte les en een lijst van leerplandoelen.
Selecteer op basis van de theorie en lesoefeningen de MAXIMAAL 3 best passende leerplandoel-codes. 

LESINHOUD:
{lesson_summary[:1500]}

MOGELIJKE DOELEN:
{json.dumps(compact_goals, ensure_ascii=False)}

Antwoord STRIKT met een JSON list met enkel de codes (bijv: ["BV3_02.01"]). 
Geen andere velden of textueel antwoord toegestaan!
"""

    custom_opts = {
        'temperature': 0.1,
        'num_predict': 1024
    }

    matched_goals = _ask_llm_json(
        prompt,
        provider=provider,
        model_name=model_name,
        custom_options=custom_opts
    )

    if isinstance(matched_goals, list):

        lesson_json["curriculum_goals"] = [
            str(c)
            for c in matched_goals
        ]

    elif isinstance(matched_goals, dict):

        for key, val in matched_goals.items():

            if isinstance(val, list):

                lesson_json["curriculum_goals"] = [
                    str(c)
                    for c in val
                ]

                break

    return lesson_json


# ============================================================
# MASTER FUNCTIE: 3-STAPPEN PIJPLIJN UITVOEREN
# ============================================================

def extract_lesson_from_text(
    ruwe_tekst: str,
    model_name: str = None,
    available_goals: list = None,
    provider: str = "Lokaal (Ollama)",
    **kwargs
) -> dict:

    """Overkoepelende Workflow."""

    schone_brontekst = (
        str(ruwe_tekst)
        .replace('< &', '(&')
        .replace('<', '&lt;')
    )

    les_duur = kwargs.get(
        "duration",
        50
    )

    didactic_draft = create_didactic_draft(
        schone_brontekst,
        duration=les_duur,
        model_name=model_name,
        provider=provider
    )

    tekst_te_parsen = (
        didactic_draft
        if didactic_draft
        and len(didactic_draft) > 50
        else schone_brontekst
    )

    result_json = _run_json_extraction(
        tekst_te_parsen,
        model_name=model_name,
        provider=provider
    )

    if (
        isinstance(result_json, dict)
        and not result_json.get("phases")
    ):

        return {
            "error": (
                "Kan JSON niet parsen uit het AI antwoord. "
                "Probeer opnieuw."
            )
        }

    if (
        available_goals
        and isinstance(result_json, dict)
        and result_json.get("phases")
    ):

        result_json = assign_curriculum_goals(
            result_json,
            available_goals,
            model_name=model_name,
            provider=provider
        )

    return result_json


# ============================================================
# POWERPOINT SUCCESCRITERIA
# ============================================================

def generate_student_goals(
    learning_objectives: str,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> list:

    if (
        not learning_objectives
        or not str(learning_objectives).strip()
    ):

        return [
            "Ik kan de lesinhoud begrijpen en toepassen."
        ]

    prompt = f"""
Vertaal onderstaande formele lesdoelen naar 2 tot 4 concrete succescriteria in leerlingentaal voor op een PowerPoint.

Formele lesdoelen:
{learning_objectives}

REGELS:
1. Formuleer elk doel als een zin die begint met 'Ik kan...'.
2. Geef UITSLUITEND een JSON-object terug:
{{
  "goals": [
    "Ik kan de zender en ontvanger aanduiden.",
    "Ik kan een voorbeeld geven van ruis."
  ]
}}
"""

    custom_opts = {
        'temperature': 0.1,
        'num_ctx': 8192
    }

    try:

        res = _ask_llm_json(
            prompt,
            provider=provider,
            model_name=model_name,
            custom_options=custom_opts
        )

        if isinstance(res, dict):

            goals = (
                res.get("goals")
                or res.get("doelen")
                or res.get("succescriteria")
            )

            if isinstance(goals, list) and goals:

                return [
                    str(g).strip()
                    for g in goals
                    if str(g).strip()
                ]

            for v in res.values():

                if isinstance(v, list) and v:

                    return [
                        str(g).strip()
                        for g in v
                        if str(g).strip()
                    ]

        elif isinstance(res, list) and res:

            return [
                str(g).strip()
                for g in res
                if str(g).strip()
            ]

    except Exception:
        pass

    lines = [
        line.strip().lstrip(
            '-•123456789. '
        )
        for line in str(
            learning_objectives
        ).split('\n')
        if line.strip()
    ]

    cleaned_goals = []

    for line in lines[:4]:

        if not line.lower().startswith(
            "ik kan"
        ):

            cleaned_goals.append(
                f"Ik kan "
                f"{line[0].lower() + line[1:] if len(line) > 1 else line.lower()}"
            )

        else:

            cleaned_goals.append(line)

    return (
        cleaned_goals
        if cleaned_goals
        else [
            "Ik kan de doelstellingen van deze les toepassen."
        ]
    )


# ============================================================
# SPACED RETRIEVAL OEFENINGEN
# ============================================================

def generate_spaced_retrieval(
    lesson_text: str,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)"
) -> list:

    if (
        not lesson_text
        or not str(lesson_text).strip()
    ):

        return []

    words = str(lesson_text).split()

    compact_text = (
        " ".join(words[:1200])
        if len(words) > 1200
        else str(lesson_text)
    )

    prompt = f"""
Ontwerp op basis van onderstaande les TWEE spaced retrieval-oefeningen.

OEFENING 1 (Volgende les): Korte termijn ophalen.
OEFENING 2 (Over 2 weken): Toepassing/Transfer.

GEEF UITSLUITEND JSON TERUG:
{{
  "exercises": [
    {{
      "timing": "Volgende les (Les + 1)",
      "activity_type": "Wisbordjes",
      "prompt": "Vraag...",
      "expected_answer": "Antwoord...",
      "rationale": "Versterkt..."
    }},
    {{
      "timing": "Over 2 à 3 weken (Les + 3)",
      "activity_type": "Duo",
      "prompt": "Vraag...",
      "expected_answer": "Antwoord...",
      "rationale": "Transfer..."
    }}
  ]
}}

Lesvoorbereiding:
{compact_text}
"""

    custom_opts = {
        'temperature': 0.3,
        'num_ctx': 8192,
        'num_predict': 1500
    }

    try:

        res = _ask_llm_json(
            prompt,
            provider=provider,
            model_name=model_name,
            custom_options=custom_opts
        )

        if isinstance(res, dict):

            exs = (
                res.get("exercises")
                or res.get("oefeningen")
                or res.get("spaced_retrieval")
            )

            if isinstance(exs, list) and len(exs) >= 2:
                return exs

            for v in res.values():

                if (
                    isinstance(v, list)
                    and len(v) >= 2
                    and isinstance(v[0], dict)
                ):

                    return v

        elif isinstance(res, list) and len(res) >= 2:

            return res

    except Exception:
        pass

    return [
        {
            "timing": "Volgende les (Les + 1)",
            "activity_type": "Wisbordjes",
            "prompt": "Schrijf de belangrijkste begrippen van de vorige les op.",
            "expected_answer": "Leerlingen reproduceren theorie zonder notities.",
            "rationale": "Activeert het geheugenspoor."
        },
        {
            "timing": "Over 2 à 3 weken (Les + 3)",
            "activity_type": "Duo",
            "prompt": "Pas de concepten toe op een nieuw praktijkvoorbeeld.",
            "expected_answer": "Leerlingen herkennen theorie in een onbekende context.",
            "rationale": "Stimuleert transfer en langetermijngeheugen."
        }
    ]


# ============================================================
# EVALUATIE 12 DIDACTISCHE BOUWSTENEN
# ============================================================

def analyze_lesson(
    les_tekst: str,
    model_name: str = None,
    provider: str = "Lokaal (Ollama)",
    **kwargs
) -> dict:

    prompt = f"""
Evalueer onderstaande les op de 12 didactische bouwstenen van 'Wijze Lessen'.

DE 12 BOUWSTENEN OM TE BEOORDELEN:
{PRINCIPLES_CONTEXT}

BEOORDELINGSREGELS:
1. Evalueer ALLE 12 bouwstenen in de array 'principles'.
2. Geef score 1 tot 5. 'total_score' is de som.
3. 'evidence': Citaat of observatie uit de les.
4. Geef voor ELKE bouwsteen DRIE concrete suggesties.
5. 'priorities': De 3 belangrijkste actiepunten.

GEEF UITSLUITEND JSON TERUG:
{{
  "total_score": 45,
  "summary": "Samenvatting...",
  "principles": [
    {{
      "number": 1,
      "score": 4,
      "status": "Goed",
      "evidence": "Bewijs...",
      "suggestion_1": "Suggestie 1...",
      "suggestion_2": "Suggestie 2...",
      "suggestion_3": "Suggestie 3..."
    }}
  ],
  "priorities": [
    "1. Prio 1",
    "2. Prio 2",
    "3. Prio 3"
  ]
}}

Lesvoorbereiding:
{les_tekst}
"""

    custom_opts = {
        'temperature': 0.1,
        'num_ctx': 8192,
        'num_predict': 4096
    }

    return _ask_llm_json(
        prompt,
        provider=provider,
        model_name=model_name,
        custom_options=custom_opts
    )


analyseer_lesvoorbereiding = analyze_lesson