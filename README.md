# Wijze Lessen Coach

Lokale Streamlit-app voor het ontwerpen, analyseren en verbeteren van lesvoorbereidingen op basis van de 12 bouwstenen uit *Wijze Lessen*.

## Vereisten
- macOS
- Anaconda/Miniconda
- Python 3.12+
- Streamlit
- Optioneel: Ollama voor lokale AI-analyse

## Starten
```bash
cd ~/Desktop/wijze-lessen-coach
conda create -n WijzeLessen python=3.12 -y
conda activate WijzeLessen
pip install -r requirements.txt
streamlit run app.py
```

Voor lokale AI met Ollama: installeer Ollama en zorg dat een model beschikbaar is, bijvoorbeeld `ollama pull llama3.2`.

## Snelle start
`start.command` activeert de Anaconda-omgeving en start Streamlit.

## Belangrijk
De app werkt ook zonder AI-model. In dat geval kun je lessen opslaan en handmatig bekijken; voor automatische analyse is Ollama vereist.
