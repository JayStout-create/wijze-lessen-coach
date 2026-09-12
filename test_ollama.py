import requests

url = "http://localhost:11434/api/generate"

data = {
    "model": "gemma4:e4b",
    "prompt": "Geef één korte zin over het belang van retrieval practice. Antwoord uitsluitend in het Nederlands.",
    "stream": False
}

response = requests.post(
    url,
    json=data,
    timeout=180
)

response.raise_for_status()

result = response.json()

print(result["response"])
