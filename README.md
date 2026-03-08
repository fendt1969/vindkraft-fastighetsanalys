# Vindkraft Fastighetsanalys

Webbverktyg som beräknar hur planerade vindkraftverk påverkar fastighetsvärden i närområdet.

## Vad gör verktyget?

1. Hämtar vindkraftsprojekt och verkpositioner från **Vindbrukskollen** (Länsstyrelsen)
2. Söker namngivna platser/fastigheter via **OpenStreetMap**
3. Beräknar värdereduktion med modellen från **Westlund & Wilhelmsson (2021)**
4. Visar resultat på interaktiv karta med **Lantmäteriets** topografiska karta
5. Genererar nedladdningsbara PDF-rapporter och HTML-karta

## Modell

Beräkningen baseras på:

> Westlund, H. & Wilhelmsson, M. (2021). *The Socio-Economic Cost of Wind Turbines: A Swedish Case Study.* Sustainability, 13, 6892.

Studien analyserar ca 69 000 småhusförsäljningar i Sverige 2013–2018 och visar att fastighetsvärden minskar med närhet till vindkraftverk:

```
β(d) = −0.2811 × exp(−0.3811 × d)
Reduktion(d) = (exp(β(d)) − 1) × 100%
```

där *d* = avstånd i km till närmaste vindkraftverk.

## Datakällor

- **Vindbrukskollen** (Länsstyrelsen) — projektområden och verkpositioner
- **OpenStreetMap / Overpass API** — namngivna platser och bebyggelse
- **Lantmäteriet** — topografisk karta (CC BY 4.0)

## Kör lokalt

```bash
pip install -r requirements.txt
streamlit run app.py
```

Eller använd startskriptet:

```bash
bash start.sh
```

Öppna sedan http://localhost:8501 i webbläsaren.
