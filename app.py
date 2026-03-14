#!/usr/bin/env python3
"""
Vindkraft Fastighetsanalys — Streamlit Web App (self-contained)
All API calls, model calculations, map/PDF generation in one file.
"""

import streamlit as st
import folium
import streamlit.components.v1 as components
import json
import math
import os
import tempfile
import time
import urllib.parse
import urllib.request
from datetime import date

# ─── Configuration ────────────────────────────────────────────────────────

VINDBRUKSKOLLEN_BASE = (
    "https://ext-geodata-applikationer.lansstyrelsen.se"
    "/arcgis/rest/services/VBK/lst_vbk_wms_vindbrukskollen/MapServer"
)
PROJECT_AREA_LAYER = 2
TURBINE_LAYERS = {
    5: "Uppförda", 6: "Beviljade", 7: "Avslagna", 8: "Handläggs",
    9: "Nedmonterade", 10: "Överklagade", 11: "Uppgift saknas", 12: "Inte aktuella",
}
OVERPASS_URL = "https://overpass-api.de/api/interpreter"
DEFAULT_RADIUS_M = 3000
DEFAULT_ASSUMED_VALUE_TKR = 3500

# ─── W&W 2025 Regional Model ─────────────────────────────────────────────
# Westlund & Wilhelmsson (2025), Table 3: hedonic coefficients per 2-km band.
# Percentage = 100 * [exp(beta) - 1], Halvorsen & Palmquist (1980).
# Only statistically significant (p<0.05) negative coefficients are used.
# Band key = midpoint km (1, 3, 5, 7).

WW2025_COEFF = {
    # Region: {midpoint_km: beta}  — only significant negative betas
    "south": {1: -0.183, 3: -0.111},                  # sign. 0-2, 2-4
    "east":  {1: -0.091, 3: -0.102, 5: -0.046},       # sign. 0-2, 2-4, 4-6
    "north": {1: -0.128},                              # sign. 0-2 only
}

# ─── W&W 2021 National Model ─────────────────────────────────────────────
# Westlund & Wilhelmsson (2021): β(d) = −0.2811 × exp(−0.3811 × d)
# ~69,000 transactions, 2013-2018, hedonic cross-section, single national curve.
WW2021_A = -0.2811
WW2021_B = -0.3811

MODEL_OPTIONS = {
    "W&W 2025 Regional": "ww2025",
    "W&W 2021 National": "ww2021",
}

# kommun → NUTS1 region mapping via län
# SE2 = south, SE1 = east, SE3 = north
_LAN_NUTS1 = {
    "Stockholm": "east", "Uppsala": "east", "Södermanland": "east",
    "Östergötland": "east", "Örebro": "east", "Västmanland": "east",
    "Jönköping": "south", "Kronoberg": "south", "Kalmar": "south",
    "Gotland": "south", "Blekinge": "south", "Skåne": "south",
    "Halland": "south", "Västra Götaland": "south",
    "Värmland": "north", "Dalarna": "north", "Gävleborg": "north",
    "Västernorrland": "north", "Jämtland": "north",
    "Västerbotten": "north", "Norrbotten": "north",
}

# All 290 kommuner → län (used to resolve NUTS1 from KOMNAMN)
_KOMMUN_LAN = {
    # Stockholm
    "Botkyrka": "Stockholm", "Danderyd": "Stockholm", "Ekerö": "Stockholm",
    "Haninge": "Stockholm", "Huddinge": "Stockholm", "Järfälla": "Stockholm",
    "Lidingö": "Stockholm", "Nacka": "Stockholm", "Norrtälje": "Stockholm",
    "Nykvarn": "Stockholm", "Nynäshamn": "Stockholm", "Salem": "Stockholm",
    "Sigtuna": "Stockholm", "Sollentuna": "Stockholm", "Solna": "Stockholm",
    "Stockholm": "Stockholm", "Sundbyberg": "Stockholm", "Södertälje": "Stockholm",
    "Tyresö": "Stockholm", "Täby": "Stockholm", "Upplands Väsby": "Stockholm",
    "Upplands-Bro": "Stockholm", "Vallentuna": "Stockholm", "Vaxholm": "Stockholm",
    "Värmdö": "Stockholm", "Österåker": "Stockholm",
    # Uppsala
    "Enköping": "Uppsala", "Heby": "Uppsala", "Håbo": "Uppsala",
    "Knivsta": "Uppsala", "Tierp": "Uppsala", "Uppsala": "Uppsala",
    "Älvkarleby": "Uppsala", "Östhammar": "Uppsala",
    # Södermanland
    "Eskilstuna": "Södermanland", "Flen": "Södermanland", "Gnesta": "Södermanland",
    "Katrineholm": "Södermanland", "Nyköping": "Södermanland",
    "Oxelösund": "Södermanland", "Strängnäs": "Södermanland",
    "Trosa": "Södermanland", "Vingåker": "Södermanland",
    # Östergötland
    "Boxholm": "Östergötland", "Finspång": "Östergötland", "Kinda": "Östergötland",
    "Linköping": "Östergötland", "Mjölby": "Östergötland", "Motala": "Östergötland",
    "Norrköping": "Östergötland", "Söderköping": "Östergötland",
    "Vadstena": "Östergötland", "Valdemarsvik": "Östergötland",
    "Ydre": "Östergötland", "Åtvidaberg": "Östergötland",
    "Ödeshög": "Östergötland",
    # Jönköping
    "Aneby": "Jönköping", "Eksjö": "Jönköping", "Gislaved": "Jönköping",
    "Gnosjö": "Jönköping", "Habo": "Jönköping", "Jönköping": "Jönköping",
    "Mullsjö": "Jönköping", "Nässjö": "Jönköping", "Sävsjö": "Jönköping",
    "Tranås": "Jönköping", "Vaggeryd": "Jönköping", "Vetlanda": "Jönköping",
    "Värnamo": "Jönköping",
    # Kronoberg
    "Alvesta": "Kronoberg", "Lessebo": "Kronoberg", "Ljungby": "Kronoberg",
    "Markaryd": "Kronoberg", "Tingsryd": "Kronoberg", "Uppvidinge": "Kronoberg",
    "Växjö": "Kronoberg", "Älmhult": "Kronoberg",
    # Kalmar
    "Borgholm": "Kalmar", "Emmaboda": "Kalmar", "Hultsfred": "Kalmar",
    "Högsby": "Kalmar", "Kalmar": "Kalmar", "Mönsterås": "Kalmar",
    "Mörbylånga": "Kalmar", "Nybro": "Kalmar", "Oskarshamn": "Kalmar",
    "Torsås": "Kalmar", "Vimmerby": "Kalmar", "Västervik": "Kalmar",
    # Gotland
    "Gotland": "Gotland",
    # Blekinge
    "Karlshamn": "Blekinge", "Karlskrona": "Blekinge", "Olofström": "Blekinge",
    "Ronneby": "Blekinge", "Sölvesborg": "Blekinge",
    # Skåne
    "Bjuv": "Skåne", "Bromölla": "Skåne", "Burlöv": "Skåne",
    "Båstad": "Skåne", "Eslöv": "Skåne", "Helsingborg": "Skåne",
    "Hässleholm": "Skåne", "Höganäs": "Skåne", "Hörby": "Skåne",
    "Höör": "Skåne", "Klippan": "Skåne", "Kristianstad": "Skåne",
    "Kävlinge": "Skåne", "Landskrona": "Skåne", "Lomma": "Skåne",
    "Lund": "Skåne", "Malmö": "Skåne", "Osby": "Skåne",
    "Perstorp": "Skåne", "Simrishamn": "Skåne", "Sjöbo": "Skåne",
    "Skurup": "Skåne", "Staffanstorp": "Skåne", "Svalöv": "Skåne",
    "Svedala": "Skåne", "Tomelilla": "Skåne", "Trelleborg": "Skåne",
    "Vellinge": "Skåne", "Ystad": "Skåne", "Åstorp": "Skåne",
    "Ängelholm": "Skåne", "Örkelljunga": "Skåne", "Östra Göinge": "Skåne",
    # Halland
    "Falkenberg": "Halland", "Halmstad": "Halland", "Hylte": "Halland",
    "Kungsbacka": "Halland", "Laholm": "Halland", "Varberg": "Halland",
    # Västra Götaland
    "Ale": "Västra Götaland", "Alingsås": "Västra Götaland",
    "Bengtsfors": "Västra Götaland", "Bollebygd": "Västra Götaland",
    "Borås": "Västra Götaland", "Dals-Ed": "Västra Götaland",
    "Essunga": "Västra Götaland", "Falköping": "Västra Götaland",
    "Färgelanda": "Västra Götaland", "Grästorp": "Västra Götaland",
    "Gullspång": "Västra Götaland", "Göteborg": "Västra Götaland",
    "Götene": "Västra Götaland", "Herrljunga": "Västra Götaland",
    "Hjo": "Västra Götaland", "Härryda": "Västra Götaland",
    "Karlsborg": "Västra Götaland", "Kungälv": "Västra Götaland",
    "Lerum": "Västra Götaland", "Lidköping": "Västra Götaland",
    "Lilla Edet": "Västra Götaland", "Lysekil": "Västra Götaland",
    "Mariestad": "Västra Götaland", "Mark": "Västra Götaland",
    "Mellerud": "Västra Götaland", "Munkedal": "Västra Götaland",
    "Mölndal": "Västra Götaland", "Orust": "Västra Götaland",
    "Partille": "Västra Götaland", "Skara": "Västra Götaland",
    "Skövde": "Västra Götaland", "Sotenäs": "Västra Götaland",
    "Stenungsund": "Västra Götaland", "Strömstad": "Västra Götaland",
    "Svenljunga": "Västra Götaland", "Tanum": "Västra Götaland",
    "Tibro": "Västra Götaland", "Tidaholm": "Västra Götaland",
    "Tjörn": "Västra Götaland", "Tranemo": "Västra Götaland",
    "Trollhättan": "Västra Götaland", "Töreboda": "Västra Götaland",
    "Uddevalla": "Västra Götaland", "Ulricehamn": "Västra Götaland",
    "Vara": "Västra Götaland", "Vårgårda": "Västra Götaland",
    "Vänersborg": "Västra Götaland", "Åmål": "Västra Götaland",
    "Öckerö": "Västra Götaland",
    # Värmland
    "Arvika": "Värmland", "Eda": "Värmland", "Filipstad": "Värmland",
    "Forshaga": "Värmland", "Grums": "Värmland", "Hagfors": "Värmland",
    "Hammarö": "Värmland", "Karlstad": "Värmland", "Kil": "Värmland",
    "Kristinehamn": "Värmland", "Munkfors": "Värmland", "Storfors": "Värmland",
    "Sunne": "Värmland", "Säffle": "Värmland", "Torsby": "Värmland",
    "Årjäng": "Värmland",
    # Örebro
    "Askersund": "Örebro", "Degerfors": "Örebro", "Hallsberg": "Örebro",
    "Hällefors": "Örebro", "Karlskoga": "Örebro", "Kumla": "Örebro",
    "Laxå": "Örebro", "Lekeberg": "Örebro", "Lindesberg": "Örebro",
    "Ljusnarsberg": "Örebro", "Nora": "Örebro", "Örebro": "Örebro",
    # Västmanland
    "Arboga": "Västmanland", "Fagersta": "Västmanland", "Hallstahammar": "Västmanland",
    "Kungsör": "Västmanland", "Köping": "Västmanland", "Norberg": "Västmanland",
    "Sala": "Västmanland", "Skinnskatteberg": "Västmanland",
    "Surahammar": "Västmanland", "Västerås": "Västmanland",
    # Dalarna
    "Avesta": "Dalarna", "Borlänge": "Dalarna", "Falun": "Dalarna",
    "Gagnef": "Dalarna", "Hedemora": "Dalarna", "Leksand": "Dalarna",
    "Ludvika": "Dalarna", "Malung-Sälen": "Dalarna", "Mora": "Dalarna",
    "Orsa": "Dalarna", "Rättvik": "Dalarna", "Smedjebacken": "Dalarna",
    "Säter": "Dalarna", "Vansbro": "Dalarna", "Älvdalen": "Dalarna",
    # Gävleborg
    "Bollnäs": "Gävleborg", "Gävle": "Gävleborg", "Hofors": "Gävleborg",
    "Hudiksvall": "Gävleborg", "Ljusdal": "Gävleborg", "Nordanstig": "Gävleborg",
    "Ockelbo": "Gävleborg", "Ovanåker": "Gävleborg", "Sandviken": "Gävleborg",
    "Söderhamn": "Gävleborg",
    # Västernorrland
    "Härnösand": "Västernorrland", "Kramfors": "Västernorrland",
    "Sollefteå": "Västernorrland", "Sundsvall": "Västernorrland",
    "Timrå": "Västernorrland", "Ånge": "Västernorrland",
    "Örnsköldsvik": "Västernorrland",
    # Jämtland
    "Berg": "Jämtland", "Bräcke": "Jämtland", "Härjedalen": "Jämtland",
    "Krokom": "Jämtland", "Ragunda": "Jämtland", "Strömsund": "Jämtland",
    "Åre": "Jämtland", "Östersund": "Jämtland",
    # Västerbotten
    "Bjurholm": "Västerbotten", "Dorotea": "Västerbotten",
    "Lycksele": "Västerbotten", "Malå": "Västerbotten",
    "Nordmaling": "Västerbotten", "Norsjö": "Västerbotten",
    "Robertsfors": "Västerbotten", "Skellefteå": "Västerbotten",
    "Sorsele": "Västerbotten", "Storuman": "Västerbotten",
    "Umeå": "Västerbotten", "Vilhelmina": "Västerbotten",
    "Vindeln": "Västerbotten", "Vännäs": "Västerbotten",
    "Åsele": "Västerbotten",
    # Norrbotten
    "Arjeplog": "Norrbotten", "Arvidsjaur": "Norrbotten",
    "Boden": "Norrbotten", "Gällivare": "Norrbotten",
    "Haparanda": "Norrbotten", "Jokkmokk": "Norrbotten",
    "Kalix": "Norrbotten", "Kiruna": "Norrbotten",
    "Luleå": "Norrbotten", "Pajala": "Norrbotten",
    "Piteå": "Norrbotten", "Älvsbyn": "Norrbotten",
    "Överkalix": "Norrbotten", "Övertorneå": "Norrbotten",
}


def get_nuts1_region(kommun_name):
    """Return NUTS1 region ('south', 'east', 'north') from kommun name."""
    lan = _KOMMUN_LAN.get(kommun_name)
    if lan:
        return _LAN_NUTS1.get(lan, "south")
    # Fuzzy: try matching start of name (handles "Marks kommun" etc.)
    for k, v in _KOMMUN_LAN.items():
        if kommun_name.startswith(k) or k.startswith(kommun_name):
            return _LAN_NUTS1.get(v, "south")
    return "south"  # default fallback


def get_nuts1_label(region):
    """Human-readable label for NUTS1 region."""
    return {"south": "Syd (SE2)", "east": "Öst (SE1)", "north": "Norr (SE3)"}[region]


# ─── Vindbrukskollen API ─────────────────────────────────────────────────

def query_vindbrukskollen(layer_id, where_clause, out_fields="*", out_sr=4326):
    params = {
        "where": where_clause, "outFields": out_fields,
        "outSR": str(out_sr), "returnGeometry": "true", "f": "json",
    }
    url = f"{VINDBRUKSKOLLEN_BASE}/{layer_id}/query?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"User-Agent": "VindkraftAnalys/1.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(f"API error: {data['error']}")
    return data.get("features", [])


def fetch_project_area(project_name):
    out_fields = "OMRID,PROJNAMN,KOMNAMN,ORGNAMN,ORGNR,ANTALVERK,CALPROD,LANSNAMN,EL_NAMN,PBYGGSTART,PDRIFT"
    features = query_vindbrukskollen(
        PROJECT_AREA_LAYER, f"PROJNAMN='{project_name}'", out_fields,
    )
    if not features:
        features = query_vindbrukskollen(
            PROJECT_AREA_LAYER, f"PROJNAMN LIKE '%{project_name}%'", out_fields,
        )
    if not features:
        return None, None
    feat = features[0]
    attrs = feat["attributes"]
    rings = feat["geometry"]["rings"]
    polygon = [[pt[1], pt[0]] for pt in rings[0]]
    return attrs, polygon


def fetch_turbines(project_name, status_filter=None):
    all_turbines = []
    for layer_id, status_name in TURBINE_LAYERS.items():
        features = query_vindbrukskollen(
            layer_id, f"PROJNAMN='{project_name}'",
            "VERKID,PROJNAMN,TOTALHOJD,MAXEFFEKT,CALPROD,STATUS,KOMNAMN,ORGNAMN",
        )
        for feat in features:
            attrs = feat["attributes"]
            geom = feat["geometry"]
            all_turbines.append({
                "verkid": attrs.get("VERKID", ""),
                "lat": geom["y"], "lon": geom["x"],
                "total_height": attrs.get("TOTALHOJD"),
                "max_power_mw": attrs.get("MAXEFFEKT"),
                "annual_gwh": attrs.get("CALPROD"),
                "status": attrs.get("STATUS", status_name),
                "status_layer": status_name,
                "operator": attrs.get("ORGNAMN", ""),
            })
    all_turbines.sort(key=lambda t: t["verkid"])
    if status_filter:
        filters = [f.strip().lower() for f in status_filter.split(",")]
        all_turbines = [t for t in all_turbines
                        if any(f in t["status_layer"].lower() for f in filters)]
    return all_turbines


# ─── Overpass API ─────────────────────────────────────────────────────────

def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2 +
         math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
         math.sin(dlon / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def fetch_properties(center_lat, center_lon, search_radius_m=4500):
    query = f"""
    [out:json][timeout:30];
    (
      node["place"~"farm|hamlet|isolated_dwelling|village|locality"](around:{search_radius_m},{center_lat},{center_lon});
      node["name"]["building"~"farm|house|residential"](around:{search_radius_m},{center_lat},{center_lon});
    );
    out body;
    """
    data = urllib.parse.urlencode({"data": query}).encode("utf-8")
    result = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(OVERPASS_URL, data=data,
                                         headers={"User-Agent": "VindkraftAnalys/1.0"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                result = json.loads(resp.read().decode("utf-8"))
            break
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
            if attempt < 2:
                time.sleep((attempt + 1) * 10)
            else:
                raise
    places = []
    seen_names = set()
    for elem in result.get("elements", []):
        name = elem.get("tags", {}).get("name", "")
        if not name or name in seen_names:
            continue
        seen_names.add(name)
        places.append({"name": name, "lat": elem["lat"], "lon": elem["lon"]})
    return places


# ─── Models ──────────────────────────────────────────────────────────────

def calc_reduction_pct_2025(distance_km, region="south"):
    """W&W 2025: regional coefficients per 2-km band with monotonicity enforcement."""
    coeff = WW2025_COEFF.get(region, WW2025_COEFF["south"])
    if not coeff:
        return 0.0
    midpoints = sorted(coeff.keys())
    reductions = [100 * (math.exp(coeff[m]) - 1) for m in midpoints]
    # Enforce monotonicity: closer bands >= farther bands
    for j in range(len(reductions) - 1):
        if reductions[j + 1] < reductions[j]:
            reductions[j] = reductions[j + 1]
    if distance_km <= midpoints[0]:
        return reductions[0]
    last_mp = midpoints[-1]
    last_red = reductions[-1]
    taper_end = last_mp + 2.0
    if distance_km >= taper_end:
        return 0.0
    for j in range(len(midpoints) - 1):
        if midpoints[j] <= distance_km <= midpoints[j + 1]:
            t = (distance_km - midpoints[j]) / (midpoints[j + 1] - midpoints[j])
            return reductions[j] + t * (reductions[j + 1] - reductions[j])
    if distance_km > last_mp:
        t = (distance_km - last_mp) / (taper_end - last_mp)
        return last_red * (1 - t)
    return 0.0


def calc_reduction_pct_2021(distance_km):
    """W&W 2021: single national exponential curve."""
    beta = WW2021_A * math.exp(WW2021_B * distance_km)
    pct = 100 * (math.exp(beta) - 1)
    return pct if pct < -0.1 else 0.0  # cut off near-zero tail


def calc_reduction_pct(distance_km, model="ww2025", region="south"):
    """Dispatch to selected model."""
    if model == "ww2021":
        return calc_reduction_pct_2021(distance_km)
    return calc_reduction_pct_2025(distance_km, region)


def analyze_properties(turbines, places, max_radius_m, region="south", model="ww2025"):
    results = []
    for place in places:
        min_dist = float("inf")
        nearest = None
        for i, t in enumerate(turbines):
            d = haversine_m(place["lat"], place["lon"], t["lat"], t["lon"])
            if d < min_dist:
                min_dist = d
                nearest = i
        if min_dist <= max_radius_m:
            dist_km = min_dist / 1000.0
            red_pct = calc_reduction_pct(dist_km, model, region)
            results.append({
                "name": place["name"], "lat": place["lat"], "lon": place["lon"],
                "distance_m": round(min_dist), "nearest_turbine": nearest,
                "reduction_pct": red_pct, "region": region,
            })
    results.sort(key=lambda r: r["distance_m"])
    return results


# ─── PDF Generation ──────────────────────────────────────────────────────

def generate_fastigheter_pdf(project_info, turbines, properties, radius_m, output_path, region="south", model="ww2025"):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.colors import HexColor, black, grey
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors

    proj_name = project_info["PROJNAMN"]
    kommun = project_info["KOMNAMN"]
    radius_km = radius_m / 1000
    n_props = len(properties)
    today = date.today().isoformat()

    if properties:
        avg_red = sum(abs(p["reduction_pct"]) for p in properties) / n_props
        max_prop, min_prop = properties[0], properties[-1]
    else:
        avg_red = 0; max_prop = min_prop = None

    max_inter = 0
    for i, t1 in enumerate(turbines):
        for t2 in turbines[i+1:]:
            d = haversine_m(t1["lat"], t1["lon"], t2["lat"], t2["lon"])
            if d > max_inter: max_inter = d

    model_tag = f"W&W 2025 regional ({get_nuts1_label(region)})" if model == "ww2025" else "W&W 2021 national"
    model_tag_html = f"W&amp;W 2025 regional ({get_nuts1_label(region)})" if model == "ww2025" else "W&amp;W 2021 national"

    page_num = [0]
    def page_footer(canvas, doc):
        page_num[0] += 1
        canvas.saveState()
        canvas.setFont("Helvetica", 8); canvas.setFillColor(grey)
        canvas.drawCentredString(A4[0]/2, A4[1]-20, "Vindkraftens inverkan p\u00e5 fastighetsv\u00e4rden")
        canvas.drawCentredString(A4[0]/2, A4[1]-32, f"{proj_name}, {kommun} kommun \u2014 {model_tag}, {radius_km:.0f} km radie")
        canvas.drawCentredString(A4[0]/2, 25, f"Sida {page_num[0]}")
        canvas.restoreState()

    doc = SimpleDocTemplate(output_path, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2.5*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle('T', parent=styles['Title'], fontSize=18, leading=22, alignment=TA_CENTER, spaceAfter=4)
    subtitle_s = ParagraphStyle('ST', parent=styles['Normal'], fontSize=10, leading=14, alignment=TA_CENTER, spaceAfter=16)
    heading_s = ParagraphStyle('H', parent=styles['Heading2'], fontSize=13, leading=16, spaceBefore=14, spaceAfter=6)
    body_s = ParagraphStyle('B', parent=styles['Normal'], fontSize=9, leading=12)
    code_s = ParagraphStyle('C', parent=styles['Normal'], fontSize=9, leading=13, fontName='Courier', leftIndent=20)
    italic_s = ParagraphStyle('I', parent=styles['Normal'], fontSize=8.5, leading=11, fontName='Helvetica-Oblique', textColor=grey)

    story = []
    story.append(Paragraph("Vindkraftens inverkan p\u00e5 fastighetsv\u00e4rden", title_s))
    story.append(Paragraph(f"{proj_name}, {kommun} kommun \u2014 {model_tag_html}, {radius_km:.0f} km radie", subtitle_s))

    # Section 1: Turbines
    story.append(Paragraph("<b>1. Planerade vindkraftverk</b>", heading_s))
    for i, t in enumerate(turbines):
        line = f"VKV {i+1} ({t['verkid']}): {t['lat']:.4f}\u00b0N, {t['lon']:.4f}\u00b0E"
        if t.get("total_height"): line += f" \u2014 {t['total_height']} m"
        if t.get("status"): line += f" [{t['status']}]"
        story.append(Paragraph(line, body_s))
    story.append(Paragraph(f"Max avst\u00e5nd mellan verken: {round(max_inter)} m", body_s))
    if project_info.get("ORGNAMN"): story.append(Paragraph(f"Operat\u00f6r: {project_info['ORGNAMN']}", body_s))
    if project_info.get("CALPROD"): story.append(Paragraph(f"Ber\u00e4knad \u00e5rsproduktion: {project_info['CALPROD']} GWh", body_s))
    story.append(Spacer(1, 8))

    # Section 2: Method
    story.append(Paragraph("<b>2. Metod och modell</b>", heading_s))
    if model == "ww2025":
        story.append(Paragraph('K\u00e4lla: Westlund &amp; Wilhelmsson (2025), "Capitalisation of onshore wind turbines on property prices in Sweden: The need to compensate for negative externalities", <i>Economic Analysis and Policy</i>, 87, 1452\u20131468.', body_s))
        story.append(Spacer(1, 4))
        story.append(Paragraph('Studien anv\u00e4nder en hedonisk prismodell med staggered difference-in-difference baserad p\u00e5 \u00f6ver 600 000 fastighetsf\u00f6rs\u00e4ljningar i Sverige 2005\u20132018. Modellen \u00e4r regionalt differentierad (NUTS1: Syd, \u00d6st, Norr) med 2 km-intervaller.', body_s))
        story.append(Spacer(1, 4))
        story.append(Paragraph("Reduktion = 100 \u00d7 [exp(\u03b2) \u2013 1] %", code_s))
        story.append(Paragraph(f"Region f\u00f6r detta projekt: {get_nuts1_label(region)}", body_s))
    else:
        story.append(Paragraph('K\u00e4lla: Westlund &amp; Wilhelmsson (2021), hedonisk prismodell baserad p\u00e5 ca 69 000 sm\u00e5husf\u00f6rs\u00e4ljningar i Sverige 2013\u20132018.', body_s))
        story.append(Spacer(1, 4))
        story.append(Paragraph('Nationell exponentiell modell (ej regionalt differentierad):', body_s))
        story.append(Paragraph("\u03b2(d) = \u22120.2811 \u00d7 exp(\u22120.3811 \u00d7 d)", code_s))
        story.append(Paragraph("Reduktion = 100 \u00d7 [exp(\u03b2) \u2013 1] %", code_s))
    story.append(Spacer(1, 8))

    # Section 3: Coefficients
    if model == "ww2025":
        story.append(Paragraph("<b>3. Regionala koefficienter (Table 3, W&amp;W 2025)</b>", heading_s))
        coeff_data = [["Avst\u00e5nd", "Syd (\u03b2)", "Syd %", "\u00d6st (\u03b2)", "\u00d6st %", "Norr (\u03b2)", "Norr %"]]
        all_bands = [
            ("0\u20132 km", {"south": -0.183, "east": -0.091, "north": -0.128}),
            ("2\u20134 km", {"south": -0.111, "east": -0.102, "north": None}),
            ("4\u20136 km", {"south": None, "east": -0.046, "north": None}),
        ]
        for label, betas in all_bands:
            row = [label]
            for reg in ["south", "east", "north"]:
                b = betas.get(reg)
                if b is not None:
                    row.append(f"{b:.3f}")
                    row.append(f"{100*(math.exp(b)-1):.1f}%")
                else:
                    row.append("\u2013")
                    row.append("ej sign.")
            coeff_data.append(row)
        ctable = Table(coeff_data, colWidths=[55, 50, 50, 50, 50, 50, 50])
        ctable.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 7.5), ('FONT', (0,1), (-1,-1), 'Helvetica', 7.5),
            ('ALIGN', (1,0), (-1,-1), 'CENTER'), ('ALIGN', (0,0), (0,-1), 'LEFT'),
            ('LINEBELOW', (0,0), (-1,0), 0.5, black), ('LINEBELOW', (0,-1), (-1,-1), 0.5, black),
            ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        story.append(ctable)
        story.append(Spacer(1, 4))
        story.append(Paragraph('Enbart statistiskt signifikanta koefficienter (p&lt;0,05) anv\u00e4nds. Linj\u00e4r interpolering mellan bandmittpunkter.', italic_s))
    else:
        story.append(Paragraph("<b>3. Modellparametrar (W&amp;W 2021)</b>", heading_s))
        story.append(Paragraph("A = \u22120.2811, B = \u22120.3811", code_s))
        story.append(Paragraph("Exponentiell avtagande kurva, ~24% vid 0 km, avtar till ~0% vid ~8 km.", body_s))
    story.append(Spacer(1, 10))

    # Section 4: Reduction curve
    curve_label = get_nuts1_label(region) if model == "ww2025" else "Nationell"
    story.append(Paragraph(f"<b>4. Reduktionskurva \u2014 {curve_label} (0\u2013{radius_km:.0f} km)</b>", heading_s))
    rdata = [["Avst\u00e5nd", "Reduktion"]]
    for d_m in range(0, radius_m + 1, 250):
        d_km = d_m / 1000.0
        r = calc_reduction_pct(d_km, model, region)
        rdata.append([f"{d_m} m", f"{r:.1f}%"])
    rtable = Table(rdata, colWidths=[80, 80])
    rtable.setStyle(TableStyle([
        ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 8), ('FONT', (0,1), (-1,-1), 'Helvetica', 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('LINEBELOW', (0,0), (-1,0), 0.5, black), ('LINEBELOW', (0,-1), (-1,-1), 0.5, black),
        ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(rtable)

    # Section 5: Property table
    story.append(PageBreak())
    story.append(Paragraph(f"<b>5. Fastigheter inom {radius_km:.0f} km \u2013 individuell ber\u00e4kning</b>", heading_s))
    story.append(Spacer(1, 6))
    col_widths = [22, 100, 48, 48, 52, 52, 52]
    header_row = ["#", "Fastighet", "Avst.(m)", "N\u00e4rmast", "Lat", "Lon", "Redukt."]
    page_size = 25
    for batch_start in range(0, n_props, page_size):
        if batch_start > 0: story.append(PageBreak())
        batch = properties[batch_start:batch_start + page_size]
        pdata = [header_row]
        for j, p in enumerate(batch):
            pdata.append([str(batch_start+j+1), p["name"], str(p["distance_m"]),
                f"VKV {p['nearest_turbine']+1}", f"{p['lat']:.4f}", f"{p['lon']:.4f}",
                f"{p['reduction_pct']:.1f}%"])
        ptable = Table(pdata, colWidths=col_widths)
        ptable.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 7), ('FONT', (0,1), (-1,-1), 'Helvetica', 7),
            ('ALIGN', (0,0), (0,-1), 'CENTER'), ('ALIGN', (2,0), (-1,-1), 'CENTER'),
            ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ('LINEBELOW', (0,0), (-1,0), 0.5, black), ('LINEBELOW', (0,-1), (-1,-1), 0.5, black),
            ('TOPPADDING', (0,0), (-1,-1), 1.5), ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, HexColor('#F5F5F5')]),
        ]))
        story.append(ptable)

    story.append(Spacer(1, 14))
    story.append(Paragraph("<b>6. Sammanfattning</b>", heading_s))
    story.append(Paragraph(f"Totalt {n_props} identifierade fastigheter inom {radius_km:.0f} km.", body_s))
    story.append(Paragraph("Ber\u00e4knad v\u00e4rdereduktion:", body_s))
    if max_prop:
        story.append(Paragraph(f"\u00a0\u00a0\u00a0- H\u00f6gst: {max_prop['reduction_pct']:.1f}% ({max_prop['name']} vid {max_prop['distance_m']} m)", body_s))
    if min_prop:
        story.append(Paragraph(f"\u00a0\u00a0\u00a0- L\u00e4gst: {min_prop['reduction_pct']:.1f}% ({min_prop['name']} vid {min_prop['distance_m']} m)", body_s))
    story.append(Paragraph(f"\u00a0\u00a0\u00a0- Medelv\u00e4rde: -{avg_red:.1f}%", body_s))
    story.append(Spacer(1, 16))
    story.append(Paragraph(f"Genererad: {today}", italic_s))
    if model == "ww2025":
        story.append(Paragraph("K\u00e4lla: Westlund &amp; Wilhelmsson (2025), Economic Analysis and Policy, 87, 1452\u20131468.", italic_s))
    else:
        story.append(Paragraph("K\u00e4lla: Westlund &amp; Wilhelmsson (2021), hedonisk prismodell, 69 000 transaktioner.", italic_s))
    story.append(Paragraph("Kartdata: Lantm\u00e4teriet CC BY 4.0, OpenStreetMap (ODbL)", italic_s))

    doc.build(story, onFirstPage=page_footer, onLaterPages=page_footer)


def generate_ekonomi_pdf(project_info, turbines, properties, radius_m, assumed_value_tkr, output_path, model="ww2025", region="south"):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.colors import HexColor, black, grey, white
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors

    proj_name = project_info["PROJNAMN"]
    kommun = project_info["KOMNAMN"]
    n_props = len(properties)
    radius_km = radius_m / 1000
    today = date.today().isoformat()

    losses = []
    total_loss = 0
    for p in properties:
        red_pct = abs(p["reduction_pct"])
        loss_tkr = red_pct / 100 * assumed_value_tkr
        losses.append({"name": p["name"], "distance_m": p["distance_m"],
            "nearest": f"VKV {p['nearest_turbine']+1}",
            "reduction_str": f"-{red_pct:.1f}%", "loss_tkr": round(loss_tkr)})
        total_loss += loss_tkr

    total_loss_rounded = round(total_loss)
    total_msek = total_loss_rounded / 1000
    avg_loss = total_loss / n_props if n_props else 0
    max_loss = losses[0] if losses else None
    min_loss = losses[-1] if losses else None

    page_num = [0]
    def page_footer(canvas, doc):
        page_num[0] += 1
        canvas.saveState()
        canvas.setFont("Helvetica", 8); canvas.setFillColor(grey)
        canvas.drawRightString(A4[0]-2*cm, A4[1]-20, f"Ekonomisk konsekvensanalys - Vindkraft vid {proj_name}")
        canvas.drawCentredString(A4[0]/2, 25, f"Sida {page_num[0]}")
        canvas.restoreState()

    doc = SimpleDocTemplate(output_path, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle('T', parent=styles['Title'], fontSize=18, leading=22, alignment=TA_CENTER, spaceAfter=4)
    subtitle_s = ParagraphStyle('ST', parent=styles['Normal'], fontSize=11, leading=14, alignment=TA_CENTER, spaceAfter=16)
    heading_s = ParagraphStyle('H', parent=styles['Heading2'], fontSize=14, leading=17, spaceBefore=14, spaceAfter=8)
    body_s = ParagraphStyle('B', parent=styles['Normal'], fontSize=9.5, leading=13)
    italic_s = ParagraphStyle('I', parent=styles['Normal'], fontSize=8.5, leading=11, fontName='Helvetica-Oblique', textColor=grey)
    light_yellow = HexColor('#FFFDE7')

    story = []
    story.append(Paragraph("<b>Ekonomisk konsekvensanalys</b>", title_s))
    story.append(Paragraph(f"Planerade vindkraftverk vid {proj_name}, {kommun} kommun", subtitle_s))
    story.append(Spacer(1, 8))

    summary_data = [
        [Paragraph('<font color="#B71C1C"><b>SAMMANFATTNING</b></font>', ParagraphStyle('c', alignment=TA_CENTER, fontSize=10))],
        [Paragraph(f'<font color="#B71C1C"><b>Total ber\u00e4knad v\u00e4rdef\u00f6rlust: {total_msek:.1f} MSEK</b></font>',
            ParagraphStyle('c', alignment=TA_CENTER, fontSize=16, leading=20))],
        [Paragraph(f'{n_props} fastigheter inom {radius_km:.0f} km | Antaget medelv\u00e4rde: {assumed_value_tkr} tkr/fastighet',
            ParagraphStyle('c', alignment=TA_CENTER, fontSize=9))],
        [Paragraph(f'Genomsnittlig f\u00f6rlust per fastighet: {round(avg_loss)} tkr',
            ParagraphStyle('c', alignment=TA_CENTER, fontSize=9))],
    ]
    if max_loss and min_loss:
        summary_data.append([Paragraph(
            f'Spridning: {min_loss["loss_tkr"]} tkr ({min_loss["name"]}) till {max_loss["loss_tkr"]} tkr ({max_loss["name"]})',
            ParagraphStyle('c', alignment=TA_CENTER, fontSize=9))])

    summary_table = Table(summary_data, colWidths=[doc.width * 0.8])
    summary_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, black), ('BACKGROUND', (0,0), (-1,-1), light_yellow),
        ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 20))

    story.append(Paragraph('<b>Vad inneb\u00e4r detta f\u00f6r fastighets\u00e4garna?</b>',
        ParagraphStyle('h', parent=heading_s, fontSize=13)))
    if max_loss and min_loss:
        story.append(Paragraph(
            f'Enligt forskning av Westlund &amp; Wilhelmsson (2025) minskar fastighetsv\u00e4rden i n\u00e4rheten av '
            f'vindkraftverk. F\u00f6r de {n_props} identifierade fastigheterna inom {radius_km:.0f} km fr\u00e5n de planerade '
            f'verken vid {proj_name} inneb\u00e4r detta en sammanlagd ber\u00e4knad v\u00e4rdef\u00f6rlust p\u00e5 ca {total_msek:.1f} MSEK '
            f'(vid ett antaget fastighetsv\u00e4rde p\u00e5 {assumed_value_tkr/1000:.1f} MSEK).', body_s))
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            f'Den n\u00e4rmaste fastigheten ({max_loss["name"]}, {max_loss["distance_m"]} m) drabbas av en ber\u00e4knad '
            f'v\u00e4rdeminskning p\u00e5 {max_loss["loss_tkr"]} tkr ({max_loss["reduction_str"]}). \u00c4ven fastigheter p\u00e5 '
            f'n\u00e4rmare {radius_km:.0f} km avst\u00e5nd ({min_loss["name"]}, {min_loss["distance_m"]} m) p\u00e5verkas med '
            f'{min_loss["loss_tkr"]} tkr ({min_loss["reduction_str"]}).', body_s))

    # Page 2: Table
    story.append(PageBreak())
    story.append(Paragraph("<b>Detaljerad fastighetslista \u2013 Ber\u00e4knad v\u00e4rdef\u00f6rlust</b>", heading_s))
    col_widths = [22, 88, 48, 42, 48, 55]
    header = ["#", "Fastighet", "Avst (m)", "N\u00e4rm.", "Red (%)", "F\u00f6rlust tkr"]
    page_size = 25
    for batch_start in range(0, n_props, page_size):
        if batch_start > 0: story.append(PageBreak())
        batch = losses[batch_start:batch_start + page_size]
        is_last = batch_start + page_size >= n_props
        tdata = [header]
        for j, row in enumerate(batch):
            tdata.append([str(batch_start+j+1), row["name"], str(row["distance_m"]),
                row["nearest"], row["reduction_str"], str(row["loss_tkr"])])
        if is_last:
            tdata.append(["",
                Paragraph("<b>SUMMA</b>", ParagraphStyle('b', fontName='Helvetica-Bold', fontSize=7.5)),
                "", "", "",
                Paragraph(f"<b>{total_loss_rounded}</b>", ParagraphStyle('b2', fontName='Helvetica-Bold', fontSize=7.5, alignment=TA_RIGHT))])
        t = Table(tdata, colWidths=col_widths)
        style_cmds = [
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 7.5), ('FONT', (0,1), (-1,-1), 'Helvetica', 7.5),
            ('ALIGN', (0,0), (0,-1), 'CENTER'), ('ALIGN', (2,0), (2,-1), 'CENTER'),
            ('ALIGN', (3,0), (3,-1), 'CENTER'), ('ALIGN', (4,0), (4,-1), 'CENTER'),
            ('ALIGN', (5,0), (5,-1), 'RIGHT'), ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ('LINEBELOW', (0,0), (-1,0), 0.5, black), ('LINEBELOW', (0,-1), (-1,-1), 0.5, black),
            ('TOPPADDING', (0,0), (-1,-1), 1.5), ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ]
        if is_last:
            style_cmds.extend([
                ('ROWBACKGROUNDS', (0,1), (-1,-2), [white, HexColor('#F5F5F5')]),
                ('BACKGROUND', (0,-1), (-1,-1), HexColor('#E0E0E0')),
                ('LINEABOVE', (0,-1), (-1,-1), 0.5, black),
            ])
        else:
            style_cmds.append(('ROWBACKGROUNDS', (0,1), (-1,-1), [white, HexColor('#F5F5F5')]))
        t.setStyle(TableStyle(style_cmds))
        story.append(t)

    story.append(Spacer(1, 10))
    story.append(Paragraph(
        f'Antaget fastighetsv\u00e4rde: {assumed_value_tkr/1000:.1f} MSEK. '
        f'Modell: {"W&amp;W 2025 regional (Table 3, NUTS1)" if model == "ww2025" else "W&amp;W 2021 national"}. '
        f'K\u00e4lla: Westlund &amp; Wilhelmsson ({"2025" if model == "ww2025" else "2021"}).',
        ParagraphStyle('S', parent=styles['Normal'], fontSize=8, leading=10, textColor=grey)))
    story.append(Spacer(1, 16))
    story.append(Paragraph("<b>Slutsats</b>", heading_s))
    if properties:
        story.append(Paragraph(
            f'De planerade vindkraftverken vid {proj_name} ber\u00e4knas p\u00e5verka {n_props} fastigheter '
            f'inom {radius_km:.0f} km med en total v\u00e4rdeminskning p\u00e5 {total_msek:.1f} MSEK.', body_s))
    story.append(Spacer(1, 20))
    story.append(Paragraph(f"Genererad: {today}", italic_s))
    if model == "ww2025":
        story.append(Paragraph("K\u00e4lla: Westlund &amp; Wilhelmsson (2025), Economic Analysis and Policy, 87, 1452\u20131468.", italic_s))
    else:
        story.append(Paragraph("K\u00e4lla: Westlund &amp; Wilhelmsson (2021), hedonisk prismodell, 69 000 transaktioner.", italic_s))
    story.append(Paragraph("Kartdata: Lantm\u00e4teriet CC BY 4.0, OpenStreetMap (ODbL)", italic_s))

    doc.build(story, onFirstPage=page_footer, onLaterPages=page_footer)


# ─── HTML Map for file save ──────────────────────────────────────────────

def generate_html_map(project_info, turbines, polygon, properties, radius_m, assumed_value_tkr):
    """Return standalone HTML string for an interactive map."""
    center_lat = sum(t["lat"] for t in turbines) / len(turbines)
    center_lon = sum(t["lon"] for t in turbines) / len(turbines)
    proj_name = project_info["PROJNAMN"]
    kommun = project_info["KOMNAMN"]
    n_props = len(properties)
    total_loss = sum(abs(p["reduction_pct"]) / 100 * assumed_value_tkr for p in properties)
    total_msek = total_loss / 1000

    turbines_js = json.dumps([{"name": f"VKV {i+1}", "lat": t["lat"], "lon": t["lon"],
        "verkid": t["verkid"], "status": t["status"], "height": t.get("total_height"),
    } for i, t in enumerate(turbines)])
    props_js = json.dumps([{"name": p["name"], "lat": p["lat"], "lon": p["lon"],
        "distance": p["distance_m"], "nearest": f"VKV {p['nearest_turbine']+1}",
        "reduction": round(p["reduction_pct"], 1),
    } for p in properties])
    polygon_js = json.dumps(polygon) if polygon else "[]"

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Vindkraft - {proj_name}</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>body{{margin:0;font-family:system-ui}}#map{{height:100vh;width:100vw}}
.info{{position:absolute;top:10px;left:60px;z-index:1000;background:rgba(255,255,255,.95);padding:12px 16px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.2);max-width:360px;font-size:13px}}
.info h3{{margin:0 0 4px;font-size:15px}}.total{{background:#fafafa;border:1px solid #ccc;border-radius:6px;padding:8px;margin:8px 0;text-align:center}}
.total .amt{{font-size:20px;font-weight:bold;color:#d32f2f}}</style></head><body>
<div id="map"></div>
<div class="info">
<h3>Vindkraft Fastighetsanalys</h3>
<div style="color:#666">{proj_name}, {kommun}</div>
<div class="total"><div>Total v\u00e4rdef\u00f6rlust</div><div class="amt">{total_msek:.1f} MSEK</div>
<div style="font-size:11px">{len(turbines)} verk | {n_props} fastigheter</div></div>
<div style="font-size:10px;color:#999">Modell: Westlund &amp; Wilhelmsson (2025)<br>Kartdata: Lantm\u00e4teriet, OSM</div>
</div>
<script>
const T={turbines_js}, P={props_js}, A={polygon_js};
const map=L.map('map').setView([{center_lat},{center_lon}],13);
L.tileLayer.wms('https://minkarta.lantmateriet.se/map/topowebb/?',{{layers:'topowebbkartan',format:'image/png',transparent:true,attribution:'Lantm\u00e4teriet CC BY 4.0'}}).addTo(map);
if(A.length>0)L.polygon(A,{{color:'#1565C0',weight:2,dashArray:'8,4',fillColor:'#1565C0',fillOpacity:0.05}}).addTo(map);
[1000,2000,3000].forEach(r=>L.circle([{center_lat},{center_lon}],{{radius:r,color:'#9e9e9e',weight:1,dashArray:'4,4',fillOpacity:0,interactive:false}}).addTo(map));
T.forEach(t=>L.circleMarker([t.lat,t.lon],{{radius:10,color:'#b71c1c',fillColor:'#d32f2f',fillOpacity:0.9,weight:2}}).bindPopup('<b>'+t.name+'</b><br>'+t.verkid+'<br>'+t.status+(t.height?' '+t.height+'m':'')).addTo(map));
function gc(r){{const a=Math.abs(r);return a>=18?'#d32f2f':a>=15?'#e65100':a>=12?'#fbc02d':'#4caf50'}}
P.forEach(p=>L.circleMarker([p.lat,p.lon],{{radius:7,color:'#0d47a1',fillColor:gc(p.reduction),fillOpacity:0.85,weight:1.5}}).bindPopup('<b>'+p.name+'</b><br>Avst: '+p.distance+'m ('+p.nearest+')<br>Red: '+p.reduction+'%').addTo(map));
</script></body></html>"""


# ─── Page Config ─────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Vindkraft Fastighetsanalys",
    page_icon="\u2699\ufe0f",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""<style>
.main-header{font-size:2rem;font-weight:700;color:#1B5E20;margin-bottom:0}
.sub-header{font-size:1rem;color:#666;margin-top:0;margin-bottom:1.5rem}
div[data-testid="stMetric"]{background:#f0f2f6;padding:.75rem;border-radius:8px}
</style>""", unsafe_allow_html=True)


# ─── Sidebar ─────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Inst\u00e4llningar")

    project_name = st.text_input("Projektnamn", placeholder="t.ex. Str\u00e4ngsered, Marb\u00e4ck",
        help="Namn enligt Vindbrukskollen")

    radius_m = st.slider("Analysradie (m)", 1000, 10000, DEFAULT_RADIUS_M, step=500)

    assumed_value = st.number_input("Antaget fastighetsv\u00e4rde (tkr)",
        min_value=500, max_value=20000, value=DEFAULT_ASSUMED_VALUE_TKR, step=500)

    model_label = st.selectbox("Modell", options=list(MODEL_OPTIONS.keys()),
        help="W&W 2025: regional (Syd/\u00d6st/Norr), diff-in-diff, 600k transaktioner.\n"
             "W&W 2021: nationell exponentiell kurva, 69k transaktioner.")
    model_key = MODEL_OPTIONS[model_label]

    status_options = list(TURBINE_LAYERS.values())
    status_filter = st.multiselect("Statusfilter (verk)", options=status_options, default=[],
        help="L\u00e4mna tomt f\u00f6r alla statusar")

    run_btn = st.button("K\u00f6r analys", type="primary", use_container_width=True,
                         disabled=not project_name.strip())

    st.markdown("---")
    st.markdown(f"**Modell:** {model_label}\n\n"
                "**Data:** Vindbrukskollen, OpenStreetMap\n\n"
                "**Kartor:** Lantm\u00e4teriet CC BY 4.0")


# ─── Color helper ────────────────────────────────────────────────────────

def reduction_color(pct):
    a = abs(pct)
    if a >= 18: return "#B71C1C"
    elif a >= 15: return "#E65100"
    elif a >= 12: return "#F57F17"
    elif a >= 10: return "#FBC02D"
    else: return "#7CB342"


def build_folium_map(turbines, polygon, properties, radius_m, center_lat, center_lon):
    m = folium.Map(location=[center_lat, center_lon], zoom_start=13, tiles='openstreetmap')
    folium.raster_layers.WmsTileLayer(
        url="https://minkarta.lantmateriet.se/map/topowebb/?",
        layers="topowebbkartan",
        fmt="image/png",
        transparent=True,
        name="Lantm\u00e4teriet Topo",
        attr="Lantm\u00e4teriet CC BY 4.0",
        overlay=True,
        control=True,
    ).add_to(m)

    if polygon:
        folium.Polygon(polygon, color='#1565C0', weight=2, fill=True,
                       fill_color='#1565C0', fill_opacity=0.1).add_to(m)

    for i, t in enumerate(turbines):
        folium.Marker([t["lat"], t["lon"]],
            popup=f"<b>VKV {i+1}</b><br>{t['verkid']}<br>{t['status']}"
                  + (f"<br>{t['total_height']}m" if t.get('total_height') else ""),
            tooltip=f"VKV {i+1}",
            icon=folium.Icon(color='red', icon='flash', prefix='fa'),
        ).add_to(m)
        folium.Circle([t["lat"], t["lon"]], radius=radius_m,
            color='#B71C1C', weight=1, fill=False, dash_array='5,5').add_to(m)

    for p in properties:
        c = reduction_color(p["reduction_pct"])
        folium.CircleMarker([p["lat"], p["lon"]], radius=7, color=c, fill=True,
            fill_color=c, fill_opacity=0.8,
            popup=f"<b>{p['name']}</b><br>{p['distance_m']}m<br>{p['reduction_pct']:.1f}%",
            tooltip=f"{p['name']} ({p['reduction_pct']:.1f}%)",
        ).add_to(m)

    folium.LayerControl().add_to(m)
    return m


# ─── Main ────────────────────────────────────────────────────────────────

st.markdown('<p class="main-header">Vindkraft Fastighetsanalys</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Ber\u00e4kning av fastighetsv\u00e4rdens p\u00e5verkan fr\u00e5n vindkraftverk</p>', unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

if run_btn and project_name.strip():
    sf = ",".join(s.lower() for s in status_filter) if status_filter else None

    with st.status("K\u00f6r analys...", expanded=True) as status:
        st.write("H\u00e4mtar projektdata fr\u00e5n Vindbrukskollen...")
        try:
            project_info, polygon = fetch_project_area(project_name.strip())
        except Exception as e:
            st.error(f"Kunde inte kontakta Vindbrukskollen: {e}")
            st.stop()

        if not project_info:
            # No project area in layer 2 — search turbine layers directly
            st.write("Inget projektomr\u00e5de hittat, s\u00f6ker bland verk...")
            turbines = fetch_turbines(project_name.strip(), sf)
            if not turbines and sf:
                turbines = fetch_turbines(project_name.strip(), None)
            if turbines:
                # Build project_info from first turbine layer hit
                canonical = project_name.strip()
                project_info = None
                # Turbine layers only have these fields (not ORGNR, LANSNAMN etc.)
                t0 = turbines[0]
                for lid in TURBINE_LAYERS:
                    try:
                        feats = query_vindbrukskollen(lid,
                            f"PROJNAMN LIKE '%{project_name.strip()}%'",
                            "PROJNAMN,KOMNAMN,ORGNAMN,CALPROD")
                    except Exception:
                        continue
                    if feats:
                        a = feats[0]["attributes"]
                        canonical = a.get("PROJNAMN", project_name.strip())
                        project_info = {
                            "PROJNAMN": canonical,
                            "OMRID": t0.get("verkid", "?").rsplit("-", 1)[0],
                            "KOMNAMN": a.get("KOMNAMN", ""),
                            "ORGNAMN": a.get("ORGNAMN"),
                            "ORGNR": None,
                            "ANTALVERK": len(turbines),
                            "CALPROD": a.get("CALPROD"),
                            "LANSNAMN": None,
                            "EL_NAMN": None,
                            "PBYGGSTART": None, "PDRIFT": None,
                        }
                        break
                if not project_info:
                    project_info = {
                        "PROJNAMN": project_name.strip(), "OMRID": "?",
                        "KOMNAMN": "", "ORGNAMN": None, "ORGNR": None,
                        "ANTALVERK": len(turbines), "CALPROD": None,
                        "LANSNAMN": None, "EL_NAMN": None,
                        "PBYGGSTART": None, "PDRIFT": None,
                    }
                polygon = None
                st.write(f"**{canonical}** \u2014 {len(turbines)} verk (inget projektomr\u00e5de)")
            else:
                suggestions = query_vindbrukskollen(PROJECT_AREA_LAYER,
                    f"PROJNAMN LIKE '%{project_name.strip()}%'", "PROJNAMN,OMRID,KOMNAMN")
                msg = f"**{project_name}** hittades inte."
                if suggestions:
                    msg += "\n\nLiknande projekt:\n"
                    for s in suggestions[:10]:
                        a = s["attributes"]
                        msg += f"- **{a['PROJNAMN']}** ({a['OMRID']}, {a['KOMNAMN']})\n"
                st.error(msg)
                st.stop()
        else:
            canonical = project_info["PROJNAMN"]
            st.write(f"**{canonical}** ({project_info['OMRID']}), {project_info['KOMNAMN']}")

            st.write("H\u00e4mtar vindkraftverk...")
            turbines = fetch_turbines(canonical, sf)
            if not turbines:
                st.error("Inga verk hittades.")
                st.stop()
            st.write(f"**{len(turbines)}** verk")

        st.write("S\u00f6ker fastigheter via OpenStreetMap...")
        center_lat = sum(t["lat"] for t in turbines) / len(turbines)
        center_lon = sum(t["lon"] for t in turbines) / len(turbines)
        try:
            places = fetch_properties(center_lat, center_lon, search_radius_m=radius_m + 1500)
        except Exception as e:
            st.error(f"Overpass API: {e}")
            st.stop()
        st.write(f"**{len(places)}** platser")

        st.write("Ber\u00e4knar v\u00e4rdereduktion...")
        kommun = project_info.get("KOMNAMN", "")
        region = get_nuts1_region(kommun)
        if model_key == "ww2025":
            st.write(f"Modell: **W&W 2025 Regional** \u2014 {get_nuts1_label(region)} ({kommun})")
        else:
            st.write(f"Modell: **W&W 2021 National** ({kommun})")
        properties = analyze_properties(turbines, places, radius_m, region, model_key)
        st.write(f"**{len(properties)}** fastigheter inom {radius_m/1000:.0f} km")

        st.write("Genererar rapporter...")
        safe = canonical.lower().replace(" ", "_")
        for ch, rep in [("\u00e4","a"),("\u00f6","o"),("\u00e5","a"),("\u00e9","e")]:
            safe = safe.replace(ch, rep)

        out_dir = tempfile.mkdtemp()
        os.makedirs(out_dir, exist_ok=True)

        pdf1 = os.path.join(out_dir, f"vindkraft_fastighetsvarden_{safe}.pdf")
        pdf2 = os.path.join(out_dir, f"vindkraft_ekonomisk_analys_{safe}.pdf")
        html_path = os.path.join(out_dir, f"vindkraft_karta_{safe}.html")

        generate_fastigheter_pdf(project_info, turbines, properties, radius_m, pdf1, region, model_key)
        generate_ekonomi_pdf(project_info, turbines, properties, radius_m, assumed_value, pdf2, model_key, region)

        html_content = generate_html_map(project_info, turbines, polygon, properties, radius_m, assumed_value)
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        status.update(label="Analys klar!", state="complete")

    st.session_state.results = {
        "project_info": project_info, "polygon": polygon, "turbines": turbines,
        "properties": properties, "radius_m": radius_m, "assumed_value": assumed_value,
        "center_lat": center_lat, "center_lon": center_lon,
        "pdf1": pdf1, "pdf2": pdf2, "html_path": html_path,
        "out_dir": out_dir, "region": region, "model": model_key,
    }


# ─── Results ─────────────────────────────────────────────────────────────

r = st.session_state.results
if r:
    pi = r["project_info"]
    turbines = r["turbines"]
    props = r["properties"]

    st.markdown("---")
    st.markdown(f"### {pi['PROJNAMN']}, {pi['KOMNAMN']}")

    # ── Project info card ──
    org_name = pi.get("ORGNAMN") or "Uppgift saknas"
    org_nr = pi.get("ORGNR") or ""
    lan = pi.get("LANSNAMN") or ""
    el_omr = pi.get("EL_NAMN") or ""
    antal_verk = pi.get("ANTALVERK")
    cal_prod = pi.get("CALPROD")

    # Format dates (epoch ms → readable)
    def _fmt_date(epoch_ms):
        if not epoch_ms:
            return None
        try:
            from datetime import datetime
            return datetime.fromtimestamp(epoch_ms / 1000).strftime("%Y-%m-%d")
        except Exception:
            return None

    byggstart = _fmt_date(pi.get("PBYGGSTART"))
    driftstart = _fmt_date(pi.get("PDRIFT"))

    info_left, info_right = st.columns(2)
    with info_left:
        st.markdown("**Exploat\u00f6r / S\u00f6kande**")
        org_line = f"**{org_name}**"
        if org_nr:
            org_line += f" &nbsp;(org.nr {org_nr})"
        st.markdown(org_line, unsafe_allow_html=True)
        if lan or el_omr:
            parts = []
            if lan:
                parts.append(f"{lan}")
            if el_omr:
                parts.append(f"Elomr\u00e5de {el_omr}")
            st.caption(" \u00b7 ".join(parts))
    with info_right:
        st.markdown("**Projektuppgifter**")
        detail_lines = []
        if antal_verk:
            detail_lines.append(f"Antal verk (ans\u00f6kan): **{antal_verk}**")
        if cal_prod:
            detail_lines.append(f"Ber\u00e4knad \u00e5rsproduktion: **{cal_prod:.1f} GWh**")
        if byggstart:
            detail_lines.append(f"Planerad byggstart: {byggstart}")
        if driftstart:
            detail_lines.append(f"Planerad drift: {driftstart}")
        detail_lines.append(f"Omr\u00e5des-ID: {pi.get('OMRID', '?')}")
        st.markdown("  \n".join(detail_lines))

    st.markdown("")  # spacing

    if props:
        total_loss = sum(abs(p["reduction_pct"]) / 100 * r["assumed_value"] for p in props)
        avg_red = sum(abs(p["reduction_pct"]) for p in props) / len(props)
        max_red = max(abs(p["reduction_pct"]) for p in props)
    else:
        total_loss = avg_red = max_red = 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Vindkraftverk", len(turbines))
    c2.metric("Fastigheter", len(props))
    c3.metric("Total f\u00f6rlust", f"{total_loss/1000:.1f} MSEK")
    c4.metric("Medelreduktion", f"-{avg_red:.1f}%")
    c5.metric("Maxreduktion", f"-{max_red:.1f}%")

    # Map
    st.markdown("### Karta")
    fmap = build_folium_map(turbines, r["polygon"], props, r["radius_m"],
                            r["center_lat"], r["center_lon"])
    map_html = fmap._repr_html_()
    components.html(map_html, height=600, scrolling=False)

    # Table
    st.markdown("### Fastigheter")
    if props:
        import pandas as pd
        df = pd.DataFrame([{
            "#": i+1, "Fastighet": p["name"], "Avst\u00e5nd (m)": p["distance_m"],
            "N\u00e4rmast": f"VKV {p['nearest_turbine']+1}",
            "Reduktion": f"{p['reduction_pct']:.1f}%",
            "F\u00f6rlust (tkr)": round(abs(p["reduction_pct"]) / 100 * r["assumed_value"]),
        } for i, p in enumerate(props)])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.markdown(f"**Totalt:** {len(props)} fastigheter | **{total_loss/1000:.1f} MSEK** "
                    f"(vid {r['assumed_value']/1000:.1f} MSEK/fastighet)")

    # Downloads
    st.markdown("### Ladda ner")
    d1, d2, d3 = st.columns(3)
    with open(r["pdf1"], "rb") as f:
        d1.download_button("Fastighetsv\u00e4rden (PDF)", f.read(),
            file_name=os.path.basename(r["pdf1"]), mime="application/pdf", use_container_width=True)
    with open(r["pdf2"], "rb") as f:
        d2.download_button("Ekonomisk analys (PDF)", f.read(),
            file_name=os.path.basename(r["pdf2"]), mime="application/pdf", use_container_width=True)
    with open(r["html_path"], "r", encoding="utf-8") as f:
        d3.download_button("Interaktiv karta (HTML)", f.read(),
            file_name=os.path.basename(r["html_path"]), mime="text/html", use_container_width=True)

    # Details
    with st.expander("Verksdetaljer"):
        for i, t in enumerate(turbines):
            st.markdown(f"**VKV {i+1}** ({t['verkid']}): {t['lat']:.6f}\u00b0N, {t['lon']:.6f}\u00b0E \u2014 "
                f"{t['status']}" + (f", {t['total_height']}m" if t.get('total_height') else ""))

    with st.expander("Om modellen"):
        cur_model = r.get("model", "ww2025")
        reg_label = get_nuts1_label(r.get("region", "south"))
        if cur_model == "ww2025":
            st.markdown(f"""**Westlund & Wilhelmsson (2025)** \u2014 *Economic Analysis and Policy*, 87, 1452\u20131468.

Hedonisk prismodell med staggered difference-in-difference, baserad p\u00e5 \u00f6ver 600 000
fastighetsf\u00f6rs\u00e4ljningar i Sverige 2005\u20132018. Regionalt differentierad (NUTS1).

Reduktion = 100 \u00d7 [exp(\u03b2) \u2013 1] %, d\u00e4r \u03b2 fr\u00e5n Table 3 per 2 km-band och region.

| Avst\u00e5nd | Syd | \u00d6st | Norr |
|---|---|---|---|
| 0\u20132 km | \u221216.7% | \u22128.7% | \u221212.0% |
| 2\u20134 km | \u221210.5% | \u22129.7% | ej sign. |
| 4\u20136 km | ej sign. | \u22124.5% | ej sign. |

**Region f\u00f6r detta projekt:** {reg_label}""")
        else:
            st.markdown("""**Westlund & Wilhelmsson (2021)** \u2014 hedonisk prismodell.

Baserad p\u00e5 ca 69 000 sm\u00e5husf\u00f6rs\u00e4ljningar i Sverige 2013\u20132018.
Nationell exponentiell modell (ej regionalt differentierad).

\u03b2(d) = \u22120.2811 \u00d7 exp(\u22120.3811 \u00d7 d)

Reduktion = 100 \u00d7 [exp(\u03b2) \u2013 1] %

| Avst\u00e5nd | Reduktion |
|---|---|
| 0 km | \u221224.4% |
| 1 km | \u221219.1% |
| 2 km | \u221214.5% |
| 3 km | \u221210.7% |
| 5 km | \u22125.2% |
| 8 km | ~0% |""")

else:
    st.info("Skriv ett projektnamn och klicka **K\u00f6r analys** f\u00f6r att b\u00f6rja.")
    st.markdown("""
#### S\u00e5 fungerar det
1. Ange projektnamn fr\u00e5n Vindbrukskollen (t.ex. *Str\u00e4ngsered*, *Marb\u00e4ck*, *Kesemossen*)
2. Justera radie och fastighetsv\u00e4rde vid behov
3. Anv\u00e4nd **Statusfilter** f\u00f6r att s\u00f6ka bland t.ex. *Avslagna* eller *Inte aktuella* projekt
4. Klicka **K\u00f6r analys**
5. Ladda ner PDF-rapporter och interaktiv karta

#### Modeller
**W&W 2025 Regional** (standard): staggered diff-in-diff, 600 000+ transaktioner,
regionalt differentierad (Syd/\u00d6st/Norr). Region best\u00e4ms automatiskt fr\u00e5n projektets kommun.

**W&W 2021 National**: exponentiell kurva, 69 000 transaktioner, en modell f\u00f6r hela Sverige.
Ger generellt h\u00f6gre reduktion \u00e4n 2025-modellen.

V\u00e4lj modell i sidopanelen under **Modell**.

#### Datak\u00e4llor
- **Vindbrukskollen** (L\u00e4nsstyrelsen) \u2014 vindkraftverkspositioner
- **OpenStreetMap** \u2014 namngivna platser/fastigheter
- **Lantm\u00e4teriet** \u2014 topografisk karta (CC BY 4.0)
    """)
