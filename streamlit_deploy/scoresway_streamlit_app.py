from __future__ import annotations

import base64
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st


APP_DIR = Path(__file__).resolve().parent
RESOURCE_FILES = ("typeId.xlsx", "qualifiers.csv", "xT_Grid.csv")
REPORT_FILES = ("Match_Report_1.png", "Match_Report.png")
FINAL_CELL_INDEX = 65
DEFAULT_SELENIUM_WAIT = 25
PROGRESS_MARKER = "__SCORESWAY_PROGRESS__"
PROGRESS_LABELS = [
    (0, "Cargando librerias y datos de 365Scores"),
    (1, "Preparando URLs del partido"),
    (7, "Capturando endpoints de Scoresway"),
    (10, "Descargando eventos del partido"),
    (12, "Descargando formaciones y estadisticas"),
    (13, "Normalizando eventos"),
    (17, "Procesando formaciones y jugadores"),
    (28, "Calculando xT y equipos"),
    (38, "Armando redes de pases"),
    (41, "Calculando bloque defensivo"),
    (44, "Procesando remates de 365Scores"),
    (51, "Preparando mapas de tiros y momentum"),
    (56, "Construyendo mapas territoriales"),
    (60, "Construyendo entradas, centros, recuperaciones y congestion"),
    (64, "Generando pagina 1 del reporte"),
    (65, "Generando pagina 2 del reporte"),
]

APP_EXPLANATORY_TEXT = """
Esta aplicación permite generar reportes visuales a partir de datos de Scoresway y 365Scores.

La base del código fue desarrollada por @adnaaan433 para trabajar con datos de WhoScored.
En esta versión, adapté el flujo para utilizarlo con partidos de Scoresway, donde actualmente está disponible el eventing de la Liga Argentina.
Mi agradecimiento a LanusStats por facilitar la extracción de datos de 365Scores como a @Vdot_Spain por su aplicación que sirvió de inspiración de generar esta adatpación para el usuario.

Modo de uso:
- Carga los links del partido a analizar ingresando a [Scoresway](https://www.scoresway.com/en_GB/soccer) y [365Scores](https://www.365scores.com/es) y colocalos en sus respectivos campos.
- Los campos muestran un link de muestra para que se vea el formato esperado.
- Elegi los colores de cada equipo.
- Completa el subtitulo del reporte.
- Genera y descarga los dos reportes colectivos.

"""

TEXTO = """

Links y otras cosas:
- Contactame! [@robaboian](https://x.com/robaboian_)
- Github de @adnaaan433: [github.com/adnaaan433](https://github.com/adnaaan433)
- Github de @LanusStats: [LanusStats](https://github.com/federicorabanos/LanusStats)
- Aplicación de @Vdot_Spain: [Eventing2csv](https://vdotspain.shinyapps.io/Eventing2csv/)

"""
# Celdas visibles importadas desde "Copia de scoresway.ipynb".
# Podes buscar y editar aca textos, titulos, labels y bloques del reporte.
NOTEBOOK_CODE_CELLS: list[tuple[int, str]] = [
    (
        0,
        r'''
# ============================================================
# LIBRERÍAS BASE
# ============================================================
import os
import re
import ast
import glob
import json
import math
import time
import random
import shutil
import warnings
from pprint import pprint
from datetime import datetime
from io import BytesIO
from urllib.parse import urlparse, parse_qs
from urllib.request import urlopen
from zoneinfo import ZoneInfo

# ============================================================
# DATA
# ============================================================
import numpy as np
import pandas as pd
from unidecode import unidecode


def limpiar_nombre_safe(x):
    """Normaliza nombres sin romper cuando vienen NaN, None o valores no-texto."""
    if pd.isna(x):
        return ""
    return unidecode(str(x))

# ============================================================
# REQUESTS / SCRAPING
# ============================================================
import requests
from bs4 import BeautifulSoup
from LanusStats import threesixfivescores


# ============================================================
# SELENIUM
# ============================================================
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

try:
    from pyvirtualdisplay import Display
except Exception:
    Display = None


_VIRTUAL_DISPLAY = None


def iniciar_virtual_display():
    """
    En Streamlit Cloud no hay pantalla real.
    Xvfb crea una pantalla virtual para correr Chrome NO-headless.
    """
    global _VIRTUAL_DISPLAY

    if os.environ.get("DISPLAY"):
        return

    if Display is None:
        raise RuntimeError(
            "pyvirtualdisplay no esta instalado. Agrega pyvirtualdisplay a requirements.txt "
            "y xvfb a packages.txt."
        )

    if _VIRTUAL_DISPLAY is None:
        _VIRTUAL_DISPLAY = Display(visible=0, size=(1400, 1000))
        _VIRTUAL_DISPLAY.start()


def crear_driver_scoresway(headless=False):
    """
    Crea un ChromeDriver compatible con Streamlit Cloud.

    Importante:
    - NO usa webdriver-manager.
    - Usa chromium/chromedriver instalados desde packages.txt.
    - Si headless=False, usa Xvfb para simular pantalla en Streamlit Cloud.
    """
    if not headless:
        iniciar_virtual_display()

    chrome_options = Options()

    if headless:
        chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1400,1000")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--lang=en-US,en")
    chrome_options.add_argument(
        "--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    )
    chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    chromium_path = (
        shutil.which("chromium")
        or shutil.which("chromium-browser")
        or shutil.which("google-chrome")
        or shutil.which("google-chrome-stable")
    )
    chromedriver_path = shutil.which("chromedriver")

    if chromium_path:
        chrome_options.binary_location = chromium_path

    if not chromedriver_path:
        raise RuntimeError(
            "No se encontro chromedriver en el sistema. "
            "En Streamlit Cloud, revisa que packages.txt tenga chromium y chromium-driver."
        )

    driver = webdriver.Chrome(
        service=Service(chromedriver_path),
        options=chrome_options,
    )

    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    });
                """
            },
        )
        driver.execute_cdp_cmd("Network.enable", {})
    except Exception:
        pass

    return driver

# ============================================================
# VISUALIZACIÓN
# ============================================================
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import matplotlib.patches as patches
import matplotlib.patheffects as path_effects
import seaborn as sns

from matplotlib import rcParams
from matplotlib.colors import to_rgba, LinearSegmentedColormap
from matplotlib.font_manager import FontProperties
from matplotlib.gridspec import GridSpec
from matplotlib.markers import MarkerStyle
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from matplotlib.cbook import get_sample_data
from matplotlib.patheffects import withStroke, Normal

from PIL import Image

# ============================================================
# MPSOCCER
# ============================================================
from mplsoccer import Pitch, VerticalPitch, FontManager, Sbopen, add_image
from mplsoccer.utils import FontManager as MplFontManager

# ============================================================
# TEXTOS / ESTILO
# ============================================================
from highlight_text import ax_text, fig_text

# ============================================================
# MACHINE LEARNING / GEOMETRÍA
# ============================================================
from sklearn.cluster import KMeans
from scipy.spatial import ConvexHull


players365 = threesixfivescores.get_players_info("https://www.365scores.com/es/football/match/liga-profesional-72/argentinos-juniors-gimnasia-la-plata-871-880-72#id=4631685")
shots365 = threesixfivescores.get_match_shotmap("https://www.365scores.com/es/football/match/liga-profesional-72/argentinos-juniors-gimnasia-la-plata-871-880-72#id=4631685")
url_partido = "https://www.scoresway.com/en_GB/soccer/liga-profesional-argentina-2026/8v84l9nq3d5t0j4gb781i3llg/match/view/f42dnr00784z4x6i1f67dnv2s/match-summary"


pd.set_option("display.max_rows", None)     
pd.set_option("display.max_columns", None) 

# specify some custom colors to use
green = '#69f900'
red = '#ff4b44'
blue = '#00a0de'
violet = '#a369ff'
bg_color= '#f5f5f5'
line_color= '#000000'
# bg_color= '#000000'
# line_color= '#ffffff'
col1 = '#ff4b44'
col2 = '#00a0de'

hcol= col1 
acol= col2
''',
    ),
    (
        1,
        r'''

match_url = re.search(
    r"/soccer/([^/]+)/([^/]+)/match/view/([^/]+)/([^/?#]+)",
    url_partido
)

if not match_url:
    raise ValueError("La URL del partido no tiene el formato esperado.")

torneo_name = match_url.group(1)
torneo_id = match_url.group(2)
partido_id = match_url.group(3)
seccion = match_url.group(4)

print("torneo_name:", torneo_name)
print("torneo_id:", torneo_id)
print("partido_id:", partido_id)
print("seccion:", seccion)

os.makedirs("json/partidos", exist_ok=True)

print("✅ Estructura de carpetas creada (si no existía).")
''',
    ),
    (
        2,
        r'''
def preparar_urls_partido_scoresway(url_partido):
    """
    Recibe cualquier URL de partido de Scoresway, por ejemplo:
    .../match-summary
    .../player-stats
    .../formations

    Y devuelve automáticamente:
    - torneo_name
    - torneo_id
    - partido_id
    - seccion_original
    - url_match_summary
    - url_player_stats
    - url_formations
    """

    url_partido = url_partido.strip().rstrip("/")

    match_url = re.search(
        r"^(https?://www\.scoresway\.com/en_GB/soccer/([^/]+)/([^/]+)/match/view/([^/]+))/([^/?#]+)",
        url_partido
    )

    if not match_url:
        raise ValueError("La URL del partido no tiene el formato esperado.")

    base_partido = match_url.group(1)
    torneo_name = match_url.group(2)
    torneo_id = match_url.group(3)
    partido_id = match_url.group(4)
    seccion_original = match_url.group(5)

    urls = {
        "torneo_name": torneo_name,
        "torneo_id": torneo_id,
        "partido_id": partido_id,
        "seccion_original": seccion_original,
        "url_match_summary": f"{base_partido}/match-summary",
        "url_player_stats": f"{base_partido}/player-stats",
        "url_formations": f"{base_partido}/formations",
    }

    return urls
''',
    ),
    (
        3,
        r'''
datos_partido = preparar_urls_partido_scoresway(url_partido)

torneo_name = datos_partido["torneo_name"]
torneo_id = datos_partido["torneo_id"]
partido_id = datos_partido["partido_id"]

url_match_summary = datos_partido["url_match_summary"]
url_player_stats = datos_partido["url_player_stats"]
url_formations = datos_partido["url_formations"]

print("torneo_name:", torneo_name)
print("torneo_id:", torneo_id)
print("partido_id:", partido_id)
print("seccion_original:", datos_partido["seccion_original"])
print("url_match_summary:", url_match_summary)
print("url_player_stats:", url_player_stats)
print("url_formations:", url_formations)

os.makedirs("json/partidos", exist_ok=True)

print("✅ Estructura de carpetas creada.")
''',
    ),
    (
        5,
        r'''
def limpiar_jsonp(texto):
    """
    Convierte una respuesta JSONP tipo:
    callback({...})
    en un dict de Python.
    """
    texto = texto.strip()

    match = re.search(r"^[^(]*\((.*)\)\s*;?$", texto, flags=re.DOTALL)

    if not match:
        raise ValueError("La respuesta no parece tener formato JSONP válido.")

    return json.loads(match.group(1))

def extraer_datos_url_player_stats(url_player_stats):
    """
    Extrae competición, torneo_id y partido_id desde una URL de Scoresway player-stats.
    """

    match_url = re.search(
        r"/soccer/([^/]+)/([^/]+)/match/view/([^/]+)/player-stats",
        url_player_stats
    )

    if not match_url:
        raise ValueError("La URL no tiene el formato esperado de player-stats.")

    return {
        "competicion": match_url.group(1),
        "torneo_id": match_url.group(2),
        "partido_id": match_url.group(3)
    }

def obtener_callback_matchevent(info_player_stats, partido_id=None):
    """
    Busca dentro de los candidatos detectados por Selenium la URL del endpoint matchevent.
    """

    candidatos = info_player_stats.get("candidatos", [])

    for c in candidatos:
        if c.get("endpoint") != "matchevent":
            continue

        if partido_id is not None:
            recurso_id_path = c.get("recurso_id_path")
            if recurso_id_path != partido_id:
                continue

        return c.get("callback_id"), c.get("api_url")

    raise ValueError("No se encontró endpoint matchevent dentro de los candidatos.")
''',
    ),
    (
        6,
        r'''
def extraer_datos_url_formations(url_formations):
    """
    Extrae competición, torneo_id y partido_id desde una URL de Scoresway formations.
    """

    match_url = re.search(
        r"/soccer/([^/]+)/([^/]+)/match/view/([^/]+)/formations",
        url_formations
    )

    if not match_url:
        raise ValueError("La URL no tiene el formato esperado de formations.")

    return {
        "competicion": match_url.group(1),
        "torneo_id": match_url.group(2),
        "partido_id": match_url.group(3)
    }

def obtener_callback_matchstats(info_formations, partido_id=None):
    """
    Busca dentro de los candidatos detectados por Selenium la URL del endpoint matchstats.
    """

    candidatos = info_formations.get("candidatos", [])

    for c in candidatos:
        if c.get("endpoint") != "matchstats":
            continue

        if partido_id is not None:
            recurso_id_path = c.get("recurso_id_path")
            if recurso_id_path != partido_id:
                continue

        return c.get("callback_id"), c.get("api_url")

    raise ValueError("No se encontró endpoint matchstats dentro de los candidatos.")
''',
    ),
    (
        7,
        r'''
def obtener_info_api_player_stats_desde_network(url_player_stats, esperar=25, headless=False):
    # En Streamlit Cloud puede correr NO-headless con Xvfb para que Scoresway cargue igual que en local.
    driver = crear_driver_scoresway(headless=headless)

    try:
        driver.get(url_player_stats)
        time.sleep(esperar)

        logs = driver.get_log("performance")

        urls = []

        for entry in logs:
            try:
                message = json.loads(entry["message"])["message"]

                if message.get("method") == "Network.requestWillBeSent":
                    request = message.get("params", {}).get("request", {})
                    request_url = request.get("url", "")

                    if request_url:
                        urls.append(request_url)

            except Exception:
                continue

        # Quitar URLs duplicadas manteniendo el orden
        urls_unicas = list(dict.fromkeys(urls))

        with open("scoresway_player_stats_network_urls.txt", "w", encoding="utf-8") as f:
            for u in urls_unicas:
                f.write(u + "\n")

        # Una sola regex alcanza
        patron = r"https?://api\.performfeeds\.com/soccerdata/(match|matchevent|nlgdynamicplayerbio|squads)/([^/?#]+)(?:/([^?#]+))?"

        candidatos = []
        urls_ya_agregadas = set()

        for u in urls_unicas:
            match = re.search(patron, u)

            if not match:
                continue

            # Evitar duplicados exactos
            if u in urls_ya_agregadas:
                continue

            urls_ya_agregadas.add(u)

            endpoint = match.group(1)
            sdapi_outlet_key = match.group(2)
            recurso_id_path = match.group(3)  # puede ser partido_id en match/matchevent

            parsed = urlparse(u)
            params = parse_qs(parsed.query)

            candidatos.append({
                "endpoint": endpoint,
                "sdapi_outlet_key": sdapi_outlet_key,
                "recurso_id_path": recurso_id_path,
                "api_url": u,
                "callback_id": params.get("_clbk", [None])[0],
                "torneo_id_api": params.get("tmcl", [None])[0],
                "contestant_id_api": params.get("ctst", [None])[0],
                "person_id_api": params.get("prsn", [None])[0],
                "params": params
            })

        if not candidatos:
            raise ValueError(
                "No se encontró ninguna URL de api.performfeeds.com/soccerdata. "
                "Revisá scoresway_player_stats_network_urls.txt."
            )

        print("✅ Candidatos encontrados:")

        for c in candidatos:
            print("\nEndpoint:", c["endpoint"])
            print("sdapi_outlet_key:", c["sdapi_outlet_key"])
            print("recurso_id_path:", c["recurso_id_path"])
            print("callback_id:", c["callback_id"])
            print("torneo_id_api:", c["torneo_id_api"])
            print("contestant_id_api:", c["contestant_id_api"])
            print("person_id_api:", c["person_id_api"])
            print("URL:", c["api_url"])

        sdapi_outlet_key = candidatos[0]["sdapi_outlet_key"]

        print("\n✅ sdapi_outlet_key final:", sdapi_outlet_key)

        return {
            "sdapi_outlet_key": sdapi_outlet_key,
            "candidatos": candidatos,
            "urls": urls_unicas
        }

    finally:
        driver.quit()
''',
    ),
    (
        8,
        r'''
def obtener_info_api_formations_desde_network(url_formations, esperar=25, headless=False):
    # En Streamlit Cloud puede correr NO-headless con Xvfb para que Scoresway cargue igual que en local.
    driver = crear_driver_scoresway(headless=headless)

    try:
        driver.get(url_formations)
        time.sleep(esperar)

        logs = driver.get_log("performance")

        urls = []

        for entry in logs:
            try:
                message = json.loads(entry["message"])["message"]

                if message.get("method") == "Network.requestWillBeSent":
                    request = message.get("params", {}).get("request", {})
                    request_url = request.get("url", "")

                    if request_url:
                        urls.append(request_url)

            except Exception:
                continue

        urls_unicas = list(dict.fromkeys(urls))

        with open("scoresway_formations_network_urls.txt", "w", encoding="utf-8") as f:
            for u in urls_unicas:
                f.write(u + "\n")

        patron = (
            r"https?://api\.performfeeds\.com/soccerdata/"
            r"(match|matchstats|rankings|squads)/([^/?#]+)(?:/([^?#]+))?"
        )

        candidatos = []
        urls_ya_agregadas = set()

        for u in urls_unicas:
            match = re.search(patron, u)

            if not match:
                continue

            if u in urls_ya_agregadas:
                continue

            urls_ya_agregadas.add(u)

            endpoint = match.group(1)
            sdapi_outlet_key = match.group(2)
            recurso_id_path = match.group(3)

            parsed = urlparse(u)
            params = parse_qs(parsed.query)

            candidatos.append({
                "endpoint": endpoint,
                "sdapi_outlet_key": sdapi_outlet_key,
                "recurso_id_path": recurso_id_path,
                "api_url": u,
                "callback_id": params.get("_clbk", [None])[0],
                "torneo_id_api": params.get("tmcl", [None])[0],
                "contestant_id_api": params.get("ctst", [None])[0],
                "params": params
            })

        if not candidatos:
            raise ValueError(
                "No se encontró ninguna URL de api.performfeeds.com/soccerdata. "
                "Revisá scoresway_formations_network_urls.txt."
            )

        print("✅ Candidatos encontrados en formations:")

        for c in candidatos:
            print("\nEndpoint:", c["endpoint"])
            print("sdapi_outlet_key:", c["sdapi_outlet_key"])
            print("recurso_id_path:", c["recurso_id_path"])
            print("callback_id:", c["callback_id"])
            print("torneo_id_api:", c["torneo_id_api"])
            print("contestant_id_api:", c["contestant_id_api"])
            print("URL:", c["api_url"])

        sdapi_outlet_key = candidatos[0]["sdapi_outlet_key"]

        print("\n✅ sdapi_outlet_key final:", sdapi_outlet_key)

        return {
            "sdapi_outlet_key": sdapi_outlet_key,
            "candidatos": candidatos,
            "urls": urls_unicas
        }

    finally:
        driver.quit()
''',
    ),
    (
        9,
        r'''
print("url_player_stats:", url_player_stats)

info_player_stats = obtener_info_api_player_stats_desde_network(
    url_player_stats,
    esperar=25,
    headless=False
)

sdapi_outlet_key = info_player_stats["sdapi_outlet_key"]

print("sdapi_outlet_key:", sdapi_outlet_key)
''',
    ),
    (
        10,
        r'''
datos_url = extraer_datos_url_player_stats(url_player_stats)

partido_id = datos_url["partido_id"]
sdapi_outlet_key = info_player_stats["sdapi_outlet_key"]

callback_id, api_url_detectada = obtener_callback_matchevent(
    info_player_stats,
    partido_id=partido_id
)

url_matchevent = (
    f"https://api.performfeeds.com/soccerdata/matchevent/"
    f"{sdapi_outlet_key}/{partido_id}"
    f"?_rt=c&_lcl=en&_fmt=jsonp&sps=widgets&_clbk={callback_id}"
)

headers = {
    "Accept": "*/*",
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    ),
    "Referer": url_player_stats,
}

response = requests.get(url_matchevent, headers=headers, timeout=30)

print("Status:", response.status_code)

json_matchevent = limpiar_jsonp(response.text)

os.makedirs("json/partidos", exist_ok=True)

file_path = f"json/partidos/matchevent_{partido_id}.json"

with open(file_path, "w", encoding="utf-8") as f:
    json.dump(json_matchevent, f, ensure_ascii=False, indent=4)

print("✅ Guardado:", file_path)
''',
    ),
    (
        11,
        r'''
def descargar_json_matchstats_formations(
    url_formations,
    info_formations,
    carpeta_salida="json/partidos",
    esperar=True
):
    """
    Descarga el JSON puro del endpoint matchstats para la sección formations.
    """

    datos_url = extraer_datos_url_formations(url_formations)

    partido_id = datos_url["partido_id"]
    sdapi_outlet_key = info_formations["sdapi_outlet_key"]

    callback_id, api_url_detectada = obtener_callback_matchstats(
        info_formations,
        partido_id=partido_id
    )

    if callback_id is None:
        raise ValueError("Se encontró matchstats, pero no se pudo extraer callback_id.")

    os.makedirs(carpeta_salida, exist_ok=True)

    url = (
        f"https://api.performfeeds.com/soccerdata/matchstats/"
        f"{sdapi_outlet_key}/{partido_id}"
        f"?_rt=c&detailed=yes&_lcl=en&_fmt=jsonp"
        f"&sps=widgets&_clbk={callback_id}"
    )

    headers = {
        "Accept": "*/*",
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/147.0.0.0 Safari/537.36"
        ),
        "Referer": url_formations,
    }

    print("URL detectada por Selenium:")
    print(api_url_detectada)

    print("\nURL reconstruida:")
    print(url)

    if esperar:
        espera = random.uniform(2, 5)
        print(f"\n⏳ Esperando {espera:.2f} segundos antes de descargar...")
        time.sleep(espera)

    response = requests.get(url, headers=headers, timeout=30)

    print("Status:", response.status_code)

    if response.status_code != 200:
        print(response.text[:500])
        raise RuntimeError(f"Error descargando matchstats: {response.status_code}")

    json_data = limpiar_jsonp(response.text)

    file_name = f"matchstats_formations_{partido_id}.json"
    file_path = os.path.join(carpeta_salida, file_name)

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)

    print(f"✅ JSON guardado en: {file_path}")

    return json_data
''',
    ),
    (
        12,
        r'''
info_formations = obtener_info_api_formations_desde_network(
    url_formations=url_formations,
    esperar=25,
    headless=False
)

json_matchstats = descargar_json_matchstats_formations(
    url_formations=url_formations,
    info_formations=info_formations,
    carpeta_salida="json/formations",
    esperar=True
)
''',
    ),
    (
        13,
        r'''

# ==== Config ====
PARTIDOS_DIR = "json/partidos"                   # carpeta con tu JSON
TYPEID_XLSX_PATH = "typeId.xlsx"  # Excel con columnas: typeId, EVENT NAME

# ==== 1) Cargar mapping (id -> nombre) ====
type_map_df = pd.read_excel(TYPEID_XLSX_PATH, dtype={"typeId": "Int64"})
TYPE_MAP = dict(zip(type_map_df["typeId"], type_map_df["EVENT NAME"]))

# ==== 2) Cargar JSONs y normalizar eventos ====
all_event_dfs = []
files = sorted(glob.glob(os.path.join(PARTIDOS_DIR, "*.json")))
if not files:
    raise FileNotFoundError(f"No se encontraron JSON en {PARTIDOS_DIR}")

for path in files:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    event_df = pd.json_normalize(data["liveData"]["event"])

    # ==== 3) Reemplazar IDs por nombres en typeId ====
    event_df["typeId"] = pd.to_numeric(event_df["typeId"], errors="coerce").astype("Int64")
    event_df["typeId"] = event_df["typeId"].map(TYPE_MAP).fillna(event_df["typeId"])

    all_event_dfs.append(event_df)

# ==== 4) Unir todo en un único DF ====
df = pd.concat(all_event_dfs, ignore_index=True)
''',
    ),
    (
        14,
        r'''

# ==== 1) Cargar mapping de qualifiers ====
QUALIFIERS_CSV_PATH = "qualifiers.csv"

# El csv tiene columnas qualifierId y QUALIFIER NAME
qual_df = pd.read_csv(QUALIFIERS_CSV_PATH, sep=";", engine="python")
QUAL_MAP = dict(zip(qual_df["qualifierId"], qual_df["QUALIFIER NAME"]))

# ==== 2) Función para convertir string a lista de dicts ====
def to_list_safe(x):
    if isinstance(x, list):
        return x
    if isinstance(x, str):
        try:
            v = ast.literal_eval(x)
            return v if isinstance(v, list) else []
        except Exception:
            return []
    return [] if pd.isna(x) else []

# ==== 3) Función para mapear qualifierIds ====
def map_qualifiers_cell(x):
    lst = to_list_safe(x)
    out = []
    for item in lst:
        if isinstance(item, dict) and "qualifierId" in item:
            qid = pd.to_numeric(item["qualifierId"], errors="coerce")
            if pd.notna(qid):
                qid = int(qid)
                item = item.copy()
                item["qualifierId"] = QUAL_MAP.get(qid, item["qualifierId"])
        out.append(item)
    return out

# ==== 4) Aplicar sobre la columna qualifier ====
df["qualifier"] = df["qualifier"].apply(map_qualifiers_cell)

''',
    ),
    (
        15,
        r'''
# --- util: asegurar lista de dicts en df['qualifiers'] ---
def to_list_safe(x):
    if isinstance(x, list):
        return x
    if isinstance(x, str):
        try:
            v = ast.literal_eval(x)
            return v if isinstance(v, list) else []
        except Exception:
            return []
    return [] if pd.isna(x) else []

df["qualifier"] = df["qualifier"].apply(to_list_safe)

# --- construir el mapeo playerId -> shirtNo a partir de filas especiales ---
def extract_player_shirt_map_from_row(quals):
    """Dado una lista de qualifiers, si contiene JerseyNumber e InvolvedPlayers,
       devuelve un dict {playerId: shirtNo}. Si no, {}.
    """
    if not isinstance(quals, list) or not quals:
        return {}

    # buscar los dos qualifiers
    q_jersey = next((q for q in quals if isinstance(q, dict) and q.get("qualifierId") == "JerseyNumber"), None)
    q_players = next((q for q in quals if isinstance(q, dict) and q.get("qualifierId") == "InvolvedPlayers"), None)
    if not q_jersey or not q_players:
        return {}

    # listas ordenadas
    jerseys = [s.strip() for s in str(q_jersey.get("value", "")).split(",") if s.strip() != ""]
    players = [s.strip() for s in str(q_players.get("value", "")).split(",") if s.strip() != ""]

    # si difieren en longitud, truncamos al mínimo para evitar desalineaciones
    n = min(len(jerseys), len(players))
    jerseys = jerseys[:n]
    players = players[:n]

    # construir dict
    # jersey a int cuando se pueda, sino dejar string
    def to_int_or_str(x):
        try:
            return int(x)
        except Exception:
            return x

    return {players[i]: to_int_or_str(jerseys[i]) for i in range(n)}

# recolectar mapas de todas las filas que tengan ambos qualifiers (suelen ser 2)
maps = []
for quals in df["qualifier"]:
    m = extract_player_shirt_map_from_row(quals)
    if m:
        maps.append(m)

# unir los dicts (si hay keys repetidas, el último visto pisa al anterior)
player_to_shirt = {}
for m in maps:
    player_to_shirt.update(m)

# --- crear la columna shirtNo para cada evento, usando df['playerId'] ---
# (ajustá el nombre si tu columna de jugador se llama distinto)
if "playerId" not in df.columns:
    raise KeyError("No encuentro la columna 'playerId' en df. Ajustá el nombre si es distinto.")

df["shirtNo"] = df["playerId"].map(player_to_shirt).astype("Int64")
''',
    ),
    (
        16,
        r'''
# =========================
# SOLO PARA INSPECCIÓN
# =========================
FORMACION_DIR = "json/formations"

formacion_files = sorted(glob.glob(os.path.join(FORMACION_DIR, "*.json")))
if not formacion_files:
    raise FileNotFoundError(f"No se encontraron JSON en {FORMACION_DIR}")

print(f"JSONs encontrados en {FORMACION_DIR}:")
for p in formacion_files:
    print(" -", os.path.basename(p))

# cargo el primero para inspección
with open(formacion_files[0], "r", encoding="utf-8") as f:
    formacion_json = json.load(f)

print("\nArchivo cargado para inspección:", os.path.basename(formacion_files[0]))
print("Tipo raíz:", type(formacion_json))

# si es dict, mirás las keys vos
if isinstance(formacion_json, dict):
    print("Keys nivel raíz:")
    print(list(formacion_json.keys()))

# si es lista
elif isinstance(formacion_json, list):
    print("Cantidad de elementos:", len(formacion_json))
''',
    ),
    (
        17,
        r'''
lineup_raw = formacion_json["liveData"]["lineUp"]

df_lineup = pd.json_normalize(
    lineup_raw,
)

df_lineup_players = df_lineup[["player", "stat"]].copy()


players_df = (
    df_lineup_players[["player"]]   # tomo SOLO la columna player
    .explode("player")      # junto jugadores de la fila 1 y 2
    .reset_index(drop=True)
)

players_df = pd.concat(
    [
        players_df.drop(columns=["player"]).reset_index(drop=True),
        pd.json_normalize(players_df["player"]).reset_index(drop=True)
    ],
    axis=1
)


# 1) Explode stats (1 fila por stat por jugador) + índice limpio
stats_long = (
    players_df
    .reset_index(drop=True)
    .reset_index(names="player_row")   # id único por jugador
    .explode("stat")
    .reset_index(drop=True)            # <- CLAVE para que concat no rompa
)

# 2) Abrir dict {type, value} sin problemas de alineación
stat_norm = pd.json_normalize(stats_long["stat"]).reset_index(drop=True)

stats_long = pd.concat(
    [stats_long.drop(columns=["stat"]).reset_index(drop=True), stat_norm],
    axis=1
)

# 3) Pivot: type -> columnas, value -> valores
stats_wide = (
    stats_long
    .pivot_table(
        index="player_row",
        columns="type",
        values="value",
        aggfunc="first"
    )
)

# 4) Merge final al players_df
players_df = (
    players_df
    .reset_index(drop=True)
    .reset_index(names="player_row")
    .drop(columns=["stat"])
    .merge(stats_wide, on="player_row", how="left")
    .drop(columns=["player_row"])
)

players_df
''',
    ),
    (
        18,
        r'''
df = df.rename(columns={
    "typeId": "type",
    "outcome": "outcomeType",
    "contestantId": "teamId",
    "timeMin": "minute",
    "timeSec": "second",
    "periodId": "period",
    "qualifier": "qualifiers"
})

df["outcomeType"] = df["outcomeType"].map({1: "Successful", 0: "Unsuccessful"})
''',
    ),
    (
        19,
        r'''
def cumulative_match_mins(events_df):
    events_out = pd.DataFrame()
    # Add cumulative time to events data, resetting for each unique match
    match_events = events_df.copy()
    match_events['cumulative_mins'] = match_events['minute'] + (1/60) * match_events['second']
    # Add time increment to cumulative minutes based on period of game.
    for period in np.arange(1, match_events['period'].max() + 1, 1):
        if period > 1:
            t_delta = match_events[match_events['period'] == period - 1]['cumulative_mins'].max() - \
                                   match_events[match_events['period'] == period]['cumulative_mins'].min()
        elif period == 1 or period == 5:
            t_delta = 0
        else:
            t_delta = 0
        match_events.loc[match_events['period'] == period, 'cumulative_mins'] += t_delta
    # Rebuild events dataframe
    events_out = pd.concat([events_out, match_events])
    return events_out
''',
    ),
    (
        20,
        r'''
df = cumulative_match_mins(df)
''',
    ),
    (
        21,
        r'''
event_list = [
    'BallTouch', 'BlockedPass', 'Claim', 'Clearance', 'Dispossessed',
    'Foul', 'Goal', 'GoodSkill', 'Interception', 'MissedShots',
    'OffsidePass', 'Pass', 'Punch', 'SavedShot', 'Save',
    'ShotOnPost', 'Smother', 'Tackle', 'TakeOn',
    'KeeperSweeper'
]

shots = ["SavedShot", "ShotOnPost", "MissedShots", "Goal"]

# nueva columna booleana
df["isTouch"] = df["type"].isin(event_list)
df["isGoal"] = df["type"] == "Goal"
df["isShot"] = df['type'].isin(shots)
''',
    ),
    (
        22,
        r'''
# --- helper: asegurar lista de dicts ---
def to_list_safe(x):
    if isinstance(x, list):
        return x
    if isinstance(x, str):
        try:
            v = ast.literal_eval(x)
            return v if isinstance(v, list) else []
        except Exception:
            return []
    return [] if pd.isna(x) else []

# --- helper: convertir '85.9' o '85,9' a número si se puede ---
def to_number_if_possible(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return pd.NA
    s = str(v).strip()
    # permite coma decimal
    s_norm = s.replace(",", ".")
    try:
        return float(s_norm)
    except Exception:
        return s  # deja string si no es numérico

# sets para saber qué columnas ya creamos
created_bool_cols = set()
created_value_cols = set()

# Recorremos fila a fila
for idx, quals in df["qualifiers"].apply(to_list_safe).items():
    if not quals:
        continue
    for q in quals:
        if not isinstance(q, dict):
            continue

        qid = q.get("qualifierId")
        if qid is None:
            continue
        qid = str(qid)  # nombre de columna

        # --- FIX clave: decidir si tiene 'value' usable sin comparar con pd.NA en tupla ---
        val = q.get("value", None)
        has_value = (val is not None) and (not pd.isna(val)) and (str(val).strip() != "")

        if has_value:
            # columna de valor (ej.: PassEndX, Angle, Length, Zone, etc.)
            if qid not in created_value_cols:
                df[qid] = pd.NA
                created_value_cols.add(qid)
            df.at[idx, qid] = to_number_if_possible(val)
        else:
            # columna booleana is<Qualifier> (ej.: isLaunch, isLongball, etc.)
            col_bool = f"is{qid}"
            if col_bool not in created_bool_cols:
                df[col_bool] = False
                created_bool_cols.add(col_bool)
            df.at[idx, col_bool] = True


# Opcional: convertir columnas de valores a numérico cuando la columna sea realmente numérica.
# En pandas nuevo, errors="ignore" puede fallar; por eso usamos coerce solo para testear
# y convertimos únicamente si no destruye valores de texto.
for col in created_value_cols:
    serie_original = df[col]
    mask_con_dato = serie_original.notna()

    if not mask_con_dato.any():
        continue

    serie_convertida = pd.to_numeric(serie_original, errors="coerce")

    # Convertir solo si todos los valores existentes pudieron pasar a número.
    # Si hay valores de texto, dejamos la columna como estaba.
    if serie_convertida[mask_con_dato].notna().all():
        df[col] = serie_convertida

# Listo: df ahora tiene columnas nuevas como:
# - isLaunch (True/False)
# - PassEndX, PassEndY, Angle, Length, Zone, etc. con valores por fila
''',
    ),
    (
        23,
        r'''
df['endX'] = df['PassEndX']
df['endY'] = df['PassEndY']
''',
    ),
    (
        24,
        r'''
tipos_a_eliminar = ['Out', 'Deleted event', 'Attempted tackle',
                    'Obstacle', 'Drop of Ball', 'Injury Time Announcement',
                    'Deleted After Review', 'Start delay', 'Contentious referee decision',
                    'End delay', 'Referee Drop Ball', 'Condition change', 'Collection End']

# filtrar dejando solo los que NO están en la lista
df = df[~df["type"].isin(tipos_a_eliminar)].reset_index(drop=False)
''',
    ),
    (
        25,
        r'''
# Aseguramos tipos numéricos
df["period"] = pd.to_numeric(df["period"], errors="coerce").astype("Int64")
df["minute"] = pd.to_numeric(df["minute"], errors="coerce")

# Inicializamos la columna; no modificamos otras filas (periods fuera de 1..5)
if "expandedMinute" not in df.columns:
    df["expandedMinute"] = np.nan

allowed_periods = {1, 2, 3, 4, 5}
prev_max = -1  # último expanded asignado

# Iterar solo por los periods permitidos y presentes
for p in sorted(set(df["period"].dropna()) & allowed_periods):
    mask = (df["period"] == p) & df["minute"].notna()
    if not mask.any():
        continue

    start_minute = df.loc[mask, "minute"].min()  # típicamente 0, 45, 90, 105, 0
    offset = (prev_max + 1) - start_minute

    df.loc[mask, "expandedMinute"] = df.loc[mask, "minute"] + offset
    prev_max = df.loc[mask, "expandedMinute"].max()

# (Opcional) si la querés como entero con NA seguros:
# df["expandedMinute"] = pd.to_numeric(df["expandedMinute"], errors="coerce").round().astype("Int64")
''',
    ),
    (
        26,
        r'''
def insert_ball_carries(events_df, min_carry_length=3, max_carry_length=60, min_carry_duration=1, max_carry_duration=10):
    events_out = pd.DataFrame()
    # Carry conditions (convert from metres to opta)
    min_carry_length = 3.0
    max_carry_length = 60.0
    min_carry_duration = 1.0
    max_carry_duration = 10.0
    # match_events = events_df[events_df['match_id'] == match_id].reset_index()
    match_events = events_df.reset_index()
    match_carries = pd.DataFrame()
    
    for idx, match_event in match_events.iterrows():

        if idx < len(match_events) - 1:
            prev_evt_team = match_event['teamId']
            next_evt_idx = idx + 1
            init_next_evt = match_events.loc[next_evt_idx]
            take_ons = 0
            incorrect_next_evt = True

            while incorrect_next_evt:

                next_evt = match_events.loc[next_evt_idx]

                if next_evt['type'] == 'TakeOn' and next_evt['outcomeType'] == 'Successful':
                    take_ons += 1
                    incorrect_next_evt = True

                elif ((next_evt['type'] == 'TakeOn' and next_evt['outcomeType'] == 'Unsuccessful')
                      or (next_evt['teamId'] != prev_evt_team and next_evt['type'] == 'Challenge' and next_evt['outcomeType'] == 'Unsuccessful')
                      or (next_evt['type'] == 'Foul')):
                    incorrect_next_evt = True

                else:
                    incorrect_next_evt = False

                next_evt_idx += 1

            # Apply some conditioning to determine whether carry criteria is satisfied
            same_team = prev_evt_team == next_evt['teamId']
            not_ball_touch = match_event['type'] != 'BallTouch'
            dx = 105*(match_event['endX'] - next_evt['x'])/100
            dy = 68*(match_event['endY'] - next_evt['y'])/100
            far_enough = dx ** 2 + dy ** 2 >= min_carry_length ** 2
            not_too_far = dx ** 2 + dy ** 2 <= max_carry_length ** 2
            dt = 60 * (next_evt['cumulative_mins'] - match_event['cumulative_mins'])
            min_time = dt >= min_carry_duration
            same_phase = dt < max_carry_duration
            same_period = match_event['period'] == next_evt['period']

            valid_carry = same_team & not_ball_touch & far_enough & not_too_far & min_time & same_phase &same_period

            if valid_carry:
                carry = pd.DataFrame()
                prev = match_event
                nex = next_evt

                carry.loc[0, 'eventId'] = prev['eventId'] + 0.5
                carry['minute'] = np.floor(((init_next_evt['minute'] * 60 + init_next_evt['second']) + (
                        prev['minute'] * 60 + prev['second'])) / (2 * 60))
                carry['second'] = (((init_next_evt['minute'] * 60 + init_next_evt['second']) +
                                    (prev['minute'] * 60 + prev['second'])) / 2) - (carry['minute'] * 60)
                carry['teamId'] = nex['teamId']
                carry['x'] = prev['endX']
                carry['y'] = prev['endY']
                carry['expandedMinute'] = np.floor(((init_next_evt['expandedMinute'] * 60 + init_next_evt['second']) +
                                                    (prev['expandedMinute'] * 60 + prev['second'])) / (2 * 60))
                carry['period'] = nex['period']
                carry['type'] = carry.apply(lambda x: {'value': 99, 'displayName': 'Carry'}, axis=1)
                carry['outcomeType'] = 'Successful'
                carry['qualifiers'] = carry.apply(lambda x: {'type': {'value': 999, 'displayName': 'takeOns'}, 'value': str(take_ons)}, axis=1)
                carry['satisfiedEventsTypes'] = carry.apply(lambda x: [], axis=1)
                carry['isTouch'] = True
                carry['playerId'] = nex['playerId']
                carry['endX'] = nex['x']
                carry['endY'] = nex['y']
                carry['blockedX'] = np.nan
                carry['blockedY'] = np.nan
                carry['goalMouthZ'] = np.nan
                carry['goalMouthY'] = np.nan
                carry['isShot'] = np.nan
                carry['relatedEventId'] = nex['eventId']
                carry['relatedPlayerId'] = np.nan
                carry['isGoal'] = np.nan
                carry['cardType'] = np.nan
                carry['isOwnGoal'] = np.nan
                carry['type'] = 'Carry'
                carry['cumulative_mins'] = (prev['cumulative_mins'] + init_next_evt['cumulative_mins']) / 2

                match_carries = pd.concat([match_carries, carry], ignore_index=True, sort=False)

    match_events_and_carries = pd.concat([match_carries, match_events], ignore_index=True, sort=False)
    match_events_and_carries = match_events_and_carries.sort_values(['period', 'cumulative_mins']).reset_index(drop=True)

    # Rebuild events dataframe
    events_out = pd.concat([events_out, match_events_and_carries])

    return events_out
''',
    ),
    (
        27,
        r'''
df = insert_ball_carries(df, min_carry_length=3, max_carry_length=60, min_carry_duration=1, max_carry_duration=10)
''',
    ),
    (
        28,
        r'''
# ===============================
# 0) Index estable para merge
# ===============================
df.columns = df.columns.str.strip()                 # evita 'type ' vs 'type'
df = df.reset_index(drop=True)
df["index"] = np.arange(1, len(df) + 1)
df = df[["index"] + [c for c in df.columns if c != "index"]]

# ===============================
# 1) Copia base + filtros xT
# ===============================
dfxT = df.copy()

# Evitar dropear filas por NaN en qualifiers
dfxT["qualifiers"] = dfxT["qualifiers"].astype(str)
dfxT = dfxT[~dfxT["qualifiers"].str.contains("Corner", na=False)]

# Sólo Pass/Carry exitosos
dfxT = dfxT[dfxT["type"].isin(["Pass", "Carry"]) & (dfxT["outcomeType"] == "Successful")]

# ===============================
# 2) Cargar grid xT y binning
# ===============================
xT = pd.read_csv("xT_Grid.csv", header=None).to_numpy()
xT_rows, xT_cols = xT.shape

# Asegurar que coords están en numérico (por si vienen como str con coma)
for c in ["x","y","endX","endY"]:
    if c in dfxT.columns:
        dfxT[c] = pd.to_numeric(dfxT[c], errors="coerce")

# Bins 0..xT_cols / 0..xT_rows (asume 0–100 en Opta)
dfxT["x1_bin_xT"] = pd.cut(dfxT["x"],    bins=xT_cols, labels=False, include_lowest=True)
dfxT["y1_bin_xT"] = pd.cut(dfxT["y"],    bins=xT_rows, labels=False, include_lowest=True)
dfxT["x2_bin_xT"] = pd.cut(dfxT["endX"], bins=xT_cols, labels=False, include_lowest=True)
dfxT["y2_bin_xT"] = pd.cut(dfxT["endY"], bins=xT_rows, labels=False, include_lowest=True)

# Tomar valores; si algún bin es NaN, dejar NaN
def _xt_val(v):
    xb, yb = v
    if pd.isna(xb) or pd.isna(yb):
        return np.nan
    return xT[int(yb)][int(xb)]

dfxT["start_zone_value_xT"] = dfxT[["x1_bin_xT","y1_bin_xT"]].apply(_xt_val, axis=1)
dfxT["end_zone_value_xT"]   = dfxT[["x2_bin_xT","y2_bin_xT"]].apply(_xt_val, axis=1)

dfxT["xT"] = dfxT["end_zone_value_xT"] - dfxT["start_zone_value_xT"]

# ===============================
# 3) QUEDARSE SÓLO CON LO NUEVO
#    (para no chocar nombres)
# ===============================
keep_new = [
    "index",                      # clave de merge
    "xT",
    "start_zone_value_xT", "end_zone_value_xT",
    "x1_bin_xT", "y1_bin_xT", "x2_bin_xT", "y2_bin_xT"
]
dfxT = dfxT[keep_new].copy()

# Si ya existían columnas xT en df (de una corrida anterior), las eliminamos
df = df.drop(columns=[c for c in keep_new if c != "index"], errors="ignore")

# ===============================
# 4) MERGE limpio (sin sufijos)
# ===============================
df = df.merge(dfxT, on="index", how="left", validate="one_to_one")
''',
    ),
    (
        29,
        r'''
teams_dict = data['matchInfo']['contestant']  # lista de dicts como la que pasaste

# =====================================================
# 1) mapping teamId -> name (o name)
# =====================================================
teamid_to_short = {t["id"]: t.get("name", t.get("name")) for t in teams_dict}
df["teamName"] = df["teamId"].map(teamid_to_short)

# =====================================================
# 2) identificar local y visitante por "position"
# =====================================================
home_team = next(t for t in teams_dict if t.get("position") == "home")
away_team = next(t for t in teams_dict if t.get("position") == "away")

hteamID = home_team["id"]
ateamID = away_team["id"]

hteamName = home_team.get("name", home_team.get("name"))
ateamName = away_team.get("name", away_team.get("name"))

# =====================================================
# 3) columna Home/Away (teamVenue)
# =====================================================
df["teamVenue"] = df["teamId"].map({hteamID: "Home", ateamID: "Away"})

# (si querés dejar NaN para eventos sin teamId, esto queda ok;
#  si preferís algo explícito:)
# df["teamVenue"] = df["teamVenue"].fillna("Unknown")

# =====================================================
# 4) oposición (rival)
# =====================================================
df["oppositionTeamName"] = df["teamVenue"].map({
    "Home": ateamName,
    "Away": hteamName
})
''',
    ),
    (
        30,
        r'''
# Reshaping the data from 100x100 to 105x68, as I use the pitch_type='uefa', in the pitch function, you can consider according to your use

df['x'] = df['x']*1.05
df['y'] = df['y']*0.68
df['endX'] = df['endX']*1.05
df['endY'] = df['endY']*0.68
df['goalMouthY'] = df['goalMouthY']*0.68
''',
    ),
    (
        31,
        r'''
df['qualifiers'] = df['qualifiers'].astype(str)
# Calculating passing distance, to find out progressive pass, this will just show the distance reduced by a pass, then will be able to filter passes which has reduced distance value more than 10yds as a progressive pass
df['prog_pass'] = np.where((df['type'] == 'Pass'), 
                           np.sqrt((105 - df['x'])**2 + (34 - df['y'])**2) - np.sqrt((105 - df['endX'])**2 + (34 - df['endY'])**2), 0)
# Calculating carrying distance, to find out progressive carry, this will just show the distance reduced by a carry, then will be able to filter carries which has reduced distance value more than 10yds as a progressive carry
df['prog_carry'] = np.where((df['type'] == 'Carry'), 
                            np.sqrt((105 - df['x'])**2 + (34 - df['y'])**2) - np.sqrt((105 - df['endX'])**2 + (34 - df['endY'])**2), 0)
df['pass_or_carry_angle'] = np.degrees(np.arctan2(df['endY'] - df['y'], df['endX'] - df['x']))
''',
    ),
    (
        32,
        r'''
df = df.rename(columns={
    "playerName": "name"})

df['name'] = df['name'].apply(limpiar_nombre_safe)
''',
    ),
    (
        33,
        r'''
def get_possession_chains(events_df, chain_check=5, suc_evts_in_chain=3):
    # 1) Copia y normalización de índice (sin crear columnas 'index'/'level_0')
    match_events_df = events_df.copy()
    match_events_df = match_events_df.reset_index(drop=True)

    # 2) Filtrar eventos que no contribuyen a posesión
    ignore_types = [
        'OffsideProvoken', 'CornerAwarded', 'Start', 'Card', 'SubstitutionOff',
        'SubstitutionOn', 'FormationChange', 'FormationSet', 'End'
    ]
    match_pos_events_df = match_events_df[~match_events_df['type'].isin(ignore_types)].copy()

    # 3) Variables auxiliares
    match_pos_events_df['outcomeBinary'] = (match_pos_events_df['outcomeType'] == 'Successful').astype(int)

    # equipo "0/1": usa el menor nombre alfabético para anclar (como tu código)
    base_team = match_pos_events_df['teamName'].min()
    match_pos_events_df['teamBinary'] = (match_pos_events_df['teamName'] == base_team).astype(int)

    # gol detectado por transición a 'Goal' (como tenías)
    goal_hit = (match_pos_events_df['type'] == 'Goal').astype(int)
    match_pos_events_df['goalBinary'] = goal_hit.diff(1).apply(lambda x: 1 if x < 0 else 0).fillna(0).astype(int)

    # 4) DataFrame para cadenas, con el mismo índice
    pos_chain_df = pd.DataFrame(index=match_pos_events_df.index)

    # si los próximos (chain_check-1) eventos son del mismo team
    for n in range(1, chain_check):
        col = f'evt_{n}_same_team'
        # diff hacia -n: 0 si mismo team, 1 si diferente; abs por seguridad
        pos_chain_df[col] = (match_pos_events_df['teamBinary'].diff(periods=-n)).abs()
        # clamp a {0,1}
        pos_chain_df[col] = pos_chain_df[col].apply(lambda x: 1 if x > 1 else (0 if pd.isna(x) else int(x)))

    # suficientes eventos del mismo team (tu lógica original)
    pos_chain_df['enough_evt_same_team'] = pos_chain_df.sum(axis=1).apply(
        lambda x: 1 if x < chain_check - suc_evts_in_chain else 0
    )
    pos_chain_df['enough_evt_same_team'] = pos_chain_df['enough_evt_same_team'].diff(1).fillna(0)
    pos_chain_df.loc[pos_chain_df['enough_evt_same_team'] < 0, 'enough_evt_same_team'] = 0

    # asegurar period numérico
    match_pos_events_df['period'] = pd.to_numeric(match_pos_events_df['period'], errors='coerce')

    # 5) Kick-offs (por gol o cambio de periodo) en la ventana futura
    pos_chain_df['upcoming_ko'] = 0
    # índices donde hay KO: cambio de periodo o goalBinary==1
    ko_idx = match_pos_events_df[
        (match_pos_events_df['goalBinary'] == 1) | (match_pos_events_df['period'].diff(1).fillna(0) != 0)
    ].index

    # marcar ventana [ko_pos - suc_evts_in_chain, ko_pos) en 'upcoming_ko'
    for ko in ko_idx:
        start = max(match_pos_events_df.index.get_loc(ko) - suc_evts_in_chain, 0)
        stop  = match_pos_events_df.index.get_loc(ko)
        # traducir posiciones a labels reales de índice
        idx_slice = match_pos_events_df.index[start:stop]
        pos_chain_df.loc[idx_slice, 'upcoming_ko'] = 1

    # 6) Posición válida de inicio de posesión
    pos_chain_df['valid_pos_start'] = pos_chain_df['enough_evt_same_team'].fillna(0) - pos_chain_df['upcoming_ko'].fillna(0)

    # 7) Sumar inicios por KO explícitos
    pos_chain_df['kick_off_period_change'] = match_pos_events_df['period'].diff(1).fillna(0)
    pos_chain_df['kick_off_goal'] = match_pos_events_df['goalBinary']  # ya es 0/1

    pos_chain_df.loc[pos_chain_df['kick_off_period_change'] == 1, 'valid_pos_start'] = 1
    pos_chain_df.loc[pos_chain_df['kick_off_goal'] == 1, 'valid_pos_start'] = 1

    # 8) Semillas iniciales
    pos_chain_df['teamName'] = match_pos_events_df['teamName']
    first_idx = pos_chain_df.index.min()
    pos_chain_df.loc[first_idx, 'valid_pos_start'] = 1
    pos_chain_df.loc[first_idx, 'possession_id'] = 1
    pos_chain_df.loc[first_idx, 'possession_team'] = pos_chain_df.loc[first_idx, 'teamName']

    # 9) Asignación de IDs de posesión
    valid_pos_start_id = pos_chain_df.index[pos_chain_df['valid_pos_start'] > 0].tolist()
    possession_id = 2
    for i in range(1, len(valid_pos_start_id)):
        curr = valid_pos_start_id[i]
        prev = valid_pos_start_id[i - 1]
        current_team = pos_chain_df.at[curr, 'teamName']
        previous_team = pos_chain_df.at[prev, 'teamName']
        if (previous_team == current_team) and (pos_chain_df.at[curr, 'kick_off_goal'] != 1) and (pos_chain_df.at[curr, 'kick_off_period_change'] != 1):
            pos_chain_df.at[curr, 'possession_id'] = np.nan
        else:
            pos_chain_df.at[curr, 'possession_id'] = possession_id
            pos_chain_df.at[curr, 'possession_team'] = current_team
            possession_id += 1

    # 10) Devolver al DF original por índice alineado
    out = match_events_df.merge(
        pos_chain_df[['possession_id', 'possession_team']],
        left_index=True, right_index=True, how='left'
    )

    # 11) Rellenar hacia adelante/atrás
    out[['possession_id', 'possession_team']] = out[['possession_id', 'possession_team']].ffill().bfill()

    return out
''',
    ),
    (
        34,
        r'''
df = get_possession_chains(df, 5, 3)
''',
    ),
    (
        35,
        r'''
df['period'] = df['period'].astype(int)

df['period'] = df['period'].replace({1: 'FirstHalf', 2: 'SecondHalf', 3: 'FirstPeriodOfExtraTime', 4: 'SecondPeriodOfExtraTime', 
                                     5: 'PenaltyShootout', 14: 'PostGame', 16: 'PreMatch'})
''',
    ),
    (
        36,
        r'''
# 1) Creamos tabla de referencia playerId -> shirtNo, name
player_ref = (
    df
    .loc[df['playerId'].notna(), ['playerId', 'shirtNo', 'name']]
    .dropna(subset=['shirtNo', 'name'])
    .drop_duplicates(subset='playerId')
    .set_index('playerId')
)

# 2) Máscara para filas Carry sin nombre o camiseta
mask_carry = (
    (df['type'] == 'Carry') &
    (df['playerId'].notna()) &
    (df['name'].isna() | df['shirtNo'].isna())
)

# 3) Asignamos name y shirtNo según playerId
df.loc[mask_carry, 'name'] = (
    df.loc[mask_carry, 'playerId']
      .map(player_ref['name'])
)

df.loc[mask_carry, 'shirtNo'] = (
    df.loc[mask_carry, 'playerId']
      .map(player_ref['shirtNo'])
)
''',
    ),
    (
        37,
        r'''
df.teamName.unique()
''',
    ),
    (
        38,
        r'''
players_df["isFirstEleven"] = np.where(
    players_df["formationPlace_x"].notna(),
    True,
    np.nan
)

players_df['name'] = players_df['matchName']
players_df["shirtNo"] = players_df["shirtNumber"]

def get_passes_df(df):
    df1 = df[~df['type'].str.contains('SubstitutionOn|FormationChange|FormationSet|Card')]
    df = df1
    df.loc[:, "receiver"] = df["playerId"].shift(-1)
    passes_ids = df.index[df['type'] == 'Pass']
    df_passes = df.loc[passes_ids, ["index", "x", "y", "endX", "endY", "teamName", "playerId", "receiver", "type", "outcomeType", "pass_or_carry_angle"]]

    return df_passes

passes_df = get_passes_df(df)
path_eff = [path_effects.Stroke(linewidth=3, foreground=bg_color), path_effects.Normal()]
''',
    ),
    (
        39,
        r'''
def get_passes_between_df(teamName, passes_df, players_df):
    passes_df = passes_df[(passes_df["teamName"] == teamName)].copy()

    dfteam = df[(df["teamName"] == teamName) &
                (~df["type"].str.contains("SubstitutionOn|FormationChange|FormationSet|Card", na=False))].copy()

    # IDs alfanuméricos -> string
    passes_df["playerId"] = passes_df["playerId"].astype("string")
    passes_df["receiver"] = passes_df["receiver"].astype("string")
    players_df = players_df.copy()
    players_df["playerId"] = players_df["playerId"].astype("string")

    # Merge titulares
    if "isFirstEleven" in players_df.columns:
        passes_df = passes_df.merge(players_df[["playerId", "isFirstEleven"]], on="playerId", how="left")
    else:
        passes_df["isFirstEleven"] = np.nan

    # Median positions
    average_locs_and_count_df = dfteam.groupby("playerId").agg({"x": ["median"], "y": ["median", "count"]})
    average_locs_and_count_df.columns = ["pass_avg_x", "pass_avg_y", "count"]
    average_locs_and_count_df = average_locs_and_count_df.reset_index()
    average_locs_and_count_df["playerId"] = average_locs_and_count_df["playerId"].astype("string")

    # ---- Merge info jugadores (name/shirtNo ya existen, pero aseguramos position/positionSide) ----
    # Detecto columnas reales en players_df (por si alguna vez quedaron con sufijos)
    pos_candidates = ["position", "Position", "position_x", "position_y"]
    side_candidates = ["positionSide", "positionSide_x", "positionSide_y"]

    pos_col = next((c for c in pos_candidates if c in players_df.columns), None)
    side_col = next((c for c in side_candidates if c in players_df.columns), None)

    need = ["playerId", "name", "shirtNo", "isFirstEleven"]
    if pos_col is not None:
        need.append(pos_col)
    if side_col is not None:
        need.append(side_col)

    average_locs_and_count_df = average_locs_and_count_df.merge(players_df[need], on="playerId", how="left")

    # Normalizo nombres para que SIEMPRE existan así:
    if pos_col is not None and pos_col != "position":
        average_locs_and_count_df = average_locs_and_count_df.rename(columns={pos_col: "position"})
    if side_col is not None and side_col != "positionSide":
        average_locs_and_count_df = average_locs_and_count_df.rename(columns={side_col: "positionSide"})

    # Si por algún motivo no venían, las creo como NaN para no romper el plot
    if "position" not in average_locs_and_count_df.columns:
        average_locs_and_count_df["position"] = np.nan
    if "positionSide" not in average_locs_and_count_df.columns:
        average_locs_and_count_df["positionSide"] = np.nan

    average_locs_and_count_df = average_locs_and_count_df.set_index("playerId")
    average_locs_and_count_df["name"] = average_locs_and_count_df["name"].apply(limpiar_nombre_safe)

    # Passes between (min/max sin floats)
    passes_player_ids_df = passes_df.loc[:, ["index", "playerId", "receiver", "teamName"]].copy()
    passes_player_ids_df = passes_player_ids_df.dropna(subset=["playerId", "receiver"])
    passes_player_ids_df = passes_player_ids_df[
        (passes_player_ids_df["playerId"] != "") & (passes_player_ids_df["receiver"] != "")
    ].copy()

    passes_player_ids_df["pos_max"] = passes_player_ids_df[["playerId", "receiver"]].max(axis=1)
    passes_player_ids_df["pos_min"] = passes_player_ids_df[["playerId", "receiver"]].min(axis=1)

    passes_between_df = passes_player_ids_df.groupby(["pos_min", "pos_max"]).index.count().reset_index()
    passes_between_df.rename({"index": "pass_count"}, axis="columns", inplace=True)

    passes_between_df = passes_between_df.merge(average_locs_and_count_df, left_on="pos_min", right_index=True)
    passes_between_df = passes_between_df.merge(
        average_locs_and_count_df, left_on="pos_max", right_index=True, suffixes=["", "_end"]
    )

    return passes_between_df, average_locs_and_count_df

# home_team_id = list(teams_dict.keys())[0]
home_passes_between_df, home_average_locs_and_count_df = get_passes_between_df(hteamName, passes_df, players_df)
# away_team_id = list(teams_dict.keys())[1]
away_passes_between_df, away_average_locs_and_count_df = get_passes_between_df(ateamName, passes_df, players_df)


def pass_network_visualization(ax, passes_between_df, average_locs_and_count_df, col, teamName, flipped=False):
    MAX_LINE_WIDTH = 15
    MAX_MARKER_SIZE = 3000

    passes_between_df['width'] = (passes_between_df.pass_count / passes_between_df.pass_count.max() * MAX_LINE_WIDTH)

    MIN_TRANSPARENCY = 0.05
    MAX_TRANSPARENCY = 0.85
    color = np.array(to_rgba(col))
    color = np.tile(color, (len(passes_between_df), 1))
    c_transparency = passes_between_df.pass_count / passes_between_df.pass_count.max()
    c_transparency = (c_transparency * (MAX_TRANSPARENCY - MIN_TRANSPARENCY)) + MIN_TRANSPARENCY
    color[:, 3] = c_transparency

    pitch = Pitch(pitch_type='uefa', corner_arcs=True, pitch_color=bg_color, line_color=line_color, linewidth=2)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)

    # lines
    pitch.lines(
        passes_between_df.pass_avg_x, passes_between_df.pass_avg_y,
        passes_between_df.pass_avg_x_end, passes_between_df.pass_avg_y_end,
        lw=passes_between_df.width, color=color, zorder=1, ax=ax
    )

    # nodes
    for _, row in average_locs_and_count_df.iterrows():
        if row.get('isFirstEleven', False) == True:
            pitch.scatter(row['pass_avg_x'], row['pass_avg_y'], s=1000, marker='o',
                          color=bg_color, edgecolor=line_color, linewidth=2, alpha=1, ax=ax)
        else:
            pitch.scatter(row['pass_avg_x'], row['pass_avg_y'], s=1000, marker='s',
                          color=bg_color, edgecolor=line_color, linewidth=2, alpha=0.75, ax=ax)

    # shirtNo
    for _, row in average_locs_and_count_df.iterrows():
        pitch.annotate(row["shirtNo"], xy=(row.pass_avg_x, row.pass_avg_y),
                       c=col, ha='center', va='center', size=18, ax=ax)

    # median height
    avgph = round(average_locs_and_count_df['pass_avg_x'].median(), 2)
    ax.axvline(x=avgph, color='gray', linestyle='--', alpha=0.75, linewidth=2)

    # --- Defense line Height (TU DEFINICIÓN: Defender + positionSide contiene Centre) ---
    center_backs_height = average_locs_and_count_df[
        (average_locs_and_count_df["position"] == "Defender") &
        (average_locs_and_count_df["positionSide"].astype(str).str.contains("Centre", na=False))
    ]

    def_line_h = round(center_backs_height["pass_avg_x"].median(), 2)

    # fallback para que nunca quede NaN
    if center_backs_height.empty or center_backs_height["pass_avg_x"].dropna().empty:
        def_line_h = round(average_locs_and_count_df["pass_avg_x"].median(), 2)

    ax.axvline(x=def_line_h, color='gray', linestyle='dotted', alpha=0.5, linewidth=2)

    # --- Forward line Height (top 2 más adelantados entre titulares) ---
    Forwards_height = average_locs_and_count_df[average_locs_and_count_df['isFirstEleven'] == True].copy()
    Forwards_height = Forwards_height.sort_values(by='pass_avg_x', ascending=False).head(2)
    fwd_line_h = round(Forwards_height['pass_avg_x'].mean(), 2)
    ax.axvline(x=fwd_line_h, color='gray', linestyle='dotted', alpha=0.5, linewidth=2)

    # middle zone fill
    ymid = [0, 0, 68, 68]
    xmid = [def_line_h, fwd_line_h, fwd_line_h, def_line_h]
    ax.fill(xmid, ymid, col, alpha=0.1)

    # verticality
    team_passes_df = passes_df[(passes_df["teamName"] == teamName)].copy()
    team_passes_df['pass_or_carry_angle'] = team_passes_df['pass_or_carry_angle'].abs()
    team_passes_df = team_passes_df[(team_passes_df['pass_or_carry_angle'] >= 0) & (team_passes_df['pass_or_carry_angle'] <= 90)]
    med_ang = team_passes_df['pass_or_carry_angle'].median()
    verticality = round((1 - med_ang/90) * 100, 2)

    # top combination
    top_combo = passes_between_df.sort_values(by='pass_count', ascending=False).head(1).reset_index(drop=True)
    most_pass_from = top_combo['name'][0]
    most_pass_to = top_combo['name_end'][0]
    most_pass_count = top_combo['pass_count'][0]

    # headings
    if teamName == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()
        ax.text(avgph - 1, 73, f"{avgph}m", fontsize=15, color=line_color, ha='left')
        ax.text(105, 73, f"Verticalidad: {verticality}%", fontsize=15, color=line_color, ha='left')
    else:
        ax.text(avgph - 1, -5, f"{avgph}m", fontsize=15, color=line_color, ha='right')
        ax.text(105, -5, f"Verticalidad: {verticality}%", fontsize=15, color=line_color, ha='right')

    if teamName == hteamName:
        ax.text(2, 66, "Círculo = Titular\nCaja = Suplente", color=hcol, size=12, ha='left', va='top')
        ax.set_title(f"{hteamName}\nMapa de pases", color=line_color, size=25, fontweight='bold')
    else:
        ax.text(2, 2, "Círculo = Titular\nCaja = Suplente", color=acol, size=12, ha='right', va='top')
        ax.set_title(f"{ateamName}\nMapa de pases", color=line_color, size=25, fontweight='bold')

    return {
        'Team_Name': teamName,
        'Defense_Line_Height': def_line_h,
        'Vericality_%': verticality,
        'Most_pass_combination_from': most_pass_from,
        'Most_pass_combination_to': most_pass_to,
        'Most_passes_in_combination': most_pass_count,
    }

''',
    ),
    (
        40,
        r'''
fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
pass_network_stats_home = pass_network_visualization(axs[0], home_passes_between_df, home_average_locs_and_count_df, hcol, hteamName)
pass_network_stats_away = pass_network_visualization(axs[1], away_passes_between_df, away_average_locs_and_count_df, acol, ateamName)
pass_network_stats_list = []
pass_network_stats_list.append(pass_network_stats_home)
pass_network_stats_list.append(pass_network_stats_away)
pass_network_stats_df = pd.DataFrame(pass_network_stats_list)

pass_network_stats_df.head()
''',
    ),
    (
        41,
        r'''
# =========================
# DEFENSIVE BLOCK (FIXED)
# - playerId alfanumérico => string (merge safe)
# - isFirstEleven True/NaN (no 1/0)
# - centrales: position == "Defender" AND positionSide contiene "Centre"
# - trae positionSide al DF de promedios (y normaliza sufijos _x/_y)
# - fallbacks para evitar NaN en compactness
# - qualifiers puede venir NaN (str.contains safe)
# =========================

def get_defensive_action_df(df):
    df_ = df.copy()

    # qualifiers safe
    if "qualifiers" in df_.columns:
        q = df_["qualifiers"].astype("string")
    else:
        q = pd.Series([""] * len(df_), index=df_.index, dtype="string")

    defensive_actions_ids = df_.index[
        ((df_['type'] == 'Aerial') & (q.str.contains('Defensive', na=False))) |
        (df_['type'] == 'BallRecovery') |
        (df_['type'] == 'BlockedPass') |
        (df_['type'] == 'Challenge') |
        (df_['type'] == 'Clearance') |
        (df_['type'] == 'Error') |
        (df_['type'] == 'Foul') |
        (df_['type'] == 'Interception') |
        (df_['type'] == 'Tackle')
    ]

    keep_cols = ["index", "x", "y", "teamName", "playerId", "type", "outcomeType"]
    keep_cols = [c for c in keep_cols if c in df_.columns]

    return df_.loc[defensive_actions_ids, keep_cols].copy()


def get_da_count_df(team_name, defensive_actions_df, players_df):
    da = defensive_actions_df[defensive_actions_df["teamName"] == team_name].copy()

    # --- playerId alfanumérico -> string en ambos ---
    if "playerId" in da.columns:
        da["playerId"] = da["playerId"].astype("string")
    players_df_ = players_df.copy()
    players_df_["playerId"] = players_df_["playerId"].astype("string")

    # merge titulares
    if "isFirstEleven" in players_df_.columns and "playerId" in da.columns:
        da = da.merge(players_df_[["playerId", "isFirstEleven"]], on="playerId", how="left")
    else:
        da["isFirstEleven"] = np.nan

    # median positions + count (solo con acciones defensivas)
    avg = da.groupby("playerId").agg({"x": ["median"], "y": ["median", "count"]})
    avg.columns = ["x", "y", "count"]
    avg = avg.reset_index()

    # ---- Merge info jugadores: name/shirtNo + position + positionSide + isFirstEleven ----
    pos_candidates  = ["position", "Position", "position_x", "position_y"]
    side_candidates = ["positionSide", "positionSide_x", "positionSide_y"]

    pos_col  = next((c for c in pos_candidates if c in players_df_.columns), None)
    side_col = next((c for c in side_candidates if c in players_df_.columns), None)

    base_need = ["playerId", "name", "shirtNo", "isFirstEleven"]
    need = base_need.copy()
    if pos_col is not None:
        need.append(pos_col)
    if side_col is not None:
        need.append(side_col)

    avg = avg.merge(players_df_[need], on="playerId", how="left")

    # normalizo nombres: SIEMPRE position / positionSide
    if pos_col is not None and pos_col != "position":
        avg = avg.rename(columns={pos_col: "position"})
    if side_col is not None and side_col != "positionSide":
        avg = avg.rename(columns={side_col: "positionSide"})

    # si faltan, creo para no romper
    if "position" not in avg.columns:
        avg["position"] = np.nan
    if "positionSide" not in avg.columns:
        avg["positionSide"] = np.nan

    avg = avg.set_index("playerId")

    return avg


def defensive_block(ax, average_locs_and_count_df, team_name, col, defensive_actions_df):
    defensive_actions_team_df = defensive_actions_df[defensive_actions_df["teamName"] == team_name].copy()

    pitch = Pitch(
        pitch_type='uefa',
        pitch_color=bg_color,
        line_color=line_color,
        linewidth=2,
        line_zorder=2,
        corner_arcs=True
    )
    pitch.draw(ax=ax)
    ax.set_facecolor(bg_color)
    ax.set_xlim(-0.5, 105.5)

    # marker size
    MAX_MARKER_SIZE = 3500
    df_plot = average_locs_and_count_df.copy()
    if "count" in df_plot.columns and df_plot["count"].max() not in [0, np.nan]:
        df_plot['marker_size'] = (df_plot['count'] / df_plot['count'].max() * MAX_MARKER_SIZE)
    else:
        df_plot['marker_size'] = 500

    # heatmap (kde)
    color_rgba = np.array(to_rgba(col))
    flamingo_cmap = LinearSegmentedColormap.from_list("Flamingo - 100 colors", [bg_color, col], N=500)
    pitch.kdeplot(
        defensive_actions_team_df.x, defensive_actions_team_df.y,
        ax=ax, fill=True, levels=5000, thresh=0.02, cut=4, cmap=flamingo_cmap
    )

    # nodes starter/sub
    df_plot = df_plot.reset_index(drop=True)
    for _, row in df_plot.iterrows():
        if row.get('isFirstEleven', False) == True:
            pitch.scatter(row['x'], row['y'], s=row['marker_size'] + 100, marker='o',
                          color=bg_color, edgecolor=line_color, linewidth=1, alpha=1, zorder=3, ax=ax)
        else:
            pitch.scatter(row['x'], row['y'], s=row['marker_size'] + 100, marker='s',
                          color=bg_color, edgecolor=line_color, linewidth=1, alpha=1, zorder=3, ax=ax)

    # acciones (puntos)
    pitch.scatter(defensive_actions_team_df.x, defensive_actions_team_df.y,
                  s=10, marker='x', color='yellow', alpha=0.2, ax=ax)

    # shirt number
    for _, row in df_plot.iterrows():
        pitch.annotate(row["shirtNo"], xy=(row.x, row.y),
                       c=line_color, ha='center', va='center', size=14, ax=ax)

    # Defensive Actions Height (media de x en jugadores con acciones)
    dah = round(df_plot['x'].mean(), 2)
    dah_show = round((dah * 1.05), 2)
    ax.axvline(x=dah, color='gray', linestyle='--', alpha=0.75, linewidth=2)

    # ---- Defense line height (TU CRITERIO centrales) ----
    center_backs_height = df_plot[
        (df_plot["position"] == "Defender") &
        (df_plot["positionSide"].astype(str).str.contains("Centre", na=False))
    ]

    def_line_h = round(center_backs_height["x"].median(), 2)

    # fallback (si no encontró centrales)
    if center_backs_height.empty or center_backs_height["x"].dropna().empty:
        def_line_h = round(df_plot["x"].median(), 2)

    ax.axvline(x=def_line_h, color='gray', linestyle='dotted', alpha=0.5, linewidth=2)

    # ---- Forward line pressing height (top 2 más adelantados entre TITULARES) ----
    forwards = df_plot[df_plot["isFirstEleven"] == True].copy()
    forwards = forwards.sort_values(by="x", ascending=False).head(2)

    fwd_line_h = round(forwards["x"].mean(), 2)

    # fallback si no hay titulares en este df (raro, pero posible)
    if forwards.empty or forwards["x"].dropna().empty:
        fwd_line_h = round(df_plot["x"].max(), 2)

    ax.axvline(x=fwd_line_h, color='gray', linestyle='dotted', alpha=0.5, linewidth=2)

    # Compactness (evitar NaN)
    compactness = round((1 - ((fwd_line_h - def_line_h) / 105)) * 100, 2)

    # textos por local/visita
    if team_name == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()
        ax.text(dah - 1, 73, f"{dah_show}m", fontsize=15, color=line_color, ha='left', va='center')
        ax.text(105, 73, f'Compacto:{compactness}%', fontsize=15, color=line_color, ha='left', va='center')
        ax.text(2, 2, "Círculo = Titular\nCaja = Suplente", color='gray', size=12, ha='right', va='top')
        ax.set_title(f"{ateamName}\nBloque defensivo ", color=line_color, fontsize=25, fontweight='bold')
    else:
        ax.text(dah - 1, -5, f"{dah_show}m", fontsize=15, color=line_color, ha='right', va='center')
        ax.text(105, -5, f'Compacto:{compactness}%', fontsize=15, color=line_color, ha='right', va='center')
        ax.text(2, 66, "Círculo = Titular\nCaja = Suplente", color='gray', size=12, ha='left', va='top')
        ax.set_title(f"{hteamName}\nBloque defensivo", color=line_color, fontsize=25, fontweight='bold')

    return {
        'Team_Name': team_name,
        'Average_Defensive_Action_Height': dah,
        'Defense_Line_Height': def_line_h,
        'Forward_Line_Pressing_Height': fwd_line_h,
        'Compact_%': compactness
    }


# =========================
# RUN
# =========================
defensive_actions_df = get_defensive_action_df(df)

defensive_home_average_locs_and_count_df = get_da_count_df(hteamName, defensive_actions_df, players_df)
defensive_away_average_locs_and_count_df = get_da_count_df(ateamName, defensive_actions_df, players_df)

# sacar GK (si tu 'position' vale "Goalkeeper", ajustá)
defensive_home_average_locs_and_count_df = defensive_home_average_locs_and_count_df[
    defensive_home_average_locs_and_count_df['position'] != 'GK'
]
defensive_away_average_locs_and_count_df = defensive_away_average_locs_and_count_df[
    defensive_away_average_locs_and_count_df['position'] != 'GK'
]

fig, axs = plt.subplots(1, 2, figsize=(20, 10), facecolor=bg_color)

defensive_block_stats_home = defensive_block(
    axs[0], defensive_home_average_locs_and_count_df, hteamName, hcol, defensive_actions_df
)
defensive_block_stats_away = defensive_block(
    axs[1], defensive_away_average_locs_and_count_df, ateamName, acol, defensive_actions_df
)

defensive_block_stats_df = pd.DataFrame([defensive_block_stats_home, defensive_block_stats_away])
''',
    ),
    (
        42,
        r'''
def draw_progressive_pass_map(ax, team_name, col):
    # filtering those passes which has reduced the distance form goal for at least 10yds and not started from defensive third, this is my condition for a progressive pass, which almost similar to opta/statsbomb conditon
    dfpro = df[(df['teamName']==team_name) & (df['prog_pass']>=9.11) & (~df['qualifiers'].str.contains('CornerTaken|Freekick')) & 
               (df['x']>=35) & (df['outcomeType']=='Successful')]
    pitch = Pitch(pitch_type='uefa', pitch_color=bg_color, line_color=line_color, linewidth=2, corner_arcs=True)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    # ax.set_ylim(-0.5, 68.5)

    if team_name == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()

    pro_count = len(dfpro)

    # calculating the counts
    left_pro = len(dfpro[dfpro['y']>=45.33])
    mid_pro = len(dfpro[(dfpro['y']>=22.67) & (dfpro['y']<45.33)])
    right_pro = len(dfpro[(dfpro['y']>=0) & (dfpro['y']<22.67)])
    left_percentage = round((left_pro/pro_count)*100)
    mid_percentage = round((mid_pro/pro_count)*100)
    right_percentage = round((right_pro/pro_count)*100)

    ax.hlines(22.67, xmin=0, xmax=105, colors=line_color, linestyle='dashed', alpha=0.35)
    ax.hlines(45.33, xmin=0, xmax=105, colors=line_color, linestyle='dashed', alpha=0.35)

    # showing the texts in the pitch
    bbox_props = dict(boxstyle="round,pad=0.3", edgecolor="None", facecolor=bg_color, alpha=0.75)
    if col == hcol:
        ax.text(8, 11.335, f'{right_pro}\n({right_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_pro}\n({mid_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_pro}\n({left_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
    else:
        ax.text(8, 11.335, f'{right_pro}\n({right_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_pro}\n({mid_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_pro}\n({left_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)

    # plotting the passes
    pro_pass = pitch.lines(dfpro.x, dfpro.y, dfpro.endX, dfpro.endY, lw=3.5, comet=True, color=col, ax=ax, alpha=0.5)
    # plotting some scatters at the end of each pass
    pro_pass_end = pitch.scatter(dfpro.endX, dfpro.endY, s=35, edgecolor=col, linewidth=1, color=bg_color, zorder=2, ax=ax)

    counttext = f"{pro_count} Passes progresivos"

    # Heading and other texts
    if col == hcol:
        ax.set_title(f"{hteamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold')
    else:
        ax.set_title(f"{ateamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold')

    return {
        'Team_Name': team_name,
        'Total_Progressive_Passes': pro_count,
        'Progressive_Passes_From_Left': left_pro,
        'Progressive_Passes_From_Center': mid_pro,
        'Progressive_Passes_From_Right': right_pro
    }

fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
Progressvie_Passes_Stats_home = draw_progressive_pass_map(axs[0], hteamName, hcol)
Progressvie_Passes_Stats_away = draw_progressive_pass_map(axs[1], ateamName, acol)
Progressvie_Passes_Stats_list = []
Progressvie_Passes_Stats_list.append(Progressvie_Passes_Stats_home)
Progressvie_Passes_Stats_list.append(Progressvie_Passes_Stats_away)
Progressvie_Passes_Stats_df = pd.DataFrame(Progressvie_Passes_Stats_list)
''',
    ),
    (
        43,
        r'''
def draw_progressive_carry_map(ax, team_name, col):
    # filtering those carries which has reduced the distance form goal for at least 10yds and not ended at defensive third, this is my condition for a progressive pass, which almost similar to opta/statsbomb conditon
    dfpro = df[(df['teamName']==team_name) & (df['prog_carry']>=9.11) & (df['endX']>=35)]
    pitch = Pitch(pitch_type='uefa', pitch_color=bg_color, line_color=line_color, linewidth=2,
                          corner_arcs=True)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    # ax.set_ylim(-5, 68.5)

    if team_name == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()

    pro_count = len(dfpro)

    # calculating the counts
    left_pro = len(dfpro[dfpro['y']>=45.33])
    mid_pro = len(dfpro[(dfpro['y']>=22.67) & (dfpro['y']<45.33)])
    right_pro = len(dfpro[(dfpro['y']>=0) & (dfpro['y']<22.67)])
    left_percentage = round((left_pro/pro_count)*100)
    mid_percentage = round((mid_pro/pro_count)*100)
    right_percentage = round((right_pro/pro_count)*100)

    ax.hlines(22.67, xmin=0, xmax=105, colors=line_color, linestyle='dashed', alpha=0.35)
    ax.hlines(45.33, xmin=0, xmax=105, colors=line_color, linestyle='dashed', alpha=0.35)

    # showing the texts in the pitch
    bbox_props = dict(boxstyle="round,pad=0.3", edgecolor="None", facecolor=bg_color, alpha=0.75)
    if col == hcol:
        ax.text(8, 11.335, f'{right_pro}\n({right_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_pro}\n({mid_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_pro}\n({left_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
    else:
        ax.text(8, 11.335, f'{right_pro}\n({right_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_pro}\n({mid_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_pro}\n({left_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)

    # plotting the carries
    for index, row in dfpro.iterrows():
        arrow = patches.FancyArrowPatch((row['x'], row['y']), (row['endX'], row['endY']), arrowstyle='->', color=col, zorder=4, mutation_scale=20, 
                                        alpha=0.9, linewidth=2, linestyle='--')
        ax.add_patch(arrow)

    counttext = f"{pro_count} Conducciones progresivas"

    # Heading and other texts
    if col == hcol:
        ax.set_title(f"{hteamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold')
    else:
        ax.set_title(f"{ateamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold')

    return {
        'Team_Name': team_name,
        'Total_Progressive_Carries': pro_count,
        'Progressive_Carries_From_Left': left_pro,
        'Progressive_Carries_From_Center': mid_pro,
        'Progressive_Carries_From_Right': right_pro
    }

fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
Progressvie_Carries_Stats_home = draw_progressive_carry_map(axs[0], hteamName, hcol)
Progressvie_Carries_Stats_away = draw_progressive_carry_map(axs[1], ateamName, acol)
Progressvie_Carries_Stats_list = []
Progressvie_Carries_Stats_list.append(Progressvie_Carries_Stats_home)
Progressvie_Carries_Stats_list.append(Progressvie_Carries_Stats_away)
Progressvie_Carries_Stats_df = pd.DataFrame(Progressvie_Carries_Stats_list)
''',
    ),
    (
        44,
        r'''
import pandas as pd
import numpy as np
from unidecode import unidecode


# ============================================================
# ADAPTADOR 365SCORES -> shots_df estándar
# ============================================================

def procesar_shots_365scores(
    shots365,
    players365,
    hteamName,
    ateamName,
    competitor_num_home=1,
    competitor_num_away=2
):
    """
    Convierte shots365 + players365 al formato estándar que veníamos usando.

    IMPORTANTE:
    - En shots365, la columna side es coordenada numérica del remate.
    - NO la pisamos con 'home'/'away'.
    - Creamos teamSide para home/away.
    """

    shots = shots365.copy()
    players = players365.copy()

    # ------------------------------------------------------------
    # 1. Preservar columnas originales de coordenadas 365
    # ------------------------------------------------------------

    shots["shot_line_365"] = pd.to_numeric(shots["line"], errors="coerce")
    shots["shot_side_365"] = pd.to_numeric(shots["side"], errors="coerce")

    # 365 trae x vacío en tu CSV, entonces reconstruimos cancha:
    # side = profundidad 0-100 -> x 0-105
    # line = lateral 0-100 -> y 0-68
    shots["x"] = shots["shot_side_365"] * 1.05
    shots["y_pitch"] = shots["shot_line_365"] * 0.68

    # OJO:
    # en el estándar del resto del código la coordenada de cancha se llama y.
    # Pero en shots365 ya existe y y esa y es goalMouth_y.
    # Entonces primero guardamos boca de arco y después usamos y_pitch como y.
    shots["goalMouth_y"] = pd.to_numeric(shots["y"], errors="coerce")
    shots["goalMouth_z"] = pd.to_numeric(shots["z"], errors="coerce")

    shots["y"] = shots["y_pitch"]

    # Alias para bloques que esperan estos nombres
    shots["goalMouthY"] = shots["goalMouth_y"]
    shots["goalMouthZ"] = shots["goalMouth_z"]

    # ------------------------------------------------------------
    # 2. Merge con jugadores
    # ------------------------------------------------------------

    players_small = players.rename(columns={
        "id": "playerId_365",
        "name": "playerName_365",
        "shortName": "shortName_365",
        "jerseyNumber": "jerseyNumber_365",
        "competitorId": "competitorId_365"
    }).copy()

    shots = shots.merge(
        players_small[
            [
                "playerId_365",
                "playerName_365",
                "shortName_365",
                "jerseyNumber_365",
                "competitorId_365"
            ]
        ],
        left_on="playerId",
        right_on="playerId_365",
        how="left"
    )

    # ------------------------------------------------------------
    # 3. Mapear equipo
    # ------------------------------------------------------------

    team_map = {
        competitor_num_home: hteamName,
        competitor_num_away: ateamName
    }

    shots["teamName"] = shots["competitorNum"].map(team_map)

    # No usar "side" porque ya existe como coordenada 365
    shots["teamSide"] = np.where(
        shots["competitorNum"] == competitor_num_home,
        "home",
        np.where(shots["competitorNum"] == competitor_num_away, "away", None)
    )

    shots["isHome"] = shots["competitorNum"] == competitor_num_home

    shots["oppositeTeam"] = shots["teamName"].apply(
        lambda x: ateamName if x == hteamName else hteamName if x == ateamName else None
    )

    # ------------------------------------------------------------
    # 4. Jugadores
    # ------------------------------------------------------------

    shots["playerName"] = shots["playerName_365"].fillna(
        shots["playerId"].astype(str)
    )

    shots["playerName"] = shots["playerName"].apply(limpiar_nombre_safe)
    shots["jerseyNumber"] = shots["jerseyNumber_365"]

    # ------------------------------------------------------------
    # 5. Tipos de remate
    # ------------------------------------------------------------
    # En shots365:
    # shot_outcome:
    # - Gol
    # - Atajado
    # - Fallado
    # - Bloqueado
    # ------------------------------------------------------------

    outcome_map = {
        "Gol": "goal",
        "Atajado": "save",
        "Fallado": "miss",
        "Bloqueado": "block"
    }

    opta_type_map = {
        "goal": "Goal",
        "save": "SavedShot",
        "miss": "MissedShots",
        "block": "BlockedShot"
    }

    shots["shotType"] = shots["shot_outcome"].map(outcome_map).fillna(
        shots["shot_outcome"].astype(str).str.lower()
    )

    shots["shotType"] = shots["shotType"].astype(str).str.strip().str.lower()

    # No usamos la columna type original de 365 porque es numérica.
    # Creamos type para compatibilidad con algunos bloques.
    shots["type_365_original"] = shots365["type"].values
    shots["type"] = shots["shotType"]

    # Y este sirve para el shotmap viejo tipo Opta
    shots["type_opta"] = shots["shotType"].map(opta_type_map)

    # ------------------------------------------------------------
    # 6. xG / xGOT
    # ------------------------------------------------------------

    shots["expectedGoals"] = pd.to_numeric(shots["xg"], errors="coerce").fillna(0)
    shots["expectedGoalsOnTarget"] = pd.to_numeric(shots["xgot"], errors="coerce").fillna(0)

    # Mantener alias
    shots["xg"] = shots["expectedGoals"]
    shots["xgot"] = shots["expectedGoalsOnTarget"]

    # ------------------------------------------------------------
    # 7. Tiempo
    # ------------------------------------------------------------

    shots["time"] = shots["time"].astype(str)

    shots["minute"] = pd.to_numeric(
        shots["time"].str.extract(r"(\d+)")[0],
        errors="coerce"
    )

    # ------------------------------------------------------------
    # 8. Flags
    # ------------------------------------------------------------

    shots["isGoal"] = shots["shotType"].eq("goal")
    shots["isOwnGoal"] = False

    # Si 365 no trae big chance, queda False.
    shots["isBigChance"] = False

    shots["goalMouthLocation"] = shots.get("goalDescription", None)
    shots["incidentType"] = "shot"

    # ------------------------------------------------------------
    # 9. Validaciones
    # ------------------------------------------------------------

    print("✅ shots_df adaptado desde 365Scores")
    print("Equipos:", shots["teamName"].dropna().unique())
    print("Tipos:", shots["shotType"].value_counts(dropna=False).to_dict())

    print("\nChequeo coordenadas:")
    display(
        shots[
            [
                "playerName",
                "teamName",
                "shotType",
                "shot_side_365",
                "shot_line_365",
                "x",
                "y",
                "goalMouth_y",
                "goalMouth_z",
                "expectedGoals",
                "expectedGoalsOnTarget"
            ]
        ].head(10)
    )

    display(
        shots.groupby(["teamName", "shotType"])
        .agg(
            shots=("shotType", "count"),
            xG=("expectedGoals", "sum"),
            xGOT=("expectedGoalsOnTarget", "sum")
        )
        .reset_index()
    )

    return shots


# ============================================================
# EJECUCIÓN
# ============================================================

shots_df = procesar_shots_365scores(
    shots365=shots365,
    players365=players365,
    hteamName=hteamName,
    ateamName=ateamName,
    competitor_num_home=1,
    competitor_num_away=2
)
''',
    ),
    (
        45,
        r'''
# ============================================================
# PREPARAR SHOTS_DF PARA EL BLOQUE POSTERIOR
# ============================================================

shots_df = shots_df.copy()

# ------------------------------------------------------------
# 1. Asegurar nombres de xG / xGOT
# ------------------------------------------------------------

if "expectedGoals" not in shots_df.columns and "xg" in shots_df.columns:
    shots_df["expectedGoals"] = shots_df["xg"]

if "expectedGoalsOnTarget" not in shots_df.columns and "xgot" in shots_df.columns:
    shots_df["expectedGoalsOnTarget"] = shots_df["xgot"]

# ------------------------------------------------------------
# 2. Asegurar columnas mínimas
# ------------------------------------------------------------

required_shots_cols = [
    "teamName",
    "playerName",
    "expectedGoals",
    "expectedGoalsOnTarget",
    "shotType",
    "x",
    "y"
]

missing_cols = [col for col in required_shots_cols if col not in shots_df.columns]

if missing_cols:
    raise ValueError(f"Faltan columnas en shots_df: {missing_cols}")

# ------------------------------------------------------------
# 3. Numéricos
# ------------------------------------------------------------

for col in [
    "expectedGoals",
    "expectedGoalsOnTarget",
    "x",
    "y",
    "goalMouth_y",
    "goalMouth_z",
    "goalMouthY",
    "goalMouthZ"
]:
    if col in shots_df.columns:
        shots_df[col] = pd.to_numeric(shots_df[col], errors="coerce")

shots_df["expectedGoals"] = shots_df["expectedGoals"].fillna(0)
shots_df["expectedGoalsOnTarget"] = shots_df["expectedGoalsOnTarget"].fillna(0)

# ------------------------------------------------------------
# 4. Limpiar nombres de jugadores
# ------------------------------------------------------------

shots_df["playerName"] = shots_df["playerName"].apply(limpiar_nombre_safe)

# ------------------------------------------------------------
# 5. Normalizar tipos
# ------------------------------------------------------------

shots_df["shotType"] = shots_df["shotType"].astype(str).str.strip().str.lower()

if "type" not in shots_df.columns:
    shots_df["type"] = shots_df["shotType"]
else:
    shots_df["type"] = shots_df["type"].astype(str).str.strip().str.lower()

# ------------------------------------------------------------
# 6. Agregar equipo rival
# ------------------------------------------------------------

def get_opposite_teamName(team):
    if team == hteamName:
        return ateamName
    elif team == ateamName:
        return hteamName
    else:
        return None

shots_df["oppositeTeam"] = shots_df["teamName"].apply(get_opposite_teamName)

# ------------------------------------------------------------
# 7. Chequeo final
# ------------------------------------------------------------

print("✅ shots_df preparado para bloques posteriores")

display(
    shots_df.groupby(["teamName", "shotType"])
    .agg(
        shots=("shotType", "count"),
        xG=("expectedGoals", "sum"),
        xGOT=("expectedGoalsOnTarget", "sum")
    )
    .reset_index()
)
''',
    ),
    (
        46,
        r'''
# ============================================================
# PREPARAR SHOTS_DF PARA EL BLOQUE POSTERIOR
# Compatible con 365Scores / evita columnas duplicadas
# ============================================================

shots_df = shots_df.copy()

# ------------------------------------------------------------
# 1. Eliminar columnas duplicadas si ya existen
# ------------------------------------------------------------
# Si por renombres anteriores quedaron dos columnas con el mismo nombre,
# pandas devuelve un DataFrame al hacer shots_df["expectedGoals"].
# Esto rompe pd.to_numeric().
# ------------------------------------------------------------

shots_df = shots_df.loc[:, ~shots_df.columns.duplicated()].copy()

# ------------------------------------------------------------
# 2. Asegurar expectedGoals / expectedGoalsOnTarget sin renombrar a ciegas
# ------------------------------------------------------------

if "expectedGoals" not in shots_df.columns:
    if "xg" in shots_df.columns:
        shots_df["expectedGoals"] = shots_df["xg"]
    else:
        shots_df["expectedGoals"] = 0

if "expectedGoalsOnTarget" not in shots_df.columns:
    if "xgot" in shots_df.columns:
        shots_df["expectedGoalsOnTarget"] = shots_df["xgot"]
    else:
        shots_df["expectedGoalsOnTarget"] = 0

# ------------------------------------------------------------
# 3. Asegurar columnas mínimas
# ------------------------------------------------------------

required_shots_cols = [
    "teamName",
    "playerName",
    "expectedGoals",
    "expectedGoalsOnTarget"
]

missing_cols = [col for col in required_shots_cols if col not in shots_df.columns]

if missing_cols:
    raise ValueError(f"Faltan columnas en shots_df: {missing_cols}")

# ------------------------------------------------------------
# 4. Asegurar numéricos
# ------------------------------------------------------------

shots_df["expectedGoals"] = pd.to_numeric(
    shots_df["expectedGoals"],
    errors="coerce"
).fillna(0)

shots_df["expectedGoalsOnTarget"] = pd.to_numeric(
    shots_df["expectedGoalsOnTarget"],
    errors="coerce"
).fillna(0)

# Mantener alias xg / xgot por compatibilidad
shots_df["xg"] = shots_df["expectedGoals"]
shots_df["xgot"] = shots_df["expectedGoalsOnTarget"]

# ------------------------------------------------------------
# 5. Limpiar nombres de jugadores
# ------------------------------------------------------------

shots_df["playerName"] = shots_df["playerName"].apply(limpiar_nombre_safe)

# ------------------------------------------------------------
# 6. Agregar equipo rival
# ------------------------------------------------------------

def get_opposite_teamName(team):
    if team == hteamName:
        return ateamName
    elif team == ateamName:
        return hteamName
    else:
        return None

shots_df["oppositeTeam"] = shots_df["teamName"].apply(get_opposite_teamName)

# ------------------------------------------------------------
# 7. Control rápido
# ------------------------------------------------------------

print("✅ shots_df preparado correctamente")
print("Columnas duplicadas:", shots_df.columns[shots_df.columns.duplicated()].tolist())

display(
    shots_df.groupby("teamName")
    .agg(
        Shots=("teamName", "count"),
        xG=("expectedGoals", "sum"),
        xGOT=("expectedGoalsOnTarget", "sum")
    )
    .reset_index()
)
''',
    ),
    (
        47,
        r'''
# ============================================================
# COLORES
# ============================================================

hcol = col1
acol = col2


# ============================================================
# HELPERS
# ============================================================

def bool_col(dataframe, col):
    """
    Devuelve una Serie booleana.
    Si la columna no existe, devuelve False para todas las filas.
    """
    if col not in dataframe.columns:
        return pd.Series(False, index=dataframe.index)

    return dataframe[col].fillna(False).astype(bool)


# ============================================================
# EVENTOS POR EQUIPO
# ============================================================

df = df.copy()

homedf = df[df["teamName"] == hteamName].copy()
awaydf = df[df["teamName"] == ateamName].copy()

# xT seguro
if "xT" in df.columns:
    df["xT"] = pd.to_numeric(df["xT"], errors="coerce").fillna(0)
    homedf["xT"] = pd.to_numeric(homedf["xT"], errors="coerce").fillna(0)
    awaydf["xT"] = pd.to_numeric(awaydf["xT"], errors="coerce").fillna(0)

    hxT = round(homedf["xT"].sum(), 2)
    axT = round(awaydf["xT"].sum(), 2)
else:
    hxT = 0
    axT = 0


# ============================================================
# GOLES DEL PARTIDO
# ============================================================
# Los goles salen del df de eventos.
# Para goles en contra:
# - si existe isOwnGoal, usamos esa columna
# - si no existe, usamos qualifiers como fallback
# ============================================================

if "isOwnGoal" in df.columns:
    home_own_goal_mask = bool_col(homedf, "isOwnGoal")
    away_own_goal_mask = bool_col(awaydf, "isOwnGoal")
else:
    if "qualifiers" not in df.columns:
        df["qualifiers"] = ""

    df["qualifiers"] = df["qualifiers"].fillna("").astype(str)
    homedf["qualifiers"] = homedf["qualifiers"].fillna("").astype(str)
    awaydf["qualifiers"] = awaydf["qualifiers"].fillna("").astype(str)

    home_own_goal_mask = homedf["qualifiers"].str.contains("OwnGoal", na=False)
    away_own_goal_mask = awaydf["qualifiers"].str.contains("OwnGoal", na=False)


# Goles propios
hgoal_count = len(
    homedf[
        (homedf["type"] == "Goal") &
        (~home_own_goal_mask)
    ]
)

agoal_count = len(
    awaydf[
        (awaydf["type"] == "Goal") &
        (~away_own_goal_mask)
    ]
)

# Goles en contra:
# Si el visitante hace un gol en contra, suma para el local.
hgoal_count += len(
    awaydf[
        (awaydf["type"] == "Goal") &
        (away_own_goal_mask)
    ]
)

# Si el local hace un gol en contra, suma para el visitante.
agoal_count += len(
    homedf[
        (homedf["type"] == "Goal") &
        (home_own_goal_mask)
    ]
)


# ============================================================
# XG / XGOT DESDE 365SCORES
# ============================================================
# Esto ya viene del shots_df adaptado:
# - expectedGoals
# - expectedGoalsOnTarget
# ============================================================

shots_df = shots_df.copy()

# Asegurar columnas por si el origen vino con xg/xgot
if "expectedGoals" not in shots_df.columns and "xg" in shots_df.columns:
    shots_df["expectedGoals"] = shots_df["xg"]

if "expectedGoalsOnTarget" not in shots_df.columns and "xgot" in shots_df.columns:
    shots_df["expectedGoalsOnTarget"] = shots_df["xgot"]

# Si por alguna razón no existen, evitar que rompa
if "expectedGoals" not in shots_df.columns:
    shots_df["expectedGoals"] = 0

if "expectedGoalsOnTarget" not in shots_df.columns:
    shots_df["expectedGoalsOnTarget"] = 0

shots_df["expectedGoals"] = pd.to_numeric(
    shots_df["expectedGoals"],
    errors="coerce"
).fillna(0)

shots_df["expectedGoalsOnTarget"] = pd.to_numeric(
    shots_df["expectedGoalsOnTarget"],
    errors="coerce"
).fillna(0)

hshots_xgdf = shots_df[shots_df["teamName"] == hteamName].copy()
ashots_xgdf = shots_df[shots_df["teamName"] == ateamName].copy()

hxg = round(hshots_xgdf["expectedGoals"].sum(), 2)
axg = round(ashots_xgdf["expectedGoals"].sum(), 2)

hxgot = round(hshots_xgdf["expectedGoalsOnTarget"].sum(), 2)
axgot = round(ashots_xgdf["expectedGoalsOnTarget"].sum(), 2)


# ============================================================
# RESUMEN DE CONTROL
# ============================================================

print(f"{hteamName} {hgoal_count} - {agoal_count} {ateamName}")

print("\nxG")
print(f"{hteamName}: {hxg}")
print(f"{ateamName}: {axg}")

print("\nxGOT")
print(f"{hteamName}: {hxgot}")
print(f"{ateamName}: {axgot}")

print("\nxT")
print(f"{hteamName}: {hxT}")
print(f"{ateamName}: {axT}")

print("\nShots desde 365")
display(
    shots_df.groupby("teamName")
    .agg(
        Shots=("shotType", "count"),
        xG=("expectedGoals", "sum"),
        xGOT=("expectedGoalsOnTarget", "sum")
    )
    .reset_index()
)
''',
    ),
    (
        48,
        r'''
# ============================================================
# PREPARAR DATAFRAME EXCLUSIVO PARA SHOTMAP DESDE 365SCORES
# ============================================================

Shotsdf_shotmap = shots_df.copy()

# ------------------------------------------------------------
# 1. Normalizar tipos de remate 365
# ------------------------------------------------------------

outcome_map_365 = {
    "Fallado": "miss",
    "Gol": "goal",
    "Atajado": "save",
    "Bloqueado": "block"
}

if "shot_outcome" in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["shotType"] = Shotsdf_shotmap["shot_outcome"].map(outcome_map_365)
else:
    Shotsdf_shotmap["shotType"] = Shotsdf_shotmap["shotType"].astype(str).str.lower()

# Por si hubiera algo ya normalizado o alguna variante
Shotsdf_shotmap["shotType"] = (
    Shotsdf_shotmap["shotType"]
    .astype(str)
    .str.strip()
    .str.lower()
)

# ------------------------------------------------------------
# 2. Mapear al formato visual anterior
# ------------------------------------------------------------

type_visual_map = {
    "goal": "Goal",
    "save": "SavedShot",
    "miss": "MissedShots",
    "block": "BlockedShot"
}

Shotsdf_shotmap["type_365"] = Shotsdf_shotmap["shotType"]
Shotsdf_shotmap["type"] = Shotsdf_shotmap["shotType"].map(type_visual_map)

# ------------------------------------------------------------
# 3. xG / xGOT
# ------------------------------------------------------------

if "expectedGoals" not in Shotsdf_shotmap.columns and "xg" in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["expectedGoals"] = Shotsdf_shotmap["xg"]

if "expectedGoalsOnTarget" not in Shotsdf_shotmap.columns and "xgot" in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["expectedGoalsOnTarget"] = Shotsdf_shotmap["xgot"]

if "expectedGoals" not in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["expectedGoals"] = 0

if "expectedGoalsOnTarget" not in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["expectedGoalsOnTarget"] = 0

Shotsdf_shotmap["expectedGoals"] = pd.to_numeric(
    Shotsdf_shotmap["expectedGoals"],
    errors="coerce"
).fillna(0)

Shotsdf_shotmap["expectedGoalsOnTarget"] = pd.to_numeric(
    Shotsdf_shotmap["expectedGoalsOnTarget"],
    errors="coerce"
).fillna(0)

# ------------------------------------------------------------
# 4. Flags
# ------------------------------------------------------------

if "isBigChance" not in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["isBigChance"] = False

if "isOwnGoal" not in Shotsdf_shotmap.columns:
    Shotsdf_shotmap["isOwnGoal"] = False

Shotsdf_shotmap["isBigChance"] = Shotsdf_shotmap["isBigChance"].fillna(False).astype(bool)
Shotsdf_shotmap["isOwnGoal"] = Shotsdf_shotmap["isOwnGoal"].fillna(False).astype(bool)

Shotsdf_shotmap["qualifiers"] = ""

# ------------------------------------------------------------
# 5. Coordenadas
# ------------------------------------------------------------

for col in ["x", "y"]:
    if col not in Shotsdf_shotmap.columns:
        Shotsdf_shotmap[col] = np.nan

Shotsdf_shotmap["x"] = pd.to_numeric(Shotsdf_shotmap["x"], errors="coerce")
Shotsdf_shotmap["y"] = pd.to_numeric(Shotsdf_shotmap["y"], errors="coerce")

if Shotsdf_shotmap["x"].isna().all():
    if "shot_side_365" in Shotsdf_shotmap.columns:
        Shotsdf_shotmap["x"] = pd.to_numeric(Shotsdf_shotmap["shot_side_365"], errors="coerce") * 1.05
    elif "side" in Shotsdf_shotmap.columns:
        Shotsdf_shotmap["x"] = pd.to_numeric(Shotsdf_shotmap["side"], errors="coerce") * 1.05
    else:
        raise ValueError("No existe x, shot_side_365 ni side para construir x.")

if Shotsdf_shotmap["y"].isna().all():
    if "shot_line_365" in Shotsdf_shotmap.columns:
        Shotsdf_shotmap["y"] = pd.to_numeric(Shotsdf_shotmap["shot_line_365"], errors="coerce") * 0.68
    elif "line" in Shotsdf_shotmap.columns:
        Shotsdf_shotmap["y"] = pd.to_numeric(Shotsdf_shotmap["line"], errors="coerce") * 0.68
    else:
        raise ValueError("No existe y, shot_line_365 ni line para construir y.")

# ------------------------------------------------------------
# 6. Chequeo antes de filtrar
# ------------------------------------------------------------

print("========== DIAGNÓSTICO SHOTMAP 365 ==========")

print("Total shots365 original:", len(shots365))
print("Total shots_df:", len(shots_df))

print("\nshot_outcome original:")
display(Shotsdf_shotmap["shot_outcome"].value_counts(dropna=False))

print("\nshotType normalizado:")
display(Shotsdf_shotmap["shotType"].value_counts(dropna=False))

print("\ntype visual:")
display(Shotsdf_shotmap["type"].value_counts(dropna=False))

# ------------------------------------------------------------
# 7. Filtrar remates válidos
# ------------------------------------------------------------

Shotsdf_shotmap = Shotsdf_shotmap[
    Shotsdf_shotmap["type"].isin([
        "Goal",
        "SavedShot",
        "MissedShots",
        "BlockedShot"
    ])
].copy()

# Chequear remates sin coordenadas
sin_coords = Shotsdf_shotmap[
    Shotsdf_shotmap["x"].isna() |
    Shotsdf_shotmap["y"].isna()
].copy()

if len(sin_coords) > 0:
    print("⚠️ Remates con tipo válido pero sin coordenadas:")
    display(
        sin_coords[
            [
                "teamName",
                "playerName",
                "shot_outcome",
                "shotType",
                "type",
                "x",
                "y",
                "shot_side_365",
                "shot_line_365"
            ]
        ]
    )

Shotsdf_shotmap = Shotsdf_shotmap.dropna(subset=["x", "y"]).copy()
Shotsdf_shotmap.reset_index(drop=True, inplace=True)

print("\n✅ Shotsdf_shotmap final:", len(Shotsdf_shotmap))

display(
    Shotsdf_shotmap.groupby(["teamName", "shot_outcome", "shotType", "type"])
    .agg(
        cantidad=("type", "count"),
        xG=("expectedGoals", "sum"),
        xGOT=("expectedGoalsOnTarget", "sum")
    )
    .reset_index()
)
''',
    ),
    (
        49,
        r'''
# ============================================================
# VARIABLES NECESARIAS PARA plot_shotmap()
# ============================================================
# Este bloque tiene que ir ANTES de definir/ejecutar plot_shotmap(ax)
# Usa Shotsdf_shotmap ya preparado desde 365Scores.

S_stats = Shotsdf_shotmap.copy()

# Asegurar numéricos
for col in ["expectedGoals", "expectedGoalsOnTarget", "x", "y"]:
    if col in S_stats.columns:
        S_stats[col] = pd.to_numeric(S_stats[col], errors="coerce")

S_stats["expectedGoals"] = S_stats["expectedGoals"].fillna(0)
S_stats["expectedGoalsOnTarget"] = S_stats["expectedGoalsOnTarget"].fillna(0)

# Dataframes por equipo
hShotsdf = S_stats[S_stats["teamName"] == hteamName].copy()
aShotsdf = S_stats[S_stats["teamName"] == ateamName].copy()

# Tipos por equipo
hGoaldf = hShotsdf[
    (hShotsdf["type"] == "Goal") &
    (~hShotsdf["isOwnGoal"])
].copy()

aGoaldf = aShotsdf[
    (aShotsdf["type"] == "Goal") &
    (~aShotsdf["isOwnGoal"])
].copy()

hSavedf = hShotsdf[hShotsdf["type"] == "SavedShot"].copy()
aSavedf = aShotsdf[aShotsdf["type"] == "SavedShot"].copy()

hMissdf = hShotsdf[hShotsdf["type"] == "MissedShots"].copy()
aMissdf = aShotsdf[aShotsdf["type"] == "MissedShots"].copy()

hBlockdf = hShotsdf[hShotsdf["type"] == "BlockedShot"].copy()
aBlockdf = aShotsdf[aShotsdf["type"] == "BlockedShot"].copy()

hogdf = hShotsdf[hShotsdf["isOwnGoal"]].copy()
aogdf = aShotsdf[aShotsdf["isOwnGoal"]].copy()

# Totales de tiros
hTotalShots = len(hShotsdf)
aTotalShots = len(aShotsdf)

# Tiros al arco: goles + atajados
hShotsOnT = len(hGoaldf) + len(hSavedf)
aShotsOnT = len(aGoaldf) + len(aSavedf)

# xG / xGOT desde 365
hxg = round(hShotsdf["expectedGoals"].sum(), 2)
axg = round(aShotsdf["expectedGoals"].sum(), 2)

hxgot = round(hShotsdf["expectedGoalsOnTarget"].sum(), 2)
axgot = round(aShotsdf["expectedGoalsOnTarget"].sum(), 2)

# xG por tiro
hxGpSh = round(hxg / hTotalShots, 2) if hTotalShots > 0 else 0
axGpSh = round(axg / aTotalShots, 2) if aTotalShots > 0 else 0

# Distancia media del tiro
given_point = (105, 34)

home_shot_distances = np.sqrt(
    (hShotsdf["x"] - given_point[0]) ** 2 +
    (hShotsdf["y"] - given_point[1]) ** 2
)

away_shot_distances = np.sqrt(
    (aShotsdf["x"] - given_point[0]) ** 2 +
    (aShotsdf["y"] - given_point[1]) ** 2
)

home_average_shot_distance = (
    round(home_shot_distances.mean(), 2)
    if len(home_shot_distances) > 0
    else 0
)

away_average_shot_distance = (
    round(away_shot_distances.mean(), 2)
    if len(away_shot_distances) > 0
    else 0
)

print("✅ Variables para plot_shotmap definidas")
print(f"{hteamName}: tiros={hTotalShots}, al arco={hShotsOnT}, xG={hxg}, xGOT={hxgot}")
print(f"{ateamName}: tiros={aTotalShots}, al arco={aShotsOnT}, xG={axg}, xGOT={axgot}")
''',
    ),
    (
        50,
        r'''
display(
    Shotsdf_shotmap.groupby(["teamName", "shot_outcome", "shotType", "type"])
    .size()
    .reset_index(name="cantidad")
)

print("Total en Shotsdf_shotmap:", len(Shotsdf_shotmap))
print("Total local:", len(Shotsdf_shotmap[Shotsdf_shotmap["teamName"] == hteamName]))
print("Total visitante:", len(Shotsdf_shotmap[Shotsdf_shotmap["teamName"] == ateamName]))
''',
    ),
    (
        51,
        r'''
def plot_shotmap(ax):
    # ============================================================
    # USAR DF EXCLUSIVO DEL SHOTMAP 365
    # ============================================================

    S = Shotsdf_shotmap.copy()

    pitch = Pitch(
        pitch_type="uefa",
        corner_arcs=True,
        pitch_color=bg_color,
        linewidth=2,
        line_color=line_color
    )

    pitch.draw(ax=ax)
    ax.set_ylim(-0.5, 68.5)
    ax.set_xlim(-0.5, 105.5)

    # ============================================================
    # OWN GOALS
    # ============================================================

    hogdf = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "Goal") &
        (S["isOwnGoal"])
    ].copy()

    aogdf = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "Goal") &
        (S["isOwnGoal"])
    ].copy()

    # ============================================================
    # HOME SHOTS
    # ============================================================

    hGoalData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "Goal") &
        (~S["isBigChance"]) &
        (~S["isOwnGoal"])
    ].copy()

    hSaveData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "SavedShot") &
        (~S["isBigChance"])
    ].copy()

    hMissData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "MissedShots") &
        (~S["isBigChance"])
    ].copy()

    hBlockData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "BlockedShot") &
        (~S["isBigChance"])
    ].copy()

    Big_C_hGoalData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "Goal") &
        (S["isBigChance"]) &
        (~S["isOwnGoal"])
    ].copy()

    Big_C_hSaveData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "SavedShot") &
        (S["isBigChance"])
    ].copy()

    Big_C_hMissData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "MissedShots") &
        (S["isBigChance"])
    ].copy()

    Big_C_hBlockData = S[
        (S["teamName"] == hteamName) &
        (S["type"] == "BlockedShot") &
        (S["isBigChance"])
    ].copy()

    total_bigC_home = (
        len(Big_C_hGoalData) +
        len(Big_C_hSaveData) +
        len(Big_C_hMissData) +
        len(Big_C_hBlockData)
    )

    bigC_miss_home = (
        len(Big_C_hSaveData) +
        len(Big_C_hMissData) +
        len(Big_C_hBlockData)
    )

    # ============================================================
    # HOME PLOT
    # Local invertido: 105-x / 68-y
    # ============================================================

    # Atajados
    pitch.scatter(
        105 - hSaveData.x,
        68 - hSaveData.y,
        s=230,
        edgecolors=hcol,
        c="None",
        hatch="///////",
        marker="o",
        linewidths=1.5,
        zorder=4,
        ax=ax
    )

    # Fallados
    pitch.scatter(
        105 - hMissData.x,
        68 - hMissData.y,
        s=230,
        edgecolors=hcol,
        c="None",
        marker="o",
        linewidths=1.8,
        zorder=3,
        ax=ax
    )

    # Bloqueados: mismo concepto que fallados, circulito vacío
    pitch.scatter(
        105 - hBlockData.x,
        68 - hBlockData.y,
        s=230,
        edgecolors=hcol,
        c="None",
        marker="o",
        linewidths=1.8,
        zorder=3,
        ax=ax
    )

    # Goles
    pitch.scatter(
        105 - hGoalData.x,
        68 - hGoalData.y,
        s=430,
        edgecolors="green",
        linewidths=1,
        c="None",
        marker="football",
        zorder=7,
        ax=ax
    )

    # Goles en contra
    pitch.scatter(
        105 - hogdf.x,
        68 - hogdf.y,
        s=430,
        edgecolors="orange",
        linewidths=1,
        c="None",
        marker="football",
        zorder=7,
        ax=ax
    )

    # Big chances home - atajados
    pitch.scatter(
        105 - Big_C_hSaveData.x,
        68 - Big_C_hSaveData.y,
        s=650,
        edgecolors=hcol,
        c="None",
        hatch="///////",
        marker="o",
        linewidths=2,
        zorder=8,
        ax=ax
    )

    # Big chances home - fallados
    pitch.scatter(
        105 - Big_C_hMissData.x,
        68 - Big_C_hMissData.y,
        s=650,
        edgecolors=hcol,
        c="None",
        marker="o",
        linewidths=2.2,
        zorder=7,
        ax=ax
    )

    # Big chances home - bloqueados
    pitch.scatter(
        105 - Big_C_hBlockData.x,
        68 - Big_C_hBlockData.y,
        s=650,
        edgecolors=hcol,
        c="None",
        marker="o",
        linewidths=2.2,
        zorder=7,
        ax=ax
    )

    # Big chances home - goles
    pitch.scatter(
        105 - Big_C_hGoalData.x,
        68 - Big_C_hGoalData.y,
        s=850,
        edgecolors="green",
        linewidths=1,
        c="None",
        marker="football",
        zorder=10,
        ax=ax
    )

    # ============================================================
    # AWAY SHOTS
    # ============================================================

    aGoalData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "Goal") &
        (~S["isBigChance"]) &
        (~S["isOwnGoal"])
    ].copy()

    aSaveData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "SavedShot") &
        (~S["isBigChance"])
    ].copy()

    aMissData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "MissedShots") &
        (~S["isBigChance"])
    ].copy()

    aBlockData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "BlockedShot") &
        (~S["isBigChance"])
    ].copy()

    Big_C_aGoalData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "Goal") &
        (S["isBigChance"]) &
        (~S["isOwnGoal"])
    ].copy()

    Big_C_aSaveData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "SavedShot") &
        (S["isBigChance"])
    ].copy()

    Big_C_aMissData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "MissedShots") &
        (S["isBigChance"])
    ].copy()

    Big_C_aBlockData = S[
        (S["teamName"] == ateamName) &
        (S["type"] == "BlockedShot") &
        (S["isBigChance"])
    ].copy()

    total_bigC_away = (
        len(Big_C_aGoalData) +
        len(Big_C_aSaveData) +
        len(Big_C_aMissData) +
        len(Big_C_aBlockData)
    )

    bigC_miss_away = (
        len(Big_C_aSaveData) +
        len(Big_C_aMissData) +
        len(Big_C_aBlockData)
    )

    # ============================================================
    # AWAY PLOT
    # Visitante sin invertir
    # ============================================================

    # Atajados
    pitch.scatter(
        aSaveData.x,
        aSaveData.y,
        s=230,
        edgecolors=acol,
        c="None",
        hatch="///////",
        marker="o",
        linewidths=1.5,
        zorder=4,
        ax=ax
    )

    # Fallados
    pitch.scatter(
        aMissData.x,
        aMissData.y,
        s=230,
        edgecolors=acol,
        c="None",
        marker="o",
        linewidths=1.8,
        zorder=3,
        ax=ax
    )

    # Bloqueados: mismo concepto que fallados, circulito vacío
    pitch.scatter(
        aBlockData.x,
        aBlockData.y,
        s=230,
        edgecolors=acol,
        c="None",
        marker="o",
        linewidths=1.8,
        zorder=3,
        ax=ax
    )

    # Goles
    pitch.scatter(
        aGoalData.x,
        aGoalData.y,
        s=430,
        edgecolors="green",
        linewidths=1,
        c="None",
        marker="football",
        zorder=7,
        ax=ax
    )

    # Goles en contra
    pitch.scatter(
        aogdf.x,
        aogdf.y,
        s=430,
        edgecolors="orange",
        linewidths=1,
        c="None",
        marker="football",
        zorder=7,
        ax=ax
    )

    # Big chances away - atajados
    pitch.scatter(
        Big_C_aSaveData.x,
        Big_C_aSaveData.y,
        s=750,
        edgecolors=acol,
        c="None",
        hatch="///////",
        marker="o",
        linewidths=2,
        zorder=8,
        ax=ax
    )

    # Big chances away - fallados
    pitch.scatter(
        Big_C_aMissData.x,
        Big_C_aMissData.y,
        s=750,
        edgecolors=acol,
        c="None",
        marker="o",
        linewidths=2.2,
        zorder=7,
        ax=ax
    )

    # Big chances away - bloqueados
    pitch.scatter(
        Big_C_aBlockData.x,
        Big_C_aBlockData.y,
        s=750,
        edgecolors=acol,
        c="None",
        marker="o",
        linewidths=2.2,
        zorder=7,
        ax=ax
    )

    # Big chances away - goles
    pitch.scatter(
        Big_C_aGoalData.x,
        Big_C_aGoalData.y,
        s=950,
        edgecolors="green",
        linewidths=1,
        c="None",
        marker="football",
        zorder=10,
        ax=ax
    )

    # ============================================================
    # STATS BAR
    # ============================================================

    shooting_stats_title = [62 - (i * 7) for i in range(7)]

    shooting_stats_home = [
        hgoal_count,
        hxg,
        hxgot,
        hTotalShots,
        hShotsOnT,
        hxGpSh,
        home_average_shot_distance
    ]

    shooting_stats_away = [
        agoal_count,
        axg,
        axgot,
        aTotalShots,
        aShotsOnT,
        axGpSh,
        away_average_shot_distance
    ]

    def safe_pair_norm(h, a, scale=20):
        h = 0 if pd.isna(h) else float(h)
        a = 0 if pd.isna(a) else float(a)

        total = h + a

        if total == 0:
            return scale / 2, scale / 2

        return (h / total) * scale, (a / total) * scale

    shooting_stats_normalized_home = []
    shooting_stats_normalized_away = []

    for hv, av in zip(shooting_stats_home, shooting_stats_away):
        nh, na = safe_pair_norm(hv, av, scale=20)
        shooting_stats_normalized_home.append(nh)
        shooting_stats_normalized_away.append(na)

    start_x = 42.5
    start_x_for_away = [x + start_x for x in shooting_stats_normalized_home]

    ax.barh(
        shooting_stats_title,
        shooting_stats_normalized_home,
        height=5,
        color=hcol,
        left=start_x,
        zorder=2
    )

    ax.barh(
        shooting_stats_title,
        shooting_stats_normalized_away,
        height=5,
        left=start_x_for_away,
        color=acol,
        zorder=2
    )

    for sp in ["top", "right", "bottom", "left"]:
        ax.spines[sp].set_visible(False)

    ax.tick_params(
        axis="both",
        which="both",
        bottom=False,
        top=False,
        left=False,
        right=False
    )

    ax.set_xticks([])
    ax.set_yticks([])

    labels = [
        "Goles",
        "xG",
        "xGOT",
        "Remates",
        "Al arco",
        "xG/remate",
        "Dist. (m)"
    ]

    for yy, lab in zip(shooting_stats_title, labels):
        ax.text(
            52.5,
            yy,
            lab,
            color=bg_color,
            fontsize=18,
            ha="center",
            va="center",
            fontweight="bold",
            zorder=11
        )

    for yy, val in zip(shooting_stats_title, shooting_stats_home):
        ax.text(
            41.5,
            yy,
            f"{val}",
            color=line_color,
            fontsize=18,
            ha="right",
            va="center",
            fontweight="bold",
            zorder=11
        )

    for yy, val in zip(shooting_stats_title, shooting_stats_away):
        ax.text(
            63.5,
            yy,
            f"{val}",
            color=line_color,
            fontsize=18,
            ha="left",
            va="center",
            fontweight="bold",
            zorder=11
        )

    ax.text(
        0,
        70,
        f"{hteamName}\n<---Remates",
        color=hcol,
        size=25,
        ha="left",
        fontweight="bold"
    )

    ax.text(
        105,
        70,
        f"{ateamName}\nRemates--->",
        color=acol,
        size=25,
        ha="right",
        fontweight="bold"
    )

    home_data = {
        "Team_Name": hteamName,
        "Goals_Scored": hgoal_count,
        "xG": hxg,
        "xGOT": hxgot,
        "Total_Shots": hTotalShots,
        "Shots_On_Target": hShotsOnT,
        "xG_per_Shot": hxGpSh,
        "Average_Shot_Distance": home_average_shot_distance
    }

    away_data = {
        "Team_Name": ateamName,
        "Goals_Scored": agoal_count,
        "xG": axg,
        "xGOT": axgot,
        "Total_Shots": aTotalShots,
        "Shots_On_Target": aShotsOnT,
        "xG_per_Shot": axGpSh,
        "Average_Shot_Distance": away_average_shot_distance
    }

    return [home_data, away_data]
''',
    ),
    (
        52,
        r'''
fig, ax = plt.subplots(figsize=(10, 10), facecolor=bg_color)

shooting_stats = plot_shotmap(ax)

shooting_stats_df = pd.DataFrame(shooting_stats)

shooting_stats_df
''',
    ),
    (
        53,
        r'''
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from mplsoccer import Pitch

# ============================================================
# CONFIGURACIÓN / FALLBACKS
# ============================================================

try:
    bg_color
except NameError:
    bg_color = "#f2f2f2"

try:
    line_color
except NameError:
    line_color = "black"

try:
    hcol
except NameError:
    hcol = "#ff4d4d"

try:
    acol
except NameError:
    acol = "#1f9ed9"


# ============================================================
# HELPERS
# ============================================================

def bool_col(dataframe, col):
    """
    Devuelve una Serie booleana.
    Si la columna no existe, devuelve False para todas las filas.
    """
    if col not in dataframe.columns:
        return pd.Series(False, index=dataframe.index)
    
    return dataframe[col].fillna(False).astype(bool)


def normalizar_shot_outcome_365(valor):
    """
    Normaliza los tipos de remate de 365Scores.
    """
    if pd.isna(valor):
        return None

    v = str(valor).strip().lower()

    mapa = {
        "gol": "goal",
        "atajados": "save",
        "atajado": "save",
        "fallado": "miss",
        "bloqueado": "block",

        # Por si ya viniera normalizado
        "goal": "goal",
        "save": "save",
        "saved": "save",
        "miss": "miss",
        "missed": "miss",
        "block": "block",
        "blocked": "block",
    }

    return mapa.get(v, v)


def convertir_goalmouth_365_a_arco(
    df_shots,
    top_goal=False,
    clip_horizontal=False,
    clip_vertical=False
):
    """
    Convierte goalMouthY / goalMouthZ al gráfico de arco.

    En 365Scores estamos usando:
    - goalMouthY desde la columna original y de shots365.
    - goalMouthZ desde la columna original z de shots365.

    Se proyecta al arco visual:
    - ancho visual: 7.5 a 97.5
    - alto visual: 0 a 30 para arco inferior
    - alto visual: 38 a 68 para arco superior
    """

    out = df_shots.copy()

    # ============================================================
    # MEDIDAS REALES
    # ============================================================

    PITCH_WIDTH_REAL = 68.0
    GOAL_WIDTH_REAL = 7.32
    GOAL_HEIGHT_REAL = 2.44
    GOAL_CENTER_Y = 34.0

    POST_LEFT = GOAL_CENTER_Y - GOAL_WIDTH_REAL / 2
    POST_RIGHT = GOAL_CENTER_Y + GOAL_WIDTH_REAL / 2

    # En nuestra lógica previa, Z=40 representa aprox. el travesaño.
    OPTA_CROSSBAR_Z = 40.0

    # ============================================================
    # MEDIDAS VISUALES
    # ============================================================

    VISUAL_LEFT = 7.5
    VISUAL_RIGHT = 97.5
    VISUAL_WIDTH = VISUAL_RIGHT - VISUAL_LEFT
    VISUAL_HEIGHT = 30.0

    # ============================================================
    # HORIZONTAL
    # ============================================================

    out["goalMouthY_opta"] = out["goalMouthY"] * PITCH_WIDTH_REAL / 100

    out["goalMouthY_plot"] = (
        VISUAL_LEFT
        + ((POST_RIGHT - out["goalMouthY_opta"]) / GOAL_WIDTH_REAL) * VISUAL_WIDTH
    )

    # ============================================================
    # VERTICAL
    # ============================================================

    out["goalMouthZ_metros"] = (
        out["goalMouthZ"] / OPTA_CROSSBAR_Z
    ) * GOAL_HEIGHT_REAL

    out["goalMouthZ_pct_arco"] = (
        out["goalMouthZ_metros"] / GOAL_HEIGHT_REAL
    )

    out["goalMouthZ_plot"] = (
        out["goalMouthZ_pct_arco"] * VISUAL_HEIGHT
    )

    if top_goal:
        out["goalMouthZ_plot"] = 38 + out["goalMouthZ_plot"]

    # ============================================================
    # CLIPS OPCIONALES
    # ============================================================

    if clip_horizontal:
        out["goalMouthY_plot"] = out["goalMouthY_plot"].clip(
            VISUAL_LEFT,
            VISUAL_RIGHT
        )

    if clip_vertical:
        if top_goal:
            out["goalMouthZ_plot"] = out["goalMouthZ_plot"].clip(38, 68)
        else:
            out["goalMouthZ_plot"] = out["goalMouthZ_plot"].clip(0, 30)

    return out


# ============================================================
# PREPARAR SHOTSDF PARA GRÁFICO DE ARQUEROS DESDE 365SCORES
# ============================================================

Shotsdf_goalpost = shots_df.copy()

# ------------------------------------------------------------
# 1. Normalizar tipo de remate
# ------------------------------------------------------------

if "shot_outcome" in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["type"] = Shotsdf_goalpost["shot_outcome"].apply(
        normalizar_shot_outcome_365
    )
elif "shotType" in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["type"] = Shotsdf_goalpost["shotType"].apply(
        normalizar_shot_outcome_365
    )
elif "type" in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["type"] = Shotsdf_goalpost["type"].apply(
        normalizar_shot_outcome_365
    )
else:
    raise ValueError("No encuentro shot_outcome, shotType ni type en shots_df.")

Shotsdf_goalpost["type"] = (
    Shotsdf_goalpost["type"]
    .astype(str)
    .str.strip()
    .str.lower()
)

# ------------------------------------------------------------
# 2. Asegurar goalMouthY / goalMouthZ
# ------------------------------------------------------------

if "goalMouthY" not in Shotsdf_goalpost.columns:
    if "goalMouth_y" in Shotsdf_goalpost.columns:
        Shotsdf_goalpost["goalMouthY"] = Shotsdf_goalpost["goalMouth_y"]
    elif "y_original_365" in Shotsdf_goalpost.columns:
        Shotsdf_goalpost["goalMouthY"] = Shotsdf_goalpost["y_original_365"]
    else:
        raise ValueError("No encuentro goalMouthY ni goalMouth_y en shots_df.")

if "goalMouthZ" not in Shotsdf_goalpost.columns:
    if "goalMouth_z" in Shotsdf_goalpost.columns:
        Shotsdf_goalpost["goalMouthZ"] = Shotsdf_goalpost["goalMouth_z"]
    elif "z" in Shotsdf_goalpost.columns:
        Shotsdf_goalpost["goalMouthZ"] = Shotsdf_goalpost["z"]
    else:
        raise ValueError("No encuentro goalMouthZ ni goalMouth_z ni z en shots_df.")

# ------------------------------------------------------------
# 3. Asegurar xG / xGOT
# ------------------------------------------------------------

if "expectedGoals" not in Shotsdf_goalpost.columns and "xg" in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["expectedGoals"] = Shotsdf_goalpost["xg"]

if "expectedGoalsOnTarget" not in Shotsdf_goalpost.columns and "xgot" in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["expectedGoalsOnTarget"] = Shotsdf_goalpost["xgot"]

if "expectedGoals" not in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["expectedGoals"] = 0

if "expectedGoalsOnTarget" not in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["expectedGoalsOnTarget"] = 0

# ------------------------------------------------------------
# 4. Columnas opcionales
# ------------------------------------------------------------

if "isBigChance" not in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["isBigChance"] = False

if "isOwnGoal" not in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["isOwnGoal"] = False

if "playerName" not in Shotsdf_goalpost.columns:
    Shotsdf_goalpost["playerName"] = ""

# ------------------------------------------------------------
# 5. Numéricos / booleanos
# ------------------------------------------------------------

Shotsdf_goalpost["goalMouthY"] = pd.to_numeric(
    Shotsdf_goalpost["goalMouthY"],
    errors="coerce"
)

Shotsdf_goalpost["goalMouthZ"] = pd.to_numeric(
    Shotsdf_goalpost["goalMouthZ"],
    errors="coerce"
)

Shotsdf_goalpost["expectedGoals"] = pd.to_numeric(
    Shotsdf_goalpost["expectedGoals"],
    errors="coerce"
).fillna(0)

Shotsdf_goalpost["expectedGoalsOnTarget"] = pd.to_numeric(
    Shotsdf_goalpost["expectedGoalsOnTarget"],
    errors="coerce"
).fillna(0)

Shotsdf_goalpost["isBigChance"] = Shotsdf_goalpost["isBigChance"].fillna(False).astype(bool)
Shotsdf_goalpost["isOwnGoal"] = Shotsdf_goalpost["isOwnGoal"].fillna(False).astype(bool)

# ------------------------------------------------------------
# 6. Filtrar SOLO tiros relevantes para gráfico de arqueros
# ------------------------------------------------------------
# Para arco:
# - goal entra
# - save entra
# - miss queda afuera
# - block queda afuera
# ------------------------------------------------------------

Shotsdf_goalpost = Shotsdf_goalpost[
    Shotsdf_goalpost["type"].isin(["goal", "save"])
].copy()

Shotsdf_goalpost = Shotsdf_goalpost.dropna(
    subset=["goalMouthY", "goalMouthZ"]
).copy()

Shotsdf_goalpost.reset_index(drop=True, inplace=True)

# ------------------------------------------------------------
# 7. Control
# ------------------------------------------------------------

print("✅ Shotsdf_goalpost preparado correctamente desde 365Scores")
print("Tipos encontrados:", Shotsdf_goalpost["type"].unique())
print("Equipos encontrados:", Shotsdf_goalpost["teamName"].unique())

print("\nConteo por equipo y tipo:")
display(
    Shotsdf_goalpost.groupby(["teamName", "type"])
    .size()
    .reset_index(name="cantidad")
)

display(
    Shotsdf_goalpost[
        [
            "teamName",
            "playerName",
            "type",
            "goalMouthY",
            "goalMouthZ",
            "expectedGoalsOnTarget",
            "isBigChance",
            "isOwnGoal"
        ]
    ].head(20)
)


# ============================================================
# FUNCIÓN PRINCIPAL: GOALPOST
# ============================================================

def plot_goalPost(ax, debug=False):
    # ============================================================
    # FILTRAR TIROS POR EQUIPO
    # ============================================================

    hShotsdf = Shotsdf_goalpost[
        Shotsdf_goalpost["teamName"] == hteamName
    ].copy()

    aShotsdf = Shotsdf_goalpost[
        Shotsdf_goalpost["teamName"] == ateamName
    ].copy()

    # ============================================================
    # CONVERTIR COORDENADAS AL GRÁFICO DE ARCO
    # ============================================================

    hShotsdf = convertir_goalmouth_365_a_arco(
        hShotsdf,
        top_goal=False,
        clip_horizontal=False,
        clip_vertical=False
    )

    aShotsdf = convertir_goalmouth_365_a_arco(
        aShotsdf,
        top_goal=True,
        clip_horizontal=False,
        clip_vertical=False
    )

    if debug:
        print("Home converted:")
        display(
            hShotsdf[
                [
                    "teamName",
                    "playerName",
                    "type",
                    "goalMouthY",
                    "goalMouthY_opta",
                    "goalMouthY_plot",
                    "goalMouthZ",
                    "goalMouthZ_metros",
                    "goalMouthZ_pct_arco",
                    "goalMouthZ_plot",
                    "expectedGoalsOnTarget"
                ]
            ]
        )

        print("Away converted:")
        display(
            aShotsdf[
                [
                    "teamName",
                    "playerName",
                    "type",
                    "goalMouthY",
                    "goalMouthY_opta",
                    "goalMouthY_plot",
                    "goalMouthZ",
                    "goalMouthZ_metros",
                    "goalMouthZ_pct_arco",
                    "goalMouthZ_plot",
                    "expectedGoalsOnTarget"
                ]
            ]
        )

    # ============================================================
    # CANCHA INVISIBLE
    # ============================================================

    pitch = Pitch(
        pitch_type="uefa",
        corner_arcs=True,
        pitch_color=bg_color,
        line_color=bg_color,
        linewidth=2
    )

    pitch.draw(ax=ax)

    ax.set_ylim(-2, 78)
    ax.set_xlim(-0.5, 105.5)

    # ============================================================
    # DIBUJAR ARCO DE ABAJO
    # ============================================================

    ax.plot([7.5, 7.5], [0, 30], color=line_color, linewidth=5)
    ax.plot([7.5, 97.5], [30, 30], color=line_color, linewidth=5)
    ax.plot([97.5, 97.5], [30, 0], color=line_color, linewidth=5)
    ax.plot([0, 105], [0, 0], color=line_color, linewidth=3)

    y_values = np.arange(0, 6) * 6
    for y in y_values:
        ax.plot([7.5, 97.5], [y, y], color=line_color, linewidth=2, alpha=0.2)

    x_values = (np.arange(0, 11) * 9) + 7.5
    for x in x_values:
        ax.plot([x, x], [0, 30], color=line_color, linewidth=2, alpha=0.2)

    # ============================================================
    # DIBUJAR ARCO DE ARRIBA
    # ============================================================

    ax.plot([7.5, 7.5], [38, 68], color=line_color, linewidth=5)
    ax.plot([7.5, 97.5], [68, 68], color=line_color, linewidth=5)
    ax.plot([97.5, 97.5], [68, 38], color=line_color, linewidth=5)
    ax.plot([0, 105], [38, 38], color=line_color, linewidth=3)

    y_values = (np.arange(0, 6) * 6) + 38
    for y in y_values:
        ax.plot([7.5, 97.5], [y, y], color=line_color, linewidth=2, alpha=0.2)

    x_values = (np.arange(0, 11) * 9) + 7.5
    for x in x_values:
        ax.plot([x, x], [38, 68], color=line_color, linewidth=2, alpha=0.2)

    # ============================================================
    # FILTROS SIN BIG CHANCE
    # ============================================================

    hSavedf = hShotsdf[
        (hShotsdf["type"] == "save") &
        (~bool_col(hShotsdf, "isBigChance"))
    ].copy()

    hGoaldf = hShotsdf[
        (hShotsdf["type"] == "goal") &
        (~bool_col(hShotsdf, "isOwnGoal")) &
        (~bool_col(hShotsdf, "isBigChance"))
    ].copy()

    aSavedf = aShotsdf[
        (aShotsdf["type"] == "save") &
        (~bool_col(aShotsdf, "isBigChance"))
    ].copy()

    aGoaldf = aShotsdf[
        (aShotsdf["type"] == "goal") &
        (~bool_col(aShotsdf, "isOwnGoal")) &
        (~bool_col(aShotsdf, "isBigChance"))
    ].copy()

    # ============================================================
    # FILTROS CON BIG CHANCE
    # ============================================================

    hSavedf_bc = hShotsdf[
        (hShotsdf["type"] == "save") &
        (bool_col(hShotsdf, "isBigChance"))
    ].copy()

    hGoaldf_bc = hShotsdf[
        (hShotsdf["type"] == "goal") &
        (~bool_col(hShotsdf, "isOwnGoal")) &
        (bool_col(hShotsdf, "isBigChance"))
    ].copy()

    aSavedf_bc = aShotsdf[
        (aShotsdf["type"] == "save") &
        (bool_col(aShotsdf, "isBigChance"))
    ].copy()

    aGoaldf_bc = aShotsdf[
        (aShotsdf["type"] == "goal") &
        (~bool_col(aShotsdf, "isOwnGoal")) &
        (bool_col(aShotsdf, "isBigChance"))
    ].copy()

    # ============================================================
    # PLOT TIROS DEL LOCAL
    # Arco inferior: tiros del local contra arquero visitante
    # ============================================================

    ax.scatter(
        hSavedf.goalMouthY_plot,
        hSavedf.goalMouthZ_plot,
        marker="o",
        c=bg_color,
        edgecolors=acol,
        hatch="/////",
        s=350,
        zorder=3
    )

    ax.scatter(
        hGoaldf.goalMouthY_plot,
        hGoaldf.goalMouthZ_plot,
        marker="*",
        c=bg_color,
        edgecolors="green",
        linewidths=2,
        s=550,
        zorder=4
    )

    # ============================================================
    # PLOT TIROS DEL VISITANTE
    # Arco superior: tiros del visitante contra arquero local
    # ============================================================

    ax.scatter(
        aSavedf.goalMouthY_plot,
        aSavedf.goalMouthZ_plot,
        marker="o",
        c=bg_color,
        edgecolors=hcol,
        hatch="/////",
        s=350,
        zorder=3
    )

    ax.scatter(
        aGoaldf.goalMouthY_plot,
        aGoaldf.goalMouthZ_plot,
        marker="*",
        c=bg_color,
        edgecolors="green",
        linewidths=2,
        s=550,
        zorder=4
    )

    # ============================================================
    # PLOT BIG CHANCES LOCAL
    # ============================================================

    ax.scatter(
        hSavedf_bc.goalMouthY_plot,
        hSavedf_bc.goalMouthZ_plot,
        marker="o",
        c=bg_color,
        edgecolors=acol,
        hatch="/////",
        s=1000,
        zorder=5
    )

    ax.scatter(
        hGoaldf_bc.goalMouthY_plot,
        hGoaldf_bc.goalMouthZ_plot,
        marker="*",
        c=bg_color,
        edgecolors="green",
        linewidths=2,
        s=1200,
        zorder=6
    )

    # ============================================================
    # PLOT BIG CHANCES VISITANTE
    # ============================================================

    ax.scatter(
        aSavedf_bc.goalMouthY_plot,
        aSavedf_bc.goalMouthZ_plot,
        marker="o",
        c=bg_color,
        edgecolors=hcol,
        hatch="/////",
        s=1000,
        zorder=5
    )

    ax.scatter(
        aGoaldf_bc.goalMouthY_plot,
        aGoaldf_bc.goalMouthZ_plot,
        marker="*",
        c=bg_color,
        edgecolors="green",
        linewidths=2,
        s=1200,
        zorder=6
    )

    # ============================================================
    # STATS DE ARQUEROS
    # ============================================================

    # xGOT generado por cada equipo
    hxgot_local = round(
        hShotsdf["expectedGoalsOnTarget"].fillna(0).sum(),
        2
    )

    axgot_away = round(
        aShotsdf["expectedGoalsOnTarget"].fillna(0).sum(),
        2
    )

    # Arquero local enfrenta tiros visitantes
    home_gk_saves = len(aSavedf) + len(aSavedf_bc)
    home_gk_big_chance_saves = len(aSavedf_bc)
    home_gk_goals_against = len(aGoaldf) + len(aGoaldf_bc)
    home_goals_prevented = round(axgot_away - home_gk_goals_against, 2)

    # Arquero visitante enfrenta tiros locales
    away_gk_saves = len(hSavedf) + len(hSavedf_bc)
    away_gk_big_chance_saves = len(hSavedf_bc)
    away_gk_goals_against = len(hGoaldf) + len(hGoaldf_bc)
    away_goals_prevented = round(hxgot_local - away_gk_goals_against, 2)

    # ============================================================
    # TEXTOS
    # ============================================================

    ax.text(
        52.5,
        74,
        f"{hteamName} atajadas",
        color=hcol,
        fontsize=30,
        ha="center",
        fontweight="bold"
    )

    ax.text(
        52.5,
        -2,
        f"{ateamName} atajadas",
        color=acol,
        fontsize=30,
        ha="center",
        va="top",
        fontweight="bold"
    )

    ax.text(
        100,
        68,
        f"Atajadas = {home_gk_saves}\n\nxGOT enfrentado:\n{axgot_away}\n\nGoles evitados:\n{home_goals_prevented}",
        color=hcol,
        fontsize=16,
        va="top",
        ha="left"
    )

    ax.text(
        100,
        2,
        f"Atajadas = {away_gk_saves}\n\nxGOT enfrentado:\n{hxgot_local}\n\nGoles evitados:\n{away_goals_prevented}",
        color=acol,
        fontsize=16,
        va="bottom",
        ha="left"
    )

    # ============================================================
    # ESTÉTICA FINAL
    # ============================================================

    for sp in ["top", "right", "bottom", "left"]:
        ax.spines[sp].set_visible(False)

    ax.tick_params(
        axis="both",
        which="both",
        bottom=False,
        top=False,
        left=False,
        right=False
    )

    ax.set_xticks([])
    ax.set_yticks([])

    # ============================================================
    # DATAFRAME DE SALIDA
    # ============================================================

    home_data = {
        "Team_Name": hteamName,
        "Shots_Saved": home_gk_saves,
        "Big_Chance_Saved": home_gk_big_chance_saves,
        "xGOT_Faced": axgot_away,
        "Goals_Against": home_gk_goals_against,
        "Goals_Prevented": home_goals_prevented
    }

    away_data = {
        "Team_Name": ateamName,
        "Shots_Saved": away_gk_saves,
        "Big_Chance_Saved": away_gk_big_chance_saves,
        "xGOT_Faced": hxgot_local,
        "Goals_Against": away_gk_goals_against,
        "Goals_Prevented": away_goals_prevented
    }

    return [home_data, away_data]


# ============================================================
# EJECUTAR
# ============================================================

fig, ax = plt.subplots(figsize=(10, 10), facecolor=bg_color)

goalkeeping_stats = plot_goalPost(ax, debug=True)

goalkeeping_stats_df = pd.DataFrame(goalkeeping_stats)

goalkeeping_stats_df
''',
    ),
    (
        54,
        r'''
Momentumdf = df.copy()
# multiplying the away teams xT values with -1 so that I can plot them in the opposite of home teams
Momentumdf.loc[Momentumdf['teamName'] == ateamName, 'end_zone_value_xT'] *= -1
# taking average xT per minute
Momentumdf = Momentumdf.groupby('minute')['end_zone_value_xT'].mean()
Momentumdf = Momentumdf.reset_index()
Momentumdf.columns = ['minute', 'average_xT']
Momentumdf['average_xT'].fillna(0, inplace=True)
# Momentumdf['average_xT'] = Momentumdf['average_xT'].rolling(window=2, min_periods=1).median()

def plot_Momentum(ax):
    # Set colors based on positive or negative values
    colors = [hcol if x > 0 else acol for x in Momentumdf['average_xT']]

    # making a list of munutes when goals are scored
    hgoal_list = homedf[(homedf['type'] == 'Goal') & (~homedf['qualifiers'].str.contains('OwnGoal'))]['minute'].tolist()
    agoal_list = awaydf[(awaydf['type'] == 'Goal') & (~awaydf['qualifiers'].str.contains('OwnGoal'))]['minute'].tolist()
    hog_list = homedf[(homedf['type'] == 'Goal') & (homedf['qualifiers'].str.contains('OwnGoal'))]['minute'].tolist()
    aog_list = awaydf[(awaydf['type'] == 'Goal') & (awaydf['qualifiers'].str.contains('OwnGoal'))]['minute'].tolist()
    hred_list = homedf[homedf['qualifiers'].str.contains('Red|SecondYellow')]['minute'].tolist()
    ared_list = awaydf[awaydf['qualifiers'].str.contains('Red|SecondYellow')]['minute'].tolist()

    # plotting scatters when goals are scored
    highest_xT = Momentumdf['average_xT'].max()
    lowest_xT = Momentumdf['average_xT'].min()
    highest_minute = Momentumdf['minute'].max()
    hscatter_y = [highest_xT]*len(hgoal_list)
    ascatter_y = [lowest_xT]*len(agoal_list)
    hogscatter_y = [highest_xT]*len(aog_list)
    aogscatter_y = [lowest_xT]*len(hog_list)
    hred_y = [highest_xT]*len(hred_list)
    ared_y = [lowest_xT]*len(ared_list)

    ax.text((45/2), lowest_xT, 'First Half', color='gray', fontsize=20, alpha=0.25, va='center', ha='center')
    ax.text((45+(45/2)), lowest_xT, 'Second Half', color='gray', fontsize=20, alpha=0.25, va='center', ha='center')

    ax.scatter(hgoal_list, hscatter_y, s=250, c='None', edgecolor='green', hatch='////', marker='o')
    ax.scatter(agoal_list, ascatter_y, s=250, c='None', edgecolor='green', hatch='////', marker='o')
    ax.scatter(hog_list, aogscatter_y, s=250, c='None', edgecolor='orange', hatch='////', marker='o')
    ax.scatter(aog_list, hogscatter_y, s=250, c='None', edgecolor='orange', hatch='////', marker='o')
    ax.scatter(hred_list, hred_y, s=250, c='None', edgecolor='red', hatch='////', marker='s')
    ax.scatter(ared_list, ared_y, s=250, c='None', edgecolor='red', hatch='////', marker='s')

    # Creating the bar plot
    ax.bar(Momentumdf['minute'], Momentumdf['average_xT'], color=colors)
    ax.set_xticks(range(0, len(Momentumdf['minute']), 5))
    ax.axvline(45, color='gray', linewidth=2, linestyle='dotted')
    # ax.axvline(90, color='gray', linewidth=2, linestyle='dotted')
    ax.set_facecolor(bg_color)
    # Hide spines
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    # # Hide ticks
    ax.tick_params(axis='both', which='both', length=0)
    ax.tick_params(axis='x', colors=line_color)
    ax.tick_params(axis='y', colors=line_color)
    # Add labels and title
    ax.set_xlabel('Minutos', color=line_color, fontsize=20)
    ax.set_ylabel('Promedio de xT por minuto', color=line_color, fontsize=20)
    ax.axhline(y=0, color=line_color, alpha=1, linewidth=2)

    ax.text(highest_minute+1,highest_xT, f"{hteamName}\nxT: {hxT}", color=hcol, fontsize=20, va='bottom', ha='left')
    ax.text(highest_minute+1,lowest_xT,  f"{ateamName}\nxT: {axT}", color=acol, fontsize=20, va='top', ha='left')

    ax.set_title('Match Momentum (xT)', color=line_color, fontsize=30, fontweight='bold')

    home_data = {
        'Team_Name': hteamName,
        'xT': hxT
    }
    
    away_data = {
        'Team_Name': ateamName,
        'xT': axT
    }
    
    return [home_data, away_data]

fig,ax=plt.subplots(figsize=(10,10), facecolor=bg_color)
plot_Momentum(ax)
xT_stats = plot_Momentum(ax)
xT_stats_df = pd.DataFrame(xT_stats)
''',
    ),
    (
        55,
        r'''
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patheffects as path_effects
from mplsoccer import Pitch

# ============================================================
# HELPERS
# ============================================================

def bool_col(dataframe, col):
    """
    Devuelve una Serie booleana.
    Si la columna no existe, devuelve False para todas las filas.
    """
    if col not in dataframe.columns:
        return pd.Series(False, index=dataframe.index)
    
    return dataframe[col].fillna(False).astype(bool)


def safe_pct(num, den, default=0):
    """
    Porcentaje seguro.
    """
    if den == 0:
        return default
    return round((num / den) * 100, 2)


def safe_div(num, den, default=0):
    """
    División segura.
    """
    if den == 0:
        return default
    return round(num / den, 2)


def safe_mean(series, default=0):
    """
    Media segura.
    """
    series = pd.to_numeric(series, errors="coerce").dropna()
    if len(series) == 0:
        return default
    return round(series.mean(), 2)


def safe_norm_home_away(home_value, away_value, scale=50):
    """
    Normaliza dos valores para barras enfrentadas.
    Home negativo, Away positivo.
    """
    home_value = 0 if pd.isna(home_value) else home_value
    away_value = 0 if pd.isna(away_value) else away_value

    total = home_value + away_value

    if total == 0:
        return 0, 0

    return -(home_value / total) * scale, (away_value / total) * scale


# ============================================================
# PREPARACIÓN GENERAL
# ============================================================

df = df.copy()

for col in ["x", "y", "endX", "endY", "Length"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

for col in ["teamName", "type", "outcomeType"]:
    if col in df.columns:
        df[col] = df[col].astype(str)


# ============================================================
# MÁSCARAS OFICIAL-LIKE OPTA
# ============================================================

# ------------------------------------------------------------
# Pases official-like
# ------------------------------------------------------------
# Replica Pass Details:
# - type == Pass
# - excluye ThrowIn
# - excluye Cross
# - excluye KeeperThrow
#
# NO excluye GoalKick en general.
# NO excluye FreekickTaken en general.
# ------------------------------------------------------------

passes_mask = (
    (df["type"] == "Pass") &
    (~bool_col(df, "isThrowIn")) &
    (~bool_col(df, "isCross")) &
    (~bool_col(df, "isKeeperThrow"))
)


# ------------------------------------------------------------
# Long Balls official-like
# ------------------------------------------------------------
# Replica Long Ball oficial:
# - type == Pass
# - isLongball == True
# - excluye ThrowIn
# - excluye Cross
# - excluye KeeperThrow
#
# NO excluye GoalKick.
# NO excluye FreekickTaken.
# ------------------------------------------------------------

longball_mask = (
    (df["type"] == "Pass") &
    (bool_col(df, "isLongball")) &
    (~bool_col(df, "isThrowIn")) &
    (~bool_col(df, "isCross")) &
    (~bool_col(df, "isKeeperThrow"))
)


# ------------------------------------------------------------
# Clearances official-like
# ------------------------------------------------------------
# Replica reporte oficial de Clearances:
# - type == Clearance
# - endX no nulo
# - endY no nulo
# - Length no nulo
# ------------------------------------------------------------

clearance_mask = (
    (df["type"] == "Clearance") &
    (df["endX"].notna()) &
    (df["endY"].notna()) &
    (df["Length"].notna())
)


# ============================================================
# PASSING STATS CORREGIDOS
# ============================================================

# ------------------------------------------------------------
# Possession % con pases official-like
# ------------------------------------------------------------

hpossdf = df[
    (df["teamName"] == hteamName) &
    passes_mask
]

apossdf = df[
    (df["teamName"] == ateamName) &
    passes_mask
]

total_poss_passes = len(hpossdf) + len(apossdf)

hposs = safe_pct(len(hpossdf), total_poss_passes)
aposs = safe_pct(len(apossdf), total_poss_passes)


# ------------------------------------------------------------
# Field Tilt %
# ------------------------------------------------------------

hftdf = df[
    (df["teamName"] == hteamName) &
    (df["isTouch"] == 1) &
    (df["x"] >= 70)
]

aftdf = df[
    (df["teamName"] == ateamName) &
    (df["isTouch"] == 1) &
    (df["x"] >= 70)
]

total_ft = len(hftdf) + len(aftdf)

hft = safe_pct(len(hftdf), total_ft)
aft = safe_pct(len(aftdf), total_ft)


# ------------------------------------------------------------
# Total Passes
# ------------------------------------------------------------

htotalPass = len(df[
    (df["teamName"] == hteamName) &
    passes_mask
])

atotalPass = len(df[
    (df["teamName"] == ateamName) &
    passes_mask
])


# ------------------------------------------------------------
# Accurate Passes
# ------------------------------------------------------------

hAccPass = len(df[
    (df["teamName"] == hteamName) &
    passes_mask &
    (df["outcomeType"] == "Successful")
])

aAccPass = len(df[
    (df["teamName"] == ateamName) &
    passes_mask &
    (df["outcomeType"] == "Successful")
])


# ------------------------------------------------------------
# Accurate Passes without defensive third
# ------------------------------------------------------------

hAccPasswdt = len(df[
    (df["teamName"] == hteamName) &
    passes_mask &
    (df["outcomeType"] == "Successful") &
    (df["endX"] > 35)
])

aAccPasswdt = len(df[
    (df["teamName"] == ateamName) &
    passes_mask &
    (df["outcomeType"] == "Successful") &
    (df["endX"] > 35)
])


# ============================================================
# LONG BALLS CORREGIDOS
# ============================================================

hLongB = len(df[
    (df["teamName"] == hteamName) &
    longball_mask
])

aLongB = len(df[
    (df["teamName"] == ateamName) &
    longball_mask
])


hAccLongB = len(df[
    (df["teamName"] == hteamName) &
    longball_mask &
    (df["outcomeType"] == "Successful")
])

aAccLongB = len(df[
    (df["teamName"] == ateamName) &
    longball_mask &
    (df["outcomeType"] == "Successful")
])


# ============================================================
# RESTO DE PASSING / SET PIECES
# ============================================================

# ------------------------------------------------------------
# Crosses
# ------------------------------------------------------------

hCrss = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCross"))
])

aCrss = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCross"))
])


hAccCrss = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCross")) &
    (df["outcomeType"] == "Successful")
])

aAccCrss = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCross")) &
    (df["outcomeType"] == "Successful")
])


# ------------------------------------------------------------
# Corners
# ------------------------------------------------------------

hCor = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCornerTaken"))
])

aCor = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isCornerTaken"))
])


# ------------------------------------------------------------
# GoalKick Length
# ------------------------------------------------------------
# Se mantiene separado del pass_mask.
# Usamos isGoalKick + Length directa.
# ------------------------------------------------------------

home_goalkick = df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isGoalKick"))
].copy()

away_goalkick = df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Pass") &
    (bool_col(df, "isGoalKick"))
].copy()

hglkl = safe_mean(home_goalkick["Length"]) if "Length" in home_goalkick.columns else 0
aglkl = safe_mean(away_goalkick["Length"]) if "Length" in away_goalkick.columns else 0


# ============================================================
# DEFENSIVE STATS
# ============================================================

# ------------------------------------------------------------
# Tackles
# ------------------------------------------------------------

htkl = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Tackle")
])

atkl = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Tackle")
])


# ------------------------------------------------------------
# Tackles Won
# ------------------------------------------------------------

htklw = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Tackle") &
    (df["outcomeType"] == "Successful")
])

atklw = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Tackle") &
    (df["outcomeType"] == "Successful")
])


# ------------------------------------------------------------
# Interceptions
# ------------------------------------------------------------

hintc = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Interception")
])

aintc = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Interception")
])


# ------------------------------------------------------------
# Clearances CORREGIDOS
# ------------------------------------------------------------

hclr = len(df[
    (df["teamName"] == hteamName) &
    clearance_mask
])

aclr = len(df[
    (df["teamName"] == ateamName) &
    clearance_mask
])


# ------------------------------------------------------------
# Aerials
# ------------------------------------------------------------

harl = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Aerial")
])

aarl = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Aerial")
])


harlw = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Aerial") &
    (df["outcomeType"] == "Successful")
])

aarlw = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Aerial") &
    (df["outcomeType"] == "Successful")
])


# ------------------------------------------------------------
# Fouls
# ------------------------------------------------------------

hfoul = len(df[
    (df["teamName"] == hteamName) &
    (df["type"] == "Foul")
])

afoul = len(df[
    (df["teamName"] == ateamName) &
    (df["type"] == "Foul")
])


# ============================================================
# PPS / SECUENCIAS CON PASES CORREGIDOS
# ============================================================

pass_df_home = df[
    (df["teamName"] == hteamName) &
    passes_mask
]

pass_counts_home = pass_df_home.groupby("possession_id").size()

PPS_home = round(pass_counts_home.mean()) if len(pass_counts_home) > 0 else 0
pass_seq_10_more_home = pass_counts_home[pass_counts_home >= 10].count()


pass_df_away = df[
    (df["teamName"] == ateamName) &
    passes_mask
]

pass_counts_away = pass_df_away.groupby("possession_id").size()

PPS_away = round(pass_counts_away.mean()) if len(pass_counts_away) > 0 else 0
pass_seq_10_more_away = pass_counts_away[pass_counts_away >= 10].count()


# ============================================================
# PPDA CON PASES CORREGIDOS
# ============================================================

home_def_acts = df[
    (df["teamName"] == hteamName) &
    (df["type"].str.contains("Interception|Foul|Challenge|BlockedPass|Tackle", na=False)) &
    (df["x"] > 35)
]

away_def_acts = df[
    (df["teamName"] == ateamName) &
    (df["type"].str.contains("Interception|Foul|Challenge|BlockedPass|Tackle", na=False)) &
    (df["x"] > 35)
]

home_pass = df[
    (df["teamName"] == hteamName) &
    passes_mask &
    (df["outcomeType"] == "Successful") &
    (df["x"] < 70)
]

away_pass = df[
    (df["teamName"] == ateamName) &
    passes_mask &
    (df["outcomeType"] == "Successful") &
    (df["x"] < 70)
]

home_ppda = safe_div(len(away_pass), len(home_def_acts), default=0)
away_ppda = safe_div(len(home_pass), len(away_def_acts), default=0)


# ============================================================
# CHEQUEO FINAL
# ============================================================

print("========== CHEQUEO FINAL ==========")

print("\nPases official-like:")
print(hteamName, htotalPass, "accurate:", hAccPass)
print(ateamName, atotalPass, "accurate:", aAccPass)

print("\nLong Balls official-like:")
print(hteamName, hLongB, "accurate:", hAccLongB)
print(ateamName, aLongB, "accurate:", aAccLongB)

print("\nClearances official-like:")
print(hteamName, hclr)
print(ateamName, aclr)

print("\nAvg. GoalKick Length:")
print(hteamName, hglkl)
print(ateamName, aglkl)


# ============================================================
# PATH EFFECTS
# ============================================================

path_eff1 = [
    path_effects.Stroke(linewidth=1.5, foreground=line_color),
    path_effects.Normal()
]

try:
    path_eff
except NameError:
    path_eff = path_eff1


# ============================================================
# VISUALIZACIÓN MATCH STATS
# ============================================================

def plotting_match_stats(ax):
    pitch = Pitch(
        pitch_type="uefa",
        corner_arcs=True,
        pitch_color=bg_color,
        line_color=bg_color,
        linewidth=2
    )

    pitch.draw(ax=ax)

    ax.set_xlim(-0.5, 105.5)
    ax.set_ylim(-5, 68.5)

    # ------------------------------------------------------------
    # Headline box
    # ------------------------------------------------------------

    head_y = [62, 68, 68, 62]
    head_x = [0, 0, 105, 105]

    ax.fill(head_x, head_y, "orange")

    ax.text(
        52.5,
        64.5,
        "Match Stats",
        ha="center",
        va="center",
        color=line_color,
        fontsize=25,
        fontweight="bold",
        path_effects=path_eff
    )

    # ------------------------------------------------------------
    # Stats values
    # ------------------------------------------------------------

    stats_title = [58 - (i * 6) for i in range(11)]

    stats_home = [
        hposs,
        hft,
        htotalPass,
        hLongB,
        hCor,
        hglkl,
        htkl,
        hintc,
        hclr,
        harl,
        home_ppda
    ]

    stats_away = [
        aposs,
        aft,
        atotalPass,
        aLongB,
        aCor,
        aglkl,
        atkl,
        aintc,
        aclr,
        aarl,
        away_ppda
    ]

    stats_normalized_home = []
    stats_normalized_away = []

    for hv, av in zip(stats_home, stats_away):
        nh, na = safe_norm_home_away(hv, av, scale=50)
        stats_normalized_home.append(nh)
        stats_normalized_away.append(na)

    start_x = 52.5

    ax.barh(
        stats_title,
        stats_normalized_home,
        height=4,
        color=hcol,
        left=start_x
    )

    ax.barh(
        stats_title,
        stats_normalized_away,
        height=4,
        color=acol,
        left=start_x
    )

    # ------------------------------------------------------------
    # Axis off
    # ------------------------------------------------------------

    for sp in ["top", "right", "bottom", "left"]:
        ax.spines[sp].set_visible(False)

    ax.tick_params(
        axis="both",
        which="both",
        bottom=False,
        top=False,
        left=False,
        right=False
    )

    ax.set_xticks([])
    ax.set_yticks([])

    # ------------------------------------------------------------
    # Labels
    # ------------------------------------------------------------

    labels = [
        "Posesión",
        "Dominio de campo",
        "Pases (Comp.)",
        "Envíos largos (Comp.)",
        "Córners",
        "Distancia saque de puerta.",
        "Entradas (ganadas)",
        "Intercepciones",
        "Despejes",
        "Duelos aéreos (ganados)",
        "PPDA"
    ]

    for yy, label in zip(stats_title, labels):
        ax.text(
            52.5,
            yy,
            label,
            color=bg_color,
            fontsize=17,
            ha="center",
            va="center",
            fontweight="bold",
            path_effects=path_eff1
        )

    # ------------------------------------------------------------
    # Home values
    # ------------------------------------------------------------

    home_texts = [
        f"{round(hposs)}%",
        f"{round(hft)}%",
        f"{htotalPass}({hAccPass})",
        f"{hLongB}({hAccLongB})",
        f"{hCor}",
        f"{hglkl} m",
        f"{htkl}({htklw})",
        f"{hintc}",
        f"{hclr}",
        f"{harl}({harlw})",
        f"{home_ppda}"
    ]

    for yy, txt in zip(stats_title, home_texts):
        ax.text(
            0,
            yy,
            txt,
            color=line_color,
            fontsize=20,
            ha="right",
            va="center",
            fontweight="bold"
        )

    # ------------------------------------------------------------
    # Away values
    # ------------------------------------------------------------

    away_texts = [
        f"{round(aposs)}%",
        f"{round(aft)}%",
        f"{atotalPass}({aAccPass})",
        f"{aLongB}({aAccLongB})",
        f"{aCor}",
        f"{aglkl} m",
        f"{atkl}({atklw})",
        f"{aintc}",
        f"{aclr}",
        f"{aarl}({aarlw})",
        f"{away_ppda}"
    ]

    for yy, txt in zip(stats_title, away_texts):
        ax.text(
            105,
            yy,
            txt,
            color=line_color,
            fontsize=20,
            ha="left",
            va="center",
            fontweight="bold"
        )

    # ------------------------------------------------------------
    # Data output
    # ------------------------------------------------------------

    home_data = {
        "Team_Name": hteamName,
        "Possession_%": hposs,
        "Field_Tilt_%": hft,
        "Total_Passes": htotalPass,
        "Accurate_Passes": hAccPass,
        "Longballs": hLongB,
        "Accurate_Longballs": hAccLongB,
        "Corners": hCor,
        "Avg.GoalKick_Length": hglkl,
        "Tackles": htkl,
        "Tackles_Won": htklw,
        "Interceptions": hintc,
        "Clearances": hclr,
        "Aerial_Duels": harl,
        "Aerial_Duels_Won": harlw,
        "Passes_Per_Defensive_Actions(PPDA)": home_ppda,
        "Average_Passes_Per_Sequences": PPS_home,
        "10+_Passing_Sequences": pass_seq_10_more_home
    }

    away_data = {
        "Team_Name": ateamName,
        "Possession_%": aposs,
        "Field_Tilt_%": aft,
        "Total_Passes": atotalPass,
        "Accurate_Passes": aAccPass,
        "Longballs": aLongB,
        "Accurate_Longballs": aAccLongB,
        "Corners": aCor,
        "Avg.GoalKick_Length": aglkl,
        "Tackles": atkl,
        "Tackles_Won": atklw,
        "Interceptions": aintc,
        "Clearances": aclr,
        "Aerial_Duels": aarl,
        "Aerial_Duels_Won": aarlw,
        "Passes_Per_Defensive_Actions(PPDA)": away_ppda,
        "Average_Passes_Per_Sequences": PPS_away,
        "10+_Passing_Sequences": pass_seq_10_more_away
    }

    return [home_data, away_data]


# ============================================================
# RUN VISUALIZACIÓN
# ============================================================

fig, ax = plt.subplots(figsize=(10, 10), facecolor=bg_color)

general_match_stats = plotting_match_stats(ax)

general_match_stats_df = pd.DataFrame(general_match_stats)

general_match_stats_df
''',
    ),
    (
        56,
        r'''
def Final_third_entry(ax, team_name, col):
    # Final third Entry means passes or carries which has started outside the Final third and ended inside the final third
    dfpass = df[(df['teamName']==team_name) & (df['type']=='Pass') & (df['x']<70) & (df['endX']>=70) & (df['outcomeType']=='Successful') &
                (~df['qualifiers'].str.contains('Freekick'))]
    dfcarry = df[(df['teamName']==team_name) & (df['type']=='Carry') & (df['x']<70) & (df['endX']>=70)]
    pitch = Pitch(pitch_type='uefa', pitch_color=bg_color, line_color=line_color, linewidth=2,
                          corner_arcs=True)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    # ax.set_ylim(-0.5, 68.5)

    if team_name == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()

    pass_count = len(dfpass) + len(dfcarry)

    # calculating the counts
    left_entry = len(dfpass[dfpass['y']>=45.33]) + len(dfcarry[dfcarry['y']>=45.33])
    mid_entry = len(dfpass[(dfpass['y']>=22.67) & (dfpass['y']<45.33)]) + len(dfcarry[(dfcarry['y']>=22.67) & (dfcarry['y']<45.33)])
    right_entry = len(dfpass[(dfpass['y']>=0) & (dfpass['y']<22.67)]) + len(dfcarry[(dfcarry['y']>=0) & (dfcarry['y']<22.67)])
    left_percentage = round((left_entry/pass_count)*100)
    mid_percentage = round((mid_entry/pass_count)*100)
    right_percentage = round((right_entry/pass_count)*100)

    ax.hlines(22.67, xmin=0, xmax=70, colors=line_color, linestyle='dashed', alpha=0.35)
    ax.hlines(45.33, xmin=0, xmax=70, colors=line_color, linestyle='dashed', alpha=0.35)
    ax.vlines(70, ymin=-2, ymax=70, colors=line_color, linestyle='dashed', alpha=0.55)

    # showing the texts in the pitch
    bbox_props = dict(boxstyle="round,pad=0.3", edgecolor="None", facecolor=bg_color, alpha=0.75)
    if col == hcol:
        ax.text(8, 11.335, f'{right_entry}\n({right_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_entry}\n({mid_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_entry}\n({left_percentage}%)', color=hcol, fontsize=24, va='center', ha='center', bbox=bbox_props)
    else:
        ax.text(8, 11.335, f'{right_entry}\n({right_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 34, f'{mid_entry}\n({mid_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)
        ax.text(8, 56.675, f'{left_entry}\n({left_percentage}%)', color=acol, fontsize=24, va='center', ha='center', bbox=bbox_props)

    # plotting the passes
    pro_pass = pitch.lines(dfpass.x, dfpass.y, dfpass.endX, dfpass.endY, lw=3.5, comet=True, color=col, ax=ax, alpha=0.5)
    # plotting some scatters at the end of each pass
    pro_pass_end = pitch.scatter(dfpass.endX, dfpass.endY, s=35, edgecolor=col, linewidth=1, color=bg_color, zorder=2, ax=ax)
    # plotting carries
    for index, row in dfcarry.iterrows():
        arrow = patches.FancyArrowPatch((row['x'], row['y']), (row['endX'], row['endY']), arrowstyle='->', color=col, zorder=4, mutation_scale=20, 
                                        alpha=1, linewidth=2, linestyle='--')
        ax.add_patch(arrow)

    counttext = f"{pass_count} Ingresos al tercio final"

    # Heading and other texts
    if col == hcol:
        ax.set_title(f"{hteamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold', path_effects=path_eff)
        ax.text(87.5, 70, '<--------------- Tercio final --------------->', color=line_color, ha='center', va='center')
        pitch.lines(53, -2, 73, -2, lw=3, transparent=True, comet=True, color=col, ax=ax, alpha=0.5)
        ax.scatter(73,-2, s=35, edgecolor=col, linewidth=1, color=bg_color, zorder=2)
        arrow = patches.FancyArrowPatch((83, -2), (103, -2), arrowstyle='->', color=col, zorder=4, mutation_scale=20, 
                                        alpha=1, linewidth=2, linestyle='--')
        ax.add_patch(arrow)
        ax.text(63, -5, f'Por pase: {len(dfpass)}', fontsize=15, color=line_color, ha='center', va='center')
        ax.text(93, -5, f'Por conducción: {len(dfcarry)}', fontsize=15, color=line_color, ha='center', va='center')
        
    else:
        ax.set_title(f"{ateamName}\n{counttext}", color=line_color, fontsize=25, fontweight='bold', path_effects=path_eff)
        ax.text(87.5, -2, '<--------------- Tercio final --------------->', color=line_color, ha='center', va='center')
        pitch.lines(53, 70, 73, 70, lw=3, transparent=True, comet=True, color=col, ax=ax, alpha=0.5)
        ax.scatter(73,70, s=35, edgecolor=col, linewidth=1, color=bg_color, zorder=2)
        arrow = patches.FancyArrowPatch((83, 70), (103, 70), arrowstyle='->', color=col, zorder=4, mutation_scale=20, 
                                        alpha=1, linewidth=2, linestyle='--')
        ax.add_patch(arrow)
        ax.text(63, 73, f'Por pase: {len(dfpass)}', fontsize=15, color=line_color, ha='center', va='center')
        ax.text(93, 73, f'Por conducción: {len(dfcarry)}', fontsize=15, color=line_color, ha='center', va='center')

    return {
        'Team_Name': team_name,
        'Total_Final_Third_Entries': pass_count,
        'Final_Third_Entries_From_Left': left_entry,
        'Final_Third_Entries_From_Center': mid_entry,
        'Final_Third_Entries_From_Right': right_entry,
        'Entry_By_Pass': len(dfpass),
        'Entry_By_Carry': len(dfcarry)
    }

fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
final_third_entry_stats_home = Final_third_entry(axs[0], hteamName, hcol)
final_third_entry_stats_away = Final_third_entry(axs[1], ateamName, acol)
final_third_entry_stats_list = []
final_third_entry_stats_list.append(final_third_entry_stats_home)
final_third_entry_stats_list.append(final_third_entry_stats_away)
final_third_entry_stats_df = pd.DataFrame(final_third_entry_stats_list)
''',
    ),
    (
        57,
        r'''
def zone14hs(ax, team_name, col):
    dfhp = df[(df['teamName']==team_name) & (df['type']=='Pass') & (df['outcomeType']=='Successful') & 
              (~df['qualifiers'].str.contains('CornerTaken|Freekick'))]
    
    pitch = Pitch(pitch_type='uefa', pitch_color=bg_color, line_color=line_color,  linewidth=2,
                          corner_arcs=True)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    ax.set_facecolor(bg_color)
    if team_name == ateamName:
      ax.invert_xaxis()
      ax.invert_yaxis()

    # setting the count varibale
    z14 = 0
    hs = 0
    lhs = 0
    rhs = 0

    path_eff = [path_effects.Stroke(linewidth=3, foreground=bg_color), path_effects.Normal()]
    # iterating ecah pass and according to the conditions plotting only zone14 and half spaces passes
    for index, row in dfhp.iterrows():
        if row['endX'] >= 70 and row['endX'] <= 88.54 and row['endY'] >= 22.66 and row['endY'] <= 45.32:
            pitch.lines(row['x'], row['y'], row['endX'], row['endY'], color='orange', comet=True, lw=3, zorder=3, ax=ax, alpha=0.75)
            ax.scatter(row['endX'], row['endY'], s=35, linewidth=1, color=bg_color, edgecolor='orange', zorder=4)
            z14 += 1
        if row['endX'] >= 70 and row['endY'] >= 11.33 and row['endY'] <= 22.66:
            pitch.lines(row['x'], row['y'], row['endX'], row['endY'], color=col, comet=True, lw=3, zorder=3, ax=ax, alpha=0.75)
            ax.scatter(row['endX'], row['endY'], s=35, linewidth=1, color=bg_color, edgecolor=col, zorder=4)
            hs += 1
            rhs += 1
        if row['endX'] >= 70 and row['endY'] >= 45.32 and row['endY'] <= 56.95:
            pitch.lines(row['x'], row['y'], row['endX'], row['endY'], color=col, comet=True, lw=3, zorder=3, ax=ax, alpha=0.75)
            ax.scatter(row['endX'], row['endY'], s=35, linewidth=1, color=bg_color, edgecolor=col, zorder=4)
            hs += 1
            lhs += 1

    # coloring those zones in the pitch
    y_z14 = [22.66, 22.66, 45.32, 45.32]
    x_z14 = [70, 88.54, 88.54, 70]
    ax.fill(x_z14, y_z14, 'orange', alpha=0.2, label='Zone14')

    y_rhs = [11.33, 11.33, 22.66, 22.66]
    x_rhs = [70, 105, 105, 70]
    ax.fill(x_rhs, y_rhs, col, alpha=0.2, label='HalfSpaces')

    y_lhs = [45.32, 45.32, 56.95, 56.95]
    x_lhs = [70, 105, 105, 70]
    ax.fill(x_lhs, y_lhs, col, alpha=0.2, label='HalfSpaces')

    # showing the counts in an attractive way
    z14name = "Zona 14"
    hsname = "Int."
    z14count = f"{z14}"
    hscount = f"{hs}"
    ax.scatter(16.46, 13.85, color=col, s=15000, edgecolor=line_color, linewidth=2, alpha=1, marker='h')
    ax.scatter(16.46, 54.15, color='orange', s=15000, edgecolor=line_color, linewidth=2, alpha=1, marker='h')
    ax.text(16.46, 13.85-4, hsname, fontsize=20, color=line_color, ha='center', va='center', path_effects=path_eff)
    ax.text(16.46, 54.15-4, z14name, fontsize=20, color=line_color, ha='center', va='center', path_effects=path_eff)
    ax.text(16.46, 13.85+2, hscount, fontsize=40, color=line_color, ha='center', va='center', path_effects=path_eff)
    ax.text(16.46, 54.15+2, z14count, fontsize=40, color=line_color, ha='center', va='center', path_effects=path_eff)

    # Headings and other texts
    if col == hcol:
      ax.set_title(f"{hteamName}\nPases a Zona14 y carriles interiores", color=line_color, fontsize=25, fontweight='bold')
    else:
      ax.set_title(f"{ateamName}\nPases a Zona14 y carriles interiores", color=line_color, fontsize=25, fontweight='bold')

    return {
        'Team_Name': team_name,
        'Total_Passes_Into_Zone14': z14,
        'Passes_Into_Halfspaces': hs,
        'Passes_Into_Left_Halfspaces': lhs,
        'Passes_Into_Right_Halfspaces': rhs
    }

fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
zonal_passing_stats_home = zone14hs(axs[0], hteamName, hcol)
zonal_passing_stats_away = zone14hs(axs[1], ateamName, acol)
zonal_passing_stats_list = []
zonal_passing_stats_list.append(zonal_passing_stats_home)
zonal_passing_stats_list.append(zonal_passing_stats_away)
zonal_passing_stats_df = pd.DataFrame(zonal_passing_stats_list)
''',
    ),
    (
        58,
        r'''
# setting the custom colormap
pearl_earring_cmaph = LinearSegmentedColormap.from_list("Pearl Earring - 10 colors",  [bg_color, hcol], N=20)
pearl_earring_cmapa = LinearSegmentedColormap.from_list("Pearl Earring - 10 colors",  [bg_color, acol], N=20)

path_eff = [path_effects.Stroke(linewidth=3, foreground=bg_color), path_effects.Normal()]

# Getting heatmap of all the end point of the successful Passes
def Pass_end_zone(ax, team_name, cm):
    pez = df[(df['teamName'] == team_name) & (df['type'] == 'Pass') & (df['outcomeType'] == 'Successful')]
    pitch = Pitch(pitch_type='uefa', line_color=line_color, goal_type='box', goal_alpha=.5, corner_arcs=True, line_zorder=2, pitch_color=bg_color, linewidth=2)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    if team_name == ateamName:
      ax.invert_xaxis()
      ax.invert_yaxis()

    pearl_earring_cmap = cm
    # binning the data points
    bin_statistic = pitch.bin_statistic(pez.endX, pez.endY, bins=(6, 5), normalize=True)
    pitch.heatmap(bin_statistic, ax=ax, cmap=pearl_earring_cmap, edgecolors=bg_color)
    pitch.scatter(df.endX, df.endY, c='gray', s=5, ax=ax)
    labels = pitch.label_heatmap(bin_statistic, color=line_color, fontsize=25, ax=ax, ha='center', va='center', str_format='{:.0%}', path_effects=path_eff)

    # Headings and other texts
    if team_name == hteamName:
      ax.set_title(f"{hteamName}\nDestino de pases", color=line_color, fontsize=25, fontweight='bold', path_effects=path_eff)
    else:
      ax.set_title(f"{ateamName}\nDestino de pases", color=line_color, fontsize=25, fontweight='bold', path_effects=path_eff)

fig,axs=plt.subplots(1,2, figsize=(20,10), facecolor=bg_color)
Pass_end_zone(axs[0], hteamName, pearl_earring_cmaph)
Pass_end_zone(axs[1], ateamName, pearl_earring_cmapa)
''',
    ),
    (
        59,
        r'''
# setting the custom colormap
pearl_earring_cmaph = LinearSegmentedColormap.from_list(
    "Pearl Earring - Home",
    [bg_color, hcol],
    N=20
)

pearl_earring_cmapa = LinearSegmentedColormap.from_list(
    "Pearl Earring - Away",
    [bg_color, acol],
    N=20
)

path_eff = [
    path_effects.Stroke(linewidth=3, foreground=bg_color),
    path_effects.Normal()
]


def Chance_creating_zone(ax, team_name, cm, col):
    df_aux = df.copy()

    # ============================================================
    # NORMALIZAR COLUMNAS
    # ============================================================

    if "keyPass" not in df_aux.columns:
        df_aux["keyPass"] = 0

    if "assist" not in df_aux.columns:
        df_aux["assist"] = 0

    df_aux["keyPass_num"] = pd.to_numeric(
        df_aux["keyPass"],
        errors="coerce"
    ).fillna(0)

    df_aux["assist_num"] = pd.to_numeric(
        df_aux["assist"],
        errors="coerce"
    ).fillna(0)

    # ============================================================
    # FILTRO CORRECTO
    # ============================================================
    # Entra si fue key pass O si fue asistencia.
    # Esto evita perder asistencias que no vienen marcadas como keyPass.
    # ============================================================

    ccp = df_aux[
        (df_aux["teamName"] == team_name) &
        (
            (df_aux["keyPass_num"] == 1) |
            (df_aux["assist_num"] == 1)
        )
    ].copy()

    # Asegurar coordenadas numéricas
    for coord_col in ["x", "y", "endX", "endY"]:
        ccp[coord_col] = pd.to_numeric(ccp[coord_col], errors="coerce")

    ccp = ccp.dropna(subset=["x", "y", "endX", "endY"]).copy()

    pitch = Pitch(
        pitch_type="uefa",
        line_color=line_color,
        corner_arcs=True,
        line_zorder=2,
        pitch_color=bg_color,
        linewidth=2
    )

    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)

    if team_name == ateamName:
        ax.invert_xaxis()
        ax.invert_yaxis()

    # ============================================================
    # HEATMAP
    # ============================================================

    bin_statistic = pitch.bin_statistic(
        ccp["x"],
        ccp["y"],
        bins=(6, 5),
        statistic="count",
        normalize=False
    )

    pitch.heatmap(
        bin_statistic,
        ax=ax,
        cmap=cm,
        edgecolors="#f8f8f8"
    )

    # ============================================================
    # LÍNEAS
    # ============================================================

    cc = 0

    for _, row in ccp.iterrows():

        # Si fue asistencia, verde.
        # Si fue key pass sin asistencia, violeta.
        if row["assist_num"] == 1:
            pass_color = green
        else:
            pass_color = violet

        pitch.lines(
            row["x"],
            row["y"],
            row["endX"],
            row["endY"],
            color=pass_color,
            comet=True,
            lw=3,
            zorder=3,
            ax=ax
        )

        ax.scatter(
            row["endX"],
            row["endY"],
            s=35,
            linewidth=1,
            color=bg_color,
            edgecolor=pass_color,
            zorder=4
        )

        cc += 1

    pitch.label_heatmap(
        bin_statistic,
        color=line_color,
        fontsize=25,
        ax=ax,
        ha="center",
        va="center",
        str_format="{:.0f}",
        path_effects=path_eff
    )

    # ============================================================
    # TEXTOS
    # ============================================================

    total_assists = int((ccp["assist_num"] == 1).sum())
    total_keypasses_no_assist = int(
        ((ccp["keyPass_num"] == 1) & (ccp["assist_num"] != 1)).sum()
    )

    if col == hcol:
        ax.text(
            105,
            -3.5,
            "Violeta = Pase clave\nVerde = Asistencia",
            color=hcol,
            size=15,
            ha="right",
            va="center"
        )

        ax.text(
            52.5,
            70,
            f"Chances creadas = {cc}",
            color=col,
            fontsize=15,
            ha="center",
            va="center"
        )

        ax.set_title(
            f"{hteamName}\nZonas de creación de chances",
            color=line_color,
            fontsize=25,
            fontweight="bold",
            path_effects=path_eff
        )

    else:
        ax.text(
            105,
            71.5,
            "Violeta = Pase clave\nVerde = Asistencia",
            color=acol,
            size=15,
            ha="left",
            va="center"
        )

        ax.text(
            52.5,
            -2,
            f"Chances creadas = {cc}",
            color=col,
            fontsize=15,
            ha="center",
            va="center"
        )

        ax.set_title(
            f"{ateamName}\nZonas de creación de chances",
            color=line_color,
            fontsize=25,
            fontweight="bold",
            path_effects=path_eff
        )

    return {
        "Team_Name": team_name,
        "Total_Chances_Created": cc,
        "Assists": total_assists,
        "Key_Passes_No_Assist": total_keypasses_no_assist
    }


fig, axs = plt.subplots(1, 2, figsize=(20, 10), facecolor=bg_color)

chance_creating_stats_home = Chance_creating_zone(
    axs[0],
    hteamName,
    pearl_earring_cmaph,
    hcol
)

chance_creating_stats_away = Chance_creating_zone(
    axs[1],
    ateamName,
    pearl_earring_cmapa,
    acol
)

chance_creating_stats_df = pd.DataFrame([
    chance_creating_stats_home,
    chance_creating_stats_away
])

chance_creating_stats_df
''',
    ),
    (
        60,
        r'''
def box_entry(ax):
    # Box Entry means passes or carries which has started outside the Opponent Penalty Box and ended inside the Opponent Penalty Box 
    bentry = df[((df['type']=='Pass')|(df['type']=='Carry')) & (df['outcomeType']=='Successful') & (df['endX']>=88.5) &
                 ~((df['x']>=88.5) & (df['y']>=13.6) & (df['y']<=54.6)) & (df['endY']>=13.6) & (df['endY']<=54.4) &
            (~df['qualifiers'].str.contains('CornerTaken|Freekick|ThrowIn'))]
    hbentry = bentry[bentry['teamName']==hteamName]
    abentry = bentry[bentry['teamName']==ateamName]

    hrigt = hbentry[hbentry['y']<68/3]
    hcent = hbentry[(hbentry['y']>=68/3) & (hbentry['y']<=136/3)]
    hleft = hbentry[hbentry['y']>136/3]

    arigt = abentry[(abentry['y']<68/3)]
    acent = abentry[(abentry['y']>=68/3) & (abentry['y']<=136/3)]
    aleft = abentry[(abentry['y']>136/3)]

    pitch = Pitch(pitch_type='uefa', line_color=line_color, corner_arcs=True, line_zorder=2, pitch_color=bg_color, linewidth=2)
    pitch.draw(ax=ax)
    ax.set_xlim(-0.5, 105.5)
    ax.set_ylim(-0.5, 68.5)

    for index, row in bentry.iterrows():
        if row['teamName'] == ateamName:
            color = acol
            x, y, endX, endY = row['x'], row['y'], row['endX'], row['endY']
        elif row['teamName'] == hteamName:
            color = hcol
            x, y, endX, endY = 105 - row['x'], 68 - row['y'], 105 - row['endX'], 68 - row['endY']
        else:
            continue  # Skip rows that don't match either team name

        if row['type'] == 'Pass':
            pitch.lines(x, y, endX, endY, lw=3.5, comet=True, color=color, ax=ax, alpha=0.5)
            pitch.scatter(endX, endY, s=35, edgecolor=color, linewidth=1, color=bg_color, zorder=2, ax=ax)
        elif row['type'] == 'Carry':
            arrow = patches.FancyArrowPatch((x, y), (endX, endY), arrowstyle='->', color=color, zorder=4, mutation_scale=20, 
                                            alpha=1, linewidth=2, linestyle='--')
            ax.add_patch(arrow)

    
    ax.text(0, 69, f'{hteamName}\nLlegadas al área: {len(hbentry)}', color=hcol, fontsize=25, fontweight='bold', ha='left', va='bottom')
    ax.text(105, 69, f'{ateamName}\nLlegadas al área: {len(abentry)}', color=acol, fontsize=25, fontweight='bold', ha='right', va='bottom')

    ax.scatter(46, 6, s=2000, marker='s', color=hcol, zorder=3)
    ax.scatter(46, 34, s=2000, marker='s', color=hcol, zorder=3)
    ax.scatter(46, 62, s=2000, marker='s', color=hcol, zorder=3)
    ax.text(46, 6, f'{len(hleft)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')
    ax.text(46, 34, f'{len(hcent)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')
    ax.text(46, 62, f'{len(hrigt)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')

    ax.scatter(59.5, 6, s=2000, marker='s', color=acol, zorder=3)
    ax.scatter(59.5, 34, s=2000, marker='s', color=acol, zorder=3)
    ax.scatter(59.5, 62, s=2000, marker='s', color=acol, zorder=3)
    ax.text(59.5, 6, f'{len(arigt)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')
    ax.text(59.5, 34, f'{len(acent)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')
    ax.text(59.5, 62, f'{len(aleft)}', fontsize=30, fontweight='bold', color=bg_color, ha='center', va='center')

    home_data = {
        'Team_Name': hteamName,
        'Total_Box_Entries': len(hbentry),
        'Box_Entry_From_Left': len(hleft),
        'Box_Entry_From_Center': len(hcent),
        'Box_Entry_From_Right': len(hrigt)
    }
    
    away_data = {
        'Team_Name': ateamName,
        'Total_Box_Entries': len(abentry),
        'Box_Entry_From_Left': len(aleft),
        'Box_Entry_From_Center': len(acent),
        'Box_Entry_From_Right': len(arigt)
    }
    
    return [home_data, away_data]

fig,ax=plt.subplots(figsize=(10,10), facecolor=bg_color)
box_entry_stats = box_entry(ax)
box_entry_stats_df = pd.DataFrame(box_entry_stats)
''',
    ),
    (
        61,
        r'''
def Crosses(ax):
    pitch = Pitch(pitch_type='uefa', corner_arcs=True, pitch_color=bg_color, line_color=line_color, linewidth=2)
    pitch.draw(ax=ax)
    ax.set_ylim(-0.5,68.5)
    ax.set_xlim(-0.5,105.5)

    home_cross = df[(df['teamName']==hteamName) & (df['type']=='Pass') & (df['qualifiers'].str.contains('Cross')) & (~df['qualifiers'].str.contains('Corner'))]
    away_cross = df[(df['teamName']==ateamName) & (df['type']=='Pass') & (df['qualifiers'].str.contains('Cross')) & (~df['qualifiers'].str.contains('Corner'))]

    hsuc = 0
    hunsuc = 0
    asuc = 0
    aunsuc = 0

    # iterating through each pass and coloring according to successful or not
    for index, row in home_cross.iterrows():
        if row['outcomeType'] == 'Successful':
            arrow = patches.FancyArrowPatch((105-row['x'], 68-row['y']), (105-row['endX'], 68-row['endY']), arrowstyle='->', mutation_scale=15, color=hcol, linewidth=1.5, alpha=1)
            ax.add_patch(arrow)
            hsuc += 1
        else:
            arrow = patches.FancyArrowPatch((105-row['x'], 68-row['y']), (105-row['endX'], 68-row['endY']), arrowstyle='->', mutation_scale=10, color=line_color, linewidth=1.5, alpha=.25)
            ax.add_patch(arrow)
            hunsuc += 1

    for index, row in away_cross.iterrows():
        if row['outcomeType'] == 'Successful':
            arrow = patches.FancyArrowPatch((row['x'], row['y']), (row['endX'], row['endY']), arrowstyle='->', mutation_scale=15, color=acol, linewidth=1.5, alpha=1)
            ax.add_patch(arrow)
            asuc += 1
        else:
            arrow = patches.FancyArrowPatch((row['x'], row['y']), (row['endX'], row['endY']), arrowstyle='->', mutation_scale=10, color=line_color, linewidth=1.5, alpha=.25)
            ax.add_patch(arrow)
            aunsuc += 1

    # Headlines and other texts
    home_left = len(home_cross[home_cross['y']>=34])
    home_right = len(home_cross[home_cross['y']<34])
    away_left = len(away_cross[away_cross['y']>=34])
    away_right = len(away_cross[away_cross['y']<34])

    ax.text(51, 2, f"Centros desde\nbanda izquierda: {home_left}", color=hcol, fontsize=15, va='bottom', ha='right')
    ax.text(51, 66, f"Centros desde\nbanda derecha: {home_right}", color=hcol, fontsize=15, va='top', ha='right')
    ax.text(54, 66, f"Centros desde\nbanda izquierda: {away_left}", color=acol, fontsize=15, va='top', ha='left')
    ax.text(54, 2, f"Centros desde\nbanda derecha: {away_right}", color=acol, fontsize=15, va='bottom', ha='left')

    ax.text(0,-2, f"Completados: {hsuc}", color=hcol, fontsize=20, ha='left', va='top')
    ax.text(0,-5.5, f"No completados: {hunsuc}", color=line_color, fontsize=20, ha='left', va='top')
    ax.text(105,-2, f"Completados: {asuc}", color=acol, fontsize=20, ha='right', va='top')
    ax.text(105,-5.5, f"No completados: {aunsuc}", color=line_color, fontsize=20, ha='right', va='top')

    ax.text(0, 70, f"{hteamName}\n<---Centros", color=hcol, size=25, ha='left', fontweight='bold')
    ax.text(105, 70, f"{ateamName}\nCentros--->", color=acol, size=25, ha='right', fontweight='bold')

    home_data = {
        'Team_Name': hteamName,
        'Total_Cross': hsuc + hunsuc,
        'Successful_Cross': hsuc,
        'Unsuccessful_Cross': hunsuc,
        'Cross_From_LeftWing': home_left,
        'Cross_From_RightWing': home_right
    }
    
    away_data = {
        'Team_Name': ateamName,
        'Total_Cross': asuc + aunsuc,
        'Successful_Cross': asuc,
        'Unsuccessful_Cross': aunsuc,
        'Cross_From_LeftWing': away_left,
        'Cross_From_RightWing': away_right
    }
    
    return [home_data, away_data]

fig,ax=plt.subplots(figsize=(10,10), facecolor=bg_color)
cross_stats = Crosses(ax)
cross_stats_df = pd.DataFrame(cross_stats)
''',
    ),
    (
        62,
        r'''
def HighTO(ax):
    pitch = Pitch(pitch_type='uefa', corner_arcs=True, pitch_color=bg_color, line_color=line_color, linewidth=2)
    pitch.draw(ax=ax)
    ax.set_ylim(-0.5,68.5)
    ax.set_xlim(-0.5,105.5)

    # High Turnover means any sequence which starts in open play and within 40 metres of the opponent's goal 
    highTO = df
    highTO['Distance'] = ((highTO['x'] - 105)**2 + (highTO['y'] - 34)**2)**0.5

    # HTO which led to Goal for away team
    agoal_count = 0
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == ateamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            possession_id = highTO.loc[i, 'possession_id']
            
            # Check the following rows within the same possession
            j = i + 1
            while j < len(highTO) and highTO.loc[j, 'possession_id'] == possession_id and highTO.loc[j, 'teamName']==ateamName:
                if highTO.loc[j, 'type'] == 'Goal' and highTO.loc[j, 'teamName']==ateamName:
                    ax.scatter(highTO.loc[i, 'x'],highTO.loc[i, 'y'], s=600, marker='*', color='green', edgecolor='k', zorder=3)
                    agoal_count += 1
                    break
                j += 1

    # HTO which led to Shot for away team
    ashot_count = 0
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == ateamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            possession_id = highTO.loc[i, 'possession_id']
            
            # Check the following rows within the same possession
            j = i + 1
            while j < len(highTO) and highTO.loc[j, 'possession_id'] == possession_id and highTO.loc[j, 'teamName']==ateamName:
                if ('Shot' in highTO.loc[j, 'type']) and (highTO.loc[j, 'teamName']==ateamName):
                    ax.scatter(highTO.loc[i, 'x'],highTO.loc[i, 'y'], s=150, color=acol, edgecolor=bg_color, zorder=2)
                    ashot_count += 1
                    break
                j += 1
    
    # other HTO for away team
    aht_count = 0
    p_list = []
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == ateamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            # Check the following rows
            j = i + 1
            if ((highTO.loc[j, 'teamName']==ateamName) and
                (highTO.loc[j, 'type']!='Dispossessed') and (highTO.loc[j, 'type']!='OffsidePass')):
                ax.scatter(highTO.loc[i, 'x'],highTO.loc[i, 'y'], s=100, color='None', edgecolor=acol)
                aht_count += 1
                p_list.append(highTO.loc[i, 'teamName'])

    # HTO which led to Goal for home team
    hgoal_count = 0
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == hteamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            possession_id = highTO.loc[i, 'possession_id']
            
            # Check the following rows within the same possession
            j = i + 1
            while j < len(highTO) and highTO.loc[j, 'possession_id'] == possession_id and highTO.loc[j, 'teamName']==hteamName:
                if highTO.loc[j, 'type'] == 'Goal' and highTO.loc[j, 'teamName']==hteamName:
                    ax.scatter(105-highTO.loc[i, 'x'],68-highTO.loc[i, 'y'], s=600, marker='*', color='green', edgecolor='k', zorder=3)
                    hgoal_count += 1
                    break
                j += 1

    # HTO which led to Shot for home team
    hshot_count = 0
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == hteamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            possession_id = highTO.loc[i, 'possession_id']
            
            # Check the following rows within the same possession
            j = i + 1
            while j < len(highTO) and highTO.loc[j, 'possession_id'] == possession_id and highTO.loc[j, 'teamName']==hteamName:
                if ('Shot' in highTO.loc[j, 'type']) and (highTO.loc[j, 'teamName']==hteamName):
                    ax.scatter(105-highTO.loc[i, 'x'],68-highTO.loc[i, 'y'], s=150, color=hcol, edgecolor=bg_color, zorder=2)
                    hshot_count += 1
                    break
                j += 1

    # other HTO for home team
    hht_count = 0
    p_list = []
    # Iterate through the DataFrame
    for i in range(len(highTO)):
        if ((highTO.loc[i, 'type'] in ['BallRecovery', 'Interception']) and 
            (highTO.loc[i, 'teamName'] == hteamName) and 
            (highTO.loc[i, 'Distance'] <= 40)):
            
            # Check the following rows
            j = i + 1
            if ((highTO.loc[j, 'teamName']==hteamName) and
                (highTO.loc[j, 'type']!='Dispossessed') and (highTO.loc[j, 'type']!='OffsidePass')):
                ax.scatter(105-highTO.loc[i, 'x'],68-highTO.loc[i, 'y'], s=100, color='None', edgecolor=hcol)
                hht_count += 1
                p_list.append(highTO.loc[i, 'teamName'])

    # Plotting the half circle
    left_circle = plt.Circle((0,34), 40, color=hcol, fill=True, alpha=0.25, linestyle='dashed')
    ax.add_artist(left_circle)
    right_circle = plt.Circle((105,34), 40, color=acol, fill=True, alpha=0.25, linestyle='dashed')
    ax.add_artist(right_circle)
    # Set the aspect ratio to be equal
    ax.set_aspect('equal', adjustable='box')
    # Headlines and other texts
    ax.text(0, 70, f"{hteamName}\nContragolpes: {hht_count}", color=hcol, size=25, ha='left', fontweight='bold')
    ax.text(105, 70, f"{ateamName}\nContragolpes: {aht_count}", color=acol, size=25, ha='right', fontweight='bold')
    ax.text(0,  -3, '<---Dirección de ataque', color=hcol, fontsize=13, ha='left', va='center')
    ax.text(105,-3, 'Dirección de ataque--->', color=acol, fontsize=13, ha='right', va='center')

    home_data = {
        'Team_Name': hteamName,
        'Total_High_Turnovers': hht_count,
        'Shot_Ending_High_Turnovers': hshot_count,
        'Goal_Ending_High_Turnovers': hgoal_count,
        'Opponent_Team_Name': ateamName
    }
    
    away_data = {
        'Team_Name': ateamName,
        'Total_High_Turnovers': aht_count,
        'Shot_Ending_High_Turnovers': ashot_count,
        'Goal_Ending_High_Turnovers': agoal_count,
        'Opponent_Team_Name': hteamName
    }
    
    return [home_data, away_data]

fig,ax=plt.subplots(figsize=(10,10), facecolor=bg_color)
high_turnover_stats = HighTO(ax)
high_turnover_stats_df = pd.DataFrame(high_turnover_stats)

high_turnover_stats_df
''',
    ),
    (
        63,
        r'''
def plot_congestion(ax):
    # Comparing open play touches of both teams in each zones of the pitch, if more than 55% touches for a team it will be coloured of that team, otherwise gray to represent contested
    pcmap = LinearSegmentedColormap.from_list("Pearl Earring - 10 colors",  [acol, 'gray', hcol], N=20)
    df1 = df[(df['teamName']==hteamName) & (df['isTouch']==1) & (~df['qualifiers'].str.contains('CornerTaken|Freekick|ThrowIn'))]
    df2 = df[(df['teamName']==ateamName) & (df['isTouch']==1) & (~df['qualifiers'].str.contains('CornerTaken|Freekick|ThrowIn'))]
    df2['x'] = 105-df2['x']
    df2['y'] =  68-df2['y']
    pitch = Pitch(pitch_type='uefa', corner_arcs=True, pitch_color=bg_color, line_color=line_color, linewidth=2, line_zorder=6)
    pitch.draw(ax=ax)
    ax.set_ylim(-0.5,68.5)
    ax.set_xlim(-0.5,105.5)

    bin_statistic1 = pitch.bin_statistic(df1.x, df1.y, bins=(6,5), statistic='count', normalize=False)
    bin_statistic2 = pitch.bin_statistic(df2.x, df2.y, bins=(6,5), statistic='count', normalize=False)

    # Assuming 'cx' and 'cy' are as follows:
    cx = np.array([[ 8.75, 26.25, 43.75, 61.25, 78.75, 96.25],
               [ 8.75, 26.25, 43.75, 61.25, 78.75, 96.25],
               [ 8.75, 26.25, 43.75, 61.25, 78.75, 96.25],
               [ 8.75, 26.25, 43.75, 61.25, 78.75, 96.25],
               [ 8.75, 26.25, 43.75, 61.25, 78.75, 96.25]])

    cy = np.array([[61.2, 61.2, 61.2, 61.2, 61.2, 61.2],
               [47.6, 47.6, 47.6, 47.6, 47.6, 47.6],
               [34.0, 34.0, 34.0, 34.0, 34.0, 34.0],
               [20.4, 20.4, 20.4, 20.4, 20.4, 20.4],
               [ 6.8,  6.8,  6.8,  6.8,  6.8,  6.8]])

    # Flatten the arrays
    cx_flat = cx.flatten()
    cy_flat = cy.flatten()

    # Create a DataFrame
    df_cong = pd.DataFrame({'cx': cx_flat, 'cy': cy_flat})

    hd_values = []
    # Loop through the 2D arrays
    for i in range(bin_statistic1['statistic'].shape[0]):
        for j in range(bin_statistic1['statistic'].shape[1]):
            stat1 = bin_statistic1['statistic'][i, j]
            stat2 = bin_statistic2['statistic'][i, j]
        
            if (stat1 / (stat1 + stat2)) > 0.55:
                hd_values.append(1)
            elif (stat1 / (stat1 + stat2)) < 0.45:
                hd_values.append(0)
            else:
                hd_values.append(0.5)

    df_cong['hd']=hd_values
    bin_stat = pitch.bin_statistic(df_cong.cx, df_cong.cy, bins=(6,5), values=df_cong['hd'], statistic='sum', normalize=False)
    pitch.heatmap(bin_stat, ax=ax, cmap=pcmap, edgecolors='#000000', lw=0, zorder=3, alpha=0.85)

    ax_text(52.5, 71, s=f"<{hteamName}>  |  Disputado  |  <{ateamName}>", highlight_textprops=[{'color':hcol}, {'color':acol}],
            color='gray', fontsize=18, ha='center', va='center', ax=ax)
    ax.set_title("Zona de dominancia del equipo", color=line_color, fontsize=30, fontweight='bold', y=1.075)
    ax.text(0,  -3, 'Dirección de ataque--->', color=hcol, fontsize=13, ha='left', va='center')
    ax.text(105,-3, '<---Dirección de ataque', color=acol, fontsize=13, ha='right', va='center')

    ax.vlines(1*(105/6), ymin=0, ymax=68, color=bg_color, lw=2, ls='--', zorder=5)
    ax.vlines(2*(105/6), ymin=0, ymax=68, color=bg_color, lw=2, ls='--', zorder=5)
    ax.vlines(3*(105/6), ymin=0, ymax=68, color=bg_color, lw=2, ls='--', zorder=5)
    ax.vlines(4*(105/6), ymin=0, ymax=68, color=bg_color, lw=2, ls='--', zorder=5)
    ax.vlines(5*(105/6), ymin=0, ymax=68, color=bg_color, lw=2, ls='--', zorder=5)

    ax.hlines(1*(68/5), xmin=0, xmax=105, color=bg_color, lw=2, ls='--', zorder=5)
    ax.hlines(2*(68/5), xmin=0, xmax=105, color=bg_color, lw=2, ls='--', zorder=5)
    ax.hlines(3*(68/5), xmin=0, xmax=105, color=bg_color, lw=2, ls='--', zorder=5)
    ax.hlines(4*(68/5), xmin=0, xmax=105, color=bg_color, lw=2, ls='--', zorder=5)
    
    return

fig,ax=plt.subplots(figsize=(10,10), facecolor=bg_color)
plot_congestion(ax)
''',
    ),
    (
        64,
        r'''
fig, axs = plt.subplots(4,3, figsize=(35,35), facecolor=bg_color)

pass_network_stats_home = pass_network_visualization(axs[0,0], home_passes_between_df, home_average_locs_and_count_df, hcol, hteamName)
shooting_stats = plot_shotmap(axs[0,1])
pass_network_stats_away = pass_network_visualization(axs[0,2], away_passes_between_df, away_average_locs_and_count_df, acol, ateamName)

defensive_block_stats_home = defensive_block(axs[1,0],defensive_home_average_locs_and_count_df, hteamName,hcol,defensive_actions_df)
goalkeeping_stats = plot_goalPost(axs[1,1])
defensive_block_stats_away = defensive_block(axs[1,2], defensive_away_average_locs_and_count_df, ateamName, acol, defensive_actions_df)

Progressvie_Passes_Stats_home = draw_progressive_pass_map(axs[2,0], hteamName, hcol)
xT_stats = plot_Momentum(axs[2,1])
Progressvie_Passes_Stats_away = draw_progressive_pass_map(axs[2,2], ateamName, acol)

Progressvie_Carries_Stats_home = draw_progressive_carry_map(axs[3,0], hteamName, hcol)
general_match_stats = plotting_match_stats(axs[3,1])
Progressvie_Carries_Stats_away = draw_progressive_carry_map(axs[3,2], ateamName, acol)

# Heading
highlight_text = [{'color':hcol}, {'color':acol}]
fig_text(0.5, 0.98, f"<{hteamName} {hgoal_count}> - <{agoal_count} {ateamName}>", color=line_color, fontsize=70, fontweight='bold',
            highlight_textprops=highlight_text, ha='center', va='center', ax=fig)

# Subtitles
fig.text(0.5, 0.95, f"Fecha 16 , LPF 2026 | Reporte Post-Match", color=line_color, fontsize=30, ha='center', va='center')
fig.text(0.5, 0.93, f"Data: Opta | Código original (whoscored.com) de @adnaaan433 | Adaptación a Scoresway por @robaboian_", color=line_color, fontsize=22.5, ha='center', va='center')

fig.text(0.125,0.1, 'Dirección de ataque ------->', color=hcol, fontsize=25, ha='left', va='center')
fig.text(0.9,0.1, '<------- Dirección de ataque', color=acol, fontsize=25, ha='right', va='center')

# Plotting Team's Logo
# Here I have choosen a very complicated process, you may know better how to plot easily
# I download any team's png logo from google and then save that file as .html, then open that html file and copy paste the url here


# Saving the final figure
fig.savefig(f"Match_Report_1.png", bbox_inches='tight')
''',
    ),
    (
        65,
        r'''
fig, axs = plt.subplots(4,3, figsize=(35,35), facecolor=bg_color)

final_third_entry_stats_home = Final_third_entry(axs[0,0], hteamName, hcol)
box_entry_stats = box_entry(axs[0,1])
final_third_entry_stats_away = Final_third_entry(axs[0,2], ateamName, acol)

zonal_passing_stats_home = zone14hs(axs[1,0], hteamName, hcol)
cross_stats = Crosses(axs[1,1])
zonal_passing_stats_away = zone14hs(axs[1,2], ateamName, acol)

Pass_end_zone(axs[2,0], hteamName, pearl_earring_cmaph)
high_turnover_stats = HighTO(axs[2,1])
Pass_end_zone(axs[2,2], ateamName, pearl_earring_cmapa)

chance_creating_stats_home = Chance_creating_zone(axs[3,0], hteamName, pearl_earring_cmaph, hcol)
plot_congestion(axs[3,1])
chance_creating_stats_away = Chance_creating_zone(axs[3,2], ateamName, pearl_earring_cmapa, acol)

# Heading
highlight_text = [{'color':hcol}, {'color':acol}]
fig_text(0.5, 0.98, f"<{hteamName} {hgoal_count}> - <{agoal_count} {ateamName}>", color=line_color, fontsize=70, fontweight='bold',
            highlight_textprops=highlight_text, ha='center', va='center', ax=fig)

# Subtitles
fig.text(0.5, 0.95, f"Fecha 16 , LPF 2026 | Reporte Post-Match 2", color=line_color, fontsize=30, ha='center', va='center')
fig.text(0.5, 0.93, f"Data: Opta | Código original (whoscored.com) de @adnaaan433 | Adaptación a Scoresway por @robaboian_", color=line_color, fontsize=22.5, ha='center', va='center')

fig.text(0.125,0.1, 'Dirección de ataque ------->', color=hcol, fontsize=25, ha='left', va='center')
fig.text(0.9,0.1, '<------- Dirección de ataque', color=acol, fontsize=25, ha='right', va='center')


# Saving the final Figure
fig.savefig("Match_Report.png", bbox_inches='tight')
''',
    ),
]
EMBEDDED_RESOURCE_FILES_B64: dict[str, str] = {
    'typeId.xlsx': (
    "UEsDBBQABgAIAAAAIQCnDOt5aAEAAA0FAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACslMtuwjAQRfeV+g+Rt1Vi6KKqKgKLPpYt"
    "UukHuPaEWPglz0Dh7+sYqKoqBSHYxEo8c8/NxDejydqaYgURtXc1G1YDVoCTXmk3r9nH7KW8ZwWScEoY76BmG0A2"
    "GV9fjWabAFikboc1a4nCA+coW7ACKx/ApZ3GRyso3cY5D0IuxBz47WBwx6V3BI5K6jTYePQEjVgaKp7X6fHWSQSD"
    "rHjcFnasmokQjJaCklO+cuoPpdwRqtSZa7DVAW+SDcZ7Cd3O/4Bd31saTdQKiqmI9CpsssHXhn/5uPj0flEdFulx"
    "6ZtGS1BeLm2aQIUhglDYApA1VV4rK7Tb+z7Az8XI8zK8sJHu/bLwER+UvjfwfD3fQpY5AkTaGMBLjz2LHiO3IoJ6"
    "p5iScXEDv7UP+UjnZhp9wJSgCKdPYR+RrrsMSQgiafgJSd9h+yGm9J09dujyrUCdypZLJG/Pxm9leuA8/8zG3wAA"
    "AP//AwBQSwMEFAAGAAgAAAAhABNevmUCAQAA3wIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1LAzEQhu+C/yHMvTvbKiLSbC9F6E1k/QEx"
    "mf1gN5mQpLr990ZBdKG2Hnqcr3eeeZn1ZrKjeKMQe3YSlkUJgpxm07tWwkv9uLgHEZNyRo3sSMKBImyq66v1M40q"
    "5aHY9T6KrOKihC4l/4AYdUdWxYI9uVxpOFiVchha9EoPqiVcleUdht8aUM00xc5ICDtzA6I++Lz5vDY3Ta9py3pv"
    "yaUjK5CmRM6QWfiQ2ULq8zWiVqGlJMGwfsrpiMr7ImMDHida/Z/o72vRUlJGJYWaA53m+ew4BbS8pEVzE3/cmUZ8"
    "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"
),
    'qualifiers.csv': (
    "cXVhbGlmaWVySWQ7UVVBTElGSUVSIE5BTUU7cXVhbGlmaWVySWRfbnVtDQoxO0xvbmdiYWxsOzENCjI7Q3Jvc3M7"
    "Mg0KMztIZWFkUGFzczszDQo0O1Rocm91Z2hiYWxsOzQNCjU7RnJlZWtpY2tUYWtlbjs1DQo2O0Nvcm5lclRha2Vu"
    "OzYNCjc7UGxheWVyQ2F1Z2h0T2Zmc2lkZTs3DQo4O0dvYWxEaXNhbGxvd2VkOzgNCjk7UGVuYWx0eTs5DQoxMDtI"
    "YW5kYmFsbDsxMA0KMTE7Ni1zZWNvbmRzIHZpb2xhdGlvbjsxMQ0KMTI7RGFuZ2Vyb3VzIHBsYXk7MTINCjEzO0Zv"
    "dWw7MTMNCjE0O0xhc3RNYW47MTQNCjE1O0hlYWQ7MTUNCjE2O1NtYWxsQm94Q2VudHJlOzE2DQoxNztCb3hDZW50"
    "cmU7MTcNCjE4O091dE9mQm94Q2VudHJlOzE4DQoxOTtUaGlydHlGaXZlUGx1c0NlbnRyZTsxOQ0KMjA7UmlnaHRG"
    "b290OzIwDQoyMTtPdGhlckJvZHlQYXJ0OzIxDQoyMjtSZWd1bGFyUGxheTsyMg0KMjM7RmFzdEJyZWFrOzIzDQoy"
    "NDtTZXRQaWVjZTsyNA0KMjU7RnJvbUNvcm5lcjsyNQ0KMjY7RGlyZWN0RnJlZWtpY2s7MjYNCjI4O093biBHb2Fs"
    "OzI4DQoyOTtBc3Npc3RlZDsyOQ0KMzA7SW52b2x2ZWRQbGF5ZXJzOzMwDQozMTtZZWxsb3c7MzENCjMyO1NlY29u"
    "ZFllbGxvdzszMg0KMzM7UmVkOzMzDQozNDtSZWZlcmVlIGFidXNlOzM0DQozNTtBcmd1bWVudDszNQ0KMzY7Vmlv"
    "bGVudCBDb25kdWN0OzM2DQozNztUaW1lIHdhc3Rpbmc7MzcNCjM4O0V4Y2Vzc2l2ZSBjZWxlYnJhdGlvbjszOA0K"
    "Mzk7Q3Jvd2QgaW50ZXJhY3Rpb247MzkNCjQwO090aGVyIHJlYXNvbjs0MA0KNDE7SW5qdXJ5OzQxDQo0MjtUYWN0"
    "aWNhbDs0Mg0KNDM7RGVsZXRlZCBldmVudDs0Mw0KNDQ7UGxheWVyUG9zaXRpb247NDQNCjQ1O1RlbXBlcmF0dXJl"
    "OzQ1DQo0NjtDb25kaXRpb25zOzQ2DQo0NztGaWVsZCBQaXRjaDs0Nw0KNDg7TGlnaHRpbmdzOzQ4DQo0OTtBdHRl"
    "bmRhbmNlIGZpZ3VyZTs0OQ0KNTA7T2ZmaWNpYWwgcG9zaXRpb247NTANCjUxO09mZmljaWFsIElEOzUxDQo1MjtQ"
    "b3NzZXNzaW9uIHRpbWU7NTINCjUzO0luanVyZWQgcGxheWVyIElEOzUzDQo1NDtFbmQgY2F1c2U7NTQNCjU1O1Jl"
    "bGF0ZWRFdmVudElkOzU1DQo1Njtab25lOzU2DQo1NztFbmQgdHlwZTs1Nw0KNTg7VGVtcCBzdG9wIHN0YXR1czs1"
    "OA0KNTk7SmVyc2V5TnVtYmVyOzU5DQo2MDtTbWFsbEJveFJpZ2h0OzYwDQo2MTtTbWFsbEJveExlZnQ7NjENCjYy"
    "O0JveCAtIERlZXAgUmlnaHQ7NjINCjYzO0JveFJpZ2h0OzYzDQo2NDtCb3hMZWZ0OzY0DQo2NTtEZWVwQm94TGVm"
    "dDs2NQ0KNjY7T3V0IG9mIGJveCAtIERlZXAgUmlnaHQ7NjYNCjY3O091dCBvZiBib3ggLSBSaWdodDs2Nw0KNjg7"
    "T3V0IG9mIGJveCAtIExlZnQ7NjgNCjY5O091dE9mQm94RGVlcExlZnQ7NjkNCjcwOzM1KyBSaWdodDs3MA0KNzE7"
    "MzUrIExlZnQ7NzENCjcyO0xlZnRGb290OzcyDQo3MztNaXNzTGVmdDs3Mw0KNzQ7TWlzc0hpZ2g7NzQNCjc1O01p"
    "c3NSaWdodDs3NQ0KNzY7TG93TGVmdDs3Ng0KNzc7SGlnaExlZnQ7NzcNCjc4O0xvd0NlbnRyZTs3OA0KNzk7SGln"
    "aENlbnRyZTs3OQ0KODA7TG93UmlnaHQ7ODANCjgxO0hpZ2hSaWdodDs4MQ0KODI7QmxvY2tlZDs4Mg0KODM7Q2xv"
    "c2UgTGVmdDs4Mw0KODQ7Q2xvc2UgUmlnaHQ7ODQNCjg1O0Nsb3NlIEhpZ2g7ODUNCjg2O0Nsb3NlIExlZnQgYW5k"
    "IEhpZ2g7ODYNCjg3O0Nsb3NlIFJpZ2h0IGFuZCBIaWdoOzg3DQo4ODtIaWdoQ2xhaW07ODgNCjg5O09uZU9uT25l"
    "Ozg5DQo5MDtEZWZsZWN0ZWQgc2F2ZTs5MA0KOTE7RGl2ZSBhbmQgZGVmbGVjdDs5MQ0KOTI7Q2F0Y2g7OTINCjkz"
    "O0RpdmUgYW5kIGNhdGNoOzkzDQo5NDtPdXRmaWVsZGVyQmxvY2s7OTQNCjk1O0JhY2sgcGFzczs5NQ0KOTY7Q29y"
    "bmVyIHNpdHVhdGlvbjs5Ng0KOTc7RGlyZWN0IGZyZWU7OTcNCjk4O1BpdGNoIFggQ29vcmRpbmF0ZTs5OA0KOTk7"
    "UGl0Y2ggWSBDb29yZGluYXRlOzk5DQoxMDA7U2l4WWFyZEJsb2NrOzEwMA0KMTAxO1NhdmVkT2ZmbGluZTsxMDEN"
    "CjEwMjtHb2FsTW91dGhZOzEwMg0KMTAzO0dvYWxNb3V0aFo7MTAzDQoxMDQ7QXR0ZW1wdCBQb3NpdGlvbiBYIENv"
    "b3JkaW5hdGU7MTA0DQoxMDU7QXR0ZW1wdCBQb3NpdGlvbiBZIENvb3JkaW5hdGU7MTA1DQoxMDY7QXR0YWNraW5n"
    "IFBhc3M7MTA2DQoxMDc7VGhyb3dJbjsxMDcNCjEwODtWb2xsZXk7MTA4DQoxMDk7T3ZlcmhlYWQ7MTA5DQoxMTA7"
    "SGFsZiBWb2xsZXk7MTEwDQoxMTE7RGl2aW5nIEhlYWRlcjsxMTENCjExMjtTY3JhbWJsZTsxMTINCjExMztTdHJv"
    "bmc7MTEzDQoxMTQ7V2VhazsxMTQNCjExNTtSaXNpbmc7MTE1DQoxMTY7RGlwcGluZzsxMTYNCjExNztMb2I7MTE3"
    "DQoxMTg7T25lIEJvdW5jZTsxMTgNCjExOTtGZXcgQm91bmNlczsxMTkNCjEyMDtTd2VydmUgTGVmdDsxMjANCjEy"
    "MTtTd2VydmUgUmlnaHQ7MTIxDQoxMjI7U3dlcnZlIE1vdmluZzsxMjINCjEyMztLZWVwZXJUaHJvdzsxMjMNCjEy"
    "NDtHb2FsS2ljazsxMjQNCjEyNTtGcmVlIEtpY2sgUG9zaXRpb24gWCBDb29yZGluYXRlOzEyNQ0KMTI2O0ZyZWUg"
    "S2ljayBQb3NpdGlvbiBZIENvb3JkaW5hdGU7MTI2DQoxMjc7RGlyZWN0aW9uIG9mIFBsYXk7MTI3DQoxMjg7UHVu"
    "Y2g7MTI4DQoxMjk7VGVuIE1pbnV0ZSBQb3NzZXNzaW9uOzEyOQ0KMTMwO1RlYW1Gb3JtYXRpb247MTMwDQoxMzE7"
    "VGVhbVBsYXllckZvcm1hdGlvbjsxMzENCjEzMjtTaW11bGF0aW9uOzEzMg0KMTMzO0RlZmxlY3Rpb247MTMzDQox"
    "MzQ7RmFyIFdpZGUgTGVmdDsxMzQNCjEzNTtGYXIgV2lkZSBSaWdodDsxMzUNCjEzNjtLZWVwZXIgVG91Y2hlZDsx"
    "MzYNCjEzNztLZWVwZXIgU2F2ZWQ7MTM3DQoxMzg7SGl0IFdvb2R3b3JrOzEzOA0KMTM5O093biBQbGF5ZXI7MTM5"
    "DQoxNDA7UGFzc0VuZFg7MTQwDQoxNDE7UGFzc0VuZFk7MTQxDQoxNDI7RmxhZyB0byBDaGVja2VyOzE0Mg0KMTQz"
    "O1N0YXIgUmF0aW5nOzE0Mw0KMTQ0O0RlbGV0ZWQgRXZlbnQgVHlwZTsxNDQNCjE0NTtGb3JtYXRpb25TbG90OzE0"
    "NQ0KMTQ2O0Jsb2NrZWRYOzE0Ng0KMTQ3O0Jsb2NrZWRZOzE0Nw0KMTQ4O0RhbmdlcjsxNDgNCjE0OTtJbnNpZGU7"
    "MTQ5DQoxNTA7T3V0c2lkZTsxNTANCjE1MTtTaG9ydDsxNTENCjE1MjtEaXJlY3Q7MTUyDQoxNTM7Tm90IHBhc3Qg"
    "Z29hbCBsaW5lOzE1Mw0KMTU0O0ludGVudGlvbmFsQXNzaXN0OzE1NA0KMTU1O0NoaXBwZWQ7MTU1DQoxNTY7TGF5"
    "T2ZmOzE1Ng0KMTU3O0xhdW5jaDsxNTcNCjE1ODtQZXJzaXN0ZW50IEluZnJpbmdlbWVudDsxNTgNCjE1OTtGb3Vs"
    "IGFuZCBBYnVzaXZlIExhbmd1YWdlOzE1OQ0KMTYwO1Rocm93aW5TZXRQaWVjZTsxNjANCjE2MTtFbmNyb2FjaG1l"
    "bnQ7MTYxDQoxNjI7TGVhdmluZyBmaWVsZDsxNjINCjE2MztFbnRlcmluZyBmaWVsZDsxNjMNCjE2NDtTcGl0dGlu"
    "ZzsxNjQNCjE2NTtQcm9mZXNzaW9uYWwgRm91bCBMYXN0IE1hbjsxNjUNCjE2NjtQcm9mZXNzaW9uYWwgRm91bCBI"
    "YW5kYmFsbDsxNjYNCjE2NztPdXQgb2YgcGxheTsxNjcNCjE2ODtGbGljay1vbjsxNjgNCjE2OTtMZWFkaW5nVG9B"
    "dHRlbXB0OzE2OQ0KMTcwO0xlYWRpbmdUb0dvYWw7MTcwDQoxNzE7UmVzY2luZGVkIENhcmQ7MTcxDQoxNzM7UGFy"
    "cmllZFNhZmU7MTczDQoxNzQ7UGFycmllZERhbmdlcjsxNzQNCjE3NTtGaW5nZXJ0aXA7MTc1DQoxNzY7Q2F1Z2h0"
    "OzE3Ng0KMTc3O0NvbGxlY3RlZDsxNzcNCjE3ODtTdGFuZGluZ1NhdmU7MTc4DQoxNzk7RGl2aW5nU2F2ZTsxNzkN"
    "CjE4MDtTdG9vcGluZzsxODANCjE4MTtSZWFjaGluZzsxODENCjE4MjtIYW5kczsxODINCjE4MztGZWV0OzE4Mw0K"
    "MTg0O0Rpc3NlbnQ7MTg0DQoxODU7QmxvY2tlZENyb3NzOzE4NQ0KMTg2O0tlZXBlck1pc3NlZDsxODYNCjE4NztL"
    "ZWVwZXJTYXZlZDsxODcNCjE4ODtNaXNzZWQ7MTg4DQoxODk7Tm90IHZpc2libGU7MTg5DQoxOTA7RnJvbVNob3RP"
    "ZmZUYXJnZXQ7MTkwDQoxOTE7T2ZmIHRoZSBiYWxsIGZvdWw7MTkxDQoxOTI7QmxvY2sgYnkgaGFuZDsxOTINCjE5"
    "MztHb2FsIG1lYXN1cmU7MTkzDQoxOTQ7Q2FwdGFpblBsYXllcklkOzE5NA0KMTk1O1B1bGwgYmFjazsxOTUNCjE5"
    "NjtTd2l0Y2ggb2YgcGxheTsxOTYNCjE5NztUZWFtIGtpdDsxOTcNCjE5ODtHSyBob29mOzE5OA0KMTk5O0dLIGtp"
    "Y2sgZnJvbSBoYW5kczsxOTkNCjIwMDtSZWZlcmVlIHN0b3A7MjAwDQoyMDE7UmVmZXJlZSBkZWxheTsyMDENCjIw"
    "MjtXZWF0aGVyIHByb2JsZW07MjAyDQoyMDM7Q3Jvd2QgdHJvdWJsZTsyMDMNCjIwNDtGaXJlOzIwNA0KMjA1O09i"
    "amVjdCB0aHJvd24gb24gcGl0Y2g7MjA1DQoyMDY7U3BlY3RhdG9yIG9uIHBpdGNoOzIwNg0KMjA3O0F3YWl0aW5n"
    "IG9mZmljaWFsJ3MgZGVjaXNpb247MjA3DQoyMDg7UmVmZXJlZSBpbmp1cnk7MjA4DQoyMDk7R2FtZSBlbmQ7MjA5"
    "DQoyMTA7U2hvdEFzc2lzdDsyMTANCjIxMTtPdmVyUnVuOzIxMQ0KMjEyO0xlbmd0aDsyMTINCjIxMztBbmdsZTsy"
    "MTMNCjIxNDtCaWdDaGFuY2U7MjE0DQoyMTU7SW5kaXZpZHVhbFBsYXk7MjE1DQoyMTY7Mm5kIHJlbGF0ZWQgZXZl"
    "bnQgSUQ7MjE2DQoyMTc7Mm5kIGFzc2lzdGVkOzIxNw0KMjE4OzJuZCBhc3Npc3Q7MjE4DQoyMTk7UGxheWVycyBv"
    "biBib3RoIHBvc3RzOzIxOQ0KMjIwO1BsYXllciBvbiBuZWFyIHBvc3Q7MjIwDQoyMjE7UGxheWVyIG9uIGZhciBw"
    "b3N0OzIyMQ0KMjIyO05vIHBsYXllcnMgb24gcG9zdHM7MjIyDQoyMjM7SW4tc3dpbmdlcjsyMjMNCjIyNDtPdXQt"
    "c3dpbmdlcjsyMjQNCjIyNTtTdHJhaWdodDsyMjUNCjIyNjtTdXNwZW5kZWQ7MjI2DQoyMjc7UmVzdW1lOzIyNw0K"
    "MjI4O093biBzaG90IGJsb2NrZWQ7MjI4DQoyMjk7UG9zdCBtYXRjaCBjb21wbGV0ZTsyMjkNCjIzMDtHSyBYIENv"
    "b3JkaW5hdGU7MjMwDQoyMzE7R0sgWSBDb29yZGluYXRlOzIzMQ0KMjMyO1VuY2hhbGxlbmdlZDsyMzINCjIzMztP"
    "cHBvc2l0ZVJlbGF0ZWRFdmVudDsyMzMNCjIzNDtIb21lIFRlYW0gUG9zc2Vzc2lvbjsyMzQNCjIzNTtBd2F5IFRl"
    "YW0gUG9zc2Vzc2lvbjsyMzUNCjIzNjtCbG9ja2VkIHBhc3M7MjM2DQoyMzc7TG93OzIzNw0KMjM4O0ZhaXIgUGxh"
    "eTsyMzgNCjIzOTtCeSBXYWxsOzIzOQ0KMjQwO0dLIFN0YXJ0OzI0MA0KMjQxO0luZGlyZWN0RnJlZWtpY2tUYWtl"
    "bjsyNDENCjI0MjtPYnN0cnVjdGlvbjsyNDINCjI0MztVbnNwb3J0aW5nIGJlaGF2aW91cjsyNDMNCjI0NDtOb3Qg"
    "UmV0cmVhdGluZzsyNDQNCjI0NTtTZXJpb3VzIEZvdWw7MjQ1DQoyNDY7RHJpbmtzIEJyZWFrOzI0Ng0KMjQ3O09m"
    "ZnNpZGU7MjQ3DQoyNDg7R29hbCBsaW5lOzI0OA0KMjQ5O1RlbXAgU2hvdCBPbjsyNDkNCjI1MDtUZW1wIEJsb2Nr"
    "ZWQ7MjUwDQoyNTE7VGVtcCBQb3N0OzI1MQ0KMjUyO1RlbXAgTWlzc2VkOzI1Mg0KMjUzO1RlbXAgTWlzcyBOb3Qg"
    "UGFzc2VkIEdvYWwgTGluZTsyNTMNCjI1NDtGb2xsb3dzIGEgRHJpYmJsZTsyNTQNCjI1NTtPcGVuIFJvb2Y7MjU1"
    "DQoyNTY7QWlyIEh1bWlkaXR5OzI1Ng0KMjU3O0FpciBQcmVzc3VyZTsyNTcNCjI1ODtTb2xkIE91dDsyNTgNCjI1"
    "OTtDZWxzaXVzIGRlZ3JlZXM7MjU5DQoyNjA7Rmxvb2RsaWdodDsyNjANCjI2MTsxIG9uIDEgY2hpcDsyNjENCjI2"
    "MjtCYWNrIGhlZWw7MjYyDQoyNjM7RGlyZWN0IGNvcm5lcjsyNjMNCjI2NDtBZXJpYWxGb3VsOzI2NA0KMjY1O0F0"
    "dGVtcHRlZCBUYWNrbGU7MjY1DQoyNjY7UHV0IFRocm91Z2g7MjY2DQoyNjc7UmlnaHQgQXJtOzI2Nw0KMjY4O0xl"
    "ZnQgQXJtOzI2OA0KMjY5O0JvdGggQXJtczsyNjkNCjI3MDtSaWdodCBMZWc7MjcwDQoyNzE7TGVmdCBMZWc7Mjcx"
    "DQoyNzE7Qm90aCBMZWdzOzI3MQ0KMjczO0hpdCBSaWdodCBQb3N0OzI3Mw0KMjc0O0hpdCBMZWZ0IFBvc3Q7Mjc0"
    "DQoyNzU7SGl0IEJhcjsyNzUNCjI3NjtPdXQgb24gc2lkZWxpbmU7Mjc2DQoyNzc7TWludXRlczsyNzcNCjI3ODtU"
    "YXA7Mjc4DQoyNzk7S2ljayBPZmY7Mjc5DQoyODA7RmFudGFzeSBBc3Npc3QgVHlwZTsyODANCjI4MTtGYW50YXN5"
    "IEFzc2lzdGVkIEJ5OzI4MQ0KMjgyO0ZhbnRhc3kgQXNzaXN0IFRlYW07MjgyDQoyODM7Q29hY2ggSUQ7MjgzDQoy"
    "ODQ7RHVlbDsyODQNCjI4NTtEZWZlbnNpdmU7Mjg1DQoyODY7T2ZmZW5zaXZlOzI4Ng0KMjg3O092ZXItYXJtOzI4"
    "Nw0KMjg4O091dCBvZiBQbGF5IFNlY3M7Mjg4DQoyODk7RGVuaWVkIGdvYWwtc2NvcmluZyBvcHA7Mjg5DQoyOTA7"
    "Q29hY2ggdHlwZXM7MjkwDQoyOTE7T3RoZXIgQmFsbCBDb250YWN0IFR5cGU7MjkxDQoyOTI7RGV0YWlsZWQgUG9z"
    "aXRpb24gSUQ7MjkyDQoyOTM7UG9zaXRpb24gU2lkZSBJRDsyOTMNCjI5NDtTaG92ZS9QdXNoOzI5NA0KMjk1O1No"
    "aXJ0IFB1bGwvSG9sZGluZzsyOTUNCjI5NztGb2xsb3dzIFNob3QgUmVib3VuZDsyOTcNCjI5ODtGb2xsb3dzIFNo"
    "b3QgQmxvY2tlZDsyOTgNCjI5OTtDbG9jayBBZmZlY3Rpbmc7Mjk5DQozMDA7U29sbyBSdW47MzAwDQozMDE7U2hv"
    "dCBmcm9tIGNyb3NzOzMwMQ0KMzAyO0NoZWNrcyBjb21wbGV0ZS9MaXZlIGNvbGxlY3Rpb24gY2hlY2tzIGNvbXBs"
    "ZXRlOzMwMg0KMzAzO0Zsb29kbGlnaHQgZmFpbHVyZTszMDMNCjMwNDtCYWxsIEluIFBsYXk7MzA0DQozMDU7QmFs"
    "bCBPdXQgb2YgUGxheTszMDUNCjMwNjtLaXQgY2hhbmdlOzMwNg0KMzA3O1BoYXNlIG9mIHBvc3Nlc3Npb24gSUQ7"
    "MzA3DQozMDg7R29lcyB0byBFeHRyYSBUaW1lOzMwOA0KMzA5O0dvZXMgdG8gUGVuYWx0aWVzOzMwOQ0KMzEwO1Bs"
    "YXllciBnb2VzIG91dDszMTANCjMxMTtQbGF5ZXIgY29tZXMgYmFjazszMTENCjMxMjtQaGFzZSBvZiBwb3NzZXNz"
    "aW9uIHN0YXJ0OzMxMg0KMzEzO0lsbGVnYWwgUmVzdGFydDszMTMNCjMxNDtFbmQgb2YgT2Zmc2lkZTszMTQNCjMx"
    "NjtQYXNzZWQgUGVuYWx0eTszMTYNCjMxNztQZW5hbHR5IFNldCBQaWVjZTszMTcNCjMxOTtDYXB0YWluIGNoYW5n"
    "ZTszMTkNCjMyMztGb2xsb3dzIGEgUmVib3VuZDszMjMNCjMyNDtGb2xsb3dzIGEgVGFrZSBPbjszMjQNCjMyNTtB"
    "YmFuZG9ubWVudCBUbyBGb2xsb3c7MzI1DQozMjg7Rmlyc3RUb3VjaDszMjgNCjMyOTtWQVIgLSBHb2FsIEF3YXJk"
    "ZWQ7MzI5DQozMzA7VkFSIC0gUGVuYWx0eSBBd2FyZGVkOzMzMA0KMzMxO1ZBUiAtIFBlbmFsdHkgTm90IEF3YXJk"
    "ZWQ7MzMxDQozMzI7VkFSIC0gKFJlZCkgQ2FyZCBVcGdyYWRlOzMzMg0KMzMzO1ZBUiAtIE1pc3Rha2VuIElkZW50"
    "aXR5OzMzMw0KMzM0O1ZBUiAtIE90aGVyOzMzNA0KMzM1O1JlZmVyZWUgRGVjaXNpb24gQ29uZmlybWVkOzMzNQ0K"
    "MzM2O1JlZmVyZWUgRGVjaXNpb24gQ2FuY2VsbGVkOzMzNg0KMzM4O0ZvbGxvd3MgYSBSZWJvdW5kIEV2ZW50IElE"
    "OzMzOA0KMzQxO1ZBUiAtIEdvYWwgTm90IEF3YXJkZWQ7MzQxDQozNDI7VkFSIC0gUmVkIENhcmQgR2l2ZW47MzQy"
    "DQozNDM7UmV2aWV3OzM0Mw0KMzQ0O1ZpZGVvIGNvdmVyYWdlIGxvc3Q7MzQ0DQozNDU7T3ZlcmhpdCBjcm9zczsz"
    "NDUNCjM0NjtOZXh0IGV2ZW50IEdvYWwtS2ljazszNDYNCjM0NztOZXh0IGV2ZW50IFRocm93LUluOzM0Nw0KMzQ4"
    "O1BlbmFsdHkgdGFrZXIgSUQ7MzQ4DQozNDk7R29hbGtlZXBlciBwdW5jaCBvdXRjb21lOzM0OQ0KMzUzO1NlY29u"
    "ZCAoMm5kKSBvcHBvc2l0ZSByZWxhdGVkIGV2ZW50IElEOzM1Mw0KMzU0O0JhbGwgaGl0cyByZWZlcmVlOzM1NA0K"
    "MzU1O0VudGVyaW5nIHJlZmVyZWUgcmV2aWV3IGFyZWE7MzU1DQozNTY7RXhjZXNzaXZlIHVzYWdlIG9mIHJldmll"
    "dyBzaWduYWw7MzU2DQozNTc7RW50ZXJpbmcgdmlkZW8gb3BlcmF0aW9ucyByb29tOzM1Nw0KMzU4O09mZmljaWFs"
    "IGJvZHk6IFJldmlld2VkIGFuZCBjb25maXJtZWQ7MzU4DQozNTk7T2ZmaWNpYWwgYm9keTogUmV2aWV3ZWQgYW5k"
    "IGNoYW5nZWQ7MzU5DQozNjE7SW5jb3JyZWN0IG91dCBvZiBwbGF5IGRlY2lzaW9uOzM2MQ0KMzYyO1ZpcmFsOzM2"
    "Mg0KMzYzO0F3YXkgYXR0ZW5kYW5jZTszNjMNCjM2NDtWQVIgRGVsYXk7MzY0DQozNjU7UmV2aWV3ZWQgZXZlbnQg"
    "SUQ7MzY1DQozNzQ7R29hbCBzaG90IHRpbWVzdGFtcDszNzQNCjM3NTtHb2FsIHNob3QgZ2FtZSBjbG9jazszNzUN"
    "CjM3NjtMb3cgR0sgaW50ZXJ2ZW50aW9uOzM3Ng0KMzc3O01lZGl1bSBHSyBpbnRlcnZlbnRpb247Mzc3DQozNzg7"
    "SGlnaCBHSyBpbnRlcnZlbnRpb247Mzc4DQozODA7T3RoZXIgb2JzdGFjbGU7MzgwDQozODE7RnVtYmxlOzM4MQ0K"
    "MzgzO1RvdWNoIHR5cGUgY29udHJvbDszODMNCjM4NDtUb3VjaCB0eXBlIHBhc3M7Mzg0DQozODU7VG91Y2ggdHlw"
    "ZSBjbGVhcmFuY2U7Mzg1DQozODY7RHJpdmVuIGNyb3NzOzM4Ng0KMzg3O0Zsb2F0ZWQgY3Jvc3M7Mzg3DQozODg7"
    "SnVtcGluZzszODgNCjM4OTtTbGlkaW5nOzM4OQ0KMzkwO0NhdXNpbmcgcGxheWVyOzM5MA0KMzkxO01pcy1oaXQ7"
    "MzkxDQozOTI7UmVja2xlc3Mgb2ZmZW5jZTszOTINCjM5MztUYWN0aWNhbCBGb3VsOzM5Mw0KMzk0O0Nvcm5lciBu"
    "b3QgdGFrZW47Mzk0DQozOTU7R0sgeCBjb29yZGluYXRlIHRpbWUgb2YgZ29hbDszOTUNCjM5NjtHSyB5IGNvb3Jk"
    "aW5hdGUgdGltZSBvZiBnb2FsOzM5Ng0KMzk3O0Jsb2NrZWQgY2xlYXJhbmNlOzM5Nw0KMzk4O0dLIENoYWxsZW5n"
    "ZTszOTgNCjM5OTtJbnRlbmRlZCB0YWNrbGUgdGFyZ2V0OzM5OQ0KNDA2O0NvbGxlY3Rpb24gY29tcGxldGU7NDA2"
    "DQo0MzY7UHJlLVJldmlldyBFdmVudCBUeXBlOzQzNg0KNDU4O05vdCBhc3Npc3RlZDs0NTgNCjQ1OTtFdmVudCB0"
    "eXBlIHJldmlldzs0NTkNCjQ2NDtUYWtlIG9uIHNwYWNlOzQ2NA0KNDY1O1Rha2Ugb24gb3ZlcnRha2U7NDY1DQo0"
    "Njc7RGVmZW5zaXZlIDEgdiAxOzQ2Nw0KNDY4O1JlbGF0ZWQgZXJyb3IgMSBJRDs0NjgNCjQ3MjtGYW50YXN5IGFz"
    "c2lzdCBJRMKgOzQ3Mg0KNDc0O1JlbGF0ZWQgZXJyb3IgMiBJRDs0NzQNCjQ3NjtOZXcgc3RhcnQgdGltZTs0NzYN"
    "CjQ3ODtPZmZpY2lhbGx5IGFubm91bmNlZDs0NzgNCjQ3OTtFc3RpbWF0ZWQ7NDc5DQo0ODQ7RHViaW91cyBzY29y"
    "ZXI7NDg0DQo0ODU7QWR2YW50YWdlIHBsYXllZDs0ODU="
),
    'xT_Grid.csv': (
    "MC4wMDYzODMwMywwLjAwNzc5NjE2LDAuMDA4NDQ4NTQsMC4wMDk3NzY1OSwwLjAxMTI2MjY3LDAuMDEyNDgzNDQs"
    "MC4wMTQ3MzU5NiwwLjAxNzQ1MDYsMC4wMjEyMjEyOSwwLjAyNzU2MzEyLDAuMDM0ODUwNzIsMC4wMzc5MjU5DQow"
    "LjAwNzUwMDcyLDAuMDA4Nzg1ODksMC4wMDk0MjM4MiwwLjAxMDU5NDksMC4wMTIxNDcxOSwwLjAxMzg0NTQsMC4w"
    "MTYxMTgxMywwLjAxODcwMzQ3LDAuMDI0MDE1MjEsMC4wMjk1MzI3MiwwLjA0MDY2OTkyLDAuMDQ2NDc3MjENCjAu"
    "MDA4ODc5OSwwLjAwOTc3NzQ1LDAuMDEwMDEzMDQsMC4wMTExMDQ2MiwwLjAxMjY5MTc0LDAuMDE0MjkxMjgsMC4w"
    "MTY4NTU5NiwwLjAxOTM1MTMyLDAuMDI0MTIyNCwwLjAyODU1MjAyLDAuMDU0OTExMzgsMC4wNjQ0MjU5NQ0KMC4w"
    "MDk0MTA1NiwwLjAxMDgyNzIyLDAuMDEwMTY1NDksMC4wMTEzMjM3NiwwLjAxMjYyNjQ2LDAuMDE0ODQ1OTgsMC4w"
    "MTY4OTUyOCwwLjAxOTk3MDcsMC4wMjM4NTE0OSwwLjAzNTExMzI2LDAuMTA4MDUxMDIsMC4yNTc0NTM2Mg0KMC4w"
    "MDk0MTA1NiwwLjAxMDgyNzIyLDAuMDEwMTY1NDksMC4wMTEzMjM3NiwwLjAxMjYyNjQ2LDAuMDE0ODQ1OTgsMC4w"
    "MTY4OTUyOCwwLjAxOTk3MDcsMC4wMjM4NTE0OSwwLjAzNTExMzI2LDAuMTA4MDUxMDIsMC4yNTc0NTM2Mg0KMC4w"
    "MDg4Nzk5LDAuMDA5Nzc3NDUsMC4wMTAwMTMwNCwwLjAxMTEwNDYyLDAuMDEyNjkxNzQsMC4wMTQyOTEyOCwwLjAx"
    "Njg1NTk2LDAuMDE5MzUxMzIsMC4wMjQxMjI0LDAuMDI4NTUyMDIsMC4wNTQ5MTEzOCwwLjA2NDQyNTk1DQowLjAw"
    "NzUwMDcyLDAuMDA4Nzg1ODksMC4wMDk0MjM4MiwwLjAxMDU5NDksMC4wMTIxNDcxOSwwLjAxMzg0NTQsMC4wMTYx"
    "MTgxMywwLjAxODcwMzQ3LDAuMDI0MDE1MjEsMC4wMjk1MzI3MiwwLjA0MDY2OTkyLDAuMDQ2NDc3MjENCjAuMDA2"
    "MzgzMDMsMC4wMDc3OTYxNiwwLjAwODQ0ODU0LDAuMDA5Nzc2NTksMC4wMTEyNjI2NywwLjAxMjQ4MzQ0LDAuMDE0"
    "NzM1OTYsMC4wMTc0NTA2LDAuMDIxMjIxMjksMC4wMjc1NjMxMiwwLjAzNDg1MDcyLDAuMDM3OTI1OQ=="
),
}


def _validate_inputs(url_365scores: str, url_scoresway: str) -> list[str]:
    errors = []

    if not url_365scores.strip():
        errors.append("Falta el link de 365Scores.")
    elif "365scores.com" not in url_365scores:
        errors.append("El primer link no parece ser de 365Scores.")

    if not url_scoresway.strip():
        errors.append("Falta el link de Scoresway.")
    elif "scoresway.com" not in url_scoresway or "/match/view/" not in url_scoresway:
        errors.append("El segundo link no parece ser un partido de Scoresway.")

    return errors


def _load_notebook_code_cells() -> list[tuple[int, str]]:
    return [
        (index, source)
        for index, source in NOTEBOOK_CODE_CELLS
        if index <= FINAL_CELL_INDEX
    ]


def _progress_label_for_cell(cell_index: int) -> str:
    label = PROGRESS_LABELS[0][1]
    for threshold, candidate in PROGRESS_LABELS:
        if cell_index >= threshold:
            label = candidate
        else:
            break
    return label


def _patch_first_cell(
    source: str,
    url_365scores: str,
    url_scoresway: str,
    home_color: str,
    away_color: str,
) -> str:
    replacements = [
        (
            r'players365\s*=\s*threesixfivescores\.get_players_info\("[^"]+"\)',
            f"players365 = threesixfivescores.get_players_info({url_365scores!r})",
        ),
        (
            r'shots365\s*=\s*threesixfivescores\.get_match_shotmap\("[^"]+"\)',
            f"shots365 = threesixfivescores.get_match_shotmap({url_365scores!r})",
        ),
        (
            r'url_partido\s*=\s*"[^"]+"',
            f"url_partido = {url_scoresway!r}",
        ),
        (
            r"col1\s*=\s*['\"][^'\"]+['\"]",
            f"col1 = {home_color!r}",
        ),
        (
            r"col2\s*=\s*['\"][^'\"]+['\"]",
            f"col2 = {away_color!r}",
        ),
    ]

    patched = source
    for pattern, replacement in replacements:
        patched, count = re.subn(pattern, replacement, patched, count=1)
        if count != 1:
            raise RuntimeError(
                "No pude reemplazar una de las URLs iniciales del notebook. "
                "Revisa si cambio la primera celda."
            )

    return patched


def _patch_report_subtitles(source: str, subtitle: str) -> str:
    return re.sub(
        r'fig\.text\(0\.5,\s*0\.95,\s*f?"[^"]*Reporte Post-Match[^"]*",\s*color=line_color,\s*fontsize=30,\s*ha=[\'"]center[\'"],\s*va=[\'"]center[\'"]\)',
        f'fig.text(0.5, 0.95, {subtitle!r}, color=line_color, fontsize=30, ha="center", va="center")',
        source,
    )


def _patch_runtime_options(source: str, selenium_wait: int, headless: bool) -> str:
    # Por defecto usamos Chrome NO-headless con Xvfb, porque Scoresway puede no exponer endpoints en headless.
    patched = source
    patched = re.sub(r"esperar=12|esperar=25", f"esperar={selenium_wait}", patched)
    patched = re.sub(r"headless=(True|False)", f"headless={headless}", patched)
    return patched


def _build_runner_code(
    url_365scores: str,
    url_scoresway: str,
    selenium_wait: int,
    headless: bool,
    home_color: str,
    away_color: str,
    subtitle: str,
) -> str:
    header = """
import os
os.environ.setdefault("MPLBACKEND", "Agg")

import matplotlib
matplotlib.use("Agg")

try:
    from IPython.display import display
except Exception:
    def display(*args, **kwargs):
        for arg in args:
            print(arg)
"""

    chunks = [header]
    code_cells = _load_notebook_code_cells()
    total_cells = len(code_cells)

    for step_number, (cell_index, source) in enumerate(code_cells, start=1):
        if cell_index == 0:
            source = _patch_first_cell(
                source,
                url_365scores,
                url_scoresway,
                home_color,
                away_color,
            )

        source = _patch_runtime_options(source, selenium_wait, headless)
        source = _patch_report_subtitles(source, subtitle)

        chunks.append(f"\n# --- Notebook cell {cell_index} ---\n")
        chunks.append(
            f'print("{PROGRESS_MARKER}:{step_number}:{total_cells}:{cell_index}", flush=True)\n'
        )
        chunks.append(source)
        chunks.append("\n")

    footer = """
from pathlib import Path

missing_reports = [name for name in ("Match_Report_1.png", "Match_Report.png") if not Path(name).exists()]
if missing_reports:
    raise FileNotFoundError(f"No se generaron los reportes esperados: {missing_reports}")

print("REPORTES_LISTOS")
"""
    chunks.append(footer)

    return "".join(chunks)


def _prepare_workdir() -> Path:
    workdir = Path(tempfile.mkdtemp(prefix="scoresway_reports_"))

    for name in RESOURCE_FILES:
        encoded = EMBEDDED_RESOURCE_FILES_B64.get(name)
        if not encoded:
            raise FileNotFoundError(f"No esta embebido el archivo auxiliar: {name}")
        (workdir / name).write_bytes(base64.b64decode(encoded))

    return workdir


def _run_report_job(
    url_365scores: str,
    url_scoresway: str,
    selenium_wait: int,
    headless: bool,
    home_color: str,
    away_color: str,
    subtitle: str,
    log_placeholder,
    progress_bar,
    progress_text,
) -> Path:
    workdir = _prepare_workdir()
    runner_path = workdir / "_scoresway_report_runner.py"
    runner_path.write_text(
        _build_runner_code(
            url_365scores,
            url_scoresway,
            selenium_wait,
            headless,
            home_color,
            away_color,
            subtitle,
        ),
        encoding="utf-8",
    )

    process = subprocess.Popen(
        [sys.executable, "-u", str(runner_path)],
        cwd=workdir,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )

    output_lines: list[str] = []
    progress_bar.progress(0)
    progress_text.caption("Inicializando generacion...")

    assert process.stdout is not None
    for line in process.stdout:
        stripped_line = line.strip()

        if stripped_line.startswith(f"{PROGRESS_MARKER}:"):
            try:
                _, step_number, total_cells, cell_index = stripped_line.split(":")
                step_number_int = int(step_number)
                total_cells_int = int(total_cells)
                cell_index_int = int(cell_index)
                progress = min(step_number_int / total_cells_int, 0.99)
                progress_bar.progress(progress)
                progress_text.caption(_progress_label_for_cell(cell_index_int))
            except ValueError:
                pass
            continue

        output_lines.append(line)
        log_placeholder.code("".join(output_lines[-120:]), language="text")

    return_code = process.wait()
    if return_code != 0:
        shutil.rmtree(workdir, ignore_errors=True)
        raise RuntimeError(
            "La ejecucion del notebook fallo.\n\n"
            + "".join(output_lines[-160:])
        )

    progress_bar.progress(1.0)
    progress_text.caption("Reportes listos.")

    return workdir


def _collect_report_bytes(workdir: Path) -> dict[str, bytes]:
    try:
        return {name: (workdir / name).read_bytes() for name in REPORT_FILES}
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def _is_scoresway_network_capture_error(exc: Exception) -> bool:
    message = str(exc)
    return (
        "api.performfeeds.com/soccerdata" in message
        and "No se encontr" in message
    )


def _run_report_job_with_retry(
    url_365scores: str,
    url_scoresway: str,
    selenium_wait: int,
    headless: bool,
    home_color: str,
    away_color: str,
    subtitle: str,
    log_placeholder,
    progress_bar,
    progress_text,
) -> Path:
    try:
        return _run_report_job(
            url_365scores=url_365scores,
            url_scoresway=url_scoresway,
            selenium_wait=selenium_wait,
            headless=headless,
            home_color=home_color,
            away_color=away_color,
            subtitle=subtitle,
            log_placeholder=log_placeholder,
            progress_bar=progress_bar,
            progress_text=progress_text,
        )
    except RuntimeError as exc:
        if not _is_scoresway_network_capture_error(exc):
            raise

        log_placeholder.warning(
            "Scoresway no expuso los endpoints en el primer intento. "
            "Reintentando con Chrome NO-headless via Xvfb y mas tiempo de espera..."
        )

        return _run_report_job(
            url_365scores=url_365scores,
            url_scoresway=url_scoresway,
            selenium_wait=max(selenium_wait + 12, 35),
            headless=False,
            home_color=home_color,
            away_color=away_color,
            subtitle=subtitle,
            log_placeholder=log_placeholder,
            progress_bar=progress_bar,
            progress_text=progress_text,
        )


def _show_reports(reports: dict[str, bytes]) -> None:
    st.success("Reportes generados.")

    tab_1, tab_2 = st.tabs(["Reporte 1", "Reporte 2"])

    with tab_1:
        image_bytes = reports["Match_Report_1.png"]
        st.image(image_bytes, use_container_width=True)
        st.download_button(
            "Descargar página 1",
            data=image_bytes,
            file_name="Match_Report_1.png",
            mime="image/png",
        )

    with tab_2:
        image_bytes = reports["Match_Report.png"]
        st.image(image_bytes, use_container_width=True)
        st.download_button(
            "Descargar página 2",
            data=image_bytes,
            file_name="Match_Report.png",
            mime="image/png",
        )


def main() -> None:
    st.set_page_config(page_title="Reportes colectivos", layout="wide")

    st.title("Reportes colectivos con data de eventing")
    st.markdown(APP_EXPLANATORY_TEXT)
    st.markdown(TEXTO)

    url_365scores = st.text_input(
        "Link 365Scores",
        placeholder="https://www.365scores.com/es/football/match/liga-profesional-72/argentinos-juniors-lanus-869-871-72#id=4710402",
    )
    url_scoresway = st.text_input(
        "Link Scoresway",
        placeholder="https://www.scoresway.com/en_GB/soccer/liga-profesional-argentina-2026/8v84l9nq3d5t0j4gb781i3llg/match/view/bgxpwu2xbgue9v5iutmqtwnx0/match-summary",
    )

    col_1, col_2, col_3 = st.columns([1, 1, 1])
    with col_1:
        home_color = st.color_picker("Color local", "#ff4b44")
    with col_2:
        away_color = st.color_picker("Color visitante", "#00a0de")
    with col_3:
        headless = st.toggle(
            "Chrome headless",
            value=False,
            help="Dejalo apagado en Streamlit Cloud: usa Xvfb para correr Chrome como no-headless."
        )

    subtitle = st.text_input(
        "Subtitulo",
        value="Jornada | Torneo | Reporte",
        placeholder="Jornada | Torneo | Reporte",
    )

    if "last_reports" in st.session_state:
        _show_reports(st.session_state["last_reports"])

    if not st.button("Generar reportes", type="primary"):
        return

    errors = _validate_inputs(url_365scores, url_scoresway)
    if errors:
        for error in errors:
            st.error(error)
        return

    with st.expander("Ver ejecucion del notebook", expanded=False):
        log_placeholder = st.empty()

    progress_bar = st.progress(0)
    progress_text = st.empty()

    with st.status("Generando reportes...", expanded=False) as status:
        try:
            workdir = _run_report_job_with_retry(
                url_365scores=url_365scores.strip(),
                url_scoresway=url_scoresway.strip(),
                selenium_wait=DEFAULT_SELENIUM_WAIT,
                headless=headless,
                home_color=home_color,
                away_color=away_color,
                subtitle=subtitle.strip() or "Jornada | Torneo | Reporte",
                log_placeholder=log_placeholder,
                progress_bar=progress_bar,
                progress_text=progress_text,
            )
            reports = _collect_report_bytes(workdir)
        except Exception as exc:
            status.update(label="Fallo la generacion", state="error", expanded=False)
            st.exception(exc)
            return

        status.update(label="Listo", state="complete", expanded=False)
        st.session_state["last_reports"] = reports

    _show_reports(reports)

    


if __name__ == "__main__":
    main()
