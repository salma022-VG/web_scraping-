"""
====================================================================
 MONITOREO DE NOTICIAS - BOGOTA Y COLOMBIA
====================================================================

QUE HACE ESTE PROGRAMA (explicado sin tecnicismos):

Este programa busca automaticamente noticias en internet (usando Google
Noticias, que junta noticias de cientos de medios: El Tiempo, El
Espectador, Semana, RCN, Caracol, Infobae, Portafolio, La Republica,
El Colombiano, etc.) y arma un archivo de Excel con:

  1) Noticias de HOY sobre Bogota relacionadas con: Hacienda, infraestructura,
     vivienda, movilidad, sector inmobiliario, el alcalde Carlos Fernando
     Galan, y Catastro / certificados catastrales.
  2) Las 3 noticias nacionales de Colombia mas importantes del dia.

Cada noticia en el Excel trae: categoria (tema), medio que la publico,
titulo, un pequeno resumen, fecha/hora y el link para abrirla.

--------------------------------------------------------------------
POR QUE SE HIZO ASI (y no "leyendo" cada pagina web una por una):

Cada periodico tiene su propia pagina, con su propio diseno, y muchos
bloquean a los robots que intentan leer su contenido automaticamente.
Si el programa tuviera que entrar pagina por pagina a cada periodico,
se romperia constantemente. Por eso usamos Google Noticias: es un
"indice" que ya junta miles de fuentes en un solo lugar, no necesita
contrasena ni permisos especiales, y siempre dice a que hora se publico
cada noticia (esto es clave para poder filtrar "las ultimas 12 horas",
por ejemplo).

--------------------------------------------------------------------
COMO DECIDE QUE HORAS DE NOTICIAS TRAER (hora de Bogota):

  - Si lo corres entre las 8:00 y las 9:00 de la manana
    -> trae noticias de las ULTIMAS 12 HORAS, contadas hasta las 8:00 am
       (es decir, desde las 8:00 pm del dia anterior).

  - Si lo corres entre las 2:00 pm y las 3:30 pm
    -> trae noticias de las ULTIMAS 6 HORAS, contadas hasta las 3:00 pm
       (es decir, desde las 9:00 am del mismo dia).

  - Si lo corres en cualquier otro horario (por ejemplo, para hacer una
    prueba), el programa usa como respaldo la ultima de esas dos franjas
    que ya haya pasado en el dia. Toda esta logica esta en la funcion
    "get_time_window" mas abajo, por si se necesita ajustar en el futuro.

--------------------------------------------------------------------
COMO AGREGAR O QUITAR PALABRAS DE BUSQUEDA (sin tocar el codigo):

Las categorias de Bogota (Hacienda, Movilidad, Catastro, etc.) y las
palabras que se buscan para cada una NO estan escritas adentro del codigo:
estan en un archivo de Excel llamado "Palabras_Clave_Bogota.xlsx", en esta
misma carpeta, con dos columnas: "categoria" y "palabra_clave".

  - Si quieres agregar una palabra nueva a una categoria que ya existe
    (por ejemplo, otro termino para "Movilidad"), abre ese Excel y agrega
    una fila nueva con la categoria en la primera columna y la palabra en
    la segunda.
  - Si quieres crear una categoria totalmente nueva, simplemente escribe
    el nombre nuevo en la columna "categoria" de una fila nueva.
  - Guarda el Excel y ya. La proxima vez que corras el programa, va a leer
    ese archivo automaticamente y va a usar la lista actualizada, sin
    necesidad de tocar este archivo de codigo.

Si ese Excel no existe (por ejemplo, la primera vez que se usa el
programa en una computadora nueva), el programa lo crea solo, con una
lista inicial de categorias y palabras ya cargada.

Ademas, cada vez que corres el programa, este mismo archivo de codigo
(monitoreo_noticias.py) se actualiza solo, dejando guardada una copia de
respaldo de las palabras que hay en el Excel en ese momento (esto es la
lista "DEFAULT_BOGOTA_QUERIES" mas abajo). Es solo un respaldo por si el
Excel se llegara a perder o borrar sin querer; el programa SIEMPRE usa lo
que este en el Excel, nunca esa copia, mientras el Excel exista.

--------------------------------------------------------------------
COMO SE USA (desde una terminal / simbolo de sistema, dentro de esta carpeta):

  python monitoreo_noticias.py
      -> Uso normal: usa la hora actual de la computadora para decidir
         que franja de noticias traer.

  python monitoreo_noticias.py --now "2026-07-28 08:30"
      -> Para PROBAR el programa como si fueran las 8:30 am de esa fecha,
         sin tener que esperar a que realmente sean esas horas.

  python monitoreo_noticias.py --window 12
      -> Para pedirle directamente "traeme las ultimas 12 horas desde
         ahora mismo", sin importar la franja horaria del reloj.

Al terminar, se crea un archivo llamado algo como
"Monitoreo_Noticias_2026-07-28_0930.xlsx" en esta misma carpeta. Ese
numero en el nombre es la fecha y hora en que se genero, para que nunca
se sobreescriba un reporte anterior por accidente.
====================================================================
"""

from __future__ import annotations

# --- Librerias que usa el programa ----------------------------------------
# (vienen instaladas por separado, ver archivo requirements.txt)
import argparse       # permite escribir opciones como --now o --window al ejecutar el programa
import os               # para revisar si el archivo de palabras clave ya existe en la carpeta
import re              # ayuda a "leer" y limpiar texto usando patrones (por ejemplo, separar el titulo del nombre del medio)
import sys              # utilidades del sistema, aqui solo se usa para avisar errores
import time              # maneja fechas/horas y pausas breves
import unicodedata       # ayuda a comparar titulos ignorando tildes/mayusculas, para no repetir noticias
from concurrent.futures import ThreadPoolExecutor, as_completed  # permite descargar varias noticias AL MISMO TIEMPO en vez de una por una (para que el programa no se demore tanto)
from dataclasses import dataclass  # forma sencilla de crear una "ficha" con los datos de cada noticia
from datetime import datetime, timedelta, timezone  # para trabajar con fechas y horas
from urllib.parse import quote  # convierte texto (como "vivienda Bogota") en un formato valido para meterlo en una direccion web
from zoneinfo import ZoneInfo  # para que todas las horas se calculen en la zona horaria de Bogota, sin importar donde este la computadora

import feedparser         # entiende el formato en que Google Noticias entrega los resultados (se llama "RSS")
import pandas as pd        # organiza las noticias en tablas, como si fueran una hoja de calculo, antes de guardarlas
import requests             # es quien realmente "toca la puerta" de internet y descarga paginas y resultados
from bs4 import BeautifulSoup          # lee el codigo de una pagina web para sacarle el resumen de la noticia
from googlenewsdecoder import gnewsdecoder  # convierte el link "raro" de Google Noticias en el link real del periodico

# Zona horaria de Bogota: TODAS las horas del programa se calculan con esta,
# para que no importe en que pais este configurada la computadora que lo ejecuta.
BOGOTA_TZ = ZoneInfo("America/Bogota")

# Configuracion regional que le decimos a Google Noticias: idioma espanol
# latinoamericano (HL), pais Colombia (GL), y la combinacion de ambos (CEID).
HL, GL, CEID = "es-419", "CO", "CO:es-419"

# Cuando el programa "visita" una pagina web, se identifica como si fuera un
# navegador normal (Chrome en Windows). Esto es necesario porque algunas
# paginas bloquean las visitas que no parecen venir de un navegador real.
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# ---------------------------------------------------------------------------
# 1. TEMAS Y PALABRAS DE BUSQUEDA PARA BOGOTA
# ---------------------------------------------------------------------------
# La lista real de temas y palabras de busqueda YA NO esta aqui en el
# codigo: vive en el archivo "Palabras_Clave_Bogota.xlsx" (ver mas abajo,
# funciones "ensure_keywords_file" y "load_bogota_queries"), para que se
# pueda editar desde Excel sin tocar el codigo.
#
# Lo que hay aqui abajo es solo la lista INICIAL con la que se crea ese
# Excel la primera vez (si todavia no existe). Una vez creado el archivo,
# el programa siempre lee de el, y esta lista de aqui ya no se vuelve a usar.

KEYWORDS_FILE = "Palabras_Clave_Bogota.xlsx"

DEFAULT_BOGOTA_QUERIES: dict[str, list[str]] = {
    "Hacienda": [
        "Secretaria de Hacienda Bogota",
        "Hacienda Distrital Bogota",
        "impuestos Bogota",
        "predial Bogota",
    ],
    "Infraestructura": [
        "infraestructura Bogota",
        "obras Bogota",
        "IDU Bogota",
        "Metro de Bogota",
    ],
    "Vivienda": [
        "vivienda Bogota",
        "vivienda VIS Bogota",
        "Metrovivienda",
        "Secretaria del Habitat Bogota",
    ],
    "Movilidad": [
        "movilidad Bogota",
        "TransMilenio",
        "Secretaria de Movilidad Bogota",
        "pico y placa Bogota",
    ],
    "Inmobiliario": [
        "sector inmobiliario Bogota",
        "mercado inmobiliario Bogota",
        "finca raiz Bogota",
    ],
    "Alcalde Carlos Fernando Galan": [
        "Carlos Fernando Galan",
        "alcalde de Bogota Galan",
    ],
    "Catastro Bogota": [
        "Catastro Bogota",
        "certificado catastral Bogota",
        "IDECA Bogota",
        "IDECA",
        "avaluo catastral Bogota",
        "Unidad Administrativa Especial de Catastro Distrital",
        "Olga Lucia Lopez Morales",
        "mapas Bogota",
    ],
}

# Direccion web de donde se sacan las noticias NACIONALES de Colombia
# (la seccion "Colombia / Nacion" de Google Noticias).
NATIONAL_TOPICS_FEED = (
    f"https://news.google.com/rss/headlines/section/topic/NATION?hl={HL}&gl={GL}&ceid={CEID}"
)


def ensure_keywords_file(path: str) -> None:
    """
    Si el archivo de palabras clave (Palabras_Clave_Bogota.xlsx) todavia no
    existe en la carpeta, lo crea con la lista inicial (DEFAULT_BOGOTA_QUERIES),
    ya con formato de tabla legible (encabezado en negrita, columnas anchas).
    Asi, la primera vez que alguien use el programa en una computadora
    nueva, no hace falta crear el Excel a mano.
    """
    if os.path.exists(path):
        return  # ya existe, no se toca (para no borrar palabras que alguien ya agrego)

    filas = [
        {"categoria": categoria, "palabra_clave": palabra}
        for categoria, palabras in DEFAULT_BOGOTA_QUERIES.items()
        for palabra in palabras
    ]
    df = pd.DataFrame(filas, columns=["categoria", "palabra_clave"])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Palabras clave", index=False)

        from openpyxl.styles import Font

        ws = writer.sheets["Palabras clave"]
        for cell in ws[1]:  # fila 1 = encabezados
            cell.font = Font(bold=True)
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value else 0) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 15), 60)
        ws.freeze_panes = "A2"

    print(f"Se creo el archivo de palabras clave: {path}")


def load_bogota_queries(path: str) -> dict[str, list[str]]:
    """
    Lee el archivo "Palabras_Clave_Bogota.xlsx" y arma, a partir de sus
    filas, la lista de categorias con sus palabras de busqueda (agrupa
    todas las filas que tengan la misma categoria). Si el archivo no
    existe todavia, primero lo crea (ver "ensure_keywords_file").

    Esto es lo que hace posible que alguien SIN saber programar pueda
    agregar una palabra de busqueda nueva: solo tiene que abrir el Excel,
    escribir una fila nueva, y guardar. La proxima corrida del programa la
    va a tomar en cuenta automaticamente.
    """
    ensure_keywords_file(path)

    df = pd.read_excel(path)
    df.columns = [str(c).strip().lower() for c in df.columns]

    queries: dict[str, list[str]] = {}
    for _, fila in df.iterrows():
        categoria = str(fila.get("categoria", "")).strip()
        palabra = str(fila.get("palabra_clave", "")).strip()
        # Se descartan filas vacias o mal diligenciadas (por ejemplo, si
        # alguien dejo una fila en blanco por error al final del Excel).
        if not categoria or not palabra or categoria.lower() == "nan" or palabra.lower() == "nan":
            continue
        queries.setdefault(categoria, []).append(palabra)

    return queries


def _build_default_queries_block(queries: dict[str, list[str]]) -> str:
    """
    Arma, como texto, el bloque de codigo "DEFAULT_BOGOTA_QUERIES = {...}"
    que refleja lo que hay en este momento en el Excel de palabras clave.
    Esto es solo para escribirlo de vuelta en el archivo .py (ver
    "sync_default_queries_in_code" mas abajo).
    """
    lineas = ["DEFAULT_BOGOTA_QUERIES: dict[str, list[str]] = {"]
    for categoria, palabras in queries.items():
        categoria_escapada = categoria.replace('"', '\\"')
        lineas.append(f'    "{categoria_escapada}": [')
        for palabra in palabras:
            palabra_escapada = palabra.replace('"', '\\"')
            lineas.append(f'        "{palabra_escapada}",')
        lineas.append("    ],")
    lineas.append("}")
    return "\n".join(lineas) + "\n"


def sync_default_queries_in_code(queries: dict[str, list[str]], py_path: str) -> None:
    """
    Deja el archivo de codigo (monitoreo_noticias.py) actualizado con una
    "copia de respaldo" de las palabras clave que hay ahora mismo en el
    Excel. Esto NO afecta el funcionamiento del programa (que siempre usa
    lo que este en el Excel), es solo para que, si el Excel se llegara a
    perder o borrar por accidente, el codigo tenga guardada la ultima
    version conocida de la lista y se pueda recrear el Excel desde ahi.

    En otras palabras: el programa reescribe una pequena parte de su
    propio archivo cada vez que corre, para mantenerlo igual al Excel.
    """
    try:
        with open(py_path, "r", encoding="utf-8") as f:
            contenido = f.read()
    except Exception as exc:  # noqa: BLE001
        print(f"  [aviso] no se pudo leer el codigo para sincronizarlo ({exc})")
        return

    patron = re.compile(r"DEFAULT_BOGOTA_QUERIES: dict\[str, list\[str\]\] = \{.*?\n\}\n", re.DOTALL)
    if not patron.search(contenido):
        print("  [aviso] no se encontro el bloque DEFAULT_BOGOTA_QUERIES en el codigo; no se actualizo.")
        return

    nuevo_bloque = _build_default_queries_block(queries)
    contenido_actualizado = patron.sub(nuevo_bloque, contenido, count=1)

    if contenido_actualizado == contenido:
        return  # no habia nada nuevo que sincronizar

    try:
        with open(py_path, "w", encoding="utf-8") as f:
            f.write(contenido_actualizado)
        print("Se actualizo la lista DEFAULT_BOGOTA_QUERIES dentro de monitoreo_noticias.py "
              "para que quede igual al Excel de palabras clave.")
    except Exception as exc:  # noqa: BLE001
        print(f"  [aviso] no se pudo guardar el codigo actualizado ({exc})")


# ---------------------------------------------------------------------------
# 2. CALCULO DE LA FRANJA HORARIA
# ---------------------------------------------------------------------------
# Esta funcion decide, segun la hora a la que se ejecuta el programa, desde
# y hasta que momento se deben buscar noticias.

def get_time_window(now: datetime | None = None, force_hours: float | None = None):
    """
    Calcula el rango de fechas/horas ("desde" y "hasta") que se va a usar
    para filtrar noticias, segun a que hora se este ejecutando el programa.

    Devuelve tres cosas: la hora de inicio, la hora final, y un texto que
    explica en palabras que regla se aplico (esto se guarda tambien en el
    Excel, en la hoja "Resumen", para que quede claro que se pidio).
    """
    # Si no nos dan una hora especifica, se usa la hora actual real (en Bogota).
    now = now or datetime.now(BOGOTA_TZ)
    if now.tzinfo is None:
        now = now.replace(tzinfo=BOGOTA_TZ)

    # Opcion manual: si alguien pide explicitamente "las ultimas X horas"
    # (con --window), se usa eso y no se aplica ninguna otra regla.
    if force_hours is not None:
        end = now
        start = end - timedelta(hours=force_hours)
        return start, end, f"ultimas {force_hours} horas (forzado, hasta {end:%H:%M})"

    # Convertimos la hora actual a "minutos desde la medianoche" para poder
    # comparar horarios facilmente (por ejemplo, 8:30 am = 8*60+30 = 510).
    hm = now.hour * 60 + now.minute

    morning_start, morning_end = 8 * 60, 9 * 60             # ventana 08:00 - 09:00
    afternoon_start, afternoon_end = 14 * 60, 15 * 60 + 30  # ventana 14:00 - 15:30

    # Caso 1: se esta ejecutando en la franja de la MANANA (8:00 a 9:00 am).
    if morning_start <= hm <= morning_end:
        end = now.replace(hour=8, minute=0, second=0, microsecond=0)
        start = end - timedelta(hours=12)
        return start, end, "franja manana: ultimas 12 horas hasta las 08:00"

    # Caso 2: se esta ejecutando en la franja de la TARDE (2:00 a 3:30 pm).
    if afternoon_start <= hm <= afternoon_end:
        end = now.replace(hour=15, minute=0, second=0, microsecond=0)
        start = end - timedelta(hours=6)
        return start, end, "franja tarde: ultimas 6 horas hasta las 15:00"

    # Caso 3 (respaldo): si se ejecuta fuera de esas dos franjas -por ejemplo,
    # alguien lo corre manualmente a otra hora para revisar algo- el programa
    # no se queda sin hacer nada: usa la ultima franja oficial que ya paso
    # hoy, para seguir dando un resultado razonable.
    if hm > afternoon_end:
        # Ya paso la franja de la tarde -> se usa esa.
        end = now.replace(hour=15, minute=0, second=0, microsecond=0)
        start = end - timedelta(hours=6)
        return start, end, "fuera de franja: usando ultima franja (tarde, 6h hasta 15:00)"
    if hm > morning_end:
        # Ya paso la franja de la manana pero no la de la tarde -> se usa la de la manana.
        end = now.replace(hour=8, minute=0, second=0, microsecond=0)
        start = end - timedelta(hours=12)
        return start, end, "fuera de franja: usando ultima franja (manana, 12h hasta 08:00)"

    # Si todavia no ha pasado ninguna franja hoy (por ejemplo, son las 3 am),
    # como ultimo recurso se traen las ultimas 12 horas desde el momento actual.
    end = now
    start = end - timedelta(hours=12)
    return start, end, "fuera de franja: ultimas 12 horas desde ahora"


# ---------------------------------------------------------------------------
# 3. DESCARGA Y LECTURA DE LOS RESULTADOS DE GOOGLE NOTICIAS
# ---------------------------------------------------------------------------

@dataclass
class NewsItem:
    """
    Esta es la "ficha" con la que el programa guarda cada noticia mientras
    trabaja. Cada noticia tiene: a que categoria/tema pertenece, su titulo,
    el medio que la publico, la fecha/hora de publicacion, el link original
    de Google Noticias, y (mas adelante en el proceso) el link real de la
    pagina y un resumen. No es una tabla de Excel todavia, es solo la forma
    en que el programa organiza la informacion mientras la procesa.
    """
    categoria: str
    titular: str
    medio: str
    fecha: datetime
    link: str
    enlace: str = ""   # se llena despues, con el link real del periodico
    resumen: str = ""  # se llena despues, con un pequeno resumen de la noticia


def _clean_title(raw_title: str) -> tuple[str, str]:
    """
    Google Noticias entrega los titulos con el nombre del medio pegado al
    final, asi: "Bogota anuncia nuevas obras - El Tiempo". Esta funcion
    separa esas dos partes: el titulo real, y el nombre del medio.
    """
    m = re.match(r"^(.*)\s+-\s+([^-]+)$", raw_title.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # Si el titulo no tiene esa estructura, se deja tal cual y el medio queda vacio
    # (mas adelante se intenta rellenar el medio de otra forma, ver parse_entries).
    return raw_title.strip(), ""


def _normalize(text: str) -> str:
    """
    "Simplifica" un texto para poder comparar titulos entre si sin que
    importen tildes, mayusculas ni signos de puntuacion. Esto se usa para
    detectar cuando dos noticias distintas en realidad son la misma
    (por ejemplo, si dos medios publicaron el mismo titulo).
    """
    text = unicodedata.normalize("NFKD", text.lower())
    text = "".join(c for c in text if not unicodedata.combining(c))  # quita tildes
    text = re.sub(r"[^a-z0-9 ]", "", text)  # quita signos de puntuacion
    return re.sub(r"\s+", " ", text).strip()


def fetch_feed(url: str, retries: int = 2, timeout: int = 12):
    """
    Descarga los resultados de una busqueda en Google Noticias (en el
    formato RSS). Si falla (por ejemplo, por un problema de conexion),
    lo intenta un par de veces mas antes de rendirse. Si definitivamente
    no se pudo descargar, avisa por consola pero el programa sigue
    funcionando con las demas busquedas (una falla no detiene todo el proceso).
    """
    last_exc = None
    for attempt in range(retries + 1):
        try:
            resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
            resp.raise_for_status()  # si la pagina respondio con un error, lo detecta aqui
            return feedparser.parse(resp.content)
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            time.sleep(1.5 * (attempt + 1))  # espera un poco antes de reintentar
    print(f"  [aviso] no se pudo descargar feed ({last_exc}): {url}", file=sys.stderr)
    return None


def search_feed_url(query: str) -> str:
    """
    Arma la direccion web de busqueda de Google Noticias para una frase
    dada (por ejemplo "vivienda Bogota"). El "when:2d" le pide a Google que
    solo devuelva noticias de los ultimos 2 dias, para no traer resultados
    viejos que de todas formas se van a descartar despues al filtrar por
    franja horaria; esto hace que el programa sea mas rapido.
    """
    q = quote(f"{query} when:2d")
    return f"https://news.google.com/rss/search?q={q}&hl={HL}&gl={GL}&ceid={CEID}"


def parse_entries(feed, categoria: str, start: datetime, end: datetime) -> list[NewsItem]:
    """
    Recorre los resultados que devolvio Google Noticias y se queda
    UNICAMENTE con las noticias publicadas dentro de la franja de horas
    que se esta buscando (por ejemplo, "entre las 9pm de ayer y las 8am de hoy").
    Las que estan fuera de ese rango se descartan aqui mismo.
    """
    items: list[NewsItem] = []
    if not feed or not getattr(feed, "entries", None):
        return items

    for entry in feed.entries:
        # Cada noticia trae su fecha de publicacion; si por algun motivo no
        # la trae, no podemos saber si esta dentro de la franja, asi que se ignora.
        published_struct = getattr(entry, "published_parsed", None)
        if not published_struct:
            continue

        # La fecha viene en horario UTC (horario internacional); se convierte
        # a la hora de Bogota para poder compararla con la franja pedida.
        pub_utc = datetime.fromtimestamp(time.mktime(published_struct), tz=timezone.utc)
        pub_bogota = pub_utc.astimezone(BOGOTA_TZ)

        if not (start <= pub_bogota <= end):
            continue  # esta noticia es mas vieja o mas nueva que la franja pedida

        titulo, medio = _clean_title(entry.title)
        if not medio:
            # Si no se pudo sacar el nombre del medio del titulo, se intenta
            # sacarlo de otro dato que a veces trae la noticia.
            source = getattr(entry, "source", None)
            medio = getattr(source, "title", "") if source else ""

        items.append(
            NewsItem(
                categoria=categoria,
                titular=titulo,
                medio=medio,
                fecha=pub_bogota,
                link=entry.link,
            )
        )
    return items


def dedupe(items: list[NewsItem]) -> list[NewsItem]:
    """
    Como el programa busca con muchas frases distintas dentro del mismo
    tema (por ejemplo, "movilidad Bogota" y "TransMilenio" pueden traer la
    misma noticia dos veces), esta funcion elimina los duplicados: si dos
    noticias tienen el mismo link, o titulos practicamente iguales, solo se
    deja una de las dos en el resultado final.
    """
    seen_links: set[str] = set()
    seen_titles: set[str] = set()
    unique: list[NewsItem] = []
    for it in items:
        norm_title = _normalize(it.titular)
        if it.link in seen_links or norm_title in seen_titles:
            continue  # ya la habiamos visto, se descarta esta copia
        seen_links.add(it.link)
        seen_titles.add(norm_title)
        unique.append(it)
    return unique


# ---------------------------------------------------------------------------
# 3b. "ENRIQUECER" CADA NOTICIA: buscar el link real y un resumen
# ---------------------------------------------------------------------------
# Los links que entrega Google Noticias no son el link final del periodico,
# sino un link intermedio de Google que redirige a la noticia real. Esta
# seccion se encarga de: (1) destapar cual es el link real del periodico,
# y (2) entrar a esa pagina para sacar un pequeno resumen de la noticia
# (el mismo texto que suele aparecer cuando se comparte la noticia en
# redes sociales).

def _resolve_real_link(google_link: str, timeout: int) -> str:
    """
    Convierte el link "envuelto" de Google Noticias en el link real de la
    pagina del periodico. Si por algun motivo no se puede destapar (por
    ejemplo, por un problema de conexion, o porque Google bloquea las
    peticiones que vienen de un servidor en la nube), se deja el link
    original de Google Noticias como respaldo, para que la noticia nunca
    quede sin link -- ver "_link_excel_seguro" en export_to_excel, que se
    encarga de que ese link de respaldo (que puede ser muy largo) nunca
    rompa el Excel.
    """
    try:
        result = gnewsdecoder(google_link, interval=1)
        if result and result.get("status") and result.get("decoded_url"):
            return result["decoded_url"]
        print(f"  [aviso] no se pudo resolver el link real (sin decoded_url): {result}", file=sys.stderr)
    except Exception as exc:  # noqa: BLE001
        print(f"  [aviso] no se pudo resolver el link real ({type(exc).__name__}: {exc})", file=sys.stderr)
    return google_link


def _fetch_summary(url: str, timeout: int) -> str:
    """
    Entra a la pagina de la noticia y busca su "descripcion corta": es un
    dato que casi todas las paginas web incluyen (aunque no se vea en
    pantalla) para que se pueda mostrar un resumen cuando la pagina se
    comparte por WhatsApp, Facebook, etc. Aqui reutilizamos ese mismo texto
    como el "resumen" de la noticia en el Excel.

    Si la pagina no tiene ese dato, o no se pudo entrar a ella, se devuelve
    un texto vacio (y mas adelante se usa el titulo como resumen de respaldo).
    """
    try:
        resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        tag = soup.find("meta", attrs={"name": "description"}) or soup.find(
            "meta", attrs={"property": "og:description"}
        )
        if tag and tag.get("content"):
            return " ".join(tag["content"].split())  # limpia espacios/saltos de linea de sobra
    except Exception:
        pass
    return ""


def enrich_item(item: NewsItem, timeout: int = 10) -> NewsItem:
    """Procesa UNA noticia: le busca el link real y el resumen."""
    item.enlace = _resolve_real_link(item.link, timeout)
    # Si no se encontro un resumen real, se usa el titulo como resumen
    # de respaldo, para que esa columna nunca quede vacia en el Excel.
    item.resumen = _fetch_summary(item.enlace, timeout) or item.titular
    return item


def enrich_items(items: list[NewsItem], max_workers: int = 15) -> list[NewsItem]:
    """
    Aplica "enrich_item" a TODAS las noticias encontradas. Esto implica
    visitar la pagina real de cada noticia, lo cual toma tiempo si se hace
    de a una por una. Por eso se procesan varias noticias EN SIMULTANEO
    (hasta 15 al mismo tiempo) para que el programa no se demore demasiado.
    Va imprimiendo el avance en pantalla (10, 20, 30... noticias procesadas)
    para que se vea que el programa sigue trabajando y no que se colgo.
    """
    if not items:
        return items
    print(f"Resolviendo enlace real y resumen de {len(items)} noticia(s)...")
    done = 0
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = [pool.submit(enrich_item, item) for item in items]
        for future in as_completed(futures):
            future.result()
            done += 1
            if done % 10 == 0 or done == len(items):
                print(f"  {done}/{len(items)} procesadas")
    return items


# ---------------------------------------------------------------------------
# 4. RECOLECCION: unir todas las busquedas en una sola lista de noticias
# ---------------------------------------------------------------------------

def collect_bogota_news(start: datetime, end: datetime, queries_por_categoria: dict[str, list[str]],
                         max_workers: int = 8) -> list[NewsItem]:
    """
    Recorre TODOS los temas y frases de busqueda que se leyeron del Excel
    "Palabras_Clave_Bogota.xlsx" (parametro "queries_por_categoria"),
    descarga los resultados de cada una, se queda solo con las noticias
    dentro de la franja horaria pedida, y al final elimina los duplicados.
    Devuelve la lista final de noticias de Bogota, ordenada de la mas
    reciente a la mas antigua.
    """
    # Se arma una lista de tareas: (tema, frase de busqueda) por cada
    # combinacion que haya en el Excel de palabras clave.
    tasks = [
        (categoria, query)
        for categoria, queries in queries_por_categoria.items()
        for query in queries
    ]

    all_items: list[NewsItem] = []
    print(f"Buscando noticias de Bogota en {len(tasks)} consultas...")

    # Igual que con "enrich_items", las busquedas se hacen varias al mismo
    # tiempo (hasta 8 a la vez) para no demorar el programa innecesariamente.
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_map = {
            pool.submit(fetch_feed, search_feed_url(query)): (categoria, query)
            for categoria, query in tasks
        }
        for future in as_completed(future_map):
            categoria, query = future_map[future]
            feed = future.result()
            found = parse_entries(feed, categoria, start, end)
            if found:
                print(f"  [{categoria}] '{query}': {len(found)} noticia(s) en la franja")
            all_items.extend(found)

    unique_items = dedupe(all_items)
    unique_items.sort(key=lambda it: it.fecha, reverse=True)  # mas recientes primero
    print(f"Total Bogota tras deduplicar: {len(unique_items)}")
    return unique_items


def collect_national_news(top_n: int = 3) -> list[NewsItem]:
    """
    Trae las noticias nacionales de Colombia mas destacadas de HOY (segun
    el orden en que Google Noticias las muestra, que ya viene ordenado por
    importancia), y se queda solo con las primeras "top_n" (por defecto 3).
    """
    today = datetime.now(BOGOTA_TZ).date()
    feed = fetch_feed(NATIONAL_TOPICS_FEED)
    items = parse_entries(
        feed,
        "Nacional Colombia",
        start=datetime.combine(today, datetime.min.time(), tzinfo=BOGOTA_TZ),
        end=datetime.now(BOGOTA_TZ),
    )
    items = dedupe(items)
    return items[:top_n]


# ---------------------------------------------------------------------------
# 5. ARMAR EL ARCHIVO DE EXCEL
# ---------------------------------------------------------------------------

# Estas son las columnas que va a tener cada hoja del Excel con noticias.
COLUMNS = ["categoria", "fuente", "titulo", "resumen", "fecha", "enlace"]


def to_dataframe(items: list[NewsItem]) -> pd.DataFrame:
    """
    Convierte la lista de noticias (las "fichas" NewsItem) en una tabla
    (lo que en programacion se llama "DataFrame"), que es basicamente lo
    mismo que una hoja de Excel pero manejada por el programa antes de
    guardarla como archivo.
    """
    if not items:
        return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(
        [
            {
                "categoria": it.categoria,
                "fuente": it.medio,
                "titulo": it.titular,
                "resumen": it.resumen,
                "fecha": it.fecha.strftime("%Y-%m-%d %H:%M"),
                "enlace": it.enlace or it.link,  # si no se pudo resolver el link real, se usa el original
            }
            for it in items
        ]
    )


def export_to_excel(bogota_items: list[NewsItem], national_items: list[NewsItem],
                     window_label: str, start: datetime, end: datetime) -> str:
    """
    Junta todo lo recolectado y genera el archivo de Excel final, con tres
    hojas (pestanas):

      - "Resumen": informacion general de la corrida (a que hora se hizo,
        que franja se aplico, cuantas noticias se encontraron).
      - "Bogota": todas las noticias de Bogota encontradas.
      - "Nacional (Top 3)": las 3 noticias nacionales mas importantes.

    Ademas, deja los links como hipervinculos que se pueden hacer clic
    directamente desde Excel, y ajusta el ancho de las columnas para que
    el texto se vea bien sin tener que ajustarlo a mano.
    """
    # El nombre del archivo incluye la fecha y hora de creacion, para que
    # cada vez que se corra el programa se genere un archivo NUEVO y nunca
    # se borre sin querer un reporte anterior.
    timestamp = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d_%H%M")
    filename = f"Monitoreo_Noticias_{timestamp}.xlsx"

    df_bogota = to_dataframe(bogota_items)
    df_nacional = to_dataframe(national_items)

    # Tabla pequena con el resumen de la corrida, para la hoja "Resumen".
    resumen = pd.DataFrame(
        {
            "Campo": ["Fecha de ejecucion", "Franja aplicada", "Desde", "Hasta",
                      "Total noticias Bogota", "Total noticias nacionales"],
            "Valor": [
                datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d %H:%M"),
                window_label,
                start.strftime("%Y-%m-%d %H:%M"),
                end.strftime("%Y-%m-%d %H:%M"),
                len(df_bogota),
                len(df_nacional),
            ],
        }
    )

    # Se abre el archivo de Excel "en modo escritura" y se van agregando
    # las tres hojas una por una.
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        df_bogota.to_excel(writer, sheet_name="Bogota", index=False)
        df_nacional.to_excel(writer, sheet_name="Nacional (Top 3)", index=False)

        from openpyxl.styles import Font

        # Excel solo admite hipervinculos de hasta 255 caracteres: si el
        # link es mas largo que eso, Excel lo guarda igual como texto en la
        # celda, pero al hacerle clic da error ("no se puede abrir el
        # archivo especificado"). Esto pasa con los links de respaldo de
        # Google Noticias cuando no se pudo destapar el link real (son
        # larguisimos). Por eso solo se activa como link clickeable cuando
        # es corto; si es largo, se deja como texto plano -- se puede
        # copiar y pegar en el navegador igual, solo que no es clickeable
        # directamente desde Excel.
        LARGO_MAXIMO_HIPERVINCULO = 255

        # Para las hojas de noticias: se convierte la columna "enlace" en un
        # link real de Excel (azul y subrayado, se puede hacer clic), y se
        # ajusta el ancho de cada columna segun el largo del texto que tiene.
        for sheet_name, df in (("Bogota", df_bogota), ("Nacional (Top 3)", df_nacional)):
            ws = writer.sheets[sheet_name]
            if "enlace" in df.columns:
                link_col = df.columns.get_loc("enlace") + 1
                for row_idx in range(2, len(df) + 2):  # la fila 1 es el encabezado, por eso se empieza en la 2
                    cell = ws.cell(row=row_idx, column=link_col)
                    if cell.value and len(str(cell.value)) <= LARGO_MAXIMO_HIPERVINCULO:
                        cell.hyperlink = cell.value
                        cell.font = Font(color="0563C1", underline="single")
            for col_cells in ws.columns:
                length = max((len(str(c.value)) if c.value else 0) for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 90)
            ws.freeze_panes = "A2"  # deja fija la fila de encabezados al desplazarse hacia abajo

        # Igual, pero para la hoja de Resumen (esta no tiene links).
        for sheet_name in ("Resumen",):
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                length = max((len(str(c.value)) if c.value else 0) for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 60)

    return filename


# ---------------------------------------------------------------------------
# 6. PROGRAMA PRINCIPAL (lo primero que se ejecuta al correr el archivo)
# ---------------------------------------------------------------------------

def main():
    """
    Esta es la funcion que organiza todo el proceso, de principio a fin:

      1) Revisa si se pidieron opciones especiales (--now o --window).
      2) Lee las categorias y palabras de busqueda del Excel de palabras clave.
      3) Calcula la franja horaria a usar.
      4) Busca las noticias de Bogota.
      5) Busca las noticias nacionales.
      6) Busca el link real y el resumen de cada noticia encontrada.
      7) Genera el archivo de Excel con todo.

    Es la funcion que se llama automaticamente al final de este archivo
    (ver la ultima linea: if __name__ == "__main__": main()).
    """
    # Se definen las opciones que se pueden escribir al ejecutar el programa
    # desde la terminal (por ejemplo: python monitoreo_noticias.py --window 12).
    parser = argparse.ArgumentParser(description="Monitoreo de noticias Bogota + Colombia")
    parser.add_argument("--now", type=str, default=None,
                         help="Forzar fecha/hora de ejecucion, ej: '2026-07-28 08:30' (hora Bogota)")
    parser.add_argument("--window", type=float, default=None,
                         help="Forzar ventana en horas hacia atras desde --now/ahora")
    args = parser.parse_args()

    now = None
    if args.now:
        now = datetime.strptime(args.now, "%Y-%m-%d %H:%M").replace(tzinfo=BOGOTA_TZ)

    # Paso 1: leer (o crear, si no existe) el Excel con las categorias y
    # palabras de busqueda de Bogota.
    queries_por_categoria = load_bogota_queries(KEYWORDS_FILE)
    total_palabras = sum(len(v) for v in queries_por_categoria.values())
    print(f"Palabras clave cargadas desde '{KEYWORDS_FILE}': "
          f"{len(queries_por_categoria)} categorias, {total_palabras} palabras")

    # Deja el codigo (este mismo archivo .py) con una copia de respaldo de
    # las palabras clave actuales, por si el Excel se llegara a perder.
    sync_default_queries_in_code(queries_por_categoria, os.path.abspath(__file__))
    print()

    # Paso 2: decidir el rango de horas a buscar.
    start, end, label = get_time_window(now=now, force_hours=args.window)
    print(f"Ventana aplicada: {label}")
    print(f"  Desde: {start:%Y-%m-%d %H:%M}")
    print(f"  Hasta: {end:%Y-%m-%d %H:%M}\n")

    # Paso 3: buscar las noticias (Bogota y luego nacionales).
    bogota_items = collect_bogota_news(start, end, queries_por_categoria)
    national_items = collect_national_news(top_n=3)

    # Paso 4: para cada noticia encontrada, buscar su link real y su resumen.
    enrich_items(bogota_items + national_items)

    # Paso 5: guardar todo en el archivo de Excel.
    filename = export_to_excel(bogota_items, national_items, label, start, end)
    print(f"\nListo. Excel generado: {filename}")
    print(f"  - Bogota: {len(bogota_items)} noticias")
    print(f"  - Nacional: {len(national_items)} noticias")


# Esta linea hace que "main()" se ejecute automaticamente cuando alguien
# corre este archivo directamente (por ejemplo, con "python monitoreo_noticias.py").
if __name__ == "__main__":
    main()
