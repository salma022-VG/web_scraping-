"""
Recolector de noticias LOCALES/NACIONALES sobre catastro: catastro multipropósito,
predios, avalúos catastrales, UAECD, IDECA, cartografía catastral, etc. (últimos 90
días por defecto, sin exigir ciudad específica). Fuentes: El Tiempo, El Espectador,
Semana, RCN Radio, Publimetro, Infobae, La República, Portafolio, Dinero, Valora
Analitik, El País (Cali), El Heraldo, Vanguardia, Q'hubo y La FM.

Control de bots: motor HTTP Scrapling (impersonación de huella TLS/JA3 de navegadores
reales + cabeceras "stealth" coherentes, mucho más difícil de detectar que `requests`
con solo un User-Agent falso), reintentos, rate-limiting por dominio y respeto a
robots.txt (ver PoliteSession).

Uso:
    python catastro_scraper.py
    python catastro_scraper.py --days 30 --output noticias.xlsx
    python catastro_scraper.py --keywords catastro "avaluo catastral" UAECD
    python catastro_scraper.py --any-of bogota   # acotar solo a Bogotá
    python catastro_scraper.py --verbose --delay-min 3 --delay-max 6
"""

from __future__ import annotations

import argparse
import calendar
import logging
import random
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib import robotparser
from urllib.parse import quote, urljoin, urlparse

import feedparser
import pandas as pd
import trafilatura
from bs4 import BeautifulSoup
from scrapling import Fetcher

logger = logging.getLogger("catastro_scraper")


# ============================================================================
# 1. CONTROL DE BOTS: sesión HTTP con rotación de UA, reintentos, rate-limit
#    y respeto a robots.txt.
# ============================================================================

# Perfiles de navegador para impersonar (huella TLS/JA3 + cabeceras coherentes).
# Scrapling genera automáticamente el resto de cabeceras (Sec-Ch-Ua, Accept, etc.)
# consistentes con el perfil elegido -- mucho más difícil de detectar que fijar
# solo un User-Agent de texto sobre una conexión TLS de `requests`/Python.
IMPERSONATE_PROFILES = ["chrome131", "chrome124", "chrome120", "firefox135", "safari184"]


class _NormalizedResponse:
    """Adapta la respuesta de Scrapling a la interfaz (.status_code/.text/.content/.url)
    que usa el resto del scraper, para no acoplar todo el código a su API."""

    __slots__ = ("status_code", "text", "content", "url", "headers")

    def __init__(self, scrapling_response):
        self.status_code = scrapling_response.status
        self.content = scrapling_response.body
        encoding = scrapling_response.encoding or "utf-8"
        self.text = scrapling_response.body.decode(encoding, errors="replace")
        self.url = scrapling_response.url
        self.headers = scrapling_response.headers


class PoliteSession:
    """Envoltorio sobre Scrapling que aplica impersonación de navegador, rate-limit
    por dominio y respeto a robots.txt."""

    def __init__(
        self,
        delay_range: tuple[float, float] = (2.0, 5.0),
        respect_robots: bool = True,
        timeout: int = 15,
    ):
        self.delay_range = delay_range
        self.respect_robots = respect_robots
        self.timeout = timeout
        self._last_request_time: dict[str, float] = {}
        self._robots_cache: dict[str, robotparser.RobotFileParser | None] = {}

    @staticmethod
    def _domain(url: str) -> str:
        return urlparse(url).netloc

    def _raw_get(self, url: str, allow_redirects: bool = True) -> _NormalizedResponse | None:
        resp = Fetcher.get(
            url,
            impersonate=random.choice(IMPERSONATE_PROFILES),
            stealthy_headers=True,
            headers={"Accept-Language": "es-CO,es;q=0.9,en;q=0.8"},
            timeout=self.timeout,
            retries=2,
            follow_redirects=allow_redirects,
        )
        if resp.status >= 400:
            raise RuntimeError(f"HTTP {resp.status}")
        return _NormalizedResponse(resp)

    def _can_fetch(self, url: str) -> bool:
        if not self.respect_robots:
            return True
        domain = self._domain(url)
        if domain not in self._robots_cache:
            robots_url = urljoin(f"https://{domain}", "/robots.txt")
            rp = robotparser.RobotFileParser()
            rp.set_url(robots_url)
            try:
                # No usamos rp.read(): internamente hace su propia petición con el
                # User-Agent genérico de urllib, que muchos sitios bloquean con 403.
                # Eso provoca que RobotFileParser asuma "disallow all" por defecto,
                # un falso negativo. En su lugar, lo pedimos nosotros mismos
                # (con impersonación realista) y se lo pasamos ya descargado.
                resp = self._raw_get(robots_url)
                rp.parse(resp.text.splitlines())
            except RuntimeError as exc:
                if "404" in str(exc):
                    self._robots_cache[domain] = None  # sin robots.txt -> todo permitido
                    return True
                self._robots_cache[domain] = None
                return True
            except Exception:
                # Si no se puede obtener/parsear robots.txt, se asume acceso permitido
                # (comportamiento estándar cuando el archivo no existe o no es accesible).
                self._robots_cache[domain] = None
                return True
            self._robots_cache[domain] = rp
        rp = self._robots_cache[domain]
        if rp is None:
            return True
        return rp.can_fetch("Mozilla/5.0 (compatible)", url)

    def _throttle(self, url: str) -> None:
        domain = self._domain(url)
        wait = random.uniform(*self.delay_range)
        last = self._last_request_time.get(domain)
        if last is not None:
            elapsed = time.time() - last
            if elapsed < wait:
                time.sleep(wait - elapsed)
        self._last_request_time[domain] = time.time()

    def get(self, url: str, allow_redirects: bool = True, **kwargs) -> _NormalizedResponse | None:
        """GET con throttling, verificación de robots.txt y manejo de errores."""
        if not self._can_fetch(url):
            logger.warning("robots.txt deniega el acceso a %s; se omite.", url)
            return None
        self._throttle(url)
        try:
            return self._raw_get(url, allow_redirects=allow_redirects)
        except Exception as exc:
            logger.warning("Error solicitando %s: %s", url, exc)
            return None


# ============================================================================
# 2. CONFIGURACIÓN: fuentes de noticias y vocabulario temático de catastro.
#
# Todas las URLs de fuentes fueron verificadas manualmente (respuesta 200 con
# contenido válido) antes de incluirlas. Aun así, los medios cambian sus feeds
# con el tiempo: si una fuente deja de responder, el scraper la omite con una
# advertencia y continúa con las demás (ver run()).
#
# Nota sobre Google News: se evaluó usarlo como agregador de respaldo, pero su
# robots.txt bloquea explícitamente a bots automatizados (incluido "Claude-Web" /
# "anthropic-ai") para la ruta de búsqueda RSS. Como este scraper respeta
# robots.txt por diseño, se optó por NO incluirlo y cubrir cada medio directamente.
# ============================================================================


@dataclass
class SourceConfig:
    name: str
    url: str
    type: str  # "rss" | "html"
    enrich: bool = False  # fuerza descargar el artículo completo para el resumen
    city_scoped: bool = False  # True si la fuente ya está acotada a una ciudad (omite el filtro "any_of")


# --- Feeds RSS directos de medios nacionales (prensa, radio y economía) ---
_RSS_SOURCES = [
    SourceConfig("El Tiempo - Colombia", "https://www.eltiempo.com/rss/colombia.xml", "rss"),
    SourceConfig("El Tiempo - Economía", "https://www.eltiempo.com/rss/economia.xml", "rss"),
    SourceConfig("El Tiempo - Bogotá", "https://www.eltiempo.com/rss/bogota.xml", "rss", city_scoped=True),
    SourceConfig(
        "El Espectador",
        "https://www.elespectador.com/arc/outboundfeeds/discover/?outputType=xml",
        "rss",
    ),
    SourceConfig("Semana", "https://www.semana.com/arc/outboundfeeds/rss/", "rss"),
    SourceConfig("RCN Radio", "https://www.rcnradio.com/rss", "rss"),
    SourceConfig("Publimetro Colombia", "https://www.publimetro.co/arc/outboundfeeds/rss/", "rss"),
    SourceConfig(
        "Infobae Colombia",
        "https://www.infobae.com/arc/outboundfeeds/rss/category/colombia/",
        "rss",
    ),
    SourceConfig("La República", "https://www.larepublica.co/rss", "rss"),
    SourceConfig("Portafolio - Economía", "https://www.portafolio.co/rss/economia.xml", "rss"),
    SourceConfig("Portafolio - Gobierno", "https://www.portafolio.co/rss/economia/gobierno.xml", "rss"),
    SourceConfig("Dinero", "https://www.dinero.com/arc/outboundfeeds/rss/", "rss"),
    SourceConfig("Valora Analitik", "https://www.valoraanalitik.com/feed/", "rss"),
    SourceConfig("El País (Cali)", "https://www.elpais.com.co/arc/outboundfeeds/rss/", "rss"),
    SourceConfig("El Heraldo (Barranquilla)", "https://www.elheraldo.co/arc/outboundfeeds/rss/", "rss"),
    SourceConfig("Vanguardia (Bucaramanga)", "https://www.vanguardia.com/arc/outboundfeeds/rss/", "rss"),
]

# --- Sitios con buscador propio pero sin RSS confiable: se arma la URL de búsqueda
#     con cada término de DEFAULT_SEARCH_QUERIES (un subconjunto corto y de alto
#     rendimiento del vocabulario completo; usar los ~60 términos completos como
#     consulta de búsqueda en cada sitio dispararía cientos de peticiones). ---
# NOTA: se verificó cada sitio comparando resultados para dos búsquedas muy distintas
# ("catastro" vs "futbol"). Caracol Radio, W Radio, Blu Radio y El Colombiano
# devuelven exactamente el mismo contenido sin importar la consulta (su "?s=" no
# filtra nada -- se confirmó también con Playwright y con Scrapling StealthyFetcher,
# así que no es un tema de detección de bots, el endpoint simplemente no funciona) y
# se excluyeron. Solo Q'hubo y La FM sí devuelven resultados distintos por consulta.
SEARCH_SITES = [
    ("Q'hubo", "https://www.qhubo.com/?s={query}"),
    ("La FM", "https://www.lafm.com.co/buscar/{query}"),
]

DEFAULT_SEARCH_QUERIES = [
    "catastro bogota",
    "avaluo catastral",
    "UAECD",
    "catastro multiproposito",
    "actualizacion catastral",
    "predios bogota",
    "catastro",
    "predios",
    "impuesto predial",
    "avaluo predio",
]


def build_sources(search_terms: list[str] | None = None) -> list[SourceConfig]:
    """Arma la lista final de fuentes: los feeds RSS fijos + una fuente HTML por
    cada combinación (sitio de búsqueda, término de `search_terms`)."""
    terms = search_terms if search_terms is not None else DEFAULT_SEARCH_QUERIES
    sources = list(_RSS_SOURCES)
    for name, template in SEARCH_SITES:
        for term in terms:
            url = template.format(query=quote(term))
            sources.append(SourceConfig(f"{name} - búsqueda \"{term}\"", url, "html"))
    return sources


# --- Vocabulario temático de catastro (para filtrar relevancia de artículos) ---
# Cada frase se expande automáticamente:
#   - "A / B"        -> dos temas independientes: "A" y "B"
#   - "A (SIGLA)"     -> dos temas: "A" y "SIGLA"
#   - "A (x, y, z)"   -> temas "A x", "A y", "A z"
_RAW_TOPICS = [
    "Catastro Bogotá",
    "Unidad Administrativa Especial de Catastro Distrital (UAECD)",
    "Catastro multipropósito",
    "Base de datos catastral",
    "Censo inmobiliario",
    "Actualización catastral",
    "Información predial",
    "Infraestructura de Datos Espaciales (IDECA)",
    "Sistema de información geográfica (SIG)",
    "Gestión catastral",
    "Inventario de predios",
    "Predio urbano",
    "Predio rural",
    "Predio mixto",
    "Número predial nacional (NPN)",
    "CHIP catastral",
    "Matrícula inmobiliaria",
    "Linderos",
    "Área de terreno",
    "Área construida",
    "Uso del suelo",
    "Clasificación del suelo",
    "Propiedad horizontal",
    "Unidad predial",
    "Titularidad / propietario",
    "Bien inmueble",
    "Avalúo catastral",
    "Avalúo comercial",
    "Avalúo de referencia",
    "Avalúo de renta",
    "Valor del suelo",
    "Valor de la construcción",
    "Valor por metro cuadrado",
    "Estimación del valor del inmueble",
    "Métodos de valuación",
    "Enfoque de mercado",
    "Enfoque de ingresos",
    "Enfoque de costos",
    "Ubicación del predio",
    "Estrato socioeconómico",
    "Área del lote",
    "Antigüedad de la construcción",
    "Uso (residencial, comercial, industrial)",
    "Estado físico del inmueble",
    "Accesibilidad",
    "Equipamientos cercanos",
    "Dinámica del mercado inmobiliario",
    "Oferta y demanda",
    "Zonificación",
    "Cartografía catastral",
    "Shapefile predial",
    "Capas geográficas",
    "Georreferenciación",
    "Coordenadas",
    "Mapas de calor",
    "Análisis espacial",
    "Zonificación urbana",
    "Manzanas catastrales",
    "Polígonos prediales",
    "Visualización geográfica",
]

 
def _expand_topic(phrase: str) -> list[str]:
    results: list[str] = []
    for alt in (p.strip() for p in phrase.split("/")):
        if not alt:
            continue
        paren = re.search(r"\(([^)]+)\)", alt)
        base = re.sub(r"\([^)]*\)", "", alt).strip()
        content = paren.group(1).strip() if paren else ""
        if "," in content:
            # p. ej. "Uso (residencial, comercial, industrial)" -> "Uso residencial", ...
            # (no se agrega "Uso" suelto: es una sola palabra demasiado genérica)
            for item in content.split(","):
                item = item.strip()
                if item and base:
                    results.append(f"{base} {item}")
        else:
            if base:
                results.append(base)
            if content and len(re.sub(r"[^a-zA-Záéíóúñ]", "", content)) >= 3:
                # sigla razonable (UAECD, IDECA, SIG, NPN, CHIP...); se descartan
                # símbolos demasiado cortos/no alfabéticos (p. ej. "m²").
                results.append(content)
    return results


def _build_default_topics() -> list[str]:
    topics: list[str] = []
    for phrase in _RAW_TOPICS:
        for topic in _expand_topic(phrase):
            if topic not in topics:
                topics.append(topic)
    return topics


DEFAULT_TOPICS: list[str] = _build_default_topics()


# ============================================================================
# 3. EXTRACCIÓN: parseo de RSS/HTML, filtro de relevancia y resumen enriquecido.
# ============================================================================

_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
_TOKEN_RE = re.compile(r"[a-z0-9]+")

# Siglas de una sola palabra que son lo bastante específicas por sí solas.
_ACRONYMS = {"uaecd", "ideca", "sig", "npn", "chip"}

# Palabras "núcleo" que anclan un texto al dominio catastral. Deliberadamente NO
# incluye "inmueble" ni "inmobiliario": son demasiado comunes en cualquier noticia
# sobre edificios (robos, incendios, ventas...) y no discriminan el tema catastral.
_CORE_CONTEXT = ["catastro", "catastral", "predio", "predial", "avaluo", "cartografia"]


@dataclass
class Article:
    fuente: str
    titulo: str
    resumen: str
    fecha: str | None
    enlace: str


def _looks_like_article_url(url: str) -> bool:
    """Heurística para descartar enlaces de navegación/pie de página (créditos, secciones,
    redes sociales) que no son noticias: exige un slug largo y con guiones en la ruta,
    típico de URLs de artículos (p. ej. /2026/07/27/una-noticia-larga-con-titulo/)."""
    path = urlparse(url).path.strip("/")
    if not path:
        return False
    last_segment = path.split("/")[-1]
    return len(last_segment) > 25 and last_segment.count("-") >= 2


def _clean_html(text: str) -> str:
    return _TAG_RE.sub(" ", text or "").strip()


def _shorten(text: str, max_len: int = 320) -> str:
    text = _WS_RE.sub(" ", text or "").strip()
    if len(text) <= max_len:
        return text
    return text[:max_len].rsplit(" ", 1)[0] + "..."


def _strip_leading_title(resumen: str, title: str) -> str:
    """trafilatura suele incluir el titular como primera línea del texto extraído;
    se quita para no duplicarlo delante del resumen."""
    resumen_norm = _WS_RE.sub(" ", resumen).strip()
    title_norm = _WS_RE.sub(" ", title).strip()
    if title_norm and resumen_norm.lower().startswith(title_norm.lower()):
        return resumen_norm[len(title_norm):].strip(" -–:|")
    return resumen_norm


def _strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def _normalize(text: str) -> str:
    return _strip_accents(text or "").lower()


def _singularize(token: str) -> str:
    """Singularización simple en español: 'catastrales'->'catastral' (consonante+'es'),
    'avalúos'/'predios'->'avalúo'/'predio' (vocal+'s')."""
    if token.endswith("es") and len(token) > 4:
        return token[:-2]
    if token.endswith("s") and len(token) > 3:
        return token[:-1]
    return token


def _singularized_text(text: str) -> str:
    """Texto normalizado (sin tildes, singularizado) como cadena de tokens separados
    por espacios. Se compara por SUBCADENA ADYACENTE (no por presencia suelta de cada
    palabra en cualquier parte del texto): así "Bien inmueble" no casa por accidente
    en un artículo que en algún punto dice "bien" y, párrafos después, "inmueble" sin
    relación entre sí -- solo casa si aparecen juntas, como frase real."""
    tokens = _TOKEN_RE.findall(_normalize(text))
    return " " + " ".join(_singularize(t) for t in tokens) + " "


def _phrase_key(phrase: str) -> str:
    words = _TOKEN_RE.findall(_normalize(phrase))
    return " " + " ".join(_singularize(w) for w in words) + " "


def _phrase_matches(phrase: str, singularized_text: str) -> bool:
    key = _phrase_key(phrase)
    return key.strip() != "" and key in singularized_text


def _topic_matches(topic: str, singularized_text: str) -> bool:
    if not _phrase_matches(topic, singularized_text):
        return False
    # Palabras sueltas y genéricas del vocabulario (p. ej. "propietario", "linderos",
    # "coordenadas", "accesibilidad") pueden aparecer en noticias sin relación con
    # catastro (ej. "el propietario de un carro robado en Bogotá"). Para esos temas
    # de una sola palabra que no son siglas, se exige además que el texto mencione
    # algún término núcleo de catastro, genuinamente específico del dominio (no
    # "inmueble", que es demasiado común en cualquier noticia sobre edificios).
    # Este guard NO aplica al filtro de ciudad (any_of), que usa _phrase_matches
    # directamente.
    words = _TOKEN_RE.findall(_normalize(topic))
    if len(words) == 1 and words[0] not in _ACRONYMS:
        return any(_phrase_matches(c, singularized_text) for c in _CORE_CONTEXT)
    return True


def is_relevant(text: str, topics: list[str], any_of: list[str] | None = None) -> bool:
    """Relevante si ALGUNO de los `topics` aparece como frase adyacente en el texto
    (tolerando singular/plural/tildes) y, si se da `any_of`, al menos uno de esos
    términos también aparece (p. ej. para acotar por ciudad)."""
    stext = _singularized_text(text)
    if not any(_topic_matches(t, stext) for t in topics):
        return False
    if any_of:
        return any(_phrase_matches(a, stext) for a in any_of)
    return True


def _entry_datetime(entry) -> datetime | None:
    for key in ("published_parsed", "updated_parsed"):
        struct = entry.get(key)
        if struct:
            return datetime.fromtimestamp(calendar.timegm(struct), tz=timezone.utc)
    return None


def _parse_iso_date(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def enrich_article(
    session: PoliteSession, url: str
) -> tuple[str | None, str | None, str | None, datetime | None]:
    """Descarga el artículo real y extrae (resumen, url_final, nombre_del_sitio, fecha_publicacion)."""
    resp = session.get(url, allow_redirects=True)
    if resp is None:
        return None, None, None, None
    final_url = resp.url
    text = trafilatura.extract(resp.text, favor_recall=True, include_comments=False)
    site_name = None
    fecha_dt = None
    try:
        metadata = trafilatura.extract_metadata(resp.text)
        if metadata:
            site_name = metadata.sitename or metadata.hostname
            fecha_dt = _parse_iso_date(metadata.date)
    except Exception:
        pass
    if not text:
        return None, final_url, site_name, fecha_dt
    return text[:600], final_url, site_name, fecha_dt


def parse_rss(
    session: PoliteSession,
    source: SourceConfig,
    topics: list[str],
    any_of: list[str] | None,
    cutoff: datetime | None = None,
    max_items: int = 20,
) -> list[Article]:
    resp = session.get(source.url)
    if resp is None:
        return []
    feed = feedparser.parse(resp.content)
    if feed.bozo and not feed.entries:
        logger.warning("Feed inválido o vacío en %s", source.url)
        return []

    effective_any_of = None if source.city_scoped else any_of

    articles: list[Article] = []
    for entry in feed.entries:
        title = entry.get("title", "").strip()
        summary_raw = _clean_html(entry.get("summary", "") or entry.get("description", ""))
        combined = f"{title} {summary_raw}"
        if not is_relevant(combined, topics, effective_any_of):
            continue

        entry_dt = _entry_datetime(entry)
        link = entry.get("link", "")
        fecha = entry.get("published", "") or entry.get("updated", "") or None
        fuente = source.name
        resumen = summary_raw

        # Si el resumen del feed es muy corto, si la fuente lo exige explícitamente,
        # o si aún no se sabe la fecha de publicación (se necesita para el filtro de
        # antigüedad), se descarga el artículo real.
        need_enrich = source.enrich or len(resumen) < 80 or entry_dt is None
        if need_enrich:
            enriched, real_link, real_source, meta_dt = enrich_article(session, link)
            if enriched:
                resumen = _strip_leading_title(enriched, title)
            if real_link:
                link = real_link
            if real_source and source.enrich:
                fuente = real_source
            if entry_dt is None and meta_dt is not None:
                entry_dt = meta_dt

        if cutoff is not None and entry_dt is not None and entry_dt < cutoff:
            continue

        articles.append(
            Article(fuente=fuente, titulo=title, resumen=_shorten(resumen), fecha=fecha, enlace=link)
        )
        if len(articles) >= max_items:
            break
    return articles


def parse_html_listing(
    session: PoliteSession,
    source: SourceConfig,
    topics: list[str],
    any_of: list[str] | None,
    cutoff: datetime | None = None,
    max_items: int = 20,
) -> list[Article]:
    """Scraping genérico: busca enlaces cuyo texto visible contenga las palabras clave."""
    resp = session.get(source.url)
    if resp is None:
        return []
    soup = BeautifulSoup(resp.text, "html.parser")

    effective_any_of = None if source.city_scoped else any_of

    seen: set[str] = set()
    articles: list[Article] = []
    for a in soup.find_all("a", href=True):
        title = a.get_text(strip=True)
        if not title or len(title) < 15:
            continue
        if not is_relevant(title, topics, effective_any_of):
            continue
        link = urljoin(source.url, a["href"])
        if link in seen or not _looks_like_article_url(link):
            continue
        seen.add(link)

        resumen, real_link, _, fecha_dt = enrich_article(session, link)
        # Si no se pudo determinar la fecha de publicación, se conserva el artículo
        # (mejor esfuerzo): es un resultado de búsqueda actual, probablemente reciente.
        if cutoff is not None and fecha_dt is not None and fecha_dt < cutoff:
            continue
        resumen = _strip_leading_title(resumen or "", title)
        articles.append(
            Article(
                fuente=source.name,
                titulo=title,
                resumen=_shorten(resumen),
                fecha=fecha_dt.isoformat() if fecha_dt else None,
                enlace=real_link or link,
            )
        )
        if len(articles) >= max_items:
            break
    return articles


# ============================================================================
# 4. CLI / ORQUESTADOR
# ============================================================================


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Recolecta noticias locales/nacionales sobre catastro en Bogotá "
        "(catastro multipropósito, predios, avalúos catastrales, UAECD, IDECA, "
        "cartografía catastral, etc.)."
    )
    p.add_argument(
        "--keywords", nargs="+", default=DEFAULT_TOPICS,
        help="Temas a buscar (basta con que aparezca UNO). Frases de varias palabras "
             "deben ir entre comillas, p. ej. \"avaluo catastral\". Por defecto se usa "
             f"el vocabulario catastral completo ({len(DEFAULT_TOPICS)} términos).",
    )
    p.add_argument(
        "--any-of", nargs="+", default=None,
        help="Opcional: exige que al menos uno de estos términos también aparezca "
             "(filtro de ciudad/región, p. ej. --any-of bogota). Desactivado por "
             "defecto para maximizar la cobertura, dado que las noticias específicas "
             "de catastro son poco frecuentes.",
    )
    p.add_argument(
        "--search-queries", nargs="+", default=DEFAULT_SEARCH_QUERIES,
        help="Términos que se escriben en el buscador de los sitios sin RSS (Q'hubo, "
             "La FM). Se prueba cada uno en cada sitio, así que conviene mantener "
             "esta lista corta.",
    )
    p.add_argument("--days", type=int, default=90, help="Solo noticias de los últimos N días. Por defecto: 90 (~3 meses).")
    p.add_argument("--max-per-source", type=int, default=25, help="Máximo de artículos por fuente.")
    p.add_argument("--delay-min", type=float, default=2.0, help="Espera mínima (s) entre solicitudes al mismo dominio.")
    p.add_argument("--delay-max", type=float, default=5.0, help="Espera máxima (s) entre solicitudes al mismo dominio.")
    p.add_argument("--output", default="noticias_catastro.csv", help="Archivo de salida (.csv o .xlsx).")
    p.add_argument("--no-robots", action="store_true", help="No respetar robots.txt (no recomendado).")
    p.add_argument("--verbose", action="store_true", help="Log detallado.")
    return p


def run(args: argparse.Namespace) -> pd.DataFrame:
    session = PoliteSession(
        delay_range=(args.delay_min, args.delay_max),
        respect_robots=not args.no_robots,
    )
    any_of = [a for a in (args.any_of or []) if a] or None
    sources = build_sources(args.search_queries)
    cutoff = datetime.now(timezone.utc) - timedelta(days=args.days) if args.days > 0 else None

    all_articles: list[Article] = []
    for source in sources:
        logger.info("Procesando fuente: %s", source.name)
        try:
            if source.type == "rss":
                items = parse_rss(session, source, args.keywords, any_of, cutoff, args.max_per_source)
            elif source.type == "html":
                items = parse_html_listing(session, source, args.keywords, any_of, cutoff, args.max_per_source)
            else:
                logger.warning("Tipo de fuente desconocido: %s", source.type)
                continue
        except Exception:
            # Una fuente rota (feed caído, HTML cambiado, etc.) no debe tumbar todo el proceso.
            logger.exception("Fallo procesando '%s'; se continúa con las demás fuentes.", source.name)
            continue

        logger.info("  -> %d artículo(s) relevante(s)", len(items))
        all_articles.extend(items)

    if not all_articles:
        logger.warning("No se encontraron noticias relevantes en esta ejecución.")
        return pd.DataFrame(columns=["fuente", "titulo", "resumen", "fecha", "enlace"])

    df = pd.DataFrame([a.__dict__ for a in all_articles])
    df.drop_duplicates(subset=["enlace"], inplace=True)
    df.sort_values(by=["fuente", "titulo"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def save(df: pd.DataFrame, output: str) -> Path:
    out_path = Path(output)
    if out_path.suffix.lower() in (".xlsx", ".xls"):
        df.to_excel(out_path, index=False)
    else:
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    return out_path.resolve()


def main(argv: list[str] | None = None) -> None:
    # En consolas de Windows (cp1252) los acentos/ñ de los titulares pueden
    # romper el print(); se fuerza UTF-8 en salida estándar para evitarlo.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    args = build_arg_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    df = run(args)
    out_path = save(df, args.output)
    if df.empty:
        logger.warning("No se encontraron noticias; se guardó igual %s solo con encabezados.", out_path)
        return

    logger.info("Se guardaron %d noticias únicas en %s", len(df), out_path)
    with pd.option_context("display.max_colwidth", 60):
        print(df[["fuente", "titulo", "fecha"]].to_string(index=False))


if __name__ == "__main__":
    main()
