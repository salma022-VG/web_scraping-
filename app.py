"""
====================================================================
 SERVIDOR WEB (BACKEND) DEL MONITOREO DE NOTICIAS
====================================================================

QUE ES ESTE ARCHIVO:

Es el "servidor" que hace posible usar el programa desde una pagina web,
en vez de tener que escribir comandos en una terminal. No repite la logica
de busqueda: reutiliza directamente las funciones que ya estan en
"monitoreo_noticias.py" (ese archivo sigue siendo el que realmente sabe
buscar noticias; este archivo solo lo conecta con el navegador).

COMO SE USA:

  1) Instala las librerias (una sola vez):
       python -m pip install -r requirements.txt

  2) Prende el servidor:
       python app.py

  3) Abre el navegador en:
       http://127.0.0.1:5000

  4) En la pagina, hay dos botones:
       - "Busqueda": corre el programa y muestra las noticias encontradas
         en pantalla (tarda entre 30 segundos y 2 minutos).
       - "Exportar a Excel": descarga el Excel con la ultima busqueda que
         se hizo en pantalla (no vuelve a buscar nada, solo descarga lo
         que ya se encontro).

  Para apagar el servidor, se vuelve a la ventana de la terminal donde
  quedo corriendo "python app.py" y se presiona Ctrl+C.
====================================================================
"""

from __future__ import annotations

import os
import threading
import uuid

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

# Aqui esta TODA la logica de busqueda ya construida (y probada) en el
# programa de consola. Este archivo solo la reutiliza, no la repite.
import monitoreo_noticias as motor

app = Flask(__name__)

# Hace que la pagina (templates/index.html) se vuelva a leer del disco en
# cada visita, en vez de quedar guardada en memoria la primera vez que se
# uso. Asi, si se edita el archivo HTML, el cambio se ve de inmediato sin
# tener que apagar y prender el servidor de nuevo.
app.config["TEMPLATES_AUTO_RELOAD"] = True


@app.context_processor
def _inyectar_version_de_archivos():
    """
    Los navegadores guardan copias de los archivos CSS/JS en su "cache"
    para no descargarlos de nuevo cada vez (asi las paginas cargan mas
    rapido). El problema es que, si nosotros cambiamos ese archivo, el
    navegador puede seguir usando la copia vieja y guardada sin darse
    cuenta de que hay una version nueva.

    Esta funcion le agrega a cada link de CSS/JS un numerito basado en la
    fecha en que se modifico el archivo por ultima vez (por ejemplo,
    "style.css?v=1738000000"). Cuando el archivo cambia, ese numero
    cambia, y el navegador entiende que es un archivo distinto y lo vuelve
    a descargar, en vez de usar la copia vieja guardada.
    """
    def version_de(nombre_archivo: str) -> int:
        ruta = os.path.join(app.static_folder, nombre_archivo)
        try:
            return int(os.path.getmtime(ruta))
        except OSError:
            return 0

    return {"version_de": version_de}


# Guarda en memoria el resultado de la ULTIMA busqueda hecha desde la
# pagina web. Sirve para que el boton "Exportar a Excel" pueda generar el
# archivo sin tener que volver a buscar todo de nuevo (eso demoraria otro
# minuto o dos innecesariamente).
ultima_busqueda: dict = {
    "bogota_items": [],
    "national_items": [],
    "label": "",
    "start": None,
    "end": None,
}

# ---------------------------------------------------------------------
# "Trabajos" en segundo plano.
#
# La busqueda puede tardar 1-2 minutos (o mas, si el servidor donde esta
# publicada la pagina tiene poca capacidad). Si se hiciera todo dentro de
# una sola peticion HTTP, algunos servicios de hosting (como Render en su
# plan gratuito) cortan la conexion antes de que termine, y el navegador
# nunca recibe respuesta.
#
# Para evitar eso, la busqueda se corre en un "hilo" aparte (en segundo
# plano): el boton "Busqueda" solo la manda a INICIAR y recibe de
# inmediato un numero de trabajo ("job_id"). Despues, la pagina pregunta
# cada pocos segundos "¿ya terminaste?" (endpoint /api/buscar/<job_id>)
# hasta que el trabajo queda listo. Cada una de esas preguntas es rapida,
# asi que ninguna conexion se queda esperando mucho tiempo.
# ---------------------------------------------------------------------
trabajos: dict[str, dict] = {}
trabajos_candado = threading.Lock()  # evita que dos hilos escriban "trabajos" al mismo tiempo


def _ejecutar_busqueda_en_segundo_plano(job_id: str, force_hours: float | None) -> None:
    """Esto es lo mismo que hacia antes api_buscar(), pero corriendo en un hilo aparte."""
    try:
        start, end, label = motor.get_time_window(force_hours=force_hours)

        queries = motor.load_bogota_queries(motor.KEYWORDS_FILE)
        motor.sync_default_queries_in_code(queries, os.path.abspath(motor.__file__))

        bogota_items = motor.collect_bogota_news(start, end, queries)
        national_items = motor.collect_national_news(start, end, top_n=3)
        motor.enrich_items(bogota_items + national_items)

        ultima_busqueda.update(
            bogota_items=bogota_items,
            national_items=national_items,
            label=label,
            start=start,
            end=end,
        )

        resultado = {
            "ventana": label,
            "desde": start.strftime("%Y-%m-%d %H:%M"),
            "hasta": end.strftime("%Y-%m-%d %H:%M"),
            "bogota": _serializar(bogota_items),
            "nacional": _serializar(national_items),
        }
        with trabajos_candado:
            trabajos[job_id] = {"estado": "listo", "resultado": resultado}

    except Exception as exc:  # noqa: BLE001
        with trabajos_candado:
            trabajos[job_id] = {"estado": "error", "error": str(exc)}


def _serializar(items: list) -> list[dict]:
    """
    Convierte la lista de noticias (las "fichas" NewsItem del motor) en
    algo que se pueda mandar como JSON al navegador: una lista de
    diccionarios simples con texto y numeros.
    """
    return [
        {
            "categoria": it.categoria,
            "fuente": it.medio,
            "titulo": it.titular,
            "resumen": it.resumen,
            "fecha": it.fecha.strftime("%Y-%m-%d %H:%M"),
            "enlace": it.enlace or it.link,
        }
        for it in items
    ]


@app.route("/")
def index():
    """Muestra la pagina principal (el archivo templates/index.html)."""
    return render_template("index.html")


@app.route("/api/buscar", methods=["POST"])
def api_buscar():
    """
    Esto es lo que se ejecuta cuando alguien aprieta el boton "Busqueda"
    en la pagina. NO hace la busqueda directamente (eso tardaria demasiado
    para una sola peticion web) -- la manda a correr en segundo plano y
    devuelve de inmediato un "job_id" para que la pagina pueda preguntar
    despues por el resultado (ver /api/buscar/<job_id> mas abajo).
    """
    datos = request.get_json(silent=True) or {}
    # El navegador puede pedir una ventana especifica (por ejemplo, "las
    # ultimas 6 horas"), o no pedir nada y dejar que el programa decida
    # solo segun la hora actual (igual que al correrlo desde la terminal).
    horas_forzadas = datos.get("window_hours")
    force_hours = float(horas_forzadas) if horas_forzadas else None

    job_id = uuid.uuid4().hex
    with trabajos_candado:
        # Se limpian trabajos anteriores: esta pagina la usa una sola
        # persona a la vez, no hace falta acumular busquedas viejas en memoria.
        trabajos.clear()
        trabajos[job_id] = {"estado": "en_progreso"}

    hilo = threading.Thread(
        target=_ejecutar_busqueda_en_segundo_plano,
        args=(job_id, force_hours),
        daemon=True,
    )
    hilo.start()

    return jsonify({"job_id": job_id})


@app.route("/api/buscar/<job_id>", methods=["GET"])
def api_estado_busqueda(job_id):
    """
    La pagina llama a esto cada pocos segundos para preguntar "¿ya
    terminaste la busqueda con este job_id?". Devuelve estado
    "en_progreso", "listo" (con el resultado adentro) o "error".
    """
    with trabajos_candado:
        trabajo = trabajos.get(job_id)

    if trabajo is None:
        return jsonify({"error": "Esa busqueda ya no existe (puede que se haya iniciado otra nueva)."}), 404

    return jsonify(trabajo)


@app.route("/api/palabras-clave", methods=["GET"])
def api_listar_palabras_clave():
    """
    Devuelve las categorias y palabras de busqueda que hay ahora mismo en
    "Palabras_Clave_Bogota.xlsx", para que la pagina web las pueda mostrar
    en la seccion "Palabras clave de busqueda".
    """
    queries = motor.load_bogota_queries(motor.KEYWORDS_FILE)
    return jsonify(queries)


@app.route("/api/palabras-clave", methods=["POST"])
def api_agregar_palabra_clave():
    """
    Agrega una palabra de busqueda nueva (con su categoria) sin necesidad
    de abrir el Excel a mano. Es lo mismo que agregar una fila en
    "Palabras_Clave_Bogota.xlsx" y guardarlo, pero hecho desde la pagina
    web -- util cuando el programa esta publicado en internet y no se
    tiene acceso directo a los archivos del servidor.
    """
    datos = request.get_json(silent=True) or {}
    categoria = str(datos.get("categoria", "")).strip()
    palabra = str(datos.get("palabra_clave", "")).strip()

    if not categoria or not palabra:
        return jsonify({"error": "Debes escribir una categoria y una palabra de busqueda."}), 400

    motor.ensure_keywords_file(motor.KEYWORDS_FILE)
    df = pd.read_excel(motor.KEYWORDS_FILE)

    ya_existe = (
        (df["categoria"].astype(str).str.strip().str.lower() == categoria.lower())
        & (df["palabra_clave"].astype(str).str.strip().str.lower() == palabra.lower())
    ).any()

    if not ya_existe:
        fila_nueva = pd.DataFrame([{"categoria": categoria, "palabra_clave": palabra}])
        df = pd.concat([df, fila_nueva], ignore_index=True)
        df.to_excel(motor.KEYWORDS_FILE, index=False)

    queries = motor.load_bogota_queries(motor.KEYWORDS_FILE)
    motor.sync_default_queries_in_code(queries, os.path.abspath(motor.__file__))
    return jsonify(queries)


@app.route("/api/palabras-clave", methods=["DELETE"])
def api_borrar_palabra_clave():
    """Quita una palabra de busqueda especifica (dentro de su categoria)."""
    datos = request.get_json(silent=True) or {}
    categoria = str(datos.get("categoria", "")).strip()
    palabra = str(datos.get("palabra_clave", "")).strip()

    if os.path.exists(motor.KEYWORDS_FILE):
        df = pd.read_excel(motor.KEYWORDS_FILE)
        se_debe_conservar = ~(
            (df["categoria"].astype(str).str.strip().str.lower() == categoria.lower())
            & (df["palabra_clave"].astype(str).str.strip().str.lower() == palabra.lower())
        )
        df = df[se_debe_conservar]
        df.to_excel(motor.KEYWORDS_FILE, index=False)

    queries = motor.load_bogota_queries(motor.KEYWORDS_FILE)
    motor.sync_default_queries_in_code(queries, os.path.abspath(motor.__file__))
    return jsonify(queries)


@app.route("/api/exportar", methods=["GET"])
def api_exportar():
    """
    Esto es lo que se ejecuta cuando alguien aprieta el boton "Exportar a
    Excel". Usa el resultado de la ULTIMA busqueda (guardado en
    "ultima_busqueda") para generar el archivo .xlsx y se lo manda al
    navegador para que lo descargue.
    """
    if not ultima_busqueda["bogota_items"] and not ultima_busqueda["national_items"]:
        return jsonify({"error": "Primero haz una busqueda antes de exportar."}), 400

    filename = motor.export_to_excel(
        ultima_busqueda["bogota_items"],
        ultima_busqueda["national_items"],
        ultima_busqueda["label"],
        ultima_busqueda["start"],
        ultima_busqueda["end"],
    )
    return send_file(filename, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    # Cuando esto corre en tu computador, no existe la variable de entorno
    # "PORT" y usa el puerto 5000 de siempre (http://127.0.0.1:5000).
    # Cuando corre publicado en un servicio como Render, ellos le dicen al
    # programa por cual puerto tiene que escuchar a traves de esa variable,
    # y por eso hay que respetarla en vez de dejar el 5000 fijo.
    puerto = int(os.environ.get("PORT", 5000))
    # threaded=True permite que el servidor pueda responder a otras
    # peticiones (por ejemplo, recargar la pagina) mientras una busqueda
    # todavia esta en proceso.
    app.run(host="0.0.0.0", port=puerto, debug=False, threaded=True)
