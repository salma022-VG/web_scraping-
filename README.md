# Monitoreo de Noticias - Bogota y Colombia

Script que busca noticias de **hoy** sobre Bogota (Hacienda, infraestructura,
vivienda, movilidad, inmobiliario, alcalde Carlos Fernando Galan, Catastro
Bogota / certificados catastrales) mas las **3 noticias nacionales de
Colombia** mas importantes del dia, y exporta todo a un Excel con titular +
medio + link.

## Como funciona

Usa los feeds RSS de Google News, que agregan cientos de medios colombianos
(El Tiempo, El Espectador, Semana, RCN, Caracol, Infobae, La Republica,
Portafolio, Bluradio, El Colombiano, etc.) con fecha de publicacion exacta,
lo que permite filtrar de forma confiable por franja horaria. Esto es mucho
mas robusto que scrapear sitio por sitio (cada medio tiene HTML distinto y
bloqueos anti-bot que rompen ese enfoque constantemente).

## Franja horaria automatica (hora de Bogota)

- **8:00 - 9:00 am** -> noticias de las **ultimas 12 horas**, hasta las 8:00 am.
- **2:00 - 3:30 pm** -> noticias de las **ultimas 6 horas**, hasta las 3:00 pm.
- Fuera de esas franjas, usa la franja mas reciente ya pasada (util si lo
  corres manualmente fuera de horario).

## Instalacion

```
pip install -r requirements.txt
```

## Opcion 1: Usarlo desde una pagina web (recomendado para uso diario)

El proyecto incluye una pagina web sencilla para no tener que usar la
terminal. Para prenderla:

```
python app.py
```

Y luego abre en el navegador (Chrome, Edge, el que uses):

```
http://127.0.0.1:5000
```

En la pagina hay tres botones:

- **Busqueda**: corre el programa y muestra las noticias encontradas en
  pantalla, en tablas (tarda entre 30 segundos y 2 minutos, se ve un
  mensaje de "Buscando..." mientras tanto). Arriba puedes elegir si
  quieres la franja automatica (segun la hora del reloj) o forzar
  "ultimas 6 / 12 / 24 horas".
- **Exportar a Excel**: descarga el Excel con exactamente los resultados
  que se ven en pantalla en ese momento (no vuelve a buscar nada).
- **Palabras clave**: abre un panel para ver, agregar o quitar las
  categorias/palabras de busqueda de Bogota directamente desde la pagina
  (sin tener que abrir el Excel a mano). Los cambios aplican desde la
  siguiente busqueda.

Para apagar el servidor: vuelve a la ventana de la terminal donde quedo
corriendo `python app.py` y presiona `Ctrl+C`.

## Opcion 2: Usarlo desde la terminal (sin pagina web)

Ejecucion normal (usa la hora actual del sistema):

```
python monitoreo_noticias.py
```

Forzar una hora especifica (para pruebas), ejemplo simulando la franja de la tarde:

```
python monitoreo_noticias.py --now "2026-07-28 15:00"
```

Forzar directamente el tamano de la ventana en horas (ignora la franja):

```
python monitoreo_noticias.py --window 6
```

## Salida

Se genera un archivo `Monitoreo_Noticias_<fecha>_<hora>.xlsx` en la misma
carpeta, con 3 hojas:

- **Resumen**: franja aplicada, rango de horas, totales.
- **Bogota**: todas las noticias encontradas, con columnas `categoria`,
  `fuente`, `titulo`, `resumen`, `fecha`, `enlace` (el enlace ya resuelto a
  la URL real del medio, no el link de redireccion de Google News).
- **Nacional (Top 3)**: las 3 noticias nacionales mas relevantes del dia,
  con las mismas columnas.

El `resumen` se extrae de la meta-descripcion de la pagina del articulo; si
un medio no la expone, se usa el titular como respaldo.

## Automatizar con el Programador de tareas de Windows

Para que corra solo a las 8:00 am y 3:00 pm todos los dias:

1. Abre "Programador de tareas" (Task Scheduler).
2. Crear tarea basica -> nombre "Monitoreo Noticias 8am".
3. Desencadenador: Diariamente, 8:00 am.
4. Accion: Iniciar un programa.
   - Programa/script: ruta a tu `python.exe`
     (por ejemplo `C:\Users\valen\AppData\Local\Python\pythoncore-3.14-64\python.exe`)
   - Argumentos: `monitoreo_noticias.py`
   - Iniciar en: `C:\Users\valen\Downloads\web_Scraping 2`
5. Repite el proceso creando otra tarea "Monitoreo Noticias 3pm" a las 3:00 pm.

El script detecta la franja automaticamente segun la hora a la que se ejecute,
asi que no hace falta pasarle argumentos cuando corre programado.

## Ajustar temas o busquedas (sin tocar codigo)

Las categorias y palabras clave de Bogota viven en el archivo
`Palabras_Clave_Bogota.xlsx` (se crea solo la primera vez que corres el
programa). Tiene dos columnas: `categoria` y `palabra_clave`.

- Para agregar una palabra a una categoria que ya existe: abre el Excel,
  agrega una fila nueva con esa categoria y la palabra, guarda.
- Para crear una categoria nueva: escribe el nombre nuevo en la columna
  `categoria` de una fila nueva.

La siguiente vez que corras `python monitoreo_noticias.py`, va a leer ese
Excel automaticamente y usar la lista actualizada — no hace falta editar
`monitoreo_noticias.py` para nada. Ademas, cada corrida deja una copia de
respaldo de esa lista dentro del propio codigo (`DEFAULT_BOGOTA_QUERIES`),
por si el Excel se llegara a borrar por accidente; ese respaldo solo se usa
para volver a crear el Excel si no existe, el programa nunca lo usa
mientras el Excel este presente.

Si estas usando la pagina web (`python app.py`), tambien puedes agregar o
quitar palabras desde el boton **"Palabras clave"**, sin tocar el Excel
directamente.

## Publicarlo en internet (Render)

Para que la pagina sea accesible desde cualquier lugar (no solo desde tu
computador), se puede publicar en [Render](https://render.com) usando su
plan gratuito. Los pasos:

1. **Sube este proyecto a un repositorio de GitHub** (crea un repo nuevo en
   github.com, y sigue las instrucciones para subir una carpeta existente
   con `git remote add origin <url-de-tu-repo>` y `git push -u origin main`).
2. Crea una cuenta gratis en [render.com](https://render.com) (puedes
   entrar directamente con tu cuenta de GitHub).
3. En el panel de Render: **New +** -> **Web Service**, y elige el
   repositorio que acabas de subir.
4. Render detecta automaticamente que es un proyecto de Python. Configuralo asi:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app` (ya viene definido en el archivo `Procfile`, Render lo deberia tomar solo)
   - **Instance Type**: Free
5. Dale clic a **Create Web Service**. La primera vez tarda unos minutos en
   instalar todo. Cuando termine, Render te da una URL publica
   (algo como `https://tu-proyecto.onrender.com`) — esa es tu pagina web,
   accesible desde cualquier celular o computador.
6. Cada vez que hagas cambios y los subas a GitHub (`git push`), Render
   vuelve a publicar la version nueva automaticamente.

**Importante sobre el plan gratuito de Render:**

- El servidor se "duerme" despues de ~15 minutos sin uso, y la primera
  visita despues de eso tarda unos 30-60 segundos extra en despertar (las
  siguientes ya van rapido).
- El disco donde vive `Palabras_Clave_Bogota.xlsx` es temporal: los
  cambios que hagas desde el boton "Palabras clave" funcionan mientras el
  servidor sigue prendido, pero se pueden perder si el servidor se
  reinicia o se vuelve a publicar. Para que un cambio de palabras clave
  quede permanente, edita el Excel en tu copia local del proyecto y vuelve
  a subirlo a GitHub (`git add`, `git commit`, `git push`).
