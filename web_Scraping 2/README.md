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

## Uso

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

## Ajustar temas o busquedas

Las categorias y palabras clave de Bogota estan en el diccionario
`BOGOTA_QUERIES` al inicio de `monitoreo_noticias.py`. Puedes agregar o quitar
terminos de busqueda ahi para afinar los resultados.
