// Este archivo conecta los botones de la pagina con el servidor (app.py).
// No hace ninguna busqueda por si mismo: solo le pide al servidor que la
// haga (con fetch, que es la forma en que una pagina web le "habla" a un
// servidor) y muestra lo que el servidor responde.

const btnBuscar = document.getElementById("btnBuscar");
const btnExportar = document.getElementById("btnExportar");
const selectVentana = document.getElementById("ventana");

const estado = document.getElementById("estado");
const estadoTexto = document.getElementById("estadoTexto");
const cajaError = document.getElementById("error");

const seccionResumen = document.getElementById("resumen");
const seccionNacional = document.getElementById("seccionNacional");
const seccionBogota = document.getElementById("seccionBogota");

const btnTogglePalabras = document.getElementById("btnTogglePalabras");
const panelPalabras = document.getElementById("panelPalabras");
const listaCategorias = document.getElementById("listaCategorias");
const listaCategoriasExistentes = document.getElementById("listaCategoriasExistentes");
const formAgregarPalabra = document.getElementById("formAgregarPalabra");
const inputCategoria = document.getElementById("inputCategoria");
const inputPalabra = document.getElementById("inputPalabra");
const palabrasError = document.getElementById("palabrasError");

function mostrar(elemento) {
  elemento.classList.remove("oculto");
}

function ocultar(elemento) {
  elemento.classList.add("oculto");
}

function escaparTexto(texto) {
  const div = document.createElement("div");
  div.textContent = texto ?? "";
  return div.innerHTML;
}

// Le asigna un color de "badge" a cada categoria, siempre el mismo color
// para la misma categoria (se calcula a partir del texto, asi que funciona
// automaticamente aunque se agreguen categorias nuevas desde el Excel de
// palabras clave, sin tener que tocar este archivo).
const COLORES_BADGE = ["badge-1", "badge-2", "badge-3", "badge-4", "badge-5", "badge-6", "badge-7", "badge-8"];

function colorParaCategoria(categoria) {
  let hash = 0;
  for (let i = 0; i < categoria.length; i++) {
    hash = (hash * 31 + categoria.charCodeAt(i)) >>> 0;
  }
  return COLORES_BADGE[hash % COLORES_BADGE.length];
}

function llenarTabla(tablaId, noticias) {
  const tabla = document.getElementById(tablaId);
  const cuerpo = tabla.querySelector("tbody");
  cuerpo.innerHTML = "";

  if (!noticias || noticias.length === 0) {
    const fila = document.createElement("tr");
    fila.innerHTML = `<td colspan="5" class="vacio">No se encontraron noticias en esta franja.</td>`;
    cuerpo.appendChild(fila);
    return;
  }

  for (const noticia of noticias) {
    const fila = document.createElement("tr");
    const colorBadge = colorParaCategoria(noticia.categoria || "");
    // El atributo "data-label" no se ve en pantallas grandes (ahi se ve la
    // tabla normal, con encabezados arriba). En celular, la tabla cambia a
    // formato de "tarjetas" (ver static/style.css) y ahi si se usa este
    // atributo para mostrar, al lado de cada dato, a que columna pertenece.
    fila.innerHTML = `
      <td class="col-categoria" data-label="Categoria"><span class="badge ${colorBadge}">${escaparTexto(noticia.categoria)}</span></td>
      <td class="col-fuente" data-label="Fuente">${escaparTexto(noticia.fuente)}</td>
      <td class="col-titulo" data-label="Titulo"><a href="${escaparTexto(noticia.enlace)}" target="_blank" rel="noopener">${escaparTexto(noticia.titulo)}</a></td>
      <td class="col-resumen" data-label="Resumen">${escaparTexto(noticia.resumen)}</td>
      <td class="col-fecha" data-label="Fecha">${escaparTexto(noticia.fecha)}</td>
    `;
    cuerpo.appendChild(fila);
  }
}

async function buscarNoticias() {
  ocultar(cajaError);
  ocultar(seccionResumen);
  ocultar(seccionNacional);
  ocultar(seccionBogota);
  mostrar(estado);

  btnBuscar.disabled = true;
  btnExportar.disabled = true;

  const horas = selectVentana.value; // "" si es "Automatica"

  try {
    const respuesta = await fetch("/api/buscar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ window_hours: horas || null }),
    });

    // Si la busqueda tardo demasiado, el servidor puede cortar la
    // conexion a mitad de camino y la respuesta llega vacia (no es un
    // JSON valido). Eso no significa que el navegador este roto: hay que
    // avisarle al usuario que lo intente de nuevo, no mostrar el error
    // tecnico crudo.
    let datos;
    try {
      datos = await respuesta.json();
    } catch {
      throw new Error(
        "La busqueda tardo demasiado y el servidor corto la conexion. " +
        "Intenta de nuevo, o elige una franja mas corta (por ejemplo, 6 horas)."
      );
    }

    if (!respuesta.ok) {
      throw new Error(datos.error || "Ocurrio un error al buscar las noticias.");
    }

    document.getElementById("resVentana").textContent = datos.ventana;
    document.getElementById("resRango").textContent = `${datos.desde} -> ${datos.hasta}`;
    document.getElementById("resTotalBogota").textContent = datos.bogota.length;
    document.getElementById("resTotalNacional").textContent = datos.nacional.length;

    llenarTabla("tablaNacional", datos.nacional);
    llenarTabla("tablaBogota", datos.bogota);

    mostrar(seccionResumen);
    mostrar(seccionNacional);
    mostrar(seccionBogota);
    btnExportar.disabled = false;
  } catch (err) {
    cajaError.textContent = err.message;
    mostrar(cajaError);
  } finally {
    ocultar(estado);
    btnBuscar.disabled = false;
  }
}

async function exportarExcel() {
  btnExportar.disabled = true;
  ocultar(cajaError);

  try {
    const respuesta = await fetch("/api/exportar");

    if (!respuesta.ok) {
      const datos = await respuesta.json().catch(() => ({}));
      throw new Error(datos.error || "No se pudo generar el Excel.");
    }

    // El servidor manda el archivo; aqui se convierte la respuesta en una
    // descarga real dentro del navegador.
    const blob = await respuesta.blob();
    const nombreArchivo = obtenerNombreArchivo(respuesta) || "Monitoreo_Noticias.xlsx";

    const enlaceTemporal = document.createElement("a");
    enlaceTemporal.href = URL.createObjectURL(blob);
    enlaceTemporal.download = nombreArchivo;
    document.body.appendChild(enlaceTemporal);
    enlaceTemporal.click();
    enlaceTemporal.remove();
  } catch (err) {
    cajaError.textContent = err.message;
    mostrar(cajaError);
  } finally {
    btnExportar.disabled = false;
  }
}

function obtenerNombreArchivo(respuesta) {
  const cabecera = respuesta.headers.get("Content-Disposition") || "";
  const match = cabecera.match(/filename="?([^"]+)"?/);
  return match ? match[1] : null;
}

// ---------------------------------------------------------------------
// Palabras clave de busqueda: ver, agregar y quitar desde la pagina web
// (sin tener que abrir el Excel a mano, util cuando la pagina esta
// publicada en internet).
// ---------------------------------------------------------------------

function dibujarCategorias(categorias) {
  listaCategorias.innerHTML = "";
  listaCategoriasExistentes.innerHTML = "";

  const nombresCategorias = Object.keys(categorias).sort();

  for (const nombre of nombresCategorias) {
    const opcion = document.createElement("option");
    opcion.value = nombre;
    listaCategoriasExistentes.appendChild(opcion);

    const tarjeta = document.createElement("div");
    tarjeta.className = "categoria-card";

    const titulo = document.createElement("h3");
    titulo.textContent = nombre;
    tarjeta.appendChild(titulo);

    const chips = document.createElement("div");
    chips.className = "chips";

    for (const palabra of categorias[nombre]) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.innerHTML = `<span>${escaparTexto(palabra)}</span>`;

      const botonQuitar = document.createElement("button");
      botonQuitar.type = "button";
      botonQuitar.textContent = "×";
      botonQuitar.title = `Quitar "${palabra}"`;
      botonQuitar.addEventListener("click", () => quitarPalabraClave(nombre, palabra));

      chip.appendChild(botonQuitar);
      chips.appendChild(chip);
    }

    tarjeta.appendChild(chips);
    listaCategorias.appendChild(tarjeta);
  }
}

async function cargarPalabrasClave() {
  ocultar(palabrasError);
  try {
    const respuesta = await fetch("/api/palabras-clave");
    const datos = await respuesta.json();
    if (!respuesta.ok) {
      throw new Error(datos.error || "No se pudieron cargar las palabras clave.");
    }
    dibujarCategorias(datos);
  } catch (err) {
    palabrasError.textContent = err.message;
    mostrar(palabrasError);
  }
}

async function agregarPalabraClave(evento) {
  evento.preventDefault();
  ocultar(palabrasError);

  const categoria = inputCategoria.value.trim();
  const palabra = inputPalabra.value.trim();
  if (!categoria || !palabra) return;

  try {
    const respuesta = await fetch("/api/palabras-clave", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ categoria, palabra_clave: palabra }),
    });
    const datos = await respuesta.json();
    if (!respuesta.ok) {
      throw new Error(datos.error || "No se pudo agregar la palabra.");
    }
    dibujarCategorias(datos);
    inputPalabra.value = "";
    inputPalabra.focus();
  } catch (err) {
    palabrasError.textContent = err.message;
    mostrar(palabrasError);
  }
}

async function quitarPalabraClave(categoria, palabra) {
  ocultar(palabrasError);
  try {
    const respuesta = await fetch("/api/palabras-clave", {
      method: "DELETE",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ categoria, palabra_clave: palabra }),
    });
    const datos = await respuesta.json();
    if (!respuesta.ok) {
      throw new Error(datos.error || "No se pudo quitar la palabra.");
    }
    dibujarCategorias(datos);
  } catch (err) {
    palabrasError.textContent = err.message;
    mostrar(palabrasError);
  }
}

function alternarPanelPalabras() {
  const estaOculto = panelPalabras.classList.contains("oculto");
  if (estaOculto) {
    mostrar(panelPalabras);
    cargarPalabrasClave();
  } else {
    ocultar(panelPalabras);
  }
}

btnBuscar.addEventListener("click", buscarNoticias);
btnExportar.addEventListener("click", exportarExcel);
btnTogglePalabras.addEventListener("click", alternarPanelPalabras);
formAgregarPalabra.addEventListener("submit", agregarPalabraClave);
