#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Versión Web en Python (Streamlit) — Buscador de PDFs Digitales
--------------------------------------------------------------
Características:
- Funciona en el navegador (local en su red o desplegada online en Streamlit Cloud / Render).
- Acepta múltiples archivos PDF o un archivo .ZIP con carpetas completas de PDFs.
- Indexación en sesión: permite buscar múltiples términos sin volver a subir los archivos.
- Búsqueda sin distinguir mayúsculas, minúsculas ni tildes.
- Muestra el archivo exacto, página exacta donde está el artículo/cliente/NIT/factura,
  vista previa visual de la página con resaltado y descarga en Excel (.xlsx).

Instalación:
    python -m pip install streamlit pymupdf openpyxl Pillow

Ejecución local en el navegador:
    streamlit run buscador_pdf_web.py
"""

import io
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import streamlit as st

try:
    import pymupdf as fitz
except ImportError:
    import fitz


def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    descompuesto = unicodedata.normalize("NFD", texto)
    sin_tildes = "".join(c for c in descompuesto if unicodedata.category(c) != "Mn")
    return sin_tildes.lower()


@dataclass
class PaginaWeb:
    file_id: str
    nombre_archivo: str
    ruta_relativa: str
    numero_pagina: int
    total_paginas: int
    texto_original: str
    texto_normalizado: str


def construir_extracto(texto_orig: str, texto_norm: str, termino_norm: str) -> str:
    pos = texto_norm.find(termino_norm)
    if pos == -1:
        linea = " ".join(texto_orig.split())
        return linea[:150] + ("..." if len(linea) > 150 else "")
    inicio = max(0, pos - 50)
    fin = min(len(texto_orig), pos + len(termino_norm) + 90)
    fragmento = " ".join(texto_orig[inicio:fin].split())
    prefijo = "..." if inicio > 0 else ""
    sufijo = "..." if fin < len(texto_orig) else ""
    return f"{prefijo}{fragmento}{sufijo}"


def indexar_entradas(entradas: List[Tuple[str, str, bytes]]) -> Tuple[List[PaginaWeb], Dict[str, bytes], List[str]]:
    paginas: List[PaginaWeb] = []
    documentos_bytes: Dict[str, bytes] = {}
    errores: List[str] = []

    progreso = st.progress(0, text="Leyendo PDFs digitales...")
    total = len(entradas)

    for idx, (ruta_rel, nombre, raw_bytes) in enumerate(entradas, start=1):
        file_id = f"{idx}_{ruta_rel}"
        try:
            with fitz.open(stream=raw_bytes, filetype="pdf") as doc:
                num_pags = len(doc)
                documentos_bytes[file_id] = raw_bytes
                for p_idx in range(num_pags):
                    texto = doc[p_idx].get_text("text") or ""
                    paginas.append(
                        PaginaWeb(
                            file_id=file_id,
                            nombre_archivo=nombre,
                            ruta_relativa=ruta_rel,
                            numero_pagina=p_idx + 1,
                            total_paginas=num_pags,
                            texto_original=texto,
                            texto_normalizado=normalizar_texto(texto),
                        )
                    )
        except Exception as exc:
            errores.append(f"{nombre}: {exc}")

        progreso.progress(idx / max(total, 1), text=f"Procesando {idx} de {total} PDFs...")

    progreso.empty()
    return paginas, documentos_bytes, errores


def generar_excel_bytes(resultados: List[dict], termino: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados PDF"

    encabezados = [
        "Archivo PDF",
        "Ruta",
        "Página Exacta",
        "Total Páginas",
        "Coincidencias",
        "Término Buscado",
        "Contexto / Artículo Encontrado",
    ]
    ws.append(encabezados)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="left", vertical="center")

    for r in resultados:
        ws.append([
            r["Archivo PDF"],
            r["Ruta"],
            r["Página Exacta"],
            r["Total Páginas"],
            r["Coincidencias"],
            termino,
            r["Contexto / Artículo"],
        ])

    for idx, ancho in enumerate([32, 36, 14, 14, 14, 24, 62], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = ancho

    ws.freeze_panes = "A2"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def renderizar_pagina_png(raw_bytes: bytes, num_pagina: int, termino: str, zoom: float) -> bytes:
    with fitz.open(stream=raw_bytes, filetype="pdf") as doc:
        page = doc[num_pagina - 1]
        terminos = [t for t in normalizar_texto(termino).split() if t]
        termino_full = normalizar_texto(termino)

        for palabra_info in page.get_text("words"):
            p_norm = normalizar_texto(palabra_info[4])
            if any(t in p_norm for t in terminos) or (len(termino_full) >= 2 and termino_full in p_norm):
                annot = page.add_highlight_annot(fitz.Rect(palabra_info[:4]))
                annot.set_colors(stroke=(1, 0.92, 0.23))
                annot.update()

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return pix.tobytes("png")


def main() -> None:
    st.set_page_config(page_title="Buscador Web de PDFs", layout="wide")
    st.title("Buscador Web de PDFs Digitales")
    st.caption(
        "Suba múltiples PDFs o un archivo .ZIP con su carpeta completa. "
        "Busca sin distinguir mayúsculas ni tildes e indica la página exacta del artículo, cliente, NIT o factura."
    )

    if "paginas" not in st.session_state:
        st.session_state.paginas = []
        st.session_state.docs = {}
        st.session_state.errores = []

    with st.sidebar:
        st.header("1. Cargar documentos")
        archivos = st.file_uploader(
            "Seleccione archivos PDF o un .ZIP de su carpeta",
            type=["pdf", "zip"],
            accept_multiple_files=True,
        )
        if st.button("Indexar archivos cargados", type="primary", use_container_width=True):
            if not archivos:
                st.warning("Seleccione al menos un PDF o ZIP.")
            else:
                entradas: List[Tuple[str, str, bytes]] = []
                for f in archivos:
                    data = f.read()
                    if f.name.lower().endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(data)) as zf:
                            for info in zf.infolist():
                                if not info.is_dir() and info.filename.lower().endswith(".pdf"):
                                    nombre = info.filename.split("/")[-1]
                                    entradas.append((info.filename, nombre, zf.read(info.filename)))
                    elif f.name.lower().endswith(".pdf"):
                        entradas.append((f.name, f.name, data))

                if not entradas:
                    st.error("No se encontraron archivos PDF válidos.")
                else:
                    paginas, docs, errores = indexar_entradas(entradas)
                    st.session_state.paginas = paginas
                    st.session_state.docs = docs
                    st.session_state.errores = errores
                    st.success(f"Listo: {len(docs)} PDFs ({len(paginas)} páginas indexadas).")

        if st.session_state.paginas:
            st.info(
                f"En memoria: {len(st.session_state.docs)} PDFs · "
                f"{len(st.session_state.paginas)} páginas listas para buscar."
            )

    if not st.session_state.paginas:
        st.info("Cargue e indexe sus archivos PDF en el panel izquierdo para comenzar a buscar.")
        return

    col_q, col_m = st.columns([3, 1])
    with col_q:
        termino = st.text_input(
            "Buscar cliente, proveedor, NIT, cédula, factura o artículo",
            placeholder="Ej: 800.245.119, Martínez, Soldador Inversor...",
        )
    with col_m:
        frase_exacta = st.toggle("Frase exacta", value=True)

    if not termino.strip():
        return

    termino_norm = normalizar_texto(termino.strip())
    palabras = [p for p in termino_norm.split() if p]

    resultados: List[dict] = []
    for pag in st.session_state.paginas:
        if frase_exacta:
            conteo = pag.texto_normalizado.count(termino_norm)
            coincide = conteo > 0
            primer = termino_norm
        else:
            coincide = bool(palabras) and all(p in pag.texto_normalizado for p in palabras)
            conteo = sum(pag.texto_normalizado.count(p) for p in palabras) if coincide else 0
            primer = palabras[0] if palabras else ""

        if coincide:
            resultados.append(
                {
                    "file_id": pag.file_id,
                    "Archivo PDF": pag.nombre_archivo,
                    "Ruta": pag.ruta_relativa,
                    "Página Exacta": pag.numero_pagina,
                    "Total Páginas": pag.total_paginas,
                    "Coincidencias": conteo,
                    "Contexto / Artículo": construir_extracto(pag.texto_original, pag.texto_normalizado, primer),
                }
            )

    st.subheader(f"Resultados ({len(resultados)} páginas encontradas)")
    if not resultados:
        st.warning(f"No se encontraron coincidencias para «{termino}».")
        return

    excel_bytes = generar_excel_bytes(resultados, termino.strip())
    st.download_button(
        label="Exportar resultados a Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"resultados_pdf_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    col_tabla, col_preview = st.columns([1.05, 1.15], gap="large")
    with col_tabla:
        tabla_visible = [
            {k: v for k, v in r.items() if k != "file_id"} for r in resultados
        ]
        st.dataframe(tabla_visible, use_container_width=True, hide_index=True)

    with col_preview:
        opciones = [
            f"{i + 1}. {r['Archivo PDF']} — Página {r['Página Exacta']} de {r['Total Páginas']}"
            for i, r in enumerate(resultados)
        ]
        idx_sel = st.selectbox("Vista previa de página encontrada", range(len(opciones)), format_func=lambda i: opciones[i])
        sel = resultados[idx_sel]
        zoom = st.slider("Zoom de vista previa", 1.0, 2.5, 1.5, 0.25)
        img_bytes = renderizar_pagina_png(
            st.session_state.docs[sel["file_id"]],
            sel["Página Exacta"],
            termino.strip(),
            zoom,
        )
        st.image(img_bytes, caption=f"{sel['Archivo PDF']} · Página {sel['Página Exacta']} de {sel['Total Páginas']}", use_container_width=True)


if __name__ == "__main__":
    main()
