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

import base64
import io
import os
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

import streamlit as st

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    TIENE_OPENPYXL = True
except ModuleNotFoundError:
    TIENE_OPENPYXL = False

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ModuleNotFoundError:
        archivos_repo = []
        try:
            archivos_repo = sorted(os.listdir(os.path.dirname(os.path.abspath(__file__)) or "."))
        except Exception:
            pass
        st.error(
            "Falta instalar la librería `pymupdf` en Streamlit Cloud.\n\n"
            f"Archivos detectados en la carpeta del repositorio: `{', '.join(archivos_repo) or 'ninguno'}`\n\n"
            "Asegúrese de tener un archivo llamado exactamente **`requirements.txt`** (en minúsculas y sin doble `.txt.txt`) "
            "junto a `buscador_pdf_web.py` con este contenido:\n\n"
            "```text\nstreamlit\npymupdf\nopenpyxl\nPillow\n```\n\n"
            "Luego en Streamlit Cloud abra **Manage app -> Delete app** y vuelva a desplegarlo (o haga un nuevo commit cambiando `requirements.txt`)."
        )
        st.stop()


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


def _generar_xlsx_puro_python(filas: List[List[object]]) -> bytes:
    """Genera un archivo .xlsx válido usando solo la librería estándar (sin openpyxl)."""
    def col_letra(idx_0: int) -> str:
        res = ""
        n = idx_0 + 1
        while n > 0:
            n, rem = divmod(n - 1, 26)
            res = chr(65 + rem) + res
        return res

    sheet_rows = []
    for r_idx, fila in enumerate(filas, start=1):
        celdas = []
        for c_idx, val in enumerate(fila):
            ref = f"{col_letra(c_idx)}{r_idx}"
            if isinstance(val, (int, float)):
                celdas.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                txt = xml_escape(str(val))
                celdas.append(f'<c r="{ref}" t="inlineStr"><is><t>{txt}</t></is></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(celdas)}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        '</worksheet>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Resultados PDF" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '</Relationships>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return buf.getvalue()


def detectar_casilla_pagina(
    raw_bytes: bytes, num_pagina: int, termino: str, frase_exacta: bool
) -> Tuple[str, Optional[Tuple[float, float, float, float]]]:
    """
    Detecta la casilla exacta (tabla/fila/columna o celda visual) donde aparece el término en la página.
    Devuelve (descripcion_casilla, rect_tuple_primera_casilla).
    """
    termino_norm = normalizar_texto(termino.strip())
    palabras_busqueda = [p for p in termino_norm.split() if p]
    if not termino_norm:
        return "—", None

    def coincide_texto(txt_norm: str) -> bool:
        if frase_exacta:
            return termino_norm in txt_norm
        return bool(palabras_busqueda) and all(p in txt_norm for p in palabras_busqueda)

    with fitz.open(stream=raw_bytes, filetype="pdf") as doc:
        page = doc[num_pagina - 1]

        # 1. Intentar detectar tablas vectoriales reales en la página
        try:
            tablas = page.find_tables()
            if tablas and getattr(tablas, "tables", None):
                casilla_global = 0
                encontradas: List[str] = []
                primer_rect: Optional[Tuple[float, float, float, float]] = None

                for t_idx, tabla in enumerate(tablas.tables, start=1):
                    for r_idx, fila_celdas in enumerate(tabla.rows, start=1):
                        for c_idx, cell_bbox in enumerate(fila_celdas.cells, start=1):
                            casilla_global += 1
                            if cell_bbox is None:
                                continue
                            rect = fitz.Rect(cell_bbox)
                            txt_celda = normalizar_texto(page.get_textbox(rect) or "")
                            if coincide_texto(txt_celda):
                                encontradas.append(
                                    f"Casilla #{casilla_global} (Tabla {t_idx} · Fila {r_idx}, Col. {c_idx})"
                                )
                                if primer_rect is None:
                                    primer_rect = (rect.x0, rect.y0, rect.x1, rect.y1)

                if encontradas:
                    resumen = encontradas[0]
                    if len(encontradas) > 1:
                        resumen += f" (+{len(encontradas) - 1} más)"
                    return resumen, primer_rect
        except Exception:
            pass

        # 2. Respaldo universal por filas y casillas visuales (para PDFs sin bordes de tabla)
        words = page.get_text("words")
        if not words:
            return "Sin casillas detectadas", None

        words_sorted = sorted(words, key=lambda w: ((w[1] + w[3]) / 2.0, w[0]))
        filas_visuales: List[List[tuple]] = []
        y_actual: Optional[float] = None

        for w in words_sorted:
            y_mid = (w[1] + w[3]) / 2.0
            if y_actual is None or abs(y_mid - y_actual) > 5.0:
                filas_visuales.append([w])
                y_actual = y_mid
            else:
                filas_visuales[-1].append(w)

        casilla_num = 0
        encontradas_vis: List[str] = []
        primer_rect_vis: Optional[Tuple[float, float, float, float]] = None

        for f_idx, fila_words in enumerate(filas_visuales, start=1):
            fila_words.sort(key=lambda w: w[0])
            celdas_fila: List[List[tuple]] = [[fila_words[0]]]
            for w in fila_words[1:]:
                prev = celdas_fila[-1][-1]
                if w[0] - prev[2] > 16.0:
                    celdas_fila.append([w])
                else:
                    celdas_fila[-1].append(w)

            fila_completa_norm = normalizar_texto(" ".join(w[4] for w in fila_words))
            coincide_en_fila = coincide_texto(fila_completa_norm)

            for c_idx, celda_words in enumerate(celdas_fila, start=1):
                casilla_num += 1
                celda_norm = normalizar_texto(" ".join(w[4] for w in celda_words))
                coincide_celda = coincide_texto(celda_norm) or (
                    coincide_en_fila
                    and any(p in celda_norm for p in palabras_busqueda)
                )
                if coincide_celda:
                    x0 = min(w[0] for w in celda_words)
                    y0 = min(w[1] for w in celda_words)
                    x1 = max(w[2] for w in celda_words)
                    y1 = max(w[3] for w in celda_words)
                    encontradas_vis.append(f"Casilla #{casilla_num} (Fila {f_idx}, Col. {c_idx})")
                    if primer_rect_vis is None:
                        primer_rect_vis = (x0, y0, x1, y1)

        if encontradas_vis:
            resumen = encontradas_vis[0]
            if len(encontradas_vis) > 1:
                resumen += f" (+{len(encontradas_vis) - 1} más)"
            return resumen, primer_rect_vis

    return "Coincidencia en página", None


def generar_excel_bytes(resultados: List[dict], termino: str) -> bytes:
    encabezados = [
        "Archivo PDF",
        "Ruta",
        "Página Exacta",
        "Total Páginas",
        "Casilla Exacta",
        "Coincidencias",
        "Término Buscado",
        "Contexto / Artículo Encontrado",
    ]
    filas = [encabezados]
    for r in resultados:
        filas.append([
            r["Archivo PDF"],
            r["Ruta"],
            r["Página Exacta"],
            r["Total Páginas"],
            r["Casilla Exacta"],
            r["Coincidencias"],
            termino,
            r["Contexto / Artículo"],
        ])

    if not TIENE_OPENPYXL:
        return _generar_xlsx_puro_python(filas)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados PDF"
    ws.append(encabezados)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="left", vertical="center")

    for fila in filas[1:]:
        ws.append(fila)

    for idx, ancho in enumerate([32, 36, 14, 14, 30, 14, 24, 62], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = ancho

    ws.freeze_panes = "A2"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def renderizar_pagina_png(
    raw_bytes: bytes,
    num_pagina: int,
    termino: str,
    zoom: float,
    casilla_rect: Optional[Tuple[float, float, float, float]] = None,
    enfocar_casilla: bool = False,
) -> bytes:
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

        clip_rect = None
        if casilla_rect is not None:
            rect_box = fitz.Rect(casilla_rect) + (-4, -3, 4, 3)
            page.draw_rect(rect_box, color=(0.11, 0.31, 0.85), width=1.8)
            if enfocar_casilla:
                margen_x = max(120.0, rect_box.width * 0.6)
                margen_y = max(75.0, rect_box.height * 2.5)
                clip_rect = (rect_box + (-margen_x, -margen_y, margen_x, margen_y)) & page.rect

        dpi_scale = max(zoom, 1.5)
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale), clip=clip_rect, alpha=False)
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
            casilla_txt, casilla_rect = detectar_casilla_pagina(
                st.session_state.docs[pag.file_id],
                pag.numero_pagina,
                termino.strip(),
                frase_exacta,
            )
            resultados.append(
                {
                    "file_id": pag.file_id,
                    "casilla_rect": casilla_rect,
                    "Archivo PDF": pag.nombre_archivo,
                    "Ruta": pag.ruta_relativa,
                    "Página Exacta": pag.numero_pagina,
                    "Total Páginas": pag.total_paginas,
                    "Casilla Exacta": casilla_txt,
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
            {k: v for k, v in r.items() if k not in ("file_id", "casilla_rect")} for r in resultados
        ]
        st.dataframe(tabla_visible, use_container_width=True, hide_index=True)

    with col_preview:
        opciones = [
            f"{i + 1}. {r['Archivo PDF']} — Página {r['Página Exacta']} de {r['Total Páginas']} · {r['Casilla Exacta']}"
            for i, r in enumerate(resultados)
        ]
        idx_sel = st.selectbox("Vista previa de página encontrada", range(len(opciones)), format_func=lambda i: opciones[i])
        sel = resultados[idx_sel]
        st.caption(f"Ubicación exacta: **{sel['Casilla Exacta']}** · Página {sel['Página Exacta']} de {sel['Total Páginas']}")

        col_z1, col_z2 = st.columns([2, 1])
        with col_z1:
            zoom_pct = st.slider("Zoom de vista previa (%)", 100, 300, 125, 25)
        with col_z2:
            enfocar_casilla = st.toggle("Enfocar casilla", value=False, disabled=sel["casilla_rect"] is None)

        zoom_factor = zoom_pct / 100.0
        img_bytes = renderizar_pagina_png(
            st.session_state.docs[sel["file_id"]],
            sel["Página Exacta"],
            termino.strip(),
            zoom_factor if not enfocar_casilla else max(zoom_factor, 2.5),
            casilla_rect=sel["casilla_rect"],
            enfocar_casilla=enfocar_casilla,
        )
        img_b64 = base64.b64encode(img_bytes).decode("ascii")
        ancho_css = "100%" if enfocar_casilla or zoom_pct == 100 else f"{zoom_pct}%"
        st.markdown(
            f'<div style="overflow:auto; max-height:720px; border:1px solid #cbd5e1; border-radius:8px; background:#e2e8f0; padding:10px;">'
            f'<img src="data:image/png;base64,{img_b64}" style="width:{ancho_css}; max-width:none; height:auto; display:block; background:#fff; box-shadow:0 2px 6px rgba(0,0,0,0.12);" />'
            f'</div>',
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
