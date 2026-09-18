import streamlit as st
import pandas as pd
import re
import io
import os
import csv
import shutil
import zipfile
import tempfile
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path


st.set_page_config(page_title="Validación BNC - R34 vs Recaudación", layout="wide")

st.title("Validación BNC — R34 vs Recaudación")
st.caption(
    "Detecta registros del R34 que cumplen los filtros definidos y verifica si existen en "
    "Recaudación Agentes Autorizados."
)


# ============================================================
# NORMALIZACIÓN
# ============================================================

def quitar_acentos(texto):
    texto = "" if texto is None else str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_texto(valor):
    s = quitar_acentos(valor).upper().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalizar_compacto(valor):
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(valor))


def limpiar_id(valor):
    """Convierte IDs numéricos a una forma comparable: 002 -> 2, 199.0 -> 199."""
    if valor is None or pd.isna(valor):
        return ""
    s = str(valor).strip()
    if not s or s.upper() in {"NAN", "NONE", "N/D", "NA", "N/A"}:
        return ""

    s_num = s.replace(" ", "").replace(",", ".")
    try:
        d = Decimal(s_num)
        if d == d.to_integral_value():
            return str(int(d))
    except (InvalidOperation, ValueError):
        pass

    # Último recurso: conservar solo dígitos.
    digitos = re.sub(r"\D", "", s)
    if not digitos:
        return ""
    return digitos.lstrip("0") or "0"


def serie_ids(serie):
    return serie.map(limpiar_id)


def buscar_columna(columnas, aliases, obligatoria=True):
    """Busca columnas por nombre, no por posición."""
    mapa = {normalizar_compacto(c): c for c in columnas}

    # 1) Coincidencia exacta normalizada.
    for alias in aliases:
        a = normalizar_compacto(alias)
        if a in mapa:
            return mapa[a]

    # 2) Coincidencia parcial, solo si es inequívoca.
    for alias in aliases:
        a = normalizar_compacto(alias)
        candidatos = [orig for norm, orig in mapa.items() if a in norm or norm in a]
        if len(candidatos) == 1:
            return candidatos[0]

    if obligatoria:
        raise ValueError(
            f"No se encontró una columna requerida. Busqué: {', '.join(aliases)}"
        )
    return None


# ============================================================
# DICCIONARIO MANEJADOR CCRPOS
# ============================================================

def cargar_diccionario(contenido):
    bio = io.BytesIO(contenido)
    xls = pd.ExcelFile(bio, engine="openpyxl")

    hoja = next(
        (s for s in xls.sheet_names if normalizar_compacto(s) == "MANEJADORCCRPOS"),
        None,
    )
    if hoja is None:
        raise ValueError(
            "El diccionario no contiene la hoja 'MANEJADOR CCRPOS'."
        )

    df = pd.read_excel(xls, sheet_name=hoja, dtype=str, keep_default_na=False)
    col_m1 = buscar_columna(df.columns, ["MANEJADOR"])
    col_m2 = buscar_columna(df.columns, ["MANEJADOR2", "MANEJADOR 2"])

    tmp = df[[col_m1, col_m2]].copy()
    tmp["_M1"] = tmp[col_m1].map(normalizar_texto)
    tmp["_M2"] = tmp[col_m2].astype(str).str.strip()
    tmp = tmp[(tmp["_M1"] != "") & (tmp["_M2"] != "")]
    tmp = tmp.drop_duplicates("_M1", keep="first")

    return dict(zip(tmp["_M1"], tmp["_M2"]))


# ============================================================
# RECAUDACIÓN
# ============================================================

def hojas_bnc(contenido):
    xls = pd.ExcelFile(io.BytesIO(contenido), engine="openpyxl")
    return [s for s in xls.sheet_names if normalizar_compacto(s).startswith("BNC")]


def hoja_bnc_esperada(nombre_archivo, hojas):
    meses = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
        "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
        "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
        "NOVIEMBRE": 11, "DICIEMBRE": 12,
    }
    meses_nombre = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }

    nombre = normalizar_texto(nombre_archivo)
    anio_match = re.search(r"20\d{2}", nombre)
    mes_encontrado = next((m for m in meses if m in nombre), None)

    if not anio_match or not mes_encontrado:
        return hojas[0] if len(hojas) == 1 else None

    anio = int(anio_match.group())
    mes = meses[mes_encontrado] - 1
    if mes == 0:
        mes = 12
        anio -= 1

    esperado = normalizar_compacto(f"BNC {meses_nombre[mes]} {anio}")
    for hoja in hojas:
        if normalizar_compacto(hoja) == esperado:
            return hoja

    return hojas[0] if len(hojas) == 1 else None


def unir_unicos(serie):
    vistos = []
    for v in serie:
        s = str(v).strip()
        if not s or s.upper() in {"NAN", "NONE"}:
            continue
        if s not in vistos:
            vistos.append(s)
    return " | ".join(vistos)


def cargar_recaudacion(contenido, hoja):
    df = pd.read_excel(
        io.BytesIO(contenido),
        sheet_name=hoja,
        dtype=str,
        keep_default_na=False,
        engine="openpyxl",
    )

    col_concat = buscar_columna(df.columns, ["CONCATENAR"], obligatoria=False)
    col_afiliado = buscar_columna(
        df.columns, ["AFILIADO", "CODIGO AFIL", "CODIGO_AFIL"], obligatoria=False
    )
    col_terminal = buscar_columna(
        df.columns, ["TERMINAL", "NUMPOS", "NUM POS"], obligatoria=False
    )

    # Preferimos construir la llave por componentes para evitar notación científica.
    if col_afiliado and col_terminal:
        df["CLAVE_CRUCE"] = serie_ids(df[col_afiliado]) + serie_ids(df[col_terminal])
    elif col_concat:
        df["CLAVE_CRUCE"] = serie_ids(df[col_concat])
    else:
        raise ValueError(
            "En Recaudación necesito 'Concatenar' o las columnas 'Afiliado' + 'Terminal'."
        )

    df = df[df["CLAVE_CRUCE"] != ""].copy()

    col_pago_ccr = buscar_columna(df.columns, ["PAGO CCR $", "PAGO CCR"], obligatoria=False)
    col_pago_ban = buscar_columna(df.columns, ["PAGO BAN $", "PAGO BAN"], obligatoria=False)
    col_pago_aa = buscar_columna(df.columns, ["PAGO A.A $", "PAGO AA $", "PAGO A.A"], obligatoria=False)
    col_fecha = buscar_columna(
        df.columns, ["FECHA", "FECHA PAGO", "FECHA COBRO", "FECHA DE PAGO"], obligatoria=False
    )

    base = (
        df.groupby("CLAVE_CRUCE", as_index=False)
        .size()
        .rename(columns={"size": "COINCIDENCIAS_RECAUDACION"})
    )

    extras = [
        (col_pago_ccr, "PAGO_CCR_RECAUDACION"),
        (col_pago_ban, "PAGO_BAN_RECAUDACION"),
        (col_pago_aa, "PAGO_AA_RECAUDACION"),
        (col_fecha, "FECHAS_RECAUDACION"),
    ]

    for original, nuevo in extras:
        if original:
            tmp = (
                df.groupby("CLAVE_CRUCE")[original]
                .apply(unir_unicos)
                .reset_index(name=nuevo)
            )
            base = base.merge(tmp, on="CLAVE_CRUCE", how="left")

    return base


# ============================================================
# R34
# ============================================================

def resolver_columnas_r34(columnas):
    return {
        "PERTENENCIA": buscar_columna(columnas, ["PERTENENCIA"]),
        "MANEJADOR": buscar_columna(columnas, ["MANEJADOR"]),
        "POS_CON_TRANSACCION": buscar_columna(
            columnas,
            ["POS_CON_TRANSACCION", "POS CON TRANSACCION", "POSCONTRANSACCION"],
        ),
        "NOMBRE_BANCO": buscar_columna(
            columnas, ["NOMBRE_BANCO", "NOMBRE BANCO", "NOMBREBANCO"]
        ),
        "CODIGO_AFIL": buscar_columna(
            columnas, ["CODIGO_AFIL", "CODIGO AFIL", "CODIGO_AFILIADO", "AFILIADO"]
        ),
        "NUMPOS": buscar_columna(
            columnas, ["NUMPOS", "NUM POS", "NUM_POS", "NUMERO POS"]
        ),
        "NOMBRE_AFILIADO": buscar_columna(
            columnas, ["NOMBRE_AFILIADO", "NOMBRE AFILIADO", "NOMBRE_AFIL"], obligatoria=False
        ),
        "RIF_AFILIADO": buscar_columna(
            columnas, ["RIF_AFILIADO", "RIF AFILIADO", "RIF"], obligatoria=False
        ),
        "CIUDAD": buscar_columna(columnas, ["CIUDAD"], obligatoria=False),
        "ESTADO": buscar_columna(columnas, ["ESTADO"], obligatoria=False),
        "TERMINAL": buscar_columna(columnas, ["TERMINAL"], obligatoria=False),
        "SERIAL": buscar_columna(columnas, ["SERIAL"], obligatoria=False),
        "AFIPOS": buscar_columna(columnas, ["AFIPOS"], obligatoria=False),
    }


def filtrar_chunk_r34(df, columnas, mapa_manejadores, banco_objetivo, archivo_origen):
    c = columnas

    pertenencia = df[c["PERTENENCIA"]].map(normalizar_compacto)
    mask_pertenencia = (
        pertenencia.str.contains("CREDICARDPOS", na=False)
        | pertenencia.str.contains("ESPECIALCENTRO", na=False)
        | pertenencia.str.contains("ESPECIALORIENTE", na=False)
        | pertenencia.str.contains("ESPECIALOCCIDENTE", na=False)
    )

    manejador_norm = df[c["MANEJADOR"]].map(normalizar_texto)
    manejador2 = manejador_norm.map(mapa_manejadores)
    mask_manejador = manejador2.notna()

    pos = pd.to_numeric(
        df[c["POS_CON_TRANSACCION"]].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    mask_pos = pos.eq(1)

    banco_norm = df[c["NOMBRE_BANCO"]].map(normalizar_compacto)
    banco_target = normalizar_compacto(banco_objetivo)
    mask_banco = banco_norm.str.contains(re.escape(banco_target), na=False)

    mask = mask_pertenencia & mask_manejador & mask_pos & mask_banco
    f = df.loc[mask].copy()

    if f.empty:
        return pd.DataFrame()

    out = pd.DataFrame(index=f.index)
    out["ARCHIVO_R34"] = archivo_origen
    out["CODIGO_AFIL"] = f[c["CODIGO_AFIL"]].astype(str).str.strip()
    out["NUMPOS"] = f[c["NUMPOS"]].astype(str).str.strip()
    out["CLAVE_CRUCE"] = serie_ids(f[c["CODIGO_AFIL"]]) + serie_ids(f[c["NUMPOS"]])

    for salida in [
        "NOMBRE_AFILIADO", "RIF_AFILIADO", "CIUDAD", "ESTADO",
        "TERMINAL", "SERIAL", "AFIPOS"
    ]:
        original = c.get(salida)
        out[salida] = f[original].astype(str).str.strip() if original else ""

    out["PERTENENCIA"] = f[c["PERTENENCIA"]].astype(str).str.strip()
    out["MANEJADOR"] = f[c["MANEJADOR"]].astype(str).str.strip()
    out["MANEJADOR2"] = manejador2.loc[f.index]
    out["NOMBRE_BANCO"] = f[c["NOMBRE_BANCO"]].astype(str).str.strip()
    out["POS_CON_TRANSACCION"] = f[c["POS_CON_TRANSACCION"]].astype(str).str.strip()

    out = out[out["CLAVE_CRUCE"] != ""].reset_index(drop=True)
    return out


def detectar_csv(path):
    raw = Path(path).read_bytes()[:120000]
    texto = None
    encoding = None

    for enc in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            texto = raw.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            continue

    if texto is None:
        texto = raw.decode("latin1", errors="replace")
        encoding = "latin1"

    try:
        dialect = csv.Sniffer().sniff(texto, delimiters=",;\t|")
        sep = dialect.delimiter
    except csv.Error:
        conteos = {s: texto.count(s) for s in [",", ";", "\t", "|"]}
        sep = max(conteos, key=conteos.get)

    return encoding, sep


def procesar_csv_r34(path, mapa_manejadores, banco_objetivo, nombre_origen):
    encoding, sep = detectar_csv(path)

    cabecera = pd.read_csv(
        path,
        sep=sep,
        encoding=encoding,
        dtype=str,
        keep_default_na=False,
        nrows=0,
    )
    columnas = resolver_columnas_r34(cabecera.columns)

    necesarias = list(dict.fromkeys(v for v in columnas.values() if v))
    partes = []

    for chunk in pd.read_csv(
        path,
        sep=sep,
        encoding=encoding,
        dtype=str,
        keep_default_na=False,
        usecols=necesarias,
        chunksize=75000,
        on_bad_lines="skip",
    ):
        filtrado = filtrar_chunk_r34(
            chunk, columnas, mapa_manejadores, banco_objetivo, nombre_origen
        )
        if not filtrado.empty:
            partes.append(filtrado)

    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()


def procesar_excel_r34(path, mapa_manejadores, banco_objetivo, nombre_origen):
    df = pd.read_excel(path, dtype=str, keep_default_na=False, engine="openpyxl")
    columnas = resolver_columnas_r34(df.columns)
    return filtrar_chunk_r34(
        df, columnas, mapa_manejadores, banco_objetivo, nombre_origen
    )


def procesar_r34_subarchivo(path, mapa_manejadores, banco_objetivo, nombre_origen):
    ext = Path(path).suffix.lower()
    if ext == ".csv" or ext == ".txt":
        return procesar_csv_r34(path, mapa_manejadores, banco_objetivo, nombre_origen)
    if ext in {".xlsx", ".xlsm"}:
        return procesar_excel_r34(path, mapa_manejadores, banco_objetivo, nombre_origen)
    raise ValueError(f"Formato R34 no soportado: {ext}")


def procesar_upload_r34(upload, mapa_manejadores, banco_objetivo):
    resultados = []

    with tempfile.TemporaryDirectory() as td:
        ruta = os.path.join(td, Path(upload.name).name)
        with open(ruta, "wb") as f:
            f.write(upload.getbuffer())

        ext = Path(ruta).suffix.lower()

        if ext == ".zip":
            with zipfile.ZipFile(ruta) as z:
                miembros = [
                    m for m in z.namelist()
                    if Path(m).suffix.lower() in {".csv", ".txt", ".xlsx", ".xlsm"}
                ]
                if not miembros:
                    raise ValueError(f"{upload.name}: el ZIP no contiene CSV/XLSX.")

                for i, miembro in enumerate(miembros, start=1):
                    destino = os.path.join(td, f"extraido_{i}{Path(miembro).suffix.lower()}")
                    with z.open(miembro) as src, open(destino, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                    nombre_origen = f"{upload.name} > {Path(miembro).name}"
                    r = procesar_r34_subarchivo(
                        destino, mapa_manejadores, banco_objetivo, nombre_origen
                    )
                    if not r.empty:
                        resultados.append(r)
        else:
            r = procesar_r34_subarchivo(
                ruta, mapa_manejadores, banco_objetivo, upload.name
            )
            if not r.empty:
                resultados.append(r)

    return pd.concat(resultados, ignore_index=True) if resultados else pd.DataFrame()


# ============================================================
# RESUMEN Y EXCEL FINAL
# ============================================================

def crear_resumen(df):
    filas = []
    for manejador, g in df.groupby("MANEJADOR2", dropna=False):
        total = len(g)
        cruzados = int((g["ESTADO_CRUCE"] == "CRUZADO").sum())
        no_encontrados = total - cruzados
        clientes = g["CODIGO_AFIL"].nunique()
        clientes_faltantes = g.loc[
            g["ESTADO_CRUCE"] == "NO ENCONTRADO", "CODIGO_AFIL"
        ].nunique()

        filas.append({
            "MANEJADOR2": manejador or "SIN MANEJADOR",
            "REGISTROS_R34": total,
            "CLIENTES_UNICOS": clientes,
            "CRUZADOS": cruzados,
            "NO_ENCONTRADOS": no_encontrados,
            "CLIENTES_UNICOS_NO_ENCONTRADOS": clientes_faltantes,
            "%_CRUCE": cruzados / total if total else 0,
            "%_NO_ENCONTRADO": no_encontrados / total if total else 0,
        })

    resumen = pd.DataFrame(filas)
    if resumen.empty:
        return resumen

    resumen = resumen.sort_values(
        ["NO_ENCONTRADOS", "REGISTROS_R34"], ascending=[False, False]
    ).reset_index(drop=True)

    total = {
        "MANEJADOR2": "TOTAL",
        "REGISTROS_R34": len(df),
        "CLIENTES_UNICOS": df["CODIGO_AFIL"].nunique(),
        "CRUZADOS": int((df["ESTADO_CRUCE"] == "CRUZADO").sum()),
        "NO_ENCONTRADOS": int((df["ESTADO_CRUCE"] == "NO ENCONTRADO").sum()),
        "CLIENTES_UNICOS_NO_ENCONTRADOS": df.loc[
            df["ESTADO_CRUCE"] == "NO ENCONTRADO", "CODIGO_AFIL"
        ].nunique(),
    }
    total["%_CRUCE"] = total["CRUZADOS"] / total["REGISTROS_R34"] if total["REGISTROS_R34"] else 0
    total["%_NO_ENCONTRADO"] = total["NO_ENCONTRADOS"] / total["REGISTROS_R34"] if total["REGISTROS_R34"] else 0

    return pd.concat([resumen, pd.DataFrame([total])], ignore_index=True)


def ajustar_anchos(worksheet, df, inicio_col=0, max_width=34):
    for i, col in enumerate(df.columns):
        valores = df[col].astype(str).head(300)
        ancho = max(len(str(col)), *(len(v) for v in valores)) + 2 if len(valores) else len(str(col)) + 2
        worksheet.set_column(inicio_col + i, inicio_col + i, min(ancho, max_width))


def crear_excel(resultado, resumen, hoja_bnc, banco_objetivo):
    no_encontrados = resultado[resultado["ESTADO_CRUCE"] == "NO ENCONTRADO"].copy()
    cruzados = resultado[resultado["ESTADO_CRUCE"] == "CRUZADO"].copy()

    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_titulo = workbook.add_format({
            "bold": True, "font_size": 16, "font_color": "#FFFFFF",
            "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"
        })
        fmt_header = workbook.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True
        })
        fmt_label = workbook.add_format({
            "bold": True, "bg_color": "#D9EAF7", "border": 1
        })
        fmt_valor = workbook.add_format({"border": 1})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_rojo = workbook.add_format({"bg_color": "#FCE4D6"})
        fmt_verde = workbook.add_format({"bg_color": "#E2F0D9"})

        # -------- RESUMEN --------
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False, startrow=9)
        ws = writer.sheets["RESUMEN"]
        ws.merge_range("A1:H1", "VALIDACIÓN BNC — R34 VS RECAUDACIÓN", fmt_titulo)
        ws.set_row(0, 28)

        total_reg = len(resultado)
        total_cruz = int((resultado["ESTADO_CRUCE"] == "CRUZADO").sum())
        total_no = int((resultado["ESTADO_CRUCE"] == "NO ENCONTRADO").sum())
        clientes_no = resultado.loc[
            resultado["ESTADO_CRUCE"] == "NO ENCONTRADO", "CODIGO_AFIL"
        ].nunique()

        metricas = [
            ("Período / hoja BNC", hoja_bnc),
            ("Banco filtrado en R34", banco_objetivo),
            ("Registros R34 válidos", total_reg),
            ("Casos cruzados", total_cruz),
            ("Casos no encontrados", total_no),
            ("Clientes únicos no encontrados", clientes_no),
            ("% cruce", total_cruz / total_reg if total_reg else 0),
        ]

        for fila, (label, valor) in enumerate(metricas, start=2):
            ws.write(fila, 0, label, fmt_label)
            if label == "% cruce":
                ws.write(fila, 1, valor, fmt_pct)
            else:
                ws.write(fila, 1, valor, fmt_valor)

        ws.write(8, 0, "Resumen por MANEJADOR2", fmt_label)
        for col_idx, col in enumerate(resumen.columns):
            ws.write(9, col_idx, col, fmt_header)
        if not resumen.empty:
            ws.autofilter(9, 0, 9 + len(resumen), len(resumen.columns) - 1)
        ws.freeze_panes(10, 0)
        ajustar_anchos(ws, resumen)
        ws.set_column(0, 0, 28)
        ws.set_column(6, 7, 16, fmt_pct)

        # -------- DETALLES --------
        for nombre_hoja, df_detalle, color_fmt in [
            ("NO ENCONTRADOS", no_encontrados, fmt_rojo),
            ("CASOS CRUZADOS", cruzados, fmt_verde),
        ]:
            df_detalle.to_excel(writer, sheet_name=nombre_hoja, index=False)
            wsd = writer.sheets[nombre_hoja]
            for col_idx, col in enumerate(df_detalle.columns):
                wsd.write(0, col_idx, col, fmt_header)
            wsd.freeze_panes(1, 0)
            if len(df_detalle) > 0:
                wsd.autofilter(0, 0, len(df_detalle), len(df_detalle.columns) - 1)
                estado_idx = df_detalle.columns.get_loc("ESTADO_CRUCE")
                wsd.set_column(estado_idx, estado_idx, 20, color_fmt)
            ajustar_anchos(wsd, df_detalle)

    salida.seek(0)
    return salida.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

col1, col2 = st.columns(2)
with col1:
    archivo_recaudacion = st.file_uploader(
        "1. Recaudación Agentes Autorizados",
        type=["xlsx"],
        key="recaudacion",
    )
with col2:
    archivo_diccionario = st.file_uploader(
        "2. Diccionario de manejadores",
        type=["xlsx"],
        key="diccionario",
    )

archivos_r34 = st.file_uploader(
    "3. R34 (puedes subir CSV, XLSX o ZIP; también varios archivos)",
    type=["csv", "txt", "xlsx", "xlsm", "zip"],
    accept_multiple_files=True,
    key="r34",
)

banco_objetivo = st.text_input(
    "Banco que debe aparecer en NOMBRE_BANCO del R34",
    value="B.O.D",
    help="Lo dejé en B.O.D porque esa fue la regla indicada. Si cambia, puedes escribir otro valor aquí.",
)

hoja_bnc = None
if archivo_recaudacion is not None:
    try:
        rec_bytes = archivo_recaudacion.getvalue()
        opciones_bnc = hojas_bnc(rec_bytes)
        if not opciones_bnc:
            st.error("No encontré ninguna hoja cuyo nombre empiece por BNC.")
        else:
            sugerida = hoja_bnc_esperada(archivo_recaudacion.name, opciones_bnc)
            indice = opciones_bnc.index(sugerida) if sugerida in opciones_bnc else 0
            hoja_bnc = st.selectbox(
                "Hoja BNC a utilizar",
                opciones_bnc,
                index=indice,
            )
    except Exception as e:
        st.error(f"No pude leer las hojas de Recaudación: {e}")

if st.button("Procesar validación", type="primary", use_container_width=True):
    if archivo_recaudacion is None or archivo_diccionario is None or not archivos_r34:
        st.error("Debes subir Recaudación, el diccionario y al menos un R34.")
    elif hoja_bnc is None:
        st.error("Debes seleccionar una hoja BNC válida.")
    else:
        try:
            with st.spinner("Procesando archivos..."):
                mapa = cargar_diccionario(archivo_diccionario.getvalue())
                rec_agg = cargar_recaudacion(archivo_recaudacion.getvalue(), hoja_bnc)

                partes = []
                for i, archivo in enumerate(archivos_r34, start=1):
                    st.write(f"Procesando R34 {i}/{len(archivos_r34)}: {archivo.name}")
                    parte = procesar_upload_r34(archivo, mapa, banco_objetivo)
                    if not parte.empty:
                        partes.append(parte)

                if not partes:
                    raise ValueError(
                        "Después de aplicar los filtros no quedó ningún registro del R34. "
                        "Revisa PERTENENCIA, MANEJADOR, POS_CON_TRANSACCION y NOMBRE_BANCO."
                    )

                r34_filtrado = pd.concat(partes, ignore_index=True)
                resultado = r34_filtrado.merge(rec_agg, on="CLAVE_CRUCE", how="left")
                resultado["ESTADO_CRUCE"] = resultado["COINCIDENCIAS_RECAUDACION"].apply(
                    lambda x: "CRUZADO" if pd.notna(x) else "NO ENCONTRADO"
                )
                resultado["COINCIDENCIAS_RECAUDACION"] = (
                    resultado["COINCIDENCIAS_RECAUDACION"].fillna(0).astype(int)
                )

                # Orden de columnas más útil para revisión.
                primeras = [
                    "ESTADO_CRUCE", "CLAVE_CRUCE", "CODIGO_AFIL", "NUMPOS",
                    "NOMBRE_AFILIADO", "RIF_AFILIADO", "PERTENENCIA",
                    "MANEJADOR", "MANEJADOR2", "NOMBRE_BANCO",
                    "POS_CON_TRANSACCION", "CIUDAD", "ESTADO",
                    "COINCIDENCIAS_RECAUDACION", "ARCHIVO_R34"
                ]
                primeras = [c for c in primeras if c in resultado.columns]
                resto = [c for c in resultado.columns if c not in primeras]
                resultado = resultado[primeras + resto]

                resumen = crear_resumen(resultado)
                excel = crear_excel(resultado, resumen, hoja_bnc, banco_objetivo)

                periodo_archivo = re.sub(r"[^A-Za-z0-9_-]+", "_", hoja_bnc).strip("_")
                nombre_salida = f"Validacion_BNC_{periodo_archivo}.xlsx"

                st.session_state["resultado_excel"] = excel
                st.session_state["nombre_salida"] = nombre_salida
                st.session_state["metricas_resultado"] = {
                    "validos": len(resultado),
                    "cruzados": int((resultado["ESTADO_CRUCE"] == "CRUZADO").sum()),
                    "no_encontrados": int((resultado["ESTADO_CRUCE"] == "NO ENCONTRADO").sum()),
                    "clientes_no": resultado.loc[
                        resultado["ESTADO_CRUCE"] == "NO ENCONTRADO", "CODIGO_AFIL"
                    ].nunique(),
                }

            st.success("Validación terminada.")

        except Exception as e:
            st.exception(e)


if st.session_state.get("resultado_excel"):
    m = st.session_state["metricas_resultado"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros R34 válidos", f"{m['validos']:,}")
    c2.metric("Casos cruzados", f"{m['cruzados']:,}")
    c3.metric("No encontrados", f"{m['no_encontrados']:,}")
    c4.metric("Clientes únicos faltantes", f"{m['clientes_no']:,}")

    st.download_button(
        "Descargar Excel de validación",
        data=st.session_state["resultado_excel"],
        file_name=st.session_state["nombre_salida"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
