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
from openpyxl import load_workbook


st.set_page_config(page_title="Validación BNC - R34 vs Recaudación", layout="wide")

st.title("Validación BNC — R34 vs Recaudación")
st.caption(
    "Filtra el R34 con las reglas definidas y compara la hoja BNC seleccionada contra esos registros."
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
    """
    Convierte identificadores numéricos a texto comparable.
    Ejemplos: 002 -> 2, 199.0 -> 199, 8.6049082E+07 -> 86049082.
    """
    if valor is None or pd.isna(valor):
        return ""

    s = str(valor).strip()
    if not s or normalizar_texto(s) in {"NAN", "NONE", "N/D", "NA", "N/A"}:
        return ""

    s_num = s.replace(" ", "").replace(",", ".")
    try:
        d = Decimal(s_num)
        if d == d.to_integral_value():
            return str(int(d))
    except (InvalidOperation, ValueError):
        pass

    digitos = re.sub(r"\D", "", s)
    if not digitos:
        return ""
    return digitos.lstrip("0") or "0"


def serie_ids(serie):
    return serie.map(limpiar_id)


def buscar_columna(columnas, aliases, obligatoria=True):
    """Busca una columna por nombre normalizado, nunca por letra/posición."""
    mapa = {normalizar_compacto(c): c for c in columnas}

    # Coincidencia exacta normalizada.
    for alias in aliases:
        a = normalizar_compacto(alias)
        if a in mapa:
            return mapa[a]

    # Coincidencia parcial solo si es inequívoca.
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
# DICCIONARIO — SOLO HOJA MANEJADOR CCRPOS
# ============================================================

def cargar_diccionario(contenido):
    bio = io.BytesIO(contenido)
    xls = pd.ExcelFile(bio, engine="openpyxl")

    hoja = next(
        (s for s in xls.sheet_names if normalizar_compacto(s) == "MANEJADORCCRPOS"),
        None,
    )
    if hoja is None:
        raise ValueError("El diccionario no contiene la hoja 'MANEJADOR CCRPOS'.")

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
    """Si el archivo dice Mayo 2026, intenta sugerir BNC Abril 2026."""
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


def cargar_recaudacion(contenido, hoja, mapa_manejadores):
    """
    Carga la hoja BNC completa.

    IMPORTANTE:
    - CONCATENAR es la llave principal.
    - Si CONCATENAR no existe, usa Afiliado + Terminal.
    - Si un CONCATENAR aparece varias veces en BNC, NO es error.
      Cada fila se conserva; para el cruce basta con que la clave exista en R34.
    """
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

    if col_concat:
        df["CLAVE_CRUCE"] = serie_ids(df[col_concat])
    elif col_afiliado and col_terminal:
        df["CLAVE_CRUCE"] = serie_ids(df[col_afiliado]) + serie_ids(df[col_terminal])
    else:
        raise ValueError(
            "En la hoja BNC necesito 'Concatenar' o las columnas 'Afiliado' + 'Terminal'."
        )

    df = df[df["CLAVE_CRUCE"] != ""].copy()

    # Columnas relevantes de BNC.
    aliases = {
        "BNC_FECHA_DOC": ["FECHA DOC.", "FECHA DOC", "FECHA"],
        "BNC_CLIENTE": ["CLIENTE"],
        "BNC_NOMBRE": ["NOMBRE 1", "NOMBRE"],
        "BNC_CONCATENAR": ["CONCATENAR"],
        "BNC_AFILIADO": ["AFILIADO"],
        "BNC_TERMINAL": ["TERMINAL"],
        "BNC_PAGO_CCR_USD": ["PAGO CCR $", "PAGO CCR"],
        "BNC_PAGO_BAN_USD": ["PAGO BAN $", "PAGO BAN"],
        "BNC_PAGO_AA_USD": ["PAGO A.A $", "PAGO AA $", "PAGO A.A"],
        "BNC_MANEJADOR_R34": ["MANEJADOR R34"],
        "BNC_MANEJADOR_R34_MODIFICADO": ["MANEJADOR R34 MODIFICADO"],
        "BNC_EQUIPO_R34": ["EQUIPO R34"],
    }

    out = pd.DataFrame(index=df.index)
    out["CLAVE_CRUCE"] = df["CLAVE_CRUCE"]

    for salida, opciones in aliases.items():
        col = buscar_columna(df.columns, opciones, obligatoria=False)
        out[salida] = df[col].astype(str).str.strip() if col else ""

    # MANEJADOR2 para el resumen.
    manejador_mod = out["BNC_MANEJADOR_R34_MODIFICADO"].astype(str).str.strip()
    manejador_orig = out["BNC_MANEJADOR_R34"].map(normalizar_texto)
    desde_diccionario = manejador_orig.map(mapa_manejadores).fillna("")
    out["MANEJADOR2_BNC"] = manejador_mod.where(manejador_mod != "", desde_diccionario)
    out.loc[out["MANEJADOR2_BNC"].astype(str).str.strip() == "", "MANEJADOR2_BNC"] = "SIN MANEJADOR"

    return out.reset_index(drop=True)


# ============================================================
# R34 — DETECCIÓN DE HOJA/ENCABEZADO + FILTROS
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
        "CONCATENAR": buscar_columna(columnas, ["CONCATENAR"], obligatoria=False),
        "CODIGO_AFIL": buscar_columna(
            columnas,
            ["CODIGO_AFIL", "CODIGO AFIL", "CODIGO_AFILIADO", "AFILIADO"],
            obligatoria=False,
        ),
        "NUMPOS": buscar_columna(
            columnas, ["NUMPOS", "NUM POS", "NUM_POS", "NUMERO POS"], obligatoria=False
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


def validar_llave_r34(columnas):
    if columnas.get("CONCATENAR"):
        return
    if columnas.get("CODIGO_AFIL") and columnas.get("NUMPOS"):
        return
    raise ValueError(
        "En el R34 necesito 'Concatenar' o las columnas 'CODIGO_AFIL' + 'NUMPOS' para crear la llave."
    )


def pertenencia_valida(serie):
    p = serie.map(normalizar_compacto)
    return (
        p.str.contains("CREDICARDPOS", na=False)
        | p.str.contains("ESPECIALCENTRO", na=False)
        | p.str.contains("ESPECIALORIENTE", na=False)
        | p.str.contains("ESPECIALOCCIDENTE", na=False)
    )


def filtrar_chunk_r34(df, columnas, mapa_manejadores, banco_objetivo, archivo_origen):
    c = columnas
    validar_llave_r34(c)

    # 1) PERTENENCIA válida.
    mask_pertenencia = pertenencia_valida(df[c["PERTENENCIA"]])

    # 2) MANEJADOR debe estar en la hoja MANEJADOR CCRPOS del diccionario.
    manejador_norm = df[c["MANEJADOR"]].map(normalizar_texto)
    manejador2 = manejador_norm.map(mapa_manejadores)
    mask_manejador = manejador2.notna()

    # 3) POS_CON_TRANSACCION = 1.
    pos = pd.to_numeric(
        df[c["POS_CON_TRANSACCION"]].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    mask_pos = pos.eq(1)

    # 4) NOMBRE_BANCO = B.O.D (normalizado, por eso B.O.D / BOD funcionan igual).
    banco_norm = df[c["NOMBRE_BANCO"]].map(normalizar_compacto)
    banco_target = normalizar_compacto(banco_objetivo)
    mask_banco = banco_norm.str.contains(re.escape(banco_target), na=False)

    mask = mask_pertenencia & mask_manejador & mask_pos & mask_banco
    f = df.loc[mask].copy()

    if f.empty:
        return pd.DataFrame()

    out = pd.DataFrame(index=f.index)
    out["ARCHIVO_R34"] = archivo_origen

    # CONCATENAR directo si existe; si no, CODIGO_AFIL + NUMPOS.
    if c.get("CONCATENAR"):
        out["CLAVE_CRUCE"] = serie_ids(f[c["CONCATENAR"]])
    else:
        out["CLAVE_CRUCE"] = serie_ids(f[c["CODIGO_AFIL"]]) + serie_ids(f[c["NUMPOS"]])

    out["R34_CODIGO_AFIL"] = (
        f[c["CODIGO_AFIL"]].astype(str).str.strip() if c.get("CODIGO_AFIL") else ""
    )
    out["R34_NUMPOS"] = (
        f[c["NUMPOS"]].astype(str).str.strip() if c.get("NUMPOS") else ""
    )

    opcionales = {
        "R34_NOMBRE_AFILIADO": "NOMBRE_AFILIADO",
        "R34_RIF_AFILIADO": "RIF_AFILIADO",
        "R34_CIUDAD": "CIUDAD",
        "R34_ESTADO": "ESTADO",
        "R34_TERMINAL": "TERMINAL",
        "R34_SERIAL": "SERIAL",
        "R34_AFIPOS": "AFIPOS",
    }
    for salida, clave in opcionales.items():
        original = c.get(clave)
        out[salida] = f[original].astype(str).str.strip() if original else ""

    out["R34_PERTENENCIA"] = f[c["PERTENENCIA"]].astype(str).str.strip()
    out["R34_MANEJADOR"] = f[c["MANEJADOR"]].astype(str).str.strip()
    out["R34_MANEJADOR2"] = manejador2.loc[f.index].astype(str).str.strip()
    out["R34_NOMBRE_BANCO"] = f[c["NOMBRE_BANCO"]].astype(str).str.strip()
    out["R34_POS_CON_TRANSACCION"] = f[c["POS_CON_TRANSACCION"]].astype(str).str.strip()

    out = out[out["CLAVE_CRUCE"] != ""].reset_index(drop=True)
    return out


def detectar_hoja_y_header_excel(path):
    """
    Encuentra automáticamente la hoja correcta del R34 y la fila de encabezados.
    Así no importa que exista una 'Hoja1' vacía ni que la hoja útil tenga otro nombre.
    """
    requeridas = {
        "PERTENENCIA",
        "MANEJADOR",
        "POSCONTRANSACCION",
        "NOMBREBANCO",
    }
    llaves_posibles = {"CONCATENAR", "CODIGOAFIL", "AFILIADO", "NUMPOS"}

    wb = load_workbook(path, read_only=True, data_only=True)
    mejor = None
    mejor_score = -1

    try:
        for ws in wb.worksheets:
            for numero_fila, fila in enumerate(
                ws.iter_rows(min_row=1, max_row=30, values_only=True), start=1
            ):
                valores = {
                    normalizar_compacto(v)
                    for v in fila
                    if v is not None and str(v).strip() != ""
                }

                score = sum(1 for req in requeridas if req in valores)
                tiene_llave = bool(valores & llaves_posibles)

                if score >= 4 and tiene_llave and score > mejor_score:
                    mejor = (ws.title, numero_fila - 1)  # header de pandas es base 0
                    mejor_score = score
    finally:
        wb.close()

    if mejor is None:
        raise ValueError(
            "No pude localizar automáticamente la hoja/encabezado del R34. "
            "Busqué PERTENENCIA, MANEJADOR, POS_CON_TRANSACCION y NOMBRE_BANCO "
            "en las primeras 30 filas de todas las hojas."
        )

    return mejor


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
    validar_llave_r34(columnas)

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
    hoja, header_row = detectar_hoja_y_header_excel(path)

    df = pd.read_excel(
        path,
        sheet_name=hoja,
        header=header_row,
        dtype=str,
        keep_default_na=False,
        engine="openpyxl",
    )
    df.columns = [str(c).strip() for c in df.columns]

    columnas = resolver_columnas_r34(df.columns)
    validar_llave_r34(columnas)

    return filtrar_chunk_r34(
        df, columnas, mapa_manejadores, banco_objetivo, nombre_origen
    )


def procesar_r34_subarchivo(path, mapa_manejadores, banco_objetivo, nombre_origen):
    ext = Path(path).suffix.lower()

    if ext in {".csv", ".txt"}:
        return procesar_csv_r34(
            path, mapa_manejadores, banco_objetivo, nombre_origen
        )

    if ext in {".xlsx", ".xlsm"}:
        return procesar_excel_r34(
            path, mapa_manejadores, banco_objetivo, nombre_origen
        )

    raise ValueError(f"Formato R34 no soportado: {ext}")


def procesar_upload_r34(upload, mapa_manejadores, banco_objetivo):
    resultados = []

    with tempfile.TemporaryDirectory() as td:
        ruta = os.path.join(td, Path(upload.name).name)
        upload.seek(0)
        with open(ruta, "wb") as f:
            shutil.copyfileobj(upload, f, length=8 * 1024 * 1024)

        ext = Path(ruta).suffix.lower()

        if ext == ".zip":
            with zipfile.ZipFile(ruta) as z:
                miembros = [
                    m for m in z.namelist()
                    if Path(m).suffix.lower() in {".csv", ".txt", ".xlsx", ".xlsm"}
                    and not Path(m).name.startswith("~$")
                    and not m.startswith("__MACOSX/")
                ]

                if not miembros:
                    raise ValueError(f"{upload.name}: el ZIP no contiene CSV/XLSX válidos.")

                for i, miembro in enumerate(miembros, start=1):
                    destino = os.path.join(td, f"extraido_{i}{Path(miembro).suffix.lower()}")
                    with z.open(miembro) as src, open(destino, "wb") as dst:
                        shutil.copyfileobj(src, dst, length=8 * 1024 * 1024)

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
# CRUCE: LA BASE DE SALIDA ES BNC
# ============================================================

def preparar_r34_para_cruce(r34_filtrado):
    """
    Deja una fila representativa por CLAVE_CRUCE para que un duplicado del R34
    no multiplique las filas de BNC al hacer el merge.
    """
    conteos = (
        r34_filtrado.groupby("CLAVE_CRUCE", as_index=False)
        .size()
        .rename(columns={"size": "COINCIDENCIAS_R34"})
    )

    detalle = r34_filtrado.drop_duplicates("CLAVE_CRUCE", keep="first").copy()
    return detalle.merge(conteos, on="CLAVE_CRUCE", how="left")


def cruzar_bnc_con_r34(bnc, r34_filtrado):
    """
    Última regla indicada:
    NO ENCONTRADOS = registros que ESTÁN EN BNC pero cuya CLAVE_CRUCE NO existe en R34.
    CASOS CRUZADOS = registros de BNC cuya CLAVE_CRUCE sí existe en R34.
    """
    r34_unico = preparar_r34_para_cruce(r34_filtrado)
    resultado = bnc.merge(r34_unico, on="CLAVE_CRUCE", how="left")

    resultado["ESTADO_CRUCE"] = resultado["COINCIDENCIAS_R34"].apply(
        lambda x: "CRUZADO" if pd.notna(x) else "NO ENCONTRADO EN R34"
    )
    resultado["COINCIDENCIAS_R34"] = (
        resultado["COINCIDENCIAS_R34"].fillna(0).astype(int)
    )

    primeras = [
        "ESTADO_CRUCE",
        "CLAVE_CRUCE",
        "BNC_FECHA_DOC",
        "BNC_CLIENTE",
        "BNC_NOMBRE",
        "BNC_CONCATENAR",
        "BNC_AFILIADO",
        "BNC_TERMINAL",
        "BNC_PAGO_CCR_USD",
        "BNC_PAGO_BAN_USD",
        "BNC_PAGO_AA_USD",
        "BNC_MANEJADOR_R34",
        "BNC_MANEJADOR_R34_MODIFICADO",
        "MANEJADOR2_BNC",
        "BNC_EQUIPO_R34",
        "COINCIDENCIAS_R34",
        "R34_CODIGO_AFIL",
        "R34_NUMPOS",
        "R34_NOMBRE_AFILIADO",
        "R34_RIF_AFILIADO",
        "R34_PERTENENCIA",
        "R34_MANEJADOR",
        "R34_MANEJADOR2",
        "R34_NOMBRE_BANCO",
        "R34_POS_CON_TRANSACCION",
        "R34_TERMINAL",
        "R34_SERIAL",
        "R34_CIUDAD",
        "R34_ESTADO",
        "ARCHIVO_R34",
    ]
    primeras = [c for c in primeras if c in resultado.columns]
    resto = [c for c in resultado.columns if c not in primeras]
    return resultado[primeras + resto]


# ============================================================
# RESUMEN Y EXCEL FINAL
# ============================================================

def crear_resumen(resultado):
    filas = []

    for manejador, g in resultado.groupby("MANEJADOR2_BNC", dropna=False):
        claves = g["CLAVE_CRUCE"].nunique()
        cruzadas = g.loc[g["ESTADO_CRUCE"] == "CRUZADO", "CLAVE_CRUCE"].nunique()
        no_encontradas = g.loc[
            g["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34", "CLAVE_CRUCE"
        ].nunique()

        filas.append({
            "MANEJADOR2": manejador if str(manejador).strip() else "SIN MANEJADOR",
            "FILAS_BNC": len(g),
            "CONCATENAR_UNICOS_BNC": claves,
            "FILAS_CRUZADAS": int((g["ESTADO_CRUCE"] == "CRUZADO").sum()),
            "FILAS_NO_ENCONTRADAS": int((g["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34").sum()),
            "CONCATENAR_UNICOS_CRUZADOS": cruzadas,
            "CONCATENAR_UNICOS_NO_ENCONTRADOS": no_encontradas,
            "%_CRUCE_UNICO": cruzadas / claves if claves else 0,
        })

    resumen = pd.DataFrame(filas)
    if resumen.empty:
        return resumen

    resumen = resumen.sort_values(
        ["CONCATENAR_UNICOS_NO_ENCONTRADOS", "CONCATENAR_UNICOS_BNC"],
        ascending=[False, False],
    ).reset_index(drop=True)

    claves_total = resultado["CLAVE_CRUCE"].nunique()
    claves_cruzadas = resultado.loc[
        resultado["ESTADO_CRUCE"] == "CRUZADO", "CLAVE_CRUCE"
    ].nunique()
    claves_no = resultado.loc[
        resultado["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34", "CLAVE_CRUCE"
    ].nunique()

    total = {
        "MANEJADOR2": "TOTAL",
        "FILAS_BNC": len(resultado),
        "CONCATENAR_UNICOS_BNC": claves_total,
        "FILAS_CRUZADAS": int((resultado["ESTADO_CRUCE"] == "CRUZADO").sum()),
        "FILAS_NO_ENCONTRADAS": int((resultado["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34").sum()),
        "CONCATENAR_UNICOS_CRUZADOS": claves_cruzadas,
        "CONCATENAR_UNICOS_NO_ENCONTRADOS": claves_no,
        "%_CRUCE_UNICO": claves_cruzadas / claves_total if claves_total else 0,
    }

    return pd.concat([resumen, pd.DataFrame([total])], ignore_index=True)


def ajustar_anchos(worksheet, df, inicio_col=0, max_width=45):
    """Ajusta columnas sin fallar con floats, fechas o NaN."""
    for i, col in enumerate(df.columns):
        if df.empty:
            ancho_datos = 0
        else:
            valores = df[col].head(300).fillna("").astype(str)
            ancho_datos = int(valores.str.len().max()) if not valores.empty else 0

        ancho = max(len(str(col)), ancho_datos) + 2
        worksheet.set_column(
            inicio_col + i,
            inicio_col + i,
            min(max(ancho, 10), max_width),
        )


def crear_excel(resultado, resumen, hoja_bnc, banco_objetivo, r34_filtrado):
    no_encontrados = resultado[
        resultado["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34"
    ].copy()
    cruzados = resultado[resultado["ESTADO_CRUCE"] == "CRUZADO"].copy()

    salida = io.BytesIO()

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_titulo = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "align": "center",
            "valign": "vcenter",
        })
        fmt_header = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#4472C4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        fmt_label = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
        })
        fmt_valor = workbook.add_format({"border": 1})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_rojo = workbook.add_format({"bg_color": "#FCE4D6"})
        fmt_verde = workbook.add_format({"bg_color": "#E2F0D9"})

        # ---------------- RESUMEN ----------------
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False, startrow=12)
        ws = writer.sheets["RESUMEN"]
        ws.merge_range("A1:H1", "VALIDACIÓN BNC — R34 VS RECAUDACIÓN", fmt_titulo)
        ws.set_row(0, 28)

        claves_total = resultado["CLAVE_CRUCE"].nunique()
        claves_cruzadas = resultado.loc[
            resultado["ESTADO_CRUCE"] == "CRUZADO", "CLAVE_CRUCE"
        ].nunique()
        claves_no = resultado.loc[
            resultado["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34", "CLAVE_CRUCE"
        ].nunique()

        metricas = [
            ("Hoja BNC analizada", hoja_bnc),
            ("Banco filtrado en R34", banco_objetivo),
            ("Registros R34 después de filtros", len(r34_filtrado)),
            ("Concatenar únicos R34", r34_filtrado["CLAVE_CRUCE"].nunique()),
            ("Filas BNC analizadas", len(resultado)),
            ("Concatenar únicos BNC", claves_total),
            ("Concatenar únicos cruzados", claves_cruzadas),
            ("Concatenar únicos NO encontrados en R34", claves_no),
            ("% cruce de Concatenar únicos", claves_cruzadas / claves_total if claves_total else 0),
        ]

        for fila, (label, valor) in enumerate(metricas, start=2):
            ws.write(fila, 0, label, fmt_label)
            if label.startswith("%"):
                ws.write(fila, 1, valor, fmt_pct)
            else:
                ws.write(fila, 1, valor, fmt_valor)

        ws.write(11, 0, "Resumen por MANEJADOR2", fmt_label)
        for col_idx, col in enumerate(resumen.columns):
            ws.write(12, col_idx, col, fmt_header)

        if not resumen.empty:
            ws.autofilter(12, 0, 12 + len(resumen), len(resumen.columns) - 1)
        ws.freeze_panes(13, 0)
        ajustar_anchos(ws, resumen)
        ws.set_column(0, 0, 30)
        if "%_CRUCE_UNICO" in resumen.columns:
            idx_pct = resumen.columns.get_loc("%_CRUCE_UNICO")
            ws.set_column(idx_pct, idx_pct, 18, fmt_pct)

        # ---------------- DETALLES ----------------
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
                wsd.set_column(estado_idx, estado_idx, 23, color_fmt)

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
    "3. R34 (CSV, XLSX o ZIP; también puedes subir varios)",
    type=["csv", "txt", "xlsx", "xlsm", "zip"],
    accept_multiple_files=True,
    key="r34",
)

banco_objetivo = st.text_input(
    "Banco que debe aparecer en NOMBRE_BANCO del R34",
    value="B.O.D",
    help="B.O.D, BOD y variantes con puntos se comparan de forma normalizada.",
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
            st.info(
                "No se filtra el R34 por FECHA_INSTALACION. Se usan todos los registros "
                "del R34 que cumplan PERTENENCIA, MANEJADOR, POS_CON_TRANSACCION y NOMBRE_BANCO."
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
                bnc = cargar_recaudacion(
                    archivo_recaudacion.getvalue(), hoja_bnc, mapa
                )

                partes = []
                for i, archivo in enumerate(archivos_r34, start=1):
                    st.write(f"Procesando R34 {i}/{len(archivos_r34)}: {archivo.name}")
                    parte = procesar_upload_r34(
                        archivo, mapa, banco_objetivo
                    )
                    if not parte.empty:
                        partes.append(parte)

                if not partes:
                    raise ValueError(
                        "Después de aplicar los filtros no quedó ningún registro del R34. "
                        "Revisa PERTENENCIA, MANEJADOR, POS_CON_TRANSACCION y NOMBRE_BANCO."
                    )

                r34_filtrado = pd.concat(partes, ignore_index=True)
                resultado = cruzar_bnc_con_r34(bnc, r34_filtrado)
                resumen = crear_resumen(resultado)
                excel = crear_excel(
                    resultado, resumen, hoja_bnc, banco_objetivo, r34_filtrado
                )

                periodo_archivo = re.sub(
                    r"[^A-Za-z0-9_-]+", "_", hoja_bnc
                ).strip("_")
                nombre_salida = f"Validacion_BNC_{periodo_archivo}.xlsx"

                claves_bnc = resultado["CLAVE_CRUCE"].nunique()
                claves_cruzadas = resultado.loc[
                    resultado["ESTADO_CRUCE"] == "CRUZADO", "CLAVE_CRUCE"
                ].nunique()
                claves_no = resultado.loc[
                    resultado["ESTADO_CRUCE"] == "NO ENCONTRADO EN R34", "CLAVE_CRUCE"
                ].nunique()

                st.session_state["resultado_excel"] = excel
                st.session_state["nombre_salida"] = nombre_salida
                st.session_state["metricas_resultado"] = {
                    "r34_validos": len(r34_filtrado),
                    "bnc_unicos": claves_bnc,
                    "cruzados_unicos": claves_cruzadas,
                    "no_encontrados_unicos": claves_no,
                }

            st.success("Validación terminada.")

        except Exception as e:
            st.exception(e)


if st.session_state.get("resultado_excel"):
    m = st.session_state["metricas_resultado"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros R34 válidos", f"{m['r34_validos']:,}")
    c2.metric("Concatenar únicos BNC", f"{m['bnc_unicos']:,}")
    c3.metric("Cruzados", f"{m['cruzados_unicos']:,}")
    c4.metric("BNC no encontrados en R34", f"{m['no_encontrados_unicos']:,}")

    st.download_button(
        "Descargar Excel de validación",
        data=st.session_state["resultado_excel"],
        file_name=st.session_state["nombre_salida"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
