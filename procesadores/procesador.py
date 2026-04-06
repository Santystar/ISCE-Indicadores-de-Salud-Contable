import pandas as pd

from config.configuracion import (
    NOMBRE_BASE_ALCON,
    NOMBRE_BASE_CERTIFICACION,
    NOMBRE_BASE_TEMPORALES,
    HOJA_ALCON,
    HOJA_TEMPORAL_REFERENCIA,
    HOJA_SABANA_REFERENCIA,
    COLUMNAS_ALCON,
    COLUMNAS_CERTIFICACION,
    COLUMNAS_TEMPORAL,
    COLUMNAS_SABANA,
    NOMBRE_ARCHIVO_SALIDA,
    MES_TRABAJO,
    CELDA_INICIO_ALCON,
    CELDA_INICIO_CERTIFICACION,
    CELDA_INICIO_TD_SALDO,
    CELDA_INICIO_TD_SABANA
)

from cargadores.cargador_excel import (
    cargar_tabla_excel,
    cargar_tabla_por_coincidencia_hoja
)

from exportadores.exportador_excel import escribir_dataframe_en_excel
from config.rutas import obtener_archivo_por_coincidencia


# ==============================
# UTILIDADES
# ==============================

def limpiar_gerencias_invalidas(df, columna="Gerencia"):
    """
    Elimina filas donde la gerencia esté vacía, en blanco o inválida.
    """
    df[columna] = df[columna].astype(str).str.strip()

    df = df[
        (df[columna] != "") &
        (df[columna].str.lower() != "nan") &
        (df[columna].str.lower() != "(en blanco)")
    ]

    return df


def convertir_porcentaje(valor):
    """
    Convierte un valor a número decimal para formato porcentaje en Excel.
    """
    try:
        return float(valor)
    except:
        return 0


# ==============================
# ALCON
# ==============================

def cargar_alcon_con_encabezado_dinamico():
    """
    Busca automáticamente la fila real de encabezado del archivo ALCON.
    """

    ruta_archivo = obtener_archivo_por_coincidencia(NOMBRE_BASE_ALCON)

    df_crudo = pd.read_excel(ruta_archivo, sheet_name=HOJA_ALCON, header=None)

    fila_encabezado = None

    for i in range(len(df_crudo)):
        fila = df_crudo.iloc[i].astype(str).tolist()

        if "Gerencia" in fila and "Cantidad alertas" in fila:
            fila_encabezado = i
            break

    if fila_encabezado is None:
        raise Exception("No se encontró encabezado en ALCON")

    df = pd.read_excel(ruta_archivo, sheet_name=HOJA_ALCON, header=fila_encabezado)

    # Limpiar nombres de columnas
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
    )

    # Buscar columnas reales por coincidencia flexible
    columnas_limpias = {col.lower(): col for col in df.columns}
    columnas_reales = []

    for col in COLUMNAS_ALCON:
        col_limpia = col.lower().strip()

        if col_limpia in columnas_limpias:
            columnas_reales.append(columnas_limpias[col_limpia])
        else:
            raise Exception(f"No se encontró la columna: {col}")

    df = df[columnas_reales].copy()
    df.columns = COLUMNAS_ALCON  # Renombrar para que quede uniforme

    return df


def procesar_alcon():
    """
    Procesa archivo ALCON y lo exporta a la hoja del mes.
    """

    df_alcon = cargar_alcon_con_encabezado_dinamico()

    df_alcon = limpiar_gerencias_invalidas(df_alcon, "Gerencia")

    df_alcon["Calidad Gerencia"] = pd.to_numeric(
        df_alcon["Calidad Gerencia"], errors="coerce"
    ).fillna(0)

    escribir_dataframe_en_excel(
        df=df_alcon,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio=CELDA_INICIO_ALCON,
        columna_porcentaje=4,
        formato_porcentaje='0.0%'
    )

    return df_alcon


# ==============================
# CERTIFICACIÓN GERENTES
# ==============================

def procesar_certificacion_gerentes():
    """
    Procesa el archivo Histórico Indicador Certificación Gerentes
    y lo exporta al archivo de salida.
    """

    df = cargar_tabla_excel(
        parte_nombre_archivo=NOMBRE_BASE_CERTIFICACION,
        columnas=COLUMNAS_CERTIFICACION
    )

    df["GERENCIA"] = df["GERENCIA"].astype(str).str.strip()

    df = df[
        (df["GERENCIA"] != "") &
        (df["GERENCIA"].str.lower() != "nan") &
        (~df["GERENCIA"].str.startswith("*"))
    ]

    df["FECHA CERTIFICACIÓN"] = pd.to_datetime(
        df["FECHA CERTIFICACIÓN"], errors="coerce", dayfirst=True
    )

    df["FECHA OBJETIVO"] = pd.to_datetime(
        df["FECHA OBJETIVO"], errors="coerce", dayfirst=True
    )

    df["INDICADOR"] = df["INDICADOR"].apply(convertir_porcentaje)

    escribir_dataframe_en_excel(
        df=df,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio=CELDA_INICIO_CERTIFICACION,
        columna_porcentaje=3,
        calcular_promedio=True,
        formato_porcentaje='0.00%',
        columnas_fecha=[1, 2]
    )

    return df


# ==============================
# TEMPORALES - TD SALDO
# ==============================

def procesar_temporales_td_saldo():
    """
    Procesa la hoja TEMPORAL y construye la tabla TD SALDO.
    """

    df = cargar_tabla_por_coincidencia_hoja(
        parte_nombre_archivo=NOMBRE_BASE_TEMPORALES,
        texto_hoja=HOJA_TEMPORAL_REFERENCIA,
        columnas_esperadas=COLUMNAS_TEMPORAL
    )

    df["gerencia_responsable"] = df["gerencia_responsable"].astype(str).str.strip()

    df = df[
        (df["gerencia_responsable"] != "") &
        (df["gerencia_responsable"].str.lower() != "nan")
    ]

    df["SALDO CONTABLE"] = pd.to_numeric(df["SALDO CONTABLE"], errors="coerce").fillna(0)
    df["PARTIDAS FUERA DE POLITICA_y"] = pd.to_numeric(
        df["PARTIDAS FUERA DE POLITICA_y"], errors="coerce"
    ).fillna(0)

    # Eliminar filas con saldo 0
    df = df[df["SALDO CONTABLE"] != 0]

    # Contar cuentas con saldo
    df["total"] = 1

    # Contar cuántas cuentas tienen fuera de política > 0
    df["fuera"] = (df["PARTIDAS FUERA DE POLITICA_y"] > 0).astype(int)

    resumen = df.groupby("gerencia_responsable", as_index=False).agg({
        "total": "sum",
        "fuera": "sum"
    })

    resumen["%"] = resumen["fuera"] / resumen["total"]

    resumen = resumen.rename(columns={
        "gerencia_responsable": "Area",
        "total": "TOTAL CUENTAS TEMPORALES CON SALDO",
        "fuera": "CUENTAS TEMPORALES FUERA DE POLITICA"
    })

    # ==============================
    # TOTAL GENERAL
    # ==============================
    total_b = resumen["TOTAL CUENTAS TEMPORALES CON SALDO"].sum()
    total_c = resumen["CUENTAS TEMPORALES FUERA DE POLITICA"].sum()
    total_d = total_c / total_b if total_b != 0 else 0

    fila_total = pd.DataFrame([{
        "Area": "Total general",
        "TOTAL CUENTAS TEMPORALES CON SALDO": total_b,
        "CUENTAS TEMPORALES FUERA DE POLITICA": total_c,
        "%": total_d
    }])

    resumen = pd.concat([resumen, fila_total], ignore_index=True)

    escribir_dataframe_en_excel(
        df=resumen,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio=CELDA_INICIO_TD_SALDO,
        columna_porcentaje=3,
        formato_porcentaje='0.0%'
    )

    return resumen


# ==============================
# TEMPORALES - TD SÁBANA
# ==============================

def procesar_temporales_td_sabana():
    """
    Procesa la hoja Sábana Temporales y construye la tabla TD SÁBANA.
    """

    df = cargar_tabla_por_coincidencia_hoja(
        parte_nombre_archivo=NOMBRE_BASE_TEMPORALES,
        texto_hoja=HOJA_SABANA_REFERENCIA,
        columnas_esperadas=COLUMNAS_SABANA
    )

    df["gerencia_responsable"] = df["gerencia_responsable"].astype(str).str.strip()

    df = df[
        (df["gerencia_responsable"] != "") &
        (df["gerencia_responsable"].str.lower() != "nan")
    ]

    df["FUERA DE POLITICA"] = (
        df["FUERA DE POLITICA"]
        .astype(str)
        .str.strip()
        .str.upper()
        .replace("SÍ", "SI")
    )

    # Cada fila es una partida
    df["total"] = 1

    # Contar solo las que están en SI
    df["fuera"] = (df["FUERA DE POLITICA"] == "SI").astype(int)

    resumen = df.groupby("gerencia_responsable", as_index=False).agg({
        "total": "sum",
        "fuera": "sum"
    })

    resumen["%"] = resumen["fuera"] / resumen["total"]

    resumen = resumen.rename(columns={
        "total": "TOTAL PARTIDAS",
        "fuera": "PARTIDAS FUERA DE POLITICA"
    })

    resumen = resumen[[
        "TOTAL PARTIDAS",
        "PARTIDAS FUERA DE POLITICA",
        "%"
    ]]

    # ==============================
    # TOTAL GENERAL
    # ==============================
    total_e = resumen["TOTAL PARTIDAS"].sum()
    total_f = resumen["PARTIDAS FUERA DE POLITICA"].sum()
    total_g = total_f / total_e if total_e != 0 else 0

    fila_total = pd.DataFrame([{
        "TOTAL PARTIDAS": total_e,
        "PARTIDAS FUERA DE POLITICA": total_f,
        "%": total_g
    }])

    resumen = pd.concat([resumen, fila_total], ignore_index=True)

    escribir_dataframe_en_excel(
        df=resumen,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio=CELDA_INICIO_TD_SABANA,
        columna_porcentaje=2,
        formato_porcentaje='0.0%'
    )

    return resumen