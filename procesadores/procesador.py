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
    Procesa la hoja Sábana Temporales y construye:

    - Débitos (positivos) → H4:J
    - Créditos (negativos) → K4:M

    Alineado exactamente con la columna A del Excel de salida.
    """

    from openpyxl import load_workbook
    from config.rutas import obtener_ruta_salida

    # ==============================
    # CARGAR DATOS DE SÁBANA
    # ==============================
    df = cargar_tabla_por_coincidencia_hoja(
        parte_nombre_archivo=NOMBRE_BASE_TEMPORALES,
        texto_hoja=HOJA_SABANA_REFERENCIA,
        columnas_esperadas=COLUMNAS_SABANA
    )

    df["gerencia_responsable"] = df["gerencia_responsable"].astype(str).str.strip()

    df["FUERA DE POLITICA"] = (
        df["FUERA DE POLITICA"]
        .astype(str)
        .str.strip()
        .str.upper()
        .replace("SÍ", "SI")
    )

    df["VALOR PARTIDA PESOS"] = pd.to_numeric(
        df["VALOR PARTIDA PESOS"], errors="coerce"
    ).fillna(0)

    # ==============================
    # LEER GERENCIAS DESDE COLUMNA A
    # ==============================
    ruta_salida = obtener_ruta_salida(NOMBRE_ARCHIVO_SALIDA)
    wb = load_workbook(ruta_salida)
    ws = wb[MES_TRABAJO]

    gerencias = []
    fila = 5

    while True:
        valor = ws.cell(row=fila, column=1).value

        if valor is None:
            break

        valor = str(valor).strip()
        gerencias.append(valor)

        if valor.upper() == "TOTAL GENERAL":
            break

        fila += 1

    # Base SIN total general para alinear cálculo
    gerencias_sin_total = [g for g in gerencias if g.upper() != "TOTAL GENERAL"]
    df_base = pd.DataFrame({"gerencia_responsable": gerencias_sin_total})

    # ==============================
    # DÉBITOS (POSITIVOS)
    # ==============================
    df_db = df[df["VALOR PARTIDA PESOS"] > 0].copy()

    df_db["TOTAL_DB"] = df_db["VALOR PARTIDA PESOS"]
    df_db["DB_FUERA"] = df_db.apply(
        lambda x: x["VALOR PARTIDA PESOS"] if x["FUERA DE POLITICA"] == "SI" else 0,
        axis=1
    )

    resumen_db = df_db.groupby("gerencia_responsable", as_index=False).agg({
        "TOTAL_DB": "sum",
        "DB_FUERA": "sum"
    })

    resumen_db = df_base.merge(
        resumen_db,
        on="gerencia_responsable",
        how="left"
    ).fillna(0)

    resumen_db["%"] = resumen_db.apply(
        lambda x: x["DB_FUERA"] / x["TOTAL_DB"] if x["TOTAL_DB"] != 0 else 0,
        axis=1
    )

    resumen_db = resumen_db.rename(columns={
        "TOTAL_DB": "TOTAL VALOR PARTIDAS PESOS DB",
        "DB_FUERA": "VALOR PARTIDAS PESOS DB (FUERA POLITICA)"
    })

    # Agregar TOTAL GENERAL
    total_db = resumen_db["TOTAL VALOR PARTIDAS PESOS DB"].sum()
    total_db_fuera = resumen_db["VALOR PARTIDAS PESOS DB (FUERA POLITICA)"].sum()
    total_db_pct = total_db_fuera / total_db if total_db != 0 else 0

    fila_total_db = pd.DataFrame([{
        "gerencia_responsable": "Total general",
        "TOTAL VALOR PARTIDAS PESOS DB": total_db,
        "VALOR PARTIDAS PESOS DB (FUERA POLITICA)": total_db_fuera,
        "%": total_db_pct
    }])

    resumen_db = pd.concat([resumen_db, fila_total_db], ignore_index=True)

    # ==============================
    # CRÉDITOS (NEGATIVOS)
    # ==============================
    df_cr = df[df["VALOR PARTIDA PESOS"] < 0].copy()

    # Para reporte visual, se muestran en positivo
    df_cr["VALOR_ABS"] = df_cr["VALOR PARTIDA PESOS"].abs()

    df_cr["TOTAL_CR"] = df_cr["VALOR_ABS"]
    df_cr["CR_FUERA"] = df_cr.apply(
        lambda x: x["VALOR_ABS"] if x["FUERA DE POLITICA"] == "SI" else 0,
        axis=1
    )

    resumen_cr = df_cr.groupby("gerencia_responsable", as_index=False).agg({
        "TOTAL_CR": "sum",
        "CR_FUERA": "sum"
    })

    resumen_cr = df_base.merge(
        resumen_cr,
        on="gerencia_responsable",
        how="left"
    ).fillna(0)

    resumen_cr["%"] = resumen_cr.apply(
        lambda x: x["CR_FUERA"] / x["TOTAL_CR"] if x["TOTAL_CR"] != 0 else 0,
        axis=1
    )

    resumen_cr = resumen_cr.rename(columns={
        "TOTAL_CR": "TOTAL VALOR PARTIDAS PESOS CR",
        "CR_FUERA": "VALOR PARTIDAS PESOS CR (FUERA POLITICA)"
    })

    # Agregar TOTAL GENERAL
    total_cr = resumen_cr["TOTAL VALOR PARTIDAS PESOS CR"].sum()
    total_cr_fuera = resumen_cr["VALOR PARTIDAS PESOS CR (FUERA POLITICA)"].sum()
    total_cr_pct = total_cr_fuera / total_cr if total_cr != 0 else 0

    fila_total_cr = pd.DataFrame([{
        "gerencia_responsable": "Total general",
        "TOTAL VALOR PARTIDAS PESOS CR": total_cr,
        "VALOR PARTIDAS PESOS CR (FUERA POLITICA)": total_cr_fuera,
        "%": total_cr_pct
    }])

    resumen_cr = pd.concat([resumen_cr, fila_total_cr], ignore_index=True)

    # ==============================
    # EXPORTAR SIN COLUMNA GERENCIA
    # ==============================
    resumen_db_export = resumen_db[[
        "TOTAL VALOR PARTIDAS PESOS DB",
        "VALOR PARTIDAS PESOS DB (FUERA POLITICA)",
        "%"
    ]].copy()

    resumen_cr_export = resumen_cr[[
        "TOTAL VALOR PARTIDAS PESOS CR",
        "VALOR PARTIDAS PESOS CR (FUERA POLITICA)",
        "%"
    ]].copy()

    escribir_dataframe_en_excel(
        df=resumen_db_export,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio="H4",
        columna_porcentaje=2,
        formato_porcentaje='0.0%'
    )

    escribir_dataframe_en_excel(
        df=resumen_cr_export,
        nombre_archivo=NOMBRE_ARCHIVO_SALIDA,
        nombre_hoja=MES_TRABAJO,
        celda_inicio="K4",
        columna_porcentaje=2,
        formato_porcentaje='0.0%'
    )

        # ==============================
    # FORMATO NUMÉRICO Y LIMPIEZA
    # ==============================
    wb = load_workbook(ruta_salida)
    ws = wb[MES_TRABAJO]

    ultima_fila = 4 + len(resumen_db_export)

    # Formato numérico para valores
    for fila_excel in range(5, ultima_fila + 1):
        for col in ["H", "I", "K", "L"]:
            ws[f"{col}{fila_excel}"].number_format = '#,##0'

    # Formato porcentaje para J y M
    for fila_excel in range(5, ultima_fila + 1):
        ws[f"J{fila_excel}"].number_format = '0.0%'
        ws[f"M{fila_excel}"].number_format = '0.0%'

    # ==============================
    # LIMPIAR COLUMNA N (sobrante)
    # ==============================
    for fila_excel in range(4, 1000):
        ws[f"N{fila_excel}"].value = None
        ws[f"N{fila_excel}"].number_format = 'General'

    wb.save(ruta_salida)

    return {
        "debitos": resumen_db,
        "creditos": resumen_cr
    }

   