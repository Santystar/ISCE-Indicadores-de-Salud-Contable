import os
import unicodedata
import pandas as pd

from config.rutas import obtener_archivo_por_coincidencia


# ==============================
# FUNCIONES AUXILIARES
# ==============================

def normalizar_texto(texto):
    """
    Convierte un texto a minúsculas y sin tildes para comparar nombres.
    """
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    return texto


# ==============================
# CARGA BÁSICA DE ARCHIVOS
# ==============================

def cargar_tabla_excel(parte_nombre_archivo, columnas, nombre_hoja=0):
    """
    Carga columnas específicas de un archivo Excel encontrado por coincidencia.
    """

    ruta_archivo = obtener_archivo_por_coincidencia(parte_nombre_archivo)

    df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja)

    columnas_faltantes = [col for col in columnas if col not in df.columns]

    if columnas_faltantes:
        raise Exception(
            f"Faltan columnas en el archivo '{os.path.basename(ruta_archivo)}': {columnas_faltantes}"
        )

    return df[columnas].copy()


def cargar_tabla_desde_fila_encabezado(parte_nombre_archivo, nombre_hoja, columnas_esperadas, fila_encabezado=0):
    """
    Carga una hoja Excel desde una fila específica de encabezado.
    """

    ruta_archivo = obtener_archivo_por_coincidencia(parte_nombre_archivo)

    df = pd.read_excel(
        ruta_archivo,
        sheet_name=nombre_hoja,
        header=fila_encabezado
    )

    columnas_faltantes = [col for col in columnas_esperadas if col not in df.columns]

    if columnas_faltantes:
        raise Exception(
            f"Faltan columnas en el archivo '{os.path.basename(ruta_archivo)}': {columnas_faltantes}"
        )

    return df[columnas_esperadas].copy()


# ==============================
# CARGA POR COINCIDENCIA DE HOJA
# ==============================

def cargar_tabla_por_coincidencia_hoja(parte_nombre_archivo, texto_hoja, columnas_esperadas):
    """
    Carga una hoja buscando coincidencia parcial en el nombre de la hoja.
    Ejemplo: 'temporal' encuentra 'TEMPORAL'
             'sabana' encuentra 'Sábana Temporales'
    """

    ruta_archivo = obtener_archivo_por_coincidencia(parte_nombre_archivo)

    archivo_excel = pd.ExcelFile(ruta_archivo)

    hoja_encontrada = None
    texto_hoja_normalizado = normalizar_texto(texto_hoja)

    for hoja in archivo_excel.sheet_names:
        hoja_normalizada = normalizar_texto(hoja)

        if texto_hoja_normalizado in hoja_normalizada:
            hoja_encontrada = hoja
            break

    if hoja_encontrada is None:
        raise Exception(
            f"No se encontró una hoja que coincida con '{texto_hoja}' "
            f"en el archivo '{os.path.basename(ruta_archivo)}'"
        )

    df = pd.read_excel(ruta_archivo, sheet_name=hoja_encontrada)

    columnas_faltantes = [col for col in columnas_esperadas if col not in df.columns]

    if columnas_faltantes:
        raise Exception(
            f"Faltan columnas en la hoja '{hoja_encontrada}' del archivo "
            f"'{os.path.basename(ruta_archivo)}': {columnas_faltantes}"
        )

    return df[columnas_esperadas].copy()