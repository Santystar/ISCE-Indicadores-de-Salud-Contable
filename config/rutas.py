import os

# ==============================
# RUTAS PRINCIPALES DEL PROYECTO
# ==============================

RUTA_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUTA_DATOS = os.path.join(RUTA_PROYECTO, "datos")
RUTA_ENTRADA = os.path.join(RUTA_DATOS, "entrada")
RUTA_SALIDA = os.path.join(RUTA_DATOS, "salida")
RUTA_TEMPORALES = os.path.join(RUTA_DATOS, "temporales")


# ==============================
# CREAR CARPETAS SI NO EXISTEN
# ==============================

def crear_carpetas_si_no_existen():
    """
    Crea las carpetas principales del proyecto si no existen.
    """
    os.makedirs(RUTA_ENTRADA, exist_ok=True)
    os.makedirs(RUTA_SALIDA, exist_ok=True)
    os.makedirs(RUTA_TEMPORALES, exist_ok=True)


# ==============================
# BUSCAR ARCHIVOS POR COINCIDENCIA
# ==============================

def obtener_archivo_por_coincidencia(parte_nombre_archivo):
    """
    Busca un archivo dentro de la carpeta de entrada
    usando coincidencia parcial en el nombre.
    """

    crear_carpetas_si_no_existen()

    archivos = os.listdir(RUTA_ENTRADA)

    coincidencias = []

    for archivo in archivos:
        nombre_archivo = archivo.lower()
        parte_buscada = parte_nombre_archivo.lower()

        if parte_buscada in nombre_archivo and archivo.endswith((".xlsx", ".xls")):
            coincidencias.append(archivo)

    if len(coincidencias) == 0:
        raise Exception(
            f"No se encontró ningún archivo que coincida con '{parte_nombre_archivo}' "
            f"en la carpeta de entrada:\n{RUTA_ENTRADA}"
        )

    if len(coincidencias) > 1:
        raise Exception(
            f"Se encontraron varios archivos que coinciden con '{parte_nombre_archivo}':\n"
            + "\n".join(coincidencias)
        )

    return os.path.join(RUTA_ENTRADA, coincidencias[0])


# ==============================
# RUTA DE ARCHIVO DE SALIDA
# ==============================

def obtener_ruta_salida(nombre_archivo):
    """
    Retorna la ruta completa de un archivo dentro de la carpeta de salida.
    """
    crear_carpetas_si_no_existen()
    return os.path.join(RUTA_SALIDA, nombre_archivo)


def obtener_ruta_temporal(nombre_archivo):
    """
    Retorna la ruta completa de un archivo dentro de la carpeta temporales.
    """
    crear_carpetas_si_no_existen()
    return os.path.join(RUTA_TEMPORALES, nombre_archivo)