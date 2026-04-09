from config.configuracion import MES_TRABAJO, MENSAJE_INICIO, MENSAJE_CONFIRMACION
from utils.mensajes import confirmar_inicio, hablar, mostrar_info, mostrar_error

from procesadores.procesador import (
    procesar_alcon,
    procesar_certificacion_gerentes,
    procesar_temporales_td_saldo,
    procesar_temporales_td_sabana
)

# ==============================
# FUNCIÓN PRINCIPAL DEL SISTEMA
# ==============================

def ejecutar_indicadores():
    """
    Ejecuta el flujo principal del sistema ISCE.
    """

    mensaje_inicio = f"{MENSAJE_CONFIRMACION}{MES_TRABAJO}"

    confirmar = confirmar_inicio(mensaje_inicio)

    if not confirmar:
        return

    hablar(MENSAJE_INICIO)

    try:
        procesar_alcon()
        procesar_certificacion_gerentes()
        procesar_temporales_td_saldo()
        procesar_temporales_td_sabana()

        hablar("Ejecución finalizada")

        mostrar_info(
            "Proceso completado",
            f"Los indicadores del mes {MES_TRABAJO} fueron procesados correctamente."
        )

    except Exception as e:
        hablar("La ejecución presentó errores")

        mostrar_error(
            "Error en la ejecución",
            f"Ocurrió un error durante el procesamiento:\n\n{str(e)}"
        )


# ==============================
# EJECUCIÓN DEL PROGRAMA
# ==============================

if __name__ == "__main__":
    ejecutar_indicadores()