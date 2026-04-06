import tkinter as tk
from tkinter import messagebox
import pyttsx3


# ==============================
# CONFIGURACIÓN DE VOZ
# ==============================

motor_voz = pyttsx3.init()
motor_voz.setProperty("rate", 170)  # velocidad de voz


# ==============================
# FUNCIONES DE MENSAJES
# ==============================

def hablar(texto):
    """
    Hace que el sistema hable el texto indicado.
    """
    try:
        motor_voz.say(texto)
        motor_voz.runAndWait()
    except:
        pass


def confirmar_inicio(mensaje):
    """
    Muestra una ventana de confirmación.
    Retorna True si el usuario acepta, False si cancela.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    respuesta = messagebox.askyesno("Confirmación de ejecución", mensaje)

    ventana.destroy()
    return respuesta


def mostrar_info(titulo, mensaje):
    """
    Muestra una ventana informativa.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    messagebox.showinfo(titulo, mensaje)

    ventana.destroy()


def mostrar_error(titulo, mensaje):
    """
    Muestra una ventana de error.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    messagebox.showerror(titulo, mensaje)

    ventana.destroy()