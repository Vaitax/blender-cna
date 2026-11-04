"""Guarda la clave de SAP en el Administrador de credenciales de Windows (no en archivos)."""
import getpass
import keyring
from blender_local import leer_config

cfg = leer_config()
srv, usr = cfg["servicio"], cfg["usuario"]
if keyring.get_password(srv, usr):
    r = input(f"Ya hay una clave guardada para {usr} (servicio {srv}). ¿Reemplazarla? (s/n): ").strip().lower()
    if r != "s":
        print("No se cambió nada.")
        raise SystemExit
clave = getpass.getpass(f"Clave de SAP para {usr} (no se muestra al escribir): ")
if clave:
    keyring.set_password(srv, usr, clave)
    print("Clave guardada en el Administrador de credenciales de Windows.")
else:
    print("No se ingresó clave; no se cambió nada.")
