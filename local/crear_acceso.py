"""Crea en el escritorio el acceso directo "Blender CNA" (abre el Blender en su propia ventana)."""
import os
import subprocess
import sys

aqui = os.path.dirname(os.path.abspath(__file__))
pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
if not os.path.exists(pyw):
    pyw = sys.executable
icono = os.path.join(os.path.dirname(aqui), "blender.ico")
script = os.path.join(aqui, "blender_local.py")
q = lambda s: s.replace("'", "''")
ps = (
    "$d=[Environment]::GetFolderPath('Desktop');"
    "$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'Blender CNA.lnk'));"
    f"$s.TargetPath='{q(pyw)}';"
    f"$s.Arguments='\"{q(script)}\" --ventana';"
    f"$s.WorkingDirectory='{q(aqui)}';"
    f"$s.IconLocation='{q(icono)}';"
    "$s.Description='Blender CNA conectado a SAP';"
    "$s.Save()"
)
r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps])
print("Acceso directo 'Blender CNA' creado en el escritorio." if r.returncode == 0
      else "No se pudo crear el acceso directo.")
