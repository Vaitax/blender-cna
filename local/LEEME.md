# Blender CNA local (conectado a SAP)

Abre el Blender en tu navegador y lee los datos por sí solo:
- **Insumos por producto**: directo de SAP Business One (listas de materiales MGM), solo lectura.
- **Precios, materias primas y aportes**: del Excel del Blender en tu OneDrive sincronizado.

Nada sale de tu PC. La app solo atiende en `localhost` y la clave de SAP queda en el
Administrador de credenciales de Windows.

## Instalación (una sola vez)
1. Descarga el repositorio desde GitHub con el botón verde **Code › Download ZIP** y descomprímelo,
   por ejemplo en `C:\BlenderCNA`.
2. Necesitas Python. Si no lo tienes, instálalo desde python.org (no pide administrador) y
   marca "Add python.exe to PATH".
3. Si en tu PC **no** está guardada la clave de la app del bono, abre `local\Guardar clave SAP.bat`.

## Uso
Abre `local\Abrir Blender.bat`. La primera vez instala lo que falta y después abre el Blender.
Deja la ventana negra abierta mientras lo usas.

- Arriba se indica de dónde vienen los datos y a qué hora se leyeron. **Sincronizar** los vuelve a leer.
- El Excel se vuelve a leer solo cuando cambia. SAP se lee cada 30 minutos o al pulsar Sincronizar
  (como máximo una vez cada 2 minutos).
- Si SAP rechaza la clave, la app **no lo vuelve a intentar** hasta que la cierres, para no
  bloquear la cuenta compartida POWERBI. Mientras tanto usa la hoja de insumos del Excel.

## Configuración
En `local\config.ini` se configuran la ruta del Excel (si la app no lo encuentra sola), la base de SAP, el usuario,
cada cuántos minutos se lee SAP y el puerto. La consulta está en `local\consultas\insumos.sql`.
