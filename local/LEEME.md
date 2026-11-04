# Blender CNA local (conectado a SAP)

Abre el Blender en tu navegador y lee los datos por sí solo:
- **Insumos por producto**: directo de SAP Business One (listas de materiales MGM), solo lectura.
- **Precios, materias primas y aportes**: guardados en tu PC (`maestros.json`). Los editas en
  las pestañas Lista de precios y Aportes por componente; los cambios quedan en `ajustes.json`.
  El botón **Importar del Excel** vuelve a copiar la base desde el Excel del Blender, si lo necesitas.

Nada sale de tu PC. La app solo atiende en `localhost` y la clave de SAP queda en el
Administrador de credenciales de Windows.

## Instalación (una sola vez)
1. Descarga el repositorio desde GitHub con el botón verde **Code › Download ZIP** y descomprímelo,
   por ejemplo en `C:\BlenderCNA`.
2. Necesitas Python. Si no lo tienes, instálalo desde python.org (no pide administrador) y
   marca "Add python.exe to PATH".
3. Si en tu PC **no** está guardada la clave de la app del bono, abre `local\Guardar clave SAP.bat`.

## Uso
La primera vez abre `local\Abrir Blender.bat`: instala lo que falta y crea en el escritorio el
acceso directo **Blender CNA**. Desde entonces usa ese acceso directo.

El Blender se abre en su propia ventana, sin pestañas ni barra de direcciones (usa Edge en modo
aplicación) y sin ventana negra. Al cerrar la ventana, la app se apaga sola en menos de un minuto.
Si algo falla, abre `local\Abrir Blender con consola.bat` para ver los mensajes, o revisa `local\blender.log`.

- Arriba se indica de dónde vienen los datos y a qué hora se leyeron. **Sincronizar** los vuelve a leer.
- SAP se lee al abrir, cada 30 minutos o al pulsar Sincronizar (como máximo una vez cada 2 minutos).
  La última lectura se guarda, así que si SAP no responde sigues trabajando con ella.
- Si SAP rechaza la clave, la app **no lo vuelve a intentar** hasta que la cierres, para no
  bloquear la cuenta compartida POWERBI. Mientras tanto usa la última lectura guardada.

## Configuración
En `local\config.ini` se configuran la ruta del Excel (si la app no lo encuentra sola), la base de SAP, el usuario,
cada cuántos minutos se lee SAP y el puerto. La consulta está en `local\consultas\insumos.sql`.

Los archivos `maestros.json`, `ajustes.json` e `insumos_sap.json` tienen datos de la empresa:
quedan solo en tu PC y no se suben a GitHub.
