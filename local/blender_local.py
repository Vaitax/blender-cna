"""Blender CNA local.

Abre el Blender en el navegador (http://localhost:8765) y le entrega los datos:
- Insumos por producto: directo de SAP Business One (HANA) por ODBC, solo lectura.
- Precios, materias primas y aportes: guardados en este PC (maestros.json). Se toman una
  vez del Excel del Blender y después se editan en la app; los cambios quedan en ajustes.json.
Si SAP no responde, usa la última lectura guardada (insumos_sap.json) o la hoja del Excel.

Cuidado con la cuenta compartida de SAP: varios intentos fallidos la bloquean y botan
los reportes de Power BI. Por eso esta app hace UN solo intento por vez, espera entre
lecturas y, si la clave es rechazada, no vuelve a intentar hasta que la cierres.
"""
import configparser
import decimal
import glob
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

AQUI = os.path.dirname(os.path.abspath(__file__))
HOJAS = ("LISTA PRECIOS", "BLENDER", "INSUMOS POR PRODUCTO")
MIN_ENTRE_INTENTOS_SAP = 120  # segundos
MAESTROS = os.path.join(AQUI, "maestros.json")      # LISTA PRECIOS y BLENDER (copia del Excel)
AJUSTES = os.path.join(AQUI, "ajustes.json")        # precios, aportes y estados cambiados en la app
CACHE_SAP = os.path.join(AQUI, "insumos_sap.json")  # última lectura buena de SAP


def leer_json(ruta, defecto=None):
    try:
        with open(ruta, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return defecto


def guardar_json(ruta, datos):
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)
    os.replace(tmp, ruta)


LOG = os.path.join(AQUI, "blender.log")


def log(*a):
    linea = " ".join([time.strftime("%H:%M:%S")] + [str(x) for x in a])
    try:
        if sys.stdout:
            print(linea, flush=True)
    except Exception:  # noqa: BLE001  (con pythonw no hay consola)
        pass
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d ") + linea + "\n")
    except OSError:
        pass


def aviso_windows(texto):
    """Mensaje visible aunque no haya consola (modo ventana)."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, texto, "Blender CNA", 0x40)
    except Exception:  # noqa: BLE001
        log(texto)


# ---------------------------------------------------------------- configuración
def leer_config():
    cp = configparser.ConfigParser(inline_comment_prefixes=(";",))
    cp.read(os.path.join(AQUI, "config.ini"), encoding="utf-8")
    g = lambda s, k, d: (cp.get(s, k, fallback=d) or d).strip()
    return {
        "excel": g("excel", "ruta", ""),
        "sap_usar": g("sap", "usar", "si").lower() in ("si", "sí", "yes", "true", "1"),
        "dsn": g("sap", "dsn", "hana"),
        "usuario": g("sap", "usuario", "POWERBI"),
        "base": g("sap", "base", "SBO_CNA_CHILE"),
        "servicio": g("sap", "servicio_clave", "bono-hana"),
        "refrescar": max(5, int(g("sap", "refrescar_cada_min", "30") or 30)) * 60,
        "puerto": int(g("app", "puerto", "8765") or 8765),
    }


def buscar_excel(cfg):
    if cfg["excel"]:
        return cfg["excel"] if os.path.exists(cfg["excel"]) else None
    casa = os.path.expanduser("~")
    candidatos = []
    for base in glob.glob(os.path.join(casa, "OneDrive*")) + [casa]:
        for pat in ("**/BLENDER*.xlsx", "**/BLENDER*.xlsm"):
            candidatos += [p for p in glob.glob(os.path.join(base, pat), recursive=True)
                           if not os.path.basename(p).startswith("~$")]
    return max(candidatos, key=os.path.getmtime) if candidatos else None


# ---------------------------------------------------------------- Excel
def col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def limpio(v):
    if v is None or isinstance(v, (int, float, str, bool)):
        return v
    if isinstance(v, decimal.Decimal):  # HANA entrega las cantidades como Decimal
        return float(v)
    return str(v)


def num(v):
    """Cantidad como número (también arregla lecturas guardadas con la cantidad como texto)."""
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", ".")) if v.strip() else 0
        except ValueError:
            return v
    return limpio(v)


def hoja_dict(ws_v, ws_f, filas=200, cols=28):
    """Hoja como {A1: {v, f}} (el mismo formato que usa la app)."""
    out = {}
    for r, (fila_v, fila_f) in enumerate(zip(
            ws_v.iter_rows(min_row=1, max_row=filas, max_col=cols, values_only=True),
            ws_f.iter_rows(min_row=1, max_row=filas, max_col=cols, values_only=True))):
        for c, (v, f) in enumerate(zip(fila_v, fila_f)):
            formula = f[1:] if isinstance(f, str) and f.startswith("=") else None
            if v is None and not formula:
                continue
            out[f"{col(c)}{r + 1}"] = {"v": limpio(v), "f": formula}
    return out


def leer_excel(ruta):
    import openpyxl
    # copia temporal: así no choca con Excel si el libro está abierto
    tmp = os.path.join(tempfile.gettempdir(), "blender_cna_lectura.xlsx")
    shutil.copy2(ruta, tmp)
    wv = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
    wf = openpyxl.load_workbook(tmp, data_only=False, read_only=True)
    nombre = {n.strip().upper(): n for n in wv.sheetnames}
    falta = [h for h in HOJAS if h not in nombre]
    if falta:
        raise RuntimeError("al Excel le falta la hoja " + ", ".join(falta))
    lp = hoja_dict(wv[nombre["LISTA PRECIOS"]], wf[nombre["LISTA PRECIOS"]], cols=2)
    bl = hoja_dict(wv[nombre["BLENDER"]], wf[nombre["BLENDER"]])
    ins = [[limpio(v) for v in fila]
           for fila in wv[nombre["INSUMOS POR PRODUCTO"]].iter_rows(max_col=6, values_only=True)]
    wv.close()
    wf.close()
    return lp, bl, ins


# ---------------------------------------------------------------- SAP
class SAP:
    def __init__(self, cfg):
        self.cfg = cfg
        self.bloqueado = None      # motivo por el que no se vuelve a intentar
        self.ultimo_intento = 0
        self.ultimo_ok = 0
        cache = leer_json(CACHE_SAP)
        self.filas = ([cache["filas"][0]] + [r[:4] + [num(r[4])] + r[5:] for r in cache["filas"][1:]]) if cache else None
        self.ultimo_ok = cache["at"] if cache else 0
        self.de_cache = bool(cache)
        self.error = None
        self.lock = threading.Lock()

    def clave(self):
        import keyring
        return keyring.get_password(self.cfg["servicio"], self.cfg["usuario"])

    def consultar(self):
        import pyodbc
        sql = open(os.path.join(AQUI, "consultas", "insumos.sql"), encoding="utf-8").read()
        sql = sql.replace("{schema}", self.cfg["base"])
        clave = self.clave()
        if not clave:
            raise PermissionError("FALTA_CLAVE")
        cn = pyodbc.connect(f"DSN={self.cfg['dsn']};UID={self.cfg['usuario']};PWD={clave}",
                            autocommit=True, timeout=20)
        try:
            cur = cn.cursor()
            cur.execute(sql)
            filas = [[limpio(v) for v in r[:4]] + [num(r[4]), limpio(r[5])] for r in cur.fetchall()]
        finally:
            cn.close()
        cab = ["Codigos", "Name", "Código Insumo", "ItemName", "Dosificación", "Unidad Medida Insumo"]
        return [cab] + filas

    def obtener(self, forzar=False):
        """Devuelve las filas de SAP en caché, o consulta si corresponde. Nunca reintenta en bucle."""
        with self.lock:
            if not self.cfg["sap_usar"] or self.bloqueado:
                return self.filas
            ahora = time.time()
            vencido = ahora - self.ultimo_ok > self.cfg["refrescar"]
            if not (forzar or vencido or self.filas is None or self.de_cache):
                return self.filas
            if ahora - self.ultimo_intento < MIN_ENTRE_INTENTOS_SAP:
                return self.filas
            self.ultimo_intento = ahora
            try:
                log("Consultando SAP…")
                self.filas = self.consultar()
                self.ultimo_ok = time.time()
                self.de_cache = False
                self.error = None
                guardar_json(CACHE_SAP, {"at": self.ultimo_ok, "filas": self.filas})
                log(f"SAP OK: {len(self.filas) - 1} líneas de insumos")
            except PermissionError:
                self.bloqueado = ("No hay clave de SAP guardada para " + self.cfg["usuario"] +
                                  ". Ejecuta 'Guardar clave SAP.bat' y vuelve a abrir el Blender.")
                log(self.bloqueado)
            except ImportError as e:
                self.bloqueado = f"Falta instalar un componente de Python ({e.name}). Ejecuta 'Abrir Blender.bat' de nuevo."
                log(self.bloqueado)
            except Exception as e:  # noqa: BLE001
                msg = str(e)
                if any(t in msg.lower() for t in ("authentication failed", "28000", "invalid username or password",
                                                  "user is locked", "password")):
                    # clave rechazada: no volver a intentar para no bloquear la cuenta compartida
                    self.bloqueado = ("SAP rechazó la clave de " + self.cfg["usuario"] +
                                      ". No se volverá a intentar para no bloquear la cuenta. Revisa la clave con quien "
                                      "administra la app del bono y vuelve a abrir el Blender.")
                else:
                    self.error = "No se pudo leer SAP: " + msg[:300]
                log(self.bloqueado or self.error)
            return self.filas


# ---------------------------------------------------------------- comparación SAP vs Excel
def firma_insumos(filas):
    prods = {}
    for r in filas[1:]:
        if not r or not r[0] or not str(r[0]).strip().upper().startswith("MGM") or not r[2]:
            continue
        try:
            q = round(float(str(r[4]).replace(",", ".")) if isinstance(r[4], str) else float(r[4] or 0), 4)
        except ValueError:
            q = 0
        prods.setdefault(str(r[0]).strip(), []).append((str(r[2]).strip(), q))
    return {k: sorted(v) for k, v in prods.items()}


def comparar(sap, excel):
    a, b = firma_insumos(sap), firma_insumos(excel)
    solo_sap = sorted(set(a) - set(b))
    solo_excel = sorted(set(b) - set(a))
    distintos = sorted(k for k in set(a) & set(b) if a[k] != b[k])
    return {"sap": len(a), "excel": len(b), "solo_sap": solo_sap[:50], "solo_excel": solo_excel[:50],
            "distintos": distintos[:50], "n_solo_sap": len(solo_sap), "n_solo_excel": len(solo_excel),
            "n_distintos": len(distintos)}


# ---------------------------------------------------------------- servidor
class Estado:
    def __init__(self):
        self.cfg = leer_config()
        self.sap = SAP(self.cfg)
        self.cache_excel = (None, 0, None)  # ruta, mtime, datos
        self.lock = threading.Lock()
        self.ultimo_ping = 0

    def maestros(self, importar=False):
        """Precios y materias primas: de maestros.json; se crea (o reimporta) desde el Excel."""
        m = None if importar else leer_json(MAESTROS)
        if m:
            return m
        ruta, (lp, bl, _ins) = self.excel()
        m = {"lp": lp, "bl": bl, "origen": os.path.basename(ruta), "at": time.time()}
        guardar_json(MAESTROS, m)
        log("Precios y aportes copiados desde", ruta)
        return m

    def insumos_excel(self):
        try:
            return self.excel()[1][2]
        except Exception:  # noqa: BLE001  (el Excel es opcional)
            return None

    def excel(self):
        ruta = buscar_excel(self.cfg)
        if not ruta:
            raise RuntimeError("no encontré el Excel del Blender. Escribe su ruta en config.ini (sección [excel]).")
        mt = os.path.getmtime(ruta)
        if self.cache_excel[0] != ruta or self.cache_excel[1] != mt:
            log("Leyendo", ruta)
            self.cache_excel = (ruta, mt, leer_excel(ruta))
        return ruta, self.cache_excel[2]

    def datos(self, forzar):
        with self.lock:
            return self._datos(forzar)

    def _datos(self, forzar):
        m = self.maestros()
        avisos = []
        filas_sap = self.sap.obtener(forzar)
        ins_excel = None
        if filas_sap:
            ins, origen = filas_sap, "SAP"
        else:
            ins_excel = self.insumos_excel()
            ins, origen = ins_excel or [], "Excel"
            if not ins_excel:
                avisos.append("No hay insumos: SAP no respondió y no encontré el Excel del Blender.")
        if self.sap.bloqueado:
            avisos.append(self.sap.bloqueado)
        elif self.sap.error:
            avisos.append(self.sap.error + (" Se usan los insumos del Excel." if origen == "Excel" else
                                            " Se mantienen los últimos datos leídos de SAP."))
        if filas_sap and self.sap.de_cache:
            avisos.append("Insumos de la última lectura de SAP (" +
                          time.strftime("%d-%m %H:%M", time.localtime(self.sap.ultimo_ok)) + ").")
        cmp = None
        if filas_sap:
            ie = self.insumos_excel()
            cmp = comparar(filas_sap, ie) if ie else None
        return {
            "name": "precios guardados en este PC", "maestrosDe": m.get("origen"), "lp": m["lp"], "bl": m["bl"],
            "ins": ins, "origen": origen, "sapAt": self.sap.ultimo_ok or None, "avisos": avisos,
            "comparacion": cmp,
        }


def buscar_html():
    for p in (os.path.join(AQUI, "taskpane.html"), os.path.join(AQUI, "..", "taskpane.html")):
        if os.path.exists(p):
            return p
    raise FileNotFoundError("taskpane.html")


def hacer_handler(estado):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def enviar(self, code, cuerpo, tipo):
            b = cuerpo if isinstance(cuerpo, bytes) else cuerpo.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", tipo)
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(b)

        def do_GET(self):
            u = urlparse(self.path)
            if u.path in ("/", "/index.html", "/taskpane.html"):
                with open(buscar_html(), "rb") as f:
                    return self.enviar(200, f.read(), "text/html; charset=utf-8")
            if u.path == "/api/datos":
                forzar = parse_qs(u.query).get("forzar", ["0"])[0] == "1"
                try:
                    d = estado.datos(forzar)
                    return self.enviar(200, json.dumps(d, ensure_ascii=False, default=str), "application/json")
                except Exception as e:  # noqa: BLE001
                    log("Error:", e)
                    return self.enviar(500, json.dumps({"error": str(e)}, ensure_ascii=False), "application/json")
            if u.path == "/api/ajustes":
                return self.enviar(200, json.dumps(leer_json(AJUSTES, {}), ensure_ascii=False), "application/json")
            if u.path == "/api/ping":
                estado.ultimo_ping = time.time()
                return self.enviar(200, "{}", "application/json")
            if u.path.endswith(".png"):
                p = os.path.join(AQUI, "..", os.path.basename(u.path))
                if os.path.exists(p):
                    with open(p, "rb") as f:
                        return self.enviar(200, f.read(), "image/png")
            self.enviar(404, "No encontrado", "text/plain; charset=utf-8")
        def do_POST(self):
            u = urlparse(self.path)
            # solo peticiones de la propia página (evita que otra web del navegador escriba aquí)
            if self.headers.get("Origin") not in (None, f"http://localhost:{estado.cfg['puerto']}",
                                                  f"http://127.0.0.1:{estado.cfg['puerto']}"):
                return self.enviar(403, "{}", "application/json")
            try:
                n = int(self.headers.get("Content-Length") or 0)
                cuerpo = json.loads(self.rfile.read(n) or b"{}")
                if u.path == "/api/ajustes":
                    datos = {k: cuerpo.get(k, {}) for k in ("prices", "leyes", "uso")}
                    guardar_json(AJUSTES, datos)
                    return self.enviar(200, '{"ok":true}', "application/json")
                if u.path == "/api/importar-excel":
                    with estado.lock:
                        m = estado.maestros(importar=True)
                    return self.enviar(200, json.dumps({"ok": True, "origen": m["origen"]}, ensure_ascii=False),
                                       "application/json")
            except Exception as e:  # noqa: BLE001
                log("Error:", e)
                return self.enviar(500, json.dumps({"error": str(e)}, ensure_ascii=False), "application/json")
            self.enviar(404, "{}", "application/json")
    return H


class Servidor(ThreadingHTTPServer):
    # en Windows, reutilizar la dirección permitiría abrir dos Blender en el mismo puerto
    allow_reuse_address = os.name != "nt"
    daemon_threads = True


def buscar_edge():
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.exists(p):
                return p
    return None


def abrir_ventana(url):
    """Abre el Blender en una ventana propia (Edge en modo app: sin pestañas ni barra de direcciones)."""
    edge = buscar_edge()
    if not edge:
        webbrowser.open(url)
        return
    perfil = os.path.join(os.environ.get("LOCALAPPDATA", AQUI), "BlenderCNA", "ventana")
    import subprocess
    subprocess.Popen([edge, f"--app={url}", f"--user-data-dir={perfil}", "--no-first-run",
                      "--no-default-browser-check", "--window-size=1440,920"])


def cerrar_si_no_hay_ventana(srv, estado):
    """En modo ventana: cuando se cierra la ventana (la página deja de dar señales), se apaga la app."""
    inicio = time.time()
    while True:
        time.sleep(10)
        ref = estado.ultimo_ping or inicio
        espera = 45 if estado.ultimo_ping else 180
        if time.time() - ref > espera:
            log("Ventana cerrada: se apaga el Blender")
            srv.shutdown()
            return


def main():
    ventana = "--ventana" in sys.argv
    estado = Estado()
    puerto = estado.cfg["puerto"]
    url = f"http://localhost:{puerto}/"
    try:
        srv = Servidor(("127.0.0.1", puerto), hacer_handler(estado))  # solo este PC
    except OSError:
        # ya hay un Blender abierto: solo se abre otra ventana
        log("El Blender ya estaba abierto; se abre otra ventana")
        abrir_ventana(url) if ventana else webbrowser.open(url)
        return
    log("Blender CNA abierto en", url, "" if ventana else "(cierra esta ventana para salir)")
    if ventana:
        threading.Timer(0.5, lambda: abrir_ventana(url)).start()
        threading.Thread(target=cerrar_si_no_hay_ventana, args=(srv, estado), daemon=True).start()
    elif "--sin-navegador" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        log("Error al iniciar:", repr(e))
        if "--ventana" in sys.argv:
            aviso_windows("No se pudo abrir el Blender: " + str(e) + "\n\nRevisa local\\blender.log")
        raise
