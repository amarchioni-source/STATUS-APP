import os
import secrets

from flask import Flask, Response, render_template, request
from werkzeug.exceptions import HTTPException

from status_core import StatusError, generar_status

app = Flask(__name__, template_folder=".", static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB entre ambos archivos

APP_USER = os.environ.get("APP_USER", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
EXTENSIONES = (".xlsx", ".xlsm")


@app.before_request
def requerir_acceso():
    if request.path == "/salud":
        return None
    if not APP_USER or not APP_PASSWORD:
        return None  # sin usuario/contraseña configurados, la app abre sin login
    auth = request.authorization
    valido = (
        auth is not None
        and auth.username is not None
        and auth.password is not None
        and secrets.compare_digest(auth.username.encode(), APP_USER.encode())
        and secrets.compare_digest(auth.password.encode(), APP_PASSWORD.encode())
    )
    if not valido:
        return Response("Acceso restringido.", 401, {"WWW-Authenticate": 'Basic realm="Status por SO"'})
    return None


@app.get("/salud")
def salud():
    return "ok"


@app.get("/")
def inicio():
    return render_template("index.html")


@app.post("/status")
def status():
    f_stock = request.files.get("stock")
    f_ventas = request.files.get("ventas")
    f_ship = request.files.get("shipment")
    so = (request.form.get("so") or "").strip()
    pedidos = (request.form.get("pedidos") or "").strip()
    final = bool(request.form.get("final"))

    def cargado(f):
        return bool(f and f.filename)

    modo = request.form.get("modo") or ("shipment" if cargado(f_ship) else "so")

    def volver(mensaje, codigo=400):
        return render_template("index.html", error=mensaje, so=so, pedidos=pedidos, final=final, modo=modo), codigo

    if not cargado(f_stock):
        return volver("Suba el Excel de stock.")
    if modo == "shipment":
        if not cargado(f_ship):
            return volver("Suba el Excel de shipment.")
        f_ventas = None
    else:
        if not cargado(f_ventas):
            return volver("Suba el Excel de ventas.")
        f_ship = None
    for f in (f_stock, f_ventas, f_ship):
        if cargado(f) and not f.filename.lower().endswith(EXTENSIONES):
            return volver(f"El archivo «{f.filename}» no es .xlsx. Guárdelo como Libro de Excel (.xlsx) y vuelva a subirlo.")

    try:
        if cargado(f_ship):
            resultado = generar_status(None, f_stock.read(), so, shipment_bytes=f_ship.read(), pedidos_txt=pedidos)
        else:
            resultado = generar_status(f_ventas.read(), f_stock.read(), so)
    except StatusError as exc:
        return volver(str(exc))
    except Exception:
        app.logger.exception("Error procesando %s", so)
        return volver("No se pudo procesar alguno de los archivos. Revise que sean los reportes originales sin modificar.", 500)

    return render_template("index.html", r=resultado, so=so, pedidos=pedidos, final=final, modo=modo)


@app.errorhandler(413)
def archivo_muy_grande(_exc):
    return render_template("index.html", error="Los archivos superan el tamaño máximo permitido (80 MB en total)."), 413


@app.errorhandler(Exception)
def error_inesperado(exc):
    if isinstance(exc, HTTPException):
        return exc
    app.logger.exception("Error no controlado")
    return (
        render_template(
            "index.html",
            error=f"Error inesperado ({type(exc).__name__}): {exc}",
            so=request.form.get("so", ""),
            pedidos=request.form.get("pedidos", ""),
            final=bool(request.form.get("final")),
            modo=request.form.get("modo", "so"),
        ),
        500,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
