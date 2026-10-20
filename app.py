import os
import secrets

from flask import Flask, Response, render_template, request

from status_core import StatusError, generar_status

app = Flask(__name__)
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
    so = (request.form.get("so") or "").strip()
    final = bool(request.form.get("final"))

    def volver(mensaje, codigo=400):
        return render_template("index.html", error=mensaje, so=so, final=final), codigo

    if not f_stock or not f_stock.filename or not f_ventas or not f_ventas.filename:
        return volver("Suba el Excel de stock y el Excel de ventas.")
    for f in (f_stock, f_ventas):
        if not f.filename.lower().endswith(EXTENSIONES):
            return volver(f"El archivo «{f.filename}» no es .xlsx. Guárdelo como Libro de Excel (.xlsx) y vuelva a subirlo.")

    try:
        resultado = generar_status(f_ventas.read(), f_stock.read(), so)
    except StatusError as exc:
        return volver(str(exc))
    except Exception:
        app.logger.exception("Error procesando %s", so)
        return volver("No se pudo procesar alguno de los archivos. Revise que sean los reportes originales sin modificar.", 500)

    return render_template("index.html", r=resultado, so=so, final=final)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
