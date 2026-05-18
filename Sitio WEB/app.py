# app.py - Sistema de Parqueadero con Detección de Placas
from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify
from datetime import datetime
import pandas as pd
import numpy as np
import cv2
import pytesseract
import pyqrcode
import base64
import io
import os
import re

app = Flask(__name__)
app.secret_key = 'clave_secreta_parqueadero'

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────

TARIFAS = {
    "carro": 5000,
    "moto": 2000,
    "moto eléctrica": 2000,
    "bicicleta": 1000,
    "patineta eléctrica": 2000
}

EXCEL_FILE = "BD_Parqueadero.xlsx"

# Si Tesseract no está en el PATH del sistema, descomenta y ajusta:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

CONFIG_OCR = r"--oem 3 --psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

# ──────────────────────────────────────────────
# CLASES
# ──────────────────────────────────────────────

class Vehiculo:
    def __init__(self, uid, placa, nombre, apellido, tipo, tipo_usuario,
                 hora_inicio=None, hora_fin=None, total=0, total_mes="", qr=None):
        self.uid = uid
        self.placa = placa
        self.nombre = nombre
        self.apellido = apellido
        self.tipo = tipo
        self.tipo_usuario = tipo_usuario
        self.hora_inicio = hora_inicio or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.hora_fin = hora_fin
        self.total = total
        self.total_mes = total_mes
        self.qr = qr

    def calcular_pago(self):
        try:
            entrada = datetime.strptime(self.hora_inicio, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            entrada = datetime.now()
        salida = datetime.now()
        horas = max(1, int((salida - entrada).total_seconds() // 3600))
        self.total = horas * TARIFAS.get(self.tipo, 0)
        self.hora_fin = salida.strftime("%Y-%m-%d %H:%M:%S")

    def to_dict(self):
        return {
            "uid": self.uid,
            "placa": self.placa,
            "nombre": self.nombre,
            "apellido": self.apellido,
            "tipo": self.tipo,
            "tipo_usuario": self.tipo_usuario,
            "hora_inicio": self.hora_inicio,
            "hora_fin": self.hora_fin,
            "total": self.total,
            "total_mes": self.total_mes,
        }


class ParqueaderoManager:
    def __init__(self):
        self.vehiculos = {}
        self.historial = []
        self.cargar_datos()

    def cargar_datos(self):
        if os.path.exists(EXCEL_FILE):
            try:
                df = pd.read_excel(EXCEL_FILE)
                for _, fila in df.iterrows():
                    uid = str(fila['id'])
                    hora_inicio = str(fila['hora Inicio']) if pd.notna(fila['hora Inicio']) else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    hora_fin = str(fila['hora Fin']) if pd.notna(fila.get('hora Fin')) else None

                    # Normalizar formato de hora si viene solo como HH:MM:SS
                    for h in [hora_inicio]:
                        if len(h) <= 8 and ':' in h:
                            hora_inicio = datetime.now().strftime("%Y-%m-%d") + " " + h

                    self.vehiculos[uid] = Vehiculo(
                        uid=uid,
                        placa=str(fila['placa']) if pd.notna(fila['placa']) else "",
                        nombre=str(fila['nombre']) if pd.notna(fila['nombre']) else "",
                        apellido=str(fila['apellido']) if pd.notna(fila['apellido']) else "",
                        tipo=str(fila['tipo de vehículo']) if pd.notna(fila['tipo de vehículo']) else "carro",
                        tipo_usuario=str(fila['tipo usuario']) if pd.notna(fila['tipo usuario']) else "",
                        hora_inicio=hora_inicio,
                        hora_fin=hora_fin,
                        total=int(fila['total']) if pd.notna(fila['total']) else 0,
                        total_mes=str(fila['total mes']) if pd.notna(fila.get('total mes')) else "",
                    )
            except Exception as e:
                print(f"[AVISO] No se pudo cargar el Excel: {e}")

    def agregar_vehiculo(self, vehiculo):
        self.vehiculos[vehiculo.uid] = vehiculo

    def eliminar_vehiculo(self, uid):
        if uid in self.vehiculos:
            ruta_qr = self.vehiculos[uid].qr
            if ruta_qr and os.path.exists(ruta_qr):
                os.remove(ruta_qr)
            del self.vehiculos[uid]

    def registrar_salida(self, uid):
        if uid in self.vehiculos:
            v = self.vehiculos[uid]
            if not v.hora_fin:
                v.calcular_pago()
                self.historial.append({
                    "placa": v.placa,
                    "nombre": v.nombre,
                    "apellido": v.apellido,
                    "tipo de vehículo": v.tipo,
                    "hora Inicio": v.hora_inicio,
                    "hora Fin": v.hora_fin,
                    "total": v.total
                })

    def get_reporte_excel(self):
        filas = []
        for v in self.vehiculos.values():
            filas.append({
                "id": v.uid, "placa": v.placa, "nombre": v.nombre,
                "apellido": v.apellido, "tipo de vehículo": v.tipo,
                "tipo usuario": v.tipo_usuario, "hora Inicio": v.hora_inicio,
                "hora Fin": v.hora_fin or "", "total": v.total,
                "total mes": v.total_mes
            })
        df = pd.DataFrame(filas)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Parqueadero', index=False)
        output.seek(0)
        return output


parqueadero = ParqueaderoManager()

# ──────────────────────────────────────────────
# DETECCIÓN DE PLACAS (OpenCV + Tesseract)
# ──────────────────────────────────────────────

def detectar_placa_desde_imagen(imagen_bgr):
    """Procesa imagen BGR y retorna el texto de la placa detectada."""
    gris = cv2.cvtColor(imagen_bgr, cv2.COLOR_BGR2GRAY)
    gris = cv2.bilateralFilter(gris, 11, 17, 17)
    bordes = cv2.Canny(gris, 30, 200)

    contornos, _ = cv2.findContours(bordes.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contornos = sorted(contornos, key=cv2.contourArea, reverse=True)[:10]

    contorno_placa = None
    for c in contornos:
        peri = cv2.arcLength(c, True)
        aprox = cv2.approxPolyDP(c, 0.018 * peri, True)
        if len(aprox) == 4:
            contorno_placa = aprox
            break

    texto = ""
    imagen_resultado = imagen_bgr.copy()

    if contorno_placa is not None:
        x, y, w, h = cv2.boundingRect(contorno_placa)
        roi_gris = gris[y:y+h, x:x+w]

        # Escalar ROI para mejor OCR
        alto_objetivo = 100
        factor = alto_objetivo / max(roi_gris.shape[0], 1)
        roi_escalado = cv2.resize(roi_gris, (int(roi_gris.shape[1] * factor), alto_objetivo), interpolation=cv2.INTER_CUBIC)
        roi_bin = cv2.adaptiveThreshold(roi_escalado, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

        texto_raw = pytesseract.image_to_string(roi_bin, config=CONFIG_OCR)
        texto = re.sub(r"[^A-Z0-9]", "", texto_raw.upper().strip())

        # Dibujar resultado en imagen
        color = (0, 200, 0) if len(texto) >= 4 else (0, 140, 255)
        cv2.drawContours(imagen_resultado, [contorno_placa], -1, color, 3)
        cv2.rectangle(imagen_resultado, (x, max(y-35, 0)), (x + w, y), color, -1)
        cv2.putText(imagen_resultado, texto or "?", (x + 5, max(y - 8, 12)),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 255, 255), 2)

    # Convertir imagen resultado a base64 para mostrar en HTML
    _, buffer = cv2.imencode('.jpg', imagen_resultado)
    img_b64 = base64.b64encode(buffer).decode('utf-8')

    return texto, img_b64


def imagen_desde_request(archivo):
    """Convierte el archivo subido en numpy array BGR."""
    contenido = archivo.read()
    arr = np.frombuffer(contenido, np.uint8)
    return cv2.imdecode(arr, cv2.IMREAD_COLOR)


# ──────────────────────────────────────────────
# QR
# ──────────────────────────────────────────────

def generar_qr(uid):
    os.makedirs("static", exist_ok=True)
    url = f"https://mipago.com/pay?vehiculo={uid}"
    qr = pyqrcode.create(url)
    # Nombre seguro para archivo (sin caracteres especiales)
    nombre_seguro = re.sub(r"[^a-zA-Z0-9_-]", "_", uid)
    ruta = f"static/qr_{nombre_seguro}.png"
    qr.png(ruta, scale=5)
    return ruta


# ──────────────────────────────────────────────
# RUTAS
# ──────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form.get('usuario') == 'admin' and request.form.get('password') == '1234':
            session['usuario'] = 'admin'
            return redirect(url_for('panel'))
        return render_template('login.html', error="Usuario o contraseña incorrectos")
    return render_template('login.html')


@app.route('/panel')
def panel():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', parqueadero=parqueadero.vehiculos)


@app.route('/ingresar', methods=["POST"])
def ingresar():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    tipo = request.form["tipo de vehículo"]
    placa = request.form.get("placa", "").upper().strip()
    uid = f"{tipo}-{placa}-{datetime.now().strftime('%H%M%S')}"
    vehiculo = Vehiculo(
        uid=uid, placa=placa,
        nombre=request.form.get("nombre", ""),
        apellido=request.form.get("apellido", ""),
        tipo=tipo,
        tipo_usuario=request.form.get("tipo_usuario", ""),
        qr=generar_qr(uid)
    )
    parqueadero.agregar_vehiculo(vehiculo)
    return redirect(url_for("panel"))


@app.route('/salir/<path:uid>')
def salir(uid):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    parqueadero.registrar_salida(uid)
    return redirect(url_for("panel"))


@app.route('/editar/<path:uid>', methods=['GET', 'POST'])
def editar(uid):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    if uid not in parqueadero.vehiculos:
        return redirect(url_for('panel'))
    if request.method == 'POST':
        v = parqueadero.vehiculos[uid]
        v.placa = request.form["placa"].upper().strip()
        v.nombre = request.form["nombre"]
        v.apellido = request.form["apellido"]
        v.tipo = request.form["tipo"]
        v.tipo_usuario = request.form["tipo_usuario"]
        v.total_mes = request.form.get("total_mes", "")
        return redirect(url_for("panel"))
    return render_template("editar.html", uid=uid, datos=parqueadero.vehiculos[uid])


@app.route('/eliminar/<path:uid>')
def eliminar(uid):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    parqueadero.eliminar_vehiculo(uid)
    return redirect(url_for('panel'))


@app.route('/historial')
def historial_view():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    return render_template("historial.html", historial=parqueadero.historial)


@app.route('/reporte')
def reporte():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    output = parqueadero.get_reporte_excel()
    return send_file(output, as_attachment=True,
                     download_name="reporte_parqueadero.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/logout')
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))


# ── RUTAS DE DETECCIÓN DE PLACA ──────────────

@app.route('/detectar_placa', methods=['POST'])
def detectar_placa():
    """Recibe imagen (archivo o webcam base64) y devuelve placa detectada + imagen procesada."""
    if 'usuario' not in session:
        return jsonify({"error": "no autorizado"}), 401

    texto = ""
    img_b64 = ""

    # Opción 1: archivo subido
    if 'imagen' in request.files and request.files['imagen'].filename:
        archivo = request.files['imagen']
        imagen = imagen_desde_request(archivo)
        if imagen is not None:
            texto, img_b64 = detectar_placa_desde_imagen(imagen)

    # Opción 2: captura base64 desde webcam
    elif request.form.get('imagen_b64'):
        data_url = request.form['imagen_b64']
        header, encoded = data_url.split(',', 1)
        img_bytes = base64.b64decode(encoded)
        arr = np.frombuffer(img_bytes, np.uint8)
        imagen = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        if imagen is not None:
            texto, img_b64 = detectar_placa_desde_imagen(imagen)

    return jsonify({"placa": texto, "imagen_resultado": img_b64})


if __name__ == '__main__':
    os.makedirs("static", exist_ok=True)
    app.run(debug=True)
