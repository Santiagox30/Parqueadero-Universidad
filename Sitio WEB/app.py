# app_clases.py
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from datetime import datetime
import pandas as pd
import os
import pyqrcode
import io

app = Flask(__name__)
app.secret_key = 'clave_secreta'

TARIFAS = {
    "carro": 5000,
    "moto": 2000,
    "moto eléctrica": 2000,
    "bicicleta": 1000,
    "patineta eléctrica": 2000
}

class Vehiculo:
    def __init__(self, uid, placa, nombre, apellido, tipo, tipo_usuario, hora_inicio=None, hora_fin=None, total=0, total_mes="", qr=None):
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
        entrada = datetime.strptime(self.hora_inicio, "%Y-%m-%d %H:%M:%S")
        salida = datetime.now()
        horas = max(1, int((salida - entrada).total_seconds() // 3600))
        self.total = horas * TARIFAS.get(self.tipo, 0)
        self.hora_fin = salida.strftime("%Y-%m-%d %H:%M:%S")

    def to_dict(self):
        return self.__dict__

class ParqueaderoManager:
    def __init__(self):
        self.vehiculos = {}
        self.historial = []
        self.archivo_excel = "BD_Parqueadero.xlsx"
        self.cargar_datos()

    def cargar_datos(self):
        if os.path.exists(self.archivo_excel):
            df = pd.read_excel(self.archivo_excel)
            for _, fila in df.iterrows():
                uid = str(fila['id'])
                self.vehiculos[uid] = Vehiculo(
                    uid=uid,
                    placa=fila['placa'],
                    nombre=fila['nombre'],
                    apellido=fila['apellido'],
                    tipo=fila['tipo de vehículo'],
                    tipo_usuario=fila['tipo usuario'],
                    hora_inicio=fila['hora Inicio'],
                    hora_fin=fila['hora Fin'],
                    total=fila['total'],
                    total_mes=fila['total mes']
                )

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
            vehiculo = self.vehiculos[uid]
            if not vehiculo.hora_fin:
                vehiculo.calcular_pago()
                self.historial.append({
                    "placa": vehiculo.placa,
                    "nombre": vehiculo.nombre,
                    "apellido": vehiculo.apellido,
                    "tipo de vehículo": vehiculo.tipo,
                    "hora Inicio": vehiculo.hora_inicio,
                    "hora Fin": vehiculo.hora_fin,
                    "total": vehiculo.total
                })

    def get_historial(self):
        return self.historial

    def get_reporte_excel(self):
        df = pd.DataFrame([v.to_dict() for v in self.vehiculos.values()])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Parqueadero')
        output.seek(0)
        return output

parqueadero = ParqueaderoManager()

def generar_qr(uid):
    url = f"https://mipago.com/pay?vehiculo={uid}"
    qr = pyqrcode.create(url)
    ruta = f"static/qr_{uid}.png"
    qr.png(ruta, scale=5)
    return ruta

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form.get('usuario') == 'admin' and request.form.get('password') == '1234':
            session['usuario'] = 'admin'
            return redirect(url_for('panel'))
        return render_template('login.html', error="Datos incorrectos")
    return render_template('login.html')

@app.route('/panel')
def panel():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', parqueadero=parqueadero.vehiculos)

@app.route('/ingresar', methods=["POST"])
def ingresar():
    tipo = request.form["tipo de vehículo"]
    placa = request.form["placa"]
    uid = f"{tipo}-{placa}"
    vehiculo = Vehiculo(
        uid=uid,
        placa=placa,
        nombre=request.form.get("nombre", ""),
        apellido=request.form.get("apellido", ""),
        tipo=tipo,
        tipo_usuario=request.form.get("tipo_usuario", ""),
        qr=generar_qr(uid)
    )
    parqueadero.agregar_vehiculo(vehiculo)
    return redirect(url_for("panel"))

@app.route('/salir/<uid>')
def salir(uid):
    parqueadero.registrar_salida(uid)
    return redirect(url_for("panel"))

@app.route('/editar/<uid>', methods=['GET', 'POST'])
def editar(uid):
    if request.method == 'POST':
        vehiculo = parqueadero.vehiculos[uid]
        vehiculo.placa = request.form["placa"]
        vehiculo.nombre = request.form["nombre"]
        vehiculo.apellido = request.form["apellido"]
        vehiculo.tipo = request.form["tipo"]
        vehiculo.tipo_usuario = request.form["tipo_usuario"]
        vehiculo.total_mes = request.form.get("total_mes", "")
        return redirect(url_for("panel"))
    return render_template("editar.html", uid=uid, datos=parqueadero.vehiculos[uid])

@app.route('/eliminar/<uid>')
def eliminar(uid):
    parqueadero.eliminar_vehiculo(uid)
    return redirect(url_for('panel'))

@app.route('/historial')
def historial_view():
    return render_template("historial.html", historial=parqueadero.get_historial())

@app.route('/reporte')
def reporte():
    output = parqueadero.get_reporte_excel()
    return send_file(
        output,
        as_attachment=True,
        download_name="reporte_parqueadero.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/logout')
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))

if __name__ == '__main__':
    app.run(debug=True)
