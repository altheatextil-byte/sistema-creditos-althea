from flask import Flask, render_template, request, redirect, url_for, session, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from datetime import datetime, timedelta
from functools import wraps
from werkzeug.middleware.proxy_fix import ProxyFix
import os
import io
import secrets
import openpyxl
import requests

app = Flask(__name__)

# Render define RENDER=true en sus servicios.
IS_PRODUCTION = os.environ.get('RENDER') == 'true' or os.environ.get('FLASK_ENV') == 'production'

# --- SECRET KEY ---
SECRET_KEY = os.environ.get('SECRET_KEY')
if not SECRET_KEY:
    if IS_PRODUCTION:
        raise RuntimeError(
            "SECRET_KEY no está configurada. Defínela en el dashboard de Render "
            "(Environment > SECRET_KEY) antes de desplegar."
        )
    SECRET_KEY = 'dev-only-insecure-key'
app.secret_key = SECRET_KEY

# Render termina TLS en su proxy; esto permite detectar HTTPS y el host real.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

# --- COOKIES DE SESIÓN ---
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=IS_PRODUCTION,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
    MAX_CONTENT_LENGTH=2 * 1024 * 1024,  # 2 MB
)

# --- BASE DE DATOS ---
basedir = os.path.abspath(os.path.dirname(__file__))
instance_path = os.path.join(basedir, 'instance')
os.makedirs(instance_path, exist_ok=True)


def normalizar_database_url(url):
    """Normaliza la URL de Render y fija el driver de PostgreSQL instalado.

    - Render entrega a veces 'postgres://' (SQLAlchemy 2.x ya no lo acepta).
    - 'postgresql://' hace que SQLAlchemy asuma psycopg2; lo dejamos explícito
      como 'postgresql+psycopg2://' para que el error, si falta el paquete,
      sea claro y no un ModuleNotFoundError genérico en tiempo de conexión.
    """
    if not url:
        return url
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    if url.startswith('postgresql://'):
        try:
            import psycopg2  # noqa: F401
        except ImportError as exc:
            raise RuntimeError(
                "El driver 'psycopg2' no está instalado. Ejecuta "
                "'pip install -r requirements.txt' y verifica que requirements.txt "
                "esté en la raíz del repositorio (Root Directory vacío en Render)."
            ) from exc
        return 'postgresql+psycopg2://' + url[len('postgresql://'):]
    return url  # SQLite u otro motor, o driver ya explícito (postgresql+psycopg://)


app.config['SQLALCHEMY_DATABASE_URI'] = normalizar_database_url(
    os.environ.get('DATABASE_URL', 'sqlite:///' + os.path.join(instance_path, 'solicitudes.db'))
)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_pre_ping': True,   # evita conexiones muertas tras el idle de Render
    'pool_recycle': 300,
}
db = SQLAlchemy(app)

# --- CREDENCIALES DEL PANEL ---
ADMIN_USER = os.environ.get('ADMIN_USER', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD')
if not ADMIN_PASSWORD:
    if IS_PRODUCTION:
        raise RuntimeError(
            "ADMIN_PASSWORD no está configurada. Defínela en Render > Environment "
            "antes de desplegar (la contraseña anterior estaba publicada en el repo)."
        )
    ADMIN_PASSWORD = 'althea2026'  # solo para desarrollo local

# --- CONFIGURACIÓN DE RESEND ---
RESEND_API_KEY = os.environ.get('RESEND_API_KEY', '')
RESEND_URL = "https://api.resend.com/emails"

def enviar_correo_resend(destinatario, asunto, html):
    if not RESEND_API_KEY:
        print("No hay API Key de Resend configurada.")
        return False
    try:
        response = requests.post(
            RESEND_URL,
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "from": "Althea Textil BG <onboarding@resend.dev>",
                "to": [destinatario],
                "subject": asunto,
                "html": html
            },
            timeout=10
        )
        if response.status_code == 200:
            print(f"Correo enviado a {destinatario}")
            return True
        else:
            print(f"Error de Resend: {response.text}")
            return False
    except Exception as e:
        print(f"Error al enviar correo: {e}")
        return False

# --- MODELO DE LA BASE DE DATOS ---
class Solicitud(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, default=datetime.now)
    estado = db.Column(db.String(20), default="Pendiente")
    razon_social = db.Column(db.String(200))
    nit = db.Column(db.String(50))
    nombre_comercial = db.Column(db.String(200))
    actividad_economica = db.Column(db.String(200))
    direccion = db.Column(db.String(200))
    ciudad = db.Column(db.String(100))
    departamento = db.Column(db.String(100))
    telefono = db.Column(db.String(50))
    correo = db.Column(db.String(100))
    fecha_constitucion = db.Column(db.String(50))
    num_empleados = db.Column(db.String(50))
    tipo_empresa = db.Column(db.String(50))
    rep_nombre = db.Column(db.String(200))
    rep_documento = db.Column(db.String(50))
    rep_cargo = db.Column(db.String(100))
    rep_telefono = db.Column(db.String(50))
    rep_correo = db.Column(db.String(100))
    rep_direccion = db.Column(db.String(200))
    regimen_tributario = db.Column(db.String(100))
    responsabilidades = db.Column(db.String(200))
    responsable_iva = db.Column(db.String(10))
    camara_comercio = db.Column(db.String(200))
    fecha_matricula = db.Column(db.String(50))
    fecha_renovacion = db.Column(db.String(50))
    compras_credito = db.Column(db.String(10))
    plazo_habitual = db.Column(db.String(50))
    ventas_mensuales = db.Column(db.String(100))
    activos = db.Column(db.String(100))
    pasivos = db.Column(db.String(100))
    patrimonio = db.Column(db.String(100))
    bancos = db.Column(db.String(200))
    antiguedad_bancaria = db.Column(db.String(100))
    obligaciones = db.Column(db.String(10))
    mora = db.Column(db.String(10))
    cupo_solicitado = db.Column(db.String(100))
    plazo_solicitado = db.Column(db.String(50))
    tipo_productos = db.Column(db.String(100))
    frecuencia_compra = db.Column(db.String(50))
    valor_compras = db.Column(db.String(100))
    ref1_empresa = db.Column(db.String(200))
    ref1_contacto = db.Column(db.String(200))
    ref1_cargo = db.Column(db.String(100))
    ref1_telefono = db.Column(db.String(50))
    ref1_correo = db.Column(db.String(100))
    ref1_tiempo = db.Column(db.String(50))
    ref2_empresa = db.Column(db.String(200))
    ref2_contacto = db.Column(db.String(200))
    ref2_cargo = db.Column(db.String(100))
    ref2_telefono = db.Column(db.String(50))
    ref2_correo = db.Column(db.String(100))
    ref2_tiempo = db.Column(db.String(50))

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- RUTAS ---
@app.route('/')
def mostrar_formulario():
    return render_template('formulario.html')

@app.route('/enviar', methods=['POST'])
def enviar_solicitud():
    nueva = Solicitud(
        razon_social=request.form['razon_social'],
        nit=request.form['nit'],
        nombre_comercial=request.form['nombre_comercial'],
        actividad_economica=request.form['actividad_economica'],
        direccion=request.form['direccion'],
        ciudad=request.form['ciudad'],
        departamento=request.form['departamento'],
        telefono=request.form['telefono'],
        correo=request.form['correo'],
        fecha_constitucion=request.form['fecha_constitucion'],
        num_empleados=request.form['num_empleados'],
        tipo_empresa=request.form['tipo_empresa'],
        rep_nombre=request.form['rep_nombre'],
        rep_documento=request.form['rep_documento'],
        rep_cargo=request.form['rep_cargo'],
        rep_telefono=request.form['rep_telefono'],
        rep_correo=request.form['rep_correo'],
        rep_direccion=request.form['rep_direccion'],
        regimen_tributario=request.form['regimen_tributario'],
        responsabilidades=request.form['responsabilidades'],
        responsable_iva=request.form['responsable_iva'],
        camara_comercio=request.form['camara_comercio'],
        fecha_matricula=request.form['fecha_matricula'],
        fecha_renovacion=request.form['fecha_renovacion'],
        compras_credito=request.form['compras_credito'],
        plazo_habitual=request.form.get('plazo_habitual', ''),
        ventas_mensuales=request.form['ventas_mensuales'],
        activos=request.form['activos'],
        pasivos=request.form['pasivos'],
        patrimonio=request.form['patrimonio'],
        bancos=request.form['bancos'],
        antiguedad_bancaria=request.form['antiguedad_bancaria'],
        obligaciones=request.form['obligaciones'],
        mora=request.form['mora'],
        cupo_solicitado=request.form['cupo_solicitado'],
        plazo_solicitado=request.form['plazo_solicitado'],
        tipo_productos=request.form['tipo_productos'],
        frecuencia_compra=request.form['frecuencia_compra'],
        valor_compras=request.form['valor_compras'],
        ref1_empresa=request.form['ref1_empresa'],
        ref1_contacto=request.form['ref1_contacto'],
        ref1_cargo=request.form['ref1_cargo'],
        ref1_telefono=request.form['ref1_telefono'],
        ref1_correo=request.form['ref1_correo'],
        ref1_tiempo=request.form['ref1_tiempo'],
        ref2_empresa=request.form['ref2_empresa'],
        ref2_contacto=request.form['ref2_contacto'],
        ref2_cargo=request.form['ref2_cargo'],
        ref2_telefono=request.form['ref2_telefono'],
        ref2_correo=request.form['ref2_correo'],
        ref2_tiempo=request.form['ref2_tiempo']
    )
    db.session.add(nueva)
    db.session.commit()
    
    # Correo al administrador usando Resend
    html_admin = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 8px;">
        <div style="background-color: #1a1a2e; padding: 20px; text-align: center;">
            <h1 style="color: #c9a05b; margin: 0;">ALTHEA</h1>
            <p style="color: white; margin: 5px 0 0 0;">Nueva Solicitud de Crédito</p>
        </div>
        <div style="padding: 30px;">
            <h2 style="color: #1a1a2e;">Resumen</h2>
            <p><strong>Empresa:</strong> {nueva.razon_social}</p>
            <p><strong>NIT:</strong> {nueva.nit}</p>
            <p><strong>Representante:</strong> {nueva.rep_nombre}</p>
            <p><strong>Cupo:</strong> {nueva.cupo_solicitado}</p>
            <p><strong>Plazo:</strong> {nueva.plazo_solicitado}</p>
            <br>
            <a href="https://sistema-creditos-althea.onrender.com/panel" style="background-color: #c9a05b; color: #1a1a2e; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">Ver en el Panel</a>
        </div>
    </div>
    """
    enviar_correo_resend('altheatextil@gmail.com', f'Nueva Solicitud - {nueva.razon_social}', html_admin)
    
    return render_template('exito.html', nombre=nueva.razon_social)

@app.route('/favicon.ico')
def favicon():
    return send_file(
        os.path.join(app.root_path, 'static', 'logo.png'),
        mimetype='image/png',
    )


@app.route('/healthz')
def healthz():
    """Health check para Render (no requiere login)."""
    try:
        db.session.execute(text('SELECT 1'))
        return {'status': 'ok'}, 200
    except Exception:
        app.logger.exception('Health check: base de datos no disponible')
        return {'status': 'degraded'}, 503


@app.after_request
def agregar_cabeceras_seguridad(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; "
        "script-src 'self'; base-uri 'self'; form-action 'self'; frame-ancestors 'self'"
    )
    if request.is_secure:
        response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario', '')
        clave = request.form.get('clave', '')
        usuario_ok = secrets.compare_digest(usuario, ADMIN_USER)
        clave_ok = secrets.compare_digest(clave, ADMIN_PASSWORD)
        if usuario_ok and clave_ok:
            session.clear()
            session['logged_in'] = True
            session.permanent = True
            return redirect(url_for('panel'))
        return render_template('login.html', error="Usuario o contraseña incorrectos")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/panel')
@login_required
def panel():
    busqueda = request.args.get('busqueda', '')
    filtro_estado = request.args.get('estado', '')
    query = Solicitud.query
    if busqueda:
        query = query.filter((Solicitud.razon_social.contains(busqueda)) | (Solicitud.nit.contains(busqueda)))
    if filtro_estado:
        query = query.filter(Solicitud.estado == filtro_estado)
    solicitudes = query.order_by(Solicitud.fecha.desc()).all()
    total = Solicitud.query.count()
    pendientes = Solicitud.query.filter_by(estado="Pendiente").count()
    aprobadas = Solicitud.query.filter_by(estado="Aprobado").count()
    rechazadas = Solicitud.query.filter_by(estado="Rechazado").count()
    return render_template('panel.html', solicitudes=solicitudes, total=total, pendientes=pendientes, aprobadas=aprobadas, rechazadas=rechazadas, busqueda=busqueda, filtro_estado=filtro_estado)

@app.route('/ver/<int:id>')
@login_required
def ver_solicitud(id):
    solicitud = Solicitud.query.get_or_404(id)
    return render_template('detalle.html', s=solicitud)

@app.route('/cambiar_estado/<int:id>/<estado>')
@login_required
def cambiar_estado(id, estado):
    solicitud = Solicitud.query.get_or_404(id)
    solicitud.estado = estado
    db.session.commit()
    
    # Correo al cliente usando Resend
    if estado == "Aprobado":
        asunto = f"Credito Aprobado - {solicitud.razon_social}"
        color = "#27ae60"
        titulo = "Su solicitud ha sido APROBADA"
    else:
        asunto = f"Credito Rechazado - {solicitud.razon_social}"
        color = "#e74c3c"
        titulo = "Solicitud No Aprobada"
    
    html_cliente = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 8px;">
        <div style="background-color: #1a1a2e; padding: 20px; text-align: center;">
            <h1 style="color: #c9a05b; margin: 0;">ALTHEA</h1>
        </div>
        <div style="padding: 30px;">
            <h2 style="color: {color};">{titulo}</h2>
            <p>Estimado(a) <strong>{solicitud.rep_nombre}</strong>,</p>
            <p>Su solicitud para la empresa <strong>{solicitud.razon_social}</strong> ha sido {estado.lower()}.</p>
            <p><strong>Cupo:</strong> {solicitud.cupo_solicitado}</p>
            <p><strong>Plazo:</strong> {solicitud.plazo_solicitado}</p>
        </div>
    </div>
    """
    enviar_correo_resend(solicitud.rep_correo, asunto, html_cliente)
    
    return redirect(url_for('ver_solicitud', id=id))

@app.route('/exportar_excel')
@login_required
def exportar_excel():
    solicitudes = Solicitud.query.all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solicitudes"
    headers = ["ID", "Fecha", "Estado", "Razón Social", "NIT", "Representante", "Teléfono", "Correo", "Cupo Solicitado", "Plazo"]
    ws.append(headers)
    for s in solicitudes:
        ws.append([s.id, s.fecha.strftime('%d/%m/%Y'), s.estado, s.razon_social, s.nit, s.rep_nombre, s.telefono, s.correo, s.cupo_solicitado, s.plazo_solicitado])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="Solicitudes_Althea.xlsx", as_attachment=True)

with app.app_context():
    try:
        db.create_all()
    except Exception:
        # No tumbamos el proceso si la BD aún no está lista: /healthz reportará 503
        # y Render mantendrá el servicio vivo mientras se corrige la variable DATABASE_URL.
        app.logger.exception('No se pudo inicializar la base de datos')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8501)), debug=not IS_PRODUCTION)