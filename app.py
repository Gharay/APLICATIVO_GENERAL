"""
app.py
Aplicación Flask principal — PostgreSQL + Flask-Migrate
"""
import os
import webbrowser
from flask import Flask, render_template
from flask_cors import CORS
from flask_migrate import Migrate
from database.models import db
from backend.routes import bp_api

app = Flask(__name__)

BASE_DIR     = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')

print(f"📁 BASE_DIR: {BASE_DIR}")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ── Configuración PostgreSQL ──────────────────────────────────
# Ajusta usuario y contraseña según tu instalación local
app.config['SQLALCHEMY_DATABASE_URI'] = (
    'postgresql://postgres:XXXXXX@localhost:5433/sic_catastral'
)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH']             = 100 * 1024 * 1024
app.config['UPLOAD_FOLDER']                  = UPLOAD_FOLDER

print(f"🔗 SQLALCHEMY_DATABASE_URI: {app.config['SQLALCHEMY_DATABASE_URI']}")

# Inicializar extensiones
db.init_app(app)
migrate = Migrate(app, db)   # ← Flask-Migrate registrado aquí

CORS(app)
app.register_blueprint(bp_api, url_prefix='/api')

# Rutas de vistas
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/editar/<int:id_predio>')
def editar_predio(id_predio):
    return render_template('form_predio.html', id_predio=id_predio)

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 INICIANDO SERVIDOR FLASK")
    print("="*60 + "\n")
    webbrowser.open('http://localhost:5000')
    app.run(debug=False, port=5000, host='127.0.0.1')