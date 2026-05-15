"""
database/models.py
Modelos SQLAlchemy para la BD del Sistema Catastral
"""
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

# ==================== TABLAS ====================

class Predio(db.Model):
    __tablename__ = 'predio'

    id_predio                        = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id                               = db.Column(db.String(100), unique=True, nullable=False)
    fmi                              = db.Column(db.String(100), nullable=False)
    propietario_vur                  = db.Column(db.String(2000))
    numero_documento_propietario_vur = db.Column(db.String(2000))
    nombre_predio_vur                = db.Column(db.String(2000))
    area_vur                         = db.Column(db.Float)
    area_final_gc                    = db.Column(db.Float)
    es_derivado                      = db.Column(db.Boolean, default=False)
    id_predio_matriz                 = db.Column(db.String(100))
    tiene_derivadas                  = db.Column(db.Boolean, default=False)
    derivadas_ids                    = db.Column(db.String(5000))
    proceso_englobe_desenglobe       = db.Column(db.Boolean, default=False)
    fecha_creacion                   = db.Column(db.DateTime, default=datetime.utcnow)
    fecha_actualizacion              = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    estado                           = db.Column(db.String(50), default='ACTIVO')
    tipo_gestion                     = db.Column(db.String(50))
    # Relaciones
    localizacion = db.relationship('Localizacion',     uselist=False, backref='predio', cascade='all, delete-orphan')
    catastral    = db.relationship('Catastral',        uselist=False, backref='predio', cascade='all, delete-orphan')
    juridica     = db.relationship('Juridica',         uselist=False, backref='predio', cascade='all, delete-orphan')
    colindantes  = db.relationship('Colindante',       backref='predio', cascade='all, delete-orphan')
    documentos   = db.relationship('Documento',        backref='predio', cascade='all, delete-orphan')


class Localizacion(db.Model):
    __tablename__ = 'localizacion'

    id_localizacion               = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_predio                     = db.Column(db.Integer, db.ForeignKey('predio.id_predio'), nullable=False)
    dpto_espacial                 = db.Column(db.String(1000))
    municipio_espacial            = db.Column(db.String(1000))
    vereda_espacial               = db.Column(db.String(1000))
    tiene_municipio_diferente_vur = db.Column(db.Boolean, default=False)
    municipio_vur                 = db.Column(db.String(1000))
    dpto_vur                      = db.Column(db.String(1000))
    vereda_vur                    = db.Column(db.String(1000))


class Catastral(db.Model):
    __tablename__ = 'catastral'

    id_catastral        = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_predio           = db.Column(db.Integer, db.ForeignKey('predio.id_predio'), nullable=False)
    numero_predial      = db.Column(db.String(1000))
    fmi_r1_r2           = db.Column(db.String(1000))
    nombre_predio_r1_r2 = db.Column(db.String(2000))
    propietario         = db.Column(db.String(2000))
    cedula_propietario  = db.Column(db.String(2000))
    area_terreno_ha     = db.Column(db.Float)
    area_terreno_ha_r1r2 = db.Column(db.Float)
    destino_economico   = db.Column(db.String(100))
    referencia_catastral= db.Column(db.String(5000))


class Juridica(db.Model):
    __tablename__ = 'juridica'

    id_juridica               = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_predio                 = db.Column(db.Integer, db.ForeignKey('predio.id_predio'), nullable=False)
    # Campos existentes
    documento_origen_inmueble  = db.Column(db.Text)
    area_juridica_ha           = db.Column(db.Float)
    nombre_predio_origen       = db.Column(db.Text)
    adjudicatario_inicial      = db.Column(db.Text)
    referencia_catastral_vur   = db.Column(db.String(2000))
    nupre                      = db.Column(db.String(100))
    documento_primera_anotacion = db.Column(db.Text)
    documento_ultima_anotacion = db.Column(db.Text)
    fecha_apertura_folio       = db.Column(db.String(20))
    tipo_instrumento           = db.Column(db.String(2000))
    fecha_instrumento          = db.Column(db.String(20))
    tipo_predio                = db.Column(db.String(20))
    estado_folio               = db.Column(db.String(2000))
    complementaciones          = db.Column(db.Text)        # Sin límite práctico
    cabida_linderos            = db.Column(db.Text)
    salvedades                 = db.Column(db.Text)
    vida_juridica              = db.Column(db.Text)  # Sin límite práctico
    nombre_predio_actual       = db.Column(db.String(2000))
    plano_juridica             = db.Column(db.String(50))  # PLANO / CARTERA / DESCRIPCION
    cas                        = db.Column(db.Text)


class Documento(db.Model):
    __tablename__ = 'documento'

    id_documento    = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_predio       = db.Column(db.Integer, db.ForeignKey('predio.id_predio'), nullable=False)
    tipo_documento  = db.Column(db.String(1000))
    nombre_archivo  = db.Column(db.String(2000), nullable=False)
    ruta_relativa   = db.Column(db.String(5000), nullable=False)
    tamaño_bytes    = db.Column(db.Float)
    fecha_carga     = db.Column(db.DateTime, default=datetime.utcnow)

class Colindante(db.Model):
    __tablename__ = 'colindante'

    id_colindante                     = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_predio                         = db.Column(db.Integer, db.ForeignKey('predio.id_predio'), nullable=False)
    numero_colindante                 = db.Column(db.Integer)
    cardinalidad                      = db.Column(db.String(500))
    fmi_colindante                    = db.Column(db.String(1000))
    numero_predial_colindante         = db.Column(db.String(1000))
    nombre_predio_colindante          = db.Column(db.String(2000))
    id_ant                            = db.Column(db.String(1000))
    tiene_levantamiento               = db.Column(db.Boolean, default=False)
    tipo_colindante                   = db.Column(db.String(2000))
    documento_origen_colindante       = db.Column(db.Text)
    propietario_inicial_colindante    = db.Column(db.Text)
    propietario_actual_vur_colindante = db.Column(db.Text)
    plano_colindante                  = db.Column(db.String(2000))
    necesario_validar_campo           = db.Column(db.Boolean, default=False)
    observacion_colindante            = db.Column(db.Text)
    colindante_primera_anotacion = db.Column(db.Text)
    colindante_ultima_anotacion = db.Column(db.Text)
    vereda_colindante = db.Column(db.String(1000))
    municipio_colindante = db.Column(db.String(1000))
    departamento_colindante = db.Column(db.String(1000))
    estado_validacion = db.Column(db.String(50))   # VALIDADO / NO_VERIFICADO / PRESUNTO_BALDIO
    fuente_validacion = db.Column(db.Text)
    requiere_verificacion_campo = db.Column(db.Boolean, default=False)
    es_presunto_baldio = db.Column(db.Boolean, default=False)
    es_derivado_colindante = db.Column(db.Boolean, default=False)
    id_predio_matriz_colindante = db.Column(db.String(100))
    tiene_derivadas_colindante = db.Column(db.Boolean, default=False)
    derivadas_ids_colindante = db.Column(db.String(5000))

    fecha_creacion                    = db.Column(db.DateTime, default=datetime.utcnow)

    documentos = db.relationship('DocumentoColindante', backref='colindante', cascade='all, delete-orphan')


class DocumentoColindante(db.Model):
    __tablename__ = 'documento_colindante'

    id_doc_colindante = db.Column(db.Integer, primary_key=True, autoincrement=True)
    id_colindante     = db.Column(db.Integer, db.ForeignKey('colindante.id_colindante'), nullable=False)
    tipo_documento    = db.Column(db.String(1000))
    nombre_archivo    = db.Column(db.String(2000), nullable=False)
    ruta_relativa     = db.Column(db.String(5000), nullable=False)
    tamaño_bytes      = db.Column(db.Float)
    fecha_carga       = db.Column(db.DateTime, default=datetime.utcnow)
