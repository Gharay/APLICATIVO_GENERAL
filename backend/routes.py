"""
backend/routes.py
Todas las rutas CRUD del Sistema de Información Catastral
"""
from flask import Blueprint, request, jsonify, current_app, send_file
from database.models import (
    db, Predio, Localizacion, Catastral, Juridica,
    Documento, Colindante, DocumentoColindante
)
from backend.export_word import exportar_word
from backend.export_excel import exportar_excel
from werkzeug.utils import secure_filename
import os
import traceback
from datetime import datetime
import re
import mimetypes

bp_api = Blueprint('api', __name__)

# ============================================================
# HELPERS DE CASTEO — usados en toda la función actualizar
# ============================================================
_VACIOS = (None, '', 'None', 'null', 'undefined')

def _to_float(val):
    """Convierte a float o retorna None si está vacío/inválido."""
    if val in _VACIOS:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def _to_int(val):
    """Convierte a int o retorna None si está vacío/inválido."""
    if val in _VACIOS:
        return None
    try:
        return int(val)
    except (ValueError, TypeError):
        return None

def _to_bool(val, default=False):
    """Convierte a bool desde string, int o bool."""
    if val in _VACIOS:
        return default
    if isinstance(val, bool):
        return val
    if isinstance(val, int):
        return bool(val)
    return str(val).lower() in ('true', '1', 'yes', 'sí', 'si')

def _to_str(val):
    """Retorna string limpio o None si está vacío."""
    if val in _VACIOS:
        return None
    return str(val).strip() or None


# ============================================================
# VALIDACIÓN Y RENOMBRADO DE ARCHIVOS
# ============================================================
EXTENSIONES_PERMITIDAS = {
    '.pdf', '.jpg', '.jpeg', '.tiff', '.tif',
    '.doc', '.docx', '.png', '.zip', '.xls', '.xlsx'
}

def _validar_archivo(archivo):
    if not archivo or not archivo.filename:
        return None, "No se envió archivo"
    ext = os.path.splitext(archivo.filename)[1].lower()
    if ext not in EXTENSIONES_PERMITIDAS:
        return None, (
            f"Extensión '{ext}' no permitida. "
            f"Extensiones válidas: {', '.join(sorted(EXTENSIONES_PERMITIDAS))}"
        )
    nombre_seguro = secure_filename(archivo.filename)
    if not nombre_seguro:
        nombre_seguro = f"archivo_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    return nombre_seguro, None

def _nombre_automatico(predio_id, predio_fmi, tipo_doc, nombre_original, carpeta):
    ext       = os.path.splitext(nombre_original)[1].lower()
    id_safe   = re.sub(r'[^\w\-]', '_', predio_id  or 'SIN_ID').strip('_')
    fmi_safe  = re.sub(r'[^\w\-]', '_', predio_fmi or 'SIN_FMI').strip('_')
    tipo_safe = re.sub(r'[^\w]',   '_', tipo_doc   or 'OTRO').strip('_')
    base      = re.sub(r'_+', '_', f"{id_safe}_{fmi_safe}_{tipo_safe}")
    nombre    = f"{base}{ext}"
    if not os.path.exists(os.path.join(carpeta, nombre)):
        return nombre
    contador = 1
    while True:
        nombre = f"{base}_{contador:03d}{ext}"
        if not os.path.exists(os.path.join(carpeta, nombre)):
            return nombre
        contador += 1


# ============================================================
# UTILIDADES DE RESPUESTA
# ============================================================
def response_success(data=None, message=""):
    return jsonify({'success': True, 'data': data, 'message': message}), 200

def response_error(error, status_code=400):
    return jsonify({'success': False, 'error': str(error)}), status_code


# ============================================================
# PREDIOS — CRUD
# ============================================================

@bp_api.route('/predios', methods=['GET'])
def obtener_predios():
    try:
        predios = db.session.query(Predio).all()
        datos = [{
            'id_predio':            p.id_predio,
            'id':                   p.id,
            'fmi':                  p.fmi,
            'propietario_vur':      p.propietario_vur,
            'nombre_predio_vur':    p.nombre_predio_vur,
            'area_vur':             float(p.area_vur) if p.area_vur else None,
            'area_final_gc':        float(p.area_final_gc) if p.area_final_gc else None,
            'tipo_gestion':         p.tipo_gestion,
            'estado':               p.estado,
            'fecha_creacion':       p.fecha_creacion.isoformat() if p.fecha_creacion else None
        } for p in predios]
        return response_success(data=datos)
    except Exception as e:
        traceback.print_exc()
        return response_error(f"Error al obtener predios: {str(e)}", 500)



@bp_api.route('/predios', methods=['POST'])
def crear_predio():
    try:
        datos = request.get_json()
        if not datos:
            return response_error("No se enviaron datos", 400)

        id_predio_str = datos.get('id', '').strip()
        fmi           = datos.get('fmi', '').strip()
        if not id_predio_str or not fmi:
            return response_error("ID y FMI son requeridos", 400)

        existe = db.session.query(Predio).filter(
            (Predio.id == id_predio_str) | (Predio.fmi == fmi)
        ).first()
        if existe:
            return response_error(
                f"Ya existe un predio con ID '{id_predio_str}' o FMI '{fmi}'", 400
            )

        nuevo_predio = Predio(id=id_predio_str, fmi=fmi, estado='ACTIVO')
        db.session.add(nuevo_predio)
        db.session.flush()

        os.makedirs(os.path.join(current_app.config['UPLOAD_FOLDER'], id_predio_str, 'predio'), exist_ok=True)
        os.makedirs(os.path.join(current_app.config['UPLOAD_FOLDER'], id_predio_str, 'colindantes'), exist_ok=True)

        db.session.add_all([
            Localizacion(id_predio=nuevo_predio.id_predio),
            Catastral(id_predio=nuevo_predio.id_predio),
            Juridica(id_predio=nuevo_predio.id_predio),
        ])
        db.session.commit()
        return jsonify({
            "success": True,
            "message": "Predio creado exitosamente",
            "data": {
                "id_predio": nuevo_predio.id_predio,
                "id":        nuevo_predio.id,
                "fmi":       nuevo_predio.fmi,
                "estado":    nuevo_predio.estado
            }
        }), 201
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al crear predio: {str(e)}", 500)


@bp_api.route('/predios/<int:id_predio>', methods=['GET'])
def obtener_predio(id_predio):
    try:
        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        localizacion = db.session.query(Localizacion).filter_by(id_predio=id_predio).first()
        catastral    = db.session.query(Catastral).filter_by(id_predio=id_predio).first()
        juridica     = db.session.query(Juridica).filter_by(id_predio=id_predio).first()
        colindantes  = db.session.query(Colindante).filter_by(id_predio=id_predio).order_by(Colindante.numero_colindante).all()
        documentos   = db.session.query(Documento).filter_by(id_predio=id_predio).all()

        datos = {
            'predio': {
                'id_predio':                        predio.id_predio,
                'id':                               predio.id,
                'fmi':                              predio.fmi,
                'propietario_vur':                  predio.propietario_vur,
                'numero_documento_propietario_vur': predio.numero_documento_propietario_vur,
                'nombre_predio_vur':                predio.nombre_predio_vur,
                'area_vur':                         float(predio.area_vur) if predio.area_vur else None,
                'area_final_gc':                    float(predio.area_final_gc) if predio.area_final_gc else None,
                 "tipo_gestion":                    predio.tipo_gestion,
                'es_derivado':                      predio.es_derivado,
                'id_predio_matriz':                 predio.id_predio_matriz,
                'tiene_derivadas':                  predio.tiene_derivadas,
                'derivadas_ids':                    predio.derivadas_ids,
                'proceso_englobe_desenglobe':       predio.proceso_englobe_desenglobe,
                'estado':                           predio.estado
            },
            'localizacion': {
                'dpto_espacial':               localizacion.dpto_espacial if localizacion else None,
                'municipio_espacial':          localizacion.municipio_espacial if localizacion else None,
                'vereda_espacial':             localizacion.vereda_espacial if localizacion else None,
                'tiene_municipio_diferente_vur': localizacion.tiene_municipio_diferente_vur if localizacion else False,
                'municipio_vur':               localizacion.municipio_vur if localizacion else None,
                'dpto_vur':                    localizacion.dpto_vur if localizacion else None,
                'vereda_vur':                  localizacion.vereda_vur if localizacion else None
            },
            'catastral': {
                'numero_predial':       catastral.numero_predial if catastral else None,
                'fmi_r1_r2':            catastral.fmi_r1_r2 if catastral else None,
                'nombre_predio_r1_r2':  catastral.nombre_predio_r1_r2 if catastral else None,
                'propietario':          catastral.propietario if catastral else None,
                'cedula_propietario':   catastral.cedula_propietario if catastral else None,
                'area_terreno_ha':      float(catastral.area_terreno_ha) if catastral and catastral.area_terreno_ha else None,
                'area_terreno_ha_r1r2': float(catastral.area_terreno_ha_r1r2) if catastral and catastral.area_terreno_ha_r1r2 else None,
                'destino_economico':    catastral.destino_economico if catastral else None,
                'referencia_catastral': catastral.referencia_catastral if catastral else None
            },
            'juridica': {
                'documento_origen_inmueble':  juridica.documento_origen_inmueble if juridica else None,
                'area_juridica_ha':           float(juridica.area_juridica_ha) if juridica and juridica.area_juridica_ha else None,
                'nombre_predio_origen':       juridica.nombre_predio_origen if juridica else None,
                'adjudicatario_inicial':      juridica.adjudicatario_inicial if juridica else None,
                'referencia_catastral_vur':   juridica.referencia_catastral_vur if juridica else None,
                'nupre':                      juridica.nupre if juridica else None,
                'documento_primera_anotacion': juridica.documento_primera_anotacion if juridica else None,
                'documento_ultima_anotacion': juridica.documento_ultima_anotacion if juridica else None,
                'fecha_apertura_folio':       juridica.fecha_apertura_folio if juridica else None,
                'tipo_instrumento':           juridica.tipo_instrumento if juridica else None,
                'fecha_instrumento':          juridica.fecha_instrumento if juridica else None,
                'tipo_predio':                juridica.tipo_predio if juridica else None,
                'estado_folio':               juridica.estado_folio if juridica else None,
                'complementaciones':          juridica.complementaciones if juridica else None,
                'cabida_linderos':            juridica.cabida_linderos if juridica else None,
                'salvedades':                 juridica.salvedades if juridica else None,
                'vida_juridica':              juridica.vida_juridica if juridica else None,
                'nombre_predio_actual':       juridica.nombre_predio_actual if juridica else None,
                'plano_juridica':             juridica.plano_juridica if juridica else None,
                'cas':                        juridica.cas if juridica else None,
            },
            'colindantes': [{
                'id_colindante':                     c.id_colindante,
                'numero_colindante':                 c.numero_colindante,
                'cardinalidad':                      c.cardinalidad,
                'fmi_colindante':                    c.fmi_colindante,
                'numero_predial_colindante':         c.numero_predial_colindante,
                'nombre_predio_colindante':          c.nombre_predio_colindante,
                'id_ant':                            c.id_ant,
                'tiene_levantamiento':               c.tiene_levantamiento,
                'tipo_colindante':                   c.tipo_colindante,
                'documento_origen_colindante':       c.documento_origen_colindante,
                'propietario_inicial_colindante':    c.propietario_inicial_colindante,
                'propietario_actual_vur_colindante': c.propietario_actual_vur_colindante,
                'plano_colindante':                  c.plano_colindante,
                'necesario_validar_campo':           c.necesario_validar_campo,
                'observacion_colindante':            c.observacion_colindante,
                'colindante_primera_anotacion': c.colindante_primera_anotacion,
                'colindante_ultima_anotacion': c.colindante_ultima_anotacion,
                'vereda_colindante': c.vereda_colindante,
                'municipio_colindante': c.municipio_colindante,
                'departamento_colindante': c.departamento_colindante,
                'estado_validacion': c.estado_validacion,
                'fuente_validacion': c.fuente_validacion,
                'requiere_verificacion_campo': c.requiere_verificacion_campo,
                'es_presunto_baldio': c.es_presunto_baldio,
                "es_derivado_colindante": c.es_derivado_colindante,
                "id_predio_matriz_colindante": c.id_predio_matriz_colindante,
                "tiene_derivadas_colindante": c.tiene_derivadas_colindante,
                "derivadas_ids_colindante": c.derivadas_ids_colindante,
                'documentos_colindante': [{
                    'id_doc_colindante': dc.id_doc_colindante,
                    'tipo_documento':    dc.tipo_documento,
                    'nombre_archivo':    dc.nombre_archivo,
                    'tamaño_bytes':      dc.tamaño_bytes,
                    'fecha_carga':       dc.fecha_carga.isoformat() if dc.fecha_carga else None
                } for dc in db.session.query(DocumentoColindante).filter_by(id_colindante=c.id_colindante).all()]
            } for c in colindantes],
            'documentos': [{
                'id_documento':   d.id_documento,
                'tipo_documento': d.tipo_documento,
                'nombre_archivo': d.nombre_archivo,
                'ruta_relativa':  d.ruta_relativa,
                'tamaño_bytes':   d.tamaño_bytes,
                'fecha_carga':    d.fecha_carga.isoformat() if d.fecha_carga else None
            } for d in documentos]
        }
        return response_success(data=datos)
    except Exception as e:
        traceback.print_exc()
        return response_error(f"Error al obtener predio: {str(e)}", 500)


@bp_api.route('/predios/<int:id_predio>', methods=['PUT'])
def actualizar_predio(id_predio):
    try:
        datos = request.get_json()
        if not datos:
            return response_error("No se enviaron datos", 400)

        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        # ── PREDIO ───────────────────────────────────────────────────────────
        if 'predio' in datos:
            p = datos['predio']
            predio.propietario_vur                  = _to_str(p.get('propietario_vur'))
            predio.numero_documento_propietario_vur = _to_str(p.get('numero_documento_propietario_vur'))
            predio.nombre_predio_vur                = _to_str(p.get('nombre_predio_vur'))
            predio.area_vur                         = _to_float(p.get('area_vur'))
            predio.area_final_gc                    = _to_float(p.get('area_final_gc'))
            predio.tipo_gestion                     = _to_str(p.get("tipo_gestion"))
            predio.es_derivado                      = _to_bool(p.get('es_derivado'))
            predio.tiene_derivadas                  = _to_bool(p.get('tiene_derivadas'))
            predio.derivadas_ids                    = _to_str(p.get('derivadas_ids'))
            predio.proceso_englobe_desenglobe       = _to_bool(p.get('proceso_englobe_desenglobe'))

            predio.id_predio_matriz = _to_str(p.get('id_predio_matriz'))

        # ── LOCALIZACIÓN ─────────────────────────────────────────────────────
        if 'localizacion' in datos:
            loc = db.session.query(Localizacion).filter_by(id_predio=id_predio).first()
            if not loc:
                loc = Localizacion(id_predio=id_predio)
                db.session.add(loc)
            l = datos['localizacion']
            loc.dpto_espacial                = _to_str(l.get('dpto_espacial'))
            loc.municipio_espacial           = _to_str(l.get('municipio_espacial'))
            loc.vereda_espacial              = _to_str(l.get('vereda_espacial'))
            loc.tiene_municipio_diferente_vur = _to_bool(l.get('tiene_municipio_diferente_vur'))
            loc.municipio_vur                = _to_str(l.get('municipio_vur'))
            loc.dpto_vur                     = _to_str(l.get('dpto_vur'))
            loc.vereda_vur                   = _to_str(l.get('vereda_vur'))

        # ── CATASTRAL ────────────────────────────────────────────────────────
        if 'catastral' in datos:
            cat = db.session.query(Catastral).filter_by(id_predio=id_predio).first()
            if not cat:
                cat = Catastral(id_predio=id_predio)
                db.session.add(cat)
            c = datos['catastral']
            cat.numero_predial       = _to_str(c.get('numero_predial'))
            cat.fmi_r1_r2            = _to_str(c.get('fmi_r1_r2'))
            cat.nombre_predio_r1_r2  = _to_str(c.get('nombre_predio_r1_r2'))
            cat.propietario          = _to_str(c.get('propietario'))
            cat.cedula_propietario   = _to_str(c.get('cedula_propietario'))
            cat.area_terreno_ha      = _to_float(c.get('area_terreno_ha'))   # ← Float
            cat.area_terreno_ha_r1r2 = _to_float(c.get('area_terreno_ha_r1r2'))   # ← Float
            cat.destino_economico    = _to_str(c.get('destino_economico'))
            cat.referencia_catastral = _to_str(c.get('referencia_catastral'))

        # ── JURÍDICA ─────────────────────────────────────────────────────────
        if 'juridica' in datos:
            jur = db.session.query(Juridica).filter_by(id_predio=id_predio).first()
            if not jur:
                jur = Juridica(id_predio=id_predio)
                db.session.add(jur)
            j = datos['juridica']
            jur.documento_origen_inmueble  = _to_str(j.get('documento_origen_inmueble'))
            jur.area_juridica_ha           = _to_float(j.get('area_juridica_ha'))   # ← Float
            jur.nombre_predio_origen       = _to_str(j.get('nombre_predio_origen'))
            jur.adjudicatario_inicial      = _to_str(j.get('adjudicatario_inicial'))
            jur.documento_ultima_anotacion = _to_str(j.get('documento_ultima_anotacion'))
            jur.fecha_apertura_folio       = _to_str(j.get('fecha_apertura_folio'))
            jur.tipo_instrumento           = _to_str(j.get('tipo_instrumento'))
            jur.fecha_instrumento          = _to_str(j.get('fecha_instrumento'))
            jur.tipo_predio                = _to_str(j.get('tipo_predio'))
            jur.estado_folio               = _to_str(j.get('estado_folio'))
            jur.complementaciones          = _to_str(j.get('complementaciones'))
            jur.cabida_linderos            = _to_str(j.get('cabida_linderos'))
            jur.nombre_predio_actual       = _to_str(j.get('nombre_predio_actual'))
            jur.plano_juridica             = _to_str(j.get('plano_juridica'))
            jur.referencia_catastral_vur   = _to_str(j.get('referencia_catastral_vur'))
            jur.nupre                      = _to_str(j.get('nupre'))
            jur.documento_primera_anotacion = _to_str(j.get('documento_primera_anotacion'))
            jur.salvedades                 = _to_str(j.get('salvedades'))
            jur.vida_juridica              = _to_str(j.get('vida_juridica'))
            jur.cas                        = _to_str(j.get('cas'))

        # ── COLINDANTES (upsert) ─────────────────────────────────────────────
        if 'colindantes' in datos:
            col_existentes = {
                c.id_colindante: c
                for c in db.session.query(Colindante).filter_by(id_predio=id_predio).all()
            }
            ids_recibidos = set()

            for col_data in datos['colindantes']:
                id_col = _to_int(col_data.get('id_colindante'))
                if id_col and id_col in col_existentes:
                    col = col_existentes[id_col]
                    ids_recibidos.add(id_col)
                else:
                    col = Colindante(id_predio=id_predio)
                    db.session.add(col)

                col.numero_colindante               = _to_int(col_data.get('numero_colindante'))   # ← Integer
                col.cardinalidad                    = _to_str(col_data.get('cardinalidad'))
                col.fmi_colindante                  = _to_str(col_data.get('fmi_colindante'))
                col.numero_predial_colindante       = _to_str(col_data.get('numero_predial_colindante'))
                col.nombre_predio_colindante        = _to_str(col_data.get('nombre_predio_colindante'))
                col.id_ant                          = _to_str(col_data.get('id_ant'))
                col.tiene_levantamiento             = _to_bool(col_data.get('tiene_levantamiento'))
                col.tipo_colindante                 = _to_str(col_data.get('tipo_colindante'))
                col.documento_origen_colindante     = _to_str(col_data.get('documento_origen_colindante'))
                col.propietario_inicial_colindante  = _to_str(col_data.get('propietario_inicial_colindante'))
                col.propietario_actual_vur_colindante = _to_str(col_data.get('propietario_actual_vur_colindante'))
                col.plano_colindante                = _to_str(col_data.get('plano_colindante'))
                col.necesario_validar_campo         = _to_bool(col_data.get('necesario_validar_campo'))
                col.observacion_colindante          = _to_str(col_data.get('observacion_colindante'))
                col.colindante_primera_anotacion      =_to_str(col_data.get('colindante_primera_anotacion'))  # ← NUEVO
                col.colindante_ultima_anotacion       = _to_str(col_data.get('colindante_ultima_anotacion')) 
                col.vereda_colindante             = _to_str(col_data.get('vereda_colindante'))
                col.municipio_colindante          = _to_str(col_data.get('municipio_colindante'))
                col.departamento_colindante       = _to_str(col_data.get('departamento_colindante'))
                col.estado_validacion             = _to_str(col_data.get('estado_validacion'))
                col.fuente_validacion             = _to_str(col_data.get('fuente_validacion'))
                col.requiere_verificacion_campo   = _to_bool(col_data.get('requiere_verificacion_campo'))
                col.es_presunto_baldio            = _to_bool(col_data.get('es_presunto_baldio'))
                col.es_derivado_colindante = _to_bool(col_data.get('es_derivado_colindante'))
                col.id_predio_matriz_colindante = _to_str(col_data.get('id_predio_matriz_colindante'))
                col.tiene_derivadas_colindante = _to_bool(col_data.get('tiene_derivadas_colindante'))
                col.derivadas_ids_colindante = _to_str(col_data.get('derivadas_ids_colindante'))

            for id_col, col in col_existentes.items():
                if id_col not in ids_recibidos:
                    # Eliminar archivos físicos de documentos del colindante
                    docs_col = db.session.query(DocumentoColindante).filter_by(id_colindante=id_col).all()
                    for dc in docs_col:
                        if dc.ruta_relativa and os.path.exists(dc.ruta_relativa):
                            try:
                                os.remove(dc.ruta_relativa)
                            except:
                                pass
                    db.session.delete(col)

        db.session.commit()
        print(f"✓ Predio {id_predio} actualizado")
        datos_actualizados = _construir_datos_completos(id_predio)
        return response_success(data=datos_actualizados, message="Predio actualizado exitosamente")

    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al actualizar predio: {str(e)}", 500)


@bp_api.route('/predios/<int:id_predio>', methods=['DELETE'])
def eliminar_predio(id_predio):
    try:
        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        db.session.query(Juridica).filter_by(id_predio=id_predio).delete()
        db.session.query(Catastral).filter_by(id_predio=id_predio).delete()
        db.session.query(Localizacion).filter_by(id_predio=id_predio).delete()

        colindantes = db.session.query(Colindante).filter_by(id_predio=id_predio).all()
        for c in colindantes:
            docs_col = db.session.query(DocumentoColindante).filter_by(id_colindante=c.id_colindante).all()
            for dc in docs_col:
                if dc.ruta_relativa and os.path.exists(dc.ruta_relativa):
                    try:
                        os.remove(dc.ruta_relativa)
                    except:
                        pass
            db.session.query(DocumentoColindante).filter_by(id_colindante=c.id_colindante).delete()
        db.session.query(Colindante).filter_by(id_predio=id_predio).delete()

        documentos = db.session.query(Documento).filter_by(id_predio=id_predio).all()
        for doc in documentos:
            if doc.ruta_relativa and os.path.exists(doc.ruta_relativa):
                try:
                    os.remove(doc.ruta_relativa)
                except:
                    pass
        db.session.query(Documento).filter_by(id_predio=id_predio).delete()
        db.session.delete(predio)
        db.session.commit()
        return response_success(message="Predio eliminado exitosamente")
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al eliminar predio: {str(e)}", 500)


# ============================================================
# COLINDANTES
# ============================================================

@bp_api.route('/predios/<int:id_predio>/colindantes', methods=['POST'])
def crear_colindante(id_predio):
    try:
        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        datos = request.get_json()
        colindante = Colindante(
            id_predio=id_predio,
            numero_colindante             = _to_int(datos.get('numero_colindante')),
            cardinalidad                  = _to_str(datos.get('cardinalidad')),
            fmi_colindante                = _to_str(datos.get('fmi_colindante')),
            numero_predial_colindante     = _to_str(datos.get('numero_predial_colindante')),
            nombre_predio_colindante      = _to_str(datos.get('nombre_predio_colindante')),
            id_ant                        = _to_str(datos.get('id_ant')),
            tiene_levantamiento           = _to_bool(datos.get('tiene_levantamiento')),
            tipo_colindante               = _to_str(datos.get('tipo_colindante')),
            documento_origen_colindante   = _to_str(datos.get('documento_origen_colindante')),
            propietario_inicial_colindante= _to_str(datos.get('propietario_inicial_colindante')),
            propietario_actual_vur_colindante = _to_str(datos.get('propietario_actual_vur_colindante')),
            plano_colindante              = _to_str(datos.get('plano_colindante')),
            necesario_validar_campo       = _to_bool(datos.get('necesario_validar_campo')),
            observacion_colindante        = _to_str(datos.get('observacion_colindante')),
            colindante_primera_anotacion  = _to_str(datos.get('colindante_primera_anotacion')),
            colindante_ultima_anotacion   = _to_str(datos.get('colindante_ultima_anotacion')),
            vereda_colindante           = _to_str(datos.get('vereda_colindante')),
            municipio_colindante        = _to_str(datos.get('municipio_colindante')),
            departamento_colindante     = _to_str(datos.get('departamento_colindante')),
            estado_validacion           = _to_str(datos.get('estado_validacion')),
            fuente_validacion           = _to_str(datos.get('fuente_validacion')),
            requiere_verificacion_campo = _to_bool(datos.get('requiere_verificacion_campo')),
            es_presunto_baldio          = _to_bool(datos.get('es_presunto_baldio')),
            es_derivado_colindante = _to_bool(datos.get('es_derivado_colindante')),
            id_predio_matriz_colindante = _to_str(datos.get('id_predio_matriz_colindante')),
            tiene_derivadas_colindante = _to_bool(datos.get('tiene_derivadas_colindante')),
            derivadas_ids_colindante = _to_str(datos.get('derivadas_ids_colindante')),
        )
        db.session.add(colindante)
        db.session.commit()
        return response_success(
            data={'id_colindante': colindante.id_colindante},
            message="Colindante creado"
        ), 201
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al crear colindante: {str(e)}", 500)


# ============================================================
# DOCUMENTOS DEL PREDIO
# ============================================================

@bp_api.route('/predios/<int:id_predio>/documentos', methods=['POST'])
def subir_documento(id_predio):
    try:
        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        archivo  = request.files.get('archivo')
        tipo_doc = request.form.get('tipo_documento', 'OTRO')

        nombre_seguro, error = _validar_archivo(archivo)
        if error:
            return response_error(error, 400)

        carpeta = os.path.join(current_app.config['UPLOAD_FOLDER'], predio.id, 'predio')
        os.makedirs(carpeta, exist_ok=True)
        nombre_final = _nombre_automatico(predio.id, predio.fmi, tipo_doc, nombre_seguro, carpeta)
        ruta = os.path.join(carpeta, nombre_final)
        archivo.save(ruta)

        doc = Documento(
            id_predio=id_predio,
            tipo_documento=tipo_doc,
            nombre_archivo=nombre_final,
            ruta_relativa=ruta,
            tamaño_bytes=os.path.getsize(ruta)
        )
        db.session.add(doc)
        db.session.commit()
        return response_success(
            data={'id_documento': doc.id_documento, 'nombre': doc.nombre_archivo},
            message="Documento guardado"
        )
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al subir documento: {str(e)}", 500)


@bp_api.route('/documentos/<int:id_documento>', methods=['DELETE'])
def eliminar_documento(id_documento):
    try:
        documento = db.session.query(Documento).filter_by(id_documento=id_documento).first()
        if not documento:
            return response_error("Documento no encontrado", 404)
        if documento.ruta_relativa and os.path.exists(documento.ruta_relativa):
            try:
                os.remove(documento.ruta_relativa)
            except:
                pass
        db.session.delete(documento)
        db.session.commit()
        return response_success(message="Documento eliminado")
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al eliminar documento: {str(e)}", 500)


@bp_api.route('/documentos/<int:id_documento>/ver', methods=['GET'])
def ver_documento(id_documento):
    try:
        doc = db.session.query(Documento).filter_by(id_documento=id_documento).first()
        if not doc:
            return response_error("Documento no encontrado", 404)
        if not doc.ruta_relativa or not os.path.exists(doc.ruta_relativa):
            return response_error("Archivo no encontrado en disco", 404)
        mime, _ = mimetypes.guess_type(doc.ruta_relativa)
        mime = mime or 'application/octet-stream'
        return send_file(doc.ruta_relativa, mimetype=mime,
                         as_attachment=False, download_name=doc.nombre_archivo)
    except Exception as e:
        traceback.print_exc()
        return response_error(str(e), 500)


# ============================================================
# DOCUMENTOS DEL COLINDANTE
# ============================================================

@bp_api.route('/predios/<int:id_predio>/colindantes/<int:id_colindante>/documentos', methods=['POST'])
def subir_documento_colindante(id_predio, id_colindante):
    try:
        predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
        if not predio:
            return response_error("Predio no encontrado", 404)

        colindante = db.session.query(Colindante).filter_by(id_colindante=id_colindante).first()
        if not colindante:
            return response_error("Colindante no encontrado", 404)

        archivo  = request.files.get('archivo')
        tipo_doc = request.form.get('tipo_documento', 'OTRO')

        nombre_seguro, error = _validar_archivo(archivo)
        if error:
            return response_error(error, 400)

        fmi_col  = colindante.fmi_colindante or f'COL_{id_colindante}'
        fmi_safe = re.sub(r'[^\w\-]', '_', fmi_col)
        carpeta  = os.path.join(current_app.config['UPLOAD_FOLDER'],
                                predio.id, 'colindantes', fmi_safe)
        os.makedirs(carpeta, exist_ok=True)
        nombre_final = _nombre_automatico(predio.id, fmi_col, tipo_doc, nombre_seguro, carpeta)
        ruta = os.path.join(carpeta, nombre_final)
        archivo.save(ruta)

        doc = DocumentoColindante(
            id_colindante=id_colindante,
            tipo_documento=tipo_doc,
            nombre_archivo=nombre_final,
            ruta_relativa=ruta,
            tamaño_bytes=os.path.getsize(ruta)
        )
        db.session.add(doc)
        db.session.commit()
        return response_success(
            data={'id_doc_colindante': doc.id_doc_colindante, 'nombre': doc.nombre_archivo},
            message="Documento guardado"
        )
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(str(e), 500)


@bp_api.route('/documentos-colindante/<int:iddoc>', methods=['DELETE'])
def eliminar_documento_colindante(iddoc):
    try:
        doc = db.session.query(DocumentoColindante).filter_by(id_doc_colindante=iddoc).first()
        if not doc:
            return response_error("Documento no encontrado", 404)

        if doc.ruta_relativa and os.path.exists(doc.ruta_relativa):
            try:
                os.remove(doc.ruta_relativa)
            except:
                pass

        db.session.delete(doc)
        db.session.commit()
        return response_success(message="Documento eliminado")
    except Exception as e:
        db.session.rollback()
        traceback.print_exc()
        return response_error(f"Error al eliminar documento: {str(e)}", 500)


@bp_api.route('/documentos-colindante/<int:id_doc>/ver', methods=['GET'])
def ver_documento_colindante(id_doc):
    try:
        doc = db.session.query(DocumentoColindante).filter_by(id_doc_colindante=id_doc).first()
        if not doc:
            return response_error("Documento no encontrado", 404)
        if not doc.ruta_relativa or not os.path.exists(doc.ruta_relativa):
            return response_error("Archivo no encontrado en disco", 404)
        mime, _ = mimetypes.guess_type(doc.ruta_relativa)
        mime = mime or 'application/octet-stream'
        return send_file(doc.ruta_relativa, mimetype=mime,
                         as_attachment=False, download_name=doc.nombre_archivo)
    except Exception as e:
        traceback.print_exc()
        return response_error(str(e), 500)


# ============================================================
# BASE EXPORTACIÓN
# ============================================================

def _construir_datos_completos(id_predio):
    predio = db.session.query(Predio).filter_by(id_predio=id_predio).first()
    if not predio:
        return None

    localizacion = db.session.query(Localizacion).filter_by(id_predio=id_predio).first()
    catastral    = db.session.query(Catastral).filter_by(id_predio=id_predio).first()
    juridica     = db.session.query(Juridica).filter_by(id_predio=id_predio).first()
    colindantes  = db.session.query(Colindante).filter_by(id_predio=id_predio)\
                       .order_by(Colindante.numero_colindante).all()
    documentos   = db.session.query(Documento).filter_by(id_predio=id_predio).all()

    return {
        'predio': {
            'id_predio':                        predio.id_predio,
            'id':                               predio.id,
            'fmi':                              predio.fmi,
            'propietario_vur':                  predio.propietario_vur,
            'numero_documento_propietario_vur': predio.numero_documento_propietario_vur,
            'nombre_predio_vur':                predio.nombre_predio_vur,
            "tipo_gestion":                     predio.tipo_gestion,
            'area_vur':                         float(predio.area_vur) if predio.area_vur else None,
            'area_final_gc':                    float(predio.area_final_gc) if predio.area_final_gc else None,
            'es_derivado':                      predio.es_derivado,
            'id_predio_matriz':                 predio.id_predio_matriz,
            'tiene_derivadas':                  predio.tiene_derivadas,
            'derivadas_ids':                    predio.derivadas_ids,
            'proceso_englobe_desenglobe':       predio.proceso_englobe_desenglobe,
            'estado':                           predio.estado,
            'fecha_creacion':                   predio.fecha_creacion.strftime('%d/%m/%Y') if predio.fecha_creacion else 'N/A'
        },
        'localizacion': {
            'dpto_espacial':               localizacion.dpto_espacial if localizacion else None,
            'municipio_espacial':          localizacion.municipio_espacial if localizacion else None,
            'vereda_espacial':             localizacion.vereda_espacial if localizacion else None,
            'tiene_municipio_diferente_vur': localizacion.tiene_municipio_diferente_vur if localizacion else False,
            'municipio_vur':               localizacion.municipio_vur if localizacion else None,
            'dpto_vur':                    localizacion.dpto_vur if localizacion else None,
            'vereda_vur':                  localizacion.vereda_vur if localizacion else None,
        },
        'catastral': {
            'numero_predial':       catastral.numero_predial if catastral else None,
            'fmi_r1_r2':            catastral.fmi_r1_r2 if catastral else None,
            'nombre_predio_r1_r2':  catastral.nombre_predio_r1_r2 if catastral else None,
            'propietario':          catastral.propietario if catastral else None,
            'cedula_propietario':   catastral.cedula_propietario if catastral else None,
            'area_terreno_ha':      float(catastral.area_terreno_ha) if catastral and catastral.area_terreno_ha else None,
            'area_terreno_ha_r1r2': float(catastral.area_terreno_ha_r1r2) if catastral and catastral.area_terreno_ha_r1r2 else None,
            'destino_economico':    catastral.destino_economico if catastral else None,
            'referencia_catastral': catastral.referencia_catastral if catastral else None,
        },
        'juridica': {
            'documento_origen_inmueble':    juridica.documento_origen_inmueble if juridica else None,
            'area_juridica_ha':             float(juridica.area_juridica_ha) if juridica and juridica.area_juridica_ha else None,
            'nombre_predio_origen':         juridica.nombre_predio_origen if juridica else None,
            'adjudicatario_inicial':        juridica.adjudicatario_inicial if juridica else None,
            'referencia_catastral_vur':     juridica.referencia_catastral_vur if juridica else None,  # ← NUEVO
            'nupre':                        juridica.nupre if juridica else None,                      # ← NUEVO
            'documento_primera_anotacion':  juridica.documento_primera_anotacion if juridica else None,# ← NUEVO
            'documento_ultima_anotacion':   juridica.documento_ultima_anotacion if juridica else None,
            'fecha_apertura_folio':         juridica.fecha_apertura_folio if juridica else None,
            'tipo_instrumento':             juridica.tipo_instrumento if juridica else None,
            'fecha_instrumento':            juridica.fecha_instrumento if juridica else None,
            'tipo_predio':                  juridica.tipo_predio if juridica else None,
            'estado_folio':                 juridica.estado_folio if juridica else None,
            'complementaciones':            juridica.complementaciones if juridica else None,
            'cabida_linderos':              juridica.cabida_linderos if juridica else None,
            'salvedades':                   juridica.salvedades if juridica else None,                  # ← NUEVO
            'vida_juridica':                juridica.vida_juridica if juridica else None,               # ← NUEVO
            'nombre_predio_actual':         juridica.nombre_predio_actual if juridica else None,
            'plano_juridica':               juridica.plano_juridica if juridica else None,
            'cas' :                         juridica.cas if juridica else None,                                    # ← NUEVO
        },
        'colindantes': [{
            'id_colindante': c.id_colindante,                          # ← FALTABA
            'numero_colindante': c.numero_colindante,
            'cardinalidad': c.cardinalidad,
            'fmi_colindante': c.fmi_colindante,
            'numero_predial_colindante': c.numero_predial_colindante,
            'nombre_predio_colindante': c.nombre_predio_colindante,
            'id_ant': c.id_ant,
            'tiene_levantamiento': c.tiene_levantamiento,
            'tipo_colindante': c.tipo_colindante,
            'documento_origen_colindante': c.documento_origen_colindante,
            'propietario_inicial_colindante': c.propietario_inicial_colindante,
            'propietario_actual_vur_colindante': c.propietario_actual_vur_colindante,
            'plano_colindante': c.plano_colindante,
            'necesario_validar_campo': c.necesario_validar_campo,
            'observacion_colindante': c.observacion_colindante,
            'colindante_primera_anotacion': c.colindante_primera_anotacion,
            'colindante_ultima_anotacion': c.colindante_ultima_anotacion,
            'vereda_colindante': c.vereda_colindante,
            'municipio_colindante': c.municipio_colindante,
            'departamento_colindante': c.departamento_colindante,
            'estado_validacion': c.estado_validacion,
            'fuente_validacion': c.fuente_validacion,
            'requiere_verificacion_campo': c.requiere_verificacion_campo,
            'es_presunto_baldio': c.es_presunto_baldio,
            'es_derivado_colindante': c.es_derivado_colindante,
            'id_predio_matriz_colindante': c.id_predio_matriz_colindante,
            'tiene_derivadas_colindante': c.tiene_derivadas_colindante,
            'derivadas_ids_colindante': c.derivadas_ids_colindante,
            'documentos_colindante': [{
                'id_doc_colindante': dc.id_doc_colindante,
                'tipo_documento': dc.tipo_documento,
                'nombre_archivo': dc.nombre_archivo,
                'tamaño_bytes': dc.tamaño_bytes,
                'fecha_carga': dc.fecha_carga.isoformat() if dc.fecha_carga else None
            } for dc in db.session.query(DocumentoColindante)
                                .filter_by(id_colindante=c.id_colindante).all()]
        } for c in colindantes],
        'documentos': [{
            'id_documento':   d.id_documento,
            'tipo_documento': d.tipo_documento,
            'nombre_archivo': d.nombre_archivo,
            'tamaño_bytes':   d.tamaño_bytes,
            'fecha_carga':    d.fecha_carga.strftime('%d/%m/%Y') if d.fecha_carga else 'N/A',
        } for d in documentos],
    }


# ============================================================
# EXPORTACIÓN
# ============================================================

@bp_api.route('/predios/<int:id_predio>/export-word', methods=['GET'])
def exportar_predio_word(id_predio):
    try:
        datos = _construir_datos_completos(id_predio)
        if not datos:
            return response_error("Predio no encontrado", 404)
        buffer = exportar_word(datos, current_app.config['UPLOAD_FOLDER'])
        nombre = f"Informe_{datos['predio']['id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        return send_file(
            buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre
        )
    except Exception as e:
        traceback.print_exc()
        return response_error(f"Error al exportar a Word: {str(e)}", 500)


@bp_api.route('/predios/<int:id_predio>/export-excel', methods=['GET'])
def exportar_predio_excel(id_predio):
    try:
        datos = _construir_datos_completos(id_predio)
        if not datos:
            return response_error("Predio no encontrado", 404)
        buffer = exportar_excel(datos)
        nombre = f"Ficha_{datos['predio']['id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nombre
        )
    except Exception as e:
        traceback.print_exc()
        return response_error(f"Error al exportar a Excel: {str(e)}", 500)


# ============================================================
# HEALTH CHECK
# ============================================================

@bp_api.route('/health', methods=['GET'])
def health_check():
    return response_success(data={'status': 'OK'})
