/**
 * static/js/ajax_handlers.js
 */

const BASE_URL = '/api';
let idPredioActual = null;

// ===== UTILIDADES =====
function getVal(id) {
    const el = document.getElementById(id) || document.querySelector(`[name="${id}"]`);
    return el ? el.value.trim() : null;
}
function getBool(id) {
    const val = getVal(id);
    return val === '1' || val === 'true' || val === 'SI';
}
function setInputValue(id, valor) {
    const el = document.getElementById(id) || document.querySelector(`[name="${id}"]`);
    if (el && valor !== null && valor !== undefined) el.value = valor;
}
function setSelectValue(id, valor) {
    const el = document.getElementById(id) || document.querySelector(`[name="${id}"]`);
    if (el && valor !== null && valor !== undefined) el.value = String(valor);
}

// ===== FUNCIONES  =====

function _esVacio(el) {
  if (!el) return true;
  if (el.tagName === 'SELECT') {
    const v = (el.value || '').trim();
    return v === '' || v === '-- Seleccione --';
  }
  const v = (el.value || '').trim();
  return v === '';
}

function marcarVaciosEnFormulario(root = document) {
  const campos = root.querySelectorAll('input, select, textarea');
  let vacios = 0;

  campos.forEach(el => {
    if (el.type === 'file' || el.disabled || el.readOnly) return;

    const vacio = _esVacio(el);
    el.classList.toggle('campo-vacio', vacio);
    if (vacio) vacios += 1;
  });

  return vacios;
}


// ===== INICIALIZACIÓN =====
document.addEventListener('DOMContentLoaded', function () {
    const pathMatch = window.location.pathname.match(/\/editar\/(\d+)/);
    if (pathMatch) {
        idPredioActual = parseInt(pathMatch[1]);
        cargarDatosPredio();

        const btnGuardar = document.getElementById('btnGuardar');
        if (btnGuardar) btnGuardar.addEventListener('click', guardarPredio);

        const btnWord = document.getElementById('btnExportarWord');
        if (btnWord) btnWord.addEventListener('click', () => {
            window.location.href = `${BASE_URL}/predios/${idPredioActual}/export-word`;
        });
        const btnExcel = document.getElementById('btnExportarExcel');
        if (btnExcel) btnExcel.addEventListener('click', () => {
            window.location.href = `${BASE_URL}/predios/${idPredioActual}/export-excel`;
        });
    }

    window.toggleMatriz = function () {
        const g = document.getElementById('grupoMatriz');
        if (g) g.style.display = getVal('predio_es_derivado') === '1' ? 'block' : 'none';
    };
    window.toggleDerivadas = function () {
        const g = document.getElementById('grupoDerivadas');
        if (g) g.style.display = getVal('predio_tiene_derivadas') === '1' ? 'block' : 'none';
    };

    window.toggleMunicipioVUR = function () {
        const g = document.getElementById('grupoMunicipioVUR');
        if (g) g.style.display = getVal('loc_municipio_diferente') === '1' ? 'block' : 'none';
    };
    window.toggleFajaParalela = function () {
        const g = document.getElementById('grupoFaja');
        if (g) g.style.display = getVal('ana_tiene_faja') === '1' ? 'block' : 'none';
    };

    const selectCompra = document.querySelector('select[name="ana_compra_tipo"]');
    if (selectCompra) {
        selectCompra.addEventListener('change', function () {
            const esParcial = this.value === 'PARCIAL';
            const gA = document.getElementById('grupoAreaCompra');
            const gM = document.getElementById('grupoMotivoCompra');
            if (gA) gA.style.display = esParcial ? 'block' : 'none';
            if (gM) gM.style.display = esParcial ? 'block' : 'none';
        });
    }
});

// ===== CARGAR DATOS =====
function cargarDatosPredio() {
    fetch(`${BASE_URL}/predios/${idPredioActual}`)
        .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
        .then(data => {
            if (data.success) llenarFormulario(data.data);
            else alert('Error al cargar predio: ' + data.error);
        })
        .catch(err => alert('Error de conexión: ' + err.message));
}

// ===== LLENAR FORMULARIO =====
function llenarFormulario(datos) {
    const span = document.getElementById('predioID');
    if (span && datos.predio) span.textContent = datos.predio.id || 'N/A';

    if (datos.predio) {
        const p = datos.predio;
        setInputValue('predio_id', p.id);
        setInputValue('predio_fmi', p.fmi);
        setInputValue('predio_propietario_vur', p.propietario_vur);
        setInputValue('predio_num_doc_propietario', p.numero_documento_propietario_vur);
        setInputValue('predio_nombre_vur', p.nombre_predio_vur);
        setSelectValue('predio_tipo_gestion', p.tipo_gestion);
        setInputValue('predio_area_vur', p.area_vur);
        setInputValue('predio_area_final_gc', p.area_final_gc);
        setSelectValue('predio_es_derivado', p.es_derivado ? '1' : '0');
        setInputValue('predio_id_matriz', p.id_predio_matriz);
        setSelectValue('predio_tiene_derivadas', p.tiene_derivadas ? '1' : '0');
        setInputValue('predio_derivadas_ids', p.derivadas_ids);
        setSelectValue('predio_englobe', p.proceso_englobe_desenglobe ? '1' : '0');
        window.toggleMatriz();
        window.toggleDerivadas();
    }

    if (datos.localizacion) {
        const loc = datos.localizacion;
        setInputValue('loc_dpto_espacial', loc.dpto_espacial);
        setInputValue('loc_municipio_espacial', loc.municipio_espacial);
        setInputValue('loc_vereda_espacial', loc.vereda_espacial);
        setSelectValue('loc_municipio_diferente', loc.tiene_municipio_diferente_vur ? '1' : '0');
        setInputValue('loc_municipio_vur', loc.municipio_vur);
        setInputValue('loc_dpto_vur', loc.dpto_vur);
        setInputValue('loc_vereda_vur', loc.vereda_vur);
        window.toggleMunicipioVUR();
    }

    if (datos.catastral) {
        const cat = datos.catastral;
        setInputValue('cat_numero_predial', cat.numero_predial);
        setInputValue('cat_fmi_r1_r2', cat.fmi_r1_r2);
        setInputValue('cat_nombre_r1_r2', cat.nombre_predio_r1_r2);
        setInputValue('cat_propietario', cat.propietario);
        setInputValue('cat_cedula', cat.cedula_propietario);
        setInputValue('cat_area_terreno', cat.area_terreno_ha);
        setInputValue('cat_area_terreno_r1r2', cat.area_terreno_ha_r1r2);
        setInputValue('cat_destino', cat.destino_economico);
        setInputValue('cat_referencia', cat.referencia_catastral);
    }

    if (datos.juridica) {
        const jur = datos.juridica;
        setInputValue('jur_documento_origen', jur.documento_origen_inmueble);
        setInputValue('jur_area_ha', jur.area_juridica_ha);
        setInputValue('jur_nombre_origen', jur.nombre_predio_origen);
        setInputValue('jur_adjudicatario', jur.adjudicatario_inicial);
        setInputValue('jur_referencia_catastral_vur', jur.referencia_catastral_vur);
        setInputValue('jur_nupre', jur.nupre);
        setInputValue('jur_salvedades', jur.salvedades);
        setInputValue('jur_vida_juridica', jur.vida_juridica);
        setInputValue('jur_documento_primera_anotacion', jur.documento_primera_anotacion);
        setInputValue('jur_documento_ultima_anotacion', jur.documento_ultima_anotacion);
        setInputValue('jur_doc_ultima_anotacion', jur.documento_ultima_anotacion);
        setInputValue('jur_fecha_apertura_folio', jur.fecha_apertura_folio);
        setInputValue('jur_tipo_instrumento', jur.tipo_instrumento);
        setInputValue('jur_fecha_instrumento', jur.fecha_instrumento);
        setInputValue('jur_tipo_predio', jur.tipo_predio);
        setInputValue('jur_estado_folio', jur.estado_folio);
        setInputValue('jur_complementaciones', jur.complementaciones);
        setInputValue('jur_cabida_linderos', jur.cabida_linderos);
        setInputValue('jur_nombre_predio_actual', jur.nombre_predio_actual);
        setSelectValue('jur_plano', jur.plano_juridica);
        setInputValue('jur_cas', jur.cas);
    }

    mostrarDocumentos(datos.documentos || []);

    // COLINDANTES — se pasan los datos directamente para poblar en el mismo paso
    const colindantes = datos.colindantes || [];
    const numInput = document.getElementById('numColindantes');
    if (numInput) numInput.value = colindantes.length;
    generarColindantes(colindantes.length, colindantes);
    marcarVaciosEnFormulario();
}

// ===== GUARDAR PREDIO =====
function guardarPredio() {

    const numActual = parseInt(document.getElementById('numColindantes')?.value) || 0;
    const numEnDOM = document.querySelectorAll('#contenedorColindantes [data-id-colindante]').length;
    if (numActual < numEnDOM) {
        const diff = numEnDOM - numActual;
        if (!confirm(`⚠️ Vas a eliminar ${diff} colindante(s) con todos sus documentos adjuntos. ¿Continuar?`)) return;
    }

    const vacios = marcarVaciosEnFormulario();

    // Opcional: mostrar contador, pero NO bloquear
    const msg = document.getElementById('mensajeError');
    if (msg) {
        if (vacios > 0) {
        msg.textContent = `Hay ${vacios} campos vacíos (marcados en rojo).`;
        msg.style.display = 'block';
        setTimeout(() => msg.style.display = 'none', 2500);
        } else {
        msg.style.display = 'none';
        }
    }

    if (!idPredioActual) { alert('Error: No se pudo identificar el predio.'); return; }

    const numColindantes = parseInt(document.getElementById('numColindantes')?.value) || 0;
    const colindantesData = [];
    for (let i = 1; i <= numColindantes; i++) {
        colindantesData.push({
            id_colindante: getVal(`col${i}_id_colindante`) || null,
            numero_colindante: i,
            cardinalidad: getVal(`col${i}_cardinalidad`),
            fmi_colindante: getVal(`col${i}_fmi`),
            numero_predial_colindante: getVal(`col${i}_numero_predial`),
            nombre_predio_colindante: getVal(`col${i}_nombre`),
            id_ant: getVal(`col${i}_id_ant`),
            tiene_levantamiento: getBool(`col${i}_tiene_levantamiento`),
            tipo_colindante: getVal(`col${i}_tipo`),
            documento_origen_colindante: getVal(`col${i}_documento_origen`),
            propietario_inicial_colindante: getVal(`col${i}_propietario_inicial`),
            propietario_actual_vur_colindante: getVal(`col${i}_propietario`),
            plano_colindante: getVal(`col${i}_plano`),
            necesario_validar_campo: getBool(`col${i}_validar_campo`),
            observacion_colindante: getVal(`col${i}_observacion`),
            colindante_primera_anotacion: getVal(`col${i}_colindante_primera_anotacion`),
            colindante_ultima_anotacion: getVal(`col${i}_colindante_ultima_anotacion`),
            vereda_colindante: getVal(`col${i}_vereda_colindante`) || null,
            municipio_colindante: getVal(`col${i}_municipio_colindante`) || null,
            departamento_colindante: getVal(`col${i}_departamento_colindante`) || null,
            estado_validacion: getVal(`col${i}_estado_validacion`) || null,
            fuente_validacion: getVal(`col${i}_fuente_validacion`) || null,
            requiere_verificacion_campo: getBool(`col${i}_requiere_verificacion_campo`),
            es_presunto_baldio: getBool(`col${i}_presunto_baldio`),
            es_derivado_colindante: getBool(`col${i}_es_derivado_colindante`),
            id_predio_matriz_colindante: getVal(`col${i}_id_predio_matriz_colindante`, null),
            tiene_derivadas_colindante: getBool(`col${i}_tiene_derivadas_colindante`),
            derivadas_ids_colindante: getVal(`col${i}_derivadas_ids_colindante`, null),
        });
    }

    const payload = {
        predio: {
            propietario_vur: getVal('predio_propietario_vur'),
            numero_documento_propietario_vur:     getVal('predio_num_doc_propietario'),
            nombre_predio_vur: getVal('predio_nombre_vur'),
            tipo_gestion: getVal('predio_tipo_gestion'),
            area_vur: parseFloat(getVal('predio_area_vur')) || null,
            area_final_gc: parseFloat(getVal('predio_area_final_gc')) || null,
            es_derivado: getBool('predio_es_derivado'),
            id_predio_matriz: getVal('predio_id_matriz') || null,
            tiene_derivadas: getBool('predio_tiene_derivadas'),
            derivadas_ids: getVal('predio_derivadas_ids'),
            proceso_englobe_desenglobe: getBool('predio_englobe')
        },
        localizacion: {
            dpto_espacial: getVal('loc_dpto_espacial'),
            municipio_espacial: getVal('loc_municipio_espacial'),
            vereda_espacial: getVal('loc_vereda_espacial'),
            tiene_municipio_diferente_vur: getBool('loc_municipio_diferente'),
            municipio_vur: getVal('loc_municipio_vur'),
            dpto_vur: getVal('loc_dpto_vur'),
            vereda_vur: getVal('loc_vereda_vur')
        },
        catastral: {
            numero_predial: getVal('cat_numero_predial'),
            fmi_r1_r2: getVal('cat_fmi_r1_r2'),
            nombre_predio_r1_r2: getVal('cat_nombre_r1_r2'),
            propietario: getVal('cat_propietario'),
            cedula_propietario: getVal('cat_cedula'),
            area_terreno_ha: parseFloat(getVal('cat_area_terreno')) || null,
            area_terreno_ha_r1r2: parseFloat(getVal('cat_area_terreno_r1r2')) || null,
            destino_economico: getVal('cat_destino'),
            referencia_catastral: getVal('cat_referencia')
        },
        juridica: {
            documento_origen_inmueble:   getVal('jur_documento_origen'),
            area_juridica_ha:            parseFloat(getVal('jur_area_ha')) || null,
            nombre_predio_origen:        getVal('jur_nombre_origen'),
            adjudicatario_inicial:       getVal('jur_adjudicatario'),
            referencia_catastral_vur:    getVal('jur_referencia_catastral_vur'),
            nupre:                       getVal('jur_nupre'),
            documento_primera_anotacion: getVal('jur_documento_primera_anotacion'),
            documento_ultima_anotacion:  getVal('jur_doc_ultima_anotacion'),
            salvedades:                  getVal('jur_salvedades'),
            vida_juridica:               getVal('jur_vida_juridica'),
            fecha_apertura_folio:        getVal('jur_fecha_apertura_folio'),
            tipo_instrumento:            getVal('jur_tipo_instrumento'),
            fecha_instrumento:           getVal('jur_fecha_instrumento'),
            tipo_predio:                 getVal('jur_tipo_predio'),
            estado_folio:                getVal('jur_estado_folio'),
            complementaciones:           getVal('jur_complementaciones'),
            cabida_linderos:             getVal('jur_cabida_linderos'),
            nombre_predio_actual:        getVal('jur_nombre_predio_actual'),
            plano_juridica:              getVal('jur_plano'),
            cas:                         getVal('jur_cas'),
        },

        colindantes: colindantesData
    };

    const btn = document.getElementById('btnGuardar');
    if (btn) { btn.disabled = true; btn.textContent = 'Guardando...'; }

    fetch(`${BASE_URL}/predios/${idPredioActual}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    })
    .then(r => r.json())
    .then(data => {
        if (data.success) {
            // ✅ Usa los datos del PUT directamente — sin segundo GET
            if (data.data) llenarFormulario(data.data);

            const msg = document.getElementById('mensajeEstado');
            if (msg) {
                msg.textContent = '✅ Predio guardado correctamente.';
                msg.style.display = 'block';
                setTimeout(() => msg.style.display = 'none', 3000);
            }
        } else {
            alert('❌ Error al guardar: ' + data.error);
        }
    })
    .catch(err => alert('❌ Error: ' + err.message))
    .finally(() => {
        if (btn) {
            btn.disabled = false;
            btn.textContent = '💾 Guardar';
        }
    });
}

// ===== GENERAR COLINDANTES =====
// num: cantidad a generar
// colindantesData: array de datos existentes de la BD (opcional)
window.generarColindantes = function (num, colindantesData) {
    const contenedor = document.getElementById('contenedorColindantes');
    if (!contenedor) return;

    const numActualDOM = parseInt(document.getElementById('numColindantes')?.value || '0', 10);

    // 1) Determinar num deseado
    if (num === undefined || num === null) {
        num = numActualDOM || 0;
    }

    // 2) Si NO nos pasan colindantesData, reconstruirla desde el DOM actual
    if (!Array.isArray(colindantesData)) {
        colindantesData = [];

        const getById = (id) => {
            const el = document.getElementById(id);
            return el ? el.value.trim() : null;
        };
        const getBool = (id) => {
            const v = getById(id);
            return v === '1' || v === 'true';
        };

        for (let i = 1; i <= numActualDOM; i++) {
            colindantesData.push({
                id_colindante: getById(`col${i}_id_colindante`) || null,

                cardinalidad: getById(`col${i}_cardinalidad`),
                fmi_colindante: getById(`col${i}_fmi`),
                numero_predial_colindante: getById(`col${i}_numero_predial`),
                nombre_predio_colindante: getById(`col${i}_nombre`),
                id_ant: getById(`col${i}_id_ant`),
                tiene_levantamiento: getBool(`col${i}_tiene_levantamiento`),
                tipo_colindante: getById(`col${i}_tipo`),

                documento_origen_colindante: getById(`col${i}_documento_origen`),
                propietario_inicial_colindante: getById(`col${i}_propietario_inicial`),
                propietario_actual_vur_colindante: getById(`col${i}_propietario`),

                colindante_primera_anotacion: getById(`col${i}_colindante_primera_anotacion`),
                colindante_ultima_anotacion: getById(`col${i}_colindante_ultima_anotacion`),

                plano_colindante: getById(`col${i}_plano`),
                necesario_validar_campo: getBool(`col${i}_validar_campo`),

                vereda_colindante: getById(`col${i}_vereda_colindante`),
                municipio_colindante: getById(`col${i}_municipio_colindante`),
                departamento_colindante: getById(`col${i}_departamento_colindante`),

                es_derivado_colindante: getBool(`col${i}_es_derivado_colindante`),
                id_predio_matriz_colindante: getById(`col${i}_id_predio_matriz_colindante`),
                tiene_derivadas_colindante: getBool(`col${i}_tiene_derivadas_colindante`),
                derivadas_ids_colindante: getById(`col${i}_derivadas_ids_colindante`),

                estado_validacion: getById(`col${i}_estado_validacion`),
                fuente_validacion: getById(`col${i}_fuente_validacion`),
                requiere_verificacion_campo: getBool(`col${i}_requiere_verificacion_campo`),
                es_presunto_baldio: getBool(`col${i}_presunto_baldio`),

                observacion_colindante: getById(`col${i}_observacion`)
            });
        }
    }

    // 3) Limpiar contenedor y regenerar según num
    contenedor.innerHTML = '';

    for (let i = 1; i <= num; i++) {
        const col = colindantesData[i - 1] || {};
        const idCol = col.id_colindante || col.idcolindante || null; // compatibilidad

        const div = document.createElement('div');
        div.className = 'colindante-card';
        if (idCol) div.setAttribute('data-id-colindante', idCol);
        div.setAttribute('data-colindante-num', i);

        const cardinalidadTxt = col.cardinalidad || 'Sin cardinalidad';
        const nombrePredioTxt = col.nombre_predio_colindante || 'Sin nombre';
        const fmiTxt = col.fmi_colindante || 'Sin FMI';

        const adjuntosHTML = idCol ? `
            <div class="form-group">
                <label>VUR General (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'VUR_GENERAL',${idCol})">
            </div>
            <div class="form-group">
                <label>VUR Anotaciones (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'VUR_ANOTACIONES',${idCol})">
            </div>
            <div class="form-group">
                <label>Copia Simple (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'COPIA_SIMPLE',${idCol})">
            </div>
            <div class="form-group">
                <label>Copia Simple del Matriz o Derivados (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'COPIA_SIMPLE_M_D',${idCol})">
            </div>
            <div class="form-group">
                <label>Certificado Catastral (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'CERTIFICADO_CATASTRAL',${idCol})">
            </div>
            <div class="form-group">
                <label>Plano (PDF)</label>
                <input type="file" accept=".pdf,.tiff,.tif,.jpg,.jpeg"
                       onchange="subirAdjuntoColindante(this,'PLANO',${idCol})">
            </div>
            <div class="form-group">
                <label>Documento Origen (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'DOCUMENTO_ORIGEN',${idCol})">
            </div>
            <div class="form-group">
                <label>Escrituras (PDF)</label>
                <input type="file" accept=".pdf"
                       onchange="subirAdjuntoColindante(this,'ESCRITURAS',${idCol})">
            </div>
            <div class="form-group">
                <label>Documento Matriz o Derivadas del Colindante (PDF)</label>
                <input type="file" id="adj_doc_matriz_derivadas_colindante" accept=".pdf"
                    onchange="subirAdjuntoColindante(this, 'DOC_MATRIZ_DERIVADAS_COL', ${idCol})">
            </div>
            <div class="form-group">
                <label>Otros Documentos (Actas Secuestro, Expedientes, etc)</label>
                <input type="file" accept=".pdf,.doc,.docx" multiple
                       onchange="subirAdjuntoColindante(this,'OTROS',${idCol})">
            </div>
            <div id="docs_col${idCol}" style="margin-top:8px;"></div>
        ` : `<p class="adjunto-warning">
                ⚠️ Guarda el predio primero para poder adjuntar documentos a este colindante.
             </p>`;

        div.innerHTML = `
            <div id="colindante-anchor-${i}" class="colindante-anchor"></div>

            <div class="colindante-sticky-header">
                <div class="colindante-header-left">
                    <span class="colindante-badge">Colindante ${i}</span>
                    <span class="colindante-meta">
                        ${cardinalidadTxt} · ${nombrePredioTxt} · ${fmiTxt}
                    </span>
                </div>

                <div class="colindante-acciones">
                    <button type="button" class="btn-col-nav" onclick="irAColindante(${i})">
                        Ver inicio
                    </button>
                    <button type="button" class="btn-col-nav" onclick="window.scrollTo({top:0,behavior:'smooth'})">
                        Ir arriba
                    </button>
                </div>
            </div>

            <div class="colindante-body">
                <div class="colindante-summary">
            <h4 class="colindante-title">
                Colindante N° ${i}
                ${idCol ? `<small class="colindante-id"> — ID BD: ${idCol}</small>` : ''}
            </h4>
            <input type="hidden" id="col${i}_id_colindante" name="col${i}_id_colindante" value="${idCol || ''}">
            <div class="form-group">
                <label>Cardinalidad</label>
                <select id="col${i}_cardinalidad" name="col${i}_cardinalidad">
                    <option value="">-- Seleccione --</option>
                    <option value="NORTE">NORTE</option>
                    <option value="SUR">SUR</option>
                    <option value="ESTE">ESTE</option>
                    <option value="OESTE">OESTE</option>
                </select>
            </div>
            <div class="form-group">
                <label>Tipo Colindante</label>
                <select id="col${i}_tipo" name="col${i}_tipo">
                    <option value="">-- Seleccione --</option>
                    <option value="REGISTRAL">REGISTRAL</option>
                    <option value="CATASTRAL">CATASTRAL</option>
                    <option value="CATASTRAL Y REGISTRAL">CATASTRAL Y REGISTRAL</option>
                    <option value="APOYO">APOYO</option>
                </select>
            </div>
            <div class="form-group">
                <label>FMI Colindante</label>
                <input type="text" id="col${i}_fmi" name="col${i}_fmi" placeholder="FMI del predio colindante">
            </div>
            <div class="form-group">
                <label>Número Predial Colindante</label>
                <input type="text" id="col${i}_numero_predial" name="col${i}_numero_predial" placeholder="Número predial">
            </div>
            <div class="form-group">
                <label>Nombre del Predio Colindante</label>
                <input type="text" id="col${i}_nombre" name="col${i}_nombre" placeholder="Nombre del predio colindante">
            </div>
            <div class="form-group">
                <label>ID ANT (si aplica)</label>
                <input type="text" id="col${i}_id_ant" name="col${i}_id_ant" placeholder="ID ANT si aplica">
            </div>
            <div class="form-group">
                <label>¿Tiene Levantamiento?</label>
                <select id="col${i}_tiene_levantamiento" name="col${i}_tiene_levantamiento">
                    <option value="0">NO</option>
                    <option value="1">SÍ</option>
                </select>
            </div>
            <div class="form-group">
                <label>Documento Origen Colindante</label>
                <input type="text" id="col${i}_documento_origen" name="col${i}_documento_origen" placeholder="Documento que da origen">
            </div>
            <div class="form-group">
                <label>Área del colindante / Observación Colindante</label>
                <textarea id="col${i}_observacion" name="col${i}_observacion" rows="2" placeholder="Área u Observaciones..."></textarea>
            </div>
            <div class="form-group">
                <label>Propietario Inicial Colindante</label>
                <input type="text" id="col${i}_propietario_inicial" name="col${i}_propietario_inicial" placeholder="Propietario inicial">
            </div>
            <div class="form-group">
                <label>Propietario Actual VUR Colindante</label>
                <input type="text" id="col${i}_propietario" name="col${i}_propietario" placeholder="Propietario actual según VUR">
            </div>

            <div class="form-group">
                <label>Primera anotación colindante</label>
                <input type="text" id="col${i}_colindante_primera_anotacion" name="col${i}_colindante_primera_anotacion" placeholder="Primera anotación">
            </div>
            <div class="form-group">
                <label>Última anotación colindante (Propietario actual)</label>
                <input type="text" id="col${i}_colindante_ultima_anotacion" name="col${i}_colindante_ultima_anotacion" placeholder="Última anotación">
            </div>
            <div class="form-group">
                <label>Plano Colindante</label>
                <select id="col${i}_plano" name="col${i}_plano">
                    <option value="">-- Seleccione --</option>
                    <option value="EN ESPERA">EN ESPERA</option>
                    <option value="SIN PLANO">SIN PLANO (NO ENCONTRADO)</option>
                    <option value="PLANO">PLANO</option>
                    <option value="CARTERA">CARTERA</option>
                    <option value="DESCRIPCION EN RES. O ESC.">DESCRIPCIÓN EN RES. O ESC.</option>
                </select>
            </div>
            <div class="form-group">
                <label>Vereda Colindante</label>
                <input type="text" id="col${i}_vereda_colindante" name="col${i}_vereda_colindante"
                        placeholder="Vereda o corregimiento">
            </div>
            <div class="form-group">
                <label>Municipio Colindante</label>
                <input type="text" id="col${i}_municipio_colindante" name="col${i}_municipio_colindante"
                        placeholder="Municipio">
            </div>

            <!-- CAMPOS MATRIZ-DERIVADAS DEL COLINDANTE -->
            <div class="form-group">
                <label>¿El colindante nace de un predio matriz?</label>
                <select id="col${i}_es_derivado_colindante" name="col${i}_es_derivado_colindante" onchange="toggleMatrizColindante(${i})">
                    <option value="0">NO</option>
                    <option value="1">Sí</option>
                </select>
            </div>

            <div class="form-group" id="grupoMatrizColindante_${i}" style="display:none;">
                <label>FMIs Predio Matriz del Colindante</label>
                <input type="text" id="col${i}_id_predio_matriz_colindante" name="col${i}_id_predio_matriz_colindante" placeholder="FMIs del predio matriz del colindante">
            </div>

            <div class="form-group">
                <label>¿El colindante tiene predios derivados?</label>
                <select id="col${i}_tiene_derivadas_colindante" name="col${i}_tiene_derivadas_colindante" onchange="toggleDerivadasColindante(${i})">
                    <option value="0">NO</option>
                    <option value="1">Sí</option>
                </select>
            </div>

            <div class="form-group" id="grupoDerivadasColindante_${i}" style="display:none;">
                <label>FMIs Predios Derivados del Colindante (separados por coma)</label>
                <input type="text" id="col${i}_derivadas_ids_colindante" name="col${i}_derivadas_ids_colindante" placeholder="Ej: 000-0001, 000-0003">
            </div>

            <div class="form-group">
                <label>Fuente de Validación</label>
                <textarea type="text" id="col${i}_fuente_validacion" name="col${i}_fuente_validacion"
                        rows="4" maxlength="65535"
                        placeholder="Ej. Catastro y Resolución de adjudicación..."></textarea>
            </div>
            <div class="form-group">
                <label>Estado de Validación</label>
                <select id="col${i}_estado_validacion" name="col${i}_estado_validacion">
                    <option value="">-- Seleccione --</option>
                    <option value="VALIDADO">VALIDADO</option>
                    <option value="NO_VERIFICADO">NO VERIFICADO</option>
                    <option value="NO_VERIFICADO">NO APLICA</option>
                    <option value="PRESUNTO_BALDIO">PRESUNTO BALDÍO</option>
                </select>
            </div>
            <div class="form-group">
                <label>¿Requiere Verificación en Campo?</label>
                <select id="col${i}_requiere_verificacion_campo" name="col${i}_requiere_verificacion_campo">
                    <option value="0">NO</option>
                    <option value="1">SÍ</option>
                </select>
            </div>
            <div class="form-group">
                <label>¿Presunto Baldío?</label>
                <select id="col${i}_presunto_baldio" name="col${i}_presunto_baldio">
                    <option value="0">NO</option>
                    <option value="1">SÍ</option>
                </select>
            </div>

            <div class="documentos-adjuntos-panel">
                <strong class="documentos-adjuntos-title">📎 Documentos Adjuntos — Colindante ${i}</strong>
                ${adjuntosHTML}
            </div>
        `;

        contenedor.appendChild(div);

        // 4) Poblar campos con datos existentes
        const setVal = (id, valor) => {
            if (valor === null || valor === undefined || valor === '') return;
            const el = document.getElementById(id);
            if (el) el.value = valor;
        };
        const setBool = (id, val) => {
            const el = document.getElementById(id);
            if (!el) return;
            el.value = val ? '1' : '0';
        };

        setVal(`col${i}_cardinalidad`, col.cardinalidad || col.cardinalidad);
        setVal(`col${i}_fmi`, col.fmi_colindante || col.fmicolindante);
        setVal(`col${i}_numero_predial`, col.numero_predial_colindante || col.numeropredialcolindante);
        setVal(`col${i}_nombre`, col.nombre_predio_colindante || col.nombreprediocolindante);
        setVal(`col${i}_id_ant`, col.id_ant || col.idant);
        setBool(`col${i}_tiene_levantamiento`, col.tiene_levantamiento || col.tienelevantamiento);
        setVal(`col${i}_tipo`, col.tipo_colindante || col.tipocolindante);

        setVal(`col${i}_documento_origen`, col.documento_origen_colindante || col.documentoorigencolindante);
        setVal(`col${i}_propietario_inicial`, col.propietario_inicial_colindante || col.propietarioinicialcolindante);
        setVal(`col${i}_propietario`, col.propietario_actual_vur_colindante || col.propietarioactualvurcolindante);

        setVal(`col${i}_colindante_primera_anotacion`, col.colindante_primera_anotacion || col.colindanteprimeraanotacion);
        setVal(`col${i}_colindante_ultima_anotacion`, col.colindante_ultima_anotacion || col.colindanteultimaanotacion);

        setVal(`col${i}_plano`, col.plano_colindante || col.planocolindante);
        setBool(`col${i}_validar_campo`, col.necesario_validar_campo || col.necesariovalidarcampo);

        setVal(`col${i}_vereda_colindante`, col.vereda_colindante || col.veredacolindante);
        setVal(`col${i}_municipio_colindante`, col.municipio_colindante || col.municipiocolindante);
        setVal(`col${i}_departamento_colindante`, col.departamento_colindante || col.departamentocolindante);

        setBool(`col${i}_es_derivado_colindante`, col.es_derivado_colindante);
        setVal(`col${i}_id_predio_matriz_colindante`, col.id_predio_matriz_colindante);
        setBool(`col${i}_tiene_derivadas_colindante`, col.tiene_derivadas_colindante);
        setVal(`col${i}_derivadas_ids_colindante`, col.derivadas_ids_colindante);

        setVal(`col${i}_estado_validacion`, col.estado_validacion || col.estadovalidacion);
        setVal(`col${i}_fuente_validacion`, col.fuente_validacion || col.fuentevalidacion);
        setBool(`col${i}_requiere_verificacion_campo`, col.requiere_verificacion_campo || col.requiereverificacioncampo);
        setBool(`col${i}_presunto_baldio`, col.es_presunto_baldio || col.espresuntobaldio);

        setVal(`col${i}_observacion`, col.observacion_colindante || col.observacioncolindante);

        // Ajustar grupos matriz/derivadas
        if (typeof toggleMatrizColindante === 'function') toggleMatrizColindante(i);
        if (typeof toggleDerivadasColindante === 'function') toggleDerivadasColindante(i);

        // 5) Mostrar documentos existentes si vienen en colindantesData
        const docsInCol = col.documentos_colindante || col.documentoscolindante;
        if (idCol && docsInCol && docsInCol.length > 0) {
            mostrarDocumentosColindante(docsInCol, idCol);
        }
    }

    // 6) Actualizar tabla resumen
    if (typeof actualizarTablaResumen === 'function') {
        actualizarTablaResumen();
    }

    // 7) Fallback: si regeneramos desde DOM (sin docs en memoria), recargar docs por GET
    const tieneDocsEnMemoria = Array.isArray(colindantesData) &&
        colindantesData.some(c =>
            c && (
                (c.documentos_colindante && c.documentos_colindante.length > 0) ||
                (c.documentoscolindante && c.documentoscolindante.length > 0)
            )
        );

    if (!tieneDocsEnMemoria && typeof idPredioActual !== 'undefined' && idPredioActual) {
        fetch(`${BASE_URL}/predios/${idPredioActual}`)
            .then(r => r.json())
            .then(d => {
                if (!d.success || !d.data || !Array.isArray(d.data.colindantes)) return;

                d.data.colindantes.forEach(c => {
                    const idCol = c.id_colindante || c.idcolindante;
                    const docs = c.documentos_colindante || c.documentoscolindante;
                    if (idCol && docs && docs.length > 0) {
                        mostrarDocumentosColindante(docs, idCol);
                    }
                });
            })
            .catch(err => {
                console.error('Error refrescando documentos de colindantes:', err);
            });
    }
};




// ===== TABLA RESUMEN COLINDANTES =====
function actualizarTablaResumen() {
    const contenedor = document.getElementById('tablaResumenColindantes');
    if (!contenedor) return;

    const num = parseInt(document.getElementById('numColindantes')?.value) || 0;

    if (num === 0) {
        contenedor.innerHTML = '';
        return;
    }

    let html = `
        <div class="resumen-colindantes-card">
            <div class="resumen-colindantes-header">
                <span>📋 Resumen de Colindantes</span>
                <span>Haz clic en una fila para ir al formulario</span>
            </div>
            <table class="resumen-colindantes-table">
                <thead>
                    <tr>
                        <th class="number-cell">N°</th>
                        <th>Cardinalidad</th>
                        <th>ID ANT</th>
                        <th>FMI</th>
                        <th>Área (ha) - Observación</th>
                        <th>Tipo</th>
                        <th>Dirección</th>
                        <th>Doc Origen</th>
                        <th>Propietario Inicial</th>
                        <th>¿Plano?</th>
                    </tr>
                </thead>
                <tbody>`;

    for (let i = 1; i <= num; i++) {
        const card      = getVal(`col${i}_cardinalidad`)  || '—';
        const idAnt     = getVal(`col${i}_id_ant`)        || '—';
        const fmi       = getVal(`col${i}_fmi`)           || '—';
        const estado    = getVal(`col${i}_observacion`)        || '—';
        const tipo      = getVal(`col${i}_tipo`)          || '—';
        const nom_col   = getVal(`col${i}_nombre`)|| '—';
        const docOrigen = getVal(`col${i}_documento_origen`) || '—';
        const propietarioinicial = getVal(`col${i}_propietario_inicial`) || '—';
        const plano     = getVal(`col${i}_plano`)          || '—';
        const rowBg     = i % 2 === 0 ? '#f3f4f6' : '#ffffff';

        html += `
            <tr onclick="irAColindante(${i})"
                style="cursor:pointer;background:${rowBg};transition:background 0.15s;"
                onmouseover="this.style.background='#f3f4f6'"
                onmouseout="this.style.background='${rowBg}'"
                title="Ir al formulario del Colindante N° ${i}">
                <td class="number-cell">${i}</td>
                <td>${card}</td>
                <td>${idAnt}</td>
                <td>${fmi}</td>
                <td>${estado}</td>
                <td>${tipo}</td>
                <td>${nom_col}</td>
                <td>${docOrigen}</td>
                <td>${propietarioinicial}</td>
                <td>${plano}</td>

            </tr>`;
    }

    html += `</tbody></table></div>`;
    contenedor.innerHTML = html;
}

window.irAColindante = function(i) {
    const div = document.querySelector(`[data-colindante-num="${i}"]`);
    if (!div) return;
    div.scrollIntoView({ behavior: 'smooth', block: 'start' });
    div.style.outline = '2px solid #3498db';
    div.style.transition = 'outline 0.3s';
    setTimeout(() => { div.style.outline = ''; }, 1800);
};

// ===== DOCUMENTOS DEL PREDIO =====
function mostrarDocumentos(docs) {
    const contenedor = document.getElementById('listaDocumentos');
    if (!contenedor) return;
    if (!docs || docs.length === 0) {
        contenedor.innerHTML = '<p class="documentos-empty">No hay documentos adjuntos.</p>';
        return;
    }
    let html = '<table class="documentos-table">';
    html += '<thead><tr><th>Tipo</th><th>Archivo</th><th>Tamaño</th><th>Acción</th></tr></thead><tbody>';
    docs.forEach(doc => {
        const kb = doc.tamaño_bytes ? (doc.tamaño_bytes / 1024).toFixed(1) + ' KB' : 'N/A';
        html += `<tr>
            <td>${doc.tipo_documento || 'N/A'}</td>
            <td>${doc.nombre_archivo}</td>
            <td>${kb}</td>
            <td>
                <a href="${BASE_URL}/documentos/${doc.id_documento}/ver"
                target="_blank"
                class="documentos-link">
                    👁 Abrir
                </a>
                <button class="btn-danger documentos-action-button"
                        onclick="eliminarDocumento(${doc.id_documento})">
                    Eliminar
                </button>
            </td></tr>`;
    });
    html += '</tbody></table>';
    contenedor.innerHTML = html;
}

function subirAdjunto(input, tipoDoc) {
    if (!input.files || !input.files.length || !idPredioActual) return;
    Array.from(input.files).forEach(file => {
        const fd = new FormData();
        fd.append('archivo', file);
        fd.append('tipo_documento', tipoDoc);
        fetch(`${BASE_URL}/predios/${idPredioActual}/documentos`, { method: 'POST', body: fd })
            .then(r => r.json())
            .then(data => {
                if (data.success) {
                    fetch(`${BASE_URL}/predios/${idPredioActual}`)
                        .then(r => r.json())
                        .then(d => { if (d.success) mostrarDocumentos(d.data.documentos); });
                } else { alert('Error al subir: ' + data.error); }
            })
            .catch(err => alert('Error: ' + err.message));
    });
}

function eliminarDocumento(idDoc) {
    if (!confirm('¿Eliminar este documento?')) return;
    fetch(`${BASE_URL}/documentos/${idDoc}`, { method: 'DELETE' })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                fetch(`${BASE_URL}/predios/${idPredioActual}`)
                    .then(r => r.json())
                    .then(d => { if (d.success) mostrarDocumentos(d.data.documentos); });
            } else { alert('Error: ' + data.error); }
        });
}

// ===== DOCUMENTOS DE COLINDANTE =====
function mostrarDocumentosColindante(docs, idColindante) {
    const contenedor = document.getElementById(`docs_col${idColindante}`);
    if (!contenedor) return;
    if (!docs || docs.length === 0) {
        contenedor.innerHTML = '<p class="documentos-empty">No hay documentos.</p>';
        return;
    }
    let html = '<table class="documentos-table">';
    html += '<thead><tr><th>Tipo</th><th>Archivo</th><th>Tamaño</th><th>Acción</th></tr></thead><tbody>';
    docs.forEach(doc => {
        const kb = doc.tamaño_bytes ? (doc.tamaño_bytes / 1024).toFixed(1) + ' KB' : 'N/A';
        html += `<tr>
            <td>${doc.tipo_documento || 'N/A'}</td>
            <td>${doc.nombre_archivo}</td>
            <td>${kb}</td>
            <td>
                <a href="${BASE_URL}/documentos-colindante/${doc.id_doc_colindante}/ver" target="_blank"
                   class="documentos-link">Abrir</a>
                <button class="btn-danger documentos-action-button"
                        onclick="eliminarDocumentoColindante(${doc.id_doc_colindante}, ${idColindante})">Eliminar</button>
            </td>
        </tr>`;
    });
    html += '</tbody></table>';
    contenedor.innerHTML = html;
}

function subirAdjuntoColindante(input, tipoDoc, idColindante) {
    if (!input.files || !input.files.length) return;
    Array.from(input.files).forEach(file => {
        const fd = new FormData();
        fd.append('archivo', file);
        fd.append('tipo_documento', tipoDoc);
        fetch(`${BASE_URL}/predios/${idPredioActual}/colindantes/${idColindante}/documentos`, {
            method: 'POST', body: fd
        })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                fetch(`${BASE_URL}/predios/${idPredioActual}`)
                    .then(r => r.json())
                    .then(d => {
                        if (d.success) {
                            const col = d.data.colindantes
                                .find(c => c.id_colindante == idColindante);
                            mostrarDocumentosColindante(col ? col.documentos_colindante : [], idColindante);
                        }
                    });
            } else {
                alert('Error al subir: ' + data.error);
            }
        })
        .catch(err => alert('Error al subir archivo: ' + err.message));
    });
}

function eliminarDocumentoColindante(idDoc, idColindante) {
    if (!confirm('¿Eliminar este documento?')) return;

    fetch(`${BASE_URL}/documentos-colindante/${idDoc}`, { method: 'DELETE' })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                fetch(`${BASE_URL}/predios/${idPredioActual}`)
                    .then(r => r.json())
                    .then(d => {
                        if (d.success) {
                            const col = d.data.colindantes
                                .find(c => c.id_colindante == idColindante);
                            mostrarDocumentosColindante(col ? col.documentos_colindante : [], idColindante);
                        } else {
                            alert('Error al recargar: ' + d.error);
                        }
                    });
            } else {
                alert('Error al eliminar: ' + data.error);
            }
        })
        .catch(err => alert('Error de conexión: ' + err.message));
}

document.addEventListener('input', (e) => {
  const el = e.target;
  if (!el.matches('input, textarea, select')) return;
  if (el.type === 'file' || el.disabled || el.readOnly) return;
  el.classList.toggle('campo-vacio', _esVacio(el));
  if (el.id && /^col\d+_(fmi|id_ant|numero_predial)$/.test(el.id)) {
        actualizarTablaResumen();
    }
});

document.addEventListener('change', (e) => {
  const el = e.target;
  if (!el.matches('select')) return;
  el.classList.toggle('campo-vacio', _esVacio(el));
  if (el.id && /^col\d+_cardinalidad$/.test(el.id)) {
        actualizarTablaResumen();
    }
});

window.toggleMatrizColindante = function(i) {
    const g = document.getElementById(`grupoMatrizColindante_${i}`);
    if (g) g.style.display = 
        document.getElementById(`col${i}_es_derivado_colindante`)?.value === '1' 
        ? 'block' : 'none';
};
window.toggleDerivadasColindante = function(i) {
    const g = document.getElementById(`grupoDerivadasColindante_${i}`);
    if (g) g.style.display = 
        document.getElementById(`col${i}_tiene_derivadas_colindante`)?.value === '1' 
        ? 'block' : 'none';
};

