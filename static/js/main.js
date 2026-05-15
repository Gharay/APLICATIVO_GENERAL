/**
 * static/js/main.js
 * Lógica del Gestor de Predios (index.html)
 */

let predios = [];
let prediosFiltrados = [];
let tipoActivo = 'TODOS';
let terminoBusqueda = '';
const BASE_URL = '/api';
let predioAEliminar = null;

// ===== INICIALIZACIÓN =====
document.addEventListener('DOMContentLoaded', function () {
    console.log('✓ App cargada');

    if (window.location.pathname === '/') {
        cargarPredios();
    }

    const btnNuevoPredio = document.getElementById('btnNuevoPredio');
    if (btnNuevoPredio) {
        btnNuevoPredio.addEventListener('click', abrirModalNuevoPredio);
    }

    const formNuevoPredio = document.getElementById('formNuevoPredio');
    if (formNuevoPredio) {
        formNuevoPredio.addEventListener('submit', crearPredio);
    }

    const searchInput = document.getElementById('searchPredios');
    if (searchInput) {
        searchInput.addEventListener('input', function (e) {
            terminoBusqueda = e.target.value.toLowerCase();
            aplicarFiltros();
        });
    }
});

// ===== CARGAR PREDIOS =====
function cargarPredios() {
    fetch(`${BASE_URL}/predios`)
        .then(r => {
            if (!r.ok) throw new Error(`HTTP ${r.status}`);
            return r.json();
        })
        .then(data => {
            if (data.success) {
                predios = data.data;
                prediosFiltrados = predios;
                console.log(`✓ ${predios.length} predios cargados`);
                actualizarContadores();
                mostrarPredios(predios);
            } else {
                console.error('Error:', data.error);
            }
        })
        .catch(err => {
            console.error('Error cargando predios:', err);
            document.getElementById('listaPredios').innerHTML =
                `<p style="color:red;">Error al conectar con el servidor: ${err.message}</p>`;
        });
}

// ===== MOSTRAR PREDIOS =====
function mostrarPredios(listaPredios) {
    const contenedor = document.getElementById('listaPredios');
    if (!contenedor) return;

    if (listaPredios.length === 0) {
        const mensajeTipo = tipoActivo === 'TODOS'
            ? 'No hay predios creados.'
            : `No hay predios en <strong>${tipoActivo}</strong>.`;
        const mensajeBusqueda = terminoBusqueda
            ? `<p style="color:#666;">Sin resultados para "<strong>${terminoBusqueda}</strong>"</p>`
            : '';
        contenedor.innerHTML = `
            <p style="color:#666; text-align:center; padding:40px;">
                ${mensajeTipo}
                ${tipoActivo === 'TODOS' && !terminoBusqueda
                    ? 'Haz clic en <strong>+ Crear Predio</strong> para comenzar.' : ''}
            </p>${mensajeBusqueda}`;
        return;
    }

    contenedor.innerHTML = '';
    listaPredios.forEach(predio => {
        const card = document.createElement('div');
        card.className = 'card-predio';

        const tipoGestion = predio.tipo_gestion || 'Sin Asignar';
        const tipoClase   = tipoGestion.toLowerCase().replace(/\s+/g, '');

        card.innerHTML = `
            <div class="tipo-gestion-badge tipo-${tipoClase}">${tipoGestion}</div>
            <div style="margin-bottom:10px; margin-top:30px;">
                <strong style="font-size:16px; color:#2c3e50;">ID: ${predio.id}</strong>
            </div>
            <div style="font-size:13px; color:#555; margin-bottom:6px;">
                <strong>FMI:</strong> ${predio.fmi}
            </div>
            <div style="font-size:13px; color:#555; margin-bottom:6px;">
                <strong>Propietario:</strong> ${predio.propietario_vur || 'No registrado'}
            </div>
            <div style="font-size:13px; color:#555; margin-bottom:6px;">
                <strong>Nombre Predio:</strong> ${predio.nombre_predio_vur || 'No registrado'}
            </div>
            <div style="font-size:12px; color:#888; margin-bottom:12px;">
                <strong>Estado:</strong> ${predio.estado} &nbsp;|&nbsp;
                <strong>Creado:</strong> ${predio.fecha_creacion
                    ? new Date(predio.fecha_creacion).toLocaleDateString('es-CO') : 'N/A'}
            </div>
            <div class="btn-group">
                <button class="btn-primary"
                    onclick="window.location.href='/editar/${predio.id_predio}'">
                    ✏️ Editar
                </button>
                <button class="btn-secondary"
                    onclick="window.location.href='${BASE_URL}/predios/${predio.id_predio}/export-word'">
                    📄 Word
                </button>
                <button class="btn-secondary"
                    onclick="window.location.href='${BASE_URL}/predios/${predio.id_predio}/export-excel'">
                    📊 Excel
                </button>
                <button class="btn-danger"
                    onclick="confirmarEliminar(${predio.id_predio}, '${predio.id}')">
                    🗑️ Eliminar
                </button>
            </div>`;
        contenedor.appendChild(card);
    });

    actualizarEstadisticas(listaPredios.length);
}

// ===== FILTRADO POR PESTAÑA =====
function filtrarPorTipo(tipo, boton) {
    tipoActivo = tipo;
    document.querySelectorAll('.tab-gestion').forEach(t => t.classList.remove('active'));
    boton.classList.add('active');
    aplicarFiltros();
}

function aplicarFiltros() {
    let resultado = predios;

    if (tipoActivo !== 'TODOS') {
        if (tipoActivo === 'SinAsignar') {
            resultado = resultado.filter(p => !p.tipo_gestion || p.tipo_gestion.trim() === '');
        } else {
            resultado = resultado.filter(p => p.tipo_gestion === tipoActivo);
        }
    }

    if (terminoBusqueda) {
        resultado = resultado.filter(p =>
            (p.id               || '').toLowerCase().includes(terminoBusqueda) ||
            (p.fmi              || '').toLowerCase().includes(terminoBusqueda) ||
            (p.propietario_vur  || '').toLowerCase().includes(terminoBusqueda) ||
            (p.nombre_predio_vur|| '').toLowerCase().includes(terminoBusqueda)
        );
    }

    prediosFiltrados = resultado;
    mostrarPredios(resultado);
}

// ===== CONTADORES =====
function actualizarContadores() {
    const tipos = ['Asignación', 'Preliminar', 'Entregado', 'Finalizado', 'Desistido'];
    const contadores = { 'TODOS': predios.length, 'SinAsignar': 0 };

    tipos.forEach(t => { contadores[t] = 0; });

    predios.forEach(p => {
        const tg = (p.tipo_gestion || '').trim();
        if (tipos.includes(tg)) {
            contadores[tg]++;
        } else {
            contadores['SinAsignar']++;
        }
    });

    Object.keys(contadores).forEach(tipo => {
        const badge = document.getElementById(`badge-${tipo}`);
        if (badge) badge.textContent = contadores[tipo];
    });
}

function actualizarEstadisticas(cantidadMostrada) {
    const statsBar    = document.getElementById('statsBar');
    const statTotal   = document.getElementById('stat-total');
    const statMostrando = document.getElementById('stat-mostrando');
    if (statsBar && statTotal && statMostrando) {
        statsBar.style.display    = 'flex';
        statTotal.textContent     = predios.length;
        statMostrando.textContent = cantidadMostrada;
    }
}

// ===== MODAL: CREAR PREDIO =====
function abrirModalNuevoPredio() {
    document.getElementById('nuevoid').value  = '';
    document.getElementById('nuevofmi').value = '';
    const err = document.getElementById('modalError');
    if (err) { err.style.display = 'none'; err.textContent = ''; }
    document.getElementById('modalNuevoPredio').style.display = 'flex';
    document.getElementById('nuevoid').focus();
}

function cerrarModalNuevoPredio() {
    document.getElementById('modalNuevoPredio').style.display = 'none';
}

function crearPredio(e) {
    if (e) e.preventDefault();
    const id  = document.getElementById('nuevoid').value.trim();
    const fmi = document.getElementById('nuevofmi').value.trim();
    const err = document.getElementById('modalError');
    const btn = document.getElementById('btnConfirmarCrear');

    if (!id || !fmi) {
        err.textContent = 'El ID y el FMI son obligatorios.';
        err.style.display = 'block';
        return;
    }

    btn.disabled = true;
    btn.textContent = 'Creando...';
    if (err) err.style.display = 'none';

    fetch(`${BASE_URL}/predios`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: id, fmi: fmi })
    })
    .then(r => r.json())
    .then(data => {
        if (data.success) {
            cerrarModalNuevoPredio();
            window.location.href = `/editar/${data.data.id_predio}`;
        } else {
            err.textContent = data.error || 'Error al crear el predio.';
            err.style.display = 'block';
        }
    })
    .catch(error => {
        err.textContent = 'Error de conexión: ' + error.message;
        err.style.display = 'block';
    })
    .finally(() => {
        btn.disabled = false;
        btn.textContent = 'Crear Predio';
    });
}

// ===== MODAL: ELIMINAR PREDIO =====
function confirmarEliminar(idPredio, nombrePredio) {
    predioAEliminar = idPredio;
    const span = document.getElementById('eliminarNombrePredio');
    if (span) span.textContent = nombrePredio;
    document.getElementById('modalEliminar').style.display = 'flex';
    const btnConfirmar = document.getElementById('btnConfirmarEliminar');
    btnConfirmar.onclick = function () { ejecutarEliminar(idPredio); };
}

function cerrarModalEliminar() {
    document.getElementById('modalEliminar').style.display = 'none';
    predioAEliminar = null;
}

function ejecutarEliminar(idPredio) {
    fetch(`${BASE_URL}/predios/${idPredio}`, { method: 'DELETE' })
        .then(r => r.json())
        .then(data => {
            cerrarModalEliminar();
            if (data.success) cargarPredios();
            else alert('Error al eliminar: ' + data.error);
        })
        .catch(err => alert('Error de conexión: ' + err.message));
}

// Cerrar modales al hacer clic fuera
window.addEventListener('click', function (e) {
    if (e.target.id === 'modalNuevoPredio') cerrarModalNuevoPredio();
    if (e.target.id === 'modalEliminar')    cerrarModalEliminar();
});
