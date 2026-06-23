/* ===========================================================================
   Control Etiquetas/Remitos - Despacho RDC Troncal Rosario
   Lógica de la aplicación: importación de Excel, comparación y export a PDF.
   =========================================================================== */

'use strict';

// --- Estado de la aplicación ---------------------------------------------
let datosA = null;
let datosB = null;
let nombreArchivoA = '';
let nombreArchivoB = '';
let reporteGenerado = null;

// --- Navegación entre vistas ---------------------------------------------
function mostrarImportarA() {
    ocultarTodo();
    document.getElementById('importarA').classList.remove('hidden');
}

function mostrarImportarB() {
    ocultarTodo();
    document.getElementById('importarB').classList.remove('hidden');
}

function ocultarTodo() {
    document.getElementById('welcomeMessage').classList.add('hidden');
    document.getElementById('importarA').classList.add('hidden');
    document.getElementById('importarB').classList.add('hidden');
    document.getElementById('reporteArea').classList.add('hidden');
    document.getElementById('loading').classList.add('hidden');
}

// --- Importación de archivos ---------------------------------------------
function procesarArchivoA(event) {
    const file = event.target.files[0];
    if (!file) return;

    nombreArchivoA = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array'
            });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                header: 1
            });

            const handlingUnitIndex = jsonData[0].indexOf('Handling Unit');
            if (handlingUnitIndex === -1) {
                mostrarEstado('statusA', 'Error: No se encontró la columna "Handling Unit"', 'error');
                return;
            }

            const codigos = [];
            for (let i = 1; i < jsonData.length; i++) {
                const codigo = jsonData[i][handlingUnitIndex];
                if (codigo) {
                    codigos.push(String(codigo).trim());
                }
            }

            datosA = {
                codigos: codigos,
                nombre: nombreArchivoA
            };

            const codigosUnicos = [...new Set(codigos)];
            const duplicados = encontrarDuplicados(codigos);

            let mensaje = `✅ Archivo cargado exitosamente: ${nombreArchivoA}<br>`;
            mensaje += `📊 Total de códigos: ${codigos.length}<br>`;
            mensaje += `🔢 Códigos únicos: ${codigosUnicos.length}`;

            if (duplicados.length > 0) {
                mensaje += `<br>⚠️ Se detectaron ${duplicados.length} códigos duplicados`;
            }

            mostrarEstado('statusA', mensaje, 'success');
            verificarBotonComparar();
        } catch (error) {
            mostrarEstado('statusA', `Error al procesar el archivo: ${error.message}`, 'error');
        }
    };

    reader.readAsArrayBuffer(file);
}

function procesarArchivoB(event) {
    const file = event.target.files[0];
    if (!file) return;

    nombreArchivoB = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array'
            });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                header: 1
            });

            const headers = jsonData[0];
            const huIndex = headers.indexOf('Handling Unit');
            const remitoIndex = headers.indexOf('Remito');

            if (huIndex === -1 || remitoIndex === -1) {
                mostrarEstado('statusB', 'Error: No se encontraron las columnas "Handling Unit" y/o "Remito"', 'error');
                return;
            }

            const registros = [];
            for (let i = 1; i < jsonData.length; i++) {
                const hu = jsonData[i][huIndex];
                const remito = jsonData[i][remitoIndex];
                if (hu && remito) {
                    registros.push({
                        handlingUnit: String(hu).trim(),
                        remito: String(remito).trim()
                    });
                }
            }

            datosB = {
                registros: registros,
                nombre: nombreArchivoB
            };

            const remitosUnicos = new Set(registros.map(r => r.remito)).size;
            const huUnicos = new Set(registros.map(r => r.handlingUnit)).size;

            let mensaje = `✅ Archivo cargado exitosamente: ${nombreArchivoB}<br>`;
            mensaje += `📊 Total de registros: ${registros.length}<br>`;
            mensaje += `📋 Remitos únicos: ${remitosUnicos}<br>`;
            mensaje += `🏷️ Handling Units únicos: ${huUnicos}`;

            mostrarEstado('statusB', mensaje, 'success');
            verificarBotonComparar();
        } catch (error) {
            mostrarEstado('statusB', `Error al procesar el archivo: ${error.message}`, 'error');
        }
    };

    reader.readAsArrayBuffer(file);
}

// --- Utilidades ----------------------------------------------------------
function encontrarDuplicados(array) {
    const conteo = {};
    const duplicados = [];

    array.forEach(item => {
        conteo[item] = (conteo[item] || 0) + 1;
    });

    Object.keys(conteo).forEach(key => {
        if (conteo[key] > 1) {
            duplicados.push({
                codigo: key,
                veces: conteo[key]
            });
        }
    });

    return duplicados;
}

function verificarBotonComparar() {
    const btnComparar = document.getElementById('btnComparar');
    if (datosA && datosB) {
        btnComparar.disabled = false;
    }
}

// --- Comparación ---------------------------------------------------------
function compararArchivos() {
    if (!datosA || !datosB) {
        alert('Debe cargar ambos archivos antes de comparar');
        return;
    }

    ocultarTodo();
    document.getElementById('loading').classList.remove('hidden');

    setTimeout(() => {
        try {
            const reporte = generarReporte();
            reporteGenerado = reporte;
            mostrarReporte(reporte);
            document.getElementById('btnExportar').disabled = false;
        } catch (error) {
            alert(`Error al comparar archivos: ${error.message}`);
        }
        document.getElementById('loading').classList.add('hidden');
    }, 500);
}

function generarReporte() {
    const codigosA = new Set(datosA.codigos);
    const codigosB = new Set(datosB.registros.map(r => r.handlingUnit));

    const soloEnA = [...codigosA].filter(c => !codigosB.has(c));
    const soloEnB = [...codigosB].filter(c => !codigosA.has(c));

    const huConMultiplesRemitos = new Map();
    datosB.registros.forEach(reg => {
        if (!huConMultiplesRemitos.has(reg.handlingUnit)) {
            huConMultiplesRemitos.set(reg.handlingUnit, new Set());
        }
        huConMultiplesRemitos.get(reg.handlingUnit).add(reg.remito);
    });

    const duplicadosEnRemitos = [];
    huConMultiplesRemitos.forEach((remitos, hu) => {
        if (remitos.size > 1) {
            duplicadosEnRemitos.push({
                handlingUnit: hu,
                remitos: [...remitos]
            });
        }
    });

    const duplicadosA = encontrarDuplicados(datosA.codigos);
    const remitosUnicos = new Set(datosB.registros.map(r => r.remito));

    const huBSinDuplicadosEnRemitos = new Set();
    datosB.registros.forEach(reg => {
        if (huConMultiplesRemitos.get(reg.handlingUnit).size === 1) {
            huBSinDuplicadosEnRemitos.add(reg.handlingUnit);
        }
    });

    const hayAnomalias = soloEnA.length > 0 || soloEnB.length > 0 || duplicadosEnRemitos.length > 0;

    return {
        nombreA: datosA.nombre,
        nombreB: datosB.nombre,
        duplicadosA: duplicadosA,
        totalCodigosA: datosA.codigos.length,
        soloEnA: soloEnA,
        soloEnB: soloEnB,
        duplicadosEnRemitos: duplicadosEnRemitos,
        hayAnomalias: hayAnomalias,
        totalRemitos: remitosUnicos.size,
        totalHUCoincidentes: [...codigosA].filter(c => codigosB.has(c)).length,
        totalHUB: huBSinDuplicadosEnRemitos.size
    };
}

function mostrarReporte(reporte) {
    ocultarTodo();
    const reporteArea = document.getElementById('reporteArea');
    reporteArea.classList.remove('hidden');

    let html = '<div class="report-section">';
    html += '<h2>📋 Reporte de Comparación</h2>';

    html += `<div class="status-info status-message">`;
    html += `<strong>Archivo A:</strong> ${reporte.nombreA}<br>`;
    html += `<strong>Archivo B:</strong> ${reporte.nombreB}`;
    html += `</div>`;

    html += '<h3>Análisis Archivo A (Etiquetas a Controlar)</h3>';
    html += `<p>En el archivo de Etiquetas a controlar "${reporte.nombreA}" hay ${reporte.totalCodigosA} código(s)`;

    if (reporte.duplicadosA.length > 0) {
        html += ` y se repiten los siguientes códigos:</p>`;
        html += '<div class="code-list"><ul>';
        reporte.duplicadosA.forEach(dup => {
            html += `<li><strong>${dup.codigo}</strong> - se repite ${dup.veces} veces</li>`;
        });
        html += '</ul></div>';
    } else {
        html += '.</p>';
    }

    html += '<h3>Resultado de la Comparación</h3>';

    if (!reporte.hayAnomalias) {
        html += `<div class="status-success status-message">`;
        html += `<p><strong>✅ Buenos días, se compararon los archivos "${reporte.nombreA}" y "${reporte.nombreB}"</strong></p>`;
        html += `<p>No se detectan anomalías, coinciden los códigos de Etiquetas a controlar con los códigos del Archivo de Remitos a Imprimir con un total de <strong>${reporte.totalRemitos} Remitos</strong> y <strong>${reporte.totalHUCoincidentes} Handling Units</strong>.</p>`;
        html += `</div>`;
    } else {
        html += `<div class="status-warning status-message">`;
        html += `<p><strong>⚠️ Buenos días, se compararon los archivos "${reporte.nombreA}" y "${reporte.nombreB}"</strong></p>`;
        html += `<p><strong>Se detectan anomalías:</strong></p>`;
        html += `</div>`;

        if (reporte.soloEnA.length > 0) {
            html += `<h4>Opción 1) Códigos en "${reporte.nombreA}" pero NO en "${reporte.nombreB}":</h4>`;
            html += '<div class="code-list"><ul>';
            reporte.soloEnA.forEach(codigo => {
                html += `<li>${codigo}</li>`;
            });
            html += '</ul></div>';
        }

        if (reporte.soloEnB.length > 0) {
            html += `<h4>Opción 2) Códigos en "${reporte.nombreB}" pero NO en "${reporte.nombreA}":</h4>`;
            html += '<div class="code-list"><ul>';
            reporte.soloEnB.forEach(codigo => {
                html += `<li>${codigo}</li>`;
            });
            html += '</ul></div>';
        }

        if (reporte.duplicadosEnRemitos.length > 0) {
            html += `<h4>Opción 3) Etiquetas que se repiten en más de un Remito:</h4>`;
            html += '<div class="code-list"><ul>';
            reporte.duplicadosEnRemitos.forEach(item => {
                html += `<li><strong>${item.handlingUnit}</strong> → Remitos: ${item.remitos.join(', ')}</li>`;
            });
            html += '</ul></div>';
        }
    }

    html += '</div>';
    reporteArea.innerHTML = html;
}

// --- Exportación a PDF ---------------------------------------------------
function exportarPDF() {
    if (!reporteGenerado) {
        alert('Primero debe realizar la comparación');
        return;
    }

    const {
        jsPDF
    } = window.jspdf;
    const doc = new jsPDF();

    const pageWidth = doc.internal.pageSize.getWidth();
    let y = 20;

    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    doc.text('Control Etiquetas/Remitos Despacho', pageWidth / 2, y, {
        align: 'center'
    });
    doc.text('RDC Troncal Rosario', pageWidth / 2, y + 8, {
        align: 'center'
    });

    y += 25;
    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');
    doc.text(`Archivo A: ${reporteGenerado.nombreA}`, 15, y);
    y += 6;
    doc.text(`Archivo B: ${reporteGenerado.nombreB}`, 15, y);
    y += 6;
    doc.text(`Fecha: ${new Date().toLocaleString('es-AR')}`, 15, y);

    y += 15;
    doc.setFontSize(12);
    doc.setFont(undefined, 'bold');
    doc.text(`Análisis Archivo A (Etiquetas a Controlar)`, 15, y);

    y += 8;
    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');
    let textoA = `En el archivo de Etiquetas a controlar "${reporteGenerado.nombreA}" hay ${reporteGenerado.totalCodigosA} código(s)`;

    if (reporteGenerado.duplicadosA.length > 0) {
        doc.text(textoA + ' y se repiten los siguientes códigos:', 15, y, {
            maxWidth: 180
        });
        y += 12;
        reporteGenerado.duplicadosA.forEach(dup => {
            if (y > 270) {
                doc.addPage();
                y = 20;
            }
            doc.text(`  - ${dup.codigo}: se repite ${dup.veces} veces`, 20, y);
            y += 6;
        });
    } else {
        doc.text(textoA + '.', 15, y, {
            maxWidth: 180
        });
        y += 8;
    }

    y += 10;
    if (y > 250) {
        doc.addPage();
        y = 20;
    }

    doc.setFontSize(12);
    doc.setFont(undefined, 'bold');
    doc.text('Resultado de la Comparación', 15, y);
    y += 10;

    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');

    if (!reporteGenerado.hayAnomalias) {
        const textoExito = `Buenos días, se compararon los archivos "${reporteGenerado.nombreA}" y "${reporteGenerado.nombreB}". No se detectan anomalías, coinciden los códigos de Etiquetas a controlar con los códigos del Archivo de Remitos a Imprimir con un total de ${reporteGenerado.totalRemitos} Remitos y ${reporteGenerado.totalHUCoincidentes} Handling Units.`;
        doc.text(textoExito, 15, y, {
            maxWidth: 180
        });
    } else {
        doc.text(`Buenos días, se compararon los archivos "${reporteGenerado.nombreA}" y "${reporteGenerado.nombreB}".`, 15, y, {
            maxWidth: 180
        });
        y += 12;
        doc.setFont(undefined, 'bold');
        doc.text('Se detectan anomalías:', 15, y);
        y += 10;
        doc.setFont(undefined, 'normal');

        if (reporteGenerado.soloEnA.length > 0) {
            if (y > 250) {
                doc.addPage();
                y = 20;
            }
            doc.setFont(undefined, 'bold');
            doc.text(`Opción 1) Códigos en "${reporteGenerado.nombreA}" pero NO en "${reporteGenerado.nombreB}":`, 15, y, {
                maxWidth: 180
            });
            y += 8;
            doc.setFont(undefined, 'normal');
            reporteGenerado.soloEnA.forEach(codigo => {
                if (y > 280) {
                    doc.addPage();
                    y = 20;
                }
                doc.text(`  - ${codigo}`, 20, y);
                y += 6;
            });
            y += 5;
        }

        if (reporteGenerado.soloEnB.length > 0) {
            if (y > 250) {
                doc.addPage();
                y = 20;
            }
            doc.setFont(undefined, 'bold');
            doc.text(`Opción 2) Códigos en "${reporteGenerado.nombreB}" pero NO en "${reporteGenerado.nombreA}":`, 15, y, {
                maxWidth: 180
            });
            y += 8;
            doc.setFont(undefined, 'normal');
            reporteGenerado.soloEnB.forEach(codigo => {
                if (y > 280) {
                    doc.addPage();
                    y = 20;
                }
                doc.text(`  - ${codigo}`, 20, y);
                y += 6;
            });
            y += 5;
        }

        if (reporteGenerado.duplicadosEnRemitos.length > 0) {
            if (y > 250) {
                doc.addPage();
                y = 20;
            }
            doc.setFont(undefined, 'bold');
            doc.text('Opción 3) Etiquetas que se repiten en más de un Remito:', 15, y, {
                maxWidth: 180
            });
            y += 8;
            doc.setFont(undefined, 'normal');
            reporteGenerado.duplicadosEnRemitos.forEach(item => {
                if (y > 275) {
                    doc.addPage();
                    y = 20;
                }
                doc.text(`  - ${item.handlingUnit} → Remitos: ${item.remitos.join(', ')}`, 20, y, {
                    maxWidth: 170
                });
                y += 6;
            });
        }
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    doc.save(`Reporte_Control_Etiquetas_${timestamp}.pdf`);
}

// --- Limpieza ------------------------------------------------------------
function limpiarDatos() {
    if (confirm('¿Está seguro de que desea limpiar todos los datos?')) {
        datosA = null;
        datosB = null;
        nombreArchivoA = '';
        nombreArchivoB = '';
        reporteGenerado = null;

        document.getElementById('fileA').value = '';
        document.getElementById('fileB').value = '';
        document.getElementById('statusA').innerHTML = '';
        document.getElementById('statusB').innerHTML = '';
        document.getElementById('btnComparar').disabled = true;
        document.getElementById('btnExportar').disabled = true;

        ocultarTodo();
        document.getElementById('welcomeMessage').classList.remove('hidden');
    }
}

function mostrarEstado(id, mensaje, tipo) {
    const elemento = document.getElementById(id);
    elemento.innerHTML = '';

    if (mensaje) {
        const div = document.createElement('div');
        div.className = `status-message status-${tipo}`;
        div.innerHTML = mensaje;
        elemento.appendChild(div);
    }
}

// --- Registro del Service Worker (PWA) -----------------------------------
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./service-worker.js').catch((err) => {
            console.error('Error al registrar Service Worker:', err);
        });
    });
}
