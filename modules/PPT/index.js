let currentSlideIndex = 0;
let slides = [];

function buildTableFromMatrix(matrix) {
    if (!matrix || matrix.length === 0) return "";
    let html = '<table class="table table-bordered table-striped table-hover text-center mb-0"><thead class="table-dark"><tr>';
    matrix[0].forEach(h => html += `<th class="align-middle">${h}</th>`);
    html += '</tr></thead><tbody>';
    
    for(let i = 1; i < matrix.length; i++) {
        let row = matrix[i];
        let rowStr = String(row[0]).toUpperCase();
        let isTotal = rowStr.includes("TOTAL") || rowStr.includes("CUMPLIMIENTO");
        
        let rowClass = isTotal ? 'table-warning fw-bold' : '';
        html += `<tr class="${rowClass}">`;
        
        row.forEach((val, idx) => {
            let text = val;
            if (String(matrix[0][idx]).includes("%") && !isNaN(val) && val !== "") {
                text = (parseFloat(val) * 100).toFixed(1) + "%";
            } else if (!isNaN(val) && val !== "") {
                text = Math.round(val);
            }
            html += `<td class="align-middle">${text}</td>`;
        });
        html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
}

function getHtmlTableComparativo() {
    if (!window.ejecutivosDetalleData || Object.keys(window.ejecutivosDetalleData).length === 0) return "<p>No hay datos disponibles.</p>";
    let allZones = ["1.- NORTE", "2.- SERENA/VIÑA", "3.- SANTIAGO", "4.- COSTA/RANCAGUA", "5.- CENTRO SUR", "6.- SUR"];
    let matrix = [["Ejecutivo", ...allZones.map(z => z.replace(/^[0-9\s.:-]+/, "").trim()), "Total Faltan"]];
    let data = window.ejecutivosDetalleData;
    let totals = new Array(allZones.length).fill(0);
    let granTotal = 0;

    for (let exec in data) {
        let row = [exec];
        let totalExec = 0;
        allZones.forEach((z, idx) => {
            let val = data[exec].zonas[z] ? Math.round(data[exec].zonas[z].faltan) : 0;
            row.push(val);
            totalExec += val;
            totals[idx] += val;
        });
        row.push(totalExec);
        granTotal += totalExec;
        matrix.push(row);
    }
    let totalRow = ["TOTAL GENERAL"];
    totals.forEach(t => totalRow.push(t));
    totalRow.push(granTotal);
    matrix.push(totalRow);
    return buildTableFromMatrix(matrix);
}

function getHtmlTableDrilldown(execName) {
    if (!window.tablaGestorCompleta) return "";
    let filteredData = [window.tablaGestorCompleta[0]];
    for(let i = 1; i < window.tablaGestorCompleta.length; i++) {
        let rowData = window.tablaGestorCompleta[i];
        if (rowData[0] === execName || String(rowData[0]).toUpperCase().includes(`TOTAL ${execName.toUpperCase()}`)) {
            filteredData.push(rowData);
        }
    }
    return buildTableFromMatrix(filteredData);
}

function getHtmlFromDomTable(tableId) {
    let originalTable = document.getElementById(tableId);
    if (!originalTable) return "<p class='text-danger fw-bold'>No se pudo cargar la tabla.</p>";
    
    let clone = originalTable.cloneNode(true);
    clone.id = ""; 
    clone.classList.add("table-bordered", "mb-0");
    let thead = clone.querySelector('thead');
    if (thead) {
        thead.className = "table-dark";
        thead.querySelectorAll('th').forEach(th => th.style.backgroundColor = '');
    }
    clone.querySelectorAll('.btn-faltan-export').forEach(s => s.remove());
    return clone.outerHTML;
}

function getHtmlTableLimpiaZona() {
    if (!window.ejecutivosDetalleData) return "";
    let zonasList = ["1.- NORTE", "2.- SERENA/VIÑA", "3.- SANTIAGO", "4.- COSTA/RANCAGUA", "5.- CENTRO SUR", "6.- SUR"];
    let matrix = [["Zona", "Gestor", "Asignado", "# Ganados", "Faltan", "% Faltante"]];
    let raw = window.ejecutivosDetalleData;
    let totGenAsigZ = 0, totGenGanZ = 0, totGenFaltanZ = 0;

    for (let z of zonasList) {
        let asigZ = 0, ganZ = 0, faltanZ = 0;
        let zName = z.replace(/^[0-9\s.:-]+/, "").trim();
        for (let g of Object.keys(raw).sort()) {
            if (raw[g].zonas[z]) {
                let d = raw[g].zonas[z];
                let fNeto = Math.round(d.faltan);
                let as = Math.round(d.asignados); 
                let gn = Math.round(d.ganados);
                if (as > 0 || gn > 0 || fNeto !== 0) {
                    let pctF = as > 0 ? fNeto/as : 0;
                    matrix.push([zName, g, as, gn, fNeto, pctF]);
                    asigZ += as; ganZ += gn; faltanZ += fNeto;
                }
            }
        }
        if (asigZ > 0 || ganZ > 0 || faltanZ !== 0) {
            let pctFaltanteZ = asigZ > 0 ? (faltanZ / asigZ) : 0;
            matrix.push([`TOTAL ${zName.toUpperCase()}`, "", asigZ, ganZ, faltanZ, pctFaltanteZ]);
            totGenAsigZ += asigZ; totGenGanZ += ganZ; totGenFaltanZ += faltanZ;
        }
    }
    matrix.push(["TOTAL GENERAL", "", totGenAsigZ, totGenGanZ, totGenFaltanZ, totGenAsigZ > 0 ? totGenFaltanZ/totGenAsigZ : 0]);
    return buildTableFromMatrix(matrix);
}

export function iniciarPresentacion() {
    const overlay = document.getElementById('presentationOverlay');
    slides = []; 
    
    let bgCoverEndEl = document.querySelector('input[name="bgCoverEnd"]:checked');
    let bgTitlesEl = document.querySelector('input[name="bgTitles"]:checked');
    let bgInfoEl = document.querySelector('input[name="bgInfo"]:checked');

    let bgCoverEnd = bgCoverEndEl ? bgCoverEndEl.value : '';
    let bgTitles = bgTitlesEl ? bgTitlesEl.value : '';
    let bgInfo = bgInfoEl ? bgInfoEl.value : '';

    slides.push({ type: 'cover', text: '', bg: bgCoverEnd });

    // --- CRECCU VIG ---
    if (document.getElementById('chkModuloVig') && document.getElementById('chkModuloVig').checked) {
        if (!window.ejecutivosDetalleData || Object.keys(window.ejecutivosDetalleData).length === 0) {
            if (window.mostrarAlerta) window.mostrarAlerta("No hay datos cargados para Creccu Vig. Sube el archivo primero.", "warning");
            return;
        }

        slides.push({ type: 'title', text: 'Creccu Vig', bg: bgTitles });

        if (document.getElementById('chkVigZonas').checked) {
            slides.push({ type: 'chart_with_table', title: 'Resumen por Zonas', chartType: 'zonas', tableHTML: getHtmlFromDomTable('tablaZonas'), bg: bgInfo });
        }
        if (document.getElementById('chkVigComp').checked) {
            slides.push({ type: 'chart_with_table', title: 'Comparativo de Faltantes por Zona', chartType: 'comparativo', tableHTML: getHtmlTableComparativo(), bg: bgInfo });
        }
        if (document.getElementById('chkVigExec').checked) {
            slides.push({ type: 'chart_with_table', title: 'Resumen por Ejecutivos', chartType: 'exec', tableHTML: getHtmlFromDomTable('tablaEjecutivos'), bg: bgInfo });
        }
        if (document.getElementById('chkVigDrill').checked) {
            for (let exec in window.ejecutivosDetalleData) {
                slides.push({ type: 'chart_with_table', title: `Desglose Ejecutivo: ${exec}`, chartType: 'drilldown', execName: exec, tableHTML: getHtmlTableDrilldown(exec), bg: bgInfo });
            }
        }
        if (document.getElementById('chkVigGestor').checked) {
            slides.push({ type: 'table_only', title: 'Resumen Tabla (Ver por Gestor)', tableHTML: buildTableFromMatrix(window.tablaGestorCompleta), bg: bgInfo });
        }
        if (document.getElementById('chkVigZona').checked) {
            slides.push({ type: 'table_only', title: 'Resumen Tabla (Ver por Zona)', tableHTML: getHtmlTableLimpiaZona(), bg: bgInfo });
        }
    }

    // --- CRECCU 1EDC ---
    if (document.getElementById('chkModulo1EDC') && document.getElementById('chkModulo1EDC').checked) {
        if (!window.main1EDCChartDataObj) {
            if (window.mostrarAlerta) window.mostrarAlerta("No hay datos cargados para Creccu 1EDC. Sube el archivo primero.", "warning");
            return;
        }

        slides.push({ type: 'title', text: 'Creccu 1EDC', bg: bgTitles });

        if (document.getElementById('chk1EDCZonas') && document.getElementById('chk1EDCZonas').checked) {
            slides.push({ type: 'chart_with_table', title: 'Resumen por Zonas 1EDC', chartType: 'zonas1EDC', tableHTML: getHtmlFromDomTable('tablaZonas1EDC'), bg: bgInfo });
        }
        if (document.getElementById('chk1EDCContact') && document.getElementById('chk1EDCContact').checked) {
            slides.push({ type: 'chart_with_table', title: 'Contactabilidad 1EDC', chartType: 'contact1EDC', tableHTML: getHtmlFromDomTable('tablaContact1EDC'), bg: bgInfo });
        }
    }

    slides.push({ type: 'end', text: '', bg: bgCoverEnd });

    if (slides.length <= 2) {
        if (window.mostrarAlerta) window.mostrarAlerta("Debes seleccionar al menos un componente de información.", "warning");
        return;
    }

    overlay.style.display = 'flex';
    currentSlideIndex = 0;
    renderSlide();
}

function renderSlide() {
    const overlay = document.getElementById('presentationOverlay');
    const container = document.getElementById('slideContent');
    const counter = document.getElementById('slideCounter');
    const slide = slides[currentSlideIndex];
    
    container.innerHTML = "";
    counter.innerText = `Slide ${currentSlideIndex + 1} / ${slides.length}`;

    if (slide.bg && slide.bg !== "") {
        overlay.style.backgroundImage = `url('${slide.bg}')`;
        overlay.style.backgroundSize = 'cover';
        overlay.style.backgroundPosition = 'center';
    } else {
        overlay.style.backgroundImage = 'none';
        overlay.style.backgroundColor = '#f1f3f5'; 
    }

    if (slide.type === 'cover' || slide.type === 'end') {
        container.innerHTML = ``;
    } 
    else if (slide.type === 'title') {
        container.innerHTML = `<div class="slide-title-card"><h1>${slide.text}</h1></div>`;
    }
    else if (slide.type === 'chart_with_table') {
        
        const header = document.createElement('h2');
        header.className = "fw-bold text-dark mb-3 slide-info-title";
        header.innerText = slide.title;
        container.appendChild(header);

        const toggleBtn = document.createElement('button');
        toggleBtn.className = "btn btn-dark fw-bold mb-3 shadow-sm";
        toggleBtn.innerHTML = "📄 Mostrar Tabla de Datos";
        container.appendChild(toggleBtn);

        const contentWrapper = document.createElement('div');
        contentWrapper.className = "w-100 position-relative d-flex justify-content-center flex-column align-items-center";
        
        // Ajustamos la altura si es la gráfica de contacto que trae tarjetas
        if (slide.chartType === 'contact1EDC' && window.main1EDCContactCardsHtml) {
            contentWrapper.style.height = "70vh";
            
            const cardsDiv = document.createElement('div');
            cardsDiv.className = "row w-100 mb-2 justify-content-center text-center";
            cardsDiv.innerHTML = window.main1EDCContactCardsHtml;
            contentWrapper.appendChild(cardsDiv);

            const chartContainer = document.createElement('div');
            chartContainer.className = "presentation-chart-box w-100 d-flex justify-content-center";
            chartContainer.style.height = "40vh"; // Reducido para que quepan las tarjetas
            const canvas = document.createElement('canvas');
            canvas.id = "presentationCanvas";
            chartContainer.appendChild(canvas);
            contentWrapper.appendChild(chartContainer);
        } else {
            contentWrapper.style.height = "65vh";
            const chartContainer = document.createElement('div');
            chartContainer.className = "presentation-chart-box w-100 h-100 d-flex justify-content-center";
            const canvas = document.createElement('canvas');
            canvas.id = "presentationCanvas";
            chartContainer.appendChild(canvas);
            contentWrapper.appendChild(chartContainer);
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = "presentation-table-box bg-white w-75 p-3 position-absolute top-0 start-50 translate-middle-x";
        tableContainer.style.display = "none";
        tableContainer.style.zIndex = "10";
        tableContainer.style.height = "100%";
        tableContainer.innerHTML = slide.tableHTML;
        contentWrapper.appendChild(tableContainer);

        container.appendChild(contentWrapper);

        setTimeout(() => {
            if (slide.chartType === 'drilldown') window.renderDrillDownChart(slide.execName, 'presentationCanvas', true);
            else if (slide.chartType === 'comparativo') window.renderFaltanComparisonChart('presentationCanvas', true);
            else if (slide.chartType === 'zonas') window.renderMainZonasChart('presentationCanvas', true);
            else if (slide.chartType === 'exec') window.renderMainExecChart('presentationCanvas', true);
            else if (slide.chartType === 'zonas1EDC') window.render1EDCZonasChart('presentationCanvas', true);
            else if (slide.chartType === 'contact1EDC') window.render1EDCContactChart('presentationCanvas', true);
        }, 50);

        let isShowingTable = false;
        toggleBtn.onclick = () => {
            isShowingTable = !isShowingTable;
            if (isShowingTable) {
                tableContainer.style.display = "block";
                toggleBtn.innerHTML = "📊 Volver a Gráfica";
                toggleBtn.className = "btn btn-danger fw-bold mb-3 shadow-sm";
            } else {
                tableContainer.style.display = "none";
                toggleBtn.innerHTML = "📄 Mostrar Tabla de Datos";
                toggleBtn.className = "btn btn-dark fw-bold mb-3 shadow-sm";
            }
        };

    } 
    else if (slide.type === 'table_only') {
        const header = document.createElement('h2');
        header.className = "fw-bold text-dark mb-4 slide-info-title";
        header.innerText = slide.title;
        container.appendChild(header);

        const tableContainer = document.createElement('div');
        tableContainer.className = "presentation-table-box";
        tableContainer.style.display = "block";
        tableContainer.innerHTML = slide.tableHTML;
        container.appendChild(tableContainer);
    }
}

window.nextSlide = () => {
    if (currentSlideIndex < slides.length - 1) {
        currentSlideIndex++;
        renderSlide();
    }
};

window.prevSlide = () => {
    if (currentSlideIndex > 0) {
        currentSlideIndex--;
        renderSlide();
    }
};

window.cerrarPresentacion = () => {
    document.getElementById('presentationOverlay').style.display = 'none';
    document.getElementById('presentationOverlay').style.backgroundImage = 'none'; 
    document.getElementById('slideContent').innerHTML = ""; 
    
    setTimeout(() => {
        window.dispatchEvent(new Event('resize'));
        if (window.currentZonasChart === 'main') window.renderMainZonasChart();
        else if (window.currentZonasChart === 'comparative') window.renderFaltanComparisonChart('graficoZonas');
        
        if (window.currentExecChart === 'main') window.toggleDrillDown('main');
        else window.toggleDrillDown(window.currentExecChart);

        if (window.render1EDCZonasChart) window.render1EDCZonasChart();
        if (window.render1EDCContactChart) window.render1EDCContactChart();
    }, 100);
};