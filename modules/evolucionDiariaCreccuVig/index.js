import { renderResumenTabla } from './ResumenTabla.js';
import { procesarResumenZonas } from './ResumenZona.js';
import { procesarResumenEjecutivos } from './ResumenEjecutivo.js';

export function normalizeZone(str) {
    let s = String(str).toUpperCase();
    if (s.includes("NORTE")) return "1.- NORTE";
    if (s.includes("SERENA") || s.includes("VIÑA")) return "2.- SERENA/VIÑA";
    if (s.includes("SANTIAGO")) return "3.- SANTIAGO";
    if (s.includes("COSTA") || s.includes("RANCAGUA")) return "4.- COSTA/RANCAGUA";
    if (s.includes("CENTRO")) return "5.- CENTRO SUR";
    if (s.includes("SUR")) return "6.- SUR"; 
    return "DESCONOCIDO";
}

export function normalizeGestor(str) {
    let s = String(str).toUpperCase();
    if (s.includes("XIOMARA")) return "Xiomara";
    if (s.includes("MARISELLE") || s.includes("MARISELLA")) return "Mariselle";
    if (s.includes("ROSA") && s.includes("L")) return "Rosa L.";
    if (s.includes("ROSA") && s.includes("A")) return "Rosa A.";
    return "DESCONOCIDO";
}

export function procesarEvolucionDiaria(data) {
    window.ejecutivosDetalleData = {}; 
    window.mainExecChartData = [];
    window.currentExecChart = 'main'; 
    window.currentZonasChart = 'main';

    let t1 = { row: -1, z: -1, g: -1, a: -1, gan: -1 };
    let t2 = { row: -1, g: -1, z: -1, a: -1, gan: -1, faltan: -1, mrut: -1, meta: -1, av: -1 };
    let t3 = { row: -1, z: -1, meta: -1, av: -1, f: -1 };

    // 1. Detección Inteligente de Tablas y Columnas
    for (let i = 0; i < data.length; i++) {
        let r = data[i].map(c => String(c).toUpperCase().trim());
        
        // TABLA 1
        if (t1.row === -1 && r.some(x => x.includes("ZONA")) && r.some(x => x.includes("GESTOR") || x.includes("EJECUTIVO"))) {
            t1.row = i + 1;
            t1.z = r.findIndex(x => x.includes("ZONA"));
            t1.g = r.findIndex(x => x.includes("GESTOR") || x.includes("EJECUTIVO"));
            t1.a = r.findIndex(x => x.includes("CUENTA") && x.includes("RUT"));
            if(t1.a === -1) t1.a = r.findIndex(x => x.includes("ASIGNADO"));
            t1.gan = r.findIndex(x => x.includes("ID") || x.includes("GANADO"));
        }
        
        // TABLA 2
        if (t2.row === -1 && r.includes("FALTA RUT POR PAGAR")) {
            t2.row = i + 1;
            t2.g = r.findIndex(x => x.includes("GESTOR") || x.includes("EJECUTIVO"));
            t2.z = r.findIndex(x => x.includes("DETALLE POR GESTOR") || x.includes("ZONA"));
            t2.a = r.findIndex(x => x.includes("CUENTA") && x.includes("RUT"));
            if(t2.a === -1) t2.a = r.findIndex(x => x.includes("ASIGNADO"));
            t2.gan = r.findIndex(x => x.includes("GANADO") || x.includes("ID"));
            t2.faltan = r.indexOf("FALTA RUT POR PAGAR");
            t2.mrut = r.findIndex(x => x.includes("META RUT"));
            t2.meta = r.indexOf("META");
            t2.av = r.findIndex(x => x.includes("AVANCE"));
        }
        
        // TABLA 3
        if (t3.row === -1 && r.includes("FALTANTE") && r.includes("META") && (r.includes("DETALLE POR GESTOR") || r.includes("ZONA"))) {
            if (t2.row === -1 || i > t2.row + 3) { 
                t3.row = i + 1;
                t3.z = r.findIndex(x => x.includes("DETALLE POR GESTOR") || x.includes("ZONA"));
                t3.meta = r.indexOf("META");
                t3.av = r.findIndex(x => x.includes("AVANCE"));
                t3.f = r.indexOf("FALTANTE");
            }
        }
    }

    if (t2.row === -1) {
        alert("Formato no reconocido. No se detectó la tabla de faltantes (Tabla 2).");
        return;
    }

    let raw_data = {}; 
    
    // 2. Extraer T1
    if (t1.row !== -1) {
        let lastZona = "DESCONOCIDO";
        for(let i = t1.row; i < data.length; i++) {
            let r = data[i]; if(!r) continue;
            let zRaw = String(r[t1.z]||"").trim();
            let gRaw = String(r[t1.g]||"").trim();
            if(zRaw && zRaw.toLowerCase() !== "nan") lastZona = zRaw;
            
            if(zRaw.toUpperCase().includes("TOTAL") || gRaw.toUpperCase().includes("TOTAL")) continue;
            
            let zNorm = normalizeZone(lastZona);
            let gNorm = normalizeGestor(gRaw);
            
            if(gNorm !== "DESCONOCIDO" && zNorm !== "DESCONOCIDO") {
                if(!raw_data[gNorm]) raw_data[gNorm] = {};
                if(!raw_data[gNorm][zNorm]) raw_data[gNorm][zNorm] = {asig:0, gan:0, faltan:0, mrut:0};
                
                if (t1.a !== -1) raw_data[gNorm][zNorm].asig = Math.round(parseFloat(r[t1.a])||0);
                if (t1.gan !== -1) raw_data[gNorm][zNorm].gan = Math.round(parseFloat(r[t1.gan])||0);
            }
            if(t2.row !== -1 && i >= t2.row - 2) break;
        }
    }

    // 3. Extraer T2 (Incluyendo captura exacta de META y sobreescritura de Asig/Ganados)
    let gestor_exact_meta = {};
    let total_general_meta = 0;

    if (t2.row !== -1) {
        let currentGestor = "DESCONOCIDO";
        for(let i = t2.row; i < data.length; i++) {
            let r = data[i]; if(!r) continue;
            let gRaw = String(r[t2.g]||"").trim();
            let zRaw = String(r[t2.z]||"").trim();

            if (gRaw.toUpperCase().startsWith("TOTAL ")) {
                if (gRaw.toUpperCase().includes("GENERAL")) {
                    if (t2.meta !== -1) total_general_meta = parseFloat(r[t2.meta]) || 0;
                } else {
                    let gN = normalizeGestor(gRaw.replace(/TOTAL/i, "").trim());
                    if (t2.meta !== -1) gestor_exact_meta[gN] = parseFloat(r[t2.meta]) || 0;
                }
            } else {
                if (gRaw && gRaw.toLowerCase()!=="nan") {
                    currentGestor = normalizeGestor(gRaw);
                    if (t2.meta !== -1) {
                        let mVal = parseFloat(r[t2.meta]);
                        if (!isNaN(mVal) && mVal > 0) gestor_exact_meta[currentGestor] = mVal;
                    }
                }
                if(zRaw && zRaw.toLowerCase()!=="nan" && !zRaw.toUpperCase().includes("TOTAL")) {
                    let zNorm = normalizeZone(zRaw);
                    let gNorm = currentGestor;
                    if(gNorm !== "DESCONOCIDO" && zNorm !== "DESCONOCIDO") {
                        if(!raw_data[gNorm]) raw_data[gNorm] = {};
                        if(!raw_data[gNorm][zNorm]) raw_data[gNorm][zNorm] = {asig:0, gan:0, faltan:0, mrut:0};

                        if (t2.faltan !== -1) raw_data[gNorm][zNorm].faltan = Math.round(parseFloat(r[t2.faltan])||0);
                        if (t2.mrut !== -1) raw_data[gNorm][zNorm].mrut = Math.round(parseFloat(r[t2.mrut])||0);

                        // Prioridad Absoluta a Tabla 2 para Asignados y Ganados
                        if (t2.a !== -1 && String(r[t2.a]).trim() !== "" && String(r[t2.a]).toLowerCase() !== "nan") {
                            raw_data[gNorm][zNorm].asig = Math.round(parseFloat(r[t2.a])||0);
                        }
                        if (t2.gan !== -1 && String(r[t2.gan]).trim() !== "" && String(r[t2.gan]).toLowerCase() !== "nan") {
                            raw_data[gNorm][zNorm].gan = Math.round(parseFloat(r[t2.gan])||0);
                        }
                    }
                }
            }
            if(t3.row !== -1 && i >= t3.row - 2) break;
        }
    }
    
    // Backup en caso de que TOTAL GENERAL no estuviera explícito en el Excel
    if (total_general_meta === 0) {
        let metaVals = Object.values(gestor_exact_meta);
        if (metaVals.length > 0) total_general_meta = metaVals[0];
    }

    // 4. Extraer T3 (Porcentajes exactos de totales por zona)
    let t3_data = {};
    if (t3.row !== -1 && t3.z !== -1) {
        for (let i = t3.row; i < data.length; i++) {
            let row = data[i]; if (!row) continue;
            let zonaRaw = String(row[t3.z] || "").trim();
            if (!zonaRaw || zonaRaw.toLowerCase() === "nan") continue;

            let isTotalRow = zonaRaw.toUpperCase().includes("TOTAL");
            let av = t3.av !== -1 ? (parseFloat(row[t3.av]) || 0) : 0;
            let meta = t3.meta !== -1 ? (parseFloat(row[t3.meta]) || 0) : 0;
            let pctF = t3.f !== -1 ? (parseFloat(row[t3.f]) || 0) : 0;

            if (isTotalRow) {
                t3_data["TOTAL"] = { avance: av, meta: meta, pctFaltante: pctF };
                break; 
            } else {
                t3_data[normalizeZone(zonaRaw)] = { avance: av, meta: meta, pctFaltante: pctF };
            }
        }
    }

    // 5. Preparar data maestra (ejecutivosDetalleData) estrictamente redondeada
    for (let g in raw_data) {
        window.ejecutivosDetalleData[g] = { asignados: 0, ganados: 0, faltan: 0, zonas: {} };
        for (let z in raw_data[g]) {
            let d = raw_data[g][z];
            window.ejecutivosDetalleData[g].zonas[z] = {
                asignados: Math.round(d.asig),
                ganados: Math.round(d.gan),
                faltan: Math.round(d.faltan), 
                faltanClamped: Math.max(0, Math.round(d.faltan)),
                meta: Math.round(d.mrut),
                metaPct: d.asig > 0 ? d.mrut / d.asig : 0
            };
            window.ejecutivosDetalleData[g].asignados += Math.round(d.asig);
            window.ejecutivosDetalleData[g].ganados += Math.round(d.gan);
            window.ejecutivosDetalleData[g].faltan += Math.max(0, Math.round(d.faltan)); 
        }
    }

    // =========================================================
    // CONSOLIDACIÓN PARA EJECUTIVOS (T2_EXECS)
    // =========================================================
    let t2_execs = {};
    for (let g of Object.keys(raw_data)) {
        let asigE = 0, ganE = 0, faltanClampedE = 0, metaE = 0;
        for (let z of Object.keys(raw_data[g])) {
            asigE += Math.round(raw_data[g][z].asig);
            ganE += Math.round(raw_data[g][z].gan);
            faltanClampedE += Math.max(0, Math.round(raw_data[g][z].faltan)); 
            metaE += Math.round(raw_data[g][z].mrut);
        }
        if (asigE > 0 || ganE > 0) {
            t2_execs[g] = {
                asignados: asigE, ganados: ganE, faltan: faltanClampedE, metaRut: metaE,
                metaPct: gestor_exact_meta[g] !== undefined ? gestor_exact_meta[g] : (asigE > 0 ? metaE/asigE : 0),
                avancePct: asigE > 0 ? ganE/asigE : 0
            };
        }
    }

    // =========================================================
    // CONSOLIDACIÓN PARA ZONAS (ZONAS FINAL DATA)
    // =========================================================
    let zonasList = ["1.- NORTE", "2.- SERENA/VIÑA", "3.- SANTIAGO", "4.- COSTA/RANCAGUA", "5.- CENTRO SUR", "6.- SUR"];
    let zonasFinalData = [];
    
    for (let zNorm of zonasList) {
        let asigZ = 0, ganZ = 0, faltanNetoZ = 0, metaZ = 0;
        for (let g of Object.keys(raw_data)) {
            if (raw_data[g][zNorm]) {
                asigZ += Math.round(raw_data[g][zNorm].asig);
                ganZ += Math.round(raw_data[g][zNorm].gan);
                faltanNetoZ += Math.round(raw_data[g][zNorm].faltan); 
                metaZ += Math.round(raw_data[g][zNorm].mrut);
            }
        }
        
        if (asigZ > 0 || ganZ > 0 || faltanNetoZ !== 0) {
            let t3Obj = t3_data[zNorm];
            zonasFinalData.push({
                nombreMostrar: zNorm.replace(/^[0-9\s.:-]+/, "").trim(), 
                faltanOrig: faltanNetoZ, 
                ganados: ganZ, 
                metaGanados: metaZ, 
                asignados: asigZ,
                metaPct: t3Obj ? t3Obj.meta : (asigZ > 0 ? metaZ/asigZ : 0),
                avance: t3Obj ? t3Obj.avance : (asigZ > 0 ? ganZ/asigZ : 0),
                pctFaltante: t3Obj ? t3Obj.pctFaltante : (asigZ > 0 ? faltanNetoZ/asigZ : 0)
            });
        }
    }

    // =========================================================
    // TABLA LIMPÍA 1: POR GESTOR (Bloquea a 0 Antes de Sumar)
    // =========================================================
    let tablaGestorArr = [["Gestor", "Zona", "Asignado", "# Ganados", "Faltan", "% Faltante"]];
    let totGenAsig = 0, totGenGan = 0, totGenFaltan = 0;

    for (let g of Object.keys(raw_data).sort()) {
        let asigG = 0, ganG = 0, faltanG = 0;
        for (let z of Object.keys(raw_data[g]).sort()) {
            let d = raw_data[g][z];
            let fC = Math.max(0, Math.round(d.faltan)); // Obligado a 0 si es negativo
            let as = Math.round(d.asig); let gn = Math.round(d.gan);
            
            if (as > 0 || gn > 0 || fC !== 0) {
                let pctF = as > 0 ? fC/as : 0;
                let zName = z.replace(/^[0-9\s.:-]+/, "").trim(); // Limpia Prefijo de Zona
                tablaGestorArr.push([g, zName, as, gn, fC, pctF]);
                asigG += as; ganG += gn; faltanG += fC; 
            }
        }
        if (asigG > 0 || ganG > 0 || faltanG !== 0) {
            let pctFaltanteG = asigG > 0 ? (faltanG / asigG) : 0; // Cálculo preciso sobre el total ya ajustado a 0
            tablaGestorArr.push([`TOTAL ${g.toUpperCase()}`, "", asigG, ganG, faltanG, pctFaltanteG]);
            totGenAsig += asigG; totGenGan += ganG; totGenFaltan += faltanG;
        }
    }
    tablaGestorArr.push(["TOTAL GENERAL", "", totGenAsig, totGenGan, totGenFaltan, totGenAsig > 0 ? totGenFaltan/totGenAsig : 0]);
    window.tablaGestorCompleta = tablaGestorArr; 

    // =========================================================
    // TABLA LIMPÍA 2: POR ZONA (Mantiene Negativos para Sobrecumplimiento)
    // =========================================================
    let tablaZonaArr = [["Zona", "Gestor", "Asignado", "# Ganados", "Faltan", "% Faltante"]];
    let totGenAsigZ = 0, totGenGanZ = 0, totGenFaltanZ = 0;

    for (let z of zonasList) {
        let asigZ = 0, ganZ = 0, faltanZ = 0;
        let zName = z.replace(/^[0-9\s.:-]+/, "").trim(); // Limpia Prefijo de Zona
        
        for (let g of Object.keys(raw_data).sort()) {
            if (raw_data[g][z]) {
                let d = raw_data[g][z];
                let fNeto = Math.round(d.faltan); // Mantiene negativos para reflejar sobrecumplimiento
                let as = Math.round(d.asig); let gn = Math.round(d.gan);
                
                if (as > 0 || gn > 0 || fNeto !== 0) {
                    let pctF = as > 0 ? fNeto/as : 0;
                    tablaZonaArr.push([zName, g, as, gn, fNeto, pctF]);
                    asigZ += as; ganZ += gn; faltanZ += fNeto;
                }
            }
        }
        if (asigZ > 0 || ganZ > 0 || faltanZ !== 0) {
            let pctFaltanteZ = asigZ > 0 ? (faltanZ / asigZ) : 0; // Cálculo preciso sobre el total neto
            tablaZonaArr.push([`TOTAL ${zName.toUpperCase()}`, "", asigZ, ganZ, faltanZ, pctFaltanteZ]);
            totGenAsigZ += asigZ; totGenGanZ += ganZ; totGenFaltanZ += faltanZ;
        }
    }
    tablaZonaArr.push(["TOTAL GENERAL", "", totGenAsigZ, totGenGanZ, totGenFaltanZ, totGenAsigZ > 0 ? totGenFaltanZ/totGenAsigZ : 0]);

    renderResumenTabla(tablaGestorArr, tablaZonaArr);
    procesarResumenZonas(zonasFinalData, t3_data["TOTAL"]);
    procesarResumenEjecutivos(t2_execs, total_general_meta);

    let btnPdf = document.getElementById('btnPdf');
    let btnExcel = document.getElementById('btnExcel');
    if (btnPdf) btnPdf.disabled = false;
    if (btnExcel) btnExcel.disabled = false;
}

window.toggleDrillDown = function(execName) {
    if (execName === 'main' || window.currentExecChart === execName) {
        window.renderTable('tablaEjecutivos', window.tablaEjecutivosPrincipal, 'execMain');
        window.renderMainExecChart();
    } else {
        let filteredTable = [window.tablaGestorCompleta[0]];
        for(let i = 1; i < window.tablaGestorCompleta.length; i++) {
            let row = window.tablaGestorCompleta[i];
            if (row[0] === execName || row[0] === `TOTAL ${execName.toUpperCase()}`) {
                filteredTable.push(row);
            }
        }
        window.renderTable('tablaEjecutivos', filteredTable, 'execDrill');
        window.renderDrillDownChart(execName);
    }
};

window.renderTable = function(tableId, data, mode = 'default') {
    const thead = document.querySelector(`#${tableId} thead`);
    const tbody = document.querySelector(`#${tableId} tbody`);
    if(!thead || !tbody) return;
    thead.innerHTML = ""; tbody.innerHTML = "";

    let trHead = document.createElement("tr");
    for (let j = 0; j < data[0].length; j++) {
        let th = document.createElement("th"); th.textContent = data[0][j]; trHead.appendChild(th);
    }
    thead.appendChild(trHead);

    for (let i = 1; i < data.length; i++) {
        let tr = document.createElement("tr");
        let isTotalRow = String(data[i][0]).toUpperCase().includes("TOTAL") || String(data[i][0]).includes("Cumplimento");
        if (isTotalRow) tr.classList.add("fw-bold", "table-warning");

        if (mode === 'zonas' || mode === 'execMain' || mode === 'execDrill') {
            tr.style.cursor = "pointer";
            tr.title = mode === 'execDrill' ? "Clic para volver a vista general" : (isTotalRow ? "Clic para ver General/Comparativo" : "Clic para ver desglose");
            tr.onmouseover = () => { tr.style.boxShadow = "inset 0 0 0 9999px rgba(0,0,0,0.05)"; };
            tr.onmouseout = () => { tr.style.boxShadow = ""; };
        }

        for (let j = 0; j < data[0].length; j++) {
            let td = document.createElement("td");
            let valor = data[i][j];

            if (!isNaN(valor) && valor !== "") {
                let num = parseFloat(valor);
                if (String(data[0][j]).includes("%")) {
                    valor = (num * 100).toFixed(1) + "%";
                } else {
                    valor = Math.round(num); // SIN DECIMALES
                }
            }
            td.innerHTML = valor;

            if (mode === 'zonas') {
                if (isTotalRow && j === 1) { 
                    td.style.backgroundColor = "#f8d7da"; 
                    td.style.color = "#842029";
                    td.innerHTML = valor + " <span class='btn-faltan-export' title='Clic para alternar gráfico Comparativo' style='font-size: 0.9em;'>📊</span>";
                    td.onclick = (e) => { 
                        e.stopPropagation(); 
                        if (window.currentZonasChart === 'comparative') window.renderMainZonasChart();
                        else window.renderFaltanComparisonChart('graficoZonas'); 
                    };
                } else {
                    td.onclick = (e) => { e.stopPropagation(); window.renderMainZonasChart(); };
                }
            } else if (mode === 'execMain') {
                if (isTotalRow && j === 1) { 
                    td.style.backgroundColor = "#f8d7da"; 
                    td.style.color = "#842029";
                }
                td.onclick = (e) => {
                    e.stopPropagation();
                    if (isTotalRow) window.toggleDrillDown('main');
                    else window.toggleDrillDown(data[i][0]);
                };
            } else if (mode === 'execDrill') {
                td.onclick = (e) => { e.stopPropagation(); window.toggleDrillDown('main'); };
            }
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
};

window.drawOverlappingChart = function(canvasId, labels, dataObj, labelTitle, totalFaltan, isExport = false) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    let existingChart = Chart.getChart(canvasId);
    if (existingChart) existingChart.destroy();

    const colorFaltan = '#dc3545'; 
    const colorGanados = '#75b798'; 
    const colorMetas = '#6ea8fe'; 
    const colorAsignados = '#ced4da'; 

    const createDiagonalPattern = () => {
        let canvas = document.createElement('canvas');
        canvas.width = 16; canvas.height = 16;
        let pctx = canvas.getContext('2d');
        pctx.fillStyle = colorFaltan; 
        pctx.fillRect(0, 0, 16, 16);
        pctx.strokeStyle = 'rgba(255, 255, 255, 0.4)'; 
        pctx.lineWidth = 4;
        pctx.beginPath(); pctx.moveTo(-4, 12); pctx.lineTo(12, -4); pctx.moveTo(4, 20); pctx.lineTo(20, 4); pctx.stroke();
        return pctx.createPattern(canvas, 'repeat');
    };
    const patternFaltan = createDiagonalPattern();

    const mainChartLabelsPlugin = {
        id: 'mainChartLabelsPlugin_' + canvasId,
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            ctx.save();
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            
            data.datasets.forEach((dataset, i) => {
                const meta = chart.getDatasetMeta(i);
                if (meta.hidden) return;
                
                if (dataset.type === 'line' || dataset.label === '% Meta del mes') {
                    ctx.font = 'bold 12px Arial';
                    meta.data.forEach((point, index) => {
                        const pct = dataset.customMetaPct[index];
                        if (pct !== undefined) {
                            const text = (pct * 100).toFixed(1) + '%'; 
                            const textWidth = ctx.measureText(text).width;
                            const rectW = textWidth + 8;
                            const rectH = 18;
                            
                            let isLast = index === meta.data.length - 1;
                            let boxX = point.x + 35;
                            let boxY = point.y - rectH / 2;
                            if (isLast) { boxX = point.x - 35 - rectW; }
                            
                            ctx.beginPath();
                            ctx.moveTo(isLast ? boxX + rectW : boxX, boxY + rectH / 2);
                            ctx.lineTo(isLast ? point.x - 8 : point.x + 8, point.y);
                            ctx.strokeStyle = '#fd7e14';
                            ctx.lineWidth = 1.5;
                            ctx.stroke();

                            ctx.beginPath();
                            ctx.moveTo(isLast ? point.x - 2 : point.x + 2, point.y);
                            ctx.lineTo(isLast ? point.x - 10 : point.x + 10, point.y - 5);
                            ctx.lineTo(isLast ? point.x - 10 : point.x + 10, point.y + 5);
                            ctx.fillStyle = '#fd7e14';
                            ctx.fill();
                            
                            ctx.fillStyle = 'rgba(255, 255, 255, 0.95)';
                            ctx.fillRect(boxX, boxY, rectW, rectH);
                            ctx.strokeRect(boxX, boxY, rectW, rectH);
                            
                            ctx.fillStyle = '#000';
                            ctx.fillText(text, boxX + rectW/2, point.y);
                        }
                    });
                } else {
                    ctx.font = 'bold 14px Arial'; 
                    let tColor = '#000';
                    if (dataset.label === 'Faltan') tColor = '#b02a37'; 
                    else if (dataset.label === 'Ganados') tColor = '#146c43'; 
                    else if (dataset.label === 'Meta Ganados') tColor = '#0a58ca'; 
                    
                    meta.data.forEach((bar, index) => {
                        const val = Math.round(dataset.data[index]); // NUNCA DECIMALES
                        if (val !== 0 && !isNaN(val)) {
                            let yPos = bar.y + 16;
                            if (dataset.label === 'Faltan') {
                                yPos = val > 0 ? bar.y + 16 : bar.y - 16;
                            } else {
                                if (val > 0) {
                                    if (bar.base - bar.y < 25) yPos = bar.y - 12; 
                                } else {
                                    yPos = bar.y - 8;
                                    if (bar.y - bar.base < 25) yPos = bar.y + 16; 
                                }
                            }
                            ctx.strokeStyle = 'rgba(255, 255, 255, 1)';
                            ctx.lineWidth = 4;
                            ctx.strokeText(val, bar.x, yPos);
                            ctx.fillStyle = tColor;
                            ctx.fillText(val, bar.x, yPos);
                        }
                    });
                }
            });
            ctx.restore();
        }
    };

    const totalBadgePlugin = {
        id: 'totalBadgePlugin_' + canvasId,
        afterDraw(chart) {
            if (totalFaltan === undefined) return;
            const { ctx, chartArea } = chart;
            if (!chartArea) return;
            const drawRoundRect = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };
            
            let fVal = Math.round(totalFaltan);
            let isCumplida = fVal <= 0;
            const text = isCumplida ? '✔ Total Cumplido' : 'Total Faltan: ' + fVal;
            
            ctx.save(); ctx.font = 'bold 14px Arial';
            const textWidth = ctx.measureText(text).width;
            const padX = 12, height = 28; const rectW = textWidth + padX * 2;
            const posX = chartArea.left; const posY = chartArea.top - 35; 
            
            ctx.fillStyle = isCumplida ? 'rgba(25, 135, 84, 0.9)' : 'rgba(220, 53, 69, 0.9)'; 
            drawRoundRect(ctx, posX, posY, rectW, height, 6); ctx.fill();
            ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
            ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(text, posX + rectW / 2, posY + height / 2);
            ctx.restore();
        }
    };

    const miniBadgeTicksPlugin = {
        id: 'miniBadgeTicksPlugin_' + canvasId,
        afterDraw(chart) {
            const { ctx, scales: { x } } = chart;
            if (!dataObj.faltanText) return;
            ctx.save(); ctx.font = 'bold 12px Arial'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            const drawRoundRect = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };
            x.ticks.forEach((tick, index) => {
                if (!dataObj.faltanText[index]) return;
                const text = dataObj.faltanText[index];
                const xPos = x.getPixelForTick(index);
                const textWidth = ctx.measureText(text).width;
                const padX = 10, rectH = 24; const rectW = textWidth + padX * 2;
                const badgeX = xPos - rectW / 2; 
                const badgeY = x.top + 38; 
                
                const isFalta = text.includes('Falta:');
                ctx.fillStyle = isFalta ? colorFaltan : '#198754';
                drawRoundRect(ctx, badgeX, badgeY, rectW, rectH, 5); ctx.fill();
                ctx.fillStyle = '#ffffff'; ctx.fillText(text, xPos, badgeY + rectH / 2);
            });
            ctx.restore();
        }
    };

    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels, 
            datasets: [
                {
                    label: '% Meta del mes',
                    data: dataObj.metas, type: 'line', yAxisID: 'y', borderWidth: 3, fill: false,
                    borderColor: '#fd7e14', backgroundColor: '#fd7e14', pointBackgroundColor: '#fd7e14',
                    pointBorderColor: '#fff', pointRadius: 6, pointHoverRadius: 8, order: 0, customMetaPct: dataObj.metaPct
                },
                { label: 'Faltan', data: dataObj.faltan, backgroundColor: patternFaltan, yAxisID: 'y', grouped: false, barPercentage: 0.3, categoryPercentage: 0.8, order: 1, minBarLength: 35 },
                { label: 'Ganados', data: dataObj.ganados, backgroundColor: colorGanados, yAxisID: 'y', grouped: false, barPercentage: 0.5, categoryPercentage: 0.8, order: 2 },
                { label: 'Meta Ganados', data: dataObj.metas, backgroundColor: colorMetas, yAxisID: 'y', grouped: false, barPercentage: 0.7, categoryPercentage: 0.8, order: 3 },
                { label: 'Asignado', data: dataObj.asignados, backgroundColor: colorAsignados, yAxisID: 'y', grouped: false, barPercentage: 0.9, categoryPercentage: 0.8, order: 4 }
            ]
        },
        options: {
            animation: isExport ? false : true, 
            responsive: true, layout: { padding: { top: 25, right: 80 } }, 
            plugins: { title: { display: true, text: labelTitle, font: { size: 16 }, color: '#000000' }, legend: { position: 'bottom' } },
            scales: { 
                y: { type: 'linear', display: true, position: 'left', beginAtZero: true, grace: '0%' }, 
                x: { 
                    ticks: { color: '#000000', font: { size: 12, weight: 'bold' } },
                    afterFit: function(axis) { axis.height = 85; }
                } 
            }
        }, plugins: [totalBadgePlugin, miniBadgeTicksPlugin, mainChartLabelsPlugin] 
    });
};

export function drawSimpleChart(canvasId, labels, datasets, labelTitle, type = 'bar', badgeText = undefined, extraOptions = {}, extraPlugins = [], isExport = false) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    let existingChart = Chart.getChart(canvasId);
    if (existingChart) existingChart.destroy();

    const simpleBadgePlugin = {
        id: 'simpleBadgePlugin_' + canvasId,
        afterDraw(chart) {
            if (badgeText === undefined) return;
            const { ctx, chartArea } = chart;
            if (!chartArea) return;
            const drawRoundRect = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };
            
            const text = badgeText;
            const isCumplida = text.includes('✔') || text.includes('Cumplida');
            
            ctx.save(); ctx.font = 'bold 14px Arial';
            const textWidth = ctx.measureText(text).width;
            const padX = 12, height = 28; const rectW = textWidth + padX * 2;
            const posX = chartArea.left; const posY = chartArea.top - 35; 
            
            ctx.fillStyle = isCumplida ? 'rgba(25, 135, 84, 0.9)' : 'rgba(220, 53, 69, 0.9)';
            drawRoundRect(ctx, posX, posY, rectW, height, 6); ctx.fill();
            ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
            ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(text, posX + rectW / 2, posY + height / 2);
            ctx.restore();
        }
    };

    let pList = badgeText !== undefined ? [simpleBadgePlugin] : [];
    if (extraPlugins && extraPlugins.length > 0) pList = pList.concat(extraPlugins);

    let layoutPadding = { top: 25, bottom: 0, right: 80 };
    if (extraOptions && extraOptions.layout && extraOptions.layout.padding) {
        layoutPadding = { ...layoutPadding, ...extraOptions.layout.padding };
    }

    new Chart(ctx, {
        type: type, 
        data: { labels: labels, datasets: datasets },
        options: { 
            animation: isExport ? false : true, 
            responsive: true, 
            layout: { padding: layoutPadding }, 
            plugins: { title: { display: true, text: labelTitle, font: { size: 16 }, color: '#000000' }, legend: { position: 'bottom' } }, 
            scales: { 
                y: { beginAtZero: true, grace: '0%' }, 
                x: { ticks: { color: '#000000', font: { size: 12, weight: 'bold' } }, afterFit: (extraOptions?.scales?.x?.afterFit) } 
            } 
        },
        plugins: pList
    });
}

window.renderMainExecChart = function(targetCanvasId = 'graficoEjecutivos', isExport = false) {
    if (!isExport) window.currentExecChart = 'main'; 
    let labels = [], faltan = [], ganados = [], metas = [], asignados = [], avance = [], metaPct = [], faltanText = [];
    let tFaltanE = window.mainExecChartData[window.mainExecChartData.length - 1][1]; 
    for (let i = 0; i < window.mainExecChartData.length - 1; i++) {
        let execName = window.mainExecChartData[i][0];
        
        let pctFaltanteReal = window.mainExecChartData[i][7]; 
        let diffText = pctFaltanteReal > 0.0001 ? `Falta: ${(pctFaltanteReal * 100).toFixed(1)}%` : 'Meta Cumplida';
        
        labels.push([execName, "", "", " "]); faltanText.push(diffText);
        faltan.push(window.mainExecChartData[i][1]); ganados.push(window.mainExecChartData[i][2]); metas.push(window.mainExecChartData[i][3]); asignados.push(window.mainExecChartData[i][4]); metaPct.push(window.mainExecChartData[i][5]); avance.push(window.mainExecChartData[i][6]);
    }
    window.drawOverlappingChart(targetCanvasId, labels, { asignados: asignados, metas: metas, ganados: ganados, faltan: faltan, avance: avance, metaPct: metaPct, faltanText: faltanText }, 'Resumen por Ejecutivos (Avance vs Meta)', tFaltanE, isExport);
};

window.renderDrillDownChart = function(execName, targetCanvasId = 'graficoEjecutivos', isExport = false) {
    if (!isExport) window.currentExecChart = execName; 
    let detail = window.ejecutivosDetalleData[execName];
    if (!detail) return;
    
    let stdZones = ["1.- NORTE", "2.- SERENA/VIÑA", "3.- SANTIAGO", "4.- COSTA/RANCAGUA", "5.- CENTRO SUR", "6.- SUR"];
    let labels = [], ganados = [], metas = [], metaPcts = [];
    let customZoneStats = {};

    for (let zona of stdZones) { 
        let z = detail.zonas[zona]; 
        if (z) { 
            let lbl = zona.replace(/^[0-9\s.:-]+/, "").trim();
            labels.push(lbl); 
            ganados.push(z.ganados); 
            metas.push(z.meta); 
            customZoneStats[lbl] = { faltan: Math.max(0, Math.round(z.faltan)) }; 
            metaPcts.push(z.metaPct || 0);
        }
    }

    let exactTotalRow = window.mainExecChartData.find(row => row[0] === execName);
    let exactTotalFaltan = exactTotalRow ? Math.round(exactTotalRow[1]) : 0;

    const drilldownLabelsPlugin = {
        id: 'drilldownLabelsPlugin',
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            data.datasets.forEach((dataset, i) => {
                const meta = chart.getDatasetMeta(i);
                if (meta.hidden) return;

                if (dataset.type === 'line' || dataset.label.startsWith('Meta Zonas')) {
                    ctx.font = 'bold 12px Arial';
                    meta.data.forEach((point, index) => {
                        const pct = dataset.customMetaPct ? dataset.customMetaPct[index] : 0;
                        if (pct > 0) {
                            const text = (pct * 100).toFixed(1) + '%';
                            const textWidth = ctx.measureText(text).width;
                            const rectW = textWidth + 8; const rectH = 18;
                            let isLast = index === meta.data.length - 1;
                            let boxX = point.x + 35; let boxY = point.y - rectH / 2;
                            if (isLast) boxX = point.x - 35 - rectW;
                            
                            ctx.beginPath(); ctx.moveTo(isLast ? boxX + rectW : boxX, boxY + rectH / 2); ctx.lineTo(isLast ? point.x - 8 : point.x + 8, point.y); ctx.strokeStyle = '#fd7e14'; ctx.lineWidth = 1.5; ctx.stroke();
                            ctx.beginPath(); ctx.moveTo(isLast ? point.x - 2 : point.x + 2, point.y); ctx.lineTo(isLast ? point.x - 10 : point.x + 10, point.y - 5); ctx.lineTo(isLast ? point.x - 10 : point.x + 10, point.y + 5); ctx.fillStyle = '#fd7e14'; ctx.fill();
                            ctx.fillStyle = 'rgba(255, 255, 255, 0.95)'; ctx.fillRect(boxX, boxY, rectW, rectH); ctx.strokeRect(boxX, boxY, rectW, rectH);
                            ctx.fillStyle = '#000'; ctx.fillText(text, boxX + rectW/2, point.y);
                        }
                    });
                } else if (dataset.label.startsWith('Ganados')) {
                    ctx.font = 'bold 14px Arial';
                    meta.data.forEach((bar, index) => {
                        const val = Math.round(dataset.data[index]); 
                        if (val > 0) {
                            let yPos = bar.y + 16;
                            if (bar.base - bar.y < 25) yPos = bar.y - 12; 
                            ctx.strokeStyle = 'rgba(255, 255, 255, 1)'; ctx.lineWidth = 4; ctx.strokeText(val, bar.x, yPos);
                            ctx.fillStyle = '#146c43'; ctx.fillText(val, bar.x, yPos);
                        }
                    });
                }
            });
            ctx.restore();
        }
    };

    const drilldownBadgePlugin = {
        id: 'drilldownBadgePlugin',
        afterDraw(chart) {
            const { ctx, scales: { x } } = chart;
            ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            const drawRoundRect = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };
            
            x.ticks.forEach((tick, index) => {
                const lbl = chart.data.labels[index];
                const stats = customZoneStats[lbl];
                if (stats) {
                    const xPos = x.getPixelForTick(index); const badgeY = x.top + 38; 
                    const fVal = Math.round(stats.faltan); const isCumplida = fVal <= 0;
                    const t1 = isCumplida ? '✔ Meta Cumplida' : 'Faltan: ' + fVal;
                    ctx.font = 'bold 12px Arial'; let w1 = ctx.measureText(t1).width + 20; const rectH = 24;
                    ctx.fillStyle = isCumplida ? '#198754' : '#dc3545'; drawRoundRect(ctx, xPos - w1/2, badgeY, w1, rectH, 5); ctx.fill();
                    ctx.fillStyle = '#fff'; ctx.fillText(t1, xPos, badgeY + rectH / 2);
                }
            });
            ctx.restore();
        }
    };

    let extraOpts = { layout: { padding: { top: 25, bottom: 0, right: 80 } }, scales: { x: { afterFit: function(axis) { axis.height = 85; } } } };
    let fValBadge = exactTotalFaltan;
    let topBadgeText = fValBadge <= 0 ? '✔ Meta Cumplida' : 'Faltan de ' + execName + ': ' + fValBadge;

    drawSimpleChart(targetCanvasId, labels, [ 
        { label: 'Ganados (' + execName + ')', data: ganados, backgroundColor: '#75b798', borderWidth: 1 }, 
        { label: 'Meta Zonas (' + execName + ')', data: metas, type: 'line', borderColor: '#fd7e14', backgroundColor: '#fd7e14', borderWidth: 2, fill: false, customMetaPct: metaPcts } 
    ], `Desglose por Zonas: ${execName}`, 'bar', topBadgeText, extraOpts, [drilldownBadgePlugin, drilldownLabelsPlugin], isExport);
};

window.renderFaltanComparisonChart = function(targetCanvasId = 'graficoZonas', isExport = false) {
    if (!isExport) window.currentZonasChart = 'comparative'; 
    let allZones = new Set();
    let zonaStats = {};
    for (let exec in window.ejecutivosDetalleData) { 
        for (let zona in window.ejecutivosDetalleData[exec].zonas) {
            allZones.add(zona);
            if (!zonaStats[zona]) zonaStats[zona] = { faltan: 0, asignados: 0 };
            zonaStats[zona].faltan += window.ejecutivosDetalleData[exec].zonas[zona].faltan; 
            zonaStats[zona].asignados += window.ejecutivosDetalleData[exec].zonas[zona].asignados;
        } 
    }
    
    let stdZones = ["1.- NORTE", "2.- SERENA/VIÑA", "3.- SANTIAGO", "4.- COSTA/RANCAGUA", "5.- CENTRO SUR", "6.- SUR"];
    let rawLabels = stdZones.filter(z => zonaStats[z] !== undefined);
    let labels = rawLabels.map(z => z.replace(/^[0-9\s.:-]+/, "").trim()); 
    let customZoneStats = {};
    rawLabels.forEach((z, i) => { customZoneStats[labels[i]] = { faltan: Math.round(zonaStats[z].faltan) }; });

    let datasets = [], colors = ['rgba(255, 99, 132, 0.7)', 'rgba(54, 162, 235, 0.7)', 'rgba(255, 206, 86, 0.7)', 'rgba(153, 102, 255, 0.7)'], c = 0;
    let textColors = ['#b02a37', '#0a58ca', '#b58500', '#6f42c1'];

    for (let exec in window.ejecutivosDetalleData) {
        let dataArr = [];
        for (let zona of rawLabels) { dataArr.push(Math.round(window.ejecutivosDetalleData[exec].zonas[zona] ? window.ejecutivosDetalleData[exec].zonas[zona].faltan : 0)); }
        datasets.push({ label: exec, data: dataArr, backgroundColor: colors[c % colors.length], customTextColor: textColors[c % textColors.length], borderWidth: 1, minBarLength: 35 }); 
        c++;
    }

    const advFaltanPlugin = {
        id: 'advFaltanPlugin',
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            ctx.save(); ctx.font = 'bold 14px Arial'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            data.datasets.forEach((dataset, i) => {
                const meta = chart.getDatasetMeta(i);
                if (!meta.hidden) {
                    meta.data.forEach((bar, index) => {
                        const val = Math.round(dataset.data[index]); 
                        if (val !== 0 && !isNaN(val)) {
                            let yPos = val > 0 ? bar.y + 16 : bar.y - 16;
                            ctx.strokeStyle = 'rgba(255, 255, 255, 1)'; ctx.lineWidth = 4; ctx.strokeText(val, bar.x, yPos);
                            ctx.fillStyle = dataset.customTextColor; ctx.fillText(val, bar.x, yPos);
                        }
                    });
                }
            });
            ctx.restore();
        }
    };

    const drilldownBadgePlugin = {
        id: 'drilldownBadgePlugin',
        afterDraw(chart) {
            const { ctx, scales: { x } } = chart;
            ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            const drawRoundRect = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };
            
            x.ticks.forEach((tick, index) => {
                const lbl = chart.data.labels[index];
                const stats = customZoneStats[lbl];
                if (stats) {
                    const xPos = x.getPixelForTick(index); const badgeY = x.top + 38; 
                    const fVal = Math.round(stats.faltan); const isCumplida = fVal <= 0;
                    const t1 = isCumplida ? '✔ Meta Cumplida' : 'Faltan: ' + fVal;
                    ctx.font = 'bold 12px Arial'; let w1 = ctx.measureText(t1).width + 20; const rectH = 24;
                    ctx.fillStyle = isCumplida ? '#198754' : '#dc3545'; drawRoundRect(ctx, xPos - w1/2, badgeY, w1, rectH, 5); ctx.fill();
                    ctx.fillStyle = '#fff'; ctx.fillText(t1, xPos, badgeY + rectH / 2);
                }
            });
            ctx.restore();
        }
    };

    let extraOpts = { layout: { padding: { top: 25, bottom: 0, right: 30 } }, scales: { x: { afterFit: function(axis) { axis.height = 85; } } } };
    let totalFaltanCmp = 0;
    for (let exec in window.ejecutivosDetalleData) { for (let zona in window.ejecutivosDetalleData[exec].zonas) { totalFaltanCmp += window.ejecutivosDetalleData[exec].zonas[zona].faltan; } }
    let fValCmp = Math.round(totalFaltanCmp);
    let topBadgeTextCmp = fValCmp <= 0 ? '✔ Meta Cumplida' : 'Total Faltan: ' + fValCmp;

    drawSimpleChart(targetCanvasId, labels, datasets, 'Comparativo de Faltantes por Zona', 'bar', topBadgeTextCmp, extraOpts, [advFaltanPlugin, drilldownBadgePlugin], isExport); 
};

window.descargarExcel = function() {
    if (typeof XLSX === "undefined") { alert("Librería XLSX no cargada."); return; }
    let spansHidden = document.querySelectorAll('.btn-faltan-export'); let cacheSpans = [];
    spansHidden.forEach(s => { cacheSpans.push({ parent: s.parentNode, element: s }); s.remove(); });

    let wb = XLSX.utils.book_new();
    let ws1 = XLSX.utils.table_to_sheet(document.getElementById('tablaZonas'));
    let ws2 = XLSX.utils.table_to_sheet(document.getElementById('tablaEjecutivos'));
    let ws3 = XLSX.utils.table_to_sheet(document.getElementById('tablaLimpiaGestor'));
    let ws4 = XLSX.utils.table_to_sheet(document.getElementById('tablaLimpiaZona'));
    
    XLSX.utils.book_append_sheet(wb, ws1, "Resumen Zonas");
    XLSX.utils.book_append_sheet(wb, ws2, "Resumen Ejecutivos");
    XLSX.utils.book_append_sheet(wb, ws3, "Tabla Por Gestor");
    XLSX.utils.book_append_sheet(wb, ws4, "Tabla Por Zona");
    
    XLSX.writeFile(wb, "Dashboard_CRECCU.xlsx");
    cacheSpans.forEach(item => item.parent.appendChild(item.element));
};

window.descargarPDF = async function() {
    if (!window.jspdf) { alert("Librería jsPDF no cargada."); return; }
    const loader = document.getElementById('loader');
    if (loader) { loader.querySelector('h3').textContent = "Generando PDF por favor espere..."; loader.style.display = 'flex'; }
    let spansHidden = document.querySelectorAll('.btn-faltan-export'); let cacheSpans = [];
    spansHidden.forEach(s => { cacheSpans.push({ parent: s.parentNode, element: s }); s.remove(); });

    try {
        const { jsPDF } = window.jspdf; const doc = new jsPDF('p', 'pt', 'a4');
        const pageWidth = doc.internal.pageSize.getWidth(); const margin = 40; const pdfWidth = pageWidth - (margin * 2);

        const autoTableCleaner = { didParseCell: function(data) { if (data.cell.text && data.cell.text.length > 0) { data.cell.text = data.cell.text.map(t => t.replace('📊', '').trim()); } } };

        doc.setFontSize(16); doc.text("Resumen por Zonas", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaZonas', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [13, 110, 253] }, ...autoTableCleaner });

        let exportDiv = document.createElement('div'); exportDiv.style.width = '1000px'; exportDiv.style.height = '500px'; exportDiv.style.position = 'fixed'; exportDiv.style.top = '-9999px';
        let exportCanvas = document.createElement('canvas'); exportCanvas.id = 'exportCanvas'; exportCanvas.width = 1000; exportCanvas.height = 500; exportDiv.appendChild(exportCanvas); document.body.appendChild(exportDiv);

        const getChartImage = (renderFunc) => { return new Promise(resolve => { renderFunc(); setTimeout(() => resolve(exportCanvas.toDataURL('image/png', 1.0)), 300); }); };

        let imgZonas = await getChartImage(() => window.renderMainZonasChart('exportCanvas', true));
        let imgProps = doc.getImageProperties(imgZonas); let pdfHeight = (imgProps.height * pdfWidth) / imgProps.width; let finalY = doc.lastAutoTable.finalY + 20;
        doc.addImage(imgZonas, 'PNG', margin, finalY, pdfWidth, pdfHeight);

        doc.addPage(); doc.text("Comparativo de Faltantes por Zona", pageWidth/2, 40, {align: 'center'});
        let imgComp = await getChartImage(() => window.renderFaltanComparisonChart('exportCanvas', true));
        let imgPropsComp = doc.getImageProperties(imgComp); let pdfHeightComp = (imgPropsComp.height * pdfWidth) / imgPropsComp.width;
        doc.addImage(imgComp, 'PNG', margin, 60, pdfWidth, pdfHeightComp);

        doc.addPage(); doc.text("Resumen por Ejecutivos (General)", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaEjecutivos', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [25, 135, 84] }, ...autoTableCleaner });
        let finalY2 = doc.lastAutoTable.finalY + 20;
        let imgExecMain = await getChartImage(() => window.renderMainExecChart('exportCanvas', true));
        let imgProps2 = doc.getImageProperties(imgExecMain); let pdfHeight2 = (imgProps2.height * pdfWidth) / imgProps2.width;
        doc.addImage(imgExecMain, 'PNG', margin, finalY2, pdfWidth, pdfHeight2);

        for (let i = 0; i < window.mainExecChartData.length - 1; i++) {
            let execName = window.mainExecChartData[i][0];
            doc.addPage(); doc.text(`Desglose por Zonas: ${execName}`, pageWidth/2, 40, {align: 'center'});

            let filteredTable = [window.tablaGestorCompleta[0]];
            for(let j = 1; j < window.tablaGestorCompleta.length; j++) {
                let row = window.tablaGestorCompleta[j];
                if (row[0] === execName || row[0] === `TOTAL ${execName.toUpperCase()}`) {
                    let rowFormatted = row.map((val, idx) => {
                        if (String(window.tablaGestorCompleta[0][idx]).includes("%") && !isNaN(val) && val !== "") return (parseFloat(val) * 100).toFixed(1) + "%";
                        if (!isNaN(val) && val !== "") return Math.round(val);
                        return val;
                    });
                    filteredTable.push(rowFormatted);
                }
            }
            doc.autoTable({ head: [filteredTable[0]], body: filteredTable.slice(1), startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [52, 58, 64] } });

            let imgDrill = await getChartImage(() => window.renderDrillDownChart(execName, 'exportCanvas', true));
            let imgPropsDrill = doc.getImageProperties(imgDrill); let pdfHeightDrill = (imgPropsDrill.height * pdfWidth) / imgPropsDrill.width;
            doc.addImage(imgDrill, 'PNG', margin, doc.lastAutoTable.finalY + 20, pdfWidth, pdfHeightDrill);
        }

        document.body.removeChild(exportDiv);

        doc.addPage(); doc.text("Resumen Tabla (Por Gestor)", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaLimpiaGestor', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [52, 58, 64] }, ...autoTableCleaner });

        doc.addPage(); doc.text("Resumen Tabla (Por Zona)", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaLimpiaZona', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [52, 58, 64] }, ...autoTableCleaner });

        doc.save("Dashboard_CRECCU.pdf");

    } catch(e) {
        console.error(e); alert("Hubo un error al generar el PDF.");
    } finally {
        cacheSpans.forEach(item => item.parent.appendChild(item.element));
        if(loader) { loader.style.display = 'none'; loader.querySelector('h3').textContent = "Procesando datos..."; }
        
        if (window.currentZonasChart === 'main') window.renderMainZonasChart();
        else if (window.currentZonasChart === 'comparative') window.renderFaltanComparisonChart('graficoZonas');
        
        if (window.currentExecChart === 'main') window.toggleDrillDown('main');
        else window.toggleDrillDown(window.currentExecChart);
    }
};