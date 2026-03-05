// ============================================================
// MODULO: Creccu 1EDC (DASHBOARD COMPLETO Y EXPORTACIONES)
// ============================================================

window.drawOverlappingChart1EDC = function(canvasId, labels, dataObj, labelTitle, totalFaltan, isExport = false) {
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
        id: 'mainChartLabelsPlugin1EDC_' + canvasId,
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
                        const val = Math.round(dataset.data[index]); 
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
        id: 'totalBadgePlugin1EDC_' + canvasId,
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
        id: 'miniBadgeTicksPlugin1EDC_' + canvasId,
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

export function renderTable1EDC(tableId, data) {
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
        let isTotalRow = String(data[i][0]).toUpperCase().includes("TOTAL") || String(data[i][0]).includes("Cumplimiento");
        if (isTotalRow) tr.classList.add("fw-bold", "table-warning");

        for (let j = 0; j < data[0].length; j++) {
            let td = document.createElement("td");
            let valor = data[i][j];

            if (!isNaN(valor) && valor !== "") {
                let num = parseFloat(valor);
                if (String(data[0][j]).includes("%")) {
                    valor = (num * 100).toFixed(1) + "%";
                } else {
                    valor = Math.round(num); // Cero decimales
                }
            }
            td.innerHTML = valor;
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
}

// ==========================================
// FUNCIONES DE DESCARGA GLOBALES SEGURAS
// ==========================================

window.descargarExcel1EDC = function() {
    if (typeof XLSX === "undefined") { 
        if (window.mostrarAlerta) window.mostrarAlerta("Librería XLSX no cargada. Revisa tu conexión.", "danger"); 
        return; 
    }
    try {
        let wb = XLSX.utils.book_new();
        let ws1 = XLSX.utils.table_to_sheet(document.getElementById('tablaZonas1EDC'));
        let ws2 = XLSX.utils.table_to_sheet(document.getElementById('tablaContact1EDC'));
        XLSX.utils.book_append_sheet(wb, ws1, "Zonas 1EDC");
        XLSX.utils.book_append_sheet(wb, ws2, "Contactabilidad");
        XLSX.writeFile(wb, "Dashboard_1EDC.xlsx");
    } catch(e) {
        console.error("Error al generar Excel 1EDC:", e);
        if (window.mostrarAlerta) window.mostrarAlerta("Hubo un error al generar el Excel.", "danger");
    }
};

window.descargarPDF1EDC = async function() {
    if (!window.jspdf) { 
        if (window.mostrarAlerta) window.mostrarAlerta("Librería jsPDF no cargada. Revisa tu conexión.", "danger"); 
        return; 
    }
    const loader = document.getElementById('loader');
    if (loader) { loader.querySelector('h3').textContent = "Generando PDF por favor espere..."; loader.style.display = 'flex'; }

    // Limpieza preventiva de canvas residuales
    let oldDiv = document.getElementById('exportDiv1EDC');
    if (oldDiv) oldDiv.remove();

    let exportDiv = document.createElement('div'); 
    exportDiv.id = 'exportDiv1EDC';
    exportDiv.style.width = '1000px'; 
    exportDiv.style.height = '500px'; 
    exportDiv.style.position = 'fixed'; 
    exportDiv.style.top = '-9999px';
    
    let exportCanvas = document.createElement('canvas'); 
    exportCanvas.id = 'exportCanvas1EDC'; 
    exportCanvas.width = 1000; 
    exportCanvas.height = 500; 
    exportDiv.appendChild(exportCanvas); 
    document.body.appendChild(exportDiv);

    try {
        const { jsPDF } = window.jspdf; 
        const doc = new jsPDF('p', 'pt', 'a4');
        const pageWidth = doc.internal.pageSize.getWidth(); 
        const margin = 40; 
        const pdfWidth = pageWidth - (margin * 2);

        doc.setFontSize(16); 
        doc.text("Resumen por Zonas 1EDC", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaZonas1EDC', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [13, 110, 253] } });

        const getChartImage = (renderFunc) => { 
            return new Promise(resolve => { 
                renderFunc('exportCanvas1EDC', true); 
                setTimeout(() => resolve(exportCanvas.toDataURL('image/png', 1.0)), 300); 
            }); 
        };

        let imgZonas = await getChartImage(window.render1EDCZonasChart);
        let imgProps = doc.getImageProperties(imgZonas); 
        let pdfHeight = (imgProps.height * pdfWidth) / imgProps.width; 
        let finalY = doc.lastAutoTable ? doc.lastAutoTable.finalY + 20 : 150;
        doc.addImage(imgZonas, 'PNG', margin, finalY, pdfWidth, pdfHeight);

        doc.addPage(); 
        doc.text("Contactabilidad 1EDC", pageWidth/2, 40, {align: 'center'});
        doc.autoTable({ html: '#tablaContact1EDC', startY: 60, styles: { fontSize: 8, halign: 'center' }, headStyles: { fillColor: [25, 135, 84] } });
        
        let currentY2 = doc.lastAutoTable ? doc.lastAutoTable.finalY + 20 : 150;

        // ==========================================
        // DIBUJAR TARJETAS EN EL PDF ANTES DE LA DONA
        // ==========================================
        let cData = window.main1EDCContactDataObj;
        if (cData && cData.labels && cData.labels.length > 0) {
            let cardWidth = (pdfWidth - 30) / 4; // 4 tarjetas con 10pt de espacio entre ellas
            let cardHeight = 45;
            
            for(let i = 0; i < cData.labels.length; i++) {
                let xPos = margin + i * (cardWidth + 10);
                
                // Color de fondo y borde Manga
                doc.setFillColor(cData.bg[i]);
                doc.setDrawColor('#000000');
                doc.setLineWidth(1.5);
                doc.roundedRect(xPos, currentY2, cardWidth, cardHeight, 4, 4, 'FD');
                
                // Contraste de texto igual que en la web
                if (cData.labels[i] === 'CONTACTO TERCERO') {
                    doc.setTextColor('#000000');
                } else {
                    doc.setTextColor('#ffffff');
                }
                
                // Titulo Tarjeta
                doc.setFontSize(7);
                doc.setFont("helvetica", "bold");
                doc.text(cData.labels[i], xPos + cardWidth/2, currentY2 + 15, {align: 'center'});
                
                // Valor Tarjeta (% en vez de número absoluto)
                doc.setFontSize(16);
                doc.text(cData.data[i].toFixed(1) + '%', xPos + cardWidth/2, currentY2 + 35, {align: 'center'});
            }
            currentY2 += cardHeight + 20; // Bajar Y para dar espacio a la gráfica
        }
        
        let imgContact = await getChartImage(window.render1EDCContactChart);
        let imgProps2 = doc.getImageProperties(imgContact); 
        let pdfHeight2 = (imgProps2.height * pdfWidth) / imgProps2.width;
        
        // Ajuste de seguridad por si la gráfica es muy grande y sobrepasa la página
        if (currentY2 + pdfHeight2 > 800) {
            pdfHeight2 = 800 - currentY2;
            let nW = (imgProps2.width * pdfHeight2) / imgProps2.height;
            let nX = margin + (pdfWidth - nW) / 2;
            doc.addImage(imgContact, 'PNG', nX, currentY2, nW, pdfHeight2);
        } else {
            doc.addImage(imgContact, 'PNG', margin, currentY2, pdfWidth, pdfHeight2);
        }

        doc.save("Dashboard_1EDC.pdf");

    } catch(e) {
        console.error("Error crítico al generar PDF 1EDC:", e); 
        if (window.mostrarAlerta) window.mostrarAlerta("Hubo un error al generar el PDF. Revisa la consola.", "danger");
    } finally {
        document.body.removeChild(exportDiv);
        if(loader) { loader.style.display = 'none'; loader.querySelector('h3').textContent = "Procesando datos..."; }
        if (window.render1EDCZonasChart) window.render1EDCZonasChart();
        if (window.render1EDCContactChart) window.render1EDCContactChart();
    }
};

// ==========================================
// FUNCIÓN PRINCIPAL DE PROCESAMIENTO
// ==========================================
export function procesarEvolucion1EDC(data) {
    let t1 = { row: -1, asig: -1, gan: -1, av: -1, zona: -1, meta: -1, falta: -1 };
    let t2 = { row: -1, label: -1, contact: -1, asig: -1 };

    for (let i = 0; i < data.length; i++) {
        if (!data[i] || data[i].length === 0) continue;
        let r = data[i].map(c => String(c).toUpperCase().trim());

        if (t1.row === -1 && r.includes("DETALLE POR GESTOR") && r.includes("CUENTA DE RUT") && r.includes("METAS")) {
            t1.row = i + 1;
            t1.zona = r.findIndex(x => x === "DETALLE POR GESTOR");
            t1.asig = r.findIndex(x => x === "CUENTA DE RUT");
            t1.gan = r.findIndex(x => x === "GANADOS");
            t1.av = r.findIndex(x => x === "% AVANCE");
            if (t1.av === -1 && r.some(x => x.includes("AVANCE"))) t1.av = r.findIndex(x => x.includes("AVANCE"));
            t1.meta = r.findIndex(x => x === "METAS");
            t1.falta = r.findIndex(x => x === "FALTA");
        }

        if (t2.row === -1 && r.includes("ETIQUETAS DE FILA") && r.includes("CONTACT.")) {
            t2.row = i + 1;
            t2.label = r.indexOf("ETIQUETAS DE FILA");
            t2.contact = r.indexOf("CONTACT.");
            t2.asig = r.indexOf("CUENTA DE RUT");
        }
    }

    if (t1.row === -1 || t2.row === -1) {
        if (window.mostrarAlerta) window.mostrarAlerta("Formato no reconocido. Faltan tablas de Zonas o Contactabilidad 1EDC.", "danger");
        return;
    }

    // ==========================================
    // EXTRACCIÓN TABLA 1: ZONAS 1EDC
    // ==========================================
    let zonasData = [];
    let totZonas = { asig: 0, gan: 0, metaGan: 0, falta: 0 };
    let totZonasExcel = { asig: 0, gan: 0, metaPct: 0, avPct: 0 };

    for (let i = t1.row; i < data.length; i++) {
        let r = data[i];
        if (!r) continue;
        let zRaw = String(r[t1.zona] || "").trim();
        
        if (!zRaw || zRaw.toLowerCase() === "nan") {
            let asigTest = parseFloat(r[t1.asig]);
            if (!isNaN(asigTest) && asigTest > 0) {
                totZonasExcel.metaPct = parseFloat(r[t1.meta]) || 0;
                totZonasExcel.avPct = parseFloat(r[t1.av]) || 0;
                break;
            }
            continue;
        }

        if (zRaw.toUpperCase().includes("TOTAL")) {
            totZonasExcel.metaPct = parseFloat(r[t1.meta]) || 0;
            totZonasExcel.avPct = parseFloat(r[t1.av]) || 0;
            break;
        }

        let zName = zRaw.replace(/^[0-9\s.:-]+/, "").trim();
        let asig = Math.round(parseFloat(r[t1.asig]) || 0);
        let gan = Math.round(parseFloat(r[t1.gan]) || 0);
        let av = parseFloat(r[t1.av]) || 0;
        let meta = parseFloat(r[t1.meta]) || 0;
        let falta = Math.round(parseFloat(r[t1.falta]) || 0);
        
        let metaGan = falta + gan;
        let pctFaltante = Math.max(0, meta - av);

        if (asig > 0 || gan > 0) {
            zonasData.push({ zona: zName, asig, gan, av, meta, falta, metaGan, pctFaltante });
            totZonas.asig += asig;
            totZonas.gan += gan;
            totZonas.falta += falta;
            totZonas.metaGan += metaGan;
        }
    }

    let tablaZonas1EDC = [["Zona", "Faltan", "# Ganados", "# Meta de ganados", "Asignado", "% Meta del mes", "% Avance del mes", "% Faltante"]];
    
    window.main1EDCChartDataObj = { labels: [], asig: [], metas: [], gan: [], falta: [], av: [], metaPct: [], faltanText: [], tFaltan: totZonas.falta };

    for (let z of zonasData) {
        tablaZonas1EDC.push([z.zona, z.falta, z.gan, z.metaGan, z.asig, z.meta, z.av, z.pctFaltante]);
        window.main1EDCChartDataObj.labels.push([z.zona, "", "", " "]);
        window.main1EDCChartDataObj.asig.push(z.asig);
        window.main1EDCChartDataObj.gan.push(z.gan);
        window.main1EDCChartDataObj.metas.push(z.metaGan);
        window.main1EDCChartDataObj.falta.push(z.falta);
        window.main1EDCChartDataObj.av.push(z.av);
        window.main1EDCChartDataObj.metaPct.push(z.meta);
        
        let diffText = z.pctFaltante > 0.0001 ? `Falta: ${(z.pctFaltante * 100).toFixed(1)}%` : 'Meta Cumplida';
        window.main1EDCChartDataObj.faltanText.push(diffText);
    }

    let finalMetaPct = totZonasExcel.metaPct > 0 ? totZonasExcel.metaPct : (totZonas.asig > 0 ? totZonas.metaGan / totZonas.asig : 0);
    let finalAvance = totZonasExcel.avPct > 0 ? totZonasExcel.avPct : (totZonas.asig > 0 ? totZonas.gan / totZonas.asig : 0);
    let finalFaltante = Math.max(0, finalMetaPct - finalAvance);
    
    tablaZonas1EDC.push(["Cumplimiento totalizado", totZonas.falta, totZonas.gan, totZonas.metaGan, totZonas.asig, finalMetaPct, finalAvance, finalFaltante]);
    
    renderTable1EDC('tablaZonas1EDC', tablaZonas1EDC);
    
    window.render1EDCZonasChart = function(targetCanvasId = 'graficoZonas1EDC', isExport = false) {
        let d = window.main1EDCChartDataObj;
        window.drawOverlappingChart1EDC(targetCanvasId, d.labels, {
            asignados: d.asig, metas: d.metas, ganados: d.gan, faltan: d.falta, avance: d.av, metaPct: d.metaPct, faltanText: d.faltanText
        }, 'Resumen por Zonas 1EDC (Avance vs Meta)', d.tFaltan, isExport);
    };
    
    try { window.render1EDCZonasChart(); } catch(e) { console.error("Error al renderizar Zonas 1EDC:", e); }

    // ==========================================
    // EXTRACCIÓN TABLA 2: CONTACTABILIDAD
    // ==========================================
    let contactData = [];
    let totContact = { asig: 0, pct: 0 };

    for (let i = t2.row; i < data.length; i++) {
        let r = data[i];
        if (!r) continue;
        let lRaw = String(r[t2.label] || "").trim();
        if (!lRaw || lRaw.toLowerCase() === "nan") continue;
        
        if (lRaw.toUpperCase().includes("TOTAL")) {
            totContact.asig = parseFloat(r[t2.asig]) || 0;
            totContact.pct = parseFloat(r[t2.contact]) || 0;
            break;
        }

        let lName = lRaw.replace(/^[0-9\s.:-]+/, "").trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
        let pct = parseFloat(r[t2.contact]) || 0;
        let asig = Math.round(parseFloat(r[t2.asig]) || 0);

        contactData.push({ label: lName, asig, pct });
    }

    const contactColors = {
        'SIN GESTION': '#6c757d', // Plomo
        'SIN CONTACTO': '#dc3545', // Rojo
        'CONTACTO DIRECTO': '#198754', // Verde
        'CONTACTO TERCERO': '#0dcaf0' // Celeste
    };

    let pLabels = [];
    let pData = [];
    let pBg = [];
    let rawAsigData = []; 
    let tablaContact1EDC = [["Etiqueta", "Cantidad", "% Contactabilidad"]];
    let cardsHtml = "";

    for (let c of contactData) {
        pLabels.push(c.label);
        pData.push(c.pct * 100);
        rawAsigData.push(c.asig);
        let bgColor = contactColors[c.label] || '#6c757d';
        pBg.push(bgColor);
        
        tablaContact1EDC.push([c.label, c.asig, c.pct]);

        let textColor = (c.label === 'CONTACTO TERCERO') ? 'text-dark' : 'text-white';
        let textShadow = (c.label === 'CONTACTO TERCERO') 
            ? 'text-shadow: 1px 1px 0px #fff, -1px -1px 0px #fff, 1px -1px 0px #fff, -1px 1px 0px #fff;' 
            : 'text-shadow: 1px 1px 0px #000, -1px -1px 0px #000, 1px -1px 0px #000, -1px 1px 0px #000;';

        cardsHtml += `
            <div class="col-md-3 px-2">
                <div class="card ${textColor} shadow mb-2" style="background-color: ${bgColor}; border: 3px solid #000; border-radius: 12px;">
                    <div class="card-body py-3">
                        <h6 class="card-title fw-bold text-uppercase mb-1" style="font-size:0.8rem; letter-spacing: 0.5px;">${c.label}</h6>
                        <h2 class="mb-0 fw-bold" style="font-size: 2rem; ${textShadow}">${(c.pct * 100).toFixed(1)}%</h2>
                    </div>
                </div>
            </div>
        `;
    }
    tablaContact1EDC.push(["TOTAL GENERAL", totContact.asig, totContact.pct]);

    document.getElementById('cardsContact1EDC').innerHTML = cardsHtml;
    window.main1EDCContactCardsHtml = cardsHtml;
    
    renderTable1EDC('tablaContact1EDC', tablaContact1EDC);

    window.main1EDCContactDataObj = { labels: pLabels, data: pData, bg: pBg, rawAsig: rawAsigData };

    // Plugin inyectado para mostrar Cuenta de RUT en la Dona (Estilo Manga)
    const pieLabelsPlugin1EDC = {
        id: 'pieLabelsPlugin1EDC',
        afterDatasetsDraw(chart) {
            try {
                const { ctx, data } = chart;
                ctx.save();
                ctx.font = 'bold 22px Arial';
                ctx.fillStyle = '#fff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                
                const meta = chart.getDatasetMeta(0);
                if (!meta || !meta.data) return;
                
                meta.data.forEach((arc, index) => {
                    const val = window.main1EDCContactDataObj.rawAsig[index];
                    if (val > 0 && typeof arc.tooltipPosition === 'function') {
                        const centerPoint = arc.tooltipPosition();
                        ctx.strokeStyle = '#000';
                        ctx.lineWidth = 4;
                        ctx.strokeText(val, centerPoint.x, centerPoint.y);
                        ctx.fillText(val, centerPoint.x, centerPoint.y);
                    }
                });
                ctx.restore();
            } catch (err) {
                console.warn("No se pudieron renderizar los textos internos de la dona:", err);
            }
        }
    };

    window.render1EDCContactChart = function(targetCanvasId = 'graficoContact1EDC', isExport = false) {
        const ctx = document.getElementById(targetCanvasId).getContext('2d');
        let existingChart = Chart.getChart(targetCanvasId);
        if (existingChart) existingChart.destroy();
        
        let d = window.main1EDCContactDataObj;
        
        new Chart(ctx, {
            type: 'doughnut', 
            data: { 
                labels: d.labels, 
                datasets: [{ 
                    data: d.data, 
                    backgroundColor: d.bg, 
                    borderColor: '#000000', 
                    borderWidth: 3, 
                    hoverOffset: 10,
                    spacing: 5 
                }] 
            },
            options: {
                animation: isExport ? false : true,
                responsive: true,
                maintainAspectRatio: false,
                cutout: '45%', 
                plugins: {
                    title: { display: true, text: 'Distribución de Contactabilidad', font: { size: 16 }, color: '#000' },
                    legend: { position: 'right', labels: { color: '#000', font: { size: 13, weight: 'bold' } } },
                    tooltip: { callbacks: { label: function(context) { return ` ${context.label}: ${context.parsed.toFixed(1)}%`; } } }
                }
            },
            plugins: [pieLabelsPlugin1EDC]
        });
    };
    
    try { window.render1EDCContactChart(); } catch(e) { console.error("Error al renderizar Dona 1EDC:", e); }

    let btnPdf = document.getElementById('btnPdf1EDC');
    let btnExcel = document.getElementById('btnExcel1EDC');
    if (btnPdf) btnPdf.disabled = false;
    if (btnExcel) btnExcel.disabled = false;
}