export function procesarResumenZonas(zonasFinalData, t3Total) {
    let tablaZonasFinal = [];
    let tGanZ = 0, tAsigZ = 0, tMetaGanZ = 0, tFaltanZ = 0;

    for (let z of zonasFinalData) {
        let ganadosZ = Math.round(z.ganados);
        let asignadosZ = Math.round(z.asignados);
        let metaGanZ = Math.round(z.metaGanados);
        let faltanZOrig = Math.round(z.faltanOrig); // Conserva los negativos (neto)

        tablaZonasFinal.push([z.nombreMostrar, faltanZOrig, ganadosZ, metaGanZ, asignadosZ, z.metaPct, z.avance, z.pctFaltante]);
        tGanZ += ganadosZ; tAsigZ += asignadosZ; tMetaGanZ += metaGanZ; tFaltanZ += faltanZOrig; 
    }

    if (tablaZonasFinal.length > 0) {
        let tMetaPctZ = t3Total ? t3Total.meta : (tAsigZ > 0 ? (tMetaGanZ / tAsigZ) : 0);
        let tAvanceZ = t3Total ? t3Total.avance : (tAsigZ > 0 ? (tGanZ / tAsigZ) : 0);
        let tPctFaltanteZ = t3Total ? t3Total.pctFaltante : Math.max(0, tMetaPctZ - tAvanceZ); 

        tablaZonasFinal.push(["Cumplimento totalizado", tFaltanZ, tGanZ, tMetaGanZ, tAsigZ, tMetaPctZ, tAvanceZ, tPctFaltanteZ]);
        
        let labelsZ = [], faltanZ = [], ganadosZ = [], metasZ = [], asignadosZ = [], avanceZ = [], metaPctZ = [], faltanTextZ = [];
        
        for (let i = 0; i < tablaZonasFinal.length - 1; i++) {
            let zonaNombre = tablaZonasFinal[i][0];
            let pctFalt = tablaZonasFinal[i][7]; 
            
            let diffText = pctFalt > 0.0001 ? `Falta: ${(pctFalt * 100).toFixed(1)}%` : 'Meta Cumplida';
            
            labelsZ.push([zonaNombre, "", "", " "]);
            faltanTextZ.push(diffText); 
            
            faltanZ.push(tablaZonasFinal[i][1]);
            ganadosZ.push(tablaZonasFinal[i][2]);
            metasZ.push(tablaZonasFinal[i][3]);
            asignadosZ.push(tablaZonasFinal[i][4]);
            metaPctZ.push(tablaZonasFinal[i][5]);
            avanceZ.push(tablaZonasFinal[i][6]);
        }
        
        tablaZonasFinal.unshift(["Zona", "Faltan", "# Ganados", "# Meta de ganados", "Asignado", "% Meta del mes", "% Avance del mes", "% Faltante"]);
        
        window.mainZonasChartDataObj = { labelsZ, asignadosZ, metasZ, ganadosZ, faltanZ, avanceZ, metaPctZ, faltanTextZ, tFaltanZ };
        window.currentZonasChart = 'main';

        window.renderMainZonasChart = function(targetCanvasId = 'graficoZonas', isExport = false) {
            if (!isExport) window.currentZonasChart = 'main';
            let d = window.mainZonasChartDataObj;
            window.drawOverlappingChart(targetCanvasId, d.labelsZ, { asignados: d.asignadosZ, metas: d.metasZ, ganados: d.ganadosZ, faltan: d.faltanZ, avance: d.avanceZ, metaPct: d.metaPctZ, faltanText: d.faltanTextZ }, 'Resumen por Zonas (Avance vs Meta)', d.tFaltanZ, isExport);
        };

        window.renderTable('tablaZonas', tablaZonasFinal, 'zonas');
        window.renderMainZonasChart();
    }
    return tFaltanZ;
}