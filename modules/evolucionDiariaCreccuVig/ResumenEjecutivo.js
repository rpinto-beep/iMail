export function procesarResumenEjecutivos(ejecutivosTotalesDerecha, total_general_meta) {
    let tablaEjecutivosFinal = [];
    let tGanE = 0, tAsigE = 0, tMetaGanE = 0, tFaltanRealExec = 0;

    for (let exec in ejecutivosTotalesDerecha) {
        let e = ejecutivosTotalesDerecha[exec];
        
        let faltan = Math.round(e.faltan); 
        let ganados = Math.round(e.ganados);
        let asignados = Math.round(e.asignados);
        let metaGanados = Math.round(e.metaRut); 
        let metaPct = e.metaPct; // Extraído exactamente como exigido
        let avance = e.avancePct; 
        
        let pctFaltante = Math.max(0, metaPct - avance);

        tablaEjecutivosFinal.push([exec, faltan, ganados, metaGanados, asignados, metaPct, avance, pctFaltante]);
        tGanE += ganados; tAsigE += asignados; tMetaGanE += metaGanados; tFaltanRealExec += faltan; 
    }

    if (tablaEjecutivosFinal.length > 0) {
        // Se le da prioridad absoluta al total general directamente sacado del Excel para MetaPct
        let tMetaPctE = total_general_meta > 0 ? total_general_meta : (tAsigE > 0 ? (tMetaGanE / tAsigE) : 0);
        let tAvanceE = tAsigE > 0 ? (tGanE / tAsigE) : 0;
        let tPctFaltanteE = Math.max(0, tMetaPctE - tAvanceE);

        tablaEjecutivosFinal.push(["Cumplimento totalizado", tFaltanRealExec, tGanE, tMetaGanE, tAsigE, tMetaPctE, tAvanceE, tPctFaltanteE]);
        window.mainExecChartData = tablaEjecutivosFinal.slice(); 
        
        tablaEjecutivosFinal.unshift(["Ejecutivo", "Faltan", "# Ganados", "# Meta de ganados", "Asignado", "% Meta del mes", "% Avance del mes", "% Faltante"]);
        
        window.tablaEjecutivosPrincipal = tablaEjecutivosFinal;
        window.renderTable('tablaEjecutivos', tablaEjecutivosFinal, 'execMain');
        window.renderMainExecChart(); 
    }
}