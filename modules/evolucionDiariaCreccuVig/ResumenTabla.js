export function renderResumenTabla(dataGestor, dataZona) {
    const containerLimpia = document.querySelector('#tablaLimpia');
    if (!containerLimpia) return;
    const parentDiv = containerLimpia.parentElement;
    
    parentDiv.innerHTML = `
        <div class="d-flex justify-content-center mb-3">
            <button class="btn btn-primary me-2 fw-bold shadow-sm" id="btnVerGestor">Ver por Gestor</button>
            <button class="btn btn-outline-primary fw-bold shadow-sm" id="btnVerZona">Ver por Zona</button>
        </div>
        <div class="table-responsive" id="divTablaGestor">
            <table class="table table-striped table-hover text-center" id="tablaLimpiaGestor">
                <thead></thead><tbody></tbody>
            </table>
        </div>
        <div class="table-responsive" id="divTablaZona" style="display:none;">
            <table class="table table-striped table-hover text-center" id="tablaLimpiaZona">
                <thead></thead><tbody></tbody>
            </table>
        </div>
    `;

    document.getElementById('btnVerGestor').onclick = function() {
        this.classList.replace('btn-outline-primary', 'btn-primary');
        document.getElementById('btnVerZona').classList.replace('btn-primary', 'btn-outline-primary');
        document.getElementById('divTablaGestor').style.display = 'block';
        document.getElementById('divTablaZona').style.display = 'none';
    };

    document.getElementById('btnVerZona').onclick = function() {
        this.classList.replace('btn-outline-primary', 'btn-primary');
        document.getElementById('btnVerGestor').classList.replace('btn-primary', 'btn-outline-primary');
        document.getElementById('divTablaZona').style.display = 'block';
        document.getElementById('divTablaGestor').style.display = 'none';
    };

    const drawTable = (tableId, dataMat) => {
        const thead = document.querySelector(`#${tableId} thead`);
        const tbody = document.querySelector(`#${tableId} tbody`);
        if(!thead || !tbody) return;
        thead.innerHTML = ""; tbody.innerHTML = "";

        if (dataMat.length === 0) return;

        let trHead = document.createElement("tr");
        for (let j = 0; j < dataMat[0].length; j++) {
            let th = document.createElement("th"); 
            th.textContent = dataMat[0][j]; 
            trHead.appendChild(th);
        }
        thead.appendChild(trHead);

        for (let i = 1; i < dataMat.length; i++) {
            let tr = document.createElement("tr");
            let isTotal = String(dataMat[i][0]).toUpperCase().includes("TOTAL") || String(dataMat[i][1]).toUpperCase().includes("TOTAL");
            if (isTotal) {
                tr.classList.add("fw-bold", "table-warning");
            }
            for (let j = 0; j < dataMat[0].length; j++) {
                let td = document.createElement("td");
                let valor = dataMat[i][j];
                if (!isNaN(valor) && valor !== "") {
                    if (String(dataMat[0][j]).includes("%")) {
                        valor = (parseFloat(valor) * 100).toFixed(1) + "%"; 
                    } else {
                        valor = Math.round(parseFloat(valor));
                    }
                }
                td.innerHTML = valor;
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
    };

    drawTable('tablaLimpiaGestor', dataGestor);
    drawTable('tablaLimpiaZona', dataZona);
}