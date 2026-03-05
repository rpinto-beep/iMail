import { procesarEvolucionDiaria } from './modules/evolucionDiariaCreccuVig/index.js';
import { procesarEvolucion1EDC } from './modules/Creccu1EDC/index.js';
import { iniciarPresentacion } from './modules/PPT/index.js';

window.iniciarPresentacion = iniciarPresentacion;

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileInput1EDC = document.getElementById('fileInput1EDC');
    const loader = document.getElementById('loader');

    // Lector Creccu Vig
    if (fileInput) {
        fileInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            loader.style.display = 'flex';

            setTimeout(() => {
                try {
                    const fileName = file.name.toLowerCase();

                    if (fileName.endsWith('.csv')) {
                        Papa.parse(file, {
                            complete: function(results) {
                                procesarEvolucionDiaria(results.data);
                                loader.style.display = 'none';
                            }
                        });
                    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                        const reader = new FileReader();
                        reader.onload = function(e) {
                            try {
                                const data = new Uint8Array(e.target.result);
                                const workbook = XLSX.read(data, {type: 'array'});
                                
                                let sheetName = workbook.SheetNames.includes("TABLA") ? "TABLA" : workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[sheetName];
                                
                                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
                                
                                procesarEvolucionDiaria(jsonData);
                            } catch (err) {
                                console.error(err);
                                window.mostrarAlerta("El archivo Excel no cumple con el formato requerido en Creccu Vig.", "danger");
                            } finally {
                                loader.style.display = 'none';
                            }
                        };
                        reader.readAsArrayBuffer(file);
                    }
                } catch (error) {
                    console.error(error);
                    window.mostrarAlerta("Error crítico al cargar el archivo de Creccu Vig.", "danger");
                    loader.style.display = 'none';
                }
            }, 100);
        });
    }

    // Lector Creccu 1EDC
    if (fileInput1EDC) {
        fileInput1EDC.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            loader.style.display = 'flex';

            setTimeout(() => {
                try {
                    const fileName = file.name.toLowerCase();

                    if (fileName.endsWith('.csv')) {
                        Papa.parse(file, {
                            complete: function(results) {
                                procesarEvolucion1EDC(results.data);
                                loader.style.display = 'none';
                            }
                        });
                    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                        const reader = new FileReader();
                        reader.onload = function(e) {
                            try {
                                const data = new Uint8Array(e.target.result);
                                const workbook = XLSX.read(data, {type: 'array'});
                                
                                let sheetName = workbook.SheetNames.includes("TABLA") ? "TABLA" : workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[sheetName];
                                
                                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
                                
                                procesarEvolucion1EDC(jsonData);
                            } catch (err) {
                                console.error(err);
                                window.mostrarAlerta("El archivo Excel no cumple con el formato requerido en Creccu 1EDC.", "danger");
                            } finally {
                                loader.style.display = 'none';
                            }
                        };
                        reader.readAsArrayBuffer(file);
                    }
                } catch (error) {
                    console.error(error);
                    window.mostrarAlerta("Error crítico al cargar el archivo de Creccu 1EDC.", "danger");
                    loader.style.display = 'none';
                }
            }, 100);
        });
    }

    const btnGenerarPPT = document.getElementById('btnGenerarPPT');
    if (btnGenerarPPT) {
        btnGenerarPPT.addEventListener('click', iniciarPresentacion);
    }
});