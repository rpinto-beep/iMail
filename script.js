let currentScriptUrl = "";
let emailsList = [];
let attachmentsList = []; // Array para múltiples archivos
let footerImageData = null;

// Variables para el control de quota local
let currentVisualQuota = 0;
// Variable para almacenar la info de los alias (Email y Nombre)
let accountData = [];

window.onload = function() { loadSavedConfig(); };

// --- UTILIDADES ---
function setLed(id, isActive) {
    const led = document.getElementById(id);
    if(isActive) led.classList.add('active');
    else led.classList.remove('active');
}

function showMessage(title, text, icon = "ℹ️") {
    document.getElementById('msg-title').innerText = title;
    document.getElementById('msg-text').innerText = text;
    document.querySelector('#messageModal .modal-content div').innerText = icon;
    document.getElementById('messageModal').style.display = 'flex';
}

function closeMessage() {
    document.getElementById('messageModal').style.display = 'none';
}

function showLoader(text, pct) {
    document.getElementById('macLoader').style.display = 'flex';
    document.getElementById('loader-text').innerText = text;
    document.getElementById('progress-fill').style.width = (pct < 10 ? 10 : pct) + "%";
}

function hideLoader() { 
    document.getElementById('macLoader').style.display = 'none'; 
    document.getElementById('progress-fill').style.width = "0%"; 
}

// --- POPUP IMAGEN ZOOM ---
function expandImage(src) {
    document.getElementById('imgZoomTarget').src = src;
    document.getElementById('imgZoomModal').style.display = 'flex';
}
function closeImgModal() {
    document.getElementById('imgZoomModal').style.display = 'none';
}

// --- CONFIGURACIÓN ---
function saveConfig() {
    const urlInput = document.getElementById('configUrlInput').value.trim();
    if(!urlInput.includes("script.google.com")) { 
        showMessage("Error", "URL inválida. Debe ser script.google.com", "⚠️"); 
        return; 
    }
    localStorage.setItem("imac_script_url", urlInput);
    currentScriptUrl = urlInput;
    checkConnection(true);
}

function loadSavedConfig() {
    const savedUrl = localStorage.getItem("imac_script_url");
    if(savedUrl) {
        currentScriptUrl = savedUrl;
        document.getElementById('configUrlInput').value = savedUrl;
        checkConnection(false);
    } else {
        switchTab('config');
    }
}

// --- WIZARD CONFIGURACIÓN SERVIDOR ---
function nextStep(stepId) {
    // Ocultar todos los pasos del wizard
    const steps = document.querySelectorAll('.wizard-step');
    steps.forEach(s => s.style.display = 'none');
    
    // Mostrar el paso solicitado
    document.getElementById('step-' + stepId).style.display = 'block';
}

// --- WIZARD AGREGAR CUENTAS (NUEVO) ---
function nextAccountStep(stepId) {
    // Ocultar pasos del wizard de cuentas
    const steps = document.querySelectorAll('.account-wizard-step');
    steps.forEach(s => s.style.display = 'none');
    
    // Mostrar el paso solicitado
    document.getElementById('account-step-' + stepId).style.display = 'block';
}


// --- CONEXIÓN ---
async function checkConnection(showUi) {
    const quotaDisplay = document.getElementById('quota-count');
    const userDisplay = document.getElementById('connected-user');
    const aliasSelect = document.getElementById('fromEmail'); // Selector de alias
    
    if(showUi) showLoader("Conectando...", 100);
    else document.getElementById('status-bar').innerText = "Conectando...";

    try {
        const response = await fetch(currentScriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: "check_quota" })
        });
        const result = await response.json();
        
        if (result.status === "success") {
            // Guardamos la cuota real inicial
            currentVisualQuota = parseInt(result.quota);
            quotaDisplay.innerText = currentVisualQuota;
            
            userDisplay.innerText = result.user ? result.user : "Desconocido";
            document.getElementById('status-bar').innerText = "Conectado como: " + userDisplay.innerText;

            // --- LLENAR SELECTOR DE ALIAS ---
            aliasSelect.innerHTML = ""; 
            accountData = []; // Reiniciamos datos guardados

            if(result.aliases && result.aliases.length > 0) {
                // Guardamos la lista completa (email y nombre)
                accountData = result.aliases;
                
                result.aliases.forEach(account => {
                    const opt = document.createElement('option');
                    opt.value = account.email;
                    // Mostramos "Nombre <email>" si el nombre existe, sino solo email
                    const displayName = account.name ? `${account.name} <${account.email}>` : account.email;
                    opt.innerText = displayName;
                    aliasSelect.appendChild(opt);
                });
                
                // Ejecutamos la actualización del nombre para el primer elemento seleccionado por defecto
                updateSenderName();

            } else {
                // Si no hay alias o no se actualizó el script, poner solo el principal
                const opt = document.createElement('option');
                opt.value = result.user;
                opt.innerText = result.user;
                aliasSelect.appendChild(opt);
            }
            // --------------------------------
            
            if(showUi) { 
                hideLoader(); 
                showMessage("Conectado", "Sistema listo.\nUsuario: " + userDisplay.innerText, "✅");
                switchTab('main'); 
            }
        }
    } catch (e) {
        userDisplay.innerText = "Error";
        if(showUi) { 
            hideLoader(); 
            showMessage("Error", "No se pudo conectar con el Script.", "❌"); 
        }
    }
}

// --- FUNCIÓN NUEVA: ACTUALIZAR NOMBRE AL SELECCIONAR CORREO ---
function updateSenderName() {
    const selectedEmail = document.getElementById('fromEmail').value;
    const nameInput = document.getElementById('fromName');
    
    // Buscamos en los datos guardados el objeto que coincida con el email
    const account = accountData.find(acc => acc.email === selectedEmail);
    
    if (account && account.name) {
        nameInput.value = account.name; // Rellenar automáticamente
    } else {
        nameInput.value = ""; // Dejar vacío si no tiene nombre configurado
    }
}

// --- ARCHIVOS ---
document.getElementById('csvFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) { setLed('led-list', false); return; }
    
    document.getElementById('file-name-display').innerText = file.name;
    const reader = new FileReader();
    reader.onload = function(evt) {
        emailsList = evt.target.result.split('\n').map(e => e.trim()).filter(e => e !== "");
        document.getElementById('status-bar').innerText = `Cargados: ${emailsList.length} correos.`;
        setLed('led-list', true);
        if(emailsList.length > 0) showMessage("Lista Cargada", `Se importaron ${emailsList.length} destinatarios.`, "📂");
    };
    reader.readAsText(file);
});

document.getElementById('footerImageInput').addEventListener('change', (e) => processFile(e.target.files, 'footer', 'led-footer'));
document.getElementById('attachment').addEventListener('change', (e) => processFile(e.target.files, 'attach', 'led-attach'));

function processFile(files, type, ledId) {
    if(!files || files.length === 0) { setLed(ledId, false); return; }
    
    if (type === 'footer') {
        const file = files[0];
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = function(evt) {
            footerImageData = evt.target.result;
            setLed(ledId, true);
        };
    } 
    else if (type === 'attach') {
        attachmentsList = []; // Reiniciar lista
        document.getElementById('file-list-container').innerHTML = ''; // Limpiar vista
        
        let loadedCount = 0;
        
        Array.from(files).forEach(file => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = function(evt) {
                const full = evt.target.result;
                attachmentsList.push({ 
                    name: file.name, 
                    data: full.split(',')[1], 
                    mime: full.match(/:(.*?);/)[1] 
                });
                
                // Crear Icono Visual
                renderFileIcon(file.name);
                
                loadedCount++;
                if(loadedCount === files.length) {
                    setLed(ledId, true);
                    document.getElementById('attach-status').innerText = `${files.length} archivo(s)`;
                }
            };
        });
    }
}

function renderFileIcon(fileName) {
    const container = document.getElementById('file-list-container');
    const ext = fileName.split('.').pop().toLowerCase();
    
    let typeClass = 'default';
    let iconChar = '📄';
    
    if(['pdf'].includes(ext)) { typeClass = 'pdf'; iconChar = 'PDF'; }
    else if(['xls', 'xlsx', 'csv'].includes(ext)) { typeClass = 'xls'; iconChar = 'XLS'; }
    else if(['doc', 'docx'].includes(ext)) { typeClass = 'doc'; iconChar = 'DOC'; }
    else if(['jpg', 'jpeg', 'png', 'gif'].includes(ext)) { typeClass = 'img'; iconChar = 'IMG'; }
    
    const chip = document.createElement('div');
    chip.className = `file-chip ${typeClass}`;
    chip.innerHTML = `<span class="file-icon-char">${iconChar}</span> ${fileName}`;
    container.appendChild(chip);
}

// --- UI (ACTUALIZADA PARA 3 TABS) ---
function switchTab(t) {
    // 1. Quitar activo de todos los botones y paneles
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    
    // 2. Activar panel específico
    document.getElementById('tab-' + t).classList.add('active');
    
    // 3. Activar botón correspondiente (Lógica manual segura)
    const btns = document.querySelectorAll('.tab-btn');
    if(t === 'main') btns[0].classList.add('active');
    else if(t === 'config') btns[1].classList.add('active');
    else if(t === 'accounts') btns[2].classList.add('active');
}

function toggleFooterMode() {
    const t = document.querySelector('input[name="footerType"]:checked').value;
    document.getElementById('footer-text-box').style.display = (t === 'text') ? 'block' : 'none';
    document.getElementById('footer-image-box').style.display = (t === 'image') ? 'block' : 'none';
}
function toggleInfoModal() {
    const m = document.getElementById('infoModal');
    m.style.display = (m.style.display === 'none' || m.style.display === '') ? 'flex' : 'none';
}
function copyCode() { 
    document.querySelector('.code-box').select(); 
    document.execCommand('copy'); 
    showMessage("Copiado", "Código copiado al portapapeles.", "📋"); 
}

// --- ENVÍO (LÓGICA ACTUALIZADA MULTI-ARCHIVO) ---
async function startSending() {
    const subject = document.getElementById('subject').value;
    const bodyText = document.getElementById('body').value;
    const footerType = document.querySelector('input[name="footerType"]:checked').value;
    const quotaDisplay = document.getElementById('quota-count');
    const whatsappNumber = document.getElementById('whatsappNumber').value.trim();
    const fromEmail = document.getElementById('fromEmail').value.trim(); // Remitente
    const fromName = document.getElementById('fromName').value.trim();   // Nombre Opcional

    if (!currentScriptUrl) { showMessage("Error", "Configura la URL primero.", "⚙️"); switchTab('config'); return; }
    if (emailsList.length === 0) { showMessage("Error", "Falta la lista de correos.", "📭"); return; }

    // Construcción del HTML
    let finalHtml = `<div style="font-family:Arial;color:#333;">${bodyText.replace(/\n/g, "<br>")}</div>`;
    
    // Insertar Botón de WhatsApp si existe
    if(whatsappNumber) {
        // Limpiar número (dejar solo dígitos)
        const cleanNumber = whatsappNumber.replace(/[^0-9]/g, '');
        finalHtml += `
        <div style="margin: 20px 0;">
            <a href="https://wa.me/${cleanNumber}" target="_blank" style="background-color:#25D366; color:#ffffff; border:1px solid #128C7E; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold; font-family:Helvetica, Arial, sans-serif; display:inline-block;">
                Contactar por WhatsApp
            </a>
        </div>`;
    }

    finalHtml += `<br><hr>`;

    if (footerType === 'text') finalHtml += `<div style="color:#666;font-size:12px;">${document.getElementById('footerText').value.replace(/\n/g, "<br>")}</div>`;
    else if (footerType === 'image' && footerImageData) finalHtml += `<img src="${footerImageData}" style="max-width:100%;">`;

    let sentCount = 0;
    
    if(quotaDisplay.innerText !== "--" && quotaDisplay.innerText !== "Offline") {
        currentVisualQuota = parseInt(quotaDisplay.innerText);
    }

    for (let i = 0; i < emailsList.length; i++) {
        let pct = Math.round(((i + 1) / emailsList.length) * 100);
        showLoader(`Enviando a ${emailsList[i]}`, pct);

        // Payload con array de adjuntos
        const payload = {
            to: emailsList[i], 
            subject: subject, 
            html: finalHtml,
            attachments: attachmentsList,
            fromAlias: fromEmail, // Enviamos el remitente
            fromName: fromName    // Enviamos el nombre forzado
        };

        try {
            const r = await fetch(currentScriptUrl, {
                method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(payload)
            });
            const res = await r.json();
            
            if (res.status === "success") {
                sentCount++;
                currentVisualQuota = currentVisualQuota - 1;
                quotaDisplay.innerText = currentVisualQuota;
                
                if(res.quota <= 0) {
                    throw new Error("Límite diario alcanzado");
                }
            } else if (res.status === "error") {
                console.error("Error envío:", res.message);
                if(i===0 && res.message.includes("From address")) {
                     throw new Error("El remitente no está configurado como alias en Gmail.");
                }
            }
        } catch (e) { 
            console.error(e); 
            if(i===0 && e.message.includes("alias")) {
                hideLoader();
                showMessage("Error de Alias", "El correo remitente no está autorizado en tu Gmail.", "⛔");
                return;
            }
        }
        
        await new Promise(r => setTimeout(r, 1200));
    }
    
    hideLoader();

    // LIMPIEZA DE CAMPOS
    emailsList = [];
    attachmentsList = [];
    footerImageData = null;
    document.getElementById('subject').value = "";
    document.getElementById('body').value = "";
    document.getElementById('whatsappNumber').value = "";
    document.getElementById('footerText').value = "";
    document.getElementById('csvFile').value = "";
    document.getElementById('footerImageInput').value = "";
    document.getElementById('attachment').value = "";
    document.getElementById('file-name-display').innerText = "Ninguno";
    setLed('led-list', false);
    setLed('led-footer', false);
    setLed('led-attach', false);
    document.getElementById('attach-status').style.display = 'none';
    document.getElementById('file-list-container').innerHTML = "";
    
    checkConnection(false); 

    showMessage("¡Finalizado!", `Enviados: ${sentCount} correos exitosamente.`, "🚀");
}