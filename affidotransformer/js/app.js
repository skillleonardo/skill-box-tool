import { processDocument, generateOutput } from './processor.js';

const fileInput = document.getElementById('file-input');
const dropZone = document.getElementById('drop-zone');
const processBtn = document.getElementById('process-btn');
const statusMsg = document.getElementById('status-msg');
const fileNameDisplay = document.getElementById('file-name');
const inputCodice = document.getElementById('input-codice');
const inputCliente = document.getElementById('input-cliente');

let selectedFile = null;

fileInput.onchange = (e) => {
    if (e.target.files.length > 0) {
        selectedFile = e.target.files[0];
        fileNameDisplay.textContent = selectedFile.name;
        fileNameDisplay.classList.add('text-orange-600');
        processBtn.disabled = false;
        statusMsg.textContent = "";
    }
};

dropZone.onclick = () => fileInput.click();

dropZone.ondragover = (e) => {
    e.preventDefault();
    dropZone.classList.add('border-orange-400', 'bg-orange-50');
};

dropZone.ondragleave = () => {
    dropZone.classList.remove('border-orange-400', 'bg-orange-50');
};

dropZone.ondrop = (e) => {
    e.preventDefault();
    dropZone.classList.remove('border-orange-400', 'bg-orange-50');
    if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
        fileInput.onchange({ target: fileInput });
    }
};

processBtn.onclick = async () => {
    if (!selectedFile) return;

    const codice = inputCodice.value.trim();
    const cliente = inputCliente.value.trim();

    if (!codice || !cliente) {
        alert("Inserisci Codice Corso e Nome Cliente.");
        return;
    }

    statusMsg.textContent = "Analisi file in corso...";
    processBtn.disabled = true;

    try {
        const reader = new FileReader();
        reader.onload = async (event) => {
            const arrayBuffer = event.target.result;
            try {
                // 1. Extract Data from Source
                const items = await processDocument(arrayBuffer);

                if (items.length === 0) {
                    throw new Error("Nessun contenuto Audio trovato nel file.");
                }

                statusMsg.textContent = `Trovati ${items.length} elementi. Generazione documento...`;

                // 2. Generate Output using Template
                const blob = await generateOutput(items, codice, cliente);

                // 3. Trigger Download
                saveAs(blob, "CODCORSO_audio_compilato.docx");

                statusMsg.textContent = "Completato! Download avviato.";
                statusMsg.className = "mt-4 text-sm font-medium text-green-600";
            } catch (err) {
                console.error(err);
                statusMsg.textContent = "Errore: " + err.message;
                statusMsg.className = "mt-4 text-sm font-medium text-red-600";
            } finally {
                processBtn.disabled = false;
            }
        };
        reader.readAsArrayBuffer(selectedFile);

    } catch (e) {
        console.error(e);
        statusMsg.textContent = "Errore generico.";
        processBtn.disabled = false;
    }
};
