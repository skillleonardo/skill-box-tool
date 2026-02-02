import { processDocument } from './processor.js';

// DOM Elements
const fileInput = document.getElementById('file-input');
const dropZone = document.getElementById('drop-zone');
const processBtn = document.getElementById('process-btn');
const statusMsg = document.getElementById('status-msg');
const statsPanel = document.getElementById('stats-panel');
const charCountDisplay = document.getElementById('char-count-display');
const cartelleDisplay = document.getElementById('cartelle-display');
const fileNameDisplay = document.getElementById('file-name');

let selectedFile = null;

// Event Listeners
fileInput.onchange = (e) => {
    if (e.target.files.length > 0) {
        selectedFile = e.target.files[0];
        fileNameDisplay.textContent = selectedFile.name;
        fileNameDisplay.classList.add('text-orange-600');
        processBtn.disabled = false;
        statsPanel.classList.add('hidden');
        statusMsg.textContent = "";
        statusMsg.className = "mt-4 text-sm font-medium text-gray-500 min-h-[20px]";
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

    statusMsg.textContent = "Elaborazione struttura tabelle in corso...";
    statusMsg.className = "mt-4 text-sm font-medium text-gray-500 min-h-[20px]";
    processBtn.disabled = true;

    const reader = new FileReader();

    reader.onload = async (event) => {
        const arrayBuffer = event.target.result;

        try {
            const result = await processDocument(arrayBuffer, selectedFile.name);

            // Update UI
            charCountDisplay.textContent = result.charCount.toLocaleString('it-IT');
            cartelleDisplay.textContent = result.cartelleCount.replace('.', ',');
            statsPanel.classList.remove('hidden');

            statusMsg.textContent = "Calcolo completato!";
            statusMsg.className = "mt-4 text-sm font-medium text-green-600";
        } catch (error) {
            statusMsg.textContent = "Errore durante l'elaborazione.";
            statusMsg.className = "mt-4 text-sm font-medium text-red-600";
        } finally {
            processBtn.disabled = false;
        }
    };

    reader.readAsArrayBuffer(selectedFile);
};
