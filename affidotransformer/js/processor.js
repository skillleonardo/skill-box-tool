import { cleanAndFormatText } from './utils.js';

// Map Table Helper (Reused)
function mapTableToGrid(table) {
    const grid = [];
    const rows = table.querySelectorAll('tr');

    rows.forEach((row, rowIndex) => {
        if (!grid[rowIndex]) grid[rowIndex] = [];
        const cells = row.querySelectorAll('td, th');
        let colIndex = 0;

        cells.forEach(cell => {
            while (grid[rowIndex][colIndex]) {
                colIndex++;
            }
            const rowspan = parseInt(cell.getAttribute('rowspan') || 1);
            const colspan = parseInt(cell.getAttribute('colspan') || 1);
            for (let r = 0; r < rowspan; r++) {
                for (let c = 0; c < colspan; c++) {
                    const targetRow = rowIndex + r;
                    const targetCol = colIndex + c;
                    if (!grid[targetRow]) grid[targetRow] = [];
                    grid[targetRow][targetCol] = { element: cell, isStart: (r === 0 && c === 0) };
                }
            }
            colIndex += colspan;
        });
    });
    return grid;
}

export async function processDocument(arrayBuffer) {
    const options = { ignoreEmptyParagraphs: false };
    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options);
    const parser = new DOMParser();
    const docDOM = parser.parseFromString(result.value, 'text/html');
    const tables = docDOM.querySelectorAll('table');

    const items = [];

    tables.forEach((table) => {
        const grid = mapTableToGrid(table);
        if (grid.length === 0) return;

        // Find "Audio" column
        let audioColIdx = -1;
        let headerRowIdx = -1;

        for (let r = 0; r < grid.length; r++) {
            const row = grid[r];
            for (let c = 0; c < row.length; c++) {
                if (row[c] && row[c].element) {
                    const txt = row[c].element.textContent.toLowerCase().trim();
                    if (txt === 'audio') {
                        audioColIdx = c;
                        headerRowIdx = r;
                        break;
                    }
                }
            }
            if (audioColIdx !== -1) break;
        }

        // If found, extract content of rows below
        if (audioColIdx !== -1) {
            const processedCells = new Set();
            for (let r = headerRowIdx + 1; r < grid.length; r++) {
                const cellObj = grid[r][audioColIdx];
                if (cellObj && cellObj.isStart && !processedCells.has(cellObj.element)) {
                    processedCells.add(cellObj.element);

                    // Extract text
                    let rawText = "";
                    // Basic text extraction preserving some structure? 
                    // Mammoth gives HTML. We just want text for the target doc.
                    // But wait, the Target Doc cell might want formatting? 
                    // Docxtemplater primarily handles Text unless using rawxml module.
                    // We will extract Clean Text.
                    rawText = cellObj.element.textContent || "";
                    rawText = cleanAndFormatText(rawText);

                    if (rawText.trim().length > 0) {
                        items.push({ content: rawText });
                    }
                }
            }
        }
    });

    return items;
}

export async function generateOutput(items, codice, cliente) {
    if (typeof window.PizZip === 'undefined') {
        throw new Error("La libreria PizZip non è stata caricata correttamente. Controlla la connessione internet.");
    }

    // Load Template
    // Note: This fetch assumes the file is accessible relative to index.html
    const response = await fetch('assets/template.docx');
    if (!response.ok) {
        throw new Error("Impossibile caricare il template (template.docx).");
    }
    const content = await response.arrayBuffer();

    const zip = new window.PizZip(content);
    const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Prepare Data
    // We need to map items to filename labels: 1-2, 1-3...
    const dataItems = items.map((item, index) => {
        return {
            filename: `1-${index + 2}`,
            content: item.content
        };
    });

    const dateStr = new Date().toLocaleDateString('it-IT');

    doc.render({
        items: dataItems,
        codice_corso: codice,
        nome_cliente: cliente,
        data_corrente: dateStr
    });

    const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    return out;
}
