import { cleanAndFormatText } from './utils.js';

/**
 * Mappa una tabella HTML in una griglia 2D gestendo rowspan e colspan.
 * @param {HTMLTableElement} table 
 * @returns {Array<Array<Object>>}
 */
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

                    grid[targetRow][targetCol] = {
                        element: cell,
                        isStart: (r === 0 && c === 0)
                    };
                }
            }
            colIndex += colspan;
        });
    });
    return grid;
}

/**
 * Elabora il file Word caricato.
 * @param {ArrayBuffer} arrayBuffer 
 * @param {string} fileName 
 * @returns {Promise<Object>} Restituisce statistiche e blob del file generato.
 */
export async function processDocument(arrayBuffer, fileName) {
    // Configurazioni Mammoth
    const options = {
        ignoreEmptyParagraphs: false,
        styleMap: ["strike => del"]
    };

    try {
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options);
        const parser = new DOMParser();
        const docDOM = parser.parseFromString(result.value, 'text/html');
        const tables = docDOM.querySelectorAll('table');

        let docParagraphs = [];
        let fullTextForStats = "";

        tables.forEach((table) => {
            const grid = mapTableToGrid(table);
            if (grid.length === 0) return;

            // 1. TROVA L'INTESTAZIONE "AUDIO"
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

            // 2. ESTRAI CONTENUTO
            if (audioColIdx !== -1) {
                let tableTextRaw = "";
                const processedCells = new Set();

                for (let r = headerRowIdx + 1; r < grid.length; r++) {
                    const row = grid[r];
                    if (row && row[audioColIdx]) {
                        const cellObj = row[audioColIdx];

                        if (cellObj.isStart && !processedCells.has(cellObj.element)) {
                            processedCells.add(cellObj.element);

                            let html = cellObj.element.innerHTML.replace(/<del>.*?<\/del>/g, '');
                            html = html.replace(/<img[^>]*>/g, '');

                            const tmp = document.createElement('div');
                            tmp.innerHTML = html;
                            tableTextRaw += tmp.textContent + " ";
                        }
                    }
                }

                const cleanedTableText = cleanAndFormatText(tableTextRaw);

                if (cleanedTableText.length > 0) {
                    fullTextForStats += cleanedTableText + " ";
                }
            }
        });

        const totalCharCount = fullTextForStats.trim().length;
        const cartelleCount = (totalCharCount / 1500).toFixed(2);

        return {
            success: true,
            charCount: totalCharCount,
            cartelleCount: cartelleCount
        };

    } catch (error) {
        console.error("Errore elaborazione:", error);
        throw error;
    }
}
