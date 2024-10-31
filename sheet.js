let data = [];
let filteredData = [];

// Load and display the Excel sheet
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Display the sheet with highlight functionality
function displaySheet(sheetData, highlightRows = [], highlightCols = []) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header;
        if (highlightCols.includes(index)) {
            th.classList.add('highlighted');
        }
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        Object.values(row).forEach((cell, colIndex) => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell;
            if (highlightRows.includes(rowIndex) || highlightCols.includes(colIndex)) {
                td.classList.add('highlighted');
            }
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Update highlighted rows and columns on range input
function updateHighlighting() {
    const rowRangeInput = document.getElementById('row-range').value.trim();
    const colRangeInput = document.getElementById('col-range').value.trim();
    const rowRange = parseRange(rowRangeInput, 'row');
    const colRange = parseRange(colRangeInput, 'col');

    displaySheet(filteredData, rowRange, colRange);
}

// Parse the row and column range inputs
function parseRange(rangeInput, type) {
    if (!rangeInput) return [];

    const rangeParts = rangeInput.split('-');
    const start = parseInt(rangeParts[0], 10);
    const end = rangeParts[1] ? parseInt(rangeParts[1], 10) : start;

    return type === 'row'
        ? Array.from({ length: end - start + 1 }, (_, i) => start + i - 1)
        : Array.from({ length: end.charCodeAt(0) - start.charCodeAt(0) + 1 }, (_, i) => start.charCodeAt(0) - 65 + i);
}

// Apply final operations on data
function applyOperation() {
    // Your existing applyOperation code
}

// Event Listeners
document.getElementById('row-range').addEventListener('input', updateHighlighting);
document.getElementById('col-range').addEventListener('input', updateHighlighting);
document.getElementById('apply-operation').addEventListener('click', applyOperation);
