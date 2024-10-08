document.getElementById('compareBtn').addEventListener('click', compareExcels);
document.getElementById('addMoreBtn').addEventListener('click', addMoreFile);
document.getElementById('exportBtn').addEventListener('click', exportResults);

let fileIndex = 2; // Keeps track of how many files are added
let resultsData = [];

function compareExcels() {
    const fileInputSections = document.querySelectorAll('.file-input-section');
    const filesData = [];

    fileInputSections.forEach((section, index) => {
        const fileInput = section.querySelector('input[type="file"]');
        const skuColumn = section.querySelector('.sku-column').value.toUpperCase();
        const priceColumn = section.querySelector('.price-column').value.toUpperCase();

        if (!fileInput.files.length) return; // Ignore empty file sections

        if (!skuColumn || !priceColumn) {
            alert(`Please specify SKU and Price columns for Excel file ${index + 1}.`);
            return;
        }

        const file = fileInput.files[0];
        filesData.push({ file, skuColumn, priceColumn });
    });

    if (filesData.length < 2) {
        alert('Please upload at least two Excel files.');
        return;
    }

    const filePromises = filesData.map(fileData => parseExcel(fileData.file));

    Promise.all(filePromises).then(dataArr => {
        compareData(dataArr, filesData);
    });
}

function parseExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            resolve(json);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function compareData(dataArr, filesData) {
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    resultsTableBody.innerHTML = ''; // Clear previous results

    resultsData = []; // Clear previous data

    const firstFileData = dataArr[0];
    const secondFileData = dataArr[1];

    const skuIndex1 = columnToIndex(filesData[0].skuColumn);
    const priceIndex1 = columnToIndex(filesData[0].priceColumn);
    const skuIndex2 = columnToIndex(filesData[1].skuColumn);
    const priceIndex2 = columnToIndex(filesData[1].priceColumn);

    firstFileData.forEach((row1) => {
        const sku1 = row1[skuIndex1];
        const price1 = parseFloat(row1[priceIndex1]);

        const matchingRow2 = secondFileData.find(row2 => row2[skuIndex2] === sku1);

        if (matchingRow2) {
            const price2 = parseFloat(matchingRow2[priceIndex2]);
            if (!isNaN(price1) && !isNaN(price2)) {
                const difference = price1 - price2;
                const row = [sku1, price1, price2, difference.toFixed(2)];

                resultsData.push(row); // Store data for export
                const rowHTML = `
                    <tr>
                        <td>${sku1}</td>
                        <td>${price1.toFixed(2)}</td>
                        <td>${price2.toFixed(2)}</td>
                        <td>${difference.toFixed(2)}</td>
                    </tr>
                `;
                resultsTableBody.insertAdjacentHTML('beforeend', rowHTML);
            }
        }
    });
}

function columnToIndex(column) {
    return column.charCodeAt(0) - 65; // Convert 'A' to 0, 'B' to 1, etc.
}

// Sorting functionality
function sortTable(columnIndex) {
    const table = document.getElementById('resultsTable');
    const rowsArray = Array.from(table.rows).slice(1); // Get table rows excluding header
    const isAscending = table.getAttribute('data-sort-asc') === 'true';

    rowsArray.sort((a, b) => {
        const aValue = parseFloat(a.cells[columnIndex].innerText);
        const bValue = parseFloat(b.cells[columnIndex].innerText);
        return isAscending ? aValue - bValue : bValue - aValue;
    });

    // Toggle sort order
    table.setAttribute('data-sort-asc', !isAscending);

    rowsArray.forEach(row => table.tBodies[0].appendChild(row)); // Reorder rows
}

// Export results to Excel
function exportResults() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['SKU', 'Price (File 1)', 'Price (File 2)', 'Difference'], ...resultsData]);
    XLSX.utils.book_append_sheet(wb, ws, 'Comparison Results');
    XLSX.writeFile(wb, 'comparison_results.xlsx');
}

// Add more file upload sections
function addMoreFile() {
    fileIndex++;
    const uploadSection = document.getElementById('upload-section');
    const newFileSection = document.createElement('div');
    newFileSection.classList.add('file-input-section');

    newFileSection.innerHTML = `
        <label for="file${fileIndex}">Excel ${fileIndex}:</label>
        <input type="file" id="file${fileIndex}" accept=".xlsx, .xls">
        <label>SKU Column:</label>
        <input type="text" class="sku-column" placeholder="e.g., A">
        <label>Price Column:</label>
        <input type="text" class="price-column" placeholder="e.g., E">
        <button class="delete-file-btn">Delete File</button>
    `;
    uploadSection.appendChild(newFileSection);

    // Add event listener for deleting file
    newFileSection.querySelector('.delete-file-btn').addEventListener('click', function() {
        newFileSection.remove();
        compareExcels(); // Recalculate the comparison after file deletion
    });
}
