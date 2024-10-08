document.getElementById('compareBtn').addEventListener('click', compareExcels);

function compareExcels() {
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    const skuColumn1 = document.getElementById('skuColumn1').value.toUpperCase();
    const priceColumn1 = document.getElementById('priceColumn1').value.toUpperCase();
    const skuColumn2 = document.getElementById('skuColumn2').value.toUpperCase();
    const priceColumn2 = document.getElementById('priceColumn2').value.toUpperCase();

    if (!file1 || !file2) {
        alert('Please upload both Excel files.');
        return;
    }

    if (!skuColumn1 || !priceColumn1 || !skuColumn2 || !priceColumn2) {
        alert('Please specify columns to compare.');
        return;
    }

    // Parse both Excel files
    parseExcel(file1, (data1) => {
        parseExcel(file2, (data2) => {
            compareData(data1, data2, skuColumn1, priceColumn1, skuColumn2, priceColumn2);
        });
    });
}

function parseExcel(file, callback) {
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Assume the first sheet
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        callback(json);
    };

    reader.readAsArrayBuffer(file);
}

function compareData(data1, data2, skuCol1, priceCol1, skuCol2, priceCol2) {
    const skuIndex1 = columnToIndex(skuCol1);
    const priceIndex1 = columnToIndex(priceCol1);
    const skuIndex2 = columnToIndex(skuCol2);
    const priceIndex2 = columnToIndex(priceCol2);

    const resultsTable = document.getElementById('resultsTable');
    resultsTable.innerHTML = ''; // Clear previous results

    data1.forEach((row1) => {
        const sku1 = row1[skuIndex1];
        const price1 = parseFloat(row1[priceIndex1]);

        // Find matching SKU in Excel 2
        const matchingRow2 = data2.find((row2) => row2[skuIndex2] === sku1);

        // Only process if there is a matching SKU and valid price data
        if (matchingRow2) {
            const price2 = parseFloat(matchingRow2[priceIndex2]);

            // Check if both prices are valid numbers
            if (!isNaN(price1) && !isNaN(price2)) {
                const difference = price1 - price2;

                // Display the row with SKU and price data
                const rowHTML = `
                    <tr>
                        <td>${sku1}</td>
                        <td>${price1.toFixed(2)}</td>
                        <td>${price2.toFixed(2)}</td>
                        <td>${difference.toFixed(2)}</td>
                    </tr>
                `;
                resultsTable.insertAdjacentHTML('beforeend', rowHTML);
            }
        }
    });
}

function columnToIndex(column) {
    return column.charCodeAt(0) - 65; // Convert 'A' to 0, 'B' to 1, etc.
}
