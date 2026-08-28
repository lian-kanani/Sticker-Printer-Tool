// Function to add a new row to the items table and copy the last row's properties if not specified
function addRow(name = "", nameFont = null, price = "", priceFont = null, currency = null, qty = 1) {
    const tbody = document.querySelector("#itemsTable tbody");
    
    // Copy the last row's properties if no specific values are provided
    if (nameFont === null || priceFont === null || currency === null) {
        const rows = tbody.querySelectorAll("tr");
        if (rows.length > 0) {
            const lastRow = rows[rows.length - 1];
            if (nameFont === null) nameFont = lastRow.querySelector('.item-name-font').value || 10;
            if (priceFont === null) priceFont = lastRow.querySelector('.item-price-font').value || 10;
            if (currency === null) currency = lastRow.querySelector('.item-currency').value;
        } else {
            if (nameFont === null) nameFont = 10;
            if (priceFont === null) priceFont = 10;
            if (currency === null) currency = "$";
        }
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="text" class="item-name" value="${name}" placeholder="الاسم" style="width: 100%;"></td>
        <td><input type="number" class="item-name-font" value="${nameFont}" step="0.5" style="width: 100%; text-align: center;"></td>
        <td><input type="text" class="item-price" value="${price}" placeholder="السعر" style="width: 100%;"></td>
        <td>
            <select class="item-currency" style="width: 100%;">
                <option value="$" ${currency === '$' ? 'selected' : ''}>$</option>
                <option value="SP" ${currency === 'SP' ? 'selected' : ''}>SP</option>
            </select>
        </td>
        <td><input type="number" class="item-price-font" value="${priceFont}" step="0.5" style="width: 100%; text-align: center;"></td>
        <td><input type="number" class="item-qty" value="${qty}" min="1" style="width: 100%; text-align: center;"></td>
        <td><button class="btn btn-danger" onclick="removeRow(this)">حذف</button></td>
    `;
    tbody.appendChild(tr);
}

// Function to remove a specific row from the table
function removeRow(btn) {
    btn.parentElement.parentElement.remove();
}

// Import data from an uploaded Excel file
document.getElementById('excelUpload').addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (!file) return;

    var reader = new FileReader();
    reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, {type: 'array'});
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        var json = XLSX.utils.sheet_to_json(firstSheet, {header: 1, defval: ""});
        
        var hasHeaders = document.getElementById('hasHeaders').checked;
        var dataRows = hasHeaders ? json.slice(1) : json;
        
        dataRows.forEach(row => {
            // Column order in your file: 0:Name, 1:Name Font, 2:Price, 3:Price Font, 4:Currency, 5:Quantity
            if(row[0] !== "" || row[2] !== "") {
                let name = row[0] || "";
                let nameFont = row[1] !== "" ? parseFloat(row[1]) : null;
                let price = row[2] !== "" ? row[2] : "";
                let priceFont = row[3] !== "" ? parseFloat(row[3]) : null;
                
                let currencyRaw = (row[4] || "").toString().toLowerCase();
                let currency = null; // Default to null, will be determined below
                
                // If the currency is Syrian Pound or similar
                if (currencyRaw.includes("sp") || currencyRaw.includes("ليرة") || currencyRaw.includes("سورية")) {
                    currency = "SP"; 
                } 
                // If the currency is US Dollar or similar
                else if (currencyRaw.includes("$") || currencyRaw.includes("دولار")) {
                    currency = "$"; 
                }
                
                let qty = row[5] !== "" ? parseInt(row[5]) : 1; 
                
                addRow(name, nameFont, price, priceFont, currency, qty);
            }
        });
        alert("✅ Data imported successfully!");
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset the file input
});

// Generate preview and populate the print container
function generatePreview() {
    const marginTop = parseFloat(document.getElementById('marginTop').value) || 0;
    const marginLeft = parseFloat(document.getElementById('marginLeft').value) || 0;
    const labelHeight = parseFloat(document.getElementById('labelHeight').value) || 1.5;
    const labelWidth = parseFloat(document.getElementById('labelWidth').value) || 2.2;
    const vStep = parseFloat(document.getElementById('vStep').value) || 1.75;
    const hStep = parseFloat(document.getElementById('hStep').value) || 2.55;
    const cols = parseInt(document.getElementById('cols').value) || 8;
    const rows = parseInt(document.getElementById('rows').value) || 16;
    const pageWidth = parseFloat(document.getElementById('pageWidth').value) || 21;
    const pageHeight = parseFloat(document.getElementById('pageHeight').value) || 29.97;
    const skipLabels = parseInt(document.getElementById('skipLabels').value) || 0;

    const colGap = Math.max(0, hStep - labelWidth); 
    const rowGap = Math.max(0, vStep - labelHeight); 

    const tableRows = document.querySelectorAll("#itemsTable tbody tr");
    let allLabels = [];

    // Adding blank labels for skipped labels
    for (let i = 0; i < skipLabels; i++) {
        allLabels.push({ blank: true });
    }

    // Gathering all labels from the table inputs
    tableRows.forEach(row => {
        const name = row.querySelector('.item-name').value.trim();
        const nameFont = row.querySelector('.item-name-font').value || 10;
        const price = row.querySelector('.item-price').value.trim();
        const currency = row.querySelector('.item-currency').value;
        const priceFont = row.querySelector('.item-price-font').value || 10;
        const qty = parseInt(row.querySelector('.item-qty').value) || 0;

        if (name === "" && price === "") return; 

        for (let i = 0; i < qty; i++) {
            allLabels.push({ name, nameFont, price, currency, priceFont, blank: false });
        }
    });

    if (allLabels.length === skipLabels) {
        alert("⚠️ No data available to print!");
        return;
    }

    const container = document.getElementById('print-container');
    container.innerHTML = ''; 
    
    const labelsPerPage = cols * rows; 
    const totalPages = Math.ceil(allLabels.length / labelsPerPage);

    for (let p = 0; p < totalPages; p++) {
        const sheetDiv = document.createElement('div');
        sheetDiv.className = 'a4-sheet';
        
        sheetDiv.style.width = pageWidth + 'cm';
        sheetDiv.style.height = pageHeight + 'cm';
        sheetDiv.style.paddingTop = marginTop + 'cm';
        sheetDiv.style.paddingLeft = marginLeft + 'cm';
        sheetDiv.style.display = 'grid';
        sheetDiv.style.gridTemplateColumns = `repeat(${cols}, ${labelWidth}cm)`;
        sheetDiv.style.gridTemplateRows = `repeat(${rows}, ${labelHeight}cm)`;
        sheetDiv.style.columnGap = colGap + 'cm';
        sheetDiv.style.rowGap = rowGap + 'cm';

        const startIndex = p * labelsPerPage;
        const endIndex = Math.min(startIndex + labelsPerPage, allLabels.length);

        for (let i = startIndex; i < endIndex; i++) {
            const item = allLabels[i];
            const labelBox = document.createElement('div');
            labelBox.className = 'label-box';
            labelBox.style.width = labelWidth + 'cm';
            labelBox.style.height = labelHeight + 'cm';
            
            if (item && !item.blank) {
                let priceHtml = '';
                if (item.price !== "") {
                    // Add a span for the currency next to the price
                    priceHtml = `
                    <div class="label-price" style="font-size: ${item.priceFont}pt;">
                        <span>${item.price}</span>
                        <span>${item.currency}</span>
                    </div>`;
                }
                
                labelBox.innerHTML = `
                    <div class="label-name" style="font-size: ${item.nameFont}pt;">${item.name}</div>
                    ${priceHtml}
                `;
            } else {
                // If it's a blank label, leave it empty without borders
                labelBox.style.border = 'none'; 
            }
            sheetDiv.appendChild(labelBox);
        }
        container.appendChild(sheetDiv);
    }

    // Show the print button after generating the preview
    document.getElementById('printBtn').style.display = 'block';
    
    // Scroll to the print button after a short delay to ensure the layout has updated
    setTimeout(() => {
        document.getElementById('printBtn').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
}

// Bind the print event to the print button
document.getElementById('printBtn').addEventListener('click', function() {
    window.print();
});