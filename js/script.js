function convertAndSave() {
    const fileInput = document.getElementById('fileInput');
    const sheetNameInput = document.getElementById('sheetNameInput');
    const startRowInput = document.getElementById('startRowInput');
    const endRowInput = document.getElementById('endRowInput');
    
    const file = fileInput.files[0];
    const sheetName = sheetNameInput.value.trim();
    const startRow = parseInt(startRowInput.value);
    const endRow = parseInt(endRowInput.value);
    
    const columnTexts = {
        'A': 'Model Id:',
        'B': 'Item Name:',
        'C': 'Description:',
        'D': 'Technical Specification:'
    };

    if (!file) {
        alert('Please select a file.');
        return;
    }

    if (!sheetName) {
        alert('Please enter a sheet name.');
        return;
    }

    if (isNaN(startRow) || isNaN(endRow) || startRow < 1 || endRow < 1 || startRow > endRow) {
        alert('Please enter valid start and end row numbers.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Check if the specified sheet exists
        if (!workbook.Sheets[sheetName]) {
            alert(`Sheet '${sheetName}' not found.`);
            return;
        }

        const sheet = workbook.Sheets[sheetName];

        for (let row = startRow; row <= endRow; row++) {
            let rowText = '';
            let itemName = ''; // Variable to store the Item Name
            for (let col in columnTexts) {
                const cellAddress = col + row;
                const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : '';
                rowText += columnTexts[col] + '\n' + cellValue + '\n\n'; // Add an extra space line after each column
                if (col === 'B') {
                    itemName = cellValue; // Get the Item Name from column B
                }
            }
            // Append each row to the text file
            const blob = new Blob([rowText], { type: 'text/plain;charset=utf-8' });
            const fileName = `${itemName}.txt`; // Use the Item Name as the filename
            saveAs(blob, fileName);
        }
    };

    reader.readAsArrayBuffer(file);
}
