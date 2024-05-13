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
    A: 'Model Id:',
    B: 'Item Name:',
    C: 'Description:',
    D: 'Technical Specification:',
  };

  //VALIDATION
  if (!file) {
    alert('Please select a file.');
    return;
  }

  if (!sheetName) {
    alert('Please enter a sheet name.');
    return;
  }

  if (
    isNaN(startRow) ||
    isNaN(endRow) ||
    startRow < 1 ||
    endRow < 1 ||
    startRow > endRow
  ) {
    alert('Please enter valid start and end row numbers.');
    return;
  }

  const reader = new FileReader();

  //FETCHING EXCEL
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    if (!workbook.Sheets[sheetName]) {
      alert(`Sheet '${sheetName}' not found.`);
      return;
    }

    //YUNG NA FETCH ILALAGAY SA ZIP
    const sheet = workbook.Sheets[sheetName];
    const zip = new JSZip();

    let zipFilename = `row-${startRow}-${endRow}.zip`;
    let promises = [];

    //LOGIC WORKS HERE
    for (let row = startRow; row <= endRow; row++) {
      let rowText = '';
      let itemName = '';
      for (let col in columnTexts) {
        const cellAddress = col + row;
        let cellValue = sheet[cellAddress]
          ? sheet[cellAddress].v.split('\n').map(line => line.trim()).join('\n')
          : 'N/A';
        if (cellValue === '') cellValue = 'N/A';
        rowText += columnTexts[col] + '\n' + cellValue + '\n\n';
        if (col === 'B') {
          itemName = cellValue
            .replace(/\//g, '-')
            .replace(/"/g, '')
            .replace(/[^a-zA-Z0-9\s-]/g, '');
        }
      }
      zip.file(`${itemName}.txt`, rowText, { createFolders: false });
    }

    zip.generateAsync({ type: 'blob' }).then(function (content) {
      saveAs(content, zipFilename);
    });
  };

  reader.readAsArrayBuffer(file);
}
