const form = document.querySelector('#myForm');

form.addEventListener('submit', (event) => {
  event.preventDefault(); // prevent the default form submission behavior

  const nameInput = document.querySelector('#nameInput').value;
  const emailInput = document.querySelector('#emailInput');
  const telInput = document.querySelector('#telInput');
  const numberInput = document.querySelector('#numberInput');

  // Load the Excel file using SheetJS
  const xhr = new XMLHttpRequest();
  xhr.open('GET', 'data.xlsx', true);
  xhr.responseType = 'arraybuffer';
  xhr.onload = function() {
    const data = new Uint8Array(xhr.response);
    const workbook = XLSX.read(data, { type: 'array' });

    // Get the first sheet in the workbook
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Search for the name in the sheet and get the corresponding data
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
      const nameCell = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
      if (nameCell && nameCell.v === nameInput) {
        const emailCell = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })];
        if (emailCell) {
          emailInput.value = emailCell.v;
        }
        const telCell = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 2 })];
        if (telCell) {
          telInput.value = telCell.v;
        }
        const numberCell = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 3 })];
        if (numberCell) {
          numberInput.value = numberCell.v;
        }
        break;
      }
    }
  };
  xhr.send();
});
