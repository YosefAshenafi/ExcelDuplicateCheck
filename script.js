document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const columnSelect = document.getElementById('columnSelect');
    const checkButton = document.getElementById('checkButton');
    const resultsSection = document.getElementById('resultsSection');
    const resultsContainer = document.getElementById('resultsContainer');

    let workbook = null;
    let worksheet = null;

    fileInput.addEventListener('change', handleFileSelect);
    checkButton.addEventListener('click', checkDuplicates);

    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // Get headers from the first row
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                const headers = [];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
                    if (cell && cell.v) headers.push(cell.v);
                }

                // Populate column select
                columnSelect.innerHTML = headers.map((header, index) => 
                    `<option value="${index}">${header}</option>`
                ).join('');
                
                columnSelect.disabled = false;
                checkButton.disabled = false;
            } catch (error) {
                alert('Error reading the Excel file. Please make sure it\'s a valid Excel file.');
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function checkDuplicates() {
        if (!worksheet || columnSelect.value === '') return;

        const columnIndex = parseInt(columnSelect.value);
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const values = [];
        
        // Collect all values from the selected column
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: columnIndex })];
            if (cell && cell.v !== undefined) {
                values.push(cell.v);
            }
        }

        // Count duplicates
        const duplicates = values.reduce((acc, value) => {
            acc[value] = (acc[value] || 0) + 1;
            return acc;
        }, {});

        // Filter only duplicates (count > 1)
        const duplicateEntries = Object.entries(duplicates)
            .filter(([_, count]) => count > 1)
            .sort((a, b) => b[1] - a[1]);

        // Display results
        if (duplicateEntries.length === 0) {
            resultsContainer.innerHTML = '<p>No duplicates found in the selected column.</p>';
        } else {
            resultsContainer.innerHTML = duplicateEntries.map(([value, count]) => `
                <div class="duplicate-item">
                    <span>${value}</span>
                    <span class="count">${count} occurrences</span>
                </div>
            `).join('');
        }

        resultsSection.style.display = 'block';
    }
}); 