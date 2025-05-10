document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const columnSelect = document.getElementById('columnSelect');
    const checkButton = document.getElementById('checkButton');
    const newButton = document.getElementById('newButton');
    const resultsSection = document.getElementById('resultsSection');
    const resultsContainer = document.getElementById('resultsContainer');
    const copyAlert = document.getElementById('copyAlert');
    const currentYearSpan = document.getElementById('currentYear');
    const selectedFileDiv = document.getElementById('selectedFile');

    // Set current year in footer
    currentYearSpan.textContent = new Date().getFullYear();

    let workbook = null;
    let worksheet = null;
    let selectedRow = null;

    fileInput.addEventListener('change', handleFileSelect);
    checkButton.addEventListener('click', checkDuplicates);
    newButton.addEventListener('click', resetApplication);

    // Add click event listener to document to remove selection
    document.addEventListener('click', (event) => {
        if (selectedRow && !event.target.closest('.duplicate-item')) {
            selectedRow.classList.remove('selected');
            selectedRow = null;
        }
    });

    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        // Show selected filename
        selectedFileDiv.innerHTML = `<i class="fas fa-file-excel me-2"></i>${file.name}`;

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
                newButton.disabled = false;
            } catch (error) {
                alert('Error reading the Excel file. Please make sure it\'s a valid Excel file.');
                console.error(error);
                selectedFileDiv.innerHTML = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function updateButtonProgress(progress) {
        checkButton.innerHTML = `
            <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>
            Processing: ${progress}%
        `;
    }

    function resetButtonState() {
        checkButton.innerHTML = '<i class="fas fa-search me-2"></i>Check Duplicates';
    }

    async function checkDuplicates() {
        if (!worksheet || columnSelect.value === '') return;

        const columnIndex = parseInt(columnSelect.value);
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const values = [];
        
        // Calculate total rows for progress
        const totalRows = range.e.r - range.s.r;
        let processedRows = 0;
        
        // Disable button and show initial progress
        checkButton.disabled = true;
        updateButtonProgress(0);
        
        // Use setTimeout to allow UI to update
        await new Promise(resolve => setTimeout(resolve, 0));
        
        // Collect all values from the selected column
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: columnIndex })];
            if (cell && cell.v !== undefined) {
                values.push(cell.v);
            }
            
            // Update progress every 100 rows or at the end
            processedRows++;
            if (processedRows % 100 === 0 || processedRows === totalRows) {
                const progress = Math.round((processedRows / totalRows) * 100);
                updateButtonProgress(progress);
                // Allow UI to update
                await new Promise(resolve => setTimeout(resolve, 0));
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
            resultsContainer.innerHTML = '<p class="text-center text-muted">No duplicates found in the selected column.</p>';
        } else {
            resultsContainer.innerHTML = duplicateEntries.map(([value, count], index) => `
                <div class="duplicate-item" data-index="${index}">
                    <span class="duplicate-value">${value}</span>
                    <span class="count"><i class="fas fa-copy me-1"></i>${count} occurrences</span>
                    <button class="btn btn-sm btn-outline-secondary copy-btn" onclick="copyToClipboard('${value.replace(/'/g, "\\'")}', ${index})">
                        <i class="fas fa-copy me-1"></i>Copy
                    </button>
                </div>
            `).join('');
        }

        // Reset button state
        resetButtonState();
        checkButton.disabled = false;
        resultsSection.style.display = 'block';
    }

    function resetApplication() {
        fileInput.value = '';
        columnSelect.innerHTML = '<option value="">Select a file first</option>';
        columnSelect.disabled = true;
        checkButton.disabled = true;
        newButton.disabled = true;
        resultsSection.style.display = 'none';
        workbook = null;
        worksheet = null;
        if (selectedRow) {
            selectedRow.classList.remove('selected');
            selectedRow = null;
        }
        resetButtonState();
        selectedFileDiv.innerHTML = '';
    }

    // Make copyToClipboard function globally available
    window.copyToClipboard = function(text, index) {
        navigator.clipboard.writeText(text).then(() => {
            // Remove selection from previously selected row
            if (selectedRow) {
                selectedRow.classList.remove('selected');
            }
            
            // Add selection to current row
            selectedRow = document.querySelector(`.duplicate-item[data-index="${index}"]`);
            if (selectedRow) {
                selectedRow.classList.add('selected');
            }

            // Remove any existing alert
            copyAlert.classList.remove('show');
            copyAlert.style.display = 'none';
            
            // Show new alert
            setTimeout(() => {
                copyAlert.style.display = 'block';
                // Force a reflow
                copyAlert.offsetHeight;
                copyAlert.classList.add('show');
                
                // Hide alert after 2 seconds
                setTimeout(() => {
                    copyAlert.classList.remove('show');
                    setTimeout(() => {
                        copyAlert.style.display = 'none';
                    }, 150);
                }, 2000);
            }, 10);
        }).catch(err => {
            console.error('Failed to copy text: ', err);
        });
    };
}); 