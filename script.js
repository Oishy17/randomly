document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const numberInput = document.getElementById('numberInput');
    const datasetInfo = document.getElementById('datasetInfo');
    const selectBtn = document.getElementById('selectBtn');
    const resetBtn = document.getElementById('resetBtn');
    const previewSection = document.getElementById('previewSection');
    const previewHeader = document.getElementById('previewHeader');
    const previewBody = document.getElementById('previewBody');
    const previewCount = document.getElementById('previewCount');
    const resultsSection = document.getElementById('resultsSection');
    const resultsHeader = document.getElementById('resultsHeader');
    const resultsBody = document.getElementById('resultsBody');
    const selectedCount = document.getElementById('selectedCount');
    const totalCount = document.getElementById('totalCount');
    const comparisonBody = document.getElementById('comparisonBody');
    const comparisonHeader = document.getElementById('comparisonHeader');
    const selectedHighlightCount = document.getElementById('selectedHighlightCount');
    const totalHighlightCount = document.getElementById('totalHighlightCount');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportCsvBtn = document.getElementById('exportCsvBtn');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const notification = document.getElementById('notification');
    const tabs = document.querySelectorAll('.tab');
    const tabContents = document.querySelectorAll('.tab-content');
    
    let dataset = [];
    let selectedPeople = [];
    let headers = [];
    let originalFileName = '';
    
    // Pagination variables
    let currentPage = 1;
    const recordsPerPage = 100;
    let paginationControls;

    // Show notification
    function showNotification(message, type) {
        notification.textContent = message;
        notification.className = `notification ${type}`;
        notification.classList.add('show');
        
        setTimeout(() => {
            notification.classList.remove('show');
        }, 3000);
    }
    
    // Tab switching functionality
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabId = tab.getAttribute('data-tab');
            
            // Update active tab
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // Show corresponding content
            tabContents.forEach(content => {
                content.classList.remove('active');
                if (content.id === `${tabId}Tab`) {
                    content.classList.add('active');
                }
            });
        });
    });
    
    // Handle file selection
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            fileName.textContent = file.name;
            originalFileName = file.name.replace(/\.[^/.]+$/, ""); // Remove extension
            
            // Show loading indicator
            loadingIndicator.style.display = 'block';
            
            // Read the Excel file
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Get the first worksheet
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // Convert to JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    if (jsonData.length > 0) {
                        dataset = jsonData;
                        headers = Object.keys(dataset[0]);
                        
                        // Update UI
                        totalCount.textContent = dataset.length;
                        previewCount.textContent = dataset.length;
                        datasetInfo.innerHTML = `
                            <strong>${dataset.length}</strong> people loaded<br>
                            <strong>${headers.length}</strong> data columns
                        `;
                        selectBtn.disabled = false;
                        
                        // Reset pagination
                        currentPage = 1;
                        
                        // Display preview with pagination
                        displayPreviewWithPagination(dataset);
                        previewSection.style.display = 'block';
                        
                        showNotification('Dataset loaded successfully!', 'success');
                        loadingIndicator.style.display = 'none';
                        
                        // Set max value for number input
                        numberInput.max = dataset.length;
                    } else {
                        alert('The Excel file appears to be empty or improperly formatted.');
                        loadingIndicator.style.display = 'none';
                    }
                } catch (error) {
                    console.error('Error processing Excel file:', error);
                    showNotification('Error processing the Excel file', 'error');
                    loadingIndicator.style.display = 'none';
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            fileName.textContent = 'No file selected';
            selectBtn.disabled = true;
        }
    });
    
    // Display dataset preview with pagination
    function displayPreviewWithPagination(data) {
        previewHeader.innerHTML = '';
        previewBody.innerHTML = '';
        
        // Create header row
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            previewHeader.appendChild(th);
        });
        
        // Calculate pagination
        const totalPages = Math.ceil(data.length / recordsPerPage);
        const startIndex = (currentPage - 1) * recordsPerPage;
        const endIndex = Math.min(startIndex + recordsPerPage, data.length);
        const pageData = data.slice(startIndex, endIndex);
        
        // Create data rows for current page
        pageData.forEach((person, index) => {
            const row = document.createElement('tr');
            
            headers.forEach(header => {
                const cell = document.createElement('td');
                cell.textContent = person[header] || 'N/A';
                row.appendChild(cell);
            });
            
            previewBody.appendChild(row);
        });
        
        // Create pagination controls if they don't exist
        if (!paginationControls) {
            paginationControls = document.createElement('div');
            paginationControls.className = 'pagination-controls';
            paginationControls.style.marginTop = '20px';
            paginationControls.style.display = 'flex';
            paginationControls.style.justifyContent = 'center';
            paginationControls.style.alignItems = 'center';
            paginationControls.style.gap = '15px';
            
            const prevButton = document.createElement('button');
            prevButton.textContent = 'Previous';
            prevButton.className = 'pagination-btn';
            
            const nextButton = document.createElement('button');
            nextButton.textContent = 'Next';
            nextButton.className = 'pagination-btn';
            
            const pageInfo = document.createElement('span');
            pageInfo.className = 'page-info';
            pageInfo.style.fontWeight = '600';
            pageInfo.style.color = '#2c3e50';
            
            paginationControls.appendChild(prevButton);
            paginationControls.appendChild(pageInfo);
            paginationControls.appendChild(nextButton);
            
            // Add event listeners for pagination
            prevButton.addEventListener('click', () => {
                if (currentPage > 1) {
                    currentPage--;
                    displayPreviewWithPagination(dataset);
                }
            });
            
            nextButton.addEventListener('click', () => {
                if (currentPage < totalPages) {
                    currentPage++;
                    displayPreviewWithPagination(dataset);
                }
            });
            
            // Add pagination controls after the table
            const tableContainer = document.querySelector('.table-container');
            tableContainer.parentNode.insertBefore(paginationControls, tableContainer.nextSibling);
        }
        
        // Update pagination info
        const pageInfo = paginationControls.querySelector('.page-info');
        pageInfo.textContent = `Page ${currentPage} of ${totalPages} (Records ${startIndex + 1}-${endIndex} of ${data.length})`;
        
        // Update button states
        const prevButton = paginationControls.querySelector('.pagination-btn:first-child');
        const nextButton = paginationControls.querySelector('.pagination-btn:last-child');
        
        prevButton.disabled = currentPage === 1;
        nextButton.disabled = currentPage === totalPages;
        
        // Style pagination buttons
        const paginationButtons = paginationControls.querySelectorAll('.pagination-btn');
        paginationButtons.forEach(button => {
            button.style.padding = '8px 16px';
            button.style.border = 'none';
            button.style.borderRadius = '5px';
            button.style.backgroundColor = '#3498db';
            button.style.color = 'white';
            button.style.cursor = 'pointer';
            button.style.fontWeight = '600';
            button.style.transition = 'all 0.3s';
            
            button.addEventListener('mouseenter', function() {
                if (!button.disabled) {
                    this.style.backgroundColor = '#2980b9';
                    this.style.transform = 'translateY(-2px)';
                }
            });
            
            button.addEventListener('mouseleave', function() {
                if (!button.disabled) {
                    this.style.backgroundColor = '#3498db';
                    this.style.transform = 'translateY(0)';
                }
            });
            
            if (button.disabled) {
                button.style.backgroundColor = '#bdc3c7';
                button.style.cursor = 'not-allowed';
                button.style.opacity = '0.6';
            }
        });
    }
    
    // Handle random selection
    selectBtn.addEventListener('click', function() {
        if (dataset.length === 0) {
            showNotification('Please upload a valid Excel file first.', 'error');
            return;
        }
        
        const count = parseInt(numberInput.value);
        
        if (isNaN(count) || count < 1) {
            showNotification('Please enter a valid number of people to select.', 'error');
            return;
        }
        
        if (count > dataset.length) {
            showNotification(`You can only select up to ${dataset.length} people from your dataset.`, 'error');
            numberInput.value = dataset.length;
            return;
        }
        
        // Use Fisher-Yates shuffle for truly random selection
        selectedPeople = getRandomSubset(dataset, count);
        
        // Display results
        displayResults(selectedPeople);
        displayComparison(dataset, selectedPeople);
        resultsSection.style.display = 'block';
        
        // Scroll to results
        resultsSection.scrollIntoView({ behavior: 'smooth' });
        
        showNotification(`Successfully selected ${count} random people!`, 'success');
    });
    
    // Handle reset
    resetBtn.addEventListener('click', function() {
        fileInput.value = '';
        fileName.textContent = 'No file selected';
        selectBtn.disabled = true;
        previewSection.style.display = 'none';
        resultsSection.style.display = 'none';
        dataset = [];
        selectedPeople = [];
        headers = [];
        datasetInfo.textContent = 'No dataset loaded';
        numberInput.value = '20';
        originalFileName = '';
        
        // Remove pagination controls if they exist
        if (paginationControls) {
            paginationControls.remove();
            paginationControls = null;
        }
        
        showNotification('Application reset successfully.', 'success');
    });
    
    // Handle Excel export
    exportExcelBtn.addEventListener('click', function() {
        if (selectedPeople.length === 0) {
            showNotification('No data to export. Please select people first.', 'error');
            return;
        }
        
        // Create a new workbook
        const wb = XLSX.utils.book_new();
        
        // Convert JSON to worksheet
        const ws = XLSX.utils.json_to_sheet(selectedPeople);
        
        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, "Selected People");
        
        // Generate Excel file and trigger download
        const fileName = `${originalFileName || 'random_selection'}_${selectedPeople.length}_people.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        showNotification('Excel file downloaded successfully!', 'success');
    });
    
    // Handle CSV export
    exportCsvBtn.addEventListener('click', function() {
        if (selectedPeople.length === 0) {
            showNotification('No data to export. Please select people first.', 'error');
            return;
        }
        
        // Convert to CSV
        const csvContent = convertToCSV(selectedPeople);
        
        // Create download link
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        const fileName = `${originalFileName || 'random_selection'}_${selectedPeople.length}_people.csv`;
        link.setAttribute('download', fileName);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showNotification('CSV file downloaded successfully!', 'success');
    });
    
    // Fisher-Yates shuffle algorithm for random selection
    function getRandomSubset(array, count) {
        const shuffled = [...array];
        let currentIndex = shuffled.length;
        let temporaryValue, randomIndex;
        
        // While there remain elements to shuffle
        while (currentIndex !== 0) {
            // Pick a remaining element
            randomIndex = Math.floor(Math.random() * currentIndex);
            currentIndex -= 1;
            
            // Swap it with the current element
            temporaryValue = shuffled[currentIndex];
            shuffled[currentIndex] = shuffled[randomIndex];
            shuffled[randomIndex] = temporaryValue;
        }
        
        // Return the first 'count' elements
        return shuffled.slice(0, Math.min(count, array.length));
    }
    
    // Display results in the table
    function displayResults(people) {
        resultsHeader.innerHTML = '';
        resultsBody.innerHTML = '';
        selectedCount.textContent = people.length;
        
        // Create header row
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            resultsHeader.appendChild(th);
        });
        
        // Create data rows
        people.forEach((person, index) => {
            const row = document.createElement('tr');
            
            headers.forEach(header => {
                const cell = document.createElement('td');
                cell.textContent = person[header] || 'N/A';
                row.appendChild(cell);
            });
            
            resultsBody.appendChild(row);
        });
    }
    
    // Display comparison view
    function displayComparison(fullData, selectedData) {
        comparisonHeader.innerHTML = '';
        comparisonBody.innerHTML = '';
        selectedHighlightCount.textContent = selectedData.length;
        totalHighlightCount.textContent = fullData.length;
        
        // Create header row
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            comparisonHeader.appendChild(th);
        });
        
        // Create data rows (show first 100 rows for performance)
        const displayData = fullData.slice(0, 100);
        displayData.forEach((person, index) => {
            const row = document.createElement('tr');
            
            // Check if this person is in the selected group
            const isSelected = selectedData.some(selected => 
                JSON.stringify(selected) === JSON.stringify(person)
            );
            
            if (isSelected) {
                row.classList.add('highlighted');
            }
            
            headers.forEach(header => {
                const cell = document.createElement('td');
                cell.textContent = person[header] || 'N/A';
                row.appendChild(cell);
            });
            
            comparisonBody.appendChild(row);
        });
        
        if (fullData.length > 100) {
            const messageRow = document.createElement('tr');
            const messageCell = document.createElement('td');
            messageCell.colSpan = headers.length;
            messageCell.textContent = `... and ${fullData.length - 100} more rows (showing first 100)`;
            messageCell.style.textAlign = 'center';
            messageCell.style.fontStyle = 'italic';
            messageCell.style.color = '#7f8c8d';
            messageRow.appendChild(messageCell);
            comparisonBody.appendChild(messageRow);
        }
    }
    
    // Convert data to CSV
    function convertToCSV(data) {
        if (data.length === 0) return '';
        
        const csvRows = [headers.join(',')];
        
        for (const row of data) {
            const values = headers.map(header => {
                const value = row[header] || '';
                // Escape quotes and wrap in quotes if contains comma
                const escaped = String(value).replace(/"/g, '""');
                return escaped.includes(',') ? `"${escaped}"` : escaped;
            });
            csvRows.push(values.join(','));
        }
        
        return csvRows.join('\n');
    }
});