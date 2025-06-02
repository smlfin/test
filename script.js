document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";

    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const reportDisplay = document.getElementById('reportDisplay');

    // *** Data Storage ***
    let allCanvassingData = []; // Stores all fetched data
    let filteredBranchData = []; // Stores data for the currently selected branch
    let selectedEmployeeEntries = []; // Stores entries for the currently selected employee

    // *** Utility Functions ***

    // Simple CSV to JSON converter
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length <= 1) return []; // Only headers or no data

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            let entry = {};
            headers.forEach((header, index) => {
                // Ensure the header exists and value is not out of bounds
                if (header && values[index] !== undefined) {
                    entry[header] = values[index];
                }
            });
            data.push(entry);
        }
        return data;
    }

    // Displays a message in the report area
    function displayMessage(message) {
        reportDisplay.innerHTML = `<p>${message}</p>`;
    }

    // Populates a dropdown with unique values
    function populateDropdown(dropdownElement, dataArray, key) {
        // Clear existing options, but keep the default "Select" option
        dropdownElement.innerHTML = `<option value="">-- Select a ${key.replace(' Name', '')} --</option>`;
        const uniqueValues = new Set();
        dataArray.forEach(item => {
            if (item[key] && item[key].trim() !== '') {
                uniqueValues.add(item[key].trim());
            }
        });

        // Sort names alphabetically for better UX
        Array.from(uniqueValues).sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            dropdownElement.appendChild(option);
        });
    }

    // Renders all detailed entries for a given employee
    function renderEmployeeDetailedEntries(entries) {
        reportDisplay.innerHTML = `<h3>Detailed Entries for ${entries[0]['Employee Name']}</h3>`;
        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this employee.</p>`;
            return;
        }

        const detailsList = document.createElement('ul');
        detailsList.className = 'employee-detailed-list';

        entries.forEach(entry => {
            const entryLi = document.createElement('li');
            entryLi.className = 'entry-detail-item';
            let detailHtml = '<h4>Canvassing Entry</h4>';

            // Iterate over all possible columns from your data and display them
            const columnsToDisplay = [
                "Timestamp", "Date", "Branch Name", "Employee Name", "Employee Code",
                "Designation", "Activity Type", "Type of Customer", "Lead Source",
                "How Contacted", "Prospect Name", "Phone Numebr(Whatsapp)", "Address",
                "Profession", "DOB/WD", "Prodcut Interested", "Remarks",
                "Next Follow-up Date", "Relation With Staff"
            ];

            columnsToDisplay.forEach(col => {
                // Exclude Branch Name and Employee Name from individual entry details
                // as they are already part of the grouping hierarchy
                if (col !== "Branch Name" && col !== "Employee Name") {
                     detailHtml += `<p><strong>${col}:</strong> ${entry[col] || 'N/A'}</p>`;
                }
            });

            entryLi.innerHTML = detailHtml;
            detailsList.appendChild(entryLi);
        });
        reportDisplay.appendChild(detailsList);
    }

    // Renders a summary for a given employee
    function renderEmployeeSummary(entries) {
        const employeeName = entries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h3>Summary for ${employeeName}</h3>`;
        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this employee.</p>`;
            return;
        }

        // --- Calculate Summary Metrics ---
        const totalEntries = entries.length;
        const activityTypeCounts = {};
        const customerTypeCounts = {};
        const leadSourceCounts = {};
        const productInterestedCounts = {};
        const professionCounts = {};

        entries.forEach(entry => {
            const activity = entry['Activity Type'];
            const customerType = entry['Type of Customer'];
            const leadSource = entry['Lead Source'];
            const product = entry['Prodcut Interested'];
            const profession = entry['Profession'];

            if (activity) activityTypeCounts[activity] = (activityTypeCounts[activity] || 0) + 1;
            if (customerType) customerTypeCounts[customerType] = (customerTypeCounts[customerType] || 0) + 1;
            if (leadSource) leadSourceCounts[leadSource] = (leadSourceCounts[leadSource] || 0) + 1;
            if (product) productInterestedCounts[product] = (productInterestedCounts[product] || 0) + 1;
            if (profession) professionCounts[profession] = (professionCounts[profession] || 0) + 1;
        });

        // --- Display Summary ---
        let summaryHtml = `<p><strong>Total Canvassing Entries:</strong> ${totalEntries}</p>`;

        summaryHtml += `<h4>Activity Types:</h4><ul>`;
        for (const type in activityTypeCounts) {
            summaryHtml += `<li>${type}: ${activityTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Customer Types:</h4><ul>`;
        for (const type in customerTypeCounts) {
            summaryHtml += `<li>${type}: ${customerTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Lead Sources:</h4><ul>`;
        for (const source in leadSourceCounts) {
            summaryHtml += `<li>${source}: ${leadSourceCounts[source]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Products Interested:</h4><ul>`;
        for (const product in productInterestedCounts) {
            summaryHtml += `<li>${product}: ${productInterestedCounts[product]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Professions:</h4><ul>`;
        for (const profession in professionCounts) {
            summaryHtml += `<li>${profession}: ${professionCounts[profession]}</li>`;
        }
        summaryHtml += `</ul>`;


        reportDisplay.innerHTML += summaryHtml;
    }

    // --- Main Fetch and Event Listeners ---

    // Initial data fetch on page load
    async function fetchData() {
        displayMessage("Fetching data from Google Sheet...");
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            console.log("Data loaded successfully:", allCanvassingData);

            if (allCanvassingData.length > 0) {
                populateDropdown(branchSelect, allCanvassingData, 'Branch Name');
                displayMessage("Data loaded. Please select a branch.");
            } else {
                displayMessage("No data found in the Google Sheet.");
            }

        } catch (error) {
            console.error("Error fetching or processing data:", error);
            displayMessage("Error loading data. Please try again later. Check browser console for details.");
            // Hide all controls if data loading fails
            branchSelect.style.display = 'none';
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
        }
    }

    // Event listener for Branch selection
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            filteredBranchData = allCanvassingData.filter(entry => entry['Branch Name'] === selectedBranch);
            populateDropdown(employeeSelect, filteredBranchData, 'Employee Name');
            employeeFilterPanel.style.display = 'block'; // Show employee filter
            viewOptions.style.display = 'none'; // Hide view options until employee selected
            employeeSelect.value = ""; // Reset employee selection
            displayMessage(`Selected Branch: ${selectedBranch}. Please select an employee.`);
        } else {
            employeeFilterPanel.style.display = 'none'; // Hide employee filter
            viewOptions.style.display = 'none'; // Hide view options
            displayMessage("Please select a branch from the dropdown above to view reports.");
        }
        reportDisplay.innerHTML = ''; // Clear report display
    });

    // Event listener for Employee selection
    employeeSelect.addEventListener('change', () => {
        const selectedEmployee = employeeSelect.value;
        if (selectedEmployee) {
            selectedEmployeeEntries = filteredBranchData.filter(entry => entry['Employee Name'] === selectedEmployee);
            viewOptions.style.display = 'block'; // Show view options
            displayMessage(`Selected Employee: ${selectedEmployee}. Choose a view option.`);
        } else {
            viewOptions.style.display = 'none'; // Hide view options
            displayMessage(`Selected Branch: ${branchSelect.value}. Please select an employee.`);
        }
        reportDisplay.innerHTML = ''; // Clear report display
    });

    // Event listeners for View Options
    viewAllEntriesBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeDetailedEntries(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee.");
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeSummary(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee.");
        }
    });

    // Initial data fetch when the page loads
    fetchData();
});