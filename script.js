document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";

    // Define monthly targets based on your clarification
    // Assuming 22 working days for daily targets
    const MONTHLY_WORKING_DAYS = 22; // As of June 2025, a common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,  // Visits/month
            'Call': 3 * MONTHLY_WORKING_DAYS, // Calls/month
            'Reference': 1 * MONTHLY_WORKING_DAYS, // References/month
            'New Customer Leads': 20 // Leads/month
        },
        'Default': { // For all other designations not explicitly defined
            'Visit': 5,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        }
    };

    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchSummaryBtn = document.getElementById('viewBranchSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
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
        dropdownElement.innerHTML = `<option value="">-- Select a ${key.replace(' Name', '')} --</option>`;
        const uniqueValues = new Set();
        dataArray.forEach(item => {
            if (item[key] && item[key].trim() !== '') {
                uniqueValues.add(item[key].trim());
            }
        });

        Array.from(uniqueValues).sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            dropdownElement.appendChild(option);
        });
    }

    // Filters data for the current month and year
    function filterDataForCurrentMonth(data) {
        const now = new Date();
        const currentMonth = now.getMonth(); // 0-indexed
        const currentYear = now.getFullYear();

        return data.filter(entry => {
            // Your 'Date' format is 'DD/MM/YY' (e.g., '01/06/25')
            if (entry['Date']) {
                const dateParts = entry['Date'].split('/');
                if (dateParts.length === 3) {
                    const entryDay = parseInt(dateParts[0], 10);
                    const entryMonth = parseInt(dateParts[1], 10) - 1; // Convert to 0-indexed for JS Date
                    let entryYear = parseInt(dateParts[2], 10);
                    // Handle 2-digit year (e.g., '25' -> 2025). Assumes years are 2000-2099
                    entryYear = (entryYear < 70 ? 2000 + entryYear : 1900 + entryYear); // Using 70 as cutoff for 2-digit year

                    return entryMonth === currentMonth && entryYear === currentYear;
                }
            }
            return false;
        });
    }


    // --- Rendering Functions ---

    // Renders all detailed entries for a given employee in a table
    function renderEmployeeDetailedEntries(entries) {
        if (entries.length === 0) {
            reportDisplay.innerHTML = `<p>No detailed entries found for this employee.</p>`;
            return;
        }

        const employeeName = entries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h3>Detailed Entries for ${employeeName}</h3>`;

        const table = document.createElement('table');
        table.className = 'data-table';

        // Table Header
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const columnsToDisplayInTable = [
            "Timestamp", "Date", "Employee Code", "Designation",
            "Activity Type", "Type of Customer", "Lead Source",
            "How Contacted", "Prospect Name", "Phone Numebr(Whatsapp)", "Address",
            "Profession", "DOB/WD", "Prodcut Interested", "Remarks",
            "Next Follow-up Date", "Relation With Staff"
        ];
        columnsToDisplayInTable.forEach(col => {
            const th = document.createElement('th');
            th.textContent = col;
            headerRow.appendChild(th);
        });

        // Table Body
        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            columnsToDisplayInTable.forEach(col => {
                const cell = row.insertCell();
                cell.textContent = entry[col] || ''; // Display value or empty string
            });
        });
        reportDisplay.appendChild(table);
    }

    // Renders a summary for a given employee (activity counts)
    function renderEmployeeSummary(entries) {
        const employeeName = entries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h3>Activity Summary for ${employeeName}</h3>`;
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

        summaryHtml += `<h4>Activity Types:</h4><ul class="summary-list">`;
        for (const type in activityTypeCounts) {
            summaryHtml += `<li>${type}: ${activityTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Customer Types:</h4><ul class="summary-list">`;
        for (const type in customerTypeCounts) {
            summaryHtml += `<li>${type}: ${customerTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Lead Sources:</h4><ul class="summary-list">`;
        for (const source in leadSourceCounts) {
            summaryHtml += `<li>${source}: ${leadSourceCounts[source]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Products Interested:</h4><ul class="summary-list">`;
        for (const product in productInterestedCounts) {
            summaryHtml += `<li>${product}: ${productInterestedCounts[product]}</li>`;
        }
        summaryHtml += `</ul>`;

        summaryHtml += `<h4>Professions:</h4><ul class="summary-list">`;
        for (const profession in professionCounts) {
            summaryHtml += `<li>${profession}: ${professionCounts[profession]}</li>`;
        }
        summaryHtml += `</ul>`;

        reportDisplay.innerHTML += summaryHtml;
    }

    // New Function: Renders a summary for all staff in a branch
    function renderBranchSummary(branchData) {
        const branchName = branchData[0]['Branch Name'];
        reportDisplay.innerHTML = `<h3>All Staff Activity Summary for ${branchName} Branch</h3>`;

        if (branchData.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this branch.</p>`;
            return;
        }

        const employeeAggregates = {};

        // Aggregate data for each employee in the branch
        branchData.forEach(entry => {
            const employee = entry['Employee Name'];
            const activity = entry['Activity Type'];
            const customerType = entry['Type of Customer'];
            const leadSource = entry['Lead Source'];
            const product = entry['Prodcut Interested'];
            const profession = entry['Profession'];

            if (!employeeAggregates[employee]) {
                employeeAggregates[employee] = {
                    totalEntries: 0,
                    activityTypeCounts: {},
                    customerTypeCounts: {},
                    leadSourceCounts: {},
                    productInterestedCounts: {},
                    professionCounts: {}
                };
            }

            employeeAggregates[employee].totalEntries++;
            if (activity) employeeAggregates[employee].activityTypeCounts[activity] = (employeeAggregates[employee].activityTypeCounts[activity] || 0) + 1;
            if (customerType) employeeAggregates[employee].customerTypeCounts[customerType] = (employeeAggregates[employee].customerTypeCounts[customerType] || 0) + 1;
            if (leadSource) employeeAggregates[employee].leadSourceCounts[leadSource] = (employeeAggregates[employee].leadSourceCounts[leadSource] || 0) + 1;
            if (product) employeeAggregates[employee].productInterestedCounts[product] = (employeeAggregates[employee].productInterestedCounts[product] || 0) + 1;
            if (profession) employeeAggregates[employee].professionCounts[profession] = (employeeAggregates[employee].professionCounts[profession] || 0) + 1;
        });

        // Display summary for each employee in the branch
        for (const employee in employeeAggregates) {
            const empSummary = employeeAggregates[employee];
            let empHtml = `<div class="employee-item">`;
            empHtml += `<h4>${employee} (${empSummary.totalEntries} entries)</h4>`;

            empHtml += `<h5>Activity Types:</h5><ul class="summary-list">`;
            for (const type in empSummary.activityTypeCounts) {
                empHtml += `<li>${type}: ${empSummary.activityTypeCounts[type]}</li>`;
            }
            empHtml += `</ul>`;

            empHtml += `<h5>Products Interested:</h5><ul class="summary-list">`;
            for (const product in empSummary.productInterestedCounts) {
                empHtml += `<li>${product}: ${empSummary.productInterestedCounts[product]}</li>`;
            }
            empHtml += `</ul>`;

            // Add other counts as desired
            empHtml += `</div>`;
            reportDisplay.innerHTML += empHtml;
        }
    }

    // New Function: Renders performance report against targets
    function renderPerformanceReport(entries, employeeName, designation) {
        reportDisplay.innerHTML = `<h3>Performance Report for ${employeeName} (${designation || 'N/A'})</h3>`;
        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries for this employee to calculate performance.</p>`;
            return;
        }

        // Filter entries for the current month for target tracking
        const currentMonthEntries = filterDataForCurrentMonth(entries);
        console.log(`Entries for current month (${new Date().getMonth() + 1}/${new Date().getFullYear()}):`, currentMonthEntries);

        // Calculate actuals for the current month based on clarified mappings
        const actuals = {
            'Visit': 0,
            'Call': 0,
            'Reference': 0, // This is now more specific
            'New Customer Leads': 0
        };

        currentMonthEntries.forEach(entry => {
            // Count Visits
            if (entry['Activity Type'] === 'Visit') {
                actuals['Visit']++;
            }
            // Count Calls
            if (entry['Activity Type'] === 'Call') {
                actuals['Call']++;
            }
            // Count Leads created (Type of Customer: "New")
            if (entry['Type of Customer'] === 'New') {
                actuals['New Customer Leads']++;
            }
            // Count Reference (Type of Customer: "New" AND Activity Type: "Referance")
            if (entry['Type of Customer'] === 'New' && entry['Activity Type'] === 'Referance') {
                actuals['Reference']++;
            }
        });

        console.log("Calculated Actuals for current month:", actuals);

        // Get targets based on designation, default if not found
        const employeeTargets = TARGETS[designation] || TARGETS['Default'];

        // Display results in a table
        let tableHtml = `<table class="performance-table">
            <thead>
                <tr>
                    <th>Metric</th>
                    <th>Actual (This Month)</th>
                    <th>Target (Monthly)</th>
                    <th>Achievement (%)</th>
                    <th>Progress</th>
                </tr>
            </thead>
            <tbody>`;

        for (const metric in employeeTargets) {
            const target = employeeTargets[metric];
            const actual = actuals[metric] || 0; // Use 0 if no actuals counted
            const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';
            const progressBarWidth = target > 0 ? Math.min(100, (actual / target) * 100) : 0; // Cap at 100% for bar

            let progressBarClass = '';
            if (progressBarWidth < 50) {
                progressBarClass = 'danger'; // Red for low progress
            } else if (progressBarWidth < 90) {
                progressBarClass = 'warning'; // Orange for moderate progress
            } else {
                progressBarClass = 'success'; // Green for good progress
            }


            tableHtml += `
                <tr>
                    <td class="performance-metric">${metric}</td>
                    <td>${actual}</td>
                    <td>${target}</td>
                    <td>${percentage}%</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${progressBarClass}" style="width: ${progressBarWidth}%;">
                                ${percentage}%
                            </div>
                        </div>
                    </td>
                </tr>`;
        }

        tableHtml += `</tbody></table>`;
        reportDisplay.innerHTML += tableHtml;
    }


    // --- Main Fetch and Event Listeners ---

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
                displayMessage("Data loaded. Please select a branch to view reports.");
                // Update total submissions on the main page if that element exists
                const totalSubmissionsSpan = document.getElementById('totalSubmissions');
                if (totalSubmissionsSpan) {
                    totalSubmissionsSpan.textContent = allCanvassingData.length;
                }
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
        employeeSelect.innerHTML = `<option value="">-- Select an Employee --</option>`; // Reset employee selection
        employeeFilterPanel.style.display = 'none'; // Hide employee filter panel initially
        viewOptions.style.display = 'none'; // Hide all view options initially
        reportDisplay.innerHTML = ''; // Clear report display

        if (selectedBranch) {
            filteredBranchData = allCanvassingData.filter(entry => entry['Branch Name'] === selectedBranch);

            // Populate Employee dropdown if there are employees in this branch
            populateDropdown(employeeSelect, filteredBranchData, 'Employee Name');
            employeeFilterPanel.style.display = 'block'; // Show employee filter

            // Show view options panel, and the Branch Summary button
            viewOptions.style.display = 'block';
            viewBranchSummaryBtn.style.display = 'inline-block'; // Always show Branch Summary
            // Hide employee-specific buttons until an employee is selected
            viewAllEntriesBtn.style.display = 'none';
            viewEmployeeSummaryBtn.style.display = 'none';
            viewPerformanceReportBtn.style.display = 'none';

            displayMessage(`Branch: ${selectedBranch}. Now select an employee or click "View All Staff Activity".`);

        } else {
            // Reset to initial state
            displayMessage("Please select a branch from the dropdown above to view reports.");
        }
    });

    // Event listener for Employee selection
    employeeSelect.addEventListener('change', () => {
        const selectedEmployee = employeeSelect.value;
        reportDisplay.innerHTML = ''; // Clear report display

        // Show/hide specific view options based on employee selection
        if (selectedEmployee) {
            selectedEmployeeEntries = filteredBranchData.filter(entry => entry['Employee Name'] === selectedEmployee);
            // Show all employee-specific buttons
            viewAllEntriesBtn.style.display = 'inline-block';
            viewEmployeeSummaryBtn.style.display = 'inline-block';
            viewPerformanceReportBtn.style.display = 'inline-block';
            displayMessage(`Employee: ${selectedEmployee}. Choose a report option.`);
        } else {
            // If no employee selected, hide employee-specific buttons
            viewAllEntriesBtn.style.display = 'none';
            viewEmployeeSummaryBtn.style.display = 'none';
            viewPerformanceReportBtn.style.display = 'none';
            displayMessage(`Branch: ${branchSelect.value}. Select an employee or click "View All Staff Activity".`);
        }
    });

    // Event listeners for View Options buttons
    viewBranchSummaryBtn.addEventListener('click', () => {
        if (filteredBranchData.length > 0) {
            renderBranchSummary(filteredBranchData);
        } else {
            displayMessage("No data available for this branch to summarize.");
        }
    });

    viewAllEntriesBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeDetailedEntries(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee to show detailed entries.");
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeSummary(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee to show summary.");
        }
    });

    viewPerformanceReportBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            // Get the employee's designation for target lookup
            const employeeDesignation = selectedEmployeeEntries[0]['Designation'] || 'Default';
            renderPerformanceReport(selectedEmployeeEntries, selectedEmployeeEntries[0]['Employee Name'], employeeDesignation);
        } else {
            displayMessage("No employee selected or no data for this employee for performance report.");
        }
    });


    // Initial data fetch when the page loads
    fetchData();
});