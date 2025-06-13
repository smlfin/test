document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // This was a placeholder, not actively used for data fetching

    // *** Header Definitions for Canvassing Data (Important for mapping CSV columns) ***
    // Ensure these match your Google Sheet column headers exactly.
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_CUSTOMER_NAME = 'Customer Name';
    const HEADER_PRODUCT_CANVASSED = 'Product Canvassed';
    const HEADER_CUSTOMER_ID = 'Customer ID'; // Assuming Customer ID exists
    const HEADER_PHONE_NUMBER = 'Phone Number'; // Assuming Phone Number exists
    const HEADER_EMAIL_ID = 'Email ID'; // Assuming Email ID exists
    const HEADER_ADDRESS = 'Address'; // Assuming Address exists
    const HEADER_CANVASSING_REMARKS = 'Canvassing Remarks';
    const HEADER_INTEREST_LEVEL = 'Interest Level';
    const HEADER_FOLLOWUP_DATE = 'Follow-up Date';
    const HEADER_DESIGNATION = 'Designation'; // For employee management


    // *** DOM Elements ***
    // Tab buttons
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const followupDueTabBtn = document.getElementById('followupDueTabBtn'); // NEW
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Sections
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const followupDueSection = document.getElementById('followupDueSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');


    // Report section elements
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
    const reportDisplay = document.getElementById('reportDisplay');

    // Customer View section elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');

    // Follow-up Due section elements (NEW)
    const followupBranchSelect = document.getElementById('followupBranchSelect');
    const followupEmployeeSelect = document.getElementById('followupEmployeeSelect');
    const followupDueList = document.getElementById('followupDueList');


    // Employee Management section elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');


    const statusMessage = document.getElementById('statusMessage');


    // *** Data Storage ***
    let canvassingData = []; // All raw data from the main sheet
    let branchesData = []; // Unique branch names
    let employeesData = {}; // Employees grouped by branch
    let customersData = {}; // Customers grouped by employee
    let nonParticipatingBranches = []; // Branches with no entries


    // *** Utility Functions ***
    function showMessage(message, isError = false, type = 'info') {
        statusMessage.textContent = message;
        statusMessage.className = 'message-container message'; // Reset classes
        if (isError) {
            statusMessage.classList.add('error');
        } else {
            statusMessage.classList.add(type);
        }
        statusMessage.style.display = 'block';
        setTimeout(() => {
            statusMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessage.textContent = message;
        employeeManagementMessage.className = 'message'; // Reset classes
        if (isError) {
            employeeManagementMessage.classList.add('error');
        } else {
            employeeManagementMessage.classList.add('success');
        }
        employeeManagementMessage.style.display = 'block';
        setTimeout(() => {
            employeeManagementMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim().replace(/"/g, ''));
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const currentLine = lines[i];
            const values = [];
            let inQuote = false;
            let currentField = '';

            for (let j = 0; j < currentLine.length; j++) {
                const char = currentLine[j];
                if (char === '"') {
                    inQuote = !inQuote;
                } else if (char === ',' && !inQuote) {
                    values.push(currentField.trim().replace(/"/g, ''));
                    currentField = '';
                } else {
                    currentField += char;
                }
            }
            values.push(currentField.trim().replace(/"/g, '')); // Add the last field

            if (values.length === headers.length) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                data.push(row);
            } else {
                console.warn(`Skipping malformed row: ${currentLine} (Expected ${headers.length} columns, got ${values.length})`);
            }
        }
        return data;
    }


    function populateDropdown(selectElement, items, defaultValue = "-- Select --") {
        if (!selectElement) {
            console.error("populateDropdown: selectElement is null. Cannot populate dropdown.");
            return;
        }
        selectElement.innerHTML = `<option value="">${defaultValue}</option>`;
        items.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            selectElement.appendChild(option);
        });
    }

    function calculateTotalActivity(data) {
        const totalActivity = {};
        data.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const branchName = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_PRODUCT_CANVASSED] || 'Unknown';
            const interestLevel = entry[HEADER_INTEREST_LEVEL] || 'Unknown';
            const timestamp = entry[HEADER_TIMESTAMP];

            if (!totalActivity[branchName]) {
                totalActivity[branchName] = {
                    totalCanvassed: 0,
                    employees: {},
                    productBreakdown: {},
                    interestLevelBreakdown: {}
                };
            }

            if (!totalActivity[branchName].employees[employeeCode]) {
                totalActivity[branchName].employees[employeeCode] = {
                    name: employeeName,
                    totalCanvassed: 0,
                    productBreakdown: {},
                    interestLevelBreakdown: {},
                    entries: [] // Store raw entries for all entries view
                };
            }

            totalActivity[branchName].totalCanvassed++;
            totalActivity[branchName].employees[employeeCode].totalCanvassed++;
            totalActivity[branchName].employees[employeeCode].entries.push(entry); // Store the full entry

            const trimmedActivityType = activityType.trim().toLowerCase();
            if (trimmedActivityType) {
                totalActivity[branchName].productBreakdown[trimmedActivityType] = (totalActivity[branchName].productBreakdown[trimmedActivityType] || 0) + 1;
                totalActivity[branchName].employees[employeeCode].productBreakdown[trimmedActivityType] = (totalActivity[branchName].employees[employeeCode].productBreakdown[trimmedActivityType] || 0) + 1;
            }

            const trimmedInterestLevel = interestLevel.trim().toLowerCase();
            if (trimmedInterestLevel) {
                totalActivity[branchName].interestLevelBreakdown[trimmedInterestLevel] = (totalActivity[branchName].interestLevelBreakdown[trimmedInterestLevel] || 0) + 1;
                totalActivity[branchName].employees[employeeCode].interestLevelBreakdown[trimmedInterestLevel] = (totalActivity[branchName].employees[employeeCode].interestLevelBreakdown[trimmedInterestLevel] || 0) + 1;
            }
        });
        return totalActivity;
    }

    function calculateProgressPercentage(current, target) {
        if (target === 0) return 0;
        return Math.min(100, (current / target) * 100);
    }

    function getProgressBarClass(percentage) {
        if (percentage >= 100) return 'success';
        if (percentage >= 75) return 'warning-high';
        if (percentage >= 50) return 'warning-medium';
        if (percentage >= 25) return 'warning-low';
        return 'danger';
    }


    // *** Data Fetching and Processing ***
    async function fetchData() {
        showMessage('Fetching latest data...', false, 'info');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvData = await response.text();
            canvassingData = parseCSV(csvData);
            processData();
            showMessage('Data loaded successfully!', false, 'success');
        } catch (error) {
            console.error('Error fetching data:', error);
            showMessage(`Failed to load data: ${error.message}. Please check the data URL.`, true);
        }
    }

    function processData() {
        branchesData = [];
        employeesData = {};
        customersData = {}; // Clear previous data
        nonParticipatingBranches = [];

        const uniqueBranches = new Set();
        const branchEmployeeMap = {}; // Branch -> Set of employees
        const employeeCustomerMap = {}; // EmployeeCode -> Array of customers

        canvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const customerName = entry[HEADER_CUSTOMER_NAME];
            const customerId = entry[HEADER_CUSTOMER_ID] || 'N/A'; // Default if missing
            const productCanvassed = entry[HEADER_PRODUCT_CANVASSED];

            uniqueBranches.add(branch);

            if (!branchEmployeeMap[branch]) {
                branchEmployeeMap[branch] = new Set();
            }
            branchEmployeeMap[branch].add(`${employeeName} (${employeeCode})`);

            if (!employeeCustomerMap[employeeCode]) {
                employeeCustomerMap[employeeCode] = [];
            }
            employeeCustomerMap[employeeCode].push({
                name: customerName,
                id: customerId,
                product: productCanvassed,
                details: entry // Store the full entry for detailed view
            });
        });

        branchesData = Array.from(uniqueBranches).sort();
        employeesData = {};
        branchesData.forEach(branch => {
            employeesData[branch] = Array.from(branchEmployeeMap[branch] || []).sort();
        });

        customersData = employeeCustomerMap; // Assign the processed customer data

        // Identify non-participating branches
        // This would ideally come from a master list of all branches,
        // but for now, we'll assume branches not in canvassingData are non-participating.
        // If you have a separate master branch list, you'd iterate that.
        // For demonstration:
        const allPossibleBranches = ['Branch A', 'Branch B', 'Branch C', 'Branch D', 'Branch E']; // Example master list
        nonParticipatingBranches = allPossibleBranches.filter(branch => !uniqueBranches.has(branch));


        // Populate initial dropdowns for the default tab
        populateDropdown(branchSelect, branchesData, "-- Select a Branch --");
        populateDropdown(customerViewBranchSelect, branchesData, "-- Select a Branch --");
        populateDropdown(followupBranchSelect, branchesData, "-- All Branches --"); // NEW: For follow-up filter
    }

    async function sendDataToGoogleAppsScript(action, data) {
        showMessage('Sending data...', false, 'info');
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({ action: action, data: JSON.stringify(data) })
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`HTTP error! status: ${response.status}. Message: ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                showMessage(result.message, false, 'success');
                // Re-fetch data if a modification was successful
                await fetchData();
                return true;
            } else {
                showMessage(result.message, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            showMessage(`Failed to send data: ${error.message}. Check your Apps Script URL and deployment.`, true);
            return false;
        }
    }


    // *** Reporting Functions ***

    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';
        const totalActivity = calculateTotalActivity(canvassingData);

        if (Object.keys(totalActivity).length === 0) {
            reportDisplay.innerHTML += '<p>No canvassing activity recorded yet.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table';

        table.innerHTML = `
            <thead>
                <tr>
                    <th>Branch Name</th>
                    <th>Total Canvassed</th>
                    <th>Product Breakdown</th>
                    <th>Interest Level Breakdown</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;
        const tbody = table.querySelector('tbody');

        branchesData.forEach(branchName => { // Use branchesData to include all known branches
            const branchActivity = totalActivity[branchName] || {
                totalCanvassed: 0,
                productBreakdown: {},
                interestLevelBreakdown: {}
            }; // Handle branches with no activity

            const productHtml = Object.entries(branchActivity.productBreakdown)
                .map(([product, count]) => `<li>${product.charAt(0).toUpperCase() + product.slice(1)}: ${count}</li>`)
                .join('');
            const interestHtml = Object.entries(branchActivity.interestLevelBreakdown)
                .map(([level, count]) => `<li>${level.charAt(0).toUpperCase() + level.slice(1)}: ${count}</li>`)
                .join('');

            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Branch Name">${branchName}</td>
                <td data-label="Total Canvassed">${branchActivity.totalCanvassed}</td>
                <td data-label="Product Breakdown"><ul>${productHtml || 'N/A'}</ul></td>
                <td data-label="Interest Level Breakdown"><ul>${interestHtml || 'N/A'}</ul></td>
            `;
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    function renderAllStaffOverallPerformance() {
        reportDisplay.innerHTML = '<h2>All Staff Performance (Overall)</h2>';
        const totalActivity = calculateTotalActivity(canvassingData);

        if (canvassingData.length === 0) {
            reportDisplay.innerHTML += '<p>No canvassing activity recorded yet.</p>';
            return;
        }

        const performanceData = [];
        Object.values(totalActivity).forEach(branch => {
            Object.values(branch.employees).forEach(employee => {
                performanceData.push(employee);
            });
        });

        // Sort by totalCanvassed in descending order
        performanceData.sort((a, b) => b.totalCanvassed - a.totalCanvassed);

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'performance-table';

        table.innerHTML = `
            <thead>
                <tr>
                    <th>Employee Name (Code)</th>
                    <th>Total Canvassed</th>
                    <th>Product Breakdown</th>
                    <th>Interest Level Breakdown</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;
        const tbody = table.querySelector('tbody');

        performanceData.forEach(employee => {
            const productHtml = Object.entries(employee.productBreakdown)
                .map(([product, count]) => `<li>${product.charAt(0).toUpperCase() + product.slice(1)}: ${count}</li>`)
                .join('');
            const interestHtml = Object.entries(employee.interestLevelBreakdown)
                .map(([level, count]) => `<li>${level.charAt(0).toUpperCase() + level.slice(1)}: ${count}</li>`)
                .join('');

            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Employee Name (Code)">${employee.name}</td>
                <td data-label="Total Canvassed">${employee.totalCanvassed}</td>
                <td data-label="Product Breakdown"><ul>${productHtml || 'N/A'}</ul></td>
                <td data-label="Interest Level Breakdown"><ul>${interestHtml || 'N/A'}</ul></td>
            `;
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';

        if (nonParticipatingBranches.length === 0) {
            reportDisplay.innerHTML += '<p>All expected branches have recorded canvassing activity!</p>';
            return;
        }

        const list = document.createElement('ul');
        list.className = 'non-participating-branch-list';
        nonParticipatingBranches.forEach(branch => {
            const li = document.createElement('li');
            li.textContent = branch;
            list.appendChild(li);
        });
        reportDisplay.appendChild(list);
        reportDisplay.innerHTML += '<p class="no-participation-message">These branches have no canvassing entries recorded in the system.</p>';
    }


    function renderBranchPerformanceReport(branchName) {
        reportDisplay.innerHTML = `<h2>Branch Performance Report: ${branchName}</h2>`;
        const totalActivity = calculateTotalActivity(canvassingData);
        const branchActivity = totalActivity[branchName];

        if (!branchActivity || Object.keys(branchActivity.employees).length === 0) {
            reportDisplay.innerHTML += `<p>No canvassing activity found for ${branchName}.</p>`;
            return;
        }

        const grid = document.createElement('div');
        grid.className = 'branch-performance-grid';

        Object.values(branchActivity.employees).forEach(employee => {
            const card = document.createElement('div');
            card.className = 'employee-performance-card';
            // FIX: Simplified this line to resolve SyntaxError: missing } in template string
            // The original complex lookup for branch name caused the error.
            card.innerHTML = `
                <h4>${employee.name}</h4> 
                <table>
                    <tbody>
                        <tr><td>Total Canvassed:</td><td><strong>${employee.totalCanvassed}</strong></td></tr>
                        <tr><td>Products:</td><td><ul>${Object.entries(employee.productBreakdown).map(([p, c]) => `<li>${p.charAt(0).toUpperCase() + p.slice(1)}: ${c}</li>`).join('') || 'N/A'}</ul></td></tr>
                        <tr><td>Interest:</td><td><ul>${Object.entries(employee.interestLevelBreakdown).map(([l, c]) => `<li>${l.charAt(0).toUpperCase() + l.slice(1)}: ${c}</li>`).join('') || 'N/A'}</ul></td></tr>
                    </tbody>
                </table>
            `;
            grid.appendChild(card);
        });
        reportDisplay.appendChild(grid);
    }

    function renderEmployeeSummaryCurrentMonth(employeeCode) {
        reportDisplay.innerHTML = `<h2>Employee Summary: ${employeeCode} (Current Month)</h2>`;
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const employeeEntries = canvassingData.filter(entry =>
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
            new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
        );

        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML += `<p>No canvassing entries found for ${employeeCode} this month.</p>`;
            return;
        }

        const totalCanvassed = employeeEntries.length;
        const productBreakdown = {};
        const interestLevelBreakdown = {};
        const followUpDueCount = 0; // Placeholder for actual calculation if followup due is needed

        employeeEntries.forEach(entry => {
            const product = entry[HEADER_PRODUCT_CANVASSED];
            const interest = entry[HEADER_INTEREST_LEVEL];
            productBreakdown[product] = (productBreakdown[product] || 0) + 1;
            interestLevelBreakdown[interest] = (interestLevelBreakdown[interest] || 0) + 1;
            // Add logic here to check if follow-up date is within next 7 days or past due
        });

        const summaryBreakdown = document.createElement('div');
        summaryBreakdown.className = 'summary-breakdown-card';

        summaryBreakdown.innerHTML = `
            <div>
                <h4>Overall Summary</h4>
                <ul class="summary-list">
                    <li>Total Canvassed: <strong>${totalCanvassed}</strong></li>
                    <li>Follow-ups Due: <strong>${followUpDueCount}</strong></li>
                    </ul>
            </div>
            <div>
                <h4>Product Interest Breakdown</h4>
                <ul class="product-interest-list">
                    ${Object.entries(productBreakdown).map(([product, count]) => `<li>${product}: <span>${count}</span></li>`).join('')}
                </ul>
            </div>
            <div>
                <h4>Interest Level Distribution</h4>
                <ul class="product-interest-list">
                    ${Object.entries(interestLevelBreakdown).map(([level, count]) => `<li>${level}: <span>${count}</span></li>`).join('')}
                </ul>
            </div>
        `;
        reportDisplay.appendChild(summaryBreakdown);
    }


    function renderAllEntries(employeeCode) {
        reportDisplay.innerHTML = `<h2>All Entries for Employee: ${employeeCode}</h2>`;
        const employeeEntries = canvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for employee ${employeeCode}.</p>`;
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');

        table.innerHTML = `
            <thead>
                <tr>
                    <th>Timestamp</th>
                    <th>Customer Name</th>
                    <th>Product Canvassed</th>
                    <th>Interest Level</th>
                    <th>Remarks</th>
                    <th>Follow-up Date</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;
        const tbody = table.querySelector('tbody');

        employeeEntries.forEach(entry => {
            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Timestamp">${entry[HEADER_TIMESTAMP] || 'N/A'}</td>
                <td data-label="Customer Name">${entry[HEADER_CUSTOMER_NAME] || 'N/A'}</td>
                <td data-label="Product Canvassed">${entry[HEADER_PRODUCT_CANVASSED] || 'N/A'}</td>
                <td data-label="Interest Level">${entry[HEADER_INTEREST_LEVEL] || 'N/A'}</td>
                <td data-label="Remarks">${entry[HEADER_CANVASSING_REMARKS] || 'N/A'}</td>
                <td data-label="Follow-up Date">${entry[HEADER_FOLLOWUP_DATE] || 'N/A'}</td>
            `;
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    function renderEmployeePerformanceReport(employeeCode) {
        reportDisplay.innerHTML = `<h2>Performance Report for Employee: ${employeeCode}</h2>`;
        const totalActivity = calculateTotalActivity(canvassingData);
        let employeePerformance = null;

        // Find the employee's performance data across all branches
        for (const branchKey in totalActivity) {
            if (totalActivity[branchKey].employees[employeeCode]) {
                employeePerformance = totalActivity[branchKey].employees[employeeCode];
                break;
            }
        }

        if (!employeePerformance) {
            reportDisplay.innerHTML += `<p>No performance data found for employee ${employeeCode}.</p>`;
            return;
        }

        const totalCanvassed = employeePerformance.totalCanvassed;
        const productBreakdown = employeePerformance.productBreakdown;
        const interestLevelBreakdown = employeePerformance.interestLevelBreakdown;

        const reportContent = `
            <div class="summary-breakdown-card">
                <div>
                    <h4>Overall Activity</h4>
                    <ul class="summary-list">
                        <li>Total Canvassed: <strong>${totalCanvassed}</strong></li>
                        </ul>
                </div>
                <div>
                    <h4>Product Breakdown</h4>
                    <ul class="product-interest-list">
                        ${Object.entries(productBreakdown).map(([product, count]) => `<li>${product.charAt(0).toUpperCase() + product.slice(1)}: <span>${count}</span></li>`).join('') || '<li>N/A</li>'}
                    </ul>
                </div>
                <div>
                    <h4>Interest Levels</h4>
                    <ul class="product-interest-list">
                        ${Object.entries(interestLevelBreakdown).map(([level, count]) => `<li>${level.charAt(0).toUpperCase() + level.slice(1)}: <span>${count}</span></li>`).join('') || '<li>N/A</li>'}
                    </ul>
                </div>
            </div>
        `;
        reportDisplay.innerHTML += reportContent;
    }


    // *** Customer View Functions ***
    function loadCustomersForView(branch, employee) {
        customerCanvassedList.innerHTML = ''; // Clear previous list
        customerDetailsContent.innerHTML = '<p>Select a customer from the list to view their details.</p>'; // Clear details

        let filteredCustomers = [];
        if (employee && customersData[employee.split('(')[1].slice(0, -1)]) { // Extract employee code from "Name (Code)"
            filteredCustomers = customersData[employee.split('(')[1].slice(0, -1)];
        } else if (branch) {
            // If only branch is selected, filter all customers by branch
            Object.values(customersData).forEach(empCustomers => {
                empCustomers.forEach(customer => {
                    if (customer.details[HEADER_BRANCH_NAME] === branch) {
                        filteredCustomers.push(customer);
                    }
                });
            });
        }

        if (filteredCustomers.length === 0) {
            customerCanvassedList.innerHTML = '<p>No customers found for this selection.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';
        filteredCustomers.forEach(customer => {
            const li = document.createElement('li');
            li.className = 'customer-list-item';
            li.dataset.customerId = customer.id; // Store ID for easy lookup
            li.textContent = `${customer.name} - ${customer.product}`;
            li.addEventListener('click', () => {
                // Remove active class from all items
                document.querySelectorAll('.customer-list-item').forEach(item => {
                    item.classList.remove('active');
                });
                // Add active class to clicked item
                li.classList.add('active');
                renderCustomerDetails(customer.details); // Pass the full details object
            });
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function renderCustomerDetails(customer) {
        if (!customer) {
            customerDetailsContent.innerHTML = '<p>No customer selected or details unavailable.</p>';
            return;
        }

        // Build the HTML structure for sections
        let detailsHtml = '';

        // Personal Information Section
        detailsHtml += `
            <div class="customer-info-section">
                <h3>Personal Information</h3>
                <div class="detail-row"><span class="detail-label">Customer Name:</span><span class="detail-value">${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Customer ID:</span><span class="detail-value">${customer[HEADER_CUSTOMER_ID] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Date Canvassed:</span><span class="detail-value">${customer[HEADER_TIMESTAMP] ? new Date(customer[HEADER_TIMESTAMP]).toLocaleDateString() : 'N/A'}</span></div>
            </div>
        `;

        // Contact Details Section
        detailsHtml += `
            <div class="customer-info-section">
                <h3>Contact Details</h3>
                <div class="detail-row"><span class="detail-label">Phone:</span><span class="detail-value">${customer[HEADER_PHONE_NUMBER] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Email:</span><span class="detail-value">${customer[HEADER_EMAIL_ID] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Address:</span><span class="detail-value">${customer[HEADER_ADDRESS] || 'N/A'}</span></div>
            </div>
        `;

        // Canvassing Details Section
        detailsHtml += `
            <div class="customer-info-section">
                <h3>Canvassing Details</h3>
                <div class="detail-row"><span class="detail-label">Canvassed By:</span><span class="detail-value">${customer[HEADER_EMPLOYEE_NAME] || 'N/A'} (${customer[HEADER_EMPLOYEE_CODE] || 'N/A'})</span></div>
                <div class="detail-row"><span class="detail-label">Branch:</span><span class="detail-value">${customer[HEADER_BRANCH_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Assigned Product:</span><span class="detail-value">${customer[HEADER_PRODUCT_CANVASSED] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Interest Level:</span><span class="detail-value">${customer[HEADER_INTEREST_LEVEL] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Follow-up Date:</span><span class="detail-value">${customer[HEADER_FOLLOWUP_DATE] || 'N/A'}</span></div>
            </div>
        `;

        // Remarks Section (full-width)
        if (customer[HEADER_CANVASSING_REMARKS]) {
            detailsHtml += `
                <div class="customer-info-section full-width-section">
                    <h3>Remarks</h3>
                    <p class="remark-text">${customer[HEADER_CANVASSING_REMARKS]}</p>
                </div>
            `;
        }

        customerDetailsContent.innerHTML = detailsHtml;
    }


    // *** Follow-up Due Functions (NEW) ***
    function renderFollowupDueCustomers(branch = null, employee = null) {
        followupDueList.innerHTML = '<h2>Customers Due for Follow-up</h2>';
        const today = new Date();
        // Set today to start of day for comparison
        today.setHours(0, 0, 0, 0);

        let filteredFollowups = canvassingData.filter(entry => {
            const followupDateStr = entry[HEADER_FOLLOWUP_DATE];
            if (!followupDateStr) return false;

            const followupDate = new Date(followupDateStr);
            followupDate.setHours(0, 0, 0, 0); // Normalize to start of day

            // Filter for dates up to next 7 days (inclusive) or past due
            const sevenDaysFromNow = new Date(today);
            sevenDaysFromNow.setDate(today.getDate() + 7);
            sevenDaysFromNow.setHours(23, 59, 59, 999); // End of day 7 days from now

            const isDue = (followupDate >= today && followupDate <= sevenDaysFromNow) || (followupDate < today);

            // Apply branch and employee filters if selected
            let passesBranchFilter = true;
            if (branch && branch !== "") {
                passesBranchFilter = entry[HEADER_BRANCH_NAME] === branch;
            }

            let passesEmployeeFilter = true;
            if (employee && employee !== "") {
                passesEmployeeFilter = entry[HEADER_EMPLOYEE_CODE] === employee.split('(')[1].slice(0, -1); // Extract code
            }

            return isDue && passesBranchFilter && passesEmployeeFilter;
        });

        if (filteredFollowups.length === 0) {
            followupDueList.innerHTML += '<p>No customers are due for follow-up in the selected criteria.</p>';
            return;
        }

        // Sort by follow-up date, oldest first (past due then upcoming)
        filteredFollowups.sort((a, b) => {
            const dateA = new Date(a[HEADER_FOLLOWUP_DATE]);
            const dateB = new Date(b[HEADER_FOLLOWUP_DATE]);
            return dateA.getTime() - dateB.getTime();
        });


        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Customer Name</th>
                    <th>Product</th>
                    <th>Interest Level</th>
                    <th>Follow-up Date</th>
                    <th>Canvassed By</th>
                    <th>Branch</th>
                    <th>Remarks</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;
        const tbody = table.querySelector('tbody');

        filteredFollowups.forEach(customer => {
            const row = tbody.insertRow();
            let followupDateStatus = customer[HEADER_FOLLOWUP_DATE] || 'N/A';
            if (customer[HEADER_FOLLOWUP_DATE]) {
                const fDate = new Date(customer[HEADER_FOLLOWUP_DATE]);
                fDate.setHours(0, 0, 0, 0);
                if (fDate < today) {
                    followupDateStatus = `<span style="color: red; font-weight: bold;">${followupDateStatus} (Overdue)</span>`;
                } else if (fDate.toDateString() === today.toDateString()) {
                     followupDateStatus = `<span style="color: orange; font-weight: bold;">${followupDateStatus} (Today)</span>`;
                }
            }


            row.innerHTML = `
                <td data-label="Customer Name">${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</td>
                <td data-label="Product">${customer[HEADER_PRODUCT_CANVASSED] || 'N/A'}</td>
                <td data-label="Interest Level">${customer[HEADER_INTEREST_LEVEL] || 'N/A'}</td>
                <td data-label="Follow-up Date">${followupDateStatus}</td>
                <td data-label="Canvassed By">${customer[HEADER_EMPLOYEE_NAME] || 'N/A'} (${customer[HEADER_EMPLOYEE_CODE] || 'N/A'})</td>
                <td data-label="Branch">${customer[HEADER_BRANCH_NAME] || 'N/A'}</td>
                <td data-label="Remarks">${customer[HEADER_CANVASSING_REMARKS] || 'N/A'}</td>
            `;
        });
        tableContainer.appendChild(table);
        followupDueList.appendChild(tableContainer);
    }


    // *** Tab Management ***
    let currentActiveTab = 'allBranchSnapshotTabBtn'; // Default active tab

    function showTab(tabId) {
        // Hide all sections
        const allSections = document.querySelectorAll('.report-section, .employee-management-section');
        allSections.forEach(section => {
            section.style.display = 'none';
        });

        // Deactivate all tab buttons
        const allTabButtons = document.querySelectorAll('.tab-button');
        allTabButtons.forEach(button => {
            button.classList.remove('active');
        });

        // Activate the selected tab button
        const selectedTabButton = document.getElementById(tabId);
        if (selectedTabButton) {
            selectedTabButton.classList.add('active');
        }

        // Show the corresponding section and load data
        if (tabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
            reportDisplay.innerHTML = '<p>Loading All Branch Snapshot...</p>'; // Loading message
            renderAllBranchSnapshot();
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
            reportDisplay.innerHTML = '<p>Loading All Staff Performance...</p>'; // Loading message
            renderAllStaffOverallPerformance();
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
            reportDisplay.innerHTML = '<p>Loading Non-Participating Branches...</p>'; // Loading message
            renderNonParticipatingBranches();
        } else if (tabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'flex';
            // Ensure dropdowns are populated and customer list is ready
            populateDropdown(customerViewBranchSelect, branchesData, "-- Select a Branch --");
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>'; // Reset
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
            customerDetailsContent.innerHTML = '<p>Select a customer from the list to view their details.</p>';
        } else if (tabId === 'followupDueTabBtn') { // NEW TAB LOGIC
            if (followupDueSection) {
                followupDueSection.style.display = 'block';
                // Ensure dropdowns are populated for filtering follow-up customers
                populateDropdown(followupBranchSelect, branchesData, "-- All Branches --");
                followupEmployeeSelect.innerHTML = '<option value="">-- All Employees --</option>'; // Reset
                renderFollowupDueCustomers(); // Initial render without filters
            }
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            employeeManagementMessage.style.display = 'none'; // Clear any previous messages
        }
    }


    // *** Event Listeners ***

    // Tab button click handlers
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    followupDueTabBtn.addEventListener('click', () => showTab('followupDueTabBtn')); // NEW
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // Report Section Dropdown & Button Listeners
    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            if (selectedBranch) {
                employeeFilterPanel.style.display = 'block';
                viewOptions.style.display = 'flex';
                populateDropdown(employeeSelect, employeesData[selectedBranch] || [], "-- Select an Employee --");
            } else {
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                reportDisplay.innerHTML = '<p>Select a branch and/or employee to view reports.</p>';
            }
        });
    }

    if (viewBranchPerformanceReportBtn) {
        viewBranchPerformanceReportBtn.addEventListener('click', () => {
            const selectedBranch = branchSelect.value;
            if (selectedBranch) {
                renderBranchPerformanceReport(selectedBranch);
            } else {
                showMessage('Please select a branch to view the report.', true);
            }
        });
    }

    if (viewEmployeeSummaryBtn) {
        viewEmployeeSummaryBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value ? employeeSelect.value.split('(')[1].slice(0, -1) : null;
            if (selectedEmployeeCode) {
                renderEmployeeSummaryCurrentMonth(selectedEmployeeCode);
            } else {
                showMessage('Please select an employee to view their summary.', true);
            }
        });
    }

    if (viewAllEntriesBtn) {
        viewAllEntriesBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value ? employeeSelect.value.split('(')[1].slice(0, -1) : null;
            if (selectedEmployeeCode) {
                renderAllEntries(selectedEmployeeCode);
            } else {
                showMessage('Please select an employee to view all entries.', true);
            }
        });
    }

    if (viewPerformanceReportBtn) {
        viewPerformanceReportBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value ? employeeSelect.value.split('(')[1].slice(0, -1) : null;
            if (selectedEmployeeCode) {
                renderEmployeePerformanceReport(selectedEmployeeCode);
            } else {
                showMessage('Please select an employee to view their performance report.', true);
            }
        });
    }


    // Customer View Dropdown Listeners
    if (customerViewBranchSelect) {
        customerViewBranchSelect.addEventListener('change', () => {
            const selectedBranch = customerViewBranchSelect.value;
            if (selectedBranch) {
                // Populate employee dropdown based on selected branch
                populateDropdown(customerViewEmployeeSelect, employeesData[selectedBranch] || [], "-- Select an Employee --");
            } else {
                customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>'; // Reset employee dropdown
            }
            // Load customers based on current branch and employee selection
            loadCustomersForView(selectedBranch, customerViewEmployeeSelect.value);
        });
    }

    if (customerViewEmployeeSelect) {
        customerViewEmployeeSelect.addEventListener('change', () => {
            const selectedBranch = customerViewBranchSelect.value;
            const selectedEmployee = customerViewEmployeeSelect.value;
            loadCustomersForView(selectedBranch, selectedEmployee);
        });
    }

    // Follow-up Due Dropdown Listeners (NEW)
    if (followupBranchSelect) {
        followupBranchSelect.addEventListener('change', () => {
            const selectedBranch = followupBranchSelect.value;
            if (selectedBranch) {
                populateDropdown(followupEmployeeSelect, employeesData[selectedBranch] || [], "-- All Employees --");
            } else {
                followupEmployeeSelect.innerHTML = '<option value="">-- All Employees --</option>'; // Reset
            }
            renderFollowupDueCustomers(selectedBranch, followupEmployeeSelect.value);
        });
    }

    if (followupEmployeeSelect) {
        followupEmployeeSelect.addEventListener('change', () => {
            const selectedBranch = followupBranchSelect.value;
            const selectedEmployee = followupEmployeeSelect.value;
            renderFollowupDueCustomers(selectedBranch, selectedEmployee);
        });
    }


    // Employee Management Form Listeners
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH_NAME]) {
                displayEmployeeManagementMessage('Employee Name, Code, and Branch are required.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const detailsText = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !detailsText) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = detailsText.split('\n').filter(line => line.trim() !== '');

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Expect at least Name and Code
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || ''
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    fetchData(); // Fetch data first
    showTab(currentActiveTab); // Then show the default tab
});
