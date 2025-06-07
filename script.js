document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";
    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE
    // NEW: URL for your published MasterEmployees sheet (as CSV)
    const EMPLOYEE_MASTER_DATA_URL = "YOUR_MASTER_EMPLOYEES_PUB_CSV_URL"; // <-- PASTE THE PUBLIC CSV URL FOR YOUR 'MasterEmployees' SHEET HERE

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Default': { // For all other designations not explicitly defined
            'Visit': 5,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        }
    };

    // --- Column Headers Mapping (IMPORTANT: Adjust these to match your actual CSV headers) ---
    // Ensure these match the exact column names in your "Form Responses 2" (or "Canvassing Data") sheet.
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_NO_OF_VISIT = 'No of Visit (Client)';
    const HEADER_NO_OF_CALL = 'No of Call';
    const HEADER_NO_OF_REFERENCE = 'No of Reference';
    const HEADER_NO_OF_NEW_CUSTOMER_LEADS = 'No of New Customer Leads';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer';

    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchSummaryBtn = document.getElementById('viewBranchSummaryBtn');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area
    const reportDisplay = document.getElementById('reportDisplay');

    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data
    let employeeMasterData = []; // Data from Employee Master sheet
    let allUniqueBranches = []; // From MasterEmployees
    let allUniqueEmployees = []; // Employee codes from MasterEmployees
    let employeeCodeToNameMap = {}; // {code: name} from MasterEmployees
    let employeeCodeToDesignationMap = {}; // {code: designation} from MasterEmployees
    let selectedBranchEntries = []; // Activity entries filtered by branch
    let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code

    // Utility to format date to YYYY-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div
    function displayMessage(message, type = 'info', targetDiv = reportDisplay) {
        if (targetDiv) {
            targetDiv.innerHTML = `<div class="message ${type}">${message}</div>`;
            targetDiv.style.display = 'block';
            setTimeout(() => {
                targetDiv.innerHTML = ''; // Clear message
                targetDiv.style.display = 'none';
            }, 5000); // Hide after 5 seconds
        }
    }

    // Specific message display for employee management forms
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessage) {
            employeeManagementMessage.textContent = message;
            employeeManagementMessage.style.color = isError ? 'red' : 'green';
            employeeManagementMessage.style.display = 'block';
            setTimeout(() => {
                employeeManagementMessage.style.display = 'none';
                employeeManagementMessage.textContent = ''; // Clear content
            }, 5000);
        }
    }

    // Function to fetch activity data from Google Sheet (Form Responses 2)
    async function fetchCanvassingData() {
        displayMessage("Fetching activity data...", 'info');
        try {
            const response = await fetch(DATA_URL);
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            displayMessage("Activity data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching canvassing data:', error);
            displayMessage('Failed to load activity data. Please check the DATA_URL and your network connection.', 'error');
            allCanvassingData = [];
        }
    }

    // Function to fetch Employee Master data from its published CSV
    async function fetchEmployeeMasterData() {
        displayMessage("Fetching employee master data...", 'info');
        try {
            const response = await fetch(EMPLOYEE_MASTER_DATA_URL);
            const csvText = await response.text();
            employeeMasterData = parseCSV(csvText); // Parse the MasterEmployees CSV

            // Populate maps and unique lists from the master data
            employeeCodeToNameMap = {};
            employeeCodeToDesignationMap = {};
            employeeMasterData.forEach(entry => {
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                const employeeName = entry[HEADER_EMPLOYEE_NAME];
                const designation = entry[HEADER_DESIGNATION];
                if (employeeCode && employeeName) {
                    employeeCodeToNameMap[employeeCode] = employeeName;
                    employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
                }
            });

            // Populate all unique branches and employees from master data
            allUniqueBranches = [...new Set(employeeMasterData.map(entry => entry[HEADER_BRANCH_NAME]))].sort();
            allUniqueEmployees = Object.keys(employeeCodeToNameMap).sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(branchSelect, allUniqueBranches); // Populate initial branch dropdown with all branches
            displayMessage("Employee master data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching employee master data:', error);
            displayMessage('Failed to load employee master data. Please check EMPLOYEE_MASTER_DATA_URL.', 'error');
            employeeMasterData = [];
            employeeCodeToNameMap = {};
            employeeCodeToDesignationMap = {};
            allUniqueBranches = [];
            allUniqueEmployees = [];
        }
    }


    // CSV parsing function (handles commas within quoted strings)
    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = parseCSVLine(lines[0]); // Headers can also contain commas in quotes
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = parseCSVLine(lines[i]);
            if (values.length !== headers.length) {
                console.warn(`Skipping malformed row ${i + 1}: Expected ${headers.length} columns, got ${values.length}. Line: "${lines[i]}"`);
                continue;
            }
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index];
            });
            data.push(entry);
        }
        return data;
    }

    // Helper to parse a single CSV line safely
    function parseCSVLine(line) {
        const result = [];
        let inQuote = false;
        let currentField = '';
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuote = !inQuote;
            } else if (char === ',' && !inQuote) {
                result.push(currentField.trim());
                currentField = '';
            } else {
                currentField += char;
            }
        }
        result.push(currentField.trim());
        return result;
    }


    // Process fetched data to populate filters and prepare for reports
    async function processData() {
        await fetchEmployeeMasterData(); // Fetch all employees from master sheet first
        await fetchCanvassingData(); // Then fetch activity data

        // Branch dropdown is already populated by fetchEmployeeMasterData
        // Employee dropdown will be populated based on branch selection.
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items, useCodeForValue = false) {
        selectElement.innerHTML = '<option value="">-- Select --</option>'; // Default option
        items.forEach(item => {
            const option = document.createElement('option');
            if (useCodeForValue) {
                option.value = item; // item is employeeCode
                option.textContent = employeeCodeToNameMap[item] || item; // Display name
            } else {
                option.value = item; // item is branch name
                option.textContent = item;
            }
            selectElement.appendChild(option);
        });
    }

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';
            // Filter employee codes based on master data for the selected branch
            const employeeCodesInBranch = employeeMasterData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            // Sort and populate employee dropdown using codes as values and names for display
            const sortedEmployeeCodesInBranch = [...new Set(employeeCodesInBranch)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

            populateDropdown(employeeSelect, sortedEmployeeCodesInBranch, true);
            viewOptions.style.display = 'flex'; // Show view options
            // Reset employee selection and employee-specific display when branch changes
            employeeSelect.value = "";
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';
        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none'; // Hide view options
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown above to view reports.</p>';
            selectedBranchEntries = []; // Clear previous activity filter
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
        }
    });

    // Handle employee selection (now based on employee CODE)
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // Filter activity data by employee code (from allCanvassingData)
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_BRANCH_NAME] === branchSelect.value
            );
            const employeeDisplayName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
            reportDisplay.innerHTML = `<p>Ready to view reports for ${employeeDisplayName}.</p>`;
        } else {
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';
        }
    });

    // Helper to calculate total activity from a set of activity entries
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 };
        entries.forEach(entry => {
            totalActivity['Visit'] += parseInt(entry[HEADER_NO_OF_VISIT]) || 0;
            totalActivity['Call'] += parseInt(entry[HEADER_NO_OF_CALL]) || 0;
            totalActivity['Reference'] += parseInt(entry[HEADER_NO_OF_REFERENCE]) || 0;
            totalActivity['New Customer Leads'] += parseInt(entry[HEADER_NO_OF_NEW_CUSTOMER_LEADS]) || 0;
        });
        return totalActivity;
    }

    // Render All Branch Snapshot (now uses allUniqueBranches from master data)
    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';
        
        allUniqueBranches.forEach(branch => {
            // Get all employees in this branch from master data
            const employeesInBranch = employeeMasterData.filter(emp => emp[HEADER_BRANCH_NAME] === branch);
            const employeeCodesInBranch = employeesInBranch.map(emp => emp[HEADER_EMPLOYEE_CODE]);

            // Filter activity data only for employees in this branch
            const branchActivityEntries = allCanvassingData.filter(entry => employeeCodesInBranch.includes(entry[HEADER_EMPLOYEE_CODE]));
            
            const totalActivity = calculateTotalActivity(branchActivityEntries);
            const displayEmployeeCount = employeeCodesInBranch.length; // Count all employees in master data for this branch

            const branchDiv = document.createElement('div');
            branchDiv.className = 'branch-snapshot';
            branchDiv.innerHTML = `<h3>${branch}</h3>
                                   <p>Total Employees: ${displayEmployeeCount}</p>
                                   <ul class="summary-list">
                                       <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
                                       <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
                                       <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
                                       <li><strong>Total New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                   </ul>`;
            reportDisplay.appendChild(branchDiv);
        });
    }

    // Render All Staff Overall Performance Report (now uses allUniqueEmployees from master data)
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>All Staff Overall Performance Report</h2><div class="branch-performance-grid"></div>';
        const performanceGrid = reportDisplay.querySelector('.branch-performance-grid');

        allUniqueEmployees.forEach(employeeCode => {
            // Get activity data for this specific employee code
            const employeeActivityEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

            const totalActivity = calculateTotalActivity(employeeActivityEntries);
            const employeeName = employeeCodeToNameMap[employeeCode] || 'Unknown Employee';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const employeeCard = document.createElement('div');
            employeeCard.className = 'employee-performance-card';
            employeeCard.innerHTML = `<h3>${employeeName} (${designation})</h3>
                                    <table class="performance-table">
                                        <thead>
                                            <tr><th>Metric</th><th>Actual</th><th>Target</th><th>% Achieved</th></tr>
                                        </thead>
                                        <tbody>
                                            ${Object.keys(targets).map(metric => {
                                                const actualValue = totalActivity[metric] || 0;
                                                const targetValue = targets[metric];
                                                const percentAchieved = performance[metric];
                                                const progressBarClass = getProgressBarClass(percentAchieved);
                                                const displayPercent = percentAchieved !== 'N/A' ? `${percentAchieved}%` : 'N/A';
                                                const progressWidth = percentAchieved !== 'N/A' ? Math.min(100, parseFloat(percentAchieved)) : 0;

                                                return `
                                                    <tr>
                                                        <td>${metric}</td>
                                                        <td>${actualValue}</td>
                                                        <td>${targetValue}</td>
                                                        <td>
                                                            <div class="progress-bar-container-small">
                                                                <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%">${displayPercent}</div>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>`;
            performanceGrid.appendChild(employeeCard);
        });
    }

    // Function to calculate performance percentage
    function calculatePerformance(actual, target) {
        const performance = {};
        for (const key in actual) {
            if (target[key] !== undefined && target[key] > 0) {
                performance[key] = ((actual[key] / target[key]) * 100).toFixed(2);
            } else {
                performance[key] = 'N/A';
            }
        }
        return performance;
    }

    // Helper for progress bar styling
    function getProgressBarClass(percentage) {
        if (percentage === 'N/A') return '';
        const p = parseFloat(percentage);
        if (p >= 100) return 'overachieved';
        if (p >= 75) return 'success';
        if (p >= 50) return 'warning';
        return 'danger';
    }


    // Render Branch Summary (now uses selectedBranchEntries which are activity entries)
    function renderBranchSummary(branchActivityEntries) {
        if (branchActivityEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity data for this branch for the selected period.</p>';
            return;
        }

        const branchName = branchActivityEntries[0][HEADER_BRANCH_NAME];
        reportDisplay.innerHTML = `<h2>Branch Activity Summary for ${branchName}</h2><div class="branch-summary-grid"></div>`;
        const summaryGrid = reportDisplay.querySelector('.branch-summary-grid');

        // Get all unique employee codes within this branch from master data
        const employeesInSelectedBranch = employeeMasterData.filter(emp => emp[HEADER_BRANCH_NAME] === branchName);
        const uniqueEmployeeCodesInBranch = [...new Set(employeesInSelectedBranch.map(e => e[HEADER_EMPLOYEE_CODE]))];

        uniqueEmployeeCodesInBranch.forEach(employeeCode => {
            // Filter activity data for this specific employee code within the selected branch
            const employeeActivities = branchActivityEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const totalActivity = calculateTotalActivity(employeeActivities); // Will be zeros if no activity
            const employeeDisplayName = employeeCodeToNameMap[employeeCode] || employeeCode;

            const employeeSummaryCard = document.createElement('div');
            employeeSummaryCard.className = 'employee-summary-card';
            employeeSummaryCard.innerHTML = `<h3>${employeeDisplayName}</h3>
                                           <ul class="summary-list">
                                               <li><strong>Visits:</strong> ${totalActivity['Visit']}</li>
                                               <li><strong>Calls:</strong> ${totalActivity['Call']}</li>
                                               <li><strong>References:</strong> ${totalActivity['Reference']}</li>
                                               <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                           </ul>`;
            summaryGrid.appendChild(employeeSummaryCard);
        });
    }

    // Render Employee Detailed Entries (uses selectedEmployeeCodeEntries which are activity entries)
    function renderEmployeeDetailedEntries(employeeCodeEntries) {
        if (employeeCodeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No detailed activity entries for this employee code.</p>';
            return;
        }

        const employeeDisplayName = employeeCodeToNameMap[employeeCodeEntries[0][HEADER_EMPLOYEE_CODE]] || employeeCodeEntries[0][HEADER_EMPLOYEE_CODE];
        reportDisplay.innerHTML = `<h2>Detailed Entries for ${employeeDisplayName}</h2>`;

        const table = document.createElement('table');
        table.className = 'data-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headersToShow = Object.keys(employeeCodeEntries[0]).filter(header =>
            header !== HEADER_EMPLOYEE_CODE &&
            header !== HEADER_BRANCH_NAME &&
            header !== HEADER_DESIGNATION
        );

        const preferredOrder = [
            HEADER_TIMESTAMP, HEADER_EMPLOYEE_NAME, 'Type of Customer', HEADER_ACTIVITY_TYPE, 'Activity Remarks',
            HEADER_NO_OF_VISIT, HEADER_NO_OF_CALL, HEADER_NO_OF_REFERENCE, HEADER_NO_OF_NEW_CUSTOMER_LEADS
        ];
        const finalHeaders = preferredOrder.filter(h => headersToShow.includes(h))
            .concat(headersToShow.filter(h => !preferredOrder.includes(h)).sort());


        finalHeaders.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        employeeCodeEntries.forEach(entry => {
            const row = tbody.insertRow();
            finalHeaders.forEach(header => {
                const cell = row.insertCell();
                if (header === HEADER_TIMESTAMP) {
                    cell.textContent = formatDate(entry[header]);
                } else {
                    cell.textContent = entry[header];
                }
            });
        });

        reportDisplay.appendChild(table);
    }

    // Render Employee Summary (Current Month) - uses selectedEmployeeCodeEntries (activity entries)
    function renderEmployeeSummary(employeeCodeEntries) {
        if (employeeCodeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity data for this employee for the selected period.</p>';
            return;
        }

        const employeeDisplayName = employeeCodeToNameMap[employeeCodeEntries[0][HEADER_EMPLOYEE_CODE]] || employeeCodeEntries[0][HEADER_EMPLOYEE_CODE];
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const currentMonthEntries = employeeCodeEntries.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        const totalActivity = calculateTotalActivity(currentMonthEntries);

        reportDisplay.innerHTML = `<h2>Employee Summary for ${employeeDisplayName} (Current Month)</h2>
                                   <ul class="summary-list">
                                       <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
                                       <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
                                       <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
                                       <li><strong>Total New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                   </ul>`;

        if (currentMonthEntries.length > 0) {
            reportDisplay.innerHTML += '<h3>Current Month\'s Detailed Entries</h3>';
            renderEmployeeDetailedEntries(currentMonthEntries);
        } else {
            reportDisplay.innerHTML += '<p>No entries for the current month.</p>';
        }
    }

    // Render Employee Performance Report (uses selectedEmployeeCodeEntries (activity entries))
    function renderPerformanceReport(employeeCodeEntries) {
        if (employeeCodeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity data for this employee for performance report.</p>';
            // Even if no activity data, we can still show a performance report against targets
            // using the employee's designation from the master data.
            const selectedEmployeeCode = employeeSelect.value; // Get the currently selected employee code
            if (selectedEmployeeCode) {
                const employeeName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
                const designation = employeeCodeToDesignationMap[selectedEmployeeCode] || 'Default';
                const targets = TARGETS[designation] || TARGETS['Default'];
                const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // All zeros
                const performance = calculatePerformance(totalActivity, targets); // Calculate performance with zeros

                reportDisplay.innerHTML = `<h2>Performance Report for ${employeeName} (${designation})</h2>
                                        <p>No activity data submitted for this employee, but showing targets.</p>
                                        <table class="performance-table">
                                            <thead>
                                                <tr><th>Metric</th><th>Actual</th><th>Target</th><th>% Achieved</th></tr>
                                            </thead>
                                            <tbody>
                                                ${Object.keys(targets).map(metric => {
                                                    const actualValue = 0;
                                                    const targetValue = targets[metric];
                                                    const percentAchieved = performance[metric];
                                                    const progressBarClass = getProgressBarClass(percentAchieved);
                                                    const displayPercent = percentAchieved !== 'N/A' ? `${percentAchieved}%` : 'N/A';
                                                    const progressWidth = percentAchieved !== 'N/A' ? Math.min(100, parseFloat(percentAchieved)) : 0;

                                                    return `
                                                        <tr>
                                                            <td>${metric}</td>
                                                            <td>${actualValue}</td>
                                                            <td>${targetValue}</td>
                                                            <td>
                                                                <div class="progress-bar-container">
                                                                    <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%">${displayPercent}</div>
                                                                </div>
                                                            </td>
                                                        </tr>
                                                    `;
                                                }).join('')}
                                            </tbody>
                                        </table>`;
            }
            return;
        }

        const employeeCode = employeeCodeEntries[0][HEADER_EMPLOYEE_CODE];
        const employeeDisplayName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        reportDisplay.innerHTML = `<h2>Performance Report for ${employeeDisplayName} (${designation})</h2>`;

        const totalActivity = calculateTotalActivity(employeeCodeEntries);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        const performanceDiv = document.createElement('div');
        performanceDiv.className = 'performance-report';
        performanceDiv.innerHTML = `
            <h3>Overall Performance</h3>
            <table class="performance-table">
                <thead>
                    <tr><th>Metric</th><th>Actual</th><th>Target</th><th>% Achieved</th></tr>
                </thead>
                <tbody>
                    ${Object.keys(targets).map(metric => {
                        const actualValue = totalActivity[metric] || 0;
                        const targetValue = targets[metric];
                        const percentAchieved = performance[metric];
                        const progressBarClass = getProgressBarClass(percentAchieved);
                        const displayPercent = percentAchieved !== 'N/A' ? `${percentAchieved}%` : 'N/A';
                        const progressWidth = percentAchieved !== 'N/A' ? Math.min(100, parseFloat(percentAchieved)) : 0;

                        return `
                            <tr>
                                <td>${metric}</td>
                                <td>${actualValue}</td>
                                <td>${targetValue}</td>
                                <td>
                                    <div class="progress-bar-container">
                                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%">${displayPercent}</div>
                                    </div>
                                </td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
            <div class="summary-details-container">
                <div>
                    <h4>Activity Breakdown by Date</h4>
                    <ul class="summary-list">
                        ${employeeCodeEntries.map(entry => `
                            <li>${formatDate(entry[HEADER_TIMESTAMP])}:
                                V:${entry[HEADER_NO_OF_VISIT] || 0} |
                                C:${entry[HEADER_NO_OF_CALL] || 0} |
                                R:${entry[HEADER_NO_OF_REFERENCE] || 0} |
                                L:${entry[HEADER_NO_OF_NEW_CUSTOMER_LEADS] || 0}
                            </li>`).join('')}
                    </ul>
                </div>
            </div>
        `;
        reportDisplay.appendChild(performanceDiv);
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(actionType, data = {}) {
        displayEmployeeManagementMessage('Processing request...', false);

        try {
            const formData = new URLSearchParams();
            formData.append('actionType', actionType);
            formData.append('data', JSON.stringify(data));

            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: formData
            });

            if (!response.ok) {
                let errorDetails = `HTTP error! status: ${response.status}`;
                try {
                    const errorJson = await response.json();
                    errorDetails += `, message: ${errorJson.message || JSON.stringify(errorJson)}`;
                } catch {
                    const errorText = await response.text();
                    errorDetails += `, message: ${errorText}`;
                }
                throw new Error(errorDetails);
            }

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Error sending data: ${error.message}`, true);
            return false;
        } finally {
            // Re-fetch all data to ensure reports are up-to-date after any employee management action
            // This is crucial to reflect added/deleted employees in dropdowns and reports
            await processData(); // Re-fetch both master and canvassing data
            // Re-render the current report or provide a success message
            // showTab(document.querySelector('.tab-button.active').id); // Re-render active tab content if needed
        }
    }


    // Event Listeners for main report buttons
    viewBranchSummaryBtn.addEventListener('click', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            // Filter activity data for the selected branch
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            renderBranchSummary(branchActivityEntries);
        } else {
            displayMessage("No branch selected.");
        }
    });

    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
             // Filter activity data for the selected branch
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            renderOverallStaffPerformanceReportForBranch(selectedBranch, branchActivityEntries);
        } else {
            displayMessage("No branch selected to show performance report.");
        }
    });

    // Helper for branch performance report (similar to overall staff performance, but for a specific branch)
    function renderOverallStaffPerformanceReportForBranch(branchName, branchActivityEntries) {
        reportDisplay.innerHTML = `<h2>All Staff Performance for ${branchName}</h2><div class="branch-performance-grid"></div>`;
        const performanceGrid = reportDisplay.querySelector('.branch-performance-grid');

        // Get all employees in this branch from master data
        const employeesInSelectedBranch = employeeMasterData.filter(emp => emp[HEADER_BRANCH_NAME] === branchName);
        const uniqueEmployeeCodesInBranch = [...new Set(employeesInSelectedBranch.map(e => e[HEADER_EMPLOYEE_CODE]))];

        uniqueEmployeeCodesInBranch.forEach(employeeCode => {
            // Filter activity data only for this employee code within the branch's activity entries
            const employeeActivities = branchActivityEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            
            const totalActivity = calculateTotalActivity(employeeActivities);
            const employeeDisplayName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const employeeCard = document.createElement('div');
            employeeCard.className = 'employee-performance-card';
            employeeCard.innerHTML = `<h3>${employeeDisplayName} (${designation})</h3>
                                    <table class="performance-table">
                                        <thead>
                                            <tr><th>Metric</th><th>Actual</th><th>Target</th><th>% Achieved</th></tr>
                                        </thead>
                                        <tbody>
                                            ${Object.keys(targets).map(metric => {
                                                const actualValue = totalActivity[metric] || 0;
                                                const targetValue = targets[metric];
                                                const percentAchieved = performance[metric];
                                                const progressBarClass = getProgressBarClass(percentAchieved);
                                                const displayPercent = percentAchieved !== 'N/A' ? `${percentAchieved}%` : 'N/A';
                                                const progressWidth = percentAchieved !== 'N/A' ? Math.min(100, parseFloat(percentAchieved)) : 0;

                                                return `
                                                    <tr>
                                                        <td>${metric}</td>
                                                        <td>${actualValue}</td>
                                                        <td>${targetValue}</td>
                                                        <td>
                                                            <div class="progress-bar-container-small">
                                                                <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%">${displayPercent}</div>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>`;
            performanceGrid.appendChild(employeeCard);
        });
    }

    viewAllEntriesBtn.addEventListener('click', () => {
        if (selectedEmployeeCodeEntries.length > 0) {
            renderEmployeeDetailedEntries(selectedEmployeeCodeEntries);
        } else {
            displayMessage("No employee selected or no activity data for this employee to show detailed entries.");
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        if (selectedEmployeeCodeEntries.length > 0) {
            renderEmployeeSummary(selectedEmployeeCodeEntries);
        } else {
            displayMessage("No employee selected or no activity data for this employee to show summary.");
        }
    });

    viewPerformanceReportBtn.addEventListener('click', () => {
        // This button now fetches the selected employee from the MasterEmployees list
        // and then filters their activities. If no activities, it shows 0s.
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // Find the full employee master data for the selected employee
            const employeeMasterEntry = employeeMasterData.find(emp => emp[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
            if (!employeeMasterEntry) {
                displayMessage("Employee details not found in Master Employees data.");
                return;
            }

            // Filter activity entries for this specific employee
            const employeeActivityEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
            renderPerformanceReport(employeeActivityEntries); // This function handles zero activity
        } else {
            displayMessage("Please select an employee to view performance report.");
        }
    });

    // Function to manage tab visibility
    function showTab(tabButtonId) {
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });
        document.getElementById(tabButtonId).classList.add('active');

        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        if (tabButtonId === 'allBranchSnapshotTabBtn' || tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'flex';
            if (tabButtonId === 'allBranchSnapshotTabBtn') {
                renderAllBranchSnapshot();
            } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
                renderOverallStaffPerformanceReport();
            }
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'none';
            displayEmployeeManagementMessage('', false);
        }
    }


    // Event listeners for main tab buttons
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();

            const employeeName = newEmployeeNameInput.value.trim();
            const employeeCode = newEmployeeCodeInput.value.trim();
            const branchName = newBranchNameInput.value.trim();
            const designation = newDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Please fill in Employee Name, Code, and Branch Name.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_BRANCH_NAME]: branchName,
                [HEADER_DESIGNATION]: designation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset();
                // processData() is now called in finally block of sendDataToGoogleAppsScript
                // showTab() is also in finally block for refresh
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeeLines = bulkDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) {
                    displayEmployeeManagementMessage(`Skipping invalid line: "${line}". Each line must have at least Employee Name and Employee Code.`, true);
                    continue;
                }

                const employeeData = {
                    [HEADER_EMPLOYEE_NAME]: parts[0],
                    [HEADER_EMPLOYEE_CODE]: parts[1],
                    [HEADER_BRANCH_NAME]: branchName,
                    [HEADER_DESIGNATION]: parts[2] || ''
                };
                employeesToAdd.push(employeeData);
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                    // processData() is now called in finally block of sendDataToGoogleAppsScript
                    // showTab() is also in finally block for refresh
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    // Event Listener for Delete Employee Form
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
                // processData() is now called in finally block of sendDataToGoogleAppsScript
                // showTab() is also in finally block for refresh
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
