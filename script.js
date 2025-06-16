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
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // This is just a placeholder, not used in frontend logic currently


    // *** Header Constants (MUST match your Google Sheet headers exactly) ***
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_CUSTOMER_NAME = 'Customer Name';
    const HEADER_CUSTOMER_CONTACT = 'Customer Contact';
    const HEADER_ADDRESS = 'Address';
    const HEADER_ACTIVITY_TYPE = 'Activity Type'; // e.g., 'Visit', 'Call', 'Other'
    const HEADER_REMARKS = 'Remarks';
    const HEADER_STATUS = 'Status'; // e.g., 'Completed', 'Pending', 'Follow-up'

    // *** Predefined list of branches for reporting and filtering ***
    // IMPORTANT: Update this array with the exact names of your branches
    const PREDEFINED_BRANCHES = [
        "Main Branch",
        "East Branch",
        "West Branch",
        "North Branch",
        "South Branch"
        // Add all your actual branch names here
    ];

    // *** DOM Elements ***
    const statusMessageDiv = document.getElementById('statusMessage');
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const reportDisplay = document.getElementById('reportDisplay');
    const branchNameInput = document.getElementById('branchNameInput');
    const employeeNameInput = document.getElementById('employeeNameInput');
    const employeeCodeInput = document.getElementById('employeeCodeInput');
    const designationInput = document.getElementById('designationInput');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const updateEmployeeForm = document.getElementById('updateEmployeeForm');
    const updateEmployeeCodeInput = document.getElementById('updateEmployeeCodeInput');
    const updateEmployeeNameInput = document.getElementById('updateEmployeeNameInput');
    const updateDesignationInput = document.getElementById('updateDesignationInput');
    const updateBranchNameInput = document.getElementById('updateBranchNameInput');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');

    // Tab buttons and sections
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const branchPerformanceTabBtn = document.getElementById('branchPerformanceTabBtn'); // NEW

    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');


    // *** Global Data Store ***
    let allCanvassingData = [];
    let allEmployeeData = []; // Will store employee data for dropdowns and management


    // *** Helper Functions ***

    function displayStatusMessage(message, isError) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = 'message-container ' + (isError ? 'error-message' : 'success-message');
        statusMessageDiv.style.display = 'block';

        // Hide after 5 seconds
        setTimeout(() => {
            statusMessageDiv.style.display = 'none';
        }, 5000);
    }

    function displayEmployeeManagementMessage(message, isError) {
        const messageDiv = document.getElementById('employeeManagementMessage'); // Assuming you have this div
        if (messageDiv) {
            messageDiv.textContent = message;
            messageDiv.className = 'message-container ' + (isError ? 'error-message' : 'success-message');
            messageDiv.style.display = 'block';
            setTimeout(() => {
                messageDiv.style.display = 'none';
            }, 5000);
        } else {
            displayStatusMessage(message, isError); // Fallback to general status message
        }
    }


    async function sendDataToGoogleAppsScript(action, data) {
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                })
            });

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayStatusMessage(result.message || 'Operation successful!', false);
                // Re-fetch data if operation might affect displayed reports
                if (action === 'add_employee' || action === 'update_employee' || action === 'delete_employee' || action === 'add_bulk_employees') {
                    await processData();
                }
                return true;
            } else {
                displayStatusMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Google Apps Script:', error);
            displayStatusMessage('Error communicating with server: ' + error.message, true);
            return false;
        }
    }


    // *** Data Fetching and Processing ***

    async function fetchData(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            return Papa.parse(csvText, { header: true, skipEmptyLines: true }).data;
        } catch (error) {
            console.error('Error fetching data:', error);
            displayStatusMessage('Failed to fetch data: ' + error.message, true);
            return [];
        }
    }

    async function processData() {
        allCanvassingData = await fetchData(DATA_URL);
        // Extract unique employees for dropdowns (assuming employee data is in canvassing data)
        const uniqueEmployees = new Map();
        allCanvassingData.forEach(row => {
            if (row[HEADER_EMPLOYEE_CODE] && row[HEADER_EMPLOYEE_NAME]) {
                uniqueEmployees.set(row[HEADER_EMPLOYEE_CODE], {
                    [HEADER_EMPLOYEE_CODE]: row[HEADER_EMPLOYEE_CODE],
                    [HEADER_EMPLOYEE_NAME]: row[HEADER_EMPLOYEE_NAME],
                    [HEADER_BRANCH_NAME]: row[HEADER_BRANCH_NAME] || 'N/A',
                    [HEADER_DESIGNATION]: row[HEADER_DESIGNATION] || 'N/A'
                });
            }
        });
        allEmployeeData = Array.from(uniqueEmployees.values());

        populateBranchDropdown();
        populateEmployeeDropdown(); // Populate after data is fetched

        // Re-render current active tab to update content with new data
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id);
        }
    }

    function populateBranchDropdown() {
        if (!branchSelect) return;
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        PREDEFINED_BRANCHES.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    function populateEmployeeDropdown(branchName = '') {
        if (!employeeSelect) return;
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';

        const filteredEmployees = branchName
            ? allEmployeeData.filter(emp => emp[HEADER_BRANCH_NAME] === branchName)
            : allEmployeeData;

        filteredEmployees.forEach(employee => {
            const option = document.createElement('option');
            option.value = employee[HEADER_EMPLOYEE_CODE];
            option.textContent = `${employee[HEADER_EMPLOYEE_NAME]} (${employee[HEADER_EMPLOYEE_CODE]}) - ${employee[HEADER_DESIGNATION]} (${employee[HEADER_BRANCH_NAME]})`;
            employeeSelect.appendChild(option);
        });
    }

    // *** Report Rendering Functions ***

    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';

        // Calculate total activities per branch
        const branchActivity = {};
        allCanvassingData.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            if (branchName) {
                if (!branchActivity[branchName]) {
                    branchActivity[branchName] = { 'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0 };
                }
                const type = activityType ? activityType.trim() : 'Other';
                branchActivity[branchName][type]++;
                branchActivity[branchName]['Total']++;
            }
        });

        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table'; // Apply class for styling
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Branch Name</th>
                    <th>Total Visits</th>
                    <th>Total Calls</th>
                    <th>Total Other Activities</th>
                    <th>Overall Total</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        const tbody = table.querySelector('tbody');

        PREDEFINED_BRANCHES.forEach(branch => { // Use PREDEFINED_BRANCHES for consistent display
            const data = branchActivity[branch] || { 'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0 };
            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Branch Name">${branch}</td>
                <td data-label="Total Visits">${data['Visit']}</td>
                <td data-label="Total Calls">${data['Call']}</td>
                <td data-label="Total Other Activities">${data['Other']}</td>
                <td data-label="Overall Total">${data['Total']}</td>
            `;
        });

        reportDisplay.appendChild(table);
    }


    function renderAllStaffOverallPerformance() {
        reportDisplay.innerHTML = '<h2>All Staff Overall Performance</h2>';

        // Group activities by employee code
        const employeeActivity = {};
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const remarks = entry[HEADER_REMARKS];
            const customerName = entry[HEADER_CUSTOMER_NAME];
            const branchName = entry[HEADER_BRANCH_NAME];
            const date = entry[HEADER_DATE];

            if (employeeCode) {
                if (!employeeActivity[employeeCode]) {
                    employeeActivity[employeeCode] = {
                        'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0,
                        'name': entry[HEADER_EMPLOYEE_NAME] || employeeCode,
                        'designation': entry[HEADER_DESIGNATION] || 'N/A',
                        'branch': branchName || 'N/A',
                        'latestActivity': {}
                    };
                }
                const type = activityType ? activityType.trim() : 'Other';
                employeeActivity[employeeCode][type]++;
                employeeActivity[employeeCode]['Total']++;

                // Track latest activity for each employee
                const currentActivityDate = new Date(date);
                const latestDate = new Date(employeeActivity[employeeCode].latestActivity.date || 0);

                if (currentActivityDate > latestDate) {
                    employeeActivity[employeeCode].latestActivity = {
                        date: date,
                        type: type,
                        customer: customerName,
                        remarks: remarks
                    };
                }
            }
        });

        const table = document.createElement('table');
        table.className = 'all-staff-performance-table'; // Apply class for styling
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Employee Name (Code)</th>
                    <th>Branch</th>
                    <th>Designation</th>
                    <th>Total Visits</th>
                    <th>Total Calls</th>
                    <th>Total Other Activities</th>
                    <th>Overall Total</th>
                    <th>Latest Activity Date</th>
                    <th>Latest Activity Type</th>
                    <th>Latest Customer</th>
                    <th>Latest Remarks</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        const tbody = table.querySelector('tbody');

        // Sort employees by overall total activity in descending order
        const sortedEmployees = Object.keys(employeeActivity).sort((a, b) => {
            return employeeActivity[b]['Total'] - employeeActivity[a]['Total'];
        });

        sortedEmployees.forEach(employeeCode => {
            const data = employeeActivity[employeeCode];
            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Employee Name (Code)">${data.name} (${employeeCode})</td>
                <td data-label="Branch">${data.branch}</td>
                <td data-label="Designation">${data.designation}</td>
                <td data-label="Total Visits">${data['Visit']}</td>
                <td data-label="Total Calls">${data['Call']}</td>
                <td data-label="Total Other Activities">${data['Other']}</td>
                <td data-label="Overall Total">${data['Total']}</td>
                <td data-label="Latest Activity Date">${data.latestActivity.date || 'N/A'}</td>
                <td data-label="Latest Activity Type">${data.latestActivity.type || 'N/A'}</td>
                <td data-label="Latest Customer">${data.latestActivity.customer || 'N/A'}</td>
                <td data-label="Latest Remarks">${data.latestActivity.remarks || 'N/A'}</td>
            `;
        });
        reportDisplay.appendChild(table);
    }

    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches (No Activity)</h2>';

        const activeBranches = new Set();
        allCanvassingData.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                activeBranches.add(entry[HEADER_BRANCH_NAME]);
            }
        });

        const nonParticipating = PREDEFINED_BRANCHES.filter(branch => !activeBranches.has(branch));

        if (nonParticipating.length > 0) {
            const list = document.createElement('ul');
            list.className = 'non-participating-list';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                list.appendChild(li);
            });
            reportDisplay.appendChild(list);
        } else {
            reportDisplay.innerHTML += '<p class="success-message">All predefined branches have recorded activity!</p>';
        }
    }

    function renderCustomerCanvassedList(branchFilter = '', employeeFilter = '') {
        customerCanvassedList.innerHTML = ''; // Clear previous list

        let filteredData = allCanvassingData;

        if (branchFilter) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === branchFilter);
        }

        if (employeeFilter) {
            filteredData = filteredData.filter(row => row[HEADER_EMPLOYEE_CODE] === employeeFilter);
        }

        if (filteredData.length === 0) {
            customerCanvassedList.innerHTML = '<p class="info-message">No customer canvassed data available for the selected filters.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';

        // Sort by Date (most recent first)
        filteredData.sort((a, b) => new Date(b[HEADER_DATE]) - new Date(a[HEADER_DATE]));

        filteredData.forEach(customer => {
            const li = document.createElement('li');
            li.className = 'customer-list-item';
            li.innerHTML = `
                <div class="customer-name">${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</div>
                <div class="customer-meta">
                    <span>${customer[HEADER_BRANCH_NAME] || 'N/A'}</span> -
                    <span>${customer[HEADER_EMPLOYEE_NAME] || 'N/A'}</span> -
                    <span>${customer[HEADER_DATE] || 'N/A'}</span>
                </div>
            `;
            li.addEventListener('click', () => displayCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function displayCustomerDetails(customer) {
        customerDetailsContent.innerHTML = `
            <h3>Customer Details: ${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</h3>
            <div class="detail-grid">
                <div class="detail-row"><span class="detail-label">Customer Contact:</span> <span class="detail-value">${customer[HEADER_CUSTOMER_CONTACT] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Address:</span> <span class="detail-value">${customer[HEADER_ADDRESS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Branch:</span> <span class="detail-value">${customer[HEADER_BRANCH_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Employee:</span> <span class="detail-value">${customer[HEADER_EMPLOYEE_NAME] || 'N/A'} (${customer[HEADER_EMPLOYEE_CODE] || 'N/A'})</span></div>
                <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Date:</span> <span class="detail-value">${customer[HEADER_DATE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Time:</span> <span class="detail-value">${customer[HEADER_TIMESTAMP] ? new Date(customer[HEADER_TIMESTAMP]).toLocaleTimeString() : 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${customer[HEADER_REMARKS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Status:</span> <span class="detail-value">${customer[HEADER_STATUS] || 'N/A'}</span></div>
            </div>
        `;
    }

    // Helper to calculate activity counts per branch
    function getBranchActivitySummary() {
        const branchSummary = {};

        allCanvassingData.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE];

            if (!branchName) return; // Skip entries without a branch name

            if (!branchSummary[branchName]) {
                branchSummary[branchName] = { 'Visit': 0, 'Call': 0 };
            }

            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';

            if (trimmedActivityType === 'visit') {
                branchSummary[branchName]['Visit']++;
            } else if (trimmedActivityType === 'calls') {
                branchSummary[branchName]['Call']++;
            }
        });

        return branchSummary;
    }

    // Render Branch Performance Reports (Max/Min Visits/Calls)
    function renderBranchPerformanceReports() {
        reportDisplay.innerHTML = '<h2>Branch Performance Overview</h2>';
        const branchActivity = getBranchActivitySummary();

        let maxVisitsBranch = { name: 'N/A', count: -1 };
        let minVisitsBranch = { name: 'N/A', count: Infinity };
        let maxCallsBranch = { name: 'N/A', count: -1 };
        let minCallsBranch = { name: 'N/A', count: Infinity };
        let hasData = false;

        for (const branchName in branchActivity) {
            hasData = true;
            const visits = branchActivity[branchName]['Visit'];
            const calls = branchActivity[branchName]['Call'];

            if (visits > maxVisitsBranch.count) {
                maxVisitsBranch = { name: branchName, count: visits };
            }
            if (visits < minVisitsBranch.count) {
                minVisitsBranch = { name: branchName, count: visits };
            }
            if (calls > maxCallsBranch.count) {
                maxCallsBranch = { name: branchName, count: calls };
            }
            if (calls < minCallsBranch.count) {
                minCallsBranch = { name: branchName, count: calls };
            }
        }

        if (!hasData) {
            reportDisplay.innerHTML += '<p class="info-message">No activity data available to generate branch performance reports.</p>';
            return;
        }

        const reportContent = `
            <div class="branch-performance-grid">
                <div class="performance-card">
                    <h3>Maximum Visits</h3>
                    <p><strong>${maxVisitsBranch.name}</strong> with ${maxVisitsBranch.count} visits</p>
                </div>
                <div class="performance-card">
                    <h3>Minimum Visits</h3>
                    <p><strong>${minVisitsBranch.name}</strong> with ${minVisitsBranch.count} visits</p>
                </div>
                <div class="performance-card">
                    <h3>Maximum Calls</h3>
                    <p><strong>${maxCallsBranch.name}</strong> with ${maxCallsBranch.count} calls</p>
                </div>
                <div class="performance-card">
                    <h3>Minimum Calls</h3>
                    <p><strong>${minCallsBranch.name}</strong> with ${minCallsBranch.count} calls</p>
                </div>
            </div>
            <h3 style="margin-top: 30px;">Branch-wise Activity Breakdown</h3>
            <table class="branch-summary-table">
                <thead>
                    <tr>
                        <th>Branch Name</th>
                        <th>Total Visits</th>
                        <th>Total Calls</th>
                    </tr>
                </thead>
                <tbody>
                    ${PREDEFINED_BRANCHES.map(branch => `
                        <tr>
                            <td>${branch}</td>
                            <td>${branchActivity[branch] ? branchActivity[branch]['Visit'] : 0}</td>
                            <td>${branchActivity[branch] ? branchActivity[branch]['Call'] : 0}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;
        reportDisplay.innerHTML += reportContent;
    }


    // *** Event Listeners ***

    // Tab switching logic
    function showTab(tabId) {
        // Deactivate all buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Activate selected button and show relevant section
        document.getElementById(tabId).classList.add('active');

        if (tabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            renderAllBranchSnapshot();
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            renderAllStaffOverallPerformance();
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            renderNonParticipatingBranches();
        } else if (tabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'block';
            renderCustomerCanvassedList();
            customerDetailsContent.innerHTML = '<p class="info-message">Select a customer from the list to view details.</p>';
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            displayEmployeeManagementMessage('', false);
        } else if (tabId === 'branchPerformanceTabBtn') { // NEW TAB CASE
            reportsSection.style.display = 'block';
            renderBranchPerformanceReports(); // Call the new function
        }
    }

    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    }
    if (detailedCustomerViewTabBtn) {
        detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }
    // NEW Event Listener for Branch Performance Reports Tab
    if (branchPerformanceTabBtn) {
        branchPerformanceTabBtn.addEventListener('click', () => showTab('branchPerformanceTabBtn'));
    }


    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            // Only populate employee dropdown if detailed view is active
            if (detailedCustomerViewSection.style.display === 'block') {
                populateEmployeeDropdown(selectedBranch);
            }
            // Always re-render customer list based on current filters
            const selectedEmployee = employeeSelect ? employeeSelect.value : '';
            renderCustomerCanvassedList(selectedBranch, selectedEmployee);
        });
    }

    if (employeeSelect) {
        employeeSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect ? branchSelect.value : '';
            const selectedEmployee = employeeSelect.value;
            renderCustomerCanvassedList(selectedBranch, selectedEmployee);
        });
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeName = employeeNameInput.value.trim();
            const employeeCode = employeeCodeInput.value.trim();
            const designation = designationInput.value.trim();
            const branchName = branchNameInput.value.trim();

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Employee Name, Code, and Branch Name are required.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_DESIGNATION]: designation,
                [HEADER_BRANCH_NAME]: branchName
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Update Employee Form
    if (updateEmployeeForm) {
        updateEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeCode = updateEmployeeCodeInput.value.trim();
            const newEmployeeName = updateEmployeeNameInput.value.trim();
            const newDesignation = updateDesignationInput.value.trim();
            const newBranchName = updateBranchNameInput.value.trim();

            if (!employeeCode) {
                displayEmployeeManagementMessage('Employee Code is required to update.', true);
                return;
            }

            const updateData = {
                [HEADER_EMPLOYEE_CODE]: employeeCode
            };
            if (newEmployeeName) updateData[HEADER_EMPLOYEE_NAME] = newEmployeeName;
            if (newDesignation) updateData[HEADER_DESIGNATION] = newDesignation;
            if (newBranchName) updateData[HEADER_BRANCH_NAME] = newBranchName;


            if (Object.keys(updateData).length === 1) { // Only employeeCode, no other updates
                displayEmployeeManagementMessage('Please provide at least one field (Name, Designation, or Branch) to update.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('update_employee', updateData);

            if (success) {
                updateEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = employeeDetails.split('\n');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2 && parts[0] && parts[1]) { // Name and Code are mandatory
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
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
