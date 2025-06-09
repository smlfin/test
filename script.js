document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

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
        'Sales Exec': {
            'Visit': 20,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 2 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 30
        },
        'Sales Officer': {
            'Visit': 15,
            'Call': 4 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 25
        }
    };

    // --- Headers for Canvassing Data (from your published Google Sheet CSV) ---
    // Make sure these match the exact headers in your CSV!
    const HEADER_DATE = 'Date';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_COUNT = 'Count';

    // --- Headers for Master Employee Data (from Liv).xlsx - Sheet1.csv) ---
    // Make sure these match the exact headers in your Master Employee Google Sheet!
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code';
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name';
    const MASTER_HEADER_BRANCH_NAME_MASTER = 'Branch Name'; // Renamed to avoid conflict with activity data branch name
    const MASTER_HEADER_DESIGNATION = 'Designation'; // Assuming designation is also in master

    // Global variables to store fetched data
    let canvassingData = [];
    let masterEmployeeData = [];
    let branches = []; // This will hold all unique branches from both data sources
    let employees = []; // This will hold all unique employees with their details

    // References to DOM elements
    const statusMessageDiv = document.getElementById('statusMessage');
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const dateFilter = document.getElementById('dateFilter');
    const reportContent = document.getElementById('reportContent');
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const reportsSection = document.getElementById('reportsSection');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    // Helper function to display messages
    function displayMessage(message, isError = false, autoHide = true) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = isError ? 'message-container error' : 'message-container success';
        statusMessageDiv.style.display = 'block';

        if (autoHide && !isError) {
            setTimeout(() => {
                statusMessageDiv.style.display = 'none';
            }, 5000); // Hide success messages after 5 seconds
        }
    }

    // Helper function to display messages specific to employee management
    function displayEmployeeManagementMessage(message, isError = false, autoHide = true) {
        const messageContainer = document.getElementById('employeeManagementMessage');
        if (messageContainer) {
            messageContainer.textContent = message;
            messageContainer.className = isError ? 'message-container error' : 'message-container success';
            messageContainer.style.display = 'block';
            if (autoHide && !isError) {
                setTimeout(() => {
                    messageContainer.style.display = 'none';
                }, 5000);
            }
        } else {
            console.warn('Employee management message container not found.');
            displayMessage(message, isError, autoHide); // Fallback to main message display
        }
    }


    // Function to parse CSV data
    function parseCSV(text) {
        const lines = text.trim().split('\n');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',');
            if (values.length === headers.length) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index].trim();
                });
                data.push(row);
            }
        }
        return data;
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, payload) {
        try {
            displayMessage('Processing request...', false, false);
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(payload)
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                displayMessage(result.message || 'Action completed successfully!', false);
                return result;
            } else {
                displayMessage(result.message || 'An error occurred.', true);
                console.error(`Apps Script Error (${action}):`, result.message);
                return null;
            }
        } catch (error) {
            displayMessage(`Network or script error for ${action}: ${error.message}`, true);
            console.error(`Fetch Error (${action}):`, error);
            return null;
        }
    }

    // Function to fetch all data (Canvassing & Master Employee)
    async function fetchData() {
        displayMessage('Fetching data...', false, false);
        try {
            // Fetch canvassing data
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}. Please check DATA_URL.`);
            }
            const csvText = await response.text();
            canvassingData = parseCSV(csvText);
            console.log('Canvassing Data:', canvassingData);

            // Fetch master employee data from Apps Script
            // Ensure WEB_APP_URL is correct and the Apps Script is deployed with 'Anyone' access
            const masterDataResponse = await sendDataToGoogleAppsScript('get_master_employees', {});
            
            if (masterDataResponse && masterDataResponse.status === 'success') {
                masterEmployeeData = masterDataResponse.data;
                console.log('Master Employee Data:', masterEmployeeData);
            } else {
                console.error('Failed to fetch master employee data:', masterDataResponse ? masterDataResponse.message : 'Unknown error in Apps Script response');
                displayMessage('Failed to load employee master data. Non-participating branches report might be inaccurate.', true);
                masterEmployeeData = []; // Ensure it's an empty array if fetch fails
            }

            processData(); // Process all data after both fetches are complete
            displayMessage('Data loaded successfully!', true, true);
        } catch (error) {
            console.error('Error fetching data:', error);
            displayMessage(`Error loading data: ${error.message}. Please check console for details.`, true, true);
        }
    }

    // Function to process data and populate global arrays
    function processData() {
        const uniqueBranches = new Set();
        const uniqueEmployees = new Map(); // Map to store employee details: code -> {name, branch, designation}

        // Add from canvassing data
        canvassingData.forEach(row => {
            if (row[HEADER_BRANCH_NAME]) {
                uniqueBranches.add(row[HEADER_BRANCH_NAME]);
            }
            if (row[HEADER_EMPLOYEE_CODE] && !uniqueEmployees.has(row[HEADER_EMPLOYEE_CODE])) {
                 uniqueEmployees.set(row[HEADER_EMPLOYEE_CODE], {
                    name: row[HEADER_EMPLOYEE_NAME],
                    branch: row[HEADER_BRANCH_NAME],
                    designation: row[HEADER_DESIGNATION]
                });
            }
        });

        // Add from master employee data (this is crucial for getting all branches/employees)
        masterEmployeeData.forEach(emp => {
            if (emp[MASTER_HEADER_BRANCH_NAME_MASTER]) {
                uniqueBranches.add(emp[MASTER_HEADER_BRANCH_NAME_MASTER]);
            }
            if (emp[MASTER_HEADER_EMPLOYEE_CODE] && !uniqueEmployees.has(emp[MASTER_HEADER_EMPLOYEE_CODE])) {
                uniqueEmployees.set(emp[MASTER_HEADER_EMPLOYEE_CODE], {
                    name: emp[MASTER_HEADER_EMPLOYEE_NAME],
                    branch: emp[MASTER_HEADER_BRANCH_NAME_MASTER],
                    designation: emp[MASTER_HEADER_DESIGNATION] || 'Unknown' // Provide a default if not present
                });
            }
        });

        branches = Array.from(uniqueBranches).sort();
        employees = Array.from(uniqueEmployees.values()).sort((a, b) => a.name.localeCompare(b.name));

        populateBranchFilter();
        populateEmployeeFilter();
        console.log('All Branches (from combined data):', branches);
        console.log('All Employees (from combined data):', employees);
    }

    // Function to filter data based on selections
    function filterData() {
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value;
        const selectedDate = dateFilter.value;

        let filtered = canvassingData;

        if (selectedBranch) {
            filtered = filtered.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
            // If a branch is selected, show employee filter
            document.getElementById('employeeFilterPanel').style.display = 'block';
            populateEmployeeFilter(selectedBranch); // Repopulate employee filter for selected branch
        } else {
            // If no branch is selected, hide employee filter and show all employees
            document.getElementById('employeeFilterPanel').style.display = 'none';
            populateEmployeeFilter(); // Populate with all employees
        }

        if (selectedEmployee) {
            filtered = filtered.filter(row => row[HEADER_EMPLOYEE_CODE] === selectedEmployee);
        }

        if (selectedDate) {
            filtered = filtered.filter(row => row[HEADER_DATE] === selectedDate);
        }

        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            renderReport(activeTabButton.id, filtered);
        }
    }

    // Populate Branch Dropdown
    function populateBranchFilter() {
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        branches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    // Populate Employee Dropdown (can be filtered by branch)
    function populateEmployeeFilter(branchName = null) {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        let employeesToDisplay = employees;

        if (branchName) {
            employeesToDisplay = employees.filter(emp => emp.branch === branchName);
        }

        employeesToDisplay.forEach(emp => {
            const option = document.createElement('option');
            option.value = emp.name; // Use employee name as value
            option.textContent = `${emp.name} (${emp.branch} - ${emp.designation})`;
            employeeSelect.appendChild(option);
        });
    }

    // Render reports based on active tab
    function renderReport(tabId, data = canvassingData) {
        reportContent.innerHTML = ''; // Clear previous content

        // Apply global filters before rendering reports
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value;
        const selectedDate = dateFilter.value;

        let filteredData = data;

        if (selectedBranch) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
        }
        if (selectedEmployee) {
            filteredData = filteredData.filter(row => row[HEADER_EMPLOYEE_CODE] === selectedEmployee || row[HEADER_EMPLOYEE_NAME] === selectedEmployee);
        }
        if (selectedDate) {
            filteredData = filteredData.filter(row => row[HEADER_DATE] === selectedDate);
        }

        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                reportsSection.style.display = 'block';
                employeeManagementSection.style.display = 'none';
                renderAllBranchSnapshot(filteredData);
                break;
            case 'allStaffOverallPerformanceTabBtn':
                reportsSection.style.display = 'block';
                employeeManagementSection.style.display = 'none';
                renderAllStaffOverallPerformance(filteredData);
                break;
            case 'nonParticipatingBranchesTabBtn': // New tab for non-participating branches
                reportsSection.style.display = 'block';
                employeeManagementSection.style.display = 'none';
                renderNonParticipatingBranches(canvassingData); // Pass unfiltered canvassingData for this report
                break;
            case 'employeeManagementTabBtn':
                reportsSection.style.display = 'none';
                employeeManagementSection.style.display = 'block';
                // No rendering needed, just showing the section
                break;
            default:
                break;
        }
    }

    // --- Report Rendering Functions ---

    // All Branch Snapshot Report
    function renderAllBranchSnapshot(data) {
        const branchSummary = {}; // Branch Name -> { totalVisits, totalCalls, totalReferences, totalLeads, employees: Set }

        data.forEach(row => {
            const branchName = row[HEADER_BRANCH_NAME];
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const count = parseInt(row[HEADER_COUNT]) || 0;
            const employeeCode = row[HEADER_EMPLOYEE_CODE];

            if (!branchSummary[branchName]) {
                branchSummary[branchName] = {
                    totalVisits: 0,
                    totalCalls: 0,
                    totalReferences: 0,
                    totalNewCustomerLeads: 0,
                    employees: new Set()
                };
            }

            switch (activityType) {
                case 'Visit':
                    branchSummary[branchName].totalVisits += count;
                    break;
                case 'Call':
                    branchSummary[branchName].totalCalls += count;
                    break;
                case 'Reference':
                    branchSummary[branchName].totalReferences += count;
                    break;
                case 'New Customer Leads':
                    branchSummary[branchName].totalNewCustomerLeads += count;
                    break;
            }
            branchSummary[branchName].employees.add(employeeCode);
        });

        let totalBranchesParticipated = Object.keys(branchSummary).length;

        let html = `
            <h2>All Branch Snapshot</h2>
            <div class="summary-details-container">
                <ul class="summary-list">
                    <li>Total Branches: <span>${branches.length}</span></li>
                    <li>Branches with Participation: <span>${totalBranchesParticipated}</span></li>
                </ul>
            </div>
        `;

        if (totalBranchesParticipated === 0) {
            html += `<p class="no-data-message">No participation data available for the selected filters.</p>`;
        } else {
            html += `
                <table class="data-table all-branch-snapshot-table">
                    <thead>
                        <tr>
                            <th>Branch Name</th>
                            <th>Total Visits</th>
                            <th>Total Calls</th>
                            <th>Total References</th>
                            <th>Total New Customer Leads</th>
                            <th>Participating Staff</th>
                        </tr>
                    </thead>
                    <tbody>
            `;
            // Sort branches alphabetically for consistent display
            const sortedBranches = Object.keys(branchSummary).sort();

            sortedBranches.forEach(branchName => {
                const summary = branchSummary[branchName];
                html += `
                    <tr>
                        <td data-label="Branch Name">${branchName}</td>
                        <td data-label="Total Visits">${summary.totalVisits}</td>
                        <td data-label="Total Calls">${summary.totalCalls}</td>
                        <td data-label="Total References">${summary.totalReferences}</td>
                        <td data-label="Total New Customer Leads">${summary.totalNewCustomerLeads}</td>
                        <td data-label="Participating Staff">${summary.employees.size}</td>
                    </tr>
                `;
            });
            html += `
                    </tbody>
                </table>
            `;
        }

        reportContent.innerHTML = html;
    }


    // All Staff Overall Performance Report
    function renderAllStaffOverallPerformance(data) {
        const employeePerformance = {}; // Employee Code -> { name, branch, designation, totalVisits, totalCalls, totalReferences, totalNewCustomerLeads }

        // Initialize all known employees (from master data) with zero performance
        employees.forEach(emp => {
            employeePerformance[emp.code] = {
                name: emp.name,
                branch: emp.branch,
                designation: emp.designation,
                totalVisits: 0,
                totalCalls: 0,
                totalReferences: 0,
                totalNewCustomerLeads: 0
            };
        });

        data.forEach(row => {
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const count = parseInt(row[HEADER_COUNT]) || 0;

            if (employeePerformance[employeeCode]) { // Ensure employee is known
                switch (activityType) {
                    case 'Visit':
                        employeePerformance[employeeCode].totalVisits += count;
                        break;
                    case 'Call':
                        employeePerformance[employeeCode].totalCalls += count;
                        break;
                    case 'Reference':
                        employeePerformance[employeeCode].totalReferences += count;
                        break;
                    case 'New Customer Leads':
                        employeePerformance[employeeCode].totalNewCustomerLeads += count;
                        break;
                }
            } else {
                console.warn(`Activity recorded for unknown employee code: ${employeeCode}`);
            }
        });

        let html = `
            <h2>All Staff Overall Performance</h2>
        `;

        const sortedEmployees = Object.values(employeePerformance).sort((a, b) => {
            // Sort by branch, then by employee name
            if (a.branch !== b.branch) {
                return a.branch.localeCompare(b.branch);
            }
            return a.name.localeCompare(b.name);
        });

        if (sortedEmployees.length === 0) {
            html += `<p class="no-data-message">No employee data available for the selected filters.</p>`;
        } else {
            html += `
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Employee Name</th>
                            <th>Branch</th>
                            <th>Designation</th>
                            <th>Total Visits</th>
                            <th>Total Calls</th>
                            <th>Total References</th>
                            <th>Total New Customer Leads</th>
                            <th>Visit Target</th>
                            <th>Call Target</th>
                            <th>Reference Target</th>
                            <th>New Customer Leads Target</th>
                        </tr>
                    </thead>
                    <tbody>
            `;
            sortedEmployees.forEach(emp => {
                const target = TARGETS[emp.designation] || {};
                html += `
                    <tr>
                        <td data-label="Employee Name">${emp.name}</td>
                        <td data-label="Branch">${emp.branch}</td>
                        <td data-label="Designation">${emp.designation}</td>
                        <td data-label="Total Visits">${emp.totalVisits}</td>
                        <td data-label="Total Calls">${emp.totalCalls}</td>
                        <td data-label="Total References">${emp.totalReferences}</td>
                        <td data-label="Total New Customer Leads">${emp.totalNewCustomerLeads}</td>
                        <td data-label="Visit Target">${target.Visit !== undefined ? target.Visit : 'N/A'}</td>
                        <td data-label="Call Target">${target.Call !== undefined ? target.Call : 'N/A'}</td>
                        <td data-label="Reference Target">${target.Reference !== undefined ? target.Reference : 'N/A'}</td>
                        <td data-label="New Customer Leads Target">${target['New Customer Leads'] !== undefined ? target['New Customer Leads'] : 'N/A'}</td>
                    </tr>
                `;
            });
            html += `
                    </tbody>
                </table>
            `;
        }

        reportContent.innerHTML = html;
    }

    // New: Render Non-Participating Branches Report
    function renderNonParticipatingBranches(allCanvassingData) {
        const participatingBranches = new Set();
        allCanvassingData.forEach(row => {
            if (row[HEADER_BRANCH_NAME]) {
                participatingBranches.add(row[HEADER_BRANCH_NAME]);
            }
        });

        const nonParticipatingBranchList = [];
        branches.forEach(branch => { // 'branches' global array contains all known branches
            if (!participatingBranches.has(branch)) {
                nonParticipatingBranchList.push(branch);
            }
        });

        let html = `
            <h2>Non-Participating Branches</h2>
            <p>This report identifies branches that have not recorded any canvassing activities for the current period (based on selected filters).</p>
        `;

        if (nonParticipatingBranchList.length === 0) {
            html += `<p class="no-participation-message">All branches have participation records for the current period!</p>`;
        } else {
            html += `<p>The following branches have no recorded participation:</p>`;
            html += `<ul class="non-participating-branch-list">`;
            nonParticipatingBranchList.forEach(branch => {
                html += `<li>${branch}</li>`;
            });
            html += `</ul>`;
        }
        reportContent.innerHTML = html;
    }

    // --- Event Listeners ---

    // Tab Navigation
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));

    function showTab(tabId) {
        // Remove 'active' from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Add 'active' to the clicked button
        document.getElementById(tabId).classList.add('active');

        // Hide all main sections
        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Show the relevant section and render report
        if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
        } else {
            reportsSection.style.display = 'block';
            filterData(); // Re-render report for the selected tab with current filters
        }
    }

    // Filters Event Listeners
    branchSelect.addEventListener('change', filterData);
    employeeSelect.addEventListener('change', filterData);
    dateFilter.addEventListener('change', filterData);

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeName = document.getElementById('employeeName').value.trim();
            const employeeCode = document.getElementById('employeeCode').value.trim();
            const employeeBranch = document.getElementById('employeeBranch').value.trim();
            const employeeDesignation = document.getElementById('employeeDesignation').value.trim();

            if (!employeeName || !employeeCode || !employeeBranch || !employeeDesignation) {
                displayEmployeeManagementMessage('All fields are required.', true);
                return;
            }

            const employeeData = {
                'Employee Name': employeeName,
                'Employee Code': employeeCode,
                'Branch Name': employeeBranch,
                'Designation': employeeDesignation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data to update reports and dropdowns
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = document.getElementById('bulkEmployeeBranchName').value.trim();
            const bulkDetails = document.getElementById('bulkEmployeeDetails').value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeeLines = bulkDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Expect at least Name,Code (Designation is optional in input but we want to store it)
                    displayEmployeeManagementMessage(`Skipping invalid entry: ${line}. Format: Name,Code,Designation (Designation is optional).`, true);
                    continue;
                }
                const employeeData = {
                    [MASTER_HEADER_EMPLOYEE_NAME]: parts[0],
                    [MASTER_HEADER_EMPLOYEE_CODE]: parts[1],
                    [MASTER_HEADER_BRANCH_NAME_MASTER]: branchName,
                    [MASTER_HEADER_DESIGNATION]: parts[2] || '' // Use empty string if designation is not provided
                };
                employeesToAdd.push(employeeData);
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset(); // Clear form after submission
                    fetchData(); // Refresh data
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

            const deleteData = { [MASTER_HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data
            }
        });
    }

    // Initial data fetch when the page loads
    fetchData();

    // Show the "All Branch Snapshot" tab by default on load
    showTab('allBranchSnapshotTabBtn');
});
