document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 10,
            'New Loan Leads': 5,
            'Existing Customer Follow-up': 20
        },
        'Sales Executive': {
            'Visit': 20,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 2 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20,
            'New Loan Leads': 10,
            'Existing Customer Follow-up': 30
        },
        'Loan Officer': {
            'Visit': 15,
            'Call': 4 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 15,
            'New Loan Leads': 15,
            'Existing Customer Follow-up': 25
        }
    };

    // Header constants for easier access and consistency
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_CUSTOMER_NAME = 'Customer Name';
    const HEADER_CUSTOMER_CONTACT = 'Customer Contact';
    const HEADER_CUSTOMER_LOCATION = 'Customer Location';
    const HEADER_LOAN_TYPE = 'Loan Type';
    const HEADER_LOAN_AMOUNT = 'Loan Amount';
    const HEADER_REFERENCE_DETAILS = 'Reference Details';
    const HEADER_REMARKS = 'Remarks';
    const HEADER_STATUS = 'Status';

    // Global variables to store processed data
    let allCanvassingData = [];
    let allBranches = [];
    let allEmployees = [];
    let employeesByBranch = {};

    // DOM Elements
    const statusMessageDiv = document.getElementById('statusMessage');
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    const reportsSection = document.getElementById('reportsSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    const allBranchSnapshotReport = document.getElementById('allBranchSnapshotReport');
    const allStaffOverallPerformanceReport = document.getElementById('allStaffOverallPerformanceReport');
    const nonParticipatingBranchesReport = document.getElementById('nonParticipatingBranchesReport');
    const employeeDetailedEntriesReport = document.getElementById('employeeDetailedEntriesReport');
    const branchDetailedEntriesReport = document.getElementById('branchDetailedEntriesReport');
    const branchDetailedEntriesTitle = document.getElementById('branchDetailedEntriesTitle');
    const backToBranchSnapshotBtn = document.getElementById('backToBranchSnapshotBtn');

    const customerDetailedEntriesReport = document.getElementById('customerDetailedEntriesReport'); // NEW
    const customerDetailedEntriesTitle = document.getElementById('customerDetailedEntriesTitle'); // NEW
    const backToEmployeeSpecificReportsBtn = document.getElementById('backToEmployeeSpecificReportsBtn'); // NEW

    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const backToEmployeeFilterBtn = document.getElementById('backToEmployeeFilterBtn');
    const employeeDetailedEntriesTitle = document.getElementById('employeeDetailedEntriesTitle');

    const allBranchSnapshotTable = document.getElementById('allBranchSnapshotTable');
    const allStaffOverallPerformanceTable = document.getElementById('allStaffOverallPerformanceTable');
    const nonParticipatingBranchesTable = document.getElementById('nonParticipatingBranchesTable');
    const employeeDetailedEntriesTable = document.getElementById('employeeDetailedEntriesTable');
    const branchDetailedEntriesTable = document.getElementById('branchDetailedEntriesTable');
    const customerDetailedEntriesTable = document.getElementById('customerDetailedEntriesTable'); // NEW

    const employeeSpecificReportsContainer = document.getElementById('employeeSpecificReportsContainer'); // NEW
    const customerVisitsTable = document.getElementById('customerVisitsTable'); // NEW
    const customerReferencesTable = document.getElementById('customerReferencesTable'); // NEW
    const newCustomerLeadsTable = document.getElementById('newCustomerLeadsTable'); // NEW
    const newLoanLeadsTable = document.getElementById('newLoanLeadsTable'); // NEW


    // Employee Management Forms
    const employeeManagementMessageDiv = document.getElementById('employeeManagementMessage');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeBranchNameInput = document.getElementById('employeeBranchName');
    const employeeDesignationInput = document.getElementById('employeeDesignation');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Utility Functions
    function showStatusMessage(message, isError = false) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = `message-container ${isError ? 'error' : 'success'}`;
        statusMessageDiv.style.display = 'block';
        setTimeout(() => {
            statusMessageDiv.style.display = 'none';
        }, 5000);
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessageDiv.textContent = message;
        employeeManagementMessageDiv.className = `message-container ${isError ? 'error' : 'success'}`;
        employeeManagementMessageDiv.style.display = 'block';
        setTimeout(() => {
            employeeManagementMessageDiv.style.none = 'none';
        }, 5000);
    }

    // Function to fetch CSV data
    async function fetchCSV(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const text = await response.text();
            return Papa.parse(text, { header: true }).data;
        } catch (error) {
            showStatusMessage(`Error fetching data: ${error.message}. Please check the DATA_URL configuration and your internet connection.`, true);
            console.error('Error fetching CSV:', error);
            return [];
        }
    }

    // Function to send data to Google Apps Script
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
                }).toString()
            });

            if (!response.ok) {
                throw new Error(`Network response was not ok: ${response.statusText}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                showStatusMessage(result.message || 'Operation successful!');
                processData(); // Re-fetch and re-render data after successful operation
                return true;
            } else {
                showStatusMessage(result.message || 'Operation failed.', true);
                console.error('Apps Script Error:', result.error);
                return false;
            }
        } catch (error) {
            showStatusMessage(`Error communicating with Google Apps Script: ${error.message}. Please check WEB_APP_URL.`, true);
            console.error('Fetch error:', error);
            return false;
        }
    }

    // Process and populate data
    async function processData() {
        showStatusMessage('Loading data...', false);
        allCanvassingData = await fetchCSV(DATA_URL);

        // Filter out empty rows that Papa Parse might include
        allCanvassingData = allCanvassingData.filter(row => Object.values(row).some(val => val !== null && val !== undefined && val !== ''));

        // Extract unique branches and employees from canvassing data
        const uniqueBranches = new Set();
        const uniqueEmployees = new Set(); // For overall staff performance dropdown
        employeesByBranch = {}; // For branch-specific employee dropdown

        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (branch) uniqueBranches.add(branch);
            if (employeeCode && employeeName && designation) {
                uniqueEmployees.add(`${employeeName} (${employeeCode}) - ${designation}`);
                if (!employeesByBranch[branch]) {
                    employeesByBranch[branch] = new Set();
                }
                employeesByBranch[branch].add({
                    name: employeeName,
                    code: employeeCode,
                    designation: designation
                });
            }
        });

        allBranches = Array.from(uniqueBranches).sort();
        allEmployees = Array.from(uniqueEmployees).sort(); // Sort all employees for overall dropdown

        populateBranchSelect();
        // Initially populate employee select with all employees if no branch is selected
        populateEmployeeSelect(); 
        renderAllBranchSnapshot();
        renderAllStaffOverallPerformance();
        renderNonParticipatingBranches();
        showStatusMessage('Data loaded successfully!', false);
    }

    function populateBranchSelect() {
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        allBranches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    function populateEmployeeSelect(branch = '') {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        let employeesToDisplay = [];

        if (branch && employeesByBranch[branch]) {
            employeesToDisplay = Array.from(employeesByBranch[branch]).map(emp => `${emp.name} (${emp.code}) - ${emp.designation}`).sort();
        } else {
            // Only populate with employees who have data, not necessarily all employees from master
            // This ensures dropdown only shows relevant employees if no branch is selected or if data is empty.
            employeesToDisplay = Array.from(new Set(allCanvassingData.map(entry => {
                if (entry[HEADER_EMPLOYEE_NAME] && entry[HEADER_EMPLOYEE_CODE] && entry[HEADER_DESIGNATION]) {
                    return `${entry[HEADER_EMPLOYEE_NAME]} (${entry[HEADER_EMPLOYEE_CODE]}) - ${entry[HEADER_DESIGNATION]}`;
                }
                return null;
            }).filter(Boolean))).sort();
        }

        employeesToDisplay.forEach(employee => {
            const option = document.createElement('option');
            option.value = employee;
            option.textContent = employee;
            employeeSelect.appendChild(option);
        });
        employeeSelect.value = ''; // Reset selected employee
    }


    // Tab and Report Display Logic
    function showTab(tabId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
        // Hide all report sections
        document.querySelectorAll('.report-content').forEach(report => report.style.display = 'none');
        // Hide employee management section
        employeeManagementSection.style.display = 'none';
        reportsSection.style.display = 'none';
        
        // Hide employee filter panel and specific reports container by default
        employeeFilterPanel.style.display = 'none'; 
        employeeSpecificReportsContainer.style.display = 'none'; 
        customerDetailedEntriesReport.style.display = 'none'; // Hide customer specific detailed report
        employeeDetailedEntriesReport.style.display = 'none'; // Hide employee detailed report
        branchDetailedEntriesReport.style.display = 'none'; // Hide branch detailed report


        // Activate selected tab button and show relevant section
        document.getElementById(tabId).classList.add('active');

        if (tabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            allBranchSnapshotReport.style.display = 'block';
            employeeFilterPanel.style.display = 'flex'; // Show employee filter under All Branch Snapshot
            branchSelect.value = ''; // Reset branch selection
            employeeSelect.value = ''; // Reset employee selection
            populateEmployeeSelect(''); // Show all employees initially in the dropdown
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            allStaffOverallPerformanceReport.style.display = 'block';
            // For this tab, the dropdowns are linked to viewing all entries for an employee
            employeeFilterPanel.style.display = 'flex'; 
            branchSelect.value = ''; // Reset branch selection
            employeeSelect.value = ''; // Reset employee selection
            populateEmployeeSelect(''); // Show all employees initially in the dropdown
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            nonParticipatingBranchesReport.style.display = 'block';
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
        }
    }

    // Render Functions for Reports

    function renderAllBranchSnapshot() {
        const branchData = {}; // branchName: { totalEntries: X, avgDailyEntries: Y }
        const today = new Date();
        const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const timestamp = new Date(entry[HEADER_TIMESTAMP]);

            if (branch && timestamp >= startOfMonth) { // Consider only current month's data
                if (!branchData[branch]) {
                    branchData[branch] = {
                        totalEntries: 0,
                        uniqueDates: new Set()
                    };
                }
                branchData[branch].totalEntries++;
                branchData[branch].uniqueDates.add(timestamp.toDateString());
            }
        });

        let tableHTML = `
            <thead>
                <tr>
                    <th>Branch Name</th>
                    <th>Total Entries (Current Month)</th>
                    <th>Avg. Daily Entries (Current Month)</th>
                </tr>
            </thead>
            <tbody>
        `;

        // Sort branches alphabetically
        const sortedBranches = Object.keys(branchData).sort();

        sortedBranches.forEach(branch => {
            const data = branchData[branch];
            const avgDailyEntries = (data.totalEntries / data.uniqueDates.size) || 0;
            tableHTML += `
                <tr>
                    <td><a href="#" class="branch-link" data-branch="${branch}">${branch}</a></td>
                    <td>${data.totalEntries}</td>
                    <td>${avgDailyEntries.toFixed(2)}</td>
                </tr>
            `;
        });

        if (sortedBranches.length === 0) {
            tableHTML += `<tr><td colspan="3">No canvassing data found for the current month.</td></tr>`;
        }

        tableHTML += `</tbody>`;
        allBranchSnapshotTable.innerHTML = tableHTML;

        // Add event listeners to newly created branch links
        document.querySelectorAll('.branch-link').forEach(link => {
            link.addEventListener('click', (event) => {
                event.preventDefault();
                const branchName = event.target.dataset.branch;
                renderBranchDetailedEntries(branchName);
            });
        });
    }

    function renderAllStaffOverallPerformance() {
        const employeePerformance = {}; // employeeCode: { name, branch, designation, totalEntries, monthlyTargetsMet: { activityType: count } }
        const today = new Date();
        const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const branch = entry[HEADER_BRANCH_NAME];
            const designation = entry[HEADER_DESIGNATION];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const timestamp = new Date(entry[HEADER_TIMESTAMP]);

            if (employeeCode && timestamp >= startOfMonth) { // Consider only current month's data
                if (!employeePerformance[employeeCode]) {
                    employeePerformance[employeeCode] = {
                        name: employeeName,
                        code: employeeCode, // Added code for easier access
                        branch: branch,
                        designation: designation,
                        totalEntries: 0,
                        activities: {}
                    };
                    Object.keys(TARGETS[designation] || {}).forEach(activity => {
                        employeePerformance[employeeCode].activities[activity] = 0;
                    });
                }
                employeePerformance[employeeCode].totalEntries++;
                if (activityType && employeePerformance[employeeCode].activities.hasOwnProperty(activityType)) {
                    employeePerformance[employeeCode].activities[activityType]++;
                }
            }
        });

        // Get all possible activity types for table headers
        const allActivityTypes = new Set();
        Object.values(TARGETS).forEach(designationTargets => {
            Object.keys(designationTargets).forEach(activity => allActivityTypes.add(activity));
        });
        const sortedActivityTypes = Array.from(allActivityTypes).sort();

        let tableHTML = `
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Branch</th>
                    <th>Designation</th>
                    <th>Total Entries (Current Month)</th>
        `;
        sortedActivityTypes.forEach(activity => {
            tableHTML += `<th>${activity}</th>`;
        });
        tableHTML += `
                </tr>
            </thead>
            <tbody>
        `;

        const sortedEmployees = Object.values(employeePerformance).sort((a, b) => a.name.localeCompare(b.name));

        sortedEmployees.forEach(emp => {
            tableHTML += `
                <tr>
                    <td>${emp.name} (${emp.code})</td>
                    <td>${emp.branch}</td>
                    <td>${emp.designation}</td>
                    <td>${emp.totalEntries}</td>
            `;
            sortedActivityTypes.forEach(activity => {
                tableHTML += `<td>${emp.activities[activity] || 0} / ${TARGETS[emp.designation]?.[activity] || '-'}</td>`;
            });
            tableHTML += `</tr>`;
        });

        if (sortedEmployees.length === 0) {
            tableHTML += `<tr><td colspan="${4 + sortedActivityTypes.length}">No employee performance data found for the current month.</td></tr>`;
        }

        tableHTML += `</tbody>`;
        allStaffOverallPerformanceTable.innerHTML = tableHTML;
    }

    function renderNonParticipatingBranches() {
        const participatingBranches = new Set();
        allCanvassingData.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                participatingBranches.add(entry[HEADER_BRANCH_NAME]);
            }
        });

        const nonParticipating = allBranches.filter(branch => !participatingBranches.has(branch));

        let tableHTML = `
            <thead>
                <tr>
                    <th>Branch Name</th>
                </tr>
            </thead>
            <tbody>
        `;

        if (nonParticipating.length > 0) {
            nonParticipating.sort().forEach(branch => {
                tableHTML += `<tr><td>${branch}</td></tr>`;
            });
        } else {
            tableHTML += `<tr><td>All branches have recorded canvassing activity.</td></tr>`;
        }

        tableHTML += `</tbody>`;
        nonParticipatingBranchesTable.innerHTML = tableHTML;
    }

    // Helper to render basic Date/Customer Name table with clickable customer names
    function renderSimpleCustomerTable(data, tableElement, title) {
        let tableHTML = `
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Customer Name</th>
                </tr>
            </thead>
            <tbody>
        `;

        if (data.length === 0) {
            tableHTML += `<tr><td colspan="2">No ${title.toLowerCase()} found.</td></tr>`;
        } else {
            // Sort by timestamp descending
            const sortedData = [...data].sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP]));

            sortedData.forEach(entry => {
                const timestamp = new Date(entry[HEADER_TIMESTAMP]);
                const formattedDate = timestamp.toLocaleDateString(); // Format date nicely
                const customerName = entry[HEADER_CUSTOMER_NAME] || 'N/A';
                const customerContact = entry[HEADER_CUSTOMER_CONTACT] || ''; // Pass contact for better uniqueness

                // Pass all necessary info for detailed view to the click handler
                tableHTML += `
                    <tr>
                        <td>${formattedDate}</td>
                        <td><a href="#" class="customer-link" 
                                data-customer-name="${customerName}" 
                                data-customer-contact="${customerContact}"
                                data-employee-code="${entry[HEADER_EMPLOYEE_CODE]}">${customerName}</a></td>
                    </tr>
                `;
            });
        }
        tableHTML += `</tbody>`;
        tableElement.innerHTML = tableHTML;

        // Attach event listeners after rendering
        tableElement.querySelectorAll('.customer-link').forEach(link => {
            link.addEventListener('click', (event) => {
                event.preventDefault();
                const customerName = event.target.dataset.customerName;
                const customerContact = event.target.dataset.customerContact;
                const employeeCode = event.target.dataset.employeeCode;
                renderCustomerDetailedEntries(customerName, customerContact, employeeCode);
            });
        });
    }

    // NEW: Functions to render the four specific tables
    function renderCustomerVisitsTable(branchName, employeeCode) {
        const filteredData = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            entry[HEADER_ACTIVITY_TYPE] === 'Visit' &&
            entry[HEADER_CUSTOMER_NAME]
        );
        renderSimpleCustomerTable(filteredData, customerVisitsTable, 'Customer Visits & Meeting Report');
    }

    function renderCustomerReferencesTable(branchName, employeeCode) {
        const filteredData = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            entry[HEADER_ACTIVITY_TYPE] === 'Reference' &&
            entry[HEADER_CUSTOMER_NAME]
        );
        renderSimpleCustomerTable(filteredData, customerReferencesTable, 'Customer Reference Generated');
    }

    function renderNewCustomerLeadsTable(branchName, employeeCode) {
        const filteredData = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            entry[HEADER_ACTIVITY_TYPE] === 'New Customer Leads' &&
            entry[HEADER_CUSTOMER_NAME]
        );
        renderSimpleCustomerTable(filteredData, newCustomerLeadsTable, 'New Customer Leads Generated');
    }

    function renderNewLoanLeadsTable(branchName, employeeCode) {
        const filteredData = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            entry[HEADER_ACTIVITY_TYPE] === 'New Loan Leads' &&
            entry[HEADER_CUSTOMER_NAME]
        );
        renderSimpleCustomerTable(filteredData, newLoanLeadsTable, 'New Loan Leads Generated');
    }


    // Unified function to render detailed entries (employee, branch, or customer)
    function renderDetailedEntriesTable(data, tableElement, titleElement, backButtonElement, backTo) {
        if (data.length === 0) {
            tableElement.innerHTML = `<tr><td colspan="9">No canvassing entries found.</td></tr>`;
            // Hide the table if no data
            tableElement.style.display = 'none'; 
            return;
        } else {
             tableElement.style.display = 'table'; // Show table if data exists
        }

        const headers = [
            HEADER_TIMESTAMP, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_EMPLOYEE_CODE,
            HEADER_DESIGNATION, HEADER_ACTIVITY_TYPE, HEADER_CUSTOMER_NAME,
            HEADER_CUSTOMER_CONTACT, HEADER_CUSTOMER_LOCATION, HEADER_LOAN_TYPE,
            HEADER_LOAN_AMOUNT, HEADER_REFERENCE_DETAILS, HEADER_REMARKS, HEADER_STATUS
        ];

        let tableHTML = `<thead><tr>`;
        headers.forEach(header => {
            tableHTML += `<th>${header}</th>`;
        });
        tableHTML += `</tr></thead><tbody>`;

        // Sort entries by timestamp descending for detailed views
        const sortedData = [...data].sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP]));


        sortedData.forEach(entry => {
            tableHTML += `<tr>`;
            headers.forEach(header => {
                tableHTML += `<td>${entry[header] || ''}</td>`;
            });
            tableHTML += `</tr>`;
        });
        tableHTML += `</tbody>`;
        tableElement.innerHTML = tableHTML;

        // Hide all main reports and show the specific detailed report
        document.querySelectorAll('.report-content').forEach(report => report.style.display = 'none');
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active')); // Deactivate all tabs
        employeeFilterPanel.style.display = 'none'; // Hide dropdowns
        employeeSpecificReportsContainer.style.display = 'none'; // Hide the four tables


        // Show the specific detailed report and its back button
        if (tableElement === employeeDetailedEntriesTable) {
            employeeDetailedEntriesReport.style.display = 'block';
            document.getElementById('allStaffOverallPerformanceTabBtn').classList.add('active'); // Keep overall performance tab active
        } else if (tableElement === branchDetailedEntriesTable) {
            branchDetailedEntriesReport.style.display = 'block';
            document.getElementById('allBranchSnapshotTabBtn').classList.add('active'); // Keep branch snapshot tab active
        } else if (tableElement === customerDetailedEntriesTable) { // NEW
            customerDetailedEntriesReport.style.display = 'block';
            document.getElementById('allBranchSnapshotTabBtn').classList.add('active'); // Keep branch snapshot tab active
        }
    }


    function renderEmployeeDetailedEntries(employeeCodeEntries) {
        const employeeName = employeeCodeEntries[0] ? employeeCodeEntries[0][HEADER_EMPLOYEE_NAME] : '';
        const employeeCode = employeeCodeEntries[0] ? employeeCodeEntries[0][HEADER_EMPLOYEE_CODE] : '';
        employeeDetailedEntriesTitle.textContent = `Customer Canvassing Details for Employee: ${employeeName} (${employeeCode})`;
        renderDetailedEntriesTable(employeeCodeEntries, employeeDetailedEntriesTable, employeeDetailedEntriesTitle, backToEmployeeFilterBtn, 'employeeFilter');
    }

    function renderBranchDetailedEntries(branchName) {
        const branchEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchName);
        branchDetailedEntriesTitle.textContent = `Customer Canvassing Details for Branch: ${branchName}`;
        renderDetailedEntriesTable(branchEntries, branchDetailedEntriesTable, branchDetailedEntriesTitle, backToBranchSnapshotBtn, 'allBranchSnapshot');
    }

    // NEW: Function to render detailed entries for a specific customer
    function renderCustomerDetailedEntries(customerName, customerContact, employeeCode) {
        // Filter by customer name, contact, and employee code for more precise results
        const customerEntries = allCanvassingData.filter(entry => 
            (entry[HEADER_CUSTOMER_NAME] === customerName || (!entry[HEADER_CUSTOMER_NAME] && customerName === 'N/A')) && 
            entry[HEADER_CUSTOMER_CONTACT] === customerContact &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode
        );
        customerDetailedEntriesTitle.textContent = `Customer Canvassing Details for: ${customerName}`;
        renderDetailedEntriesTable(customerEntries, customerDetailedEntriesTable, customerDetailedEntriesTitle, backToEmployeeSpecificReportsBtn, 'employeeSpecificReports');
    }


    // Event Listeners for Tab Buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));

    // Event Listener for Branch Select
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            populateEmployeeSelect(selectedBranch);
            employeeFilterPanel.style.display = 'flex'; // Show employee filter when branch is selected
        } else {
            // If no branch selected, reset employee dropdown and hide specific reports
            populateEmployeeSelect(); // Show all employees if no branch selected
            employeeFilterPanel.style.display = 'flex'; // Keep filter visible
            employeeSelect.value = ''; // Reset employee select
            employeeSpecificReportsContainer.style.display = 'none'; // Hide the four tables
            customerDetailedEntriesReport.style.display = 'none'; // Hide customer specific detailed report
            employeeDetailedEntriesReport.style.display = 'none'; // Hide employee detailed report
        }
        // Always try to render employee specific reports if both are selected, or hide them
        const selectedEmployeeText = employeeSelect.value;
        const employeeCodeMatch = selectedEmployeeText ? selectedEmployeeText.match(/\((EMP\d+)\)/) : null;
        const employeeCode = employeeCodeMatch ? employeeCodeMatch[1] : null;
        if (selectedBranch && employeeCode) {
             renderEmployeeSpecificReports(selectedBranch, employeeCode);
        } else {
            employeeSpecificReportsContainer.style.display = 'none';
        }
    });

    // NEW: Event listener for Employee Select
    employeeSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        const selectedEmployeeText = employeeSelect.value;

        if (selectedBranch && selectedEmployeeText) {
            const employeeCodeMatch = selectedEmployeeText.match(/\((EMP\d+)\)/);
            if (employeeCodeMatch) {
                const employeeCode = employeeCodeMatch[1];
                renderEmployeeSpecificReports(selectedBranch, employeeCode);
            } else {
                showStatusMessage('Could not parse employee code from selection.', true);
                employeeSpecificReportsContainer.style.display = 'none';
            }
        } else {
            employeeSpecificReportsContainer.style.display = 'none';
        }

        // Hide customer detailed view if employee selection changes
        customerDetailedEntriesReport.style.display = 'none';
        // Hide employee detailed entries (from "View All Entries (Employee)" button) if employee selection changes
        employeeDetailedEntriesReport.style.display = 'none';
    });

    // Function to render all four specific tables
    function renderEmployeeSpecificReports(branchName, employeeCode) {
        renderCustomerVisitsTable(branchName, employeeCode);
        renderCustomerReferencesTable(branchName, employeeCode);
        renderNewCustomerLeadsTable(branchName, employeeCode);
        renderNewLoanLeadsTable(branchName, employeeCode);
        employeeSpecificReportsContainer.style.display = 'block';
    }


    // Event Listener for View All Entries (Employee) Button (This is still for the "All Staff Performance (Overall)" use case)
    viewAllEntriesBtn.addEventListener('click', () => {
        const selectedEmployeeText = employeeSelect.value;
        if (!selectedEmployeeText) {
            showStatusMessage('Please select an employee to view entries.', true);
            return;
        }

        const employeeCodeMatch = selectedEmployeeText.match(/\((EMP\d+)\)/);
        if (!employeeCodeMatch) {
            showStatusMessage('Could not parse employee code from selection.', true);
            return;
        }
        const employeeCode = employeeCodeMatch[1];

        const employeeCodeEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
        renderEmployeeDetailedEntries(employeeCodeEntries);
    });

    // Event Listener for Back to Employee Filter Button (from Employee Detailed Entries)
    backToEmployeeFilterBtn.addEventListener('click', () => {
        employeeDetailedEntriesReport.style.display = 'none';
        showTab('allStaffOverallPerformanceTabBtn'); // This back button always goes to All Staff Performance
    });

    // Event Listener for Back to Branch Snapshot Button (from Branch Detailed Entries)
    backToBranchSnapshotBtn.addEventListener('click', () => {
        branchDetailedEntriesReport.style.display = 'none';
        showTab('allBranchSnapshotTabBtn'); // This back button always goes to All Branch Snapshot
        // Reset dropdowns to ensure initial state on return
        branchSelect.value = ''; 
        employeeSelect.value = '';
        employeeSpecificReportsContainer.style.display = 'none'; // Ensure the four tables are hidden
    });

    // NEW: Event Listener for Back to Employee Specific Reports (from Customer Detailed Entries)
    backToEmployeeSpecificReportsBtn.addEventListener('click', () => {
        customerDetailedEntriesReport.style.display = 'none';
        // Re-show the main reports section and the All Branch Snapshot content
        reportsSection.style.display = 'block';
        allBranchSnapshotReport.style.display = 'block';
        employeeFilterPanel.style.display = 'flex'; // Show dropdowns
        employeeSpecificReportsContainer.style.display = 'block'; // Show the four tables again
    });


    // Employee Management Forms Event Listeners

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: employeeBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: employeeDesignationInput.value.trim()
            };

            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH_NAME] || !employeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = employeeDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Minimum Name, Code
                    displayEmployeeManagementMessage(`Invalid entry: "${line}". Each line should be 'Name,Code,Designation'.`, true);
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
    showTab('allBranchSnapshotTabBtn'); // Default tab to display on load
});
