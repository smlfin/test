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
            'Referance': 1 * MONTHLY_WORKING_DAYS, // Stick to your spelling
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Referance': 1 * MONTHLY_WORKING_DAYS, // Stick to your spelling
            'New Customer Leads': 20
        },
        'Relationship Officer': { // Added Relationship Officer
            'Visit': 25,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Referance': 1 * MONTHLY_WORKING_DAYS, // Stick to your spelling
            'New Customer Leads': 15
        },
        'Telecaller': { // Added Telecaller
            'Call': 10 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 10,
            'Visit': 0, // No visit target for telecaller
            'Referance': 0 // No reference target for telecaller // Stick to your spelling
        },
        'Other': { // Default/Fallback for other designations
            'Visit': 5,
            'Call': 2 * MONTHLY_WORKING_DAYS,
            'Referance': 0.5 * MONTHLY_WORKING_DAYS, // Stick to your spelling
            'New Customer Leads': 5
        }
    };

    // --- Headers for Canvassing Data Sheet ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_VISIT = 'Visit';
    const HEADER_CALL = 'Call';
    const HEADER_REFERENCE = 'Referance'; // Changed to match your spelling
    const HEADER_NEW_CUSTOMER_LEADS = 'New Customer Leads'; // This will be calculated in JS
    const HEADER_REMARKS = 'Remarks';

    // *** NEW HEADERS FOR CALCULATION ***
    const HEADER_ACTIVITY_TYPE = 'Activity Type'; // Make sure this matches your actual column name
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer'; // Make sure this matches your actual column name

    // --- Headers for Master Employee Data (from Liv).xlsx - Sheet1.csv) ---
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code';
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name';
    const MASTER_HEADER_BRANCH_NAME_MASTER = 'Branch Name'; // Renamed to avoid conflict with activity data branch name

    // Function to fetch and parse Master Employee Data
    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED_FOR_FRONTEND_REPORTING";


    // --- Elements ---
    const statusMessageElement = document.getElementById('statusMessage');
    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const allStaffPerformanceTableBody = document.getElementById('allStaffPerformanceTableBody');
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const allStaffPerformanceTable = document.getElementById('allStaffPerformanceTable');
    const singleEmployeePerformanceContainer = document.getElementById('singleEmployeePerformanceContainer');
    const employeeNameDisplay = document.getElementById('employeeNameDisplay');
    const employeeCodeDisplay = document.getElementById('employeeCodeDisplay');
    const employeeBranchDisplay = document.getElementById('employeeBranchDisplay');
    const employeeDesignationDisplay = document.getElementById('employeeDesignationDisplay');
    const singleEmployeePerformanceTableBody = document.getElementById('singleEmployeePerformanceTableBody');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeBranchInput = document.getElementById('employeeBranch');
    const employeeDesignationInput = document.getElementById('employeeDesignation');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const noParticipationMessageContainer = document.getElementById('noParticipationMessageContainer');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewOptions = document.getElementById('viewOptions');
    const dateFilterStart = document.getElementById('dateFilterStart');
    const dateFilterEnd = document.getElementById('dateFilterEnd');
    const applyDateFilterBtn = document.getElementById('applyDateFilterBtn');
    const dateRangeDisplay = document.getElementById('dateRangeDisplay');
    const selectedDateRange = document.getElementById('selectedDateRange');
    const clearDateFilterBtn = document.getElementById('clearDateFilterBtn');


    // --- Global Data Storage ---
    let canvassingData = [];
    let masterEmployeeData = []; // Will be populated from Apps Script or a predefined list
    let branches = [];
    let employees = [];

    // --- Tab Navigation ---
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');

    const allBranchSnapshotContainer = document.getElementById('allBranchSnapshotContainer');
    const allStaffOverallPerformanceContainer = document.getElementById('allStaffOverallPerformanceContainer');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const nonParticipatingBranchesContainer = document.getElementById('nonParticipatingBranchesContainer');

    function showTab(tabButtonId) {
        // Hide all content sections
        allBranchSnapshotContainer.style.display = 'none';
        allStaffOverallPerformanceContainer.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        singleEmployeePerformanceContainer.style.display = 'none'; // Ensure single employee view is hidden
        nonParticipatingBranchesContainer.style.display = 'none';

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the selected section and activate its button
        let targetContainer;
        if (tabButtonId === 'allBranchSnapshotTabBtn') {
            targetContainer = allBranchSnapshotContainer;
            viewOptions.style.display = 'flex';
            employeeFilterPanel.style.display = 'none';
        } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            targetContainer = allStaffOverallPerformanceContainer;
            viewOptions.style.display = 'flex';
            employeeFilterPanel.style.display = 'flex';
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            targetContainer = employeeManagementSection;
            viewOptions.style.display = 'none'; // Hide date filters for employee management
            employeeFilterPanel.style.display = 'none';
            dateRangeDisplay.style.display = 'none';
            // Clear selections if coming from a report view
            branchSelect.value = '';
            employeeSelect.value = '';
        } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
            targetContainer = nonParticipatingBranchesContainer;
            viewOptions.style.display = 'none';
            employeeFilterPanel.style.display = 'none';
            dateRangeDisplay.style.display = 'none';
        }

        if (targetContainer) {
            targetContainer.style.display = 'block';
            document.getElementById(tabButtonId).classList.add('active');
            // Re-render reports based on current filters when tab is shown
            if (tabButtonId !== 'employeeManagementTabBtn') {
                renderReports();
            } else {
                displayEmployeeManagementMessage(''); // Clear any messages when switching to management
            }
        }
    }

    // --- Data Fetching and Processing ---
    async function fetchData() {
        displayMessage('Fetching data...', false, false);
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            canvassingData = parseCSV(csvText);
            console.log('Canvassing Data:', canvassingData);

            // Fetch master employee data from Apps Script
            const masterDataResponse = await sendDataToGoogleAppsScript('get_master_employees', {});
            if (masterDataResponse && masterDataResponse.status === 'success') {
                masterEmployeeData = masterDataResponse.data;
                console.log('Master Employee Data:', masterEmployeeData);
            } else {
                console.error('Failed to fetch master employee data:', masterDataResponse ? masterDataResponse.message : 'Unknown error');
                displayMessage('Failed to load employee master data. Some features might be limited.', true);
                masterEmployeeData = []; // Ensure it's an empty array if fetch fails
            }

            processData(); // Process all data after both fetches are complete
            displayMessage('Data loaded successfully!', true, true);
        } catch (error) {
            console.error('Error fetching data:', error);
            displayMessage(`Error loading data: ${error.message}. Please check console for details.`, true, true);
        }
    }

    function parseCSV(csvText) {
        const lines = csvText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            if (values.length === headers.length) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                data.push(row);
            }
        }
        return data;
    }

    function processData() {
        // Collect unique branches and employees from combined data (canvassing + master)
        const uniqueBranches = new Set();
        const uniqueEmployees = new Set(); // Stores Employee Code for uniqueness

        // Add from canvassing data
        canvassingData.forEach(row => {
            if (row[HEADER_BRANCH_NAME]) {
                uniqueBranches.add(row[HEADER_BRANCH_NAME]);
            }
            if (row[HEADER_EMPLOYEE_CODE] && row[HEADER_EMPLOYEE_NAME] && row[HEADER_BRANCH_NAME] && row[HEADER_DESIGNATION]) {
                uniqueEmployees.add(JSON.stringify({
                    code: row[HEADER_EMPLOYEE_CODE],
                    name: row[HEADER_EMPLOYEE_NAME],
                    branch: row[HEADER_BRANCH_NAME],
                    designation: row[HEADER_DESIGNATION]
                }));
            }
        });

        // Add from master employee data
        masterEmployeeData.forEach(emp => {
            if (emp[MASTER_HEADER_BRANCH_NAME_MASTER]) {
                uniqueBranches.add(emp[MASTER_HEADER_BRANCH_NAME_MASTER]);
            }
            if (emp[MASTER_HEADER_EMPLOYEE_CODE] && emp[MASTER_HEADER_EMPLOYEE_NAME] && emp[MASTER_HEADER_BRANCH_NAME_MASTER]) {
                uniqueEmployees.add(JSON.stringify({
                    code: emp[MASTER_HEADER_EMPLOYEE_CODE],
                    name: emp[MASTER_HEADER_EMPLOYEE_NAME],
                    branch: emp[MASTER_HEADER_BRANCH_NAME_MASTER],
                    designation: emp.Designation || 'Other' // Master data should ideally have designation
                }));
            }
        });

        branches = Array.from(uniqueBranches).sort();
        employees = Array.from(uniqueEmployees).map(empStr => JSON.parse(empStr)).sort((a, b) => a.name.localeCompare(b.name));

        populateFilters();
        renderReports();
    }


    function populateFilters() {
        // Populate Branch Select
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        branches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });

        // Populate Employee Select (initially for all employees)
        populateEmployeeFilter(employees);
    }

    function populateEmployeeFilter(employeesToFilter) {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        employeesToFilter.forEach(employee => {
            const option = document.createElement('option');
            option.value = employee.code;
            option.textContent = `${employee.name} (${employee.branch})`;
            employeeSelect.appendChild(option);
        });
    }

    // --- Report Rendering ---
    function renderReports() {
        const currentBranch = branchSelect.value;
        const currentEmployeeCode = employeeSelect.value;
        const startDate = dateFilterStart.value ? new Date(dateFilterStart.value) : null;
        const endDate = dateFilterEnd.value ? new Date(dateFilterEnd.value) : null;

        // Filter data based on selections and dates
        const filteredData = canvassingData.filter(row => {
            let matchesBranch = true;
            let matchesEmployee = true;
            let matchesDate = true;

            if (currentBranch && row[HEADER_BRANCH_NAME] !== currentBranch) {
                matchesBranch = false;
            }

            if (currentEmployeeCode && row[HEADER_EMPLOYEE_CODE] !== currentEmployeeCode) {
                matchesEmployee = false;
            }

            if (startDate || endDate) {
                const rowDate = new Date(row[HEADER_TIMESTAMP]);
                rowDate.setHours(0, 0, 0, 0); // Normalize to start of day for comparison
                if (startDate && rowDate < startDate) {
                    matchesDate = false;
                }
                if (endDate && rowDate > endDate) {
                    matchesDate = false;
                }
            }

            return matchesBranch && matchesEmployee && matchesDate;
        });

        // Determine which report to render based on active tab
        const activeTab = document.querySelector('.tab-button.active');
        if (!activeTab) return; // No active tab, do nothing

        if (activeTab.id === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot(filteredData);
            dateRangeDisplay.style.display = (startDate || endDate) ? 'flex' : 'none';
        } else if (activeTab.id === 'allStaffOverallPerformanceTabBtn') {
            if (currentEmployeeCode) {
                renderSingleEmployeePerformance(filteredData, currentEmployeeCode);
                singleEmployeePerformanceContainer.style.display = 'block';
                allStaffOverallPerformanceContainer.style.display = 'none';
            } else {
                renderAllStaffOverallPerformance(filteredData);
                singleEmployeePerformanceContainer.style.display = 'none';
                allStaffOverallPerformanceContainer.style.display = 'block';
            }
            dateRangeDisplay.style.display = (startDate || endDate) ? 'flex' : 'none';
        } else if (activeTab.id === 'nonParticipatingBranchesTabBtn') {
            renderNonParticipatingBranches(canvassingData); // Non-participating branches always uses full data
            dateRangeDisplay.style.display = 'none'; // Date filter does not apply here
        } else if (activeTab.id === 'employeeManagementTabBtn') {
            // Do nothing, employee management section handled by its own forms
        }
        updateDateRangeDisplay(startDate, endDate);
    }

    function renderAllBranchSnapshot(data) {
        const branchSummary = {};
        const employeeSet = new Set(); // To count unique employees overall

        data.forEach(row => {
            const branchName = row[HEADER_BRANCH_NAME];
            const employeeCode = row[HEADER_EMPLOYEE_CODE];

            if (!branchSummary[branchName]) {
                branchSummary[branchName] = {
                    employees: new Set(),
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0
                };
            }
            branchSummary[branchName].employees.add(employeeCode);
            branchSummary[branchName].calls += parseInt(row[HEADER_CALL] || 0);
            branchSummary[branchName].visits += parseInt(row[HEADER_VISIT] || 0);
            // The original script was parsing HEADER_REFERENCE here.
            // Since you clarified the spelling, it should be HEADER_REFERENCE which is 'Referance'
            branchSummary[branchName].references += parseInt(row[HEADER_REFERENCE] || 0);

            // *** Calculate New Customer Leads here for Branch Snapshot ***
            // This assumes 'Activity Type' and 'Type of Customer' columns exist in your CSV
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const customerType = row[HEADER_TYPE_OF_CUSTOMER];

            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                branchSummary[branchName].newCustomerLeads += 1;
            }
            // *** End Calculation ***

            employeeSet.add(employeeCode);
        });

        // Calculate overall totals
        let overallCalls = 0;
        let overallVisits = 0;
        let overallReferences = 0;
        let overallNewCustomerLeads = 0;

        allBranchSnapshotTableBody.innerHTML = '';
        for (const branchName of Object.keys(branchSummary).sort()) {
            const summary = branchSummary[branchName];
            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = summary.employees.size;
            row.insertCell().textContent = summary.calls;
            row.insertCell().textContent = summary.visits;
            row.insertCell().textContent = summary.references;
            row.insertCell().textContent = summary.newCustomerLeads;

            overallCalls += summary.calls;
            overallVisits += summary.visits;
            overallReferences += summary.references;
            overallNewCustomerLeads += summary.newCustomerLeads;
        }

        document.getElementById('totalBranchesParticipated').textContent = Object.keys(branchSummary).length;
        document.getElementById('overallEmployeesParticipated').textContent = employeeSet.size;
        document.getElementById('overallCalls').textContent = overallCalls;
        document.getElementById('overallVisits').textContent = overallVisits;
        document.getElementById('overallReferences').textContent = overallReferences;
        document.getElementById('overallNewCustomerLeads').textContent = overallNewCustomerLeads;
    }

    function renderAllStaffOverallPerformance(data) {
        const employeePerformance = {};

        data.forEach(row => {
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            if (!employeePerformance[employeeCode]) {
                // Find employee details from the 'employees' global array (which includes master data)
                const employeeDetail = employees.find(emp => emp.code === employeeCode);
                employeePerformance[employeeCode] = {
                    name: employeeDetail ? employeeDetail.name : row[HEADER_EMPLOYEE_NAME] || 'Unknown',
                    branch: employeeDetail ? employeeDetail.branch : row[HEADER_BRANCH_NAME] || 'Unknown',
                    designation: employeeDetail ? employeeDetail.designation : row[HEADER_DESIGNATION] || 'Other',
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0 // Initialize to 0 for calculation
                };
            }
            employeePerformance[employeeCode].calls += parseInt(row[HEADER_CALL] || 0);
            employeePerformance[employeeCode].visits += parseInt(row[HEADER_VISIT] || 0);
            employeePerformance[employeeCode].references += parseInt(row[HEADER_REFERENCE] || 0); // Uses your spelling

            // *** Calculate New Customer Leads here for All Staff Performance ***
            // This assumes 'Activity Type' and 'Type of Customer' columns exist in your CSV
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const customerType = row[HEADER_TYPE_OF_CUSTOMER];

            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                employeePerformance[employeeCode].newCustomerLeads += 1;
            }
            // *** End Calculation ***
        });

        allStaffPerformanceTableBody.innerHTML = '';
        const sortedEmployeeCodes = Object.keys(employeePerformance).sort((a, b) => {
            const empA = employeePerformance[a];
            const empB = employeePerformance[b];
            return empA.name.localeCompare(empB.name);
        });

        sortedEmployeeCodes.forEach(employeeCode => {
            const performance = employeePerformance[employeeCode];
            const designation = performance.designation || 'Other';
            const targets = TARGETS[designation] || TARGETS['Other'];

            const row = allStaffPerformanceTableBody.insertRow();
            const nameCell = row.insertCell();
            nameCell.textContent = performance.name;
            nameCell.classList.add('employee-name-cell');
            nameCell.dataset.employeeCode = employeeCode; // Store code for click event
            nameCell.dataset.employeeName = performance.name;
            nameCell.dataset.employeeBranch = performance.branch;
            nameCell.dataset.employeeDesignation = designation;

            row.insertCell().textContent = employeeCode;
            row.insertCell().textContent = performance.branch;
            row.insertCell().textContent = designation;
            row.insertCell().textContent = performance.calls;
            row.insertCell().textContent = performance.visits;
            row.insertCell().textContent = performance.references;
            row.insertCell().textContent = performance.newCustomerLeads; // Display calculated value
            row.insertCell().textContent = targets.Call;
            row.insertCell().textContent = targets.Visit;
            row.insertCell().textContent = targets.Referance; // Uses your spelling
            row.insertCell().textContent = targets['New Customer Leads'];
        });
    }

    function renderSingleEmployeePerformance(data, employeeCode) {
        // Filter data for the specific employee and also calculate New Customer Leads for each row
        const employeeDataWithLeads = data.filter(row => row[HEADER_EMPLOYEE_CODE] === employeeCode).map(row => {
            const newCustomerLead = ((row[HEADER_ACTIVITY_TYPE] === 'Visit' || row[HEADER_ACTIVITY_TYPE] === 'Calls') && row[HEADER_TYPE_OF_CUSTOMER] === 'New') ? 1 : 0;
            return {
                ...row, // Keep all original row data
                calculatedNewCustomerLead: newCustomerLead // Add the calculated lead
            };
        });

        // Find the employee's full details (name, branch, designation) from the global employees array
        const employeeDetail = employees.find(emp => emp.code === employeeCode);
        if (employeeDetail) {
            employeeNameDisplay.textContent = employeeDetail.name;
            employeeCodeDisplay.textContent = employeeDetail.code;
            employeeBranchDisplay.textContent = employeeDetail.branch;
            employeeDesignationDisplay.textContent = employeeDetail.designation || 'N/A';
        } else {
            // Fallback if employee detail not found in global list
            employeeNameDisplay.textContent = 'Unknown Employee';
            employeeCodeDisplay.textContent = employeeCode;
            employeeBranchDisplay.textContent = 'N/A';
            employeeDesignationDisplay.textContent = 'N/A';
        }

        singleEmployeePerformanceTableBody.innerHTML = '';
        // Sort by date descending, using the new array with calculated leads
        employeeDataWithLeads.sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP]));

        employeeDataWithLeads.forEach(row => {
            const rowElement = singleEmployeePerformanceTableBody.insertRow();
            rowElement.insertCell().textContent = new Date(row[HEADER_TIMESTAMP]).toLocaleDateString();
            rowElement.insertCell().textContent = row[HEADER_CALL] || 0;
            rowElement.insertCell().textContent = row[HEADER_VISIT] || 0;
            rowElement.insertCell().textContent = row[HEADER_REFERENCE] || 0; // Uses your spelling
            rowElement.insertCell().textContent = row.calculatedNewCustomerLead; // Display the calculated value
            rowElement.insertCell().textContent = row[HEADER_REMARKS] || 'N/A';
        });
    }

    function renderNonParticipatingBranches(allCanvassingData) {
        const participatingBranches = new Set();
        allCanvassingData.forEach(row => {
            if (row[HEADER_BRANCH_NAME]) {
                participatingBranches.add(row[HEADER_BRANCH_NAME]);
            }
        });

        const allBranches = new Set(branches); // Use the globally collected branches
        const nonParticipating = Array.from(allBranches).filter(branch => !participatingBranches.has(branch)).sort();

        nonParticipatingBranchesList.innerHTML = '';
        if (nonParticipating.length > 0) {
            noParticipationMessageContainer.style.display = 'block';
            document.getElementById('noParticipationMessage').textContent = 'The following branches have no participation records:';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        } else {
            noParticipationMessageContainer.style.display = 'block';
            document.getElementById('noParticipationMessage').textContent = 'All recorded branches have participation records.';
            nonParticipatingBranchesList.innerHTML = ''; // Clear list if all participate
        }
    }


    // --- Helper Functions ---
    function displayMessage(message, isError = false, autoHide = true) {
        statusMessageElement.textContent = message;
        statusMessageElement.classList.remove('success', 'error');
        if (isError) {
            statusMessageElement.classList.add('error');
        } else {
            statusMessageElement.classList.add('success');
        }
        statusMessageElement.style.display = 'block';

        if (autoHide) {
            setTimeout(() => {
                statusMessageElement.style.display = 'none';
            }, 5000); // Hide after 5 seconds
        }
    }

    function updateDateRangeDisplay(startDate, endDate) {
        if (startDate && endDate) {
            selectedDateRange.textContent = `${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`;
            dateRangeDisplay.style.display = 'flex';
        } else if (startDate) {
            selectedDateRange.textContent = `From ${startDate.toLocaleDateString()}`;
            dateRangeDisplay.style.display = 'flex';
        } else if (endDate) {
            selectedDateRange.textContent = `To ${endDate.toLocaleDateString()}`;
            dateRangeDisplay.style.display = 'flex';
        } else {
            dateRangeDisplay.style.display = 'none';
        }
    }

    // --- Google Apps Script Communication ---
    async function sendDataToGoogleAppsScript(action, data) {
        displayMessage('Processing request...', false, false);
        try {
            const formData = new FormData();
            formData.append('action', action);
            formData.append('data', JSON.stringify(data));

            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            console.log('Apps Script Response:', result);

            if (result.status === 'success') {
                displayMessage(result.message || 'Operation successful!', false);
                return result; // Return the full result object
            } else {
                displayMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayMessage(`Communication error: ${error.message}`, true);
            return false;
        }
    }

    function displayEmployeeManagementMessage(message, isError) {
        const messageElement = document.getElementById('employeeManagementMessage'); // Assuming you add this div
        if (!messageElement) {
            // Fallback to general status message if specific one not found
            displayMessage(message, isError);
            return;
        }
        messageElement.textContent = message;
        messageElement.classList.remove('success', 'error');
        if (isError) {
            messageElement.classList.add('error');
        } else {
            messageElement.classList.add('success');
        }
        messageElement.style.display = 'block';
        setTimeout(() => {
            messageElement.style.display = 'none';
        }, 5000);
    }


    // --- Event Listeners ---

    // Tab button click events
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));


    // Filter change events
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            const employeesInBranch = employees.filter(emp => emp.branch === selectedBranch);
            populateEmployeeFilter(employeesInBranch);
        } else {
            populateEmployeeFilter(employees); // Show all employees if no branch selected
        }
        employeeSelect.value = ''; // Reset employee selection when branch changes
        renderReports();
    });

    employeeSelect.addEventListener('change', renderReports);

    applyDateFilterBtn.addEventListener('click', renderReports);
    clearDateFilterBtn.addEventListener('click', () => {
        dateFilterStart.value = '';
        dateFilterEnd.value = '';
        renderReports();
    });

    // Delegated event listener for employee name clicks in All Staff Performance table
    if (allStaffPerformanceTableBody) {
        allStaffPerformanceTableBody.addEventListener('click', (event) => {
            const targetCell = event.target.closest('.employee-name-cell');
            if (targetCell) {
                const employeeCode = targetCell.dataset.employeeCode;
                // Set the employeeSelect to the clicked employee
                employeeSelect.value = employeeCode;
                // Switch to the overall performance tab and render single employee view
                showTab('allStaffOverallPerformanceTabBtn');
                renderReports(); // This will trigger renderSingleEmployeePerformance because employeeSelect has a value
            }
        });
    }


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: employeeBranchInput.value.trim(),
                [HEADER_DESIGNATION]: employeeDesignationInput.value.trim()
            };

            // Basic validation
            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH_NAME] || !employeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const employeeLines = bulkDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Expect at least Name,Code
                    console.warn(`Skipping malformed bulk entry line: ${line}`);
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
                    bulkAddEmployeeForm.reset(); // Clear form after submission
                    fetchData(); // Refresh data
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    // Existing: Event Listener for Delete Employee Form
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
                deleteEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    fetchData();
    showTab('allBranchSnapshotTabBtn');
});
