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
    // Removed direct headers for Visit, Call, Referance as they are derived from Activity Type
    const HEADER_NEW_CUSTOMER_LEADS = 'New Customer Leads'; // This will be calculated in JS
    const HEADER_REMARKS = 'Remarks';

    // *** NEW HEADERS FOR CALCULATION (MUST MATCH YOUR CSV COLUMN NAMES EXACTLY) ***
    // Based on your clarification that 'Activity Type' column exists.
    const HEADER_ACTIVITY_TYPE = 'Activity Type'; // Make sure this matches your actual column name (e.g., 'ActivityType', 'Activity_Type')
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer'; // Make sure this matches your actual column name

    // --- Headers for Master Employee Data (from Liv).xlsx - Sheet1.csv) ---
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code';
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name';
    const MASTER_HEADER_BRANCH_NAME_MASTER = 'Branch Name'; // Renamed to avoid conflict with activity data branch name
    const MASTER_HEADER_DESIGNATION = 'Designation'; // Assuming master data also has designation


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
    let masterEmployeeData = [];
    let branches = []; // Array of normalized branch names
    let employees = []; // Array of canonical employee objects (unique by code)

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
        // Use a Set for unique branch names (normalized)
        const uniqueBranchesSet = new Set();
        // Use a Map to store canonical employee details, keyed by Employee Code
        const employeeMap = new Map();

        // 1. Populate employeeMap with master data first (highest priority for employee details)
        masterEmployeeData.forEach(emp => {
            const employeeCode = emp[MASTER_HEADER_EMPLOYEE_CODE];
            const employeeName = emp[MASTER_HEADER_EMPLOYEE_NAME];
            const branchName = emp[MASTER_HEADER_BRANCH_NAME_MASTER];
            const designation = emp[MASTER_HEADER_DESIGNATION] || 'Other'; // Assuming master data has Designation

            if (branchName) {
                uniqueBranchesSet.add(branchName.trim().toLowerCase()); // Normalize branch name
            }

            if (employeeCode) {
                employeeMap.set(employeeCode, {
                    code: employeeCode,
                    name: employeeName || 'Unknown',
                    branch: branchName || 'Unknown',
                    designation: designation
                });
            }
        });

        // 2. Process canvassing data
        canvassingData.forEach(row => {
            const branchName = row[HEADER_BRANCH_NAME];
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const employeeName = row[HEADER_EMPLOYEE_NAME];
            const designation = row[HEADER_DESIGNATION] || 'Other';

            if (branchName) {
                uniqueBranchesSet.add(branchName.trim().toLowerCase()); // Normalize branch name
            }

            if (employeeCode) {
                // If employee code already in map (from master data), do not overwrite details
                if (!employeeMap.has(employeeCode)) {
                    employeeMap.set(employeeCode, {
                        code: employeeCode,
                        name: employeeName || 'Unknown',
                        branch: branchName || 'Unknown',
                        designation: designation
                    });
                }
            }
        });

        // Convert the Set of normalized branches to a sorted Array
        branches = Array.from(uniqueBranchesSet).sort();
        // Convert the Map values to a sorted Array of canonical employee objects
        employees = Array.from(employeeMap.values()).sort((a, b) => a.name.localeCompare(b.name));

        populateFilters();
        renderReports();
    }


    function populateFilters() {
        // Populate Branch Select
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        // The 'branches' array now contains normalized (lowercase) names
        branches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); // Capitalize for display
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
            // Use the canonical name and branch from the 'employees' array
            option.textContent = `${employee.name} (${employee.branch})`;
            employeeSelect.appendChild(option);
        });
    }

    // --- Report Rendering ---
    function renderReports() {
        const currentBranch = branchSelect.value; // This will be the normalized lowercase name
        const currentEmployeeCode = employeeSelect.value;
        const startDate = dateFilterStart.value ? new Date(dateFilterStart.value) : null;
        const endDate = dateFilterEnd.value ? new Date(dateFilterEnd.value) : null;

        // Filter data based on selections and dates
        const filteredData = canvassingData.filter(row => {
            let matchesBranch = true;
            let matchesEmployee = true;
            let matchesDate = true;

            // Compare filtered branch (normalized) with row's branch (normalized)
            if (currentBranch && row[HEADER_BRANCH_NAME].trim().toLowerCase() !== currentBranch) {
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
            const branchNameNormalized = row[HEADER_BRANCH_NAME].trim().toLowerCase(); // Use normalized name for grouping
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const activityType = row[HEADER_ACTIVITY_TYPE];

            if (!branchSummary[branchNameNormalized]) {
                branchSummary[branchNameNormalized] = {
                    employees: new Set(),
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0
                };
            }
            branchSummary[branchNameNormalized].employees.add(employeeCode);

            // Increment based on Activity Type
            if (activityType === 'Calls') {
                branchSummary[branchNameNormalized].calls += 1;
            } else if (activityType === 'Visit') {
                branchSummary[branchNameNormalized].visits += 1;
            } else if (activityType === 'Referance') {
                branchSummary[branchNameNormalized].references += 1;
            }

            // Calculate New Customer Leads
            const customerType = row[HEADER_TYPE_OF_CUSTOMER];
            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                branchSummary[branchNameNormalized].newCustomerLeads += 1;
            }

            employeeSet.add(employeeCode);
        });

        // Calculate overall totals
        let overallCalls = 0;
        let overallVisits = 0;
        let overallReferences = 0;
        let overallNewCustomerLeads = 0;

        allBranchSnapshotTableBody.innerHTML = '';
        for (const branchNameKey of Object.keys(branchSummary).sort()) { // Sort by normalized key
            const summary = branchSummary[branchNameKey];
            const row = allBranchSnapshotTableBody.insertRow();
            // Display the branch name with capitalized first letters for better readability
            row.insertCell().textContent = branchNameKey.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
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
            const activityType = row[HEADER_ACTIVITY_TYPE];

            if (!employeePerformance[employeeCode]) {
                // Find employee details from the 'employees' global array (which now has canonical data)
                const employeeDetail = employees.find(emp => emp.code === employeeCode);
                employeePerformance[employeeCode] = {
                    name: employeeDetail ? employeeDetail.name : row[HEADER_EMPLOYEE_NAME] || 'Unknown',
                    branch: employeeDetail ? employeeDetail.branch : row[HEADER_BRANCH_NAME] || 'Unknown',
                    designation: employeeDetail ? employeeDetail.designation : row[HEADER_DESIGNATION] || 'Other',
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0
                };
            }

            // Increment based on Activity Type
            if (activityType === 'Calls') {
                employeePerformance[employeeCode].calls += 1;
            } else if (activityType === 'Visit') {
                employeePerformance[employeeCode].visits += 1;
            } else if (activityType === 'Referance') {
                employeePerformance[employeeCode].references += 1;
            }

            // Calculate New Customer Leads
            const customerType = row[HEADER_TYPE_OF_CUSTOMER];
            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                employeePerformance[employeeCode].newCustomerLeads += 1;
            }
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
            row.insertCell().textContent = performance.newCustomerLeads;
            row.insertCell().textContent = targets.Call;
            row.insertCell().textContent = targets.Visit;
            row.insertCell().textContent = targets.Referance;
            row.insertCell().textContent = targets['New Customer Leads'];
        });
    }

    function renderSingleEmployeePerformance(data, employeeCode) {
        // Filter data for the specific employee and also calculate all metrics for each row
        const employeeDataWithMetrics = data.filter(row => row[HEADER_EMPLOYEE_CODE] === employeeCode).map(row => {
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const customerType = row[HEADER_TYPE_OF_CUSTOMER];

            let calculatedCalls = 0;
            let calculatedVisits = 0;
            let calculatedReferences = 0;
            let calculatedNewCustomerLead = 0;

            // Increment based on Activity Type for individual rows
            if (activityType === 'Calls') {
                calculatedCalls = 1;
            } else if (activityType === 'Visit') {
                calculatedVisits = 1;
            } else if (activityType === 'Referance') {
                calculatedReferences = 1;
            }

            // Calculate New Customer Leads for individual rows
            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                calculatedNewCustomerLead = 1;
            }

            return {
                ...row, // Keep all original row data
                calculatedCalls: calculatedCalls,
                calculatedVisits: calculatedVisits,
                calculatedReferences: calculatedReferences,
                calculatedNewCustomerLead: calculatedNewCustomerLead
            };
        });

        // Find the employee's full details (name, branch, designation) from the global employees array
        // This 'employees' array now contains the canonical details.
        const employeeDetail = employees.find(emp => emp.code === employeeCode);
        if (employeeDetail) {
            employeeNameDisplay.textContent = employeeDetail.name;
            employeeCodeDisplay.textContent = employeeDetail.code;
            employeeBranchDisplay.textContent = employeeDetail.branch;
            employeeDesignationDisplay.textContent = employeeDetail.designation || 'N/A';
        } else {
            // Fallback if employee detail not found (should be rare with new logic)
            employeeNameDisplay.textContent = 'Unknown Employee';
            employeeCodeDisplay.textContent = employeeCode;
            employeeBranchDisplay.textContent = 'N/A';
            employeeDesignationDisplay.textContent = 'N/A';
        }

        singleEmployeePerformanceTableBody.innerHTML = '';
        // Sort by date descending, using the new array with calculated leads
        employeeDataWithMetrics.sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP]));

        employeeDataWithMetrics.forEach(row => {
            const rowElement = singleEmployeePerformanceTableBody.insertRow();
            rowElement.insertCell().textContent = new Date(row[HEADER_TIMESTAMP]).toLocaleDateString();
            rowElement.insertCell().textContent = row.calculatedCalls;
            rowElement.insertCell().textContent = row.calculatedVisits;
            rowElement.insertCell().textContent = row.calculatedReferences;
            rowElement.insertCell().textContent = row.calculatedNewCustomerLead; // Display the calculated value
            rowElement.insertCell().textContent = row[HEADER_REMARKS] || 'N/A';
        });
    }

    function renderNonParticipatingBranches(allCanvassingData) {
        const participatingBranchesNormalized = new Set();
        allCanvassingData.forEach(row => {
            if (row[HEADER_BRANCH_NAME]) {
                participatingBranchesNormalized.add(row[HEADER_BRANCH_NAME].trim().toLowerCase());
            }
        });

        // Use the globally collected and normalized branches for comparison
        const allBranchesNormalized = new Set(branches); // 'branches' already contains normalized names
        const nonParticipating = Array.from(allBranchesNormalized).filter(branch => !participatingBranchesNormalized.has(branch)).sort();

        nonParticipatingBranchesList.innerHTML = '';
        if (nonParticipating.length > 0) {
            noParticipationMessageContainer.style.display = 'block';
            document.getElementById('noParticipationMessage').textContent = 'The following branches have no participation records:';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); // Capitalize for display
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
        const selectedBranchNormalized = branchSelect.value; // This is already normalized
        if (selectedBranchNormalized) {
            // Filter employees based on their normalized branch name
            const employeesInBranch = employees.filter(emp => emp.branch.trim().toLowerCase() === selectedBranchNormalized);
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
