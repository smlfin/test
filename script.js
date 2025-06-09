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
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Sales Executive': {
            'Visit': 5 * MONTHLY_WORKING_DAYS,
            'Call': 10 * MONTHLY_WORKING_DAYS,
            'Reference': 2 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 40
        }
    };

    // --- Headers for Activity Data (from DATA_URL) ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_COUNT = 'Count';
    const HEADER_REMARKS = 'Remarks';

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED..."; // This is intentionally unused as master data is now hardcoded

    // --- Headers for Master Employee Data (from Liv).xlsx - Sheet1.csv) ---
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code';
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name';
    const MASTER_HEADER_BRANCH_NAME_MASTER = 'Branch Name'; // Renamed to avoid conflict with activity data branch name
    const MASTER_HEADER_DESIGNATION = 'Designation'; // Assuming designation is also part of master data

    // Hardcoded Master Employee Data
    // IMPORTANT: Populate this array with your actual employee and branch details.
    // Ensure 'Employee Code' is unique for deduplication.
    const HARDCODED_MASTER_EMPLOYEES = [
        {
            [MASTER_HEADER_EMPLOYEE_NAME]: 'Alice',
            [MASTER_HEADER_EMPLOYEE_CODE]: 'EMP001',
            [MASTER_HEADER_BRANCH_NAME_MASTER]: 'Main Branch',
            [MASTER_HEADER_DESIGNATION]: 'Manager'
        },
        {
            [MASTER_HEADER_EMPLOYEE_NAME]: 'Bob',
            [MASTER_HEADER_EMPLOYEE_CODE]: 'EMP002',
            [MASTER_HEADER_BRANCH_NAME_MASTER]: 'North Branch',
            [MASTER_HEADER_DESIGNATION]: 'Sales Executive'
        },
        {
            [MASTER_HEADER_EMPLOYEE_NAME]: 'Charlie',
            [MASTER_HEADER_EMPLOYEE_CODE]: 'EMP003',
            [MASTER_HEADER_BRANCH_NAME_MASTER]: 'Main Branch',
            [MASTER_HEADER_DESIGNATION]: 'Investment Staff'
        },
        {
            [MASTER_HEADER_EMPLOYEE_NAME]: 'David',
            [MASTER_HEADER_EMPLOYEE_CODE]: 'EMP004',
            [MASTER_HEADER_BRANCH_NAME_MASTER]: 'South Branch',
            [MASTER_HEADER_DESIGNATION]: 'Sales Executive'
        },
        {
            [MASTER_HEADER_EMPLOYEE_NAME]: 'Eve',
            [MASTER_HEADER_EMPLOYEE_CODE]: 'EMP005',
            [MASTER_HEADER_BRANCH_NAME_MASTER]: 'North Branch',
            [MASTER_HEADER_DESIGNATION]: 'Manager'
        }
        // Add more employee data here following the same structure
    ];

    // Function to fetch (now, return hardcoded) Master Employee Data
    async function fetchMasterEmployeeData() {
        try {
            // Since master employee data is now hardcoded in the script,
            // we simply return the predefined array.
            return HARDCODED_MASTER_EMPLOYEES;
        } catch (error) {
            console.error("Error retrieving hardcoded master employee data:", error);
            displayEmployeeManagementMessage('Failed to load hardcoded master employee data. Please check the console for details.', true);
            return []; // Return empty array on error
        }
    }


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    const applyFilterBtn = document.getElementById('applyFilterBtn');
    const reportsSection = document.getElementById('reportsSection');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewOptions = document.getElementById('viewOptions');
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn'); // Restored tab
    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const allStaffOverallPerformanceTableBody = document.getElementById('allStaffOverallPerformanceTableBody');
    const employeeDetailsPerformanceBody = document.getElementById('employeeDetailsPerformanceBody');
    const totalVisitsElement = document.getElementById('totalVisits');
    const totalCallsElement = document.getElementById('totalCalls');
    const totalReferencesElement = document.getElementById('totalReferences');
    const totalNewCustomerLeadsElement = document.getElementById('totalNewCustomerLeads');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const statusMessage = document.getElementById('statusMessage'); // Message container
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const noParticipationMessage = document.getElementById('noParticipationMessage');
    const nonParticipatingBranchList = document.getElementById('nonParticipatingBranchList');


    // *** Data Variables ***
    let allCanvassingData = [];
    let masterEmployees = []; // This will now hold the hardcoded data
    let uniqueBranches = new Set();
    let uniqueEmployeesMap = new Map(); // Key: Employee Code, Value: Employee Name (for deduplication)


    // *** Helper Functions ***

    function displayEmployeeManagementMessage(message, isError) {
        statusMessage.textContent = message;
        statusMessage.className = isError ? 'message-container error' : 'message-container success';
        statusMessage.style.display = 'block';
        setTimeout(() => {
            statusMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }


    async function fetchData() {
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvData = await response.text();
            return Papa.parse(csvData, { header: true }).data;
        } catch (error) {
            console.error("Failed to fetch canvassing data:", error);
            displayEmployeeManagementMessage('Failed to load canvassing data. Please check the console for details.', true);
            return []; // Return empty array on error
        }
    }

    // Function to populate dropdowns
    function populateBranchDropdown(branches) {
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        branches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    function populateEmployeeDropdown(employees) {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        employees.forEach(employee => {
            const option = document.createElement('option');
            option.value = employee[HEADER_EMPLOYEE_CODE];
            option.textContent = `${employee[HEADER_EMPLOYEE_NAME]} (${employee[HEADER_EMPLOYEE_CODE]})`;
            employeeSelect.appendChild(option);
        });
    }

    // Function to filter data by date range
    function filterDataByDateRange(data, startDate, endDate) {
        const start = startDate ? new Date(startDate) : null;
        const end = endDate ? new Date(endDate) : null;

        if (!start && !end) {
            return data;
        }

        return data.filter(row => {
            const rowDate = new Date(row[HEADER_TIMESTAMP]);
            // Compare only date parts, ignore time for consistent filtering
            const rowDateOnly = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());

            let isAfterStart = true;
            if (start) {
                const startDateOnly = new Date(start.getFullYear(), start.getMonth(), start.getDate());
                isAfterStart = rowDateOnly >= startDateOnly;
            }

            let isBeforeEnd = true;
            if (end) {
                const endDateOnly = new Date(end.getFullYear(), end.getMonth(), end.getDate());
                isBeforeEnd = rowDateOnly <= endDateOnly;
            }

            return isAfterStart && isBeforeEnd;
        });
    }


    // Date range defaulting
    function setDefaultDateRange() {
        const today = new Date();
        const firstDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

        // Format dates as YYYY-MM-DD for input fields
        const formatDate = (date) => {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        };

        startDateInput.value = formatDate(firstDayOfMonth);
        endDateInput.value = formatDate(today);
    }


    function calculatePerformance(data, type, employeeDesignation) {
        const sum = data.reduce((acc, row) => {
            return acc + (row[HEADER_ACTIVITY_TYPE] === type ? parseInt(row[HEADER_COUNT] || 0) : 0);
        }, 0);

        const target = TARGETS[employeeDesignation] ? TARGETS[employeeDesignation][type] : 0;
        const percentage = target > 0 ? ((sum / target) * 100).toFixed(2) : (sum > 0 ? '100.00' : '0.00');

        return { sum, percentage };
    }

    function generatePerformanceSummary(filteredData, employeeDesignation) {
        const summary = {};
        for (const activityType in TARGETS[employeeDesignation]) {
            const { sum, percentage } = calculatePerformance(filteredData, activityType, employeeDesignation);
            summary[activityType] = { sum, percentage };
        }
        return summary;
    }


    function updateAllBranchSnapshotTable(filteredData) {
        allBranchSnapshotTableBody.innerHTML = '';
        const branchData = new Map(); // Key: Branch Name, Value: { totalVisits, totalCalls, totalReferences, totalNewCustomerLeads, employees: Set }

        // Initialize branch data from master list to ensure all branches are listed, even if no activity
        masterEmployees.forEach(emp => {
            const branchName = emp[MASTER_HEADER_BRANCH_NAME_MASTER];
            if (!branchData.has(branchName)) {
                branchData.set(branchName, {
                    totalVisits: 0,
                    totalCalls: 0,
                    totalReferences: 0,
                    totalNewCustomerLeads: 0,
                    employees: new Set()
                });
            }
        });


        filteredData.forEach(row => {
            const branch = row[HEADER_BRANCH_NAME];
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const count = parseInt(row[HEADER_COUNT] || 0);

            if (!branchData.has(branch)) {
                branchData.set(branch, {
                    totalVisits: 0,
                    totalCalls: 0,
                    totalReferences: 0,
                    totalNewCustomerLeads: 0,
                    employees: new Set()
                });
            }

            const currentBranch = branchData.get(branch);
            if (activityType === 'Visit') currentBranch.totalVisits += count;
            if (activityType === 'Call') currentBranch.totalCalls += count;
            if (activityType === 'Reference') currentBranch.totalReferences += count;
            if (activityType === 'New Customer Leads') currentBranch.totalNewCustomerLeads += count;
            currentBranch.employees.add(employeeCode); // Track employees with activity
        });

        // Sort branches alphabetically
        const sortedBranches = Array.from(branchData.keys()).sort();

        sortedBranches.forEach(branchName => {
            const data = branchData.get(branchName);
            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell(0).textContent = branchName;
            row.insertCell(1).textContent = data.employees.size; // Number of unique employees in this branch with activity
            row.insertCell(2).textContent = data.totalVisits;
            row.insertCell(3).textContent = data.totalCalls;
            row.insertCell(4).textContent = data.totalReferences;
            row.insertCell(5).textContent = data.totalNewCustomerLeads;
        });

        if (sortedBranches.length === 0) {
            const row = allBranchSnapshotTableBody.insertRow();
            const cell = row.insertCell(0);
            cell.colSpan = 6;
            cell.textContent = 'No data available for the selected period.';
            cell.style.textAlign = 'center';
            cell.style.padding = '20px';
        }
    }


    function updateAllStaffOverallPerformanceTable(filteredData) {
        allStaffOverallPerformanceTableBody.innerHTML = '';
        const employeePerformance = new Map(); // Key: Employee Code, Value: { name, branch, designation, activities: {type: sum} }

        // Initialize employee data from master list
        masterEmployees.forEach(emp => {
            const employeeCode = emp[MASTER_HEADER_EMPLOYEE_CODE];
            employeePerformance.set(employeeCode, {
                name: emp[MASTER_HEADER_EMPLOYEE_NAME],
                branch: emp[MASTER_HEADER_BRANCH_NAME_MASTER],
                designation: emp[MASTER_HEADER_DESIGNATION],
                activities: {
                    'Visit': 0,
                    'Call': 0,
                    'Reference': 0,
                    'New Customer Leads': 0
                }
            });
        });


        filteredData.forEach(row => {
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const count = parseInt(row[HEADER_COUNT] || 0);

            if (employeePerformance.has(employeeCode)) {
                employeePerformance.get(employeeCode).activities[activityType] += count;
            } else {
                // This case should ideally not happen if masterEmployees is comprehensive
                // but included for robustness if data exists for unlisted employees.
                console.warn(`Activity data found for unlisted employee code: ${employeeCode}`);
                employeePerformance.set(employeeCode, {
                    name: row[HEADER_EMPLOYEE_NAME] || 'N/A', // Use name from activity data if available
                    branch: row[HEADER_BRANCH_NAME] || 'N/A',
                    designation: row[HEADER_DESIGNATION] || 'N/A',
                    activities: {
                        'Visit': activityType === 'Visit' ? count : 0,
                        'Call': activityType === 'Call' ? count : 0,
                        'Reference': activityType === 'Reference' ? count : 0,
                        'New Customer Leads': activityType === 'New Customer Leads' ? count : 0
                    }
                });
            }
        });

        // Convert map to array and sort by employee name
        const sortedEmployees = Array.from(employeePerformance.values()).sort((a, b) => a.name.localeCompare(b.name));

        sortedEmployees.forEach(emp => {
            const row = allStaffOverallPerformanceTableBody.insertRow();
            row.insertCell(0).textContent = emp.name;
            row.insertCell(1).textContent = emp.branch;
            row.insertCell(2).textContent = emp.designation;
            row.insertCell(3).textContent = emp.activities['Visit'];
            row.insertCell(4).textContent = emp.activities['Call'];
            row.insertCell(5).textContent = emp.activities['Reference'];
            row.insertCell(6).textContent = emp.activities['New Customer Leads'];
        });

        if (sortedEmployees.length === 0) {
            const row = allStaffOverallPerformanceTableBody.insertRow();
            const cell = row.insertCell(0);
            cell.colSpan = 7;
            cell.textContent = 'No data available for the selected period.';
            cell.style.textAlign = 'center';
            cell.style.padding = '20px';
        }
    }


    function updateEmployeeDetailsPerformance(filteredData, selectedEmployeeCode) {
        employeeDetailsPerformanceBody.innerHTML = '';
        totalVisitsElement.textContent = '0';
        totalCallsElement.textContent = '0';
        totalReferencesElement.textContent = '0';
        totalNewCustomerLeadsElement.textContent = '0';

        const employeeData = masterEmployees.find(emp => emp[MASTER_HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);

        if (!employeeData) {
            // This scenario should ideally not happen if dropdown is populated from master data
            console.error("Employee not found in master data for selected code:", selectedEmployeeCode);
            return;
        }

        const employeeDesignation = employeeData[MASTER_HEADER_DESIGNATION];

        // Filter data for the specific employee
        const specificEmployeeData = filteredData.filter(row => row[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);

        if (specificEmployeeData.length === 0) {
            const row = employeeDetailsPerformanceBody.insertRow();
            const cell = row.insertCell(0);
            cell.colSpan = 4;
            cell.textContent = 'No activity data for this employee in the selected period.';
            cell.style.textAlign = 'center';
            cell.style.padding = '20px';
            return;
        }

        const summary = generatePerformanceSummary(specificEmployeeData, employeeDesignation);

        // Populate summary metrics
        totalVisitsElement.textContent = summary['Visit'].sum;
        totalCallsElement.textContent = summary['Call'].sum;
        totalReferencesElement.textContent = summary['Reference'].sum;
        totalNewCustomerLeadsElement.textContent = summary['New Customer Leads'].sum;

        // Populate detailed activity table
        for (const activityType in summary) {
            const row = employeeDetailsPerformanceBody.insertRow();
            row.insertCell(0).textContent = activityType;
            row.insertCell(1).textContent = summary[activityType].sum;
            row.insertCell(2).textContent = TARGETS[employeeDesignation][activityType];
            row.insertCell(3).textContent = summary[activityType].percentage + '%';
        }
    }

    function updateNonParticipatingBranches(allData, masterEmployees, startDate, endDate) {
        noParticipationMessage.style.display = 'none';
        nonParticipatingBranchList.innerHTML = '';

        const allBranchesFromMaster = new Set(masterEmployees.map(emp => emp[MASTER_HEADER_BRANCH_NAME_MASTER]));
        const participatingBranches = new Set();

        const filteredByDate = filterDataByDateRange(allData, startDate, endDate);

        filteredByDate.forEach(row => {
            participatingBranches.add(row[HEADER_BRANCH_NAME]);
        });

        const nonParticipating = [...allBranchesFromMaster].filter(
            branch => !participatingBranches.has(branch)
        ).sort();

        if (nonParticipating.length > 0) {
            noParticipationMessage.style.display = 'block';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchList.appendChild(li);
            });
        }
    }


    // Main function to process all data and update UI
    async function processData() {
        try {
            // Fetch canvassing data
            allCanvassingData = await fetchData();
            // Fetch master employee data (now from hardcoded source)
            masterEmployees = await fetchMasterEmployeeData();

            // Populate unique branches and employees from master data for dropdowns
            uniqueBranches = new Set();
            uniqueEmployeesMap.clear();

            masterEmployees.forEach(emp => {
                const branchName = emp[MASTER_HEADER_BRANCH_NAME_MASTER];
                const employeeCode = emp[MASTER_HEADER_EMPLOYEE_CODE];
                const employeeName = emp[MASTER_HEADER_EMPLOYEE_NAME];
                const designation = emp[MASTER_HEADER_DESIGNATION];

                if (branchName) {
                    uniqueBranches.add(branchName);
                }
                if (employeeCode) {
                    uniqueEmployeesMap.set(employeeCode, { name: employeeName, code: employeeCode, branch: branchName, designation: designation });
                }
            });

            // Convert uniqueEmployeesMap values to an array for dropdowns
            const employeesForDropdown = Array.from(uniqueEmployeesMap.values()).sort((a, b) => a.name.localeCompare(b.name));


            populateBranchDropdown(Array.from(uniqueBranches).sort());
            populateEmployeeDropdown(employeesForDropdown);

            // Set default date range on load
            setDefaultDateRange();

            // Initial display of all branch snapshot
            const initialStartDate = startDateInput.value;
            const initialEndDate = endDateInput.value;
            const filteredData = filterDataByDateRange(allCanvassingData, initialStartDate, initialEndDate);
            updateAllBranchSnapshotTable(filteredData);
            updateAllStaffOverallPerformanceTable(filteredData); // Also update overall performance initially
            updateNonParticipatingBranches(allCanvassingData, masterEmployees, initialStartDate, initialEndDate);

        } catch (error) {
            console.error('Error processing data:', error);
            displayEmployeeManagementMessage('An error occurred during data processing. See console for details.', true);
        }
    }


    // *** Event Listeners ***

    // Tab switching logic
    function showTab(tabId) {
        // Hide all sections
        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Remove 'active' class from all buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the selected section and add 'active' class to button
        if (tabId === 'allBranchSnapshotTabBtn') {
            document.getElementById('allBranchSnapshotSection').style.display = 'block';
            document.getElementById('allStaffOverallPerformanceSection').style.display = 'none';
            document.getElementById('nonParticipatingBranchesSection').style.display = 'none'; // Ensure this is hidden
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'none'; // Hide employee filter for branch snapshot
            viewOptions.style.display = 'flex'; // Show date filters for branch snapshot
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            document.getElementById('allBranchSnapshotSection').style.display = 'none';
            document.getElementById('allStaffOverallPerformanceSection').style.display = 'block';
            document.getElementById('nonParticipatingBranchesSection').style.display = 'none'; // Ensure this is hidden
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'flex'; // Show employee filter for staff performance
            viewOptions.style.display = 'flex'; // Show date filters
        } else if (tabId === 'nonParticipatingBranchesTabBtn') { // Handler for the restored tab
            document.getElementById('allBranchSnapshotSection').style.display = 'none';
            document.getElementById('allStaffOverallPerformanceSection').style.display = 'none';
            document.getElementById('nonParticipatingBranchesSection').style.display = 'block';
            reportsSection.style.display = 'block';
            employeeFilterPanel.style.display = 'none'; // No employee filter needed for this tab
            viewOptions.style.display = 'flex'; // Date filters are still relevant
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            reportsSection.style.display = 'none'; // Hide reports section
        }
        document.getElementById(tabId).classList.add('active');

        // Re-run data processing relevant to the tab to ensure fresh data
        if (tabId !== 'employeeManagementTabBtn') {
             const currentStartDate = startDateInput.value;
             const currentEndDate = endDateInput.value;
             const filteredData = filterDataByDateRange(allCanvassingData, currentStartDate, currentEndDate);

             if (tabId === 'allBranchSnapshotTabBtn') {
                 updateAllBranchSnapshotTable(filteredData);
             } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
                 const selectedEmployeeCode = employeeSelect.value;
                 if (selectedEmployeeCode) {
                    updateEmployeeDetailsPerformance(filteredData, selectedEmployeeCode);
                 } else {
                    updateAllStaffOverallPerformanceTable(filteredData);
                 }
             } else if (tabId === 'nonParticipatingBranchesTabBtn') {
                updateNonParticipatingBranches(allCanvassingData, masterEmployees, currentStartDate, currentEndDate);
             }
        }
    }


    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));


    // Apply filter button click
    applyFilterBtn.addEventListener('click', () => {
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;
        const startDate = startDateInput.value;
        const endDate = endDateInput.value;

        let filteredData = filterDataByDateRange(allCanvassingData, startDate, endDate);

        if (selectedBranch) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
        }

        const activeTab = document.querySelector('.tab-button.active').id;

        if (activeTab === 'allBranchSnapshotTabBtn') {
            updateAllBranchSnapshotTable(filteredData);
        } else if (activeTab === 'allStaffOverallPerformanceTabBtn') {
            if (selectedEmployeeCode) {
                updateEmployeeDetailsPerformance(filteredData, selectedEmployeeCode);
            } else {
                updateAllStaffOverallPerformanceTable(filteredData);
            }
        } else if (activeTab === 'nonParticipatingBranchesTabBtn') {
            updateNonParticipatingBranches(allCanvassingData, masterEmployees, startDate, endDate);
        }
    });

    // Branch selection changes
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        const startDate = startDateInput.value;
        const endDate = endDateInput.value;
        let filteredData = filterDataByDateRange(allCanvassingData, startDate, endDate);

        let employeesInSelectedBranch = [];
        if (selectedBranch) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
            // Filter master employees by selected branch
            employeesInSelectedBranch = masterEmployees.filter(emp => emp[MASTER_HEADER_BRANCH_NAME_MASTER] === selectedBranch);
        } else {
            // If no branch selected, show all employees from master list
            employeesInSelectedBranch = Array.from(uniqueEmployeesMap.values());
        }
        populateEmployeeDropdown(employeesInSelectedBranch);

        const activeTab = document.querySelector('.tab-button.active').id;
        if (activeTab === 'allBranchSnapshotTabBtn') {
            updateAllBranchSnapshotTable(filteredData);
        } else if (activeTab === 'allStaffOverallPerformanceTabBtn') {
             // If branch changes, reset employee selection and show overall staff performance
            employeeSelect.value = '';
            updateAllStaffOverallPerformanceTable(filteredData);
        }
    });


    // Employee selection changes
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        const startDate = startDateInput.value;
        const endDate = endDateInput.value;

        let filteredData = filterDataByDateRange(allCanvassingData, startDate, endDate);
        const selectedBranch = branchSelect.value; // Get current branch filter

        if (selectedBranch) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
        }

        if (selectedEmployeeCode) {
            updateEmployeeDetailsPerformance(filteredData, selectedEmployeeCode);
        } else {
            // If no employee selected, show overall staff performance for current branch filter
            updateAllStaffOverallPerformanceTable(filteredData);
        }
    });


    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(actionType, payload) {
        displayEmployeeManagementMessage('Processing request...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: actionType,
                    data: JSON.stringify(payload)
                }).toString()
            });
            const result = await response.json(); // Assuming the Apps Script returns JSON

            if (result.status === 'success') {
                displayEmployeeManagementMessage(result.message, false);
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error("Error sending data to Google Apps Script:", error);
            displayEmployeeManagementMessage('Failed to connect to the server. Please check your network.', true);
            return false;
        }
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required.', true);
                return;
            }

            const employeeLines = employeeDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Expect at least Name, Code
                    displayEmployeeManagementMessage(`Invalid format for employee entry: "${line}". Expected "Name,Code,Designation".`, true);
                    continue;
                }
                const employeeData = {
                    [HEADER_EMPLOYEE_NAME]: parts[0],
                    [HEADER_EMPLOYEE_CODE]: parts[1],
                    [HEADER_BRANCH_NAME]: branchName,
                    [HEADER_DESIGNATION]: parts[2] || '' // Designation is optional
                };
                employeesToAdd.push(employeeData);
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                    // After adding, refresh the data to update dropdowns
                    await processData(); // Re-process all data including master employees
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
                // After deletion, refresh the data to update dropdowns
                await processData(); // Re-process all data including master employees
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
