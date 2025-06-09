document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec";
    const MONTHLY_WORKING_DAYS = 22;

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': {
            'Visit': 30, // Increased target for Investment Staff
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Financial Advisor': { // Added Financial Advisor with custom Visit target
            'Visit': 15,
            'Call': 4 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 15
        },
        'Default': { // For any other designation not explicitly listed
            'Visit': 5,
            'Call': 2 * MONTHLY_WORKING_DAYS,
            'Reference': 0.5 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 10
        }
    };

    // --- Headers for Canvassing Data (from your CSV) ---
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_CALLS = 'Calls';
    const HEADER_VISITS = 'Visits';
    const HEADER_REFERENCES = 'References';
    const HEADER_NEW_CUSTOMER_LEADS = 'New Customer Leads';
    const HEADER_REMARKS = 'Remarks';

    // --- Headers for Master Employee Data (from Liv).xlsx - Sheet1.csv) ---
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code';
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name';
    const MASTER_HEADER_BRANCH_NAME_MASTER = 'Branch Name'; // Renamed to avoid conflict with activity data branch name

    // Global variables to hold data
    let masterEmployeeData = []; // To store data from Master Employees sheet
    let canvassingData = [];     // To store data from Canvassing Data sheet
    let allBranches = new Set(); // To store all unique branch names from master data

    // *** DOM Elements ***
    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const allStaffOverallPerformanceTableBody = document.getElementById('allStaffOverallPerformanceTableBody');
    const singleEmployeePerformanceTableBody = document.getElementById('singleEmployeePerformanceTableBody');

    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewOptions = document.getElementById('viewOptions');
    const dateFilterStart = document.getElementById('dateFilterStart'); // New Start Date filter
    const dateFilterEnd = document.getElementById('dateFilterEnd');     // New End Date filter
    const applyDateFilterBtn = document.getElementById('applyDateFilterBtn'); // New Apply Filter Button
    const clearDateFilterBtn = document.getElementById('clearDateFilterBtn'); // New Clear Filter Button
    const dateRangeDisplay = document.getElementById('dateRangeDisplay');
    const selectedDateRange = document.getElementById('selectedDateRange');


    const totalBranchesParticipated = document.getElementById('totalBranchesParticipated');
    const overallEmployeesParticipated = document.getElementById('overallEmployeesParticipated');
    const overallCalls = document.getElementById('overallCalls');
    const overallVisits = document.getElementById('overallVisits');
    const overallReferences = document.getElementById('overallReferences');
    const overallNewCustomerLeads = document.getElementById('overallNewCustomerLeads');

    const employeeNameDisplay = document.getElementById('employeeNameDisplay');
    const employeeCodeDisplay = document.getElementById('employeeCodeDisplay');
    const employeeBranchDisplay = document.getElementById('employeeBranchDisplay');
    const employeeDesignationDisplay = document.getElementById('employeeDesignationDisplay');

    const allBranchSnapshotContainer = document.getElementById('allBranchSnapshotContainer');
    const allStaffOverallPerformanceContainer = document.getElementById('allStaffOverallPerformanceContainer');
    const singleEmployeePerformanceContainer = document.getElementById('singleEmployeePerformanceContainer');
    const nonParticipatingBranchesContainer = document.getElementById('nonParticipatingBranchesContainer');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    const noParticipationMessageContainer = document.getElementById('noParticipationMessageContainer');
    const noParticipationMessage = document.getElementById('noParticipationMessage');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');

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

    const statusMessageContainer = document.getElementById('statusMessage');

    // *** Helper Functions ***

    function displayStatusMessage(message, isError = false) {
        statusMessageContainer.textContent = message;
        statusMessageContainer.className = `message-container ${isError ? 'error-message' : 'success-message'}`;
        statusMessageContainer.style.display = 'block';
        setTimeout(() => {
            statusMessageContainer.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        // Use the same statusMessage container for employee management messages
        displayStatusMessage(message, isError);
    }

    async function parseCsv(text) {
        const lines = text.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim());
        return lines.slice(1).map(line => {
            const values = line.split(',').map(value => value.trim());
            return headers.reduce((obj, header, index) => {
                obj[header] = values[index];
                return obj;
            }, {});
        });
    }

    // Function to fetch and parse Master Employee Data from Apps Script
    async function fetchMasterEmployeeData() {
        try {
            const response = await sendDataToGoogleAppsScript('get_master_employees', {});
            if (response && response.status === 'SUCCESS' && response.data) {
                masterEmployeeData = response.data;
                allBranches = new Set(masterEmployeeData.map(emp => emp[MASTER_HEADER_BRANCH_NAME_MASTER]));
                console.log("Master Employee Data fetched:", masterEmployeeData);
                console.log("All Branches from Master Data:", Array.from(allBranches));
            } else {
                console.error("Failed to fetch master employee data:", response.message || "Unknown error");
                displayStatusMessage(`Failed to load master employee data: ${response.message || 'Unknown error. Check Apps Script deployment.'}`, true);
                masterEmployeeData = []; // Ensure it's empty on failure
                allBranches = new Set();
            }
        } catch (error) {
            console.error("Error fetching master employee data:", error);
            displayStatusMessage(`Error loading master employee data: ${error.message}. Check browser console.`, true);
            masterEmployeeData = [];
            allBranches = new Set();
        }
    }


    // Function to fetch and parse Canvassing Data
    async function fetchCanvassingData() {
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const text = await response.text();
            canvassingData = await parseCsv(text);
            console.log("Canvassing Data fetched:", canvassingData);
        } catch (error) {
            console.error("Failed to fetch canvassing data:", error);
            displayStatusMessage(`Failed to load canvassing data: ${error.message}. Check DATA_URL.`, true);
            canvassingData = []; // Ensure it's empty on failure
        }
    }

    async function sendDataToGoogleAppsScript(action, data) {
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayStatusMessage(result.message);
                return { status: 'SUCCESS', message: result.message, data: result.data };
            } else {
                displayStatusMessage(result.message, true);
                return { status: 'ERROR', message: result.message };
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayStatusMessage(`Error: ${error.message}. Check WEB_APP_URL and Apps Script deployment.`, true);
            return { status: 'ERROR', message: error.message };
        }
    }

    function getTarget(designation, metric) {
        return (TARGETS[designation] && TARGETS[designation][metric]) ? TARGETS[designation][metric] : TARGETS['Default'][metric];
    }

    function formatDateForInput(date) {
        const d = new Date(date);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    // *** Reporting Functions ***

    function calculateOverallSummary(filteredData) {
        const summary = {
            totalBranchesParticipated: new Set(),
            totalEmployeesParticipated: new Set(),
            overallCalls: 0,
            overallVisits: 0,
            overallReferences: 0,
            overallNewCustomerLeads: 0
        };

        filteredData.forEach(entry => {
            summary.totalBranchesParticipated.add(entry[HEADER_BRANCH_NAME]);
            summary.totalEmployeesParticipated.add(entry[HEADER_EMPLOYEE_CODE]);
            summary.overallCalls += parseInt(entry[HEADER_CALLS] || 0);
            summary.overallVisits += parseInt(entry[HEADER_VISITS] || 0);
            summary.overallReferences += parseInt(entry[HEADER_REFERENCES] || 0);
            summary.overallNewCustomerLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        return {
            totalBranchesParticipated: summary.totalBranchesParticipated.size,
            overallEmployeesParticipated: summary.totalEmployeesParticipated.size,
            overallCalls: summary.overallCalls,
            overallVisits: summary.overallVisits,
            overallReferences: summary.overallReferences,
            overallNewCustomerLeads: summary.overallNewCustomerLeads
        };
    }

    function renderAllBranchSnapshot(data) {
        const branchSummary = {};
        data.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            if (!branchSummary[branchName]) {
                branchSummary[branchName] = {
                    employees: new Set(),
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0
                };
            }
            branchSummary[branchName].employees.add(entry[HEADER_EMPLOYEE_CODE]);
            branchSummary[branchName].calls += parseInt(entry[HEADER_CALLS] || 0);
            branchSummary[branchName].visits += parseInt(entry[HEADER_VISITS] || 0);
            branchSummary[branchName].references += parseInt(entry[HEADER_REFERENCES] || 0);
            branchSummary[branchName].newCustomerLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        allBranchSnapshotTableBody.innerHTML = '';
        for (const branchName in branchSummary) {
            const summary = branchSummary[branchName];
            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = summary.employees.size;
            row.insertCell().textContent = summary.calls;
            row.insertCell().textContent = summary.visits;
            row.insertCell().textContent = summary.references;
            row.insertCell().textContent = summary.newCustomerLeads;
        }

        const overall = calculateOverallSummary(data);
        totalBranchesParticipated.textContent = overall.totalBranchesParticipated;
        overallEmployeesParticipated.textContent = overall.overallEmployeesParticipated;
        overallCalls.textContent = overall.overallCalls;
        overallVisits.textContent = overall.overallVisits;
        overallReferences.textContent = overall.overallReferences;
        overallNewCustomerLeads.textContent = overall.overallNewCustomerLeads;
    }

    function renderAllStaffPerformance(data) {
        const employeePerformance = {};

        data.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (!employeePerformance[employeeCode]) {
                employeePerformance[employeeCode] = {
                    name: entry[HEADER_EMPLOYEE_NAME],
                    branch: entry[HEADER_BRANCH_NAME],
                    designation: entry[HEADER_DESIGNATION],
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0
                };
            }
            employeePerformance[employeeCode].calls += parseInt(entry[HEADER_CALLS] || 0);
            employeePerformance[employeeCode].visits += parseInt(entry[HEADER_VISITS] || 0);
            employeePerformance[employeeCode].references += parseInt(entry[HEADER_REFERENCES] || 0);
            employeePerformance[employeeCode].newCustomerLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        allStaffOverallPerformanceTableBody.innerHTML = '';
        for (const code in employeePerformance) {
            const emp = employeePerformance[code];
            const row = allStaffOverallPerformanceTableBody.insertRow();
            row.insertCell().textContent = emp.name;
            row.insertCell().textContent = code;
            row.insertCell().textContent = emp.branch;
            row.insertCell().textContent = emp.designation;
            row.insertCell().textContent = emp.calls;
            row.insertCell().textContent = emp.visits;
            row.insertCell().textContent = emp.references;
            row.insertCell().textContent = emp.newCustomerLeads;
            row.insertCell().textContent = getTarget(emp.designation, 'Call');
            row.insertCell().textContent = getTarget(emp.designation, 'Visit');
            row.insertCell().textContent = getTarget(emp.designation, 'Reference');
            row.insertCell().textContent = getTarget(emp.designation, 'New Customer Leads');
        }
    }

    function renderSingleEmployeePerformance(employeeCode, data) {
        const employeeData = data.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

        if (employeeData.length > 0) {
            const firstEntry = employeeData[0];
            employeeNameDisplay.textContent = firstEntry[HEADER_EMPLOYEE_NAME];
            employeeCodeDisplay.textContent = firstEntry[HEADER_EMPLOYEE_CODE];
            employeeBranchDisplay.textContent = firstEntry[HEADER_BRANCH_NAME];
            employeeDesignationDisplay.textContent = firstEntry[HEADER_DESIGNATION];
        } else {
            employeeNameDisplay.textContent = 'N/A';
            employeeCodeDisplay.textContent = employeeCode;
            employeeBranchDisplay.textContent = 'N/A';
            employeeDesignationDisplay.textContent = 'N/A';
        }

        singleEmployeePerformanceTableBody.innerHTML = '';
        employeeData.sort((a, b) => new Date(a[HEADER_DATE]) - new Date(b[HEADER_DATE])); // Sort by date
        employeeData.forEach(entry => {
            const row = singleEmployeePerformanceTableBody.insertRow();
            row.insertCell().textContent = entry[HEADER_DATE];
            row.insertCell().textContent = entry[HEADER_CALLS];
            row.insertCell().textContent = entry[HEADER_VISITS];
            row.insertCell().textContent = entry[HEADER_REFERENCES];
            row.insertCell().textContent = entry[HEADER_NEW_CUSTOMER_LEADS];
            row.insertCell().textContent = entry[HEADER_REMARKS];
        });
    }

    function renderNonParticipatingBranches() {
        nonParticipatingBranchesList.innerHTML = '';
        noParticipationMessageContainer.style.display = 'block';

        if (masterEmployeeData.length === 0) {
            noParticipationMessage.textContent = 'Master employee data not available. Please ensure it is fetched correctly.';
            return;
        }

        const participatingBranches = new Set(canvassingData.map(entry => entry[HEADER_BRANCH_NAME]));
        const nonParticipating = Array.from(allBranches).filter(branch => !participatingBranches.has(branch));

        if (nonParticipating.length > 0) {
            noParticipationMessage.textContent = 'The following branches have not participated in canvassing activities:';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        } else {
            noParticipationMessage.textContent = 'All branches have participated in canvassing activities!';
            noParticipationMessage.style.color = '#4CAF50'; // Green color for success
        }
    }


    // *** Data Processing and Filtering ***

    function filterData(branch, employeeCode, startDateStr, endDateStr) {
        let filtered = canvassingData;

        if (branch) {
            filtered = filtered.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
        }

        if (employeeCode) {
            filtered = filtered.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
        }

        if (startDateStr && endDateStr) {
            const startDate = new Date(startDateStr);
            const endDate = new Date(endDateStr);

            filtered = filtered.filter(entry => {
                const entryDate = new Date(entry[HEADER_DATE]);
                // Set hours to 0 for accurate date comparison without time component
                entryDate.setHours(0,0,0,0);
                startDate.setHours(0,0,0,0);
                endDate.setHours(0,0,0,0);
                return entryDate >= startDate && entryDate <= endDate;
            });

            selectedDateRange.textContent = `${formatDateForInput(startDate)} to ${formatDateForInput(endDate)}`;
            dateRangeDisplay.style.display = 'block';
        } else {
            dateRangeDisplay.style.display = 'none';
            selectedDateRange.textContent = '';
        }

        return filtered;
    }

    function processData() {
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value;
        const startDate = dateFilterStart.value;
        const endDate = dateFilterEnd.value;

        const currentActiveTab = document.querySelector('.tab-button.active').id;

        let filteredData;
        if (currentActiveTab === 'singleEmployeePerformanceTabBtn') {
             // For single employee, filter based on selected employee only
            filteredData = filterData(null, selectedEmployee, startDate, endDate);
        } else {
            // For other tabs, filter based on branch, and dates
            filteredData = filterData(selectedBranch, null, startDate, endDate);
        }

        if (currentActiveTab === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot(filteredData);
        } else if (currentActiveTab === 'allStaffOverallPerformanceTabBtn') {
            renderAllStaffPerformance(filteredData);
        } else if (currentActiveTab === 'singleEmployeePerformanceTabBtn') {
             renderSingleEmployeePerformance(selectedEmployee, filteredData);
        }
        // Non-participating branches tab does not use filtered data from canvassing data
    }

    // *** UI Handlers ***

    function showTab(tabId) {
        // Hide all report containers
        allBranchSnapshotContainer.style.display = 'none';
        allStaffOverallPerformanceContainer.style.display = 'none';
        singleEmployeePerformanceContainer.style.display = 'none';
        nonParticipatingBranchesContainer.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Hide filter controls by default
        employeeFilterPanel.style.display = 'none';
        viewOptions.style.display = 'none';
        dateRangeDisplay.style.display = 'none'; // Ensure date range display is hidden too

        // Remove 'active' class from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the selected tab and add 'active' class to its button
        document.getElementById(tabId).classList.add('active');

        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                allBranchSnapshotContainer.style.display = 'block';
                // No specific filters needed for this initial view, but still show dropdowns for future filtering
                branchSelect.value = ''; // Reset branch filter
                employeeSelect.value = ''; // Reset employee filter
                // Clear date filters
                dateFilterStart.value = '';
                dateFilterEnd.value = '';
                processData(); // Re-process data for this tab
                break;
            case 'allStaffOverallPerformanceTabBtn':
                allStaffOverallPerformanceContainer.style.display = 'block';
                employeeFilterPanel.style.display = 'block'; // Show employee filter for this tab
                viewOptions.style.display = 'flex'; // Show date filters
                branchSelect.value = ''; // Reset branch filter
                employeeSelect.value = ''; // Reset employee filter
                 // Clear date filters
                dateFilterStart.value = '';
                dateFilterEnd.value = '';
                processData(); // Re-process data for this tab
                break;
            case 'nonParticipatingBranchesTabBtn':
                nonParticipatingBranchesContainer.style.display = 'block';
                renderNonParticipatingBranches(); // This tab relies on masterEmployeeData and canvassingData
                break;
            case 'employeeManagementTabBtn':
                employeeManagementSection.style.display = 'block';
                break;
             case 'singleEmployeePerformanceTabBtn': // This tab is shown when an employee is selected
                singleEmployeePerformanceContainer.style.display = 'block';
                employeeFilterPanel.style.display = 'block'; // Show employee filter for this tab
                viewOptions.style.display = 'flex'; // Show date filters
                branchSelect.value = ''; // Reset branch filter
                 // Clear date filters
                dateFilterStart.value = '';
                dateFilterEnd.value = '';
                // Employee select value will be set by the event listener that triggers this view
                processData(); // Re-process data for this tab
                break;
        }
    }


    // *** Event Listeners ***

    // Tab buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));

    // Branch and Employee Selectors
    branchSelect.addEventListener('change', () => {
        // If a branch is selected, filter employees for that branch
        const selectedBranch = branchSelect.value;
        let employeesInBranch = [];
        if (selectedBranch) {
            employeesInBranch = masterEmployeeData.filter(emp => emp[MASTER_HEADER_BRANCH_NAME_MASTER] === selectedBranch);
        } else {
            employeesInBranch = masterEmployeeData; // All employees if no branch selected
        }

        // Populate employee dropdown
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        employeesInBranch.forEach(emp => {
            const option = document.createElement('option');
            option.value = emp[MASTER_HEADER_EMPLOYEE_CODE];
            option.textContent = `${emp[MASTER_HEADER_EMPLOYEE_NAME]} (${emp[MASTER_HEADER_DESIGNATION]})`;
            employeeSelect.appendChild(option);
        });

        // If the current active tab uses these filters, re-process data
        const currentActiveTab = document.querySelector('.tab-button.active').id;
        if (currentActiveTab === 'allStaffOverallPerformanceTabBtn' || currentActiveTab === 'singleEmployeePerformanceTabBtn') {
            processData();
        }
    });

    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // If an employee is selected, switch to single employee performance tab
            showTab('singleEmployeePerformanceTabBtn');
        } else {
            // If no employee selected, revert to All Staff Performance
            showTab('allStaffOverallPerformanceTabBtn');
        }
    });

    // New: Date filter buttons
    applyDateFilterBtn.addEventListener('click', processData); // Re-process data when apply button clicked
    clearDateFilterBtn.addEventListener('click', () => {
        dateFilterStart.value = '';
        dateFilterEnd.value = '';
        processData(); // Re-process data to clear filter
    });


    // Employee Management Forms
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [MASTER_HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [MASTER_HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [MASTER_HEADER_BRANCH_NAME_MASTER]: employeeBranchInput.value.trim(),
                [MASTER_HEADER_DESIGNATION]: employeeDesignationInput.value.trim()
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset();
                fetchData(); // Refresh all data including master employee data
            }
        });
    }

    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeeLines = employeeDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Name, Code, (Optional) Designation
                    const employeeData = {
                        [MASTER_HEADER_EMPLOYEE_NAME]: parts[0],
                        [MASTER_HEADER_EMPLOYEE_CODE]: parts[1],
                        [MASTER_HEADER_BRANCH_NAME_MASTER]: branchName,
                        [MASTER_HEADER_DESIGNATION]: parts[2] || '' // Designation is optional
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                    fetchData(); // Refresh all data including master employee data
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

            const deleteData = { [MASTER_HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
                fetchData(); // Refresh all data including master employee data
            }
        });
    }

    // Function to fetch all necessary data
    async function fetchData() {
        await Promise.all([
            fetchCanvassingData(),
            fetchMasterEmployeeData()
        ]);
        processData(); // Process and render data after both fetches are complete
        populateBranchDropdown(); // Populate branch dropdown after master data is fetched
    }

    function populateBranchDropdown() {
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        Array.from(allBranches).sort().forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }


    // Initial data fetch and tab display when the page loads
    fetchData();
    showTab('allBranchSnapshotTabBtn');
});
