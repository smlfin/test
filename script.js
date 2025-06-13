document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFD03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; 

    const MONTHLY_WORKING_DAYS = 22; // Assumed number of working days in a month for target calculations

    // Define targets for different designations
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
        'Seniors': { // Added Investment Staff with custom Visit target
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

    // Predefined list of branches for selection and reporting
    const PREDEFINED_BRANCHES = [
        'Central', 'North', 'South', 'East', 'West', 'Overseas'
    ];

    // CSV Headers (ensure these match your Google Sheet headers exactly)
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer';
    const HEADER_CUSTOMER_NAME = 'Customer Name';
    const HEADER_REMARKS = 'Remarks';

    // *** DOM Elements ***
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const branchDiagnosisTabBtn = document.getElementById('branchDiagnosisTabBtn'); // New Tab Button

    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const branchDiagnosisSection = document.getElementById('branchDiagnosisSection'); // New Section

    const statusMessageDiv = document.getElementById('statusMessage');
    const reportDisplay = document.getElementById('reportDisplay');
    const noParticipationMessage = document.getElementById('noParticipationMessage');

    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const applyDateFilterBtn = document.getElementById('applyDateFilterBtn');
    const clearDateFilterBtn = document.getElementById('clearDateFilterBtn');
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');

    const detailedCustomerBranchSelect = document.getElementById('detailedCustomerBranchSelect');
    const detailedCustomerEmployeeSelect = document.getElementById('detailedCustomerEmployeeSelect');
    const detailedCustomerStartDateInput = document.getElementById('detailedCustomerStartDate');
    const detailedCustomerEndDateInput = document.getElementById('detailedCustomerEndDate');
    const applyDetailedCustomerDateFilterBtn = document.getElementById('applyDetailedCustomerDateFilterBtn');
    const clearDetailedCustomerDateFilterBtn = document.getElementById('clearDetailedCustomerDateFilterBtn');
    const detailedCustomerViewDisplay = document.getElementById('detailedCustomerViewDisplay');
    const noDetailedCustomerDataMessage = document.getElementById('noDetailedCustomerDataMessage');

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


    // *** Global Variables for Data ***
    let allCanvassingData = []; // Stores all fetched canvassing data
    let employeeCodeToDesignationMap = new Map(); // Map employeeCode to Designation
    let employeeCodeToNameMap = new Map(); // Map employeeCode to Employee Name


    // *** Helper Functions ***

    function showStatusMessage(message, isError = false) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = isError ? 'message-container error' : 'message-container success';
        statusMessageDiv.style.display = 'block';
        setTimeout(() => {
            statusMessageDiv.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    function calculateTotalActivity(entries) {
        const activitySummary = {
            'Visit': 0,
            'Call': 0,
            'Reference': 0,
            'New Customer Leads': 0
        };
        entries.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            if (activitySummary.hasOwnProperty(activityType)) {
                activitySummary[activityType]++;
            }
        });
        return activitySummary;
    }

    function parseCSV(csv) {
        const lines = csv.split('\n');
        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (line === '') continue;

            const values = line.split(',').map(value => value.trim());
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

    // New Helper: Get number of working days passed in the current month
    function getWorkingDaysPassedThisMonth() {
        const today = new Date();
        const year = today.getFullYear();
        const month = today.getMonth(); // 0-indexed
        const currentDay = today.getDate();

        let workingDaysCount = 0;
        for (let day = 1; day <= currentDay; day++) {
            const date = new Date(year, month, day);
            const dayOfWeek = date.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
            if (dayOfWeek >= 1 && dayOfWeek <= 5) { // Monday to Friday
                workingDaysCount++;
            }
        }
        return workingDaysCount;
    }

    // New Helper: Get targets for an employee designation
    function calculateTargetForEmployee(designation) {
        return TARGETS[designation] || TARGETS['Default'];
    }


    // *** Data Processing ***
    async function processData() {
        showStatusMessage('Fetching data...');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            
            // Populate employeeCodeToDesignationMap and employeeCodeToNameMap
            allCanvassingData.forEach(entry => {
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                const designation = entry[HEADER_DESIGNATION];
                const employeeName = entry[HEADER_EMPLOYEE_NAME];
                if (employeeCode && designation && !employeeCodeToDesignationMap.has(employeeCode)) {
                    employeeCodeToDesignationMap.set(employeeCode, designation);
                }
                if (employeeCode && employeeName && !employeeCodeToNameMap.has(employeeCode)) {
                    employeeCodeToNameMap.set(employeeCode, employeeName);
                }
            });

            populateBranchDropdowns();
            populateEmployeeDropdowns();
            showStatusMessage('Data fetched successfully!', false);
        } catch (error) {
            console.error('Error fetching data:', error);
            showStatusMessage('Failed to fetch data: ' + error.message, true);
        }
    }

    function populateBranchDropdowns() {
        // Clear existing options
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        detailedCustomerBranchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';

        // Populate with predefined branches
        PREDEFINED_BRANCHES.forEach(branch => {
            const option1 = document.createElement('option');
            option1.value = branch;
            option1.textContent = branch;
            branchSelect.appendChild(option1);

            const option2 = document.createElement('option');
            option2.value = branch;
            option2.textContent = branch;
            detailedCustomerBranchSelect.appendChild(option2);
        });
    }

    function populateEmployeeDropdowns(branchName = '') {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        detailedCustomerEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';

        let filteredEmployeeCodes = new Set();
        let employeesInBranch = [];

        if (branchName) {
            employeesInBranch = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchName);
            employeesInBranch.forEach(entry => {
                filteredEmployeeCodes.add(entry[HEADER_EMPLOYEE_CODE]);
            });
        } else {
            // If no branch selected, show all employees from the map
            employeeCodeToNameMap.forEach((name, code) => filteredEmployeeCodes.add(code));
        }

        const sortedEmployeeCodes = Array.from(filteredEmployeeCodes).sort((a, b) => {
            const nameA = employeeCodeToNameMap.get(a) || '';
            const nameB = employeeCodeToNameMap.get(b) || '';
            return nameA.localeCompare(nameB);
        });

        sortedEmployeeCodes.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap.get(employeeCode);
            if (employeeName) { // Ensure employee name exists
                const option1 = document.createElement('option');
                option1.value = employeeCode;
                option1.textContent = `${employeeName} (${employeeCode})`;
                employeeSelect.appendChild(option1);

                const option2 = document.createElement('option');
                option2.value = employeeCode;
                option2.textContent = `${employeeName} (${employeeCode})`;
                detailedCustomerEmployeeSelect.appendChild(option2);
            }
        });
    }


    // *** Reporting Functions ***

    function filterCanvassingData(data, branch, employeeCode, startDate, endDate) {
        return data.filter(entry => {
            const entryDate = new Date(entry[HEADER_DATE]);
            const matchesBranch = branch === '' || entry[HEADER_BRANCH_NAME] === branch;
            const matchesEmployee = employeeCode === '' || entry[HEADER_EMPLOYEE_CODE] === employeeCode;
            const matchesStartDate = !startDate || entryDate >= startDate;
            const matchesEndDate = !endDate || entryDate <= endDate;
            return matchesBranch && matchesEmployee && matchesStartDate && matchesEndDate;
        });
    }

    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '';
        noParticipationMessage.style.display = 'none';
        employeeFilterPanel.style.display = 'none';

        const filteredData = filterCanvassingData(
            allCanvassingData,
            branchSelect.value,
            '', // Employee filter not used here
            startDateInput.value ? new Date(startDateInput.value) : null,
            endDateInput.value ? new Date(endDateInput.value) : null
        );

        const branchActivity = {};
        PREDEFINED_BRANCHES.forEach(branch => {
            branchActivity[branch] = calculateTotalActivity([]); // Initialize with zeros
        });

        filteredData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            if (branchActivity.hasOwnProperty(branch)) {
                branchActivity[branch][entry[HEADER_ACTIVITY_TYPE]]++;
            }
        });

        let tableHtml = `
            <table class="all-branch-snapshot-table">
                <thead>
                    <tr>
                        <th>Branch Name</th>
                        <th>Visits</th>
                        <th>Calls</th>
                        <th>References</th>
                        <th>New Customer Leads</th>
                    </tr>
                </thead>
                <tbody>
        `;

        let hasData = false;
        PREDEFINED_BRANCHES.forEach(branch => {
            const activities = branchActivity[branch];
            if (Object.values(activities).some(count => count > 0)) {
                hasData = true;
            }
            tableHtml += `
                <tr>
                    <td data-label="Branch Name">${branch}</td>
                    <td data-label="Visits">${activities['Visit']}</td>
                    <td data-label="Calls">${activities['Call']}</td>
                    <td data-label="References">${activities['Reference']}</td>
                    <td data-label="New Customer Leads">${activities['New Customer Leads']}</td>
                </tr>
            `;
        });

        tableHtml += `
                </tbody>
            </table>
        `;

        if (hasData) {
            reportDisplay.innerHTML = tableHtml;
        } else {
            noParticipationMessage.style.display = 'block';
        }
    }

    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '';
        noParticipationMessage.style.display = 'none';
        employeeFilterPanel.style.display = 'none';

        const filteredData = filterCanvassingData(
            allCanvassingData,
            '', // No specific branch filter for this report's initial data
            '', // Employee filter not used here
            startDateInput.value ? new Date(startDateInput.value) : null,
            endDateInput.value ? new Date(endDateInput.value) : null
        );

        const participatingBranches = new Set(filteredData.map(entry => entry[HEADER_BRANCH_NAME]));
        const nonParticipating = PREDEFINED_BRANCHES.filter(branch => !participatingBranches.has(branch));

        if (nonParticipating.length > 0) {
            let listHtml = `
                <h2>Non-Participating Branches</h2>
                <p>The following branches had no canvassing activity during the selected period:</p>
                <ul>
            `;
            nonParticipating.forEach(branch => {
                listHtml += `<li>${branch}</li>`;
            });
            listHtml += `</ul>`;
            reportDisplay.innerHTML = listHtml;
        } else {
            reportDisplay.innerHTML = '<p>All branches had activity during the selected period!</p>';
        }
    }

    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '';
        noParticipationMessage.style.display = 'none';
        employeeFilterPanel.style.display = 'block'; // Show employee filter for this report

        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;

        const filteredData = filterCanvassingData(
            allCanvassingData,
            selectedBranch,
            selectedEmployeeCode,
            startDateInput.value ? new Date(startDateInput.value) : null,
            endDateInput.value ? new Date(endDateInput.value) : null
        );

        const employeePerformance = {};

        // Get unique employees based on filtered data or all if no filters
        let relevantEmployees = new Map(); // employeeCode -> {name, designation}
        if (selectedEmployeeCode) {
            const name = employeeCodeToNameMap.get(selectedEmployeeCode);
            const designation = employeeCodeToDesignationMap.get(selectedEmployeeCode);
            if (name && designation) {
                relevantEmployees.set(selectedEmployeeCode, {name, designation});
            }
        } else {
            // If no specific employee, consider all employees in the selected branch or all employees
            const employeesInScope = selectedBranch ? 
                allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch) :
                allCanvassingData;
            
            employeesInScope.forEach(entry => {
                const code = entry[HEADER_EMPLOYEE_CODE];
                const name = entry[HEADER_EMPLOYEE_NAME];
                const designation = entry[HEADER_DESIGNATION];
                if (code && name && designation && !relevantEmployees.has(code)) {
                    relevantEmployees.set(code, {name, designation});
                }
            });
        }


        relevantEmployees.forEach((empInfo, code) => {
            employeePerformance[code] = {
                name: empInfo.name,
                designation: empInfo.designation,
                activity: calculateTotalActivity([])
            };
        });

        filteredData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (employeePerformance.hasOwnProperty(employeeCode)) {
                employeePerformance[employeeCode].activity[entry[HEADER_ACTIVITY_TYPE]]++;
            }
        });

        let tableHtml = `
            <table class="performance-table">
                <thead>
                    <tr>
                        <th>Employee Name</th>
                        <th>Employee Code</th>
                        <th>Branch</th>
                        <th>Designation</th>
                        <th>Visits</th>
                        <th>Calls</th>
                        <th>References</th>
                        <th>New Customer Leads</th>
                    </tr>
                </thead>
                <tbody>
        `;

        let hasData = false;
        const sortedEmployeeCodes = Array.from(Object.keys(employeePerformance)).sort((a,b) => {
            return employeePerformance[a].name.localeCompare(employeePerformance[b].name);
        });

        sortedEmployeeCodes.forEach(code => {
            const emp = employeePerformance[code];
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === code)?.[HEADER_BRANCH_NAME] || 'N/A'; // Find branch for employee
            if (Object.values(emp.activity).some(count => count > 0)) {
                hasData = true;
            }
            tableHtml += `
                <tr>
                    <td data-label="Employee Name">${emp.name}</td>
                    <td data-label="Employee Code">${code}</td>
                    <td data-label="Branch">${branchName}</td>
                    <td data-label="Designation">${emp.designation}</td>
                    <td data-label="Visits">${emp.activity['Visit']}</td>
                    <td data-label="Calls">${emp.activity['Call']}</td>
                    <td data-label="References">${emp.activity['Reference']}</td>
                    <td data-label="New Customer Leads">${emp.activity['New Customer Leads']}</td>
                </tr>
            `;
        });

        tableHtml += `
                </tbody>
            </table>
        `;

        if (hasData) {
            reportDisplay.innerHTML = tableHtml;
        } else {
            noParticipationMessage.style.display = 'block';
        }
    }


    function renderDetailedCustomerView() {
        detailedCustomerViewDisplay.innerHTML = '';
        noDetailedCustomerDataMessage.style.display = 'none';

        const selectedBranch = detailedCustomerBranchSelect.value;
        const selectedEmployeeCode = detailedCustomerEmployeeSelect.value;

        const filteredData = filterCanvassingData(
            allCanvassingData,
            selectedBranch,
            selectedEmployeeCode,
            detailedCustomerStartDateInput.value ? new Date(detailedCustomerStartDateInput.value) : null,
            detailedCustomerEndDateInput.value ? new Date(detailedCustomerEndDateInput.value) : null
        );

        if (filteredData.length === 0) {
            noDetailedCustomerDataMessage.style.display = 'block';
            return;
        }

        let html = '<div class="customer-details-list">';
        filteredData.sort((a, b) => new Date(a[HEADER_DATE]) - new Date(b[HEADER_DATE])).forEach(entry => {
            html += `
                <div class="customer-card">
                    <h3>${entry[HEADER_CUSTOMER_NAME] || 'N/A'}</h3>
                    <div class="detail-row"><span class="detail-label">Date:</span> <span class="detail-value">${entry[HEADER_DATE]}</span></div>
                    <div class="detail-row"><span class="detail-label">Branch:</span> <span class="detail-value">${entry[HEADER_BRANCH_NAME]}</span></div>
                    <div class="detail-row"><span class="detail-label">Employee:</span> <span class="detail-value">${entry[HEADER_EMPLOYEE_NAME]} (${entry[HEADER_EMPLOYEE_CODE]})</span></div>
                    <div class="detail-row"><span class="detail-label">Designation:</span> <span class="detail-value">${entry[HEADER_DESIGNATION]}</span></div>
                    <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${entry[HEADER_ACTIVITY_TYPE]}</span></div>
                    <div class="detail-row"><span class="detail-label">Type of Customer:</span> <span class="detail-value">${entry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</span></div>
                    <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${entry[HEADER_REMARKS] || 'No remarks'}</span></div>
                </div>
            `;
        });
        html += '</div>';
        detailedCustomerViewDisplay.innerHTML = html;
    }

    // New Report: Branch Diagnosis
    function renderBranchDiagnosisReport() {
        branchDiagnosisDisplay.innerHTML = '';
        noParticipationMessage.style.display = 'none'; // Reuse if no data
        employeeFilterPanel.style.display = 'none';

        const workingDaysPassed = getWorkingDaysPassedThisMonth();
        const idealFactor = workingDaysPassed / MONTHLY_WORKING_DAYS;

        let tableHtml = `
            <table class="branch-diagnosis-table">
                <thead>
                    <tr>
                        <th>Branch Name</th>
                        <th>Total Target (Visits)</th>
                        <th>Ideal (Visits)</th>
                        <th>Actual (Visits)</th>
                        <th>Diff (Visits)</th>
                        <th>Total Target (Calls)</th>
                        <th>Ideal (Calls)</th>
                        <th>Actual (Calls)</th>
                        <th>Diff (Calls)</th>
                        <th>Total Target (References)</th>
                        <th>Ideal (References)</th>
                        <th>Actual (References)</th>
                        <th>Diff (References)</th>
                        <th>Total Target (New Leads)</th>
                        <th>Ideal (New Leads)</th>
                        <th>Actual (New Leads)</th>
                        <th>Diff (New Leads)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        let hasData = false;
        PREDEFINED_BRANCHES.forEach(branch => {
            const employeesInBranch = Array.from(new Set(
                allCanvassingData
                    .filter(entry => entry[HEADER_BRANCH_NAME] === branch)
                    .map(entry => entry[HEADER_EMPLOYEE_CODE])
            ));

            const branchTotalTargets = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 };
            
            employeesInBranch.forEach(employeeCode => {
                const designation = employeeCodeToDesignationMap.get(employeeCode);
                if (designation) {
                    const employeeTargets = calculateTargetForEmployee(designation);
                    branchTotalTargets['Visit'] += employeeTargets['Visit'] || 0;
                    branchTotalTargets['Call'] += employeeTargets['Call'] || 0;
                    branchTotalTargets['Reference'] += employeeTargets['Reference'] || 0;
                    branchTotalTargets['New Customer Leads'] += employeeTargets['New Customer Leads'] || 0;
                }
            });

            const branchIdealTargets = {
                'Visit': Math.round(branchTotalTargets['Visit'] * idealFactor),
                'Call': Math.round(branchTotalTargets['Call'] * idealFactor),
                'Reference': Math.round(branchTotalTargets['Reference'] * idealFactor),
                'New Customer Leads': Math.round(branchTotalTargets['New Customer Leads'] * idealFactor)
            };

            const branchActualActivities = calculateTotalActivity(
                allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch)
            );

            const visitDiff = branchActualActivities['Visit'] - branchIdealTargets['Visit'];
            const callDiff = branchActualActivities['Call'] - branchIdealTargets['Call'];
            const refDiff = branchActualActivities['Reference'] - branchIdealTargets['Reference'];
            const newLeadsDiff = branchActualActivities['New Customer Leads'] - branchIdealTargets['New Customer Leads'];

            // Check if any actual activity or target exists for the branch to show it
            if (Object.values(branchActualActivities).some(count => count > 0) || Object.values(branchTotalTargets).some(count => count > 0)) {
                hasData = true;
            }

            tableHtml += `
                <tr>
                    <td data-label="Branch Name">${branch}</td>
                    <td data-label="Total Target (Visits)">${branchTotalTargets['Visit']}</td>
                    <td data-label="Ideal (Visits)">${branchIdealTargets['Visit']}</td>
                    <td data-label="Actual (Visits)">${branchActualActivities['Visit']}</td>
                    <td data-label="Diff (Visits)" class="${visitDiff >= 0 ? 'positive-diagnosis' : 'negative-diagnosis'}">${visitDiff}</td>
                    <td data-label="Total Target (Calls)">${branchTotalTargets['Call']}</td>
                    <td data-label="Ideal (Calls)">${branchIdealTargets['Call']}</td>
                    <td data-label="Actual (Calls)">${branchActualActivities['Call']}</td>
                    <td data-label="Diff (Calls)" class="${callDiff >= 0 ? 'positive-diagnosis' : 'negative-diagnosis'}">${callDiff}</td>
                    <td data-label="Total Target (References)">${branchTotalTargets['Reference']}</td>
                    <td data-label="Ideal (References)">${branchIdealTargets['Reference']}</td>
                    <td data-label="Actual (References)">${branchActualActivities['Reference']}</td>
                    <td data-label="Diff (References)" class="${refDiff >= 0 ? 'positive-diagnosis' : 'negative-diagnosis'}">${refDiff}</td>
                    <td data-label="Total Target (New Leads)">${branchTotalTargets['New Customer Leads']}</td>
                    <td data-label="Ideal (New Leads)">${branchIdealTargets['New Customer Leads']}</td>
                    <td data-label="Actual (New Leads)">${branchActualActivities['New Customer Leads']}</td>
                    <td data-label="Diff (New Leads)" class="${newLeadsDiff >= 0 ? 'positive-diagnosis' : 'negative-diagnosis'}">${newLeadsDiff}</td>
                </tr>
            `;
        });

        tableHtml += `
                </tbody>
            </table>
        `;

        if (hasData) {
            branchDiagnosisDisplay.innerHTML = tableHtml;
        } else {
            branchDiagnosisDisplay.innerHTML = '<p>No branch diagnosis data available.</p>'; // Specific message for this report
        }
    }


    // *** Tab Switching Logic ***
    function showTab(tabId) {
        // Hide all sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        branchDiagnosisSection.style.display = 'none'; // Hide new section

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the selected section and activate its button
        let currentSection;
        let currentButton;

        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                currentSection = reportsSection;
                currentButton = allBranchSnapshotTabBtn;
                renderAllBranchSnapshot();
                break;
            case 'allStaffOverallPerformanceTabBtn':
                currentSection = reportsSection;
                currentButton = allStaffOverallPerformanceTabBtn;
                renderOverallStaffPerformanceReport();
                break;
            case 'nonParticipatingBranchesTabBtn':
                currentSection = reportsSection;
                currentButton = nonParticipatingBranchesTabBtn;
                renderNonParticipatingBranches();
                break;
            case 'detailedCustomerViewTabBtn':
                currentSection = detailedCustomerViewSection;
                currentButton = detailedCustomerViewTabBtn;
                renderDetailedCustomerView();
                break;
            case 'employeeManagementTabBtn':
                currentSection = employeeManagementSection;
                currentButton = employeeManagementTabBtn;
                // No rendering function call needed directly for employee management, just show the section
                break;
            case 'branchDiagnosisTabBtn': // New Tab Case
                currentSection = branchDiagnosisSection;
                currentButton = branchDiagnosisTabBtn;
                renderBranchDiagnosisReport();
                break;
            default:
                // Fallback to default report if an unknown tabId is passed
                currentSection = reportsSection;
                currentButton = allBranchSnapshotTabBtn;
                renderAllBranchSnapshot();
                break;
        }

        if (currentSection) {
            currentSection.style.display = 'block';
        }
        if (currentButton) {
            currentButton.classList.add('active');
        }
    }


    // *** Event Listeners ***

    // Main Report Tabs
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    branchDiagnosisTabBtn.addEventListener('click', () => showTab('branchDiagnosisTabBtn')); // New Event Listener

    // Main Report Controls
    branchSelect.addEventListener('change', () => {
        populateEmployeeDropdowns(branchSelect.value); // Update employee dropdown based on branch
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id); // Re-render current report with new filter
        }
    });

    employeeSelect.addEventListener('change', () => {
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton && activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
            renderOverallStaffPerformanceReport();
        }
    });

    applyDateFilterBtn.addEventListener('click', () => {
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id); // Re-render current report with new date filter
        }
    });

    clearDateFilterBtn.addEventListener('click', () => {
        startDateInput.value = '';
        endDateInput.value = '';
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id); // Re-render current report clearing date filter
        }
    });

    // Detailed Customer View Controls
    detailedCustomerBranchSelect.addEventListener('change', () => {
        populateEmployeeDropdowns(detailedCustomerBranchSelect.value); // Update employee dropdown based on branch
        renderDetailedCustomerView();
    });

    detailedCustomerEmployeeSelect.addEventListener('change', renderDetailedCustomerView);
    applyDetailedCustomerDateFilterBtn.addEventListener('click', renderDetailedCustomerView);
    clearDetailedCustomerDateFilterBtn.addEventListener('click', () => {
        detailedCustomerStartDateInput.value = '';
        detailedCustomerEndDateInput.value = '';
        renderDetailedCustomerView();
    });

    // Employee Management Forms (Simplified for front-end, assumes Apps Script backend)

    async function sendDataToGoogleAppsScript(action, data) {
        showStatusMessage(`Sending ${action} request...`);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Required for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action: action, data: data }),
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                showStatusMessage(`${action} successful: ${result.message}`, false);
                // Re-fetch data if successful to update reports
                await processData(); 
                // Re-render the current active report after data update
                const activeTabButton = document.querySelector('.tab-button.active');
                if (activeTabButton) {
                    showTab(activeTabButton.id);
                }
                return true;
            } else {
                showStatusMessage(`${action} failed: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error(`Error during ${action}:`, error);
            showStatusMessage(`Error during ${action}: ${error.message}`, true);
            return false;
        }
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        const msgDiv = document.getElementById('employeeManagementMessage') || document.createElement('div');
        msgDiv.id = 'employeeManagementMessage';
        msgDiv.textContent = message;
        msgDiv.className = isError ? 'message-container error' : 'message-container success';
        msgDiv.style.display = 'block';
        employeeManagementSection.prepend(msgDiv); // Add to the top of the section
        setTimeout(() => {
            msgDiv.style.display = 'none';
            msgDiv.remove(); // Clean up
        }, 5000);
    }


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

            if (Object.values(employeeData).some(val => !val)) {
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

            const lines = employeeDetails.split('\n');
            const employeesToAdd = [];
            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 3 && parts[0] && parts[1] && parts[2]) {
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
