document.addEventListener('DOMContentLoaded', () => {

    // --- START: TWO-TIERED FRONT-END PASSWORD PROTECTION ---
    const ACCESS_PASSWORD_FULL = "sml4576"; // Full access password
    const ACCESS_PASSWORD_LIMITED = "123";  // Limited access password
    let currentAccessLevel = null; // To store 'full' or 'limited'
    const accessDeniedOverlay = document.getElementById('accessDeniedOverlay');
    const dashboardContent = document.getElementById('dashboardContent');
    const secretPasswordInputContainer = document.getElementById('secretPasswordInputContainer');
    const secretPasswordInput = document.getElementById('secretPasswordInput');
    const submitSecretPasswordBtn = document.getElementById('submitSecretPassword');
    const passwordErrorMessage = document.getElementById('passwordErrorMessage');
   // Get references to buttons/tabs that need conditional access
    const downloadOverallStaffPerformanceReportBtn = document.getElementById('downloadOverallStaffPerformanceReportBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const viewAllEntriesButton = document.getElementById('viewAllEntriesBtn'); // <--- ADD THIS LINE (OR UPDATE THE EXISTING PLACEHOLDER)
       if (secretPasswordInput) {
        secretPasswordInput.focus();
    }
    submitSecretPasswordBtn.addEventListener('click', () => {
        checkAndSetAccess();
    });
    secretPasswordInput.addEventListener('keypress', (event) => {
        if (event.key === 'Enter') {
            checkAndSetAccess();
        }
    });

    function checkAndSetAccess() {
        const enteredPassword = secretPasswordInput.value;
        passwordErrorMessage.textContent = ''; // Clear previous messages

        if (enteredPassword === ACCESS_PASSWORD_FULL) {
            currentAccessLevel = 'full';
            hideAccessDeniedOverlay();
            initializeDashboardUI();
        } else if (enteredPassword === ACCESS_PASSWORD_LIMITED) {
            currentAccessLevel = 'limited';
            hideAccessDeniedOverlay();
            initializeDashboardUI();
        } else {
            passwordErrorMessage.textContent = 'Incorrect password. Please try again.';
            secretPasswordInput.value = ''; // Clear input
            secretPasswordInput.focus();
        }
    }

    function hideAccessDeniedOverlay() {
        accessDeniedOverlay.style.display = 'none';
        dashboardContent.style.display = 'block'; // Show dashboard
    }

    function initializeDashboardUI() {
        // Show/hide elements based on access level
        if (downloadOverallStaffPerformanceReportBtn) {
            downloadOverallStaffPerformanceReportBtn.style.display = (currentAccessLevel === 'full') ? 'block' : 'none';
        }
        if (detailedCustomerViewTabBtn) {
            detailedCustomerViewTabBtn.style.display = (currentAccessLevel === 'full') ? 'block' : 'none';
        }
         // Show/hide view all entries button
        if (viewAllEntriesButton) {
            viewAllEntriesButton.style.display = (currentAccessLevel === 'full') ? 'inline-block' : 'none';
        }
        // Initially show the first tab and populate data
        processData(); // Initial data fetch
        showTab('allBranchSnapshotTabBtn'); // Show the default tab
    }
    // --- END: TWO-TIERED FRONT-END PASSWORD PROTECTION ---

    // Constants for Google Apps Script Web App URL and Sheet Headers
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzzbK48R7f3P7w9l3f2rJc5t9v-02pP1o2cM7s9b8n7e6x/exec'; // Replace with your Web App URL
    const DATA_URL = 'https://script.google.com/macros/s/AKfycbzzbK48R7f3P7w9l3f2rJc5t9v-02pP1o2cM7s9b8n7e6x/exec?action=getData'; // URL to fetch data

    // Sheet Headers (ensure these match your Google Sheet header names exactly)
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer';
    const HEADER_STATUS = 'Status';
    const HEADER_REMARKS = 'Remarks';
    const HEADER_TARGET_VISIT = 'Target Visit';
    const HEADER_TARGET_CALL = 'Target Call';
    const HEADER_TARGET_REFERENCE = 'Target Reference';
    const HEADER_TARGET_NEW_LEAD = 'Target New Lead';

    // Global variables to store data
    let canvassingData = [];
    let employeeCodeToNameMap = {};
    let employeeNameToCodeMap = {};
    let employeeBranchesMap = {}; // New map to store branch for each employee code
    let employeeDesignationsMap = {}; // New map to store designation for each employee code

    // Global filter variables
    let selectedMonth;
    let selectedYear;

    // Get DOM elements
    const messageContainer = document.getElementById('messageContainer');
    const employeeManagementMessageContainer = document.getElementById('employeeManagementMessageContainer');

    const monthSelect = document.getElementById('monthSelect');
    const yearSelect = document.getElementById('yearSelect');

    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');

    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const employeePerformanceTabBtn = document.getElementById('employeePerformanceTabBtn');
    const staffParticipationTabBtn = document.getElementById('staffParticipationTabBtn');
    // const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn'); // Already declared at the top

    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeBranchInput = document.getElementById('employeeBranch');
    const employeeDesignationInput = document.getElementById('employeeDesignation');

    const editEmployeeForm = document.getElementById('editEmployeeForm');
    const editEmployeeCodeSelect = document.getElementById('editEmployeeCode');
    const editEmployeeNameInput = document.getElementById('editEmployeeName');
    const editEmployeeBranchInput = document.getElementById('editEmployeeBranch');
    const editEmployeeDesignationInput = document.getElementById('editEmployeeDesignation');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    const dashboardSections = document.querySelectorAll('.dashboard-section');

    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const employeePerformanceTableBody = document.getElementById('employeePerformanceTableBody');

    const employeeSummaryName = document.getElementById('employeeSummaryName');
    const employeeSummaryCode = document.getElementById('employeeSummaryCode');
    const employeeSummaryBranch = document.getElementById('employeeSummaryBranch');
    const employeeSummaryDesignation = document.getElementById('employeeSummaryDesignation');

    const employeeSummaryVisits = document.getElementById('employeeSummaryVisits');
    const employeeSummaryCalls = document.getElementById('employeeSummaryCalls');
    const employeeSummaryReferences = document.getElementById('employeeSummaryReferences');
    const employeeSummaryNewLeads = document.getElementById('employeeSummaryNewLeads');

    const employeeSummaryVisitsTarget = document.getElementById('employeeSummaryVisitsTarget');
    const employeeSummaryCallsTarget = document.getElementById('employeeSummaryCallsTarget');
    const employeeSummaryReferencesTarget = document.getElementById('employeeSummaryReferencesTarget');
    const employeeSummaryNewLeadsTarget = document.getElementById('employeeSummaryNewLeadsTarget');

    const employeeSummaryVisitPercentage = document.getElementById('employeeSummaryVisitPercentage');
    const employeeSummaryCallPercentage = document.getElementById('employeeSummaryCallPercentage');
    const employeeSummaryReferencePercentage = document.getElementById('employeeSummaryReferencePercentage');
    const employeeSummaryNewLeadPercentage = document.getElementById('employeeSummaryNewLeadPercentage');

    const employeeSummaryOverallCompletion = document.getElementById('employeeSummaryOverallCompletion');

    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');

    const customerCard1 = document.getElementById('customerCard1');
    const customerCard2 = document.getElementById('customerCard2');
    const customerCard3 = document.getElementById('customerCard3');

    // Function to display messages to the user
    function displayMessage(message, isError = false) {
        if (messageContainer) {
            messageContainer.textContent = message;
            messageContainer.className = isError ? 'message error' : 'message success';
            setTimeout(() => {
                messageContainer.textContent = '';
                messageContainer.className = 'message';
            }, 5000); // Message disappears after 5 seconds
        }
    }

    // Function to display messages for employee management
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessageContainer) {
            employeeManagementMessageContainer.textContent = message;
            employeeManagementMessageContainer.className = isError ? 'message error' : 'message success';
            setTimeout(() => {
                employeeManagementMessageContainer.textContent = '';
                employeeManagementMessageContainer.className = 'message';
            }, 5000); // Message disappears after 5 seconds
        }
    }

    // Function to fetch data from Google Apps Script
    async function fetchCanvassingData() {
        displayMessage('Fetching data...', 'info');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                const errorText = await response.text(); // Get more details if response is not ok
                console.error(`HTTP error! status: ${response.status}. Details: ${errorText}`);
                displayMessage(`Error fetching data: ${response.status}. Please check WEB_APP_URL and Apps Script deployment.`, true);
                return null;
            }
            const data = await response.json();
            displayMessage('Data fetched successfully!', false);
            return data;
        } catch (error) {
            console.error('Error fetching data:', error);
            displayMessage(`Error fetching data: ${error.message || 'Network error'}.`, true);
            return null;
        }
    }

    // Function to parse the raw data and populate global variables
    function parseCanvassingData(data) {
        canvassingData = data;
        employeeCodeToNameMap = {};
        employeeNameToCodeMap = {};
        employeeBranchesMap = {};
        employeeDesignationsMap = {};

        const uniqueEmployees = new Set(); // Use a set to track unique employee codes for dropdowns
        const uniqueBranches = new Set(); // Use a set to track unique branches for dropdowns

        canvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const branchName = entry[HEADER_BRANCH_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (employeeCode && employeeName) {
                employeeCodeToNameMap[employeeCode] = employeeName;
                employeeNameToCodeMap[employeeName] = employeeCode;
                uniqueEmployees.add(employeeCode);
            }
            if (branchName) {
                uniqueBranches.add(branchName);
            }
            if (employeeCode && branchName) {
                employeeBranchesMap[employeeCode] = branchName; // Store branch for employee
            }
             if (employeeCode && designation) {
                employeeDesignationsMap[employeeCode] = designation; // Store designation for employee
            }
        });

        // Populate dropdowns after data is parsed
        populateBranchDropdown(Array.from(uniqueBranches).sort());
        populateEmployeeDropdown(Array.from(uniqueEmployees).sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        }));
         // Populate edit employee dropdowns
        populateEditEmployeeDropdown(Array.from(uniqueEmployees).sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        }));
    }

    // Populate dropdowns
    function populateMonthDropdown() {
        const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        monthSelect.innerHTML = months.map((month, index) => `<option value="${index}">${month}</option>`).join('');
    }

    function populateYearDropdown() {
        const currentYear = new Date().getFullYear();
        for (let i = currentYear - 5; i <= currentYear + 1; i++) {
            const option = document.createElement('option');
            option.value = i;
            option.textContent = i;
            yearSelect.appendChild(option);
        }
    }

    function populateBranchDropdown(branches) {
        branchSelect.innerHTML = '<option value="">All Branches</option>' +
            branches.map(branch => `<option value="${branch}">${branch}</option>`).join('');

        if (customerViewBranchSelect) {
            customerViewBranchSelect.innerHTML = '<option value="">Select Branch</option>' +
                branches.map(branch => `<option value="${branch}">${branch}</option>`).join('');
        }
    }

    function populateEmployeeDropdown(employeeCodes) {
        employeeSelect.innerHTML = '<option value="">All Employees</option>' +
            employeeCodes.map(code => `<option value="${code}">${employeeCodeToNameMap[code] || code}</option>`).join('');
    }

    function populateEditEmployeeDropdown(employeeCodes) {
        editEmployeeCodeSelect.innerHTML = '<option value="">Select Employee Code</option>' +
            employeeCodes.map(code => `<option value="${code}">${code} - ${employeeCodeToNameMap[code] || code}</option>`).join('');
    }

    // Get filtered data based on month and year
    function getFilteredCanvassingData() {
        return canvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === selectedMonth && entryDate.getFullYear() === selectedYear;
        });
    }

    // Process data (fetch and parse)
    async function processData() {
        const data = await fetchCanvassingData();
        if (data) {
            parseCanvassingData(data);
        }
    }

    // --- Tab Management ---
    function showTab(activeTabBtnId) {
        // Remove 'active' class from all tab buttons
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        // Add 'active' class to the clicked tab button
        const activeTabBtn = document.getElementById(activeTabBtnId);
        if (activeTabBtn) {
            activeTabBtn.classList.add('active');
        }

        // Hide all dashboard sections
        dashboardSections.forEach(section => section.style.display = 'none');

        // Show the relevant section based on the active tab
        if (activeTabBtnId === 'allBranchSnapshotTabBtn') {
            document.getElementById('allBranchSnapshotSection').style.display = 'block';
        } else if (activeTabBtnId === 'employeePerformanceTabBtn') {
            document.getElementById('employeePerformanceSection').style.display = 'block';
        } else if (activeTabBtnId === 'staffParticipationTabBtn') {
            document.getElementById('staffParticipationSection').style.display = 'block';
        } else if (activeTabBtnId === 'detailedCustomerViewTabBtn') {
            document.getElementById('detailedCustomerViewSection').style.display = 'block';
            // Trigger rendering of customer view specific dropdowns and list
            const currentSelectedBranch = customerViewBranchSelect.value;
            const currentSelectedEmployee = customerViewEmployeeSelect.value;

            // Re-populate employee dropdown for customer view based on current branch
            if (currentSelectedBranch) {
                 const employeeSelectElement = document.getElementById('customerViewEmployeeSelect');
                employeeSelectElement.innerHTML = '<option value="">Select Employee</option>';
                 const employeesInBranch = [...new Set(getFilteredCanvassingData()
                    .filter(entry => entry[HEADER_BRANCH_NAME] === currentSelectedBranch)
                    .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                        const nameA = employeeCodeToNameMap[codeA] || codeA;
                        const nameB = employeeCodeToNameMap[codeB] || codeB;
                        return nameA.localeCompare(nameB);
                    });

                employeesInBranch.forEach(employeeCode => {
                    const option = document.createElement('option');
                    option.value = employeeCode;
                    option.textContent = employeeCodeToNameMap[employeeCode] || employeeCode;
                    employeeSelectElement.appendChild(option);
                });
            }

            // If a branch and employee are already selected, render the customer list
            if (currentSelectedBranch && currentSelectedEmployee) {
                renderCustomersCanvassedForEmployee(currentSelectedBranch, currentSelectedEmployee);
            } else {
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
            }
             renderCustomerDetails(null); // Clear customer detail cards when tab changes
        } else if (activeTabBtnId === 'employeeManagementTabBtn') {
            document.getElementById('employeeManagementSection').style.display = 'block';
        }
    }

    // --- Rendering Functions ---
    function renderReportsBasedOnFilters() {
        const allBranchSnapshotSection = document.getElementById('allBranchSnapshotSection');
        const employeePerformanceSection = document.getElementById('employeePerformanceSection');
        const staffParticipationSection = document.getElementById('staffParticipationSection');
        const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
        const employeeManagementSection = document.getElementById('employeeManagementSection'); // Reference to employee management section

        if (allBranchSnapshotSection.style.display === 'block') {
            renderOverallStaffPerformanceReport();
        } else if (employeePerformanceSection.style.display === 'block') {
            const selectedEmployeeCode = employeeSelect.value;
            if (selectedEmployeeCode) {
                const currentBranch = branchSelect.value;
                const selectedEmployeeCodeEntries = getFilteredCanvassingData().filter(entry =>
                    entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                    entry[HEADER_BRANCH_NAME] === currentBranch
                );
                renderEmployeeSummary(selectedEmployeeCode, selectedEmployeeCodeEntries);
            } else {
                // Clear employee summary if no employee selected
                employeeSummaryName.textContent = 'N/A';
                employeeSummaryCode.textContent = 'N/A';
                employeeSummaryBranch.textContent = 'N/A';
                employeeSummaryDesignation.textContent = 'N/A';

                employeeSummaryVisits.textContent = '0';
                employeeSummaryCalls.textContent = '0';
                employeeSummaryReferences.textContent = '0';
                employeeSummaryNewLeads.textContent = '0';

                employeeSummaryVisitsTarget.textContent = '0';
                employeeSummaryCallsTarget.textContent = '0';
                employeeSummaryReferencesTarget.textContent = '0';
                employeeSummaryNewLeadsTarget.textContent = '0';

                employeeSummaryVisitPercentage.textContent = '0%';
                employeeSummaryCallPercentage.textContent = '0%';
                employeeSummaryReferencePercentage.textContent = '0%';
                employeeSummaryNewLeadPercentage.textContent = '0%';
                employeeSummaryOverallCompletion.textContent = '0%';
            }
        } else if (staffParticipationSection.style.display === 'block') {
            renderStaffParticipation();
        } else if (detailedCustomerViewSection.style.display === 'block') {
            // Re-render customer view if active
            const selectedBranch = customerViewBranchSelect.value;
            const selectedEmployeeCode = customerViewEmployeeSelect.value;
            customerCanvassedList.innerHTML = '';
            if (selectedBranch && selectedEmployeeCode) {
                renderCustomersCanvassedForEmployee(selectedBranch, selectedEmployeeCode);
            } else {
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
            }
            // Clear the content of the cards
            if (customerCard1) customerCard1.innerHTML = '<h3>Canvassing Activity</h3><p>Select a customer from the list to view their details.</p>';
            if (customerCard2) customerCard2.innerHTML = '<h3>Customer Overview</h3><p>Select a customer from the list to view their details.</p>';
            if (customerCard3) customerCard3.innerHTML = '<h3>More Details</h3><p>Select a customer from the list to view their details.</p>';
        }
    }

// Function to render customers canvassed by a specific employee in a branch
const renderCustomersCanvassedForEmployee = (branchName, employeeCode) => {
    // Clear previous list
    customerCanvassedList.innerHTML = '';

    // Ensure elements exist
    if (!customerCanvassedList) {
        console.error('Customer canvassed list element not found.');
        return;
    }

    if (employeeCode && branchName) {
        // Filter unique customers canvassed by this employee in this branch for the selected month/year
        const filteredData = getFilteredCanvassingData(); // Use filtered data here
        const customersCanvassed = filteredData.filter(entry =>
            entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_PROSPECT_NAME] // Ensure Prospect Name exists
        ).map(entry => {
            // Return the whole entry for detailed view later, but ensure uniqueness by name
            // To handle multiple entries for the same customer, we'll pick the most recent one or the first one.
            return {
                name: entry[HEADER_PROSPECT_NAME],
                entry: entry // Store the full entry for later display
            };
        }).reduce((acc, current) => {
            // Ensure unique customers by name, keep the last entry if duplicates exist
            if (!acc.find(item => item.name === current.name)) {
                acc.push(current);
            }
            return acc;
        }, []);


        if (customersCanvassed.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'customer-list';
            customersCanvassed.sort((a, b) => a.name.localeCompare(b.name)).forEach(customer => {
                const li = document.createElement('li');
                li.textContent = customer.name;
                li.dataset.prospectName = customer.name; // Store prospect name for lookup
                li.classList.add('customer-list-item');
                ul.appendChild(li);
            });
            customerCanvassedList.appendChild(ul);
        } else {
            customerCanvassedList.innerHTML = '<p>No customers found for this employee in the selected branch.</p>';
        }
    } else {
        customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
    }
}; // --- ADDED MISSING CLOSING BRACE FOR renderCustomersCanvassedForEmployee ---

    function renderOverallStaffPerformanceReport() {
        const performanceSummaryTableBody = document.getElementById('performanceSummaryTableBody');
        performanceSummaryTableBody.innerHTML = ''; // Clear previous data

        const filteredData = getFilteredCanvassingData();

        const employeePerformanceSummary = {};

        // Initialize summary for all known employees
        Object.keys(employeeCodeToNameMap).forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode];
            const branch = employeeBranchesMap[employeeCode] || 'N/A'; // Get branch from map
            const designation = employeeDesignationsMap[employeeCode] || 'N/A'; // Get designation from map

            employeePerformanceSummary[employeeCode] = {
                name: employeeName,
                branch: branch,
                designation: designation,
                visits: 0,
                calls: 0,
                references: 0,
                newLeads: 0,
                targetVisit: 0,
                targetCall: 0,
                targetReference: 0,
                targetNewLead: 0
            };
        });

        // Aggregate activities and targets for the selected month
        filteredData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (employeePerformanceSummary[employeeCode]) {
                const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
                const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER] ? entry[HEADER_TYPE_OF_CUSTOMER].trim().toLowerCase() : '';

                if (activityType === 'visit') {
                    employeePerformanceSummary[employeeCode].visits++;
                } else if (activityType === 'calls') {
                    employeePerformanceSummary[employeeCode].calls++;
                } else if (activityType === 'referance') { // Typo: Should be 'reference'
                    employeePerformanceSummary[employeeCode].references++;
                }
                if (typeOfCustomer === 'new') {
                    employeePerformanceSummary[employeeCode].newLeads++;
                }

                // Aggregate targets (assuming targets are per employee per month, can be simplified for now)
                // If multiple entries for targets exist for the same employee in the same month,
                // this will take the target from the last processed entry. Adjust if target logic is more complex.
                employeePerformanceSummary[employeeCode].targetVisit = parseInt(entry[HEADER_TARGET_VISIT] || 0);
                employeePerformanceSummary[employeeCode].targetCall = parseInt(entry[HEADER_TARGET_CALL] || 0);
                employeePerformanceSummary[employeeCode].targetReference = parseInt(entry[HEADER_TARGET_REFERENCE] || 0);
                employeePerformanceSummary[employeeCode].targetNewLead = parseInt(entry[HEADER_TARGET_NEW_LEAD] || 0);
            }
        });


        // Render summary rows
        Object.values(employeePerformanceSummary).forEach(summary => {
            const row = performanceSummaryTableBody.insertRow();
            row.insertCell().textContent = summary.name;
            row.insertCell().textContent = summary.branch;
            row.insertCell().textContent = summary.designation;

            const visitCompletion = summary.targetVisit > 0 ? ((summary.visits / summary.targetVisit) * 100).toFixed(0) : (summary.visits > 0 ? 100 : 0);
            const callCompletion = summary.targetCall > 0 ? ((summary.calls / summary.targetCall) * 100).toFixed(0) : (summary.calls > 0 ? 100 : 0);
            const referenceCompletion = summary.targetReference > 0 ? ((summary.references / summary.targetReference) * 100).toFixed(0) : (summary.references > 0 ? 100 : 0);
            const newLeadCompletion = summary.targetNewLead > 0 ? ((summary.newLeads / summary.targetNewLead) * 100).toFixed(0) : (summary.newLeads > 0 ? 100 : 0);

            const calculateOverallCompletion = () => {
                let totalActual = summary.visits + summary.calls + summary.references + summary.newLeads;
                let totalTarget = summary.targetVisit + summary.targetCall + summary.targetReference + summary.targetNewLead;
                return totalTarget > 0 ? ((totalActual / totalTarget) * 100).toFixed(0) : (totalActual > 0 ? 100 : 0);
            };

            row.insertCell().textContent = `${summary.visits}/${summary.targetVisit} (${visitCompletion}%)`;
            row.insertCell().textContent = `${summary.calls}/${summary.targetCall} (${callCompletion}%)`;
            row.insertCell().textContent = `${summary.references}/${summary.targetReference} (${referenceCompletion}%)`;
            row.insertCell().textContent = `${summary.newLeads}/${summary.targetNewLead} (${newLeadCompletion}%)`;
            row.insertCell().textContent = `${calculateOverallCompletion()}%`;
        });
        if (Object.keys(employeePerformanceSummary).length === 0) {
            const row = performanceSummaryTableBody.insertRow();
            row.insertCell().textContent = 'No performance data available for selected month/year.';
            row.insertCell().colSpan = 7;
        }
    }


    function renderEmployeeSummary(employeeCode, employeeEntries) {
        // Get employee details from map
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const employeeBranch = employeeBranchesMap[employeeCode] || 'N/A';
        const employeeDesignation = employeeDesignationsMap[employeeCode] || 'N/A';

        employeeSummaryName.textContent = employeeName;
        employeeSummaryCode.textContent = employeeCode;
        employeeSummaryBranch.textContent = employeeBranch;
        employeeSummaryDesignation.textContent = employeeDesignation;

        let totalVisits = 0;
        let totalCalls = 0;
        let totalReferences = 0;
        let totalNewLeads = 0;

        let targetVisit = 0;
        let targetCall = 0;
        let targetReference = 0;
        let targetNewLead = 0;

        employeeEntries.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER] ? entry[HEADER_TYPE_OF_CUSTOMER].trim().toLowerCase() : '';

            if (activityType === 'visit') totalVisits++;
            else if (activityType === 'calls') totalCalls++;
            else if (activityType === 'referance') totalReferences++; // Typo: should be 'reference'

            if (typeOfCustomer === 'new') totalNewLeads++;

            // Assuming targets might be present in each relevant entry, or take from the last entry
            targetVisit = parseInt(entry[HEADER_TARGET_VISIT] || targetVisit);
            targetCall = parseInt(entry[HEADER_TARGET_CALL] || targetCall);
            targetReference = parseInt(entry[HEADER_TARGET_REFERENCE] || targetReference);
            targetNewLead = parseInt(entry[HEADER_TARGET_NEW_LEAD] || targetNewLead);
        });

        employeeSummaryVisits.textContent = totalVisits;
        employeeSummaryCalls.textContent = totalCalls;
        employeeSummaryReferences.textContent = totalReferences;
        employeeSummaryNewLeads.textContent = totalNewLeads;

        employeeSummaryVisitsTarget.textContent = targetVisit;
        employeeSummaryCallsTarget.textContent = targetCall;
        employeeSummaryReferencesTarget.textContent = targetReference;
        employeeSummaryNewLeadsTarget.textContent = targetNewLead;

        const visitPercentage = targetVisit > 0 ? ((totalVisits / targetVisit) * 100).toFixed(0) : (totalVisits > 0 ? 100 : 0);
        const callPercentage = targetCall > 0 ? ((totalCalls / targetCall) * 100).toFixed(0) : (totalCalls > 0 ? 100 : 0);
        const referencePercentage = targetReference > 0 ? ((totalReferences / targetReference) * 100).toFixed(0) : (totalReferences > 0 ? 100 : 0);
        const newLeadPercentage = targetNewLead > 0 ? ((totalNewLeads / targetNewLead) * 100).toFixed(0) : (totalNewLeads > 0 ? 100 : 0);

        employeeSummaryVisitPercentage.textContent = `${visitPercentage}%`;
        employeeSummaryCallPercentage.textContent = `${callPercentage}%`;
        employeeSummaryReferencePercentage.textContent = `${referencePercentage}%`;
        employeeSummaryNewLeadPercentage.textContent = `${newLeadPercentage}%`;

        const overallActual = totalVisits + totalCalls + totalReferences + totalNewLeads;
        const overallTarget = targetVisit + targetCall + targetReference + targetNewLead;
        const overallCompletion = overallTarget > 0 ? ((overallActual / overallTarget) * 100).toFixed(0) : (overallActual > 0 ? 100 : 0);
        employeeSummaryOverallCompletion.textContent = `${overallCompletion}%`;
    }

    function renderCustomerDetails(customerEntry) {
        if (!customerEntry) {
            if (customerCard1) customerCard1.innerHTML = '<h3>Canvassing Activity</h3><p>Select a customer from the list to view their details.</p>';
            if (customerCard2) customerCard2.innerHTML = '<h3>Customer Overview</h3><p>Select a customer from the list to view their details.</p>';
            if (customerCard3) customerCard3.innerHTML = '<h3>More Details</h3><p>Select a customer from the list to view their details.</p>';
            return;
        }

        // Card 1: Canvassing Activity
        if (customerCard1) {
            customerCard1.innerHTML = `
                <h3>Canvassing Activity</h3>
                <p><strong>Timestamp:</strong> ${new Date(customerEntry[HEADER_TIMESTAMP]).toLocaleString()}</p>
                <p><strong>Activity Type:</strong> ${customerEntry[HEADER_ACTIVITY_TYPE] || 'N/A'}</p>
                <p><strong>Status:</strong> ${customerEntry[HEADER_STATUS] || 'N/A'}</p>
                <p><strong>Remarks:</strong> ${customerEntry[HEADER_REMARKS] || 'N/A'}</p>
            `;
        }

        // Card 2: Customer Overview
        if (customerCard2) {
            customerCard2.innerHTML = `
                <h3>Customer Overview</h3>
                <p><strong>Prospect Name:</strong> ${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</p>
                <p><strong>Type of Customer:</strong> ${customerEntry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</p>
            `;
        }

        // Card 3: More Details (Employee info for this entry)
        if (customerCard3) {
            customerCard3.innerHTML = `
                <h3>Canvassed By</h3>
                <p><strong>Employee:</strong> ${customerEntry[HEADER_EMPLOYEE_NAME] || 'N/A'} (${customerEntry[HEADER_EMPLOYEE_CODE] || 'N/A'})</p>
                <p><strong>Branch:</strong> ${customerEntry[HEADER_BRANCH_NAME] || 'N/A'}</p>
                <p><strong>Designation:</strong> ${customerEntry[HEADER_DESIGNATION] || 'N/A'}</p>
            `;
        }
    }


    function downloadOverallStaffPerformanceReportCSV() {
        const data = Object.values(employeePerformanceSummary).map(summary => {
            const visitCompletion = summary.targetVisit > 0 ? ((summary.visits / summary.targetVisit) * 100).toFixed(0) : (summary.visits > 0 ? 100 : 0);
            const callCompletion = summary.targetCall > 0 ? ((summary.calls / summary.targetCall) * 100).toFixed(0) : (summary.calls > 0 ? 100 : 0);
            const referenceCompletion = summary.targetReference > 0 ? ((summary.references / summary.targetReference) * 100).toFixed(0) : (summary.references > 0 ? 100 : 0);
            const newLeadCompletion = summary.targetNewLead > 0 ? ((summary.newLeads / summary.targetNewLead) * 100).toFixed(0) : (summary.newLeads > 0 ? 100 : 0);

            let totalActual = summary.visits + summary.calls + summary.references + summary.newLeads;
            let totalTarget = summary.targetVisit + summary.targetCall + summary.targetReference + summary.targetNewLead;
            const overallCompletion = totalTarget > 0 ? ((totalActual / totalTarget) * 100).toFixed(0) : (totalActual > 0 ? 100 : 0);

            return [
                summary.name,
                summary.branch,
                summary.designation,
                `${summary.visits}/${summary.targetVisit} (${visitCompletion}%)`,
                `${summary.calls}/${summary.targetCall} (${callCompletion}%)`,
                `${summary.references}/${summary.targetReference} (${referenceCompletion}%)`,
                `${summary.newLeads}/${summary.targetNewLead} (${newLeadCompletion}%)`,
                `${overallCompletion}%`
            ];
        });

        const headers = ['Employee Name', 'Branch', 'Designation', 'Visit (Actual/Target/%)', 'Call (Actual/Target/%)', 'Reference (Actual/Target/%)', 'New Lead (Actual/Target/%)', 'Overall Completion'];
        const csvContent = "data:text/csv;charset=utf-8,"
            + headers.map(e => `"${e}"`).join(',') + "\n"
            + data.map(row => row.map(e => `"${e}"`).join(',')).join("\n");

        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `OverallStaffPerformance_${selectedMonth + 1}-${selectedYear}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function downloadEmployeeSummaryCSV() {
        const employeeCode = employeeSelect.value;
        if (!employeeCode) {
            displayMessage('Please select an employee to download their summary.', true);
            return;
        }

        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const employeeBranch = employeeBranchesMap[employeeCode] || 'N/A';
        const employeeDesignation = employeeDesignationsMap[employeeCode] || 'N/A';

        const filteredData = getFilteredCanvassingData().filter(entry =>
            entry[HEADER_EMPLOYEE_CODE] === employeeCode
        );

        let totalVisits = 0;
        let totalCalls = 0;
        let totalReferences = 0;
        let totalNewLeads = 0;
        let targetVisit = 0;
        let targetCall = 0;
        let targetReference = 0;
        let targetNewLead = 0;

        filteredData.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER] ? entry[HEADER_TYPE_OF_CUSTOMER].trim().toLowerCase() : '';

            if (activityType === 'visit') totalVisits++;
            else if (activityType === 'calls') totalCalls++;
            else if (activityType === 'referance') totalReferences++;
            if (typeOfCustomer === 'new') totalNewLeads++;

            targetVisit = parseInt(entry[HEADER_TARGET_VISIT] || targetVisit);
            targetCall = parseInt(entry[HEADER_TARGET_CALL] || targetCall);
            targetReference = parseInt(entry[HEADER_TARGET_REFERENCE] || targetReference);
            targetNewLead = parseInt(entry[HEADER_TARGET_NEW_LEAD] || targetNewLead);
        });

        const visitCompletion = targetVisit > 0 ? ((totalVisits / targetVisit) * 100).toFixed(0) : (totalVisits > 0 ? 100 : 0);
        const callCompletion = targetCall > 0 ? ((totalCalls / targetCall) * 100).toFixed(0) : (totalCalls > 0 ? 100 : 0);
        const referenceCompletion = targetReference > 0 ? ((totalReferences / targetReference) * 100).toFixed(0) : (targetReference > 0 ? 100 : 0);
        const newLeadCompletion = targetNewLead > 0 ? ((totalNewLeads / targetNewLead) * 100).toFixed(0) : (totalNewLeads > 0 ? 100 : 0);

        const overallActual = totalVisits + totalCalls + totalReferences + totalNewLeads;
        const overallTarget = targetVisit + targetCall + targetReference + targetNewLead;
        const overallCompletion = overallTarget > 0 ? ((overallActual / overallTarget) * 100).toFixed(0) : (overallActual > 0 ? 100 : 0);

        const data = [
            ['Employee Name', employeeName],
            ['Employee Code', employeeCode],
            ['Branch', employeeBranch],
            ['Designation', employeeDesignation],
            ['Total Visits (Actual/Target/%)', `${totalVisits}/${targetVisit} (${visitCompletion}%)`],
            ['Total Calls (Actual/Target/%)', `${totalCalls}/${targetCall} (${callCompletion}%)`],
            ['Total References (Actual/Target/%)', `${totalReferences}/${targetReference} (${referenceCompletion}%)`],
            ['Total New Leads (Actual/Target/%)', `${totalNewLeads}/${targetNewLead} (${newLeadCompletion}%)`],
            ['Overall Completion', `${overallCompletion}%`]
        ];

        const csvContent = "data:text/csv;charset=utf-8,"
            + data.map(row => row.map(e => `"${e}"`).join(',')).join("\n");

        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `EmployeeSummary_${employeeCode}_${selectedMonth + 1}-${selectedYear}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function renderStaffParticipation() {
        const staffParticipationTableContainer = document.getElementById('staffParticipationTableContainer');
        staffParticipationTableContainer.innerHTML = ''; // Clear previous table

        const filteredData = getFilteredCanvassingData();

        const employeeActivitySummary = {};

        // Initialize summary for all known employees
        Object.keys(employeeCodeToNameMap).forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode];
            const branch = employeeBranchesMap[employeeCode] || 'N/A';
            const designation = employeeDesignationsMap[employeeCode] || 'N/A';
            employeeActivitySummary[employeeCode] = {
                name: employeeName,
                branch: branch,
                designation: designation,
                totalActivities: {
                    'Visit': 0,
                    'Call': 0,
                    'Reference': 0,
                    'New Customer Leads': 0
                },
                hasActivityThisMonth: false
            };
        });

        // Aggregate activities for the selected month
        filteredData.forEach(entry => { // Use filteredData here
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (employeeActivitySummary[employeeCode]) {
                const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
                const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER] ? entry[HEADER_TYPE_OF_CUSTOMER].trim().toLowerCase() : '';

                if (activityType === 'visit') {
                    employeeActivitySummary[employeeCode].totalActivities['Visit']++;
                } else if (activityType === 'calls') {
                    employeeActivitySummary[employeeCode].totalActivities['Call']++;
                } else if (activityType === 'referance') { // Typo: Should be 'reference'
                    employeeActivitySummary[employeeCode].totalActivities['Reference']++;
                }
                if (typeOfCustomer === 'new') {
                    employeeActivitySummary[employeeCode].totalActivities['New Customer Leads']++;
                }
                employeeActivitySummary[employeeCode].hasActivityThisMonth = true;
            }
        }); // Corrected: Removed extra '}' here.

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling

        const table = document.createElement('table');
        table.className = 'performance-table'; // Reuse existing table style

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = ['Employee Name', 'Branch', 'Designation', 'Total Visits', 'Total Calls', 'Total References', 'Total New Customer Leads', 'Participation Status'];
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        Object.values(employeeActivitySummary).forEach(summary => {
            const row = tbody.insertRow();
            row.insertCell().textContent = summary.name;
            row.insertCell().textContent = summary.branch;
            row.insertCell().textContent = summary.designation;
            row.insertCell().textContent = summary.totalActivities['Visit'];
            row.insertCell().textContent = summary.totalActivities['Call'];
            row.insertCell().textContent = summary.totalActivities['Reference'];
            row.insertCell().textContent = summary.totalActivities['New Customer Leads'];
            row.insertCell().textContent = summary.hasActivityThisMonth ? 'Participated' : 'Not Participated';
            row.classList.add(summary.hasActivityThisMonth ? 'participated' : 'not-participated');
        });
        tableContainer.appendChild(table);
        staffParticipationTableContainer.appendChild(tableContainer);

        if (Object.keys(employeeActivitySummary).length === 0) {
            staffParticipationTableContainer.innerHTML = '<p>No staff participation data available for selected month/year.</p>';
        }
    }


    // Function to send data to Google Apps Script (Add/Edit/Delete Employee)
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage(`Sending data for ${action}...`, 'info');
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data }),
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error from Apps Script! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to send data to Apps Script. Status: ${response.status}. Please check WEB_APP_URL and Apps Script deployment.`);
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
            displayEmployeeManagementMessage(`Error sending data: ${error.message}. Please check WEB_APP_URL and Apps Script deployment.`, true);
            return false;
        } finally {
            // Re-fetch all data to ensure reports are up-to-date after any employee management action
            await processData(); // Re-fetch canvassing data and re-populate maps/dropdowns
            renderReportsBasedOnFilters(); // Re-render the current report based on selected filters
        }
    }

    // --- Event Listeners ---

    // Initial data fetch and tab display when the page loads
    // Initial data fetch and tab display when the page loads
    populateMonthDropdown();
    populateYearDropdown();

    // Set default selected month and year to current month and year
    const today = new Date();
    selectedMonth = today.getMonth();
    selectedYear = today.getFullYear();
    monthSelect.value = selectedMonth;
    yearSelect.value = selectedYear;

    // Event listeners for month and year changes
    monthSelect.addEventListener('change', (event) => {
        selectedMonth = parseInt(event.target.value);
        renderReportsBasedOnFilters(); // This new function will re-render all reports
    });

    yearSelect.addEventListener('change', (event) => {
        selectedYear = parseInt(event.target.value);
        renderReportsBasedOnFilters(); // This new function will re-render all reports
    });

    // Event listeners for branch and employee dropdowns (for performance tab)
    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            // Filter employees based on selected branch
            const employeesInBranch = selectedBranch ? [...new Set(getFilteredCanvassingData()
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                }) : Object.keys(employeeCodeToNameMap).sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                });

            employeeSelect.innerHTML = '<option value="">All Employees</option>' +
                employeesInBranch.map(code => `<option value="${code}">${employeeCodeToNameMap[code] || code}</option>`).join('');

            // If an employee was previously selected and is no longer in the filtered list, reset employeeSelect
            if (employeeSelect.value && !employeesInBranch.includes(employeeSelect.value)) {
                employeeSelect.value = '';
            }
            renderReportsBasedOnFilters(); // Re-render reports based on the new filter
        });
    }

    if (employeeSelect) {
        employeeSelect.addEventListener('change', () => {
            renderReportsBasedOnFilters(); // Re-render reports based on the new filter
        });
    }

    // --- Tab Button Event Listeners ---
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => {
            showTab('allBranchSnapshotTabBtn');
            renderReportsBasedOnFilters();
        });
    }

    if (employeePerformanceTabBtn) {
        employeePerformanceTabBtn.addEventListener('click', () => {
            showTab('employeePerformanceTabBtn');
            renderReportsBasedOnFilters();
        });
    }

    if (staffParticipationTabBtn) {
        staffParticipationTabBtn.addEventListener('click', () => {
            showTab('staffParticipationTabBtn');
            renderReportsBasedOnFilters();
        });
    }

    // --- NEW: Event Listener for Detailed Customer View Tab Button ---
    if (detailedCustomerViewTabBtn) {
        detailedCustomerViewTabBtn.addEventListener('click', () => {
            showTab('detailedCustomerViewTabBtn');
            renderReportsBasedOnFilters();
        });
    }
    // --- END NEW ---

    // --- NEW: Event listener for Customer View Branch Select ---
    if (customerViewBranchSelect) {
        customerViewBranchSelect.addEventListener('change', () => {
            const selectedBranch = customerViewBranchSelect.value;
            const employeeSelectElement = document.getElementById('customerViewEmployeeSelect');
            // Clear previous options
            employeeSelectElement.innerHTML = '<option value="">Select Employee</option>';

            if (selectedBranch) {
                // Get unique employees who made entries in this branch for the selected month/year
                const employeesInBranch = [...new Set(getFilteredCanvassingData() // Use filtered data
                    .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                    .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                        const nameA = employeeCodeToNameMap[codeA] || codeA;
                        const nameB = employeeCodeToNameMap[codeB] || codeB;
                        return nameA.localeCompare(nameB);
                    });

                employeesInBranch.forEach(employeeCode => {
                    const option = document.createElement('option');
                    option.value = employeeCode;
                    option.textContent = employeeCodeToNameMap[employeeCode] || employeeCode;
                    employeeSelectElement.appendChild(option);
                });
            }
            // Always reset customer list when branch changes
            customerCanvassedList.innerHTML = '';
            renderCustomerDetails(null); // Clear customer detail cards
        });
    }

    // --- NEW: Event listener for Customer View Employee Select ---
    if (customerViewEmployeeSelect) {
        customerViewEmployeeSelect.addEventListener('change', () => {
            const selectedBranch = customerViewBranchSelect.value;
            const selectedEmployeeCode = customerViewEmployeeSelect.value;

            // Render customer list only if both branch and employee are selected
            if (selectedBranch && selectedEmployeeCode) {
                renderCustomersCanvassedForEmployee(selectedBranch, selectedEmployeeCode);
            } else {
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
            }
            renderCustomerDetails(null); // Clear customer detail cards when employee changes
        });
    }

    // --- NEW: Event Listener for Customer Canvassed List (Click to show details) ---
    if (customerCanvassedList) {
        customerCanvassedList.addEventListener('click', (event) => {
            if (event.target.classList.contains('customer-list-item')) {
                // Remove active class from previously selected item
                document.querySelectorAll('.customer-list-item').forEach(item => item.classList.remove('active'));
                // Add active class to clicked item
                event.target.classList.add('active');

                const prospectName = event.target.dataset.prospectName;
                const selectedEmployeeCode = customerViewEmployeeSelect.value;
                const selectedBranch = customerViewBranchSelect.value;

                // Find the *latest* entry for this specific prospect, employee, and branch
                // This ensures we get the most up-to-date details if a customer has multiple entries
                const filteredData = getFilteredCanvassingData();
                const customerEntry = filteredData
                    .filter(entry => entry[HEADER_PROSPECT_NAME] === prospectName && entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode && entry[HEADER_BRANCH_NAME] === selectedBranch)
                    .sort((a, b) => {
                        const dateA = new Date(a[HEADER_TIMESTAMP]);
                        const dateB = new Date(b[HEADER_TIMESTAMP]);
                        return dateB.getTime() - dateA.getTime();
                    })[0];

                if (customerEntry) {
                    renderCustomerDetails(customerEntry);
                } else {
                    // If no entry found, clear details in cards
                    if (customerCard1) customerCard1.innerHTML = '<h3>Canvassing Activity</h3><p>Details not found for this customer.</p>';
                    if (customerCard2) customerCard2.innerHTML = '<h3>Customer Overview</h3><p>Details not found for this customer.</p>';
                    if (customerCard3) customerCard3.innerHTML = '<h3>More Details</h3><p>Details not found for this customer.</p>';
                }
            }
        });
    }
    // --- END NEW ---

    // --- Employee Management Tab Event Listener ---
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => {
            showTab('employeeManagementTabBtn');
             // Clear forms and messages when tab is opened
            addEmployeeForm.reset();
            editEmployeeForm.reset();
            deleteEmployeeForm.reset();
            displayEmployeeManagementMessage(''); // Clear any previous messages

            // Re-populate edit/delete employee dropdowns in case data changed
            populateEditEmployeeDropdown(Object.keys(employeeCodeToNameMap).sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            }));
        });
    }

    // --- Employee Management Forms Submission ---

    // Add Employee
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault(); // Prevent default form submission

            const newEmployeeCode = employeeCodeInput.value.trim();
            const newEmployeeName = employeeNameInput.value.trim();
            const newEmployeeBranch = employeeBranchInput.value.trim();
            const newEmployeeDesignation = employeeDesignationInput.value.trim();

            if (!newEmployeeCode || !newEmployeeName || !newEmployeeBranch || !newEmployeeDesignation) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const addData = {
                [HEADER_EMPLOYEE_CODE]: newEmployeeCode,
                [HEADER_EMPLOYEE_NAME]: newEmployeeName,
                [HEADER_BRANCH_NAME]: newEmployeeBranch,
                [HEADER_DESIGNATION]: newEmployeeDesignation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', addData);

            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Edit Employee
    if (editEmployeeForm) {
        editEmployeeCodeSelect.addEventListener('change', () => {
            const selectedCode = editEmployeeCodeSelect.value;
            if (selectedCode) {
                editEmployeeNameInput.value = employeeCodeToNameMap[selectedCode] || '';
                editEmployeeBranchInput.value = employeeBranchesMap[selectedCode] || '';
                editEmployeeDesignationInput.value = employeeDesignationsMap[selectedCode] || '';
            } else {
                editEmployeeNameInput.value = '';
                editEmployeeBranchInput.value = '';
                editEmployeeDesignationInput.value = '';
            }
        });

        editEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();

            const employeeCodeToEdit = editEmployeeCodeSelect.value.trim();
            const updatedEmployeeName = editEmployeeNameInput.value.trim();
            const updatedEmployeeBranch = editEmployeeBranchInput.value.trim();
            const updatedEmployeeDesignation = editEmployeeDesignationInput.value.trim();

            if (!employeeCodeToEdit || !updatedEmployeeName || !updatedEmployeeBranch || !updatedEmployeeDesignation) {
                displayEmployeeManagementMessage('All fields are required for editing an employee.', true);
                return;
            }

            const editData = {
                [HEADER_EMPLOYEE_CODE]: employeeCodeToEdit,
                [HEADER_EMPLOYEE_NAME]: updatedEmployeeName,
                [HEADER_BRANCH_NAME]: updatedEmployeeBranch,
                [HEADER_DESIGNATION]: updatedEmployeeDesignation
            };

            const success = await sendDataToGoogleAppsScript('edit_employee', editData);

            if (success) {
                editEmployeeForm.reset();
            }
        });
    }

    // Delete Employee
    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();

            const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
                return;
            }

            // Optional: Add a confirmation dialog for deletion
            if (!confirm(`Are you sure you want to delete employee with code: ${employeeCodeToDelete}? This action cannot be undone.`)) {
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
            }
        });
    }

    // NEW: Event Listener for "Download Overall Staff Performance CSV" button
    if (downloadOverallStaffPerformanceReportBtn) {
        downloadOverallStaffPerformanceReportBtn.addEventListener('click', () => {
            downloadOverallStaffPerformanceReportCSV();
        });
    }
    // --- END NEW ---

    // Initial data fetch and tab display when the page loads
    populateMonthDropdown();
    populateYearDropdown();

    // Set default selected month and year to current month and year
    const today = new Date();
    selectedMonth = today.getMonth();
    selectedYear = today.getFullYear();
    monthSelect.value = selectedMonth;
    yearSelect.value = selectedYear;

    // Event listeners for month and year changes
    monthSelect.addEventListener('change', (event) => {
        selectedMonth = parseInt(event.target.value);
        renderReportsBasedOnFilters(); // This new function will re-render all reports
    });

    yearSelect.addEventListener('change', (event) => {
        selectedYear = parseInt(event.target.value);
        renderReportsBasedOnFilters(); // This new function will re-render all reports
    });

    // Call processData once to fetch the raw data, then render initial reports
    processData().then(() => {
        renderReportsBasedOnFilters(); // Render initial reports with current month/year
    });
}); // This is the closing brace for DOMContentLoaded
