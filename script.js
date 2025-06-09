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
        // Predefined list of branches for the dropdown and "no participation" check
    const PREDEFINED_BRANCHES = [
        "Angamaly", "Corporate Office", "Edappally", "Harippad", "Koduvayur", "Kuzhalmannam",
        "Mattanchery", "Mavelikara", "Nedumkandom", "Nenmara", "Paravoor", "Perumbavoor",
        "Thiruwillamala", "Thodupuzha", "Chengannur", "Alathur", "Kottayam", "Kattapana",
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "HO KKM" // Corrected "Pathanamthitta" typo if it existed previously
    ].sort();

  // --- Column Headers Mapping (IMPORTANT: These must EXACTLY match the column names in your "Form Responses 2" Google Sheet) ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer'; // !!! CORRECTED TYPO HERE !!!
    const HEADER_R_LEAD_SOURCE = 'rLead Source';      // Keeping user's provided interpretation of split header
    const HEADER_HOW_CONTACTED = 'How Contacted';
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)'; // Keeping user's provided typo
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested'; // Keeping user's provided typo
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';

    let canvassingData = []; // Stores the fetched and processed data

    // --- DOM Elements ---
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    const allBranchSnapshotTable = document.getElementById('allBranchSnapshotTable').querySelector('tbody');
    const allStaffOverallPerformanceTable = document.getElementById('allStaffOverallPerformanceTable').querySelector('tbody');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');

    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const statusMessageDiv = document.getElementById('statusMessage');

    // --- Helper Functions ---

    // Generic function to display status messages
    function displayStatusMessage(message, isError = false) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = isError ? 'message-container error' : 'message-container success';
        statusMessageDiv.style.display = 'block';
        setTimeout(() => {
            statusMessageDiv.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    // Specific message for Employee Management form
    function displayEmployeeManagementMessage(message, isError = false) {
        const msgDiv = document.getElementById('employeeManagementMessage'); // Assuming you add this div in your HTML
        if (!msgDiv) {
            console.error("Employee management message div not found!");
            return;
        }
        msgDiv.textContent = message;
        msgDiv.className = isError ? 'message-container error' : 'message-container success';
        msgDiv.style.display = 'block';
        setTimeout(() => {
            msgDiv.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }


    // Parses CSV text into an array of objects
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
                    let value = values[index];
                    // Robustly parse numeric headers
                    if ([HEADER_CALL, HEADER_VISIT, HEADER_REFERENCE].includes(header)) {
                        const parsed = parseFloat(value);
                        row[header] = isNaN(parsed) ? 0 : Math.round(parsed);
                    } else {
                        row[header] = value; // Keep as string for other headers like Activity Type, Type of Customer, Remarks
                    }
                });
                data.push(row);
            }
        }
        return data;
    }

    // New function to process parsed data and calculate derived fields
    function processCanvassingData(data) {
        return data.map(row => {
            let newCustomerLeads = 0;
            const activityType = row[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = row[HEADER_TYPE_OF_CUSTOMER];

            // Calculate New Customer Leads based on condition
            if (activityType && typeOfCustomer &&
                (activityType.toLowerCase() === 'calls' || activityType.toLowerCase() === 'visit') &&
                typeOfCustomer.toLowerCase() === 'new') {
                newCustomerLeads = 1; // Count as 1 new customer lead for this activity
            }
            return { ...row, newCustomerLeads: newCustomerLeads }; // Add newCustomerLeads to the row object
        });
    }

     // Function to fetch activity data from Google Sheet (Form Responses 2)
    async function fetchCanvassingData() {
        displayMessage("Fetching activity data...", 'info');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error fetching Canvassing Data! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to fetch canvassing data. Status: ${response.status}. Please check DATA_URL.`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            console.log('--- Fetched Canvassing Data: ---');
            console.log(allCanvassingData); // Log canvassing data for debugging
            if (allCanvassingData.length > 0) {
                console.log('Canvassing Data Headers (first entry):', Object.keys(allCanvassingData[0]));
            }
            displayMessage("Activity data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching canvassing data:', error);
            displayMessage(`Failed to load activity data: ${error.message}. Please ensure the sheet is published correctly to CSV and the URL is accurate.`, 'error');
            allCanvassingData = [];
        }
    }


    // Populates the branch dropdown filter
    function populateBranchFilter(data) {
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        const branches = new Set(data.map(row => row[HEADER_BRANCH_NAME]).filter(Boolean));
        branches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    // Populates the employee dropdown filter based on selected branch
    function populateEmployeeFilter(data) {
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        employeeFilterPanel.style.display = 'none'; // Hide by default

        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';
            const employeesInBranch = new Set();
            data.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch)
                .forEach(row => {
                    // Combine code and name for display if available, otherwise just code
                    const employeeDisplay = row[HEADER_EMPLOYEE_NAME] && row[HEADER_EMPLOYEE_CODE]
                        ? `${row[HEADER_EMPLOYEE_NAME]} (${row[HEADER_EMPLOYEE_CODE]})`
                        : row[HEADER_EMPLOYEE_CODE];
                    if (employeeDisplay) {
                        employeesInBranch.add(employeeDisplay);
                    }
                });

            employeesInBranch.forEach(employeeDisplay => {
                const option = document.createElement('option');
                option.value = employeeDisplay.match(/\(([^)]+)\)/)?.[1] || employeeDisplay; // Extract code if format is "Name (Code)"
                option.textContent = employeeDisplay;
                employeeSelect.appendChild(option);
            });
        }
    }

    // Filters data based on selected branch and employee
    function getFilteredData(data) {
        let filteredData = data;
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value;

        if (selectedBranch) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === selectedBranch);
        }
        if (selectedEmployee) {
            filteredData = filteredData.filter(row => row[HEADER_EMPLOYEE_CODE] === selectedEmployee);
        }
        return filteredData;
    }

    // Renders the All Branch Snapshot report
    function renderAllBranchSnapshot(data) {
        allBranchSnapshotTable.innerHTML = ''; // Clear previous data
        const filteredData = getFilteredData(data);

        const branchSummary = {};
        const employeeSet = new Set(); // To count unique employees overall

        filteredData.forEach(row => {
            const branchName = row[HEADER_BRANCH_NAME];
            const employeeCode = row[HEADER_EMPLOYEE_CODE];

            if (!branchSummary[branchName]) {
                branchSummary[branchName] = {
                    employees: new Set(),
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0 // Will be summed from pre-calculated field
                };
            }

            branchSummary[branchName].employees.add(employeeCode);
            branchSummary[branchName].calls += row[HEADER_CALL];
            branchSummary[branchName].visits += row[HEADER_VISIT];
            branchSummary[branchName].references += row[HEADER_REFERENCE];
            branchSummary[branchName].newCustomerLeads += row.newCustomerLeads; // Sum pre-calculated field

            employeeSet.add(employeeCode);
        });

        // Add overall summary row if no branch is selected
        if (!branchSelect.value) {
            let totalEmployees = employeeSet.size;
            let totalCalls = 0;
            let totalVisits = 0;
            let totalReferences = 0;
            let totalNewCustomerLeads = 0;

            for (const branch in branchSummary) {
                totalCalls += branchSummary[branch].calls;
                totalVisits += branchSummary[branch].visits;
                totalReferences += branchSummary[branch].references;
                totalNewCustomerLeads += branchSummary[branch].newCustomerLeads;
            }

            const overallRow = allBranchSnapshotTable.insertRow();
            overallRow.classList.add('overall-summary-row'); // Add class for styling
            overallRow.insertCell().textContent = 'Overall Total';
            overallRow.insertCell().textContent = totalEmployees;
            overallRow.insertCell().textContent = totalCalls;
            overallRow.insertCell().textContent = totalVisits;
            overallRow.insertCell().textContent = totalReferences;
            overallRow.insertCell().textContent = totalNewCustomerLeads;
        }


        // Render branch-wise data
        for (const branchName in branchSummary) {
            const summary = branchSummary[branchName];
            const row = allBranchSnapshotTable.insertRow();
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = summary.employees.size;
            row.insertCell().textContent = summary.calls;
            row.insertCell().textContent = summary.visits;
            row.insertCell().textContent = summary.references;
            row.insertCell().textContent = summary.newCustomerLeads;
        }

        if (Object.keys(branchSummary).length === 0) {
            const row = allBranchSnapshotTable.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 6; // Span all columns
            cell.textContent = "No data available for the selected filters.";
            cell.style.textAlign = "center";
        }
    }


    // Renders the All Staff Overall Performance report
    function renderAllStaffOverallPerformance(data) {
        allStaffOverallPerformanceTable.innerHTML = ''; // Clear previous data
        const filteredData = getFilteredData(data);

        const employeePerformance = {};
        const allEmployees = new Set(); // Keep track of all employees to show non-participating ones

        filteredData.forEach(row => {
            const employeeCode = row[HEADER_EMPLOYEE_CODE];
            const employeeName = row[HEADER_EMPLOYEE_NAME];
            const designation = row[HEADER_DESIGNATION];
            const branchName = row[HEADER_BRANCH_NAME];

            allEmployees.add(employeeCode); // Add all employees for overall list

            if (!employeePerformance[employeeCode]) {
                employeePerformance[employeeCode] = {
                    name: employeeName,
                    designation: designation,
                    branch: branchName,
                    calls: 0,
                    visits: 0,
                    references: 0,
                    newCustomerLeads: 0, // Will be summed from pre-calculated field
                    remarks: []
                };
            }

            employeePerformance[employeeCode].calls += row[HEADER_CALL];
            employeePerformance[employeeCode].visits += row[HEADER_VISIT];
            employeePerformance[employeeCode].references += row[HEADER_REFERENCE];
            employeePerformance[employeeCode].newCustomerLeads += row.newCustomerLeads; // Sum pre-calculated field
            if (row[HEADER_REMARK]) {
                employeePerformance[employeeCode].remarks.push(row[HEADER_REMARK]);
            }
        });

        // Get all known employees from the canvassing data, even if they have no current activities
        const allKnownEmployeeData = new Map(); // employeeCode -> {name, designation, branch}
        canvassingData.forEach(row => {
            if (!allKnownEmployeeData.has(row[HEADER_EMPLOYEE_CODE])) {
                allKnownEmployeeData.set(row[HEADER_EMPLOYEE_CODE], {
                    name: row[HEADER_EMPLOYEE_NAME],
                    designation: row[HEADER_DESIGNATION],
                    branch: row[HEADER_BRANCH_NAME]
                });
            }
        });

        // Merge and display all employees, including those with zero activities
        for (const empCode of allKnownEmployeeData.keys()) {
            const empDetails = allKnownEmployeeData.get(empCode);
            const performance = employeePerformance[empCode] || {
                name: empDetails.name,
                designation: empDetails.designation,
                branch: empDetails.branch,
                calls: 0,
                visits: 0,
                references: 0,
                newCustomerLeads: 0,
                remarks: []
            };

            // Apply branch/employee filter if present for overall staff performance
            if (branchSelect.value && performance.branch !== branchSelect.value) {
                continue;
            }
            if (employeeSelect.value && empCode !== employeeSelect.value) {
                continue;
            }

            const row = allStaffOverallPerformanceTable.insertRow();
            row.insertCell().textContent = performance.branch;
            row.insertCell().textContent = performance.name || 'N/A';
            row.insertCell().textContent = empCode;
            row.insertCell().textContent = performance.designation || 'N/A';
            row.insertCell().textContent = performance.calls;
            row.insertCell().textContent = performance.visits;
            row.insertCell().textContent = performance.references;
            row.insertCell().textContent = performance.newCustomerLeads;
            row.insertCell().textContent = performance.remarks.join('; '); // Join remarks with a semicolon
        }

        if (Object.keys(employeePerformance).length === 0 && getFilteredData(data).length === 0) {
            const row = allStaffOverallPerformanceTable.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 9; // Span all columns
            cell.textContent = "No data available for the selected filters.";
            cell.style.textAlign = "center";
        }
    }


    // Renders the Non-Participating Branches report
    function renderNonParticipatingBranches(data) {
        nonParticipatingBranchesList.innerHTML = ''; // Clear previous data

        const allBranches = new Set(data.map(row => row[HEADER_BRANCH_NAME]).filter(Boolean));
        const participatingBranches = new Set(data.filter(row => row[HEADER_CALL] > 0 || row[HEADER_VISIT] > 0 || row[HEADER_REFERENCE] > 0).map(row => row[HEADER_BRANCH_NAME]));

        const nonParticipating = [...allBranches].filter(branch => !participatingBranches.has(branch));

        if (nonParticipating.length > 0) {
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        } else {
            const messageDiv = document.createElement('div');
            messageDiv.className = 'no-participation-message-container';
            messageDiv.innerHTML = '<p class="no-participation-message">🎉 All branches have participation! 🎉</p>';
            nonParticipatingBranchesList.appendChild(messageDiv);
        }
    }

    // Shows the selected report tab and hides others
    function showTab(tabId) {
        // Remove 'active' class from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all report sections
        document.querySelectorAll('.report-section > div').forEach(section => {
            section.style.display = 'none';
        });

        // Show controls panel by default, then hide if employee management is selected
        document.querySelector('.controls-panel').style.display = 'flex';

        // Add 'active' class to the clicked tab button
        document.getElementById(tabId).classList.add('active');

        // Show the corresponding section
        if (tabId === 'allBranchSnapshotTabBtn') {
            document.getElementById('allBranchSnapshotSection').style.display = 'block';
            renderAllBranchSnapshot(canvassingData);
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            document.getElementById('allStaffOverallPerformanceSection').style.display = 'block';
            renderAllStaffOverallPerformance(canvassingData);
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            document.getElementById('nonParticipatingBranchesSection').style.display = 'block';
            renderNonParticipatingBranches(canvassingData);
            document.querySelector('.controls-panel').style.display = 'none'; // Hide filters for this tab
        } else if (tabId === 'employeeManagementTabBtn') {
            document.getElementById('employeeManagementSection').style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'none'; // Hide filters for this tab
        }
    }

    // --- Event Listeners for Tabs ---
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // --- Event Listeners for Filters ---
    branchSelect.addEventListener('change', () => {
        populateEmployeeFilter(canvassingData); // Repopulate employee filter based on branch
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id); // Re-render current tab with new filter
        }
    });

    employeeSelect.addEventListener('change', () => {
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton) {
            showTab(activeTabButton.id); // Re-render current tab with new filter
        }
    });

    // --- Google Apps Script Integration ---
    async function sendDataToGoogleAppsScript(action, data) {
        try {
            displayStatusMessage("Sending data...", false);
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'no-cors', // Required for Google Apps Script Web Apps when not using CORS headers
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                }).toString()
            });

            // For 'no-cors', response.ok will always be false, and response.status will be 0.
            // We rely on the Apps Script to confirm success (e.g., by updating the sheet
            // which will then be reflected on the next fetchData call).
            displayStatusMessage("Data sent. Refreshing data shortly...", false);
            return true; // Assume success for no-cors
        } catch (error) {
            console.error("Error sending data to Apps Script:", error);
            displayStatusMessage("Error sending data. Please check console.", true);
            return false;
        }
    }

    // --- Event Listener for Add Employee Form ---
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeNameInput = document.getElementById('employeeName');
            const employeeCodeInput = document.getElementById('employeeCode');
            const employeeDesignationInput = document.getElementById('employeeDesignation');
            const employeeBranchNameInput = document.getElementById('employeeBranchName');

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [HEADER_DESIGNATION]: employeeDesignationInput.value.trim(),
                [HEADER_BRANCH_NAME]: employeeBranchNameInput.value.trim()
            };

            // Basic validation
            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_DESIGNATION] || !employeeData[HEADER_BRANCH_NAME]) {
                displayEmployeeManagementMessage('All fields are required.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data
            }
        });
    }

    // --- Event Listener for Bulk Add Employee Form ---
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const details = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !details) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeesToAdd = [];
            const detailLines = details.split('\n').map(line => line.trim()).filter(line => line.length > 0);

            detailLines.forEach(line => {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length === 3) {
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_DESIGNATION]: parts[2],
                        [HEADER_BRANCH_NAME]: branchName
                    };
                    employeesToAdd.push(employeeData);
                } else {
                    displayEmployeeManagementMessage(`Skipping invalid line: "${line}". Format: Name,Code,Designation`, true);
                }
            });

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
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
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
