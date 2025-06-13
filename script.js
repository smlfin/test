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
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // Marked as UNUSED for clarity, won't be fetched for reports

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
    const HEADER_HOW_CONTACTED = 'How Contacted'; // This is not in the list provided by user, but is in the original script. Keeping it.
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)'; // Keeping user's provided typo
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested'; // Keeping user's provided typo
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';
    // NEW: Customer Detail Headers as provided by user
    const HEADER_FAMILY_DETAILS_1 = 'Family Deatils -1 Name of wife/Husband';
    const HEADER_FAMILY_DETAILS_2 = 'Family Deatils -2 Job of wife/Husband';
    const HEADER_FAMILY_DETAILS_3 = 'Family Deatils -3 Names of Children';
    const HEADER_FAMILY_DETAILS_4 = 'Family Deatils -4 Deatils of Children';
    const HEADER_PROFILE_OF_CUSTOMER = 'Profile of Customer';


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area
    const reportDisplay = document.getElementById('reportDisplay');
    // Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const followupDueTabBtn = document.getElementById('followupDueTabBtn'); // NEW
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const followupDueSection = document.getElementById('followupDueSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');

    // NEW: Followup Due Elements
    const followupBranchSelect = document.getElementById('followupBranchSelect');
    const followupEmployeeSelect = document.getElementById('followupEmployeeSelect');
    const followupDateInput = document.getElementById('followupDateInput');
    const filterFollowupBtn = document.getElementById('filterFollowupBtn');
    const resetFollowupBtn = document.getElementById('resetFollowupBtn');
    const followupDueDisplay = document.getElementById('followupDueDisplay');


    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES
    let allUniqueEmployees = []; // Employee codes from Canvassing Data
    let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
    let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data
    let selectedBranchEntries = []; // Activity entries filtered by branch (for main reports section)
    let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code (for main reports section)


    // Utility to format date to ISO-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div (now targets statusMessageDiv)
    function displayMessage(message, type = 'info') {
        if (statusMessageDiv) {
            statusMessageDiv.innerHTML = `<div class="message ${type}">${message}</div>`;
            statusMessageDiv.style.display = 'block';
            setTimeout(() => {
                statusMessageDiv.innerHTML = ''; // Clear message
                statusMessageDiv.style.display = 'none';
            }, 5000); // Hide after 5 seconds
        }
    }

    // Specific message display for employee management forms
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessage) {
            employeeManagementMessage.textContent = message;
            employeeManagementMessage.style.color = isError ? 'red' : 'green';
            employeeManagementMessage.style.display = 'block';
            setTimeout(() => {
                employeeManagementMessage.style.display = 'none';
                employeeManagementMessage.textContent = ''; // Clear content
            }, 5000);
        }
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

    // CSV parsing function (handles commas within quoted strings)
    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = parseCSVLine(lines[0]); // Headers can also contain commas in quotes
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = parseCSVLine(lines[i]);
            if (values.length !== headers.length) {
                console.warn(`Skipping malformed row ${i + 1}: Expected ${headers.length} columns, got ${values.length}. Line: "${lines[i]}"`);
                continue;
            }
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index];
            });
            data.push(entry);
        }
        return data;
    }

    // Helper to parse a single CSV line safely
    function parseCSVLine(line) {
        const result = [];
        let inQuote = false;
        let currentField = '';
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuote = !inQuote;
            } else if (char === ',' && !inQuote) {
                result.push(currentField.trim());
                currentField = '';
            } else {
                currentField += char;
            }
        }
        result.push(currentField.trim());
        return result;
    }


    // Process fetched data to populate filters and prepare for reports
    async function processData() {
        // Only fetch canvassing data, ignoring MasterEmployees for front-end reports
        await fetchCanvassingData();

        // Re-initialize allUniqueBranches from the predefined list
        allUniqueBranches = [...PREDEFINED_BRANCHES].sort(); // Use the hardcoded list

        // Populate employeeCodeToNameMap and employeeCodeToDesignationMap ONLY from Canvassing Data
        employeeCodeToNameMap = {}; // Reset map before populating
        employeeCodeToDesignationMap = {}; // Reset map before populating
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (employeeCode) {
                // If an employee code exists in canvassing data, use its name/designation
                employeeCodeToNameMap[employeeCode] = employeeName || employeeCode;
                employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
            }
        });

        // Re-populate allUniqueEmployees based ONLY on canvassing data
        allUniqueEmployees = [...new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
            // Use the name from the map if available, otherwise use the code for sorting and display
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches); // Populate branch dropdown with predefined branches
        populateDropdown(customerViewBranchSelect, allUniqueBranches); // Populate branch dropdown for detailed customer view
        populateDropdown(followupBranchSelect, allUniqueBranches); // Populate branch dropdown for followup view
        populateDropdown(followupEmployeeSelect, allUniqueEmployees, true); // Populate employee dropdown for followup view
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot(); // Render the default "All Branch Snapshot" report
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items, useCodeForValue = false) {
        selectElement.innerHTML = '<option value="">-- Select --</option>'; // Default option
        items.forEach(item => {
            const option = document.createElement('option');
            if (useCodeForValue) {
                // Display name from map or code itself
                option.value = item; // item is employeeCode
                option.textContent = employeeCodeToNameMap[item] || item;
            } else {
                option.value = item; // item is branch name
                option.textContent = item;
            }
            selectElement.appendChild(option);
        });
    }

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';

            // Get employee codes ONLY from Canvassing Data for the selected branch
            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            // Combine and unique all employee codes for the selected branch
            const combinedEmployeeCodes = new Set([
                ...employeeCodesInBranchFromCanvassing
            ]);

            // Convert Set back to array and sort
            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                // Use the name from the map if available, otherwise use the code for sorting and display
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

            populateDropdown(employeeSelect, sortedEmployeeCodesInBranch, true);
            viewOptions.style.display = 'flex'; // Show view options
            // Reset employee selection and employee-specific display when branch changes
            employeeSelect.value = "";
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';

            // Deactivate all buttons in viewOptions and then reactivate the appropriate ones
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));

        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none'; // Hide view options
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown above to view reports.</p>';
            selectedBranchEntries = []; // Clear previous activity filter
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
        }
    });

    // Handle employee selection (now based on employee CODE)
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // Filter activity data by employee code (from allCanvassingData)
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_BRANCH_NAME] === branchSelect.value // Filter by selected branch as well
            );
            const employeeDisplayName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
            reportDisplay.innerHTML = `<p>Ready to view reports for ${employeeDisplayName}.</p>`;

            // Automatically trigger the Employee Summary (d4.PNG style)
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewEmployeeSummaryBtn.classList.add('active'); // Set Employee Summary as active
            renderEmployeeSummary(selectedEmployeeCodeEntries); // Render the Employee Summary

        } else {
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.S</p>';
            // Clear active button if employee selection is cleared
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        }
    });

    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // Initialize counters
        const productInterests = new Set(); // To collect unique product interests

        console.log('Calculating total activity for entries:', entries.length); // Log entries being processed
        entries.forEach((entry, index) => {
            let activityType = entry[HEADER_ACTIVITY_TYPE];
            let typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];
            let productInterested = entry[HEADER_PRODUCT_INTERESTED]; // Get product interested

            // Trim and convert to lowercase for robust comparison
            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';
            const trimmedProductInterested = productInterested ? productInterested.trim() : ''; // Don't lowercase products unless explicitly asked

            console.log(`--- Entry ${index + 1} Debug ---`);
            console.log(`  Processed Activity Type (trimmed, lowercase): '${trimmedActivityType}'`);
            console.log(`  Processed Type of Customer (trimmed, lowercase): '${trimmedTypeOfCustomer}'`);
            console.log(`  Processed Product Interested (trimmed): '${trimmedProductInterested}'`);


            // Direct matching to user's provided sheet values (now lowercase)
            if (trimmedActivityType === 'visit') {
                totalActivity['Visit']++;
            } else if (trimmedActivityType === 'calls') { // Matches "Calls" from sheet, now lowercase
                totalActivity['Call']++;
            } else if (trimmedActivityType === 'referance') { // Matches "Referance" (with typo) from sheet, now lowercase
                totalActivity['Reference']++;
            } else {
                // If it's not one of the direct activity types, log for debugging
                console.warn(`  Unknown or unhandled Activity Type encountered (trimmed, lowercase): '${trimmedActivityType}'.`);
            }

            // --- UPDATED LOGIC FOR 'New Customer Leads' ---
            // Based on the user's previously working script, New Customer Leads are counted
            // if the 'Type of Customer' (now correctly spelled) is simply 'new', regardless of 'Activity Type'.
            if (trimmedTypeOfCustomer === 'new') {
                totalActivity['New Customer Leads']++;
                console.log(`  New Customer Lead INCREMENTED based on Type of Customer === 'new'.`);
            } else {
                console.log(`  New Customer Lead NOT INCREMENTED: Type of Customer is not 'new'.`);
            }
            // --- END UPDATED LOGIC ---

            // Collect unique product interests
            if (trimmedProductInterested) {
                productInterests.add(trimmedProductInterested);
            }
            console.log(`--- End Entry ${index + 1} Debug ---`);
        });
        console.log('Calculated Total Activity Final:', totalActivity);

        // Return both total activities and product interests
        return { totalActivity, productInterests: [...productInterests] };
    }

    // Render All Branch Snapshot (now uses PREDEFINED_BRANCHES and checks for participation)
    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';

        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = ['Branch Name', 'Employees with Activity', 'Total Visits', 'Total Calls', 'Total References', 'Total New Customer Leads'];
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries); // Destructure to get totalActivity
            const employeeCodesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))];
            const displayEmployeeCount = employeeCodesInBranch.length;

            const row = tbody.insertRow();
            // Assign data-label for mobile view
            row.insertCell().setAttribute('data-label', 'Branch Name');
            row.lastChild.textContent = branch;

            row.insertCell().setAttribute('data-label', 'Employees with Activity');
            row.lastChild.textContent = displayEmployeeCount;

            row.insertCell().setAttribute('data-label', 'Total Visits');
            row.lastChild.textContent = totalActivity['Visit'];

            row.insertCell().setAttribute('data-label', 'Total Calls');
            row.lastChild.textContent = totalActivity['Call'];

            row.insertCell().setAttribute('data-label', 'Total References');
            row.lastChild.textContent = totalActivity['Reference'];

            row.insertCell().setAttribute('data-label', 'Total New Customer Leads');
            row.lastChild.textContent = totalActivity['New Customer Leads'];
        });

        reportDisplay.appendChild(table);
    }

    // NEW: Render Non-Participating Branches Report
    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';
        const nonParticipatingBranches = [];

        PREDEFINED_BRANCHES.forEach(branch => {
            const hasActivity = allCanvassingData.some(entry => entry[HEADER_BRANCH_NAME] === branch);
            if (!hasActivity) {
                nonParticipatingBranches.push(branch);
            }
        });

        if (nonParticipatingBranches.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-branch-list';
            nonParticipatingBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded activity!</p>';
        }
    }

    // Render All Staff Overall Performance Report (for d1.PNG)
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>Overall Staff Performance Report (This Month)</h2>';
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling

        const table = document.createElement('table');
        table.className = 'performance-table';

        const thead = table.createTHead();
        let headerRow = thead.insertRow();

        // Main Headers
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Branch Name';
        headerRow.insertCell().textContent = 'Designation';

        // Define metrics for the performance table
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // 'Actual', 'Target', '%'
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        // Sub-headers
        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty cell for Employee Name
        headerRow.insertCell(); // Empty cell for Branch Name
        headerRow.insertCell(); // Empty cell for Designation
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Get unique employees who have made at least one entry this month
        const employeesWithActivityThisMonth = [...new Set(allCanvassingData
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });


        if (employeesWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p>No employee activity found for the current month.</p>';
            return;
        }

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );
            const { totalActivity } = calculateTotalActivity(employeeActivities);

            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;
            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0; // Ensure target is 0 if undefined
                let percentValue = performance[metric];
                let displayPercent;
                let progressBarClass;
                let progressWidth;
                if (isNaN(percentValue) || targetValue === 0) { // If target is 0, it's N/A
                    displayPercent = 'N/A';
                    progressWidth = 0;
                    progressBarClass = 'no-activity';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                    progressWidth = Math.min(100, Math.round(percentValue));
                    progressBarClass = getProgressBarClass(percentValue);
                }
                // Special handling for 0 actuals with positive targets to show 0% and danger color
                if (actualValue === 0 && targetValue > 0) {
                    displayPercent = '0%';
                    progressWidth = 0;
                    progressBarClass = 'danger';
                }
                row.insertCell().textContent = actualValue;
                row.insertCell().textContent = targetValue;
                const percentCell = row.insertCell();
                percentCell.innerHTML = `
                    <div class="progress-bar-container-small">
                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth === 0 && displayPercent !== 'N/A' ? '30px' : progressWidth}%">
                            ${displayPercent}
                        </div>
                    </div>
                `;
            });
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Function to calculate performance percentage
    function calculatePerformance(actuals, targets) {
        const performance = {};
        for (const metric in targets) {
            const actual = actuals[metric] || 0;
            const target = targets[metric];
            if (target > 0) {
                performance[metric] = (actual / target) * 100;
            } else {
                performance[metric] = NaN; // Or 0, depending on how you want to handle no target
            }
        }
        return performance;
    }

    // Helper to determine progress bar class based on percentage
    function getProgressBarClass(percentage) {
        if (percentage >= 100) return 'success';
        if (percentage >= 75) return 'warning-high';
        if (percentage >= 50) return 'warning-medium';
        if (percentage > 0) return 'warning-low';
        return 'danger';
    }

    // Function to render Branch Performance Report (d3.PNG)
    function renderBranchPerformanceReport(branchName) {
        reportDisplay.innerHTML = `<h2>Branch Performance Report: ${branchName} (This Month)</h2>`;
        const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchName);
        if (branchActivityEntries.length === 0) {
            reportDisplay.innerHTML += '<p>No activity found for this branch this month.</p>';
            return;
        }

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const employeesInBranch = [...new Set(branchActivityEntries
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

        if (employeesInBranch.length === 0) {
            reportDisplay.innerHTML += '<p>No employee activity found for this branch for the current month.</p>';
            return;
        }

        const branchPerformanceGrid = document.createElement('div');
        branchPerformanceGrid.className = 'branch-performance-grid';

        employeesInBranch.forEach(employeeCode => {
            const employeeActivities = branchActivityEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const { totalActivity } = calculateTotalActivity(employeeActivities); // Destructure
            const employeeDisplayName = employeeCodeToNameMap[employeeCode] || employeeCode; // Use name from map or code
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const employeeCard = document.createElement('div');
            employeeCard.className = 'employee-performance-card';
            employeeCard.innerHTML = `
                <h4>${employeeDisplayName} (${designation})</h4>
                <div style="overflow-x: auto;">
                    <table class="performance-table">
                        <thead>
                            <tr>
                                <th>Metric</th>
                                <th>Act</th>
                                <th>Tgt</th>
                                <th>%</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Visit</td>
                                <td>${totalActivity['Visit'] || 0}</td>
                                <td>${targets['Visit'] || 0}</td>
                                <td>
                                    <div class="progress-bar-container-small">
                                        <div class="progress-bar ${getProgressBarClass(performance['Visit'])}" style="width: ${isNaN(performance['Visit']) || targets['Visit'] === 0 ? '0' : Math.min(100, Math.round(performance['Visit']))}%">
                                            ${isNaN(performance['Visit']) || targets['Visit'] === 0 ? 'N/A' : `${Math.round(performance['Visit'])}%`}
                                        </div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td>Call</td>
                                <td>${totalActivity['Call'] || 0}</td>
                                <td>${targets['Call'] || 0}</td>
                                <td>
                                    <div class="progress-bar-container-small">
                                        <div class="progress-bar ${getProgressBarClass(performance['Call'])}" style="width: ${isNaN(performance['Call']) || targets['Call'] === 0 ? '0' : Math.min(100, Math.round(performance['Call']))}%">
                                            ${isNaN(performance['Call']) || targets['Call'] === 0 ? 'N/A' : `${Math.round(performance['Call'])}%`}
                                        </div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td>Reference</td>
                                <td>${totalActivity['Reference'] || 0}</td>
                                <td>${targets['Reference'] || 0}</td>
                                <td>
                                    <div class="progress-bar-container-small">
                                        <div class="progress-bar ${getProgressBarClass(performance['Reference'])}" style="width: ${isNaN(performance['Reference']) || targets['Reference'] === 0 ? '0' : Math.min(100, Math.round(performance['Reference']))}%">
                                            ${isNaN(performance['Reference']) || targets['Reference'] === 0 ? 'N/A' : `${Math.round(performance['Reference'])}%`}
                                        </div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td>New Customer Leads</td>
                                <td>${totalActivity['New Customer Leads'] || 0}</td>
                                <td>${targets['New Customer Leads'] || 0}</td>
                                <td>
                                    <div class="progress-bar-container-small">
                                        <div class="progress-bar ${getProgressBarClass(performance['New Customer Leads'])}" style="width: ${isNaN(performance['New Customer Leads']) || targets['New Customer Leads'] === 0 ? '0' : Math.min(100, Math.round(performance['New Customer Leads']))}%">
                                            ${isNaN(performance['New Customer Leads']) || targets['New Customer Leads'] === 0 ? 'N/A' : `${Math.round(performance['New Customer Leads'])}%`}
                                        </div>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            `;
            branchPerformanceGrid.appendChild(employeeCard);
        });
        reportDisplay.appendChild(branchPerformanceGrid);
    }

    // Function to render Employee Summary (d4.PNG)
    function renderEmployeeSummary(employeeActivities) {
        if (!employeeActivities || employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No activities found for this employee this month.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const branchName = employeeActivities[0][HEADER_BRANCH_NAME];
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        const { totalActivity, productInterests } = calculateTotalActivity(employeeActivities);
        const targets = TARGETS[designation] || TARGETS['Default'];

        reportDisplay.innerHTML = `
            <h2>Employee Summary: ${employeeName} (${designation})</h2>
            <div class="summary-breakdown-card">
                <div class="summary-section">
                    <h3>Summary (This Month)</h3>
                    <table>
                        <tr><th>Metric</th><th>Actual</th><th>Target</th><th>%</th></tr>
                        <tr>
                            <td>Visits</td>
                            <td>${totalActivity['Visit'] || 0}</td>
                            <td>${targets['Visit'] || 0}</td>
                            <td>${isNaN(calculatePerformance(totalActivity, targets)['Visit']) ? 'N/A' : `${Math.round(calculatePerformance(totalActivity, targets)['Visit'])}%`}</td>
                        </tr>
                        <tr>
                            <td>Calls</td>
                            <td>${totalActivity['Call'] || 0}</td>
                            <td>${targets['Call'] || 0}</td>
                            <td>${isNaN(calculatePerformance(totalActivity, targets)['Call']) ? 'N/A' : `${Math.round(calculatePerformance(totalActivity, targets)['Call'])}%`}</td>
                        </tr>
                        <tr>
                            <td>References</td>
                            <td>${totalActivity['Reference'] || 0}</td>
                            <td>${targets['Reference'] || 0}</td>
                            <td>${isNaN(calculatePerformance(totalActivity, targets)['Reference']) ? 'N/A' : `${Math.round(calculatePerformance(totalActivity, targets)['Reference'])}%`}</td>
                        </tr>
                        <tr>
                            <td>New Customer Leads</td>
                            <td>${totalActivity['New Customer Leads'] || 0}</td>
                            <td>${targets['New Customer Leads'] || 0}</td>
                            <td>${isNaN(calculatePerformance(totalActivity, targets)['New Customer Leads']) ? 'N/A' : `${Math.round(calculatePerformance(totalActivity, targets)['New Customer Leads'])}%`}</td>
                        </tr>
                    </table>
                </div>
                <div class="summary-section">
                    <h3>Product Interests</h3>
                    <ul class="product-interest-list">
                        ${productInterests.length > 0 ? productInterests.map(p => `<li>${p}</li>`).join('') : '<li>No product interests recorded this month.</li>'}
                    </ul>
                </div>
            </div>
        `;
    }

    // Function to render all entries for a selected employee
    function renderAllEmployeeEntries(employeeActivities) {
        if (!employeeActivities || employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No entries found for this employee.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        reportDisplay.innerHTML = `<h2>All Entries for ${employeeName}</h2>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'customer-details-table'; // Re-using table style

        const thead = table.createTHead();
        const headerRow = thead.insertRow();

        // Dynamically get headers from the first entry to ensure all columns are displayed
        const headers = Object.keys(employeeActivities[0]).filter(header =>
            header !== HEADER_TIMESTAMP && // Exclude timestamp as it's typically internal
            header !== HEADER_FAMILY_DETAILS_1 && // Exclude family details as they are for specific customer view
            header !== HEADER_FAMILY_DETAILS_2 &&
            header !== HEADER_FAMILY_DETAILS_3 &&
            header !== HEADER_FAMILY_DETAILS_4 &&
            header !== HEADER_PROFILE_OF_CUSTOMER
        );

        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        employeeActivities.sort((a, b) => new Date(b[HEADER_DATE]) - new Date(a[HEADER_DATE])).forEach(entry => {
            const row = tbody.insertRow();
            headers.forEach(header => {
                const cell = row.insertCell();
                let value = entry[header] || '';
                if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    value = formatDate(value); // Format dates consistently
                }
                cell.textContent = value;
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Function to render Employee Performance Report (d2.PNG)
    function renderEmployeePerformanceReport(employeeActivities) {
        if (!employeeActivities || employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No activities found for this employee.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        reportDisplay.innerHTML = `<h2>Employee Performance Report: ${employeeName} (${designation})</h2>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'performance-table';

        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Date';
        headerRow.insertCell().textContent = 'Activity Type';
        headerRow.insertCell().textContent = 'Customer Type';
        headerRow.insertCell().textContent = 'Prospect Name';
        headerRow.insertCell().textContent = 'Remarks';
        headerRow.insertCell().textContent = 'Next Follow-up Date';

        const tbody = table.createTBody();

        // Sort entries by date descending for the report
        employeeActivities.sort((a, b) => new Date(b[HEADER_DATE]) - new Date(a[HEADER_DATE]));

        employeeActivities.forEach(entry => {
            const row = tbody.insertRow();
            row.insertCell().textContent = formatDate(entry[HEADER_DATE] || '');
            row.insertCell().textContent = entry[HEADER_ACTIVITY_TYPE] || '';
            row.insertCell().textContent = entry[HEADER_TYPE_OF_CUSTOMER] || '';
            row.insertCell().textContent = entry[HEADER_PROSPECT_NAME] || '';
            row.insertCell().textContent = entry[HEADER_REMARKS] || '';
            row.insertCell().textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE] || '');
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // NEW: Function to render customer details (d5.PNG)
    function renderEmployeeCustomerDetails(customerEntry) {
        if (!customerEntry) {
            customerDetailsContent.innerHTML = '<p>Select a customer to view details.</p>';
            customerDetailsContent.style.display = 'none';
            return;
        }

        customerDetailsContent.style.display = 'block';
        customerDetailsContent.innerHTML = `
            <h3>Customer Details: ${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</h3>
            <div class="detail-row"><span class="detail-label">Phone Number (Whatsapp):</span> <span class="detail-value">${customerEntry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Address:</span> <span class="detail-value">${customerEntry[HEADER_ADDRESS] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Profession:</span> <span class="detail-value">${customerEntry[HEADER_PROFESSION] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">DOB/WD:</span> <span class="detail-value">${customerEntry[HEADER_DOB_WD] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Product Interested:</span> <span class="detail-value">${customerEntry[HEADER_PRODUCT_INTERESTED] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Next Follow-up Date:</span> <span class="detail-value">${formatDate(customerEntry[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Relation With Staff:</span> <span class="detail-value">${customerEntry[HEADER_RELATION_WITH_STAFF] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${customerEntry[HEADER_REMARKS] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details 1 (Wife/Husband Name):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_1] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details 2 (Wife/Husband Job):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_2] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details 3 (Children Names):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_3] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details 4 (Children Details):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_4] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Profile of Customer:</span> <span class="detail-value">${customerEntry[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</span></div>
        `;
    }

    // NEW: Render Followup Due Report
    function renderFollowupDueReport(branch = '', employeeCode = '', dueDate = '') {
        console.log(`Rendering followup due report for branch: ${branch}, employeeCode: ${employeeCode}, dueDate: ${dueDate}`);
        followupDueDisplay.innerHTML = '<h3>Loading Follow-ups...</h3>';

        let filteredFollowups = allCanvassingData;

        // Filter by due date first
        if (dueDate) {
            filteredFollowups = filteredFollowups.filter(entry => {
                const followUpDate = entry[HEADER_NEXT_FOLLOW_UP_DATE];
                // Ensure both dates are valid and in 'YYYY-MM-DD' format for comparison
                return formatDate(followUpDate) === dueDate;
            });
            console.log(`Filtered by due date (${dueDate}):`, filteredFollowups.length, 'entries');
        } else {
            // If no specific date is selected, default to today's follow-ups
            const today = new Date();
            const todayFormatted = today.toISOString().split('T')[0];
            filteredFollowups = filteredFollowups.filter(entry => {
                const followUpDate = entry[HEADER_NEXT_FOLLOW_UP_DATE];
                return formatDate(followUpDate) === todayFormatted;
            });
            console.log(`Filtered by today's date (default):`, filteredFollowups.length, 'entries');
        }

        // Further filter by branch
        if (branch) {
            filteredFollowups = filteredFollowups.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            console.log(`Filtered by branch (${branch}):`, filteredFollowups.length, 'entries');
        }

        // Further filter by employee
        if (employeeCode) {
            filteredFollowups = filteredFollowups.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            console.log(`Filtered by employee (${employeeCode}):`, filteredFollowups.length, 'entries');
        }


        if (filteredFollowups.length === 0) {
            followupDueDisplay.innerHTML = '<p>No follow-ups due found for the selected criteria.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'customer-details-table'; // Re-using customer-details-table class

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_EMPLOYEE_CODE,
            HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER, HEADER_REMARKS, HEADER_NEXT_FOLLOW_UP_DATE
        ];

        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        filteredFollowups.forEach(entry => {
            const row = tbody.insertRow();
            headers.forEach(header => {
                const cell = row.insertCell();
                let value = entry[header] || '';
                // Format Date and Next Follow-up Date
                if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    value = formatDate(value);
                }
                cell.textContent = value;
            });
        });

        followupDueDisplay.innerHTML = ''; // Clear loading message
        followupDueDisplay.appendChild(table);
    }


    // Function to handle tab switching
    function showTab(tabId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        followupDueSection.style.display = 'none'; // NEW
        employeeManagementSection.style.display = 'none';


        // Activate the selected tab button and show its content
        document.getElementById(tabId).classList.add('active');
        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                reportsSection.style.display = 'block';
                renderAllBranchSnapshot();
                break;
            case 'allStaffOverallPerformanceTabBtn':
                reportsSection.style.display = 'block';
                renderOverallStaffPerformanceReport();
                break;
            case 'nonParticipatingBranchesTabBtn':
                reportsSection.style.display = 'block';
                renderNonParticipatingBranches();
                break;
            case 'detailedCustomerViewTabBtn':
                detailedCustomerViewSection.style.display = 'block';
                // Reset dropdowns for detailed customer view
                customerViewBranchSelect.value = '';
                customerViewEmployeeSelect.innerHTML = '<option value="">-- Select --</option>';
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
                customerDetailsContent.style.display = 'none';
                break;
            case 'followupDueTabBtn':
                followupDueSection.style.display = 'block';
                // Set today's date as default for followupDateInput
                const today = new Date();
                followupDateInput.value = today.toISOString().split('T')[0];
                renderFollowupDueReport(); // Initial render for today's followups
                break;
            case 'employeeManagementTabBtn':
                employeeManagementSection.style.display = 'block';
                break;
        }
    }


    // --- Event Listeners for Tab Buttons ---
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
    if (followupDueTabBtn) {
        followupDueTabBtn.addEventListener('click', () => {
            showTab('followupDueTabBtn');
            // Set today's date as default for followupDateInput
            const today = new Date();
            followupDateInput.value = today.toISOString().split('T')[0];
            renderFollowupDueReport(); // Initial render for today's followups
        });
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }

    // --- Event Listeners for Main Report Section (reportsSection) ---
    if (viewBranchPerformanceReportBtn) {
        viewBranchPerformanceReportBtn.addEventListener('click', () => {
            const selectedBranch = branchSelect.value;
            if (selectedBranch) {
                document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
                viewBranchPerformanceReportBtn.classList.add('active');
                renderBranchPerformanceReport(selectedBranch);
            } else {
                displayMessage('Please select a branch first.', 'error');
            }
        });
    }

    if (viewEmployeeSummaryBtn) {
        viewEmployeeSummaryBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value;
            if (selectedEmployeeCode && selectedEmployeeCodeEntries.length > 0) {
                document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
                viewEmployeeSummaryBtn.classList.add('active');
                renderEmployeeSummary(selectedEmployeeCodeEntries);
            } else {
                displayMessage('Please select an employee first.', 'error');
            }
        });
    }

    if (viewAllEntriesBtn) {
        viewAllEntriesBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value;
            if (selectedEmployeeCode && selectedEmployeeCodeEntries.length > 0) {
                document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
                viewAllEntriesBtn.classList.add('active');
                renderAllEmployeeEntries(selectedEmployeeCodeEntries);
            } else {
                displayMessage('Please select an employee first.', 'error');
            }
        });
    }

    if (viewPerformanceReportBtn) {
        viewPerformanceReportBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value;
            if (selectedEmployeeCode && selectedEmployeeCodeEntries.length > 0) {
                document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
                viewPerformanceReportBtn.classList.add('active');
                renderEmployeePerformanceReport(selectedEmployeeCodeEntries);
            } else {
                displayMessage('Please select an employee first.', 'error');
            }
        });
    }

    // --- Event Listeners for Detailed Customer View Section (detailedCustomerViewSection) ---
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        customerDetailsContent.style.display = 'none'; // Hide details when branch changes
        if (selectedBranch) {
            const employeesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            const uniqueEmployeeCodes = [...new Set(employeesInBranchFromCanvassing)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, uniqueEmployeeCodes, true);
            customerCanvassedList.innerHTML = '<p>Select an employee to view canvassed customers.</p>';
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select --</option>';
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        customerDetailsContent.style.display = 'none'; // Hide details when employee changes

        if (selectedBranch && selectedEmployeeCode) {
            const employeeCustomers = allCanvassingData.filter(entry =>
                entry[HEADER_BRANCH_NAME] === selectedBranch &&
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
            );
            renderCustomerCanvassedList(employeeCustomers);
        } else {
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
        }
    });

    function renderCustomerCanvassedList(customerEntries) {
        customerCanvassedList.innerHTML = '';
        if (customerEntries.length === 0) {
            customerCanvassedList.innerHTML = '<p>No canvassed customers found for this employee in this branch.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';
        // Sort by prospect name for easier navigation
        customerEntries.sort((a, b) => (a[HEADER_PROSPECT_NAME] || '').localeCompare(b[HEADER_PROSPECT_NAME] || '')).forEach(entry => {
            const li = document.createElement('li');
            li.textContent = entry[HEADER_PROSPECT_NAME] || 'Unknown Customer';
            li.className = 'customer-list-item';
            li.addEventListener('click', () => renderEmployeeCustomerDetails(entry));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    // Event Listeners for Followup Due filters
    if (filterFollowupBtn) {
        filterFollowupBtn.addEventListener('click', () => {
            const branch = followupBranchSelect.value;
            const employeeCode = followupEmployeeSelect.value;
            const dueDate = followupDateInput.value;
            renderFollowupDueReport(branch, employeeCode, dueDate);
        });
    }

    if (resetFollowupBtn) {
        resetFollowupBtn.addEventListener('click', () => {
            followupBranchSelect.value = '';
            followupEmployeeSelect.value = '';
            const today = new Date();
            followupDateInput.value = today.toISOString().split('T')[0]; // Reset to today
            renderFollowupDueReport(); // Render with today's date and no other filters
        });
    }


    // --- Employee Management Form Submissions ---
    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault(); // Prevent default form submission

            const newEmployeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            // Basic validation
            if (!newEmployeeData[HEADER_EMPLOYEE_NAME] || !newEmployeeData[HEADER_EMPLOYEE_CODE] || !newEmployeeData[HEADER_BRANCH_NAME] || !newEmployeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding a new employee.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', newEmployeeData);

            if (success) {
                addEmployeeForm.reset(); // Clear form fields
                // Optionally, re-fetch data or update dropdowns if needed immediately
                // For simplicity, we assume data will be reloaded on next full report generation if needed.
            }
        });
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage(`Sending data to server for ${action}...`, false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for CORS
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data }),
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`Apps Script error for ${action}:`, response.status, errorText);
                displayEmployeeManagementMessage(`Error during ${action}: ${errorText || response.statusText}`, true);
                return false;
            }

            const result = await response.json();
            console.log(`Apps Script response for ${action}:`, result);

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(`${action.replace(/_/g, ' ')} successful!`, false);
                return true;
            } else {
                displayEmployeeManagementMessage(`Failed to ${action.replace(/_/g, ' ')}: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error(`Network or unexpected error during ${action}:`, error);
            displayEmployeeManagementMessage(`Network error during ${action}: ${error.message}`, true);
            return false;
        }
    }


    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName) {
                displayEmployeeManagementMessage('Branch Name is required for bulk entry.', true);
                return;
            }
            if (!employeeDetails) {
                displayEmployeeManagementMessage('Employee Details are required for bulk entry.', true);
                return;
            }

            const lines = employeeDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Expecting at least Name, Code
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
