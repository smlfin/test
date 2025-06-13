document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?output=csv"; 

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
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn'); // NEW
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    // NEW TABS
    const todaysFollowUpsTabBtn = document.getElementById('todaysFollowUpsTabBtn');
    const branchStatusTargetTabBtn = document.getElementById('branchStatusTargetTabBtn');


    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    // NEW SECTIONS
    const todaysFollowUpsSection = document.getElementById('todaysFollowUpsSection');
    const branchStatusTargetSection = document.getElementById('branchStatusTargetSection');

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');

    // NEW: Follow-up & Branch Status Display Elements
    const todaysFollowUpsDisplay = document.getElementById('todaysFollowUpsDisplay');
    const branchStatusTargetDisplay = document.getElementById('branchStatusTargetDisplay');


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
        let headerRow = thead.insertRow(); // Main Headers
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


        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            // Get the branch from any of their entries (assuming an employee is in one branch at a time for simplicity)
            const employeeBranch = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';

            // Filter activities for this employee for the current month
            const employeeMonthlyActivities = allCanvassingData.filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                       entryDate.getMonth() === currentMonth &&
                       entryDate.getFullYear() === currentYear;
            });

            const { totalActivity } = calculateTotalActivity(employeeMonthlyActivities);
            const employeeTargets = TARGETS[employeeDesignation] || TARGETS['Default'];

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = employeeBranch;
            row.insertCell().textContent = employeeDesignation;

            metrics.forEach(metric => {
                const actual = totalActivity[metric] || 0;
                const target = employeeTargets[metric] || 0;
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(1) : (actual > 0 ? 100 : 0).toFixed(1);

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                
                const percentageCell = row.insertCell();
                const progressBarContainer = document.createElement('div');
                progressBarContainer.className = 'progress-bar-container-small';
                const progressBar = document.createElement('div');
                progressBar.className = 'progress-bar';
                progressBar.style.width = `${Math.min(100, percentage)}%`; // Cap at 100% for visual
                progressBar.textContent = `${percentage}%`;

                if (percentage >= 100) {
                    progressBar.classList.add('success');
                } else if (percentage >= 75) {
                    progressBar.classList.add('warning-high');
                } else if (percentage >= 50) {
                    progressBar.classList.add('warning-medium');
                } else if (percentage >= 25) {
                    progressBar.classList.add('warning-low');
                } else {
                    progressBar.classList.add('danger');
                }
                
                if (actual === 0 && target === 0) { // No activity and no target
                     progressBar.classList.add('no-activity');
                     progressBar.textContent = 'N/A';
                     progressBar.style.width = '100%'; // Full grey bar
                }


                progressBarContainer.appendChild(progressBar);
                percentageCell.appendChild(progressBarContainer);
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);

        if (employeesWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No staff recorded activity this month.</p>';
        }
    }

    // Render Employee Summary (d4.PNG style)
    function renderEmployeeSummary(entries) {
        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity found for the selected employee in this branch.</p>';
            return;
        }

        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const employeeBranch = entries[0][HEADER_BRANCH_NAME]; // Get branch from one of their entries

        reportDisplay.innerHTML = `
            <h2>Summary for ${employeeName} (${employeeCode}) - ${employeeBranch}</h2>
            <div class="summary-breakdown-card">
                <div class="summary-section">
                    <h3>Employee Details</h3>
                    <p><strong>Designation:</strong> ${employeeDesignation}</p>
                    <p><strong>Branch:</strong> ${employeeBranch}</p>
                </div>
                <div class="summary-section">
                    <h3>Overall Activity (Current Month)</h3>
                    <div id="employeeActivitySummary"></div>
                </div>
                <div class="summary-section">
                    <h3>Product Interests (Recent)</h3>
                    <div id="employeeProductInterests"></div>
                </div>
            </div>
        `;

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Filter entries for the current month
        const monthlyEntries = entries.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        const { totalActivity, productInterests } = calculateTotalActivity(monthlyEntries);
        const employeeTargets = TARGETS[employeeDesignation] || TARGETS['Default'];

        const activitySummaryDiv = document.getElementById('employeeActivitySummary');
        const productInterestsDiv = document.getElementById('employeeProductInterests');

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const actual = totalActivity[metric] || 0;
            const target = employeeTargets[metric] || 0;
            const percentage = target > 0 ? ((actual / target) * 100).toFixed(1) : (actual > 0 ? 100 : 0).toFixed(1);

            const p = document.createElement('p');
            p.innerHTML = `<strong>${metric}:</strong> ${actual} / ${target} (${percentage}%)`;
            activitySummaryDiv.appendChild(p);

            const progressBarContainer = document.createElement('div');
            progressBarContainer.className = 'progress-bar-container';
            const progressBar = document.createElement('div');
            progressBar.className = 'progress-bar';
            progressBar.style.width = `${Math.min(100, percentage)}%`; // Cap at 100% for visual
            progressBar.textContent = `${percentage}%`;

            if (percentage >= 100) {
                progressBar.classList.add('success');
            } else if (percentage >= 75) {
                progressBar.classList.add('warning-high');
            } else if (percentage >= 50) {
                progressBar.classList.add('warning-medium');
            } else if (percentage >= 25) {
                progressBar.classList.add('warning-low');
            } else {
                progressBar.classList.add('danger');
            }

            if (actual === 0 && target === 0) { // No activity and no target
                progressBar.classList.add('no-activity');
                progressBar.textContent = 'N/A';
                progressBar.style.width = '100%'; // Full grey bar
            }

            progressBarContainer.appendChild(progressBar);
            activitySummaryDiv.appendChild(progressBarContainer);
        });

        if (productInterests.length > 0) {
            productInterestsDiv.innerHTML = `<p>${productInterests.join(', ')}</p>`;
        } else {
            productInterestsDiv.innerHTML = '<p>No specific product interests recorded.</p>';
        }
    }


    // Render All Entries for a specific Employee
    function renderAllEntries(entries) {
        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML = '<p>No entries found for the selected employee in this branch.</p>';
            return;
        }

        const employeeName = employeeCodeToNameMap[entries[0][HEADER_EMPLOYEE_CODE]] || entries[0][HEADER_EMPLOYEE_CODE];
        reportDisplay.innerHTML = `<h2>All Entries for ${employeeName}</h2>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        // Dynamically create headers from the first entry, excluding 'Timestamp' and 'Date' if desired
        const headersToShow = Object.keys(entries[0]).filter(header => 
            header !== HEADER_TIMESTAMP && header !== HEADER_DATE && header !== HEADER_EMPLOYEE_NAME
        );
        // Add Date and Employee Name at the beginning if desired
        const displayHeaders = ['Date', 'Employee Name', ...headersToShow];

        displayHeaders.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            displayHeaders.forEach(header => {
                const cell = row.insertCell();
                if (header === 'Date') {
                    cell.textContent = formatDate(entry[HEADER_DATE] || entry[HEADER_TIMESTAMP]);
                } else if (header === 'Employee Name') {
                     cell.textContent = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_CODE];
                }
                else {
                    cell.textContent = entry[header] || '';
                }
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Function to handle tab switching and section display
    function showTab(tabId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all report sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        todaysFollowUpsSection.style.display = 'none'; // NEW
        branchStatusTargetSection.style.display = 'none'; // NEW


        // Activate the clicked tab button and show the corresponding section
        document.getElementById(tabId).classList.add('active');
        if (tabId === 'allBranchSnapshotTabBtn' || tabId === 'allStaffOverallPerformanceTabBtn' || tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            // Clear selections and reports when switching tabs
            branchSelect.value = '';
            employeeSelect.value = '';
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
            reportDisplay.innerHTML = '<p>Select a branch and/or employee to view reports.</p>';
        } else if (tabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'block';
            customerViewBranchSelect.value = '';
            customerViewEmployeeSelect.value = '';
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
            customerDetailsContent.style.display = 'none';
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            // Populate branch datalist for add new employee form
            const branchNamesDatalist = document.getElementById('branchNames');
            branchNamesDatalist.innerHTML = '';
            PREDEFINED_BRANCHES.forEach(branch => {
                const option = document.createElement('option');
                option.value = branch;
                branchNamesDatalist.appendChild(option);
            });
        } else if (tabId === 'todaysFollowUpsTabBtn') { // NEW
            todaysFollowUpsSection.style.display = 'block';
            renderTodaysFollowUps();
        } else if (tabId === 'branchStatusTargetTabBtn') { // NEW
            branchStatusTargetSection.style.display = 'block';
            renderBranchStatusVsTarget();
        }
    }

    // NEW FUNCTION: Render Today's Follow-ups
    function renderTodaysFollowUps() {
        todaysFollowUpsDisplay.innerHTML = '<h2>Follow-ups Due Today</h2>';
        const today = new Date();
        const todayFormatted = formatDate(today); // Get today's date in YYYY-MM-DD

        const followUpsToday = allCanvassingData.filter(entry => {
            const followUpDate = entry[HEADER_NEXT_FOLLOW_UP_DATE];
            return followUpDate && formatDate(followUpDate) === todayFormatted;
        });

        if (followUpsToday.length > 0) {
            const tableContainer = document.createElement('div');
            tableContainer.className = 'data-table-container';
            const table = document.createElement('table');

            const thead = table.createTHead();
            const headerRow = thead.insertRow();
            const headers = [
                HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME,
                HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_REMARKS,
                HEADER_NEXT_FOLLOW_UP_DATE
            ];
            headers.forEach(headerText => {
                const th = document.createElement('th');
                th.textContent = headerText;
                headerRow.appendChild(th);
            });

            const tbody = table.createTBody();
            followUpsToday.forEach(entry => {
                const row = tbody.insertRow();
                headers.forEach(header => {
                    const cell = row.insertCell();
                    if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                        cell.textContent = formatDate(entry[header]) || '';
                    } else if (header === HEADER_EMPLOYEE_NAME) {
                        cell.textContent = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_NAME] || '';
                    }
                    else {
                        cell.textContent = entry[header] || '';
                    }
                });
            });
            tableContainer.appendChild(table);
            todaysFollowUpsDisplay.appendChild(tableContainer);
        } else {
            todaysFollowUpsDisplay.innerHTML += '<p class="no-participation-message">No follow-ups due today.</p>';
        }
    }

    // NEW FUNCTION: Render Branch Status vs. Target
    function renderBranchStatusVsTarget() {
        branchStatusTargetDisplay.innerHTML = '<h2>Branch Status vs. Target (Current Month)</h2>';

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const branchData = {}; // To store actuals and calculated targets for each branch

        PREDEFINED_BRANCHES.forEach(branch => {
            branchData[branch] = {
                actuals: { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 },
                targets: { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 },
                employees: new Set()
            };
        });

        // Calculate actuals and identify employees per branch for the current month
        allCanvassingData.forEach(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
                const branch = entry[HEADER_BRANCH_NAME];
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                if (branchData[branch]) {
                    const { totalActivity } = calculateTotalActivity([entry]); // Calculate activity for single entry
                    for (const metric in totalActivity) {
                        branchData[branch].actuals[metric] += totalActivity[metric];
                    }
                    if (employeeCode) {
                        branchData[branch].employees.add(employeeCode);
                    }
                }
            }
        });

        // Calculate total branch targets by summing individual employee targets for the current month
        for (const branch in branchData) {
            branchData[branch].employees.forEach(employeeCode => {
                const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
                const employeeTargets = TARGETS[employeeDesignation] || TARGETS['Default'];
                
                for (const metric in employeeTargets) {
                    branchData[branch].targets[metric] += employeeTargets[metric];
                }
            });
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'performance-table'; // Use existing performance table styles

        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Branch Name';
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3;
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        headerRow = thead.insertRow();
        headerRow.insertCell(); // Empty for Branch Name
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        let hasActivity = false;
        for (const branch of PREDEFINED_BRANCHES) {
            const data = branchData[branch];
            // Only show branches that have either actual activity or a calculated target
            if (Object.values(data.actuals).some(val => val > 0) || Object.values(data.targets).some(val => val > 0)) {
                hasActivity = true;
                const row = tbody.insertRow();
                row.insertCell().textContent = branch;

                metrics.forEach(metric => {
                    const actual = data.actuals[metric];
                    const target = data.targets[metric];
                    const percentage = target > 0 ? ((actual / target) * 100).toFixed(1) : (actual > 0 ? 100 : 0).toFixed(1);

                    row.insertCell().textContent = actual;
                    row.insertCell().textContent = target;

                    const percentageCell = row.insertCell();
                    const progressBarContainer = document.createElement('div');
                    progressBarContainer.className = 'progress-bar-container-small';
                    const progressBar = document.createElement('div');
                    progressBar.className = 'progress-bar';
                    progressBar.style.width = `${Math.min(100, percentage)}%`;
                    progressBar.textContent = `${percentage}%`;

                    if (percentage >= 100) {
                        progressBar.classList.add('success');
                    } else if (percentage >= 75) {
                        progressBar.classList.add('warning-high');
                    } else if (percentage >= 50) {
                        progressBar.classList.add('warning-medium');
                    } else if (percentage >= 25) {
                        progressBar.classList.add('warning-low');
                    } else {
                        progressBar.classList.add('danger');
                    }
                    
                    if (actual === 0 && target === 0) {
                        progressBar.classList.add('no-activity');
                        progressBar.textContent = 'N/A';
                        progressBar.style.width = '100%';
                    }

                    progressBarContainer.appendChild(progressBar);
                    percentageCell.appendChild(progressBarContainer);
                });
            }
        }

        tableContainer.appendChild(table);
        branchStatusTargetDisplay.appendChild(tableContainer);

        if (!hasActivity) {
            branchStatusTargetDisplay.innerHTML += '<p class="no-participation-message">No branch activity or targets found for the current month.</p>';
        }
    }


    // Function to send data to Google Apps Script Web App
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Sending data to server...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded', // Required by Apps Script
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                })
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Server error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message || 'Operation successful!', false);
                // Re-process data after successful update to reflect changes in reports
                await processData(); 
                // If a specific management form was used, show the management tab again
                showTab('employeeManagementTabBtn');
                return true;
            } else {
                displayEmployeeManagementMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Google Apps Script:', error);
            displayEmployeeManagementMessage(`Error: ${error.message}. Please check your web app URL and deployment.`, true);
            return false;
        }
    }


    // Event Listener for Tab Buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => {
        showTab('allBranchSnapshotTabBtn');
        renderAllBranchSnapshot();
    });
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => {
        showTab('allStaffOverallPerformanceTabBtn');
        renderOverallStaffPerformanceReport();
    });
    nonParticipatingBranchesTabBtn.addEventListener('click', () => {
        showTab('nonParticipatingBranchesTabBtn');
        renderNonParticipatingBranches();
    });
    detailedCustomerViewTabBtn.addEventListener('click', () => {
        showTab('detailedCustomerViewTabBtn');
        // Initial population of branch dropdown for detailed customer view
        populateDropdown(customerViewBranchSelect, allUniqueBranches);
        // Clear previous customer list/details when opening tab
        customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
        customerDetailsContent.style.display = 'none';
    });
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    // NEW TAB Event Listeners
    todaysFollowUpsTabBtn.addEventListener('click', () => showTab('todaysFollowUpsTabBtn'));
    branchStatusTargetTabBtn.addEventListener('click', () => showTab('branchStatusTargetTabBtn'));


    // Initial setup for Detailed Customer View tab
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        customerDetailsContent.style.display = 'none'; // Hide details when branch changes
        if (selectedBranch) {
            // Filter employees who have activity in the selected branch
            const employeesInBranch = [...new Set(allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                });
            populateDropdown(customerViewEmployeeSelect, employeesInBranch, true);
            customerCanvassedList.innerHTML = '<p>Select an employee to view canvassed customers.</p>';
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>'; // Clear employee dropdown
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        customerDetailsContent.style.display = 'none'; // Hide details when employee changes
        if (selectedBranch && selectedEmployeeCode) {
            // Filter entries for the selected branch and employee
            const employeeCustomerEntries = allCanvassingData.filter(entry =>
                entry[HEADER_BRANCH_NAME] === selectedBranch &&
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
            );
            renderCustomerCanvassedList(employeeCustomerEntries);
        } else {
            customerCanvassedList.innerHTML = '<p>Select an employee to view canvassed customers.</p>';
        }
    });

    function renderCustomerCanvassedList(entries) {
        if (!entries || entries.length === 0) {
            customerCanvassedList.innerHTML = '<p>No canvassed customers found for the selected employee.</p>';
            return;
        }

        customerCanvassedList.innerHTML = `<h3>Canvassed Customers</h3>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'customer-list-table'; // Add a specific class for styling

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_DATE, HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER, HEADER_PRODUCT_INTERESTED, HEADER_NEXT_FOLLOW_UP_DATE, HEADER_REMARKS
        ];
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            row.style.cursor = 'pointer'; // Make rows clickable
            row.addEventListener('click', () => displayCustomerDetails(entry));

            headers.forEach(header => {
                const cell = row.insertCell();
                if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    cell.textContent = formatDate(entry[header]) || '';
                } else {
                    cell.textContent = entry[header] || '';
                }
            });
        });
        tableContainer.appendChild(table);
        customerCanvassedList.appendChild(tableContainer);
    }

    function displayCustomerDetails(customerEntry) {
        customerDetailsContent.innerHTML = `
            <h3>Customer Details for ${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</h3>
            <div class="customer-details-grid">
                <div class="detail-row"><span class="detail-label">Date:</span> <span class="detail-value">${formatDate(customerEntry[HEADER_DATE]) || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Prospect Name:</span> <span class="detail-value">${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Phone/WhatsApp:</span> <span class="detail-value">${customerEntry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Address:</span> <span class="detail-value">${customerEntry[HEADER_ADDRESS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Profession:</span> <span class="detail-value">${customerEntry[HEADER_PROFESSION] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">DOB/WD:</span> <span class="detail-value">${customerEntry[HEADER_DOB_WD] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${customerEntry[HEADER_ACTIVITY_TYPE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Type of Customer:</span> <span class="detail-value">${customerEntry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Product Interested:</span> <span class="detail-value">${customerEntry[HEADER_PRODUCT_INTERESTED] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Next Follow-up Date:</span> <span class="detail-value">${formatDate(customerEntry[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${customerEntry[HEADER_REMARKS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Relation With Staff:</span> <span class="detail-value">${customerEntry[HEADER_RELATION_WITH_STAFF] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Lead Source:</span> <span class="detail-value">${customerEntry[HEADER_R_LEAD_SOURCE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Spouse Name):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_1] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Spouse Job):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_2] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Children Names):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_3] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Children Details):</span> <span class="detail-value">${customerEntry[HEADER_FAMILY_DETAILS_4] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Profile of Customer:</span> <span class="detail-value">${customerEntry[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</span></div>
            </div>
            <button class="btn" onclick="customerDetailsContent.style.display='none';">Hide Details</button>
        `;
        customerDetailsContent.style.display = 'block';
        customerDetailsContent.scrollIntoView({ behavior: 'smooth' }); // Scroll to view details
    }


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            if (employeeData[HEADER_EMPLOYEE_NAME] && employeeData[HEADER_EMPLOYEE_CODE] && employeeData[HEADER_BRANCH_NAME] && employeeData[HEADER_DESIGNATION]) {
                const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
                if (success) {
                    addEmployeeForm.reset();
                }
            } else {
                displayEmployeeManagementMessage('All fields are required to add an employee.', true);
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = employeeDetails.split('\n');
            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Name, Code, (Designation optional)
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
