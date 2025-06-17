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
            'Visit': 15,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Seniors': { // Added Investment Staff with custom Visit target
            'Visit': 15,
            'Call': 3 * MONTHLY_WORKING_DAYS,
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
    // Add these lines along with other DOM Elements
    const viewBranchVisitLeaderboardBtn = document.getElementById('viewBranchVisitLeaderboardBtn');
    const viewBranchCallLeaderboardBtn = document.getElementById('viewBranchCallLeaderboardBtn');
    const viewStaffParticipationBtn = document.getElementById('viewStaffParticipationBtn');

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
    // NEW: Total Staff Participation Tab Button
    const totalStaffParticipationTabBtn = document.getElementById('totalStaffParticipationTabBtn');


    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    // NEW: Total Staff Participation Section
    const totalStaffParticipationSection = document.getElementById('totalStaffParticipationSection');


    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent'); // Parent of the cards

    // Get the specific card elements for direct population
    const customerCard1 = document.getElementById('customerCard1');
    const customerCard2 = document.getElementById('customerCard2');
    const customerCard3 = document.getElementById('customerCard3');


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

   
    // NEW: Render Non-Participating Branches Report (now Zero Visit Branches)
    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Zero Visit Branches</h2>'; // Changed title
        const zeroVisitBranches = [];

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries); // Get total activities

            // Check if total visits for this branch is 0
            if (totalActivity['Visit'] === 0) {
                zeroVisitBranches.push(branch);
            }
        });

        if (zeroVisitBranches.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-branch-list'; // Reusing existing class
            zeroVisitBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded visits!</p>'; // Changed message
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
                const targetValue = targets[metric] || 0;
                // Ensure target is 0 if undefined
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
                            <tr><th>Metric</th><th>Actual</th><th>Target</th><th>%</th></tr>
                        </thead>
                        <tbody>
                            ${Object.keys(targets).map(metric => {
                                const actualValue = totalActivity[metric] || 0;
                                const targetValue = targets[metric];
                                let percentValue = performance[metric];
                                let displayPercent;
                                let progressBarClass;
                                let progressWidth;

                                if (isNaN(percentValue) || targetValue === 0) {
                                    displayPercent = 'N/A';
                                    progressWidth = 0;
                                    progressBarClass = 'no-activity';
                                } else {
                                    displayPercent = `${Math.round(percentValue)}%`;
                                    progressWidth = Math.min(100, Math.round(percentValue));
                                    progressBarClass = getProgressBarClass(percentValue);
                                }
                                if (actualValue === 0 && targetValue > 0) {
                                    displayPercent = '0%';
                                    progressWidth = 0;
                                    progressBarClass = 'danger';
                                }

                                return `
                                    <tr>
                                        <td>${metric}</td>
                                        <td>${actualValue}</td>
                                        <td>${targetValue}</td>
                                        <td>
                                            <div class="progress-bar-container-small">
                                                <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth === 0 && displayPercent !== 'N/A' ? '30px' : progressWidth}%">
                                                    ${displayPercent}
                                                </div>
                                            </div>
                                        </td>
                                    </tr>
                                `;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
            `;
            branchPerformanceGrid.appendChild(employeeCard);
        });
        reportDisplay.appendChild(branchPerformanceGrid);
    }

    // Function to render employee summary (d4.PNG style)
    function renderEmployeeSummary(entries) {
        const reportArea = document.getElementById('reportDisplay');
        reportArea.innerHTML = '<h2>Employee Activity Summary (This Month)</h2>';

        if (entries.length === 0) {
            reportArea.innerHTML += '<p>No activity records for the selected employee in the current month.</p>';
            return;
        }

        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        const currentMonthEntries = entries.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            const currentMonth = new Date().getMonth();
            const currentYear = new Date().getFullYear();
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        if (currentMonthEntries.length === 0) {
            reportArea.innerHTML += `<p>No activity records for ${employeeName} for the current month.</p>`;
            return;
        }

        const { totalActivity, productInterests } = calculateTotalActivity(currentMonthEntries);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        const cardContainer = document.createElement('div');
        cardContainer.className = 'employee-summary-cards';

        // Card 1: Overall Performance Summary (Visits, Calls, References, New Customer Leads)
        let performanceTableHtml = `
            <div class="card">
                <h3>Overall Performance - ${employeeName}</h3>
                <p>Designation: ${designation}</p>
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th>Metric</th>
                            <th>Actual</th>
                            <th>Target</th>
                            <th>%</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        const metricsOrder = ['Visit', 'Call', 'Reference', 'New Customer Leads']; // Ensure consistent order
        metricsOrder.forEach(metric => {
            const actual = totalActivity[metric] || 0;
            const target = targets[metric] || 0;
            let percentValue = performance[metric];
            let displayPercent;
            let progressBarClass;
            let progressWidth;

            if (isNaN(percentValue) || target === 0) {
                displayPercent = 'N/A';
                progressWidth = 0;
                progressBarClass = 'no-activity';
            } else {
                displayPercent = `${Math.round(percentValue)}%`;
                progressWidth = Math.min(100, Math.round(percentValue));
                progressBarClass = getProgressBarClass(percentValue);
            }
            if (actual === 0 && target > 0) {
                displayPercent = '0%';
                progressWidth = 0;
                progressBarClass = 'danger';
            }

            performanceTableHtml += `
                <tr>
                    <td>${metric}</td>
                    <td>${actual}</td>
                    <td>${target}</td>
                    <td>
                        <div class="progress-bar-container-small">
                            <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth === 0 && displayPercent !== 'N/A' ? '30px' : progressWidth}%">
                                ${displayPercent}
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        });
        performanceTableHtml += `
                    </tbody>
                </table>
            </div>
        `;
        cardContainer.innerHTML += performanceTableHtml;

        // Card 2: Recent Activities (Last 5)
        const recentActivities = currentMonthEntries
            .sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP])) // Sort by most recent
            .slice(0, 5); // Get last 5

        let recentActivitiesHtml = `
            <div class="card">
                <h3>Recent Activities</h3>
                ${recentActivities.length > 0 ? `
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Activity</th>
                            <th>Prospect</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${recentActivities.map(entry => `
                            <tr>
                                <td>${formatDate(entry[HEADER_TIMESTAMP])}</td>
                                <td>${entry[HEADER_ACTIVITY_TYPE]}</td>
                                <td>${entry[HEADER_PROSPECT_NAME] || 'N/A'}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
                ` : '<p>No recent activities.</p>'}
            </div>
        `;
        cardContainer.innerHTML += recentActivitiesHtml;

        // Card 3: Product Interests (unique list)
        let productInterestsHtml = `
            <div class="card">
                <h3>Product Interests</h3>
                ${productInterests.length > 0 ? `
                    <ul>
                        ${productInterests.map(product => `<li>${product}</li>`).join('')}
                    </ul>
                ` : '<p>No specific product interests recorded.</p>'}
            </div>
        `;
        cardContainer.innerHTML += productInterestsHtml;

        reportArea.appendChild(cardContainer);
    }

    // Function to render all entries (Raw Data view, d5.PNG style)
    function renderAllEntries(entries) {
        const reportArea = document.getElementById('reportDisplay');
        reportArea.innerHTML = '<h2>All Canvassing Entries</h2>';
        
        if (entries.length === 0) {
            reportArea.innerHTML += '<p>No entries to display.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'full-data-table';
        
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        
        // Define the order of headers as requested/useful for display
        const displayHeaders = [
            HEADER_TIMESTAMP, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_EMPLOYEE_CODE,
            HEADER_DESIGNATION, HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER,
            HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ADDRESS,
            HEADER_PROFESSION, HEADER_PRODUCT_INTERESTED, HEADER_REMARKS, HEADER_NEXT_FOLLOW_UP_DATE,
            HEADER_RELATION_WITH_STAFF, HEADER_FAMILY_DETAILS_1, HEADER_FAMILY_DETAILS_2,
            HEADER_FAMILY_DETAILS_3, HEADER_FAMILY_DETAILS_4, HEADER_PROFILE_OF_CUSTOMER
        ];

        displayHeaders.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText.replace('HEADER_', '').replace(/_/g, ' '); // Basic formatting for display
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            displayHeaders.forEach(headerKey => {
                const cell = row.insertCell();
                let cellValue = entry[headerKey] || '';
                // Format date columns for readability
                if (headerKey === HEADER_TIMESTAMP || headerKey === HEADER_NEXT_FOLLOW_UP_DATE) {
                    cellValue = formatDate(cellValue);
                }
                cell.textContent = cellValue;
            });
        });

        tableContainer.appendChild(table);
        reportArea.appendChild(tableContainer);
    }

    // --- Tab Switching Logic ---
    function showTab(tabId) {
        // Hide all sections
        reportsSection.classList.remove('active');
        detailedCustomerViewSection.classList.remove('active');
        employeeManagementSection.classList.remove('active');
        // NEW: Hide Total Staff Participation Section
        totalStaffParticipationSection.classList.remove('active');


        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the selected section and activate its button
        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                reportsSection.classList.add('active');
                allBranchSnapshotTabBtn.classList.add('active');
                renderAllBranchSnapshot(); // Render default report for this tab
                // Reset main report area filters
                branchSelect.value = "";
                employeeSelect.value = "";
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                break;
            case 'allStaffOverallPerformanceTabBtn':
                reportsSection.classList.add('active');
                allStaffOverallPerformanceTabBtn.classList.add('active');
                renderOverallStaffPerformanceReport(); // Render overall staff performance
                // Reset main report area filters
                branchSelect.value = "";
                employeeSelect.value = "";
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                break;
            case 'nonParticipatingBranchesTabBtn':
                reportsSection.classList.add('active');
                nonParticipatingBranchesTabBtn.classList.add('active');
                renderNonParticipatingBranches(); // Render non-participating branches
                // Reset main report area filters
                branchSelect.value = "";
                employeeSelect.value = "";
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                break;
            case 'branchPerformanceTabBtn':
                reportsSection.classList.add('active');
                branchPerformanceTabBtn.classList.add('active');
                // For Branch Performance, show branch select and clear other filters
                branchSelect.value = "";
                employeeSelect.value = "";
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                reportDisplay.innerHTML = '<p>Please select a branch to view its performance reports.</p>';
                break;
            case 'detailedCustomerViewTabBtn': // NEW case for Detailed Customer View
                detailedCustomerViewSection.classList.add('active');
                detailedCustomerViewTabBtn.classList.add('active');
                // Reset customer view filters when tab is opened
                customerViewBranchSelect.value = "";
                customerViewEmployeeSelect.value = "";
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
                customerDetailsContent.innerHTML = ''; // Clear customer card details
                populateDropdown(customerViewEmployeeSelect, [], true); // Clear employee dropdown initially
                break;
            case 'employeeManagementTabBtn': // NEW case for Employee Management
                employeeManagementSection.classList.add('active');
                employeeManagementTabBtn.classList.add('active');
                // Optionally clear forms or display messages when tab is opened
                addEmployeeForm.reset();
                bulkAddEmployeeForm.reset();
                deleteEmployeeForm.reset();
                employeeManagementMessage.textContent = '';
                employeeManagementMessage.style.display = 'none';
                break;
            // NEW case for Total Staff Participation Tab
            case 'totalStaffParticipationTabBtn':
                totalStaffParticipationSection.classList.add('active');
                totalStaffParticipationTabBtn.classList.add('active');
                renderTotalStaffParticipationReport(); // Call the new report function
                break;
            default:
                reportsSection.classList.add('active');
                allBranchSnapshotTabBtn.classList.add('active');
                renderAllBranchSnapshot();
                break;
        }
    }

    // --- Specific Report Button Event Listeners (Main Report Area) ---
    if (viewBranchPerformanceReportBtn) {
        viewBranchPerformanceReportBtn.addEventListener('click', () => {
            const selectedBranch = branchSelect.value;
            if (selectedBranch) {
                renderBranchPerformanceReport(selectedBranch);
            } else {
                displayMessage('Please select a branch first.', 'error');
            }
        });
    }

    if (viewEmployeeSummaryBtn) {
        viewEmployeeSummaryBtn.addEventListener('click', () => {
            if (selectedEmployeeCodeEntries.length > 0) {
                renderEmployeeSummary(selectedEmployeeCodeEntries);
            } else {
                displayMessage('Please select an employee first.', 'error');
            }
        });
    }

    if (viewAllEntriesBtn) {
        viewAllEntriesBtn.addEventListener('click', () => {
            const selectedBranch = branchSelect.value;
            const selectedEmployeeCode = employeeSelect.value;
            let entriesToDisplay = [];

            if (selectedBranch && selectedEmployeeCode) {
                // If both branch and employee are selected, show only employee's entries for that branch
                entriesToDisplay = allCanvassingData.filter(entry =>
                    entry[HEADER_BRANCH_NAME] === selectedBranch &&
                    entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
                );
            } else if (selectedBranch) {
                // If only branch is selected, show all entries for that branch
                entriesToDisplay = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            } else {
                // If neither is selected, show all data (this is less likely to be used, but as a fallback)
                entriesToDisplay = allCanvassingData;
            }
            renderAllEntries(entriesToDisplay);
        });
    }
    
    // NEW: Detailed Customer View Logic
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        customerCanvassedList.innerHTML = '<p>Select an employee to view customers.</p>'; // Clear
        customerDetailsContent.innerHTML = ''; // Clear customer card details

        if (selectedBranch) {
            const employeesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);
            
            const combinedEmployeeCodes = new Set(employeesInBranchFromCanvassing);

            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            
            populateDropdown(customerViewEmployeeSelect, sortedEmployeeCodesInBranch, true);
        } else {
            populateDropdown(customerViewEmployeeSelect, [], true); // Clear employee dropdown
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        customerDetailsContent.innerHTML = ''; // Clear customer card details

        if (selectedBranch && selectedEmployeeCode) {
            const employeeCustomers = allCanvassingData.filter(entry =>
                entry[HEADER_BRANCH_NAME] === selectedBranch &&
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
            ).sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP])); // Sort by recent for display

            renderCustomerList(employeeCustomers);
        } else {
            customerCanvassedList.innerHTML = '<p>Select both a branch and an employee.</p>';
        }
    });

    function renderCustomerList(customers) {
        customerCanvassedList.innerHTML = ''; // Clear previous list

        if (customers.length === 0) {
            customerCanvassedList.innerHTML = '<p>No canvassed customers found for this employee.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list'; // Add a class for styling

        customers.forEach(customer => {
            const li = document.createElement('li');
            li.textContent = `${customer[HEADER_PROSPECT_NAME] || 'N/A'} - ${formatDate(customer[HEADER_TIMESTAMP])}`;
            li.dataset.phoneNumber = customer[HEADER_PHONE_NUMBER_WHATSAPP]; // Store phone for click
            li.dataset.address = customer[HEADER_ADDRESS]; // Store address for click
            li.dataset.profession = customer[HEADER_PROFESSION]; // Store profession for click
            li.dataset.productInterested = customer[HEADER_PRODUCT_INTERESTED]; // Store product interested
            li.dataset.remarks = customer[HEADER_REMARKS]; // Store remarks
            li.dataset.nextFollowUpDate = customer[HEADER_NEXT_FOLLOW_UP_DATE]; // Store next follow-up date
            li.dataset.relationWithStaff = customer[HEADER_RELATION_WITH_STAFF]; // Store relation with staff

            // Store family details and profile
            li.dataset.familyDetails1 = customer[HEADER_FAMILY_DETAILS_1] || 'N/A';
            li.dataset.familyDetails2 = customer[HEADER_FAMILY_DETAILS_2] || 'N/A';
            li.dataset.familyDetails3 = customer[HEADER_FAMILY_DETAILS_3] || 'N/A';
            li.dataset.familyDetails4 = customer[HEADER_FAMILY_DETAILS_4] || 'N/A';
            li.dataset.profileOfCustomer = customer[HEADER_PROFILE_OF_CUSTOMER] || 'N/A';


            li.addEventListener('click', () => {
                displayCustomerDetails(li.dataset); // Pass all dataset values
            });
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);

        // Automatically display details of the first customer
        if (customers.length > 0) {
            displayCustomerDetails(ul.firstChild.dataset);
        }
    }

    function displayCustomerDetails(customerData) {
        customerDetailsContent.innerHTML = ''; // Clear previous details

        // Card 1: Basic Contact & Activity
        customerCard1.innerHTML = `
            <h3>${customerData.prospectName || 'Customer Details'}</h3>
            <p><strong>Phone:</strong> ${customerData.phoneNumber || 'N/A'}</p>
            <p><strong>Address:</strong> ${customerData.address || 'N/A'}</p>
            <p><strong>Profession:</strong> ${customerData.profession || 'N/A'}</p>
            <p><strong>Product Interested:</strong> ${customerData.productInterested || 'N/A'}</p>
            <p><strong>Remarks:</strong> ${customerData.remarks || 'N/A'}</p>
            <p><strong>Next Follow-up Date:</strong> ${formatDate(customerData.nextFollowUpDate) || 'N/A'}</p>
            <p><strong>Relation With Staff:</strong> ${customerData.relationWithStaff || 'N/A'}</p>
        `;
        // Card 2: Family Details
        customerCard2.innerHTML = `
            <h3>Family Details</h3>
            <p><strong>Spouse Name:</strong> ${customerData.familyDetails1}</p>
            <p><strong>Spouse Job:</strong> ${customerData.familyDetails2}</p>
            <p><strong>Children Names:</strong> ${customerData.familyDetails3}</p>
            <p><strong>Children Details:</strong> ${customerData.familyDetails4}</p>
        `;
        // Card 3: Profile of Customer
        customerCard3.innerHTML = `
            <h3>Profile of Customer</h3>
            <p>${customerData.profileOfCustomer}</p>
        `;

        // Append cards to the display area
        customerDetailsContent.appendChild(customerCard1);
        customerDetailsContent.appendChild(customerCard2);
        customerDetailsContent.appendChild(customerCard3);
    }

    // --- Google Apps Script Integration (for Employee Management) ---
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Sending data...', false); // Info message
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
                console.error(`HTTP error! Status: ${response.status}, Response: ${errorText}`);
                displayEmployeeManagementMessage(`Error: ${response.status} - ${errorText}`, true);
                return false;
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                // After successful operation, re-process data to update front-end
                await processData(); 
                return true;
            } else {
                displayEmployeeManagementMessage(result.message || 'An unknown error occurred.', true);
                return false;
            }
        } catch (error) {
            console.error('Network or parsing error:', error);
            displayEmployeeManagementMessage(`Network error: ${error.message}`, true);
            return false;
        }
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const newEmployeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            // Basic validation
            if (!newEmployeeData[HEADER_EMPLOYEE_NAME] || !newEmployeeData[HEADER_EMPLOYEE_CODE] || !newEmployeeData[HEADER_BRANCH_NAME] || !newEmployeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', newEmployeeData);
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
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = employeeDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Expect at least Name, Code
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


    // NEW: renderTotalStaffParticipationReport Function
    async function renderTotalStaffParticipationReport() {
        const reportDisplay = document.getElementById('totalStaffParticipationReportDisplay');
        if (!reportDisplay) {
            console.error("Total Staff Participation Report Display element not found.");
            return;
        }

        reportDisplay.innerHTML = '<h3>Loading Total Staff Participation Data...</h3>';

        // Ensure data is loaded; processData calls fetchCanvassingData
        if (allCanvassingData.length === 0 || Object.keys(employeeCodeToDesignationMap).length === 0) {
            await processData(); // This will fetch data and populate maps
            if (allCanvassingData.length === 0) {
                reportDisplay.innerHTML = '<p>No data available to generate the Total Staff Participation report.</p>';
                return;
            }
        }

        // Initialize participation statistics for each activity type
        const participationStats = {
            'Visit': { totalStaff: new Set(), fullyCompletedStaff: new Set(), partiallyCompletedStaff: new Set() },
            'Calls': { totalStaff: new Set(), fullyCompletedStaff: new Set(), partiallyCompletedStaff: new Set() },
            'Referance': { totalStaff: new Set(), fullyCompletedStaff: new Set(), partiallyCompletedStaff: new Set() },
            'New Customer Leads': { totalStaff: new Set(), fullyCompletedStaff: new Set(), partiallyCompletedStaff: new Set() }
        };

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Group activities by employee code for the current month
        const employeeActivitiesThisMonth = {};
        allCanvassingData.forEach(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                if (employeeCode) {
                    if (!employeeActivitiesThisMonth[employeeCode]) {
                        employeeActivitiesThisMonth[employeeCode] = [];
                    }
                    employeeActivitiesThisMonth[employeeCode].push(entry);
                }
            }
        });

        // Iterate through each employee who had activity this month
        for (const employeeCode in employeeActivitiesThisMonth) {
            const employeeEntries = employeeActivitiesThisMonth[employeeCode];
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            const employeeTargets = TARGETS[designation] || TARGETS['Default'];

            // Use the existing calculateTotalActivity to get actual counts for this specific employee
            const { totalActivity: employeeActuals } = calculateTotalActivity(employeeEntries);

            // Evaluate each activity type for this employee
            // Using an array to maintain order and map keys correctly
            [
                { key: 'Visit', targetKey: 'Visit' },
                { key: 'Calls', targetKey: 'Call' }, // Map 'Calls' to 'Call' for TARGETS lookup
                { key: 'Referance', targetKey: 'Reference' },
                { key: 'New Customer Leads', targetKey: 'New Customer Leads' }
            ].forEach(({ key: activityReportKey, targetKey: activityTargetKey }) => {

                const actualCount = employeeActuals[activityTargetKey] || 0;
                const targetCount = employeeTargets[activityTargetKey] || 0;

                // Determine participation status
                if (actualCount > 0) {
                    participationStats[activityReportKey].totalStaff.add(employeeCode); // Staff participated if actual > 0
                }

                // Determine fully completed or partially completed based on targets
                if (targetCount > 0) { // If there's a specific target defined
                    if (actualCount >= targetCount) {
                        participationStats[activityReportKey].fullyCompletedStaff.add(employeeCode);
                    } else if (actualCount > 0 && actualCount < targetCount) {
                        participationStats[activityReportKey].partiallyCompletedStaff.add(employeeCode);
                    }
                } else if (actualCount > 0) {
                    // If no specific target (targetCount is 0 or undefined), but employee performed activity,
                    // consider them as 'fully completed' for this report's purpose if they did any activity.
                    participationStats[activityReportKey].fullyCompletedStaff.add(employeeCode);
                }
            });
        }

        // Generate HTML for the report table
        let tableHtml = `
            <h3>Overall Staff Participation (Current Month)</h3>
            <table class="report-table">
                <thead>
                    <tr>
                        <th>Activity Type</th>
                        <th>Total Staff Participated</th>
                        <th>Fully Completed</th>
                        <th>Partially Completed (Target)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        // Order of activities as requested by user
        const orderedActivityTypes = ['Visit', 'Calls', 'Referance', 'New Customer Leads'];

        orderedActivityTypes.forEach(type => {
            const stats = participationStats[type];
            const totalParticipated = stats.totalStaff.size;
            const fullyCompletedCount = stats.fullyCompletedStaff.size;
            
            // Partially completed staff are those who participated (totalStaff) but did NOT fully complete (fullyCompletedStaff)
            const partiallyCompletedTargetCount = totalParticipated - fullyCompletedCount;


            tableHtml += `
                <tr>
                    <td>${type}</td>
                    <td>${totalParticipated}</td>
                    <td>${fullyCompletedCount}</td>
                    <td>${partiallyCompletedTargetCount}</td>
                </tr>
            `;
        });

        tableHtml += `
                </tbody>
            </table>
        `;

        reportDisplay.innerHTML = tableHtml;
    }


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
