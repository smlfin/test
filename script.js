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
    // REMOVED: const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
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
    const dailyProgressMonitoringTabBtn = document.getElementById('dailyProgressMonitoringTabBtn'); // NEW

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const dailyProgressMonitoringSection = document.getElementById('dailyProgressMonitoringSection'); // NEW

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent'); // Parent of the cards
    const customerActivityDate = document.getElementById('customerActivityDate');


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

    // NEW: Daily Progress Monitoring Elements
    const progressDateInput = document.getElementById('progressDate');
    const progressBranchSelect = document.getElementById('progressBranchSelect');
    const progressEmployeeSelect = document.getElementById('progressEmployeeSelect');
    const showDailyProgressBtn = document.getElementById('showDailyProgressBtn');
    const dailyProgressDisplay = document.getElementById('dailyProgressDisplay');


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
        populateDropdown(progressBranchSelect, ['-- All Branches --', ...allUniqueBranches]); // Populate for daily progress, with all branches option
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot(); // Render the default "All Branch Snapshot" report
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items, useCodeForValue = false) {
        selectElement.innerHTML = ''; // Clear existing options
        if (selectElement.id === 'progressBranchSelect' || selectElement.id === 'progressEmployeeSelect') {
             // For daily progress, first option should be "-- All Branches --" or "-- All Employees --"
            const allOption = document.createElement('option');
            allOption.value = '';
            allOption.textContent = selectElement.id === 'progressBranchSelect' ? '-- All Branches --' : '-- All Employees --';
            selectElement.appendChild(allOption);
        } else {
            // For other dropdowns, use "-- Select --" as the default
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = '-- Select --';
            selectElement.appendChild(defaultOption);
        }

        items.forEach(item => {
            // Skip the '-- All Branches --' or '-- All Employees --' if it's already added or if it's the value passed for other dropdowns
            if (item === '-- All Branches --' || item === '-- All Employees --') {
                if (selectElement.id === 'progressBranchSelect' || selectElement.id === 'progressEmployeeSelect') {
                    // Already added as the first option, so skip
                    return;
                }
            }
            const option = document.createElement('option');
            if (useCodeForValue) {
                // item is employeeCode
                option.value = item; 
                option.textContent = employeeCodeToNameMap[item] || item; // Display name from map or code itself
            } else {
                // item is branch name
                option.value = item; 
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
        const employeesWithActivityThisMonth = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_DATE]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        }).map(entry => entry[HEADER_EMPLOYEE_CODE]);

        const uniqueEmployeesThisMonth = [...new Set(employeesWithActivityThisMonth)];

        if (uniqueEmployeesThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No staff activity recorded this month.</p>';
            return;
        }

        uniqueEmployeesThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivitiesThisMonth = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_DATE]).getMonth() === currentMonth &&
                new Date(entry[HEADER_DATE]).getFullYear() === currentYear
            );

            const { totalActivity } = calculateTotalActivity(employeeActivitiesThisMonth);
            const employeeTargets = TARGETS[designation] || TARGETS['Default'];

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actual = totalActivity[metric] || 0;
                const target = employeeTargets[metric] || 0;
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) + '%' : 'N/A';

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                row.insertCell().textContent = percentage;
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Summary Report (d4.PNG)
    function renderEmployeeSummary(entries) {
        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML = '<p class="no-participation-message">No activity found for this employee in the selected branch.</p>';
            return;
        }

        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const branchName = entries[0][HEADER_BRANCH_NAME];
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const employeeTargets = TARGETS[designation] || TARGETS['Default'];

        reportDisplay.innerHTML = `
            <h2>Employee Summary: ${employeeName} (${branchName})</h2>
            <div class="employee-summary-cards">
                <div class="summary-card">
                    <h3>Designation</h3>
                    <p>${designation}</p>
                </div>
                <div class="summary-card">
                    <h3>Total Entries</h3>
                    <p>${entries.length}</p>
                </div>
                </div>
        `;

        const { totalActivity, productInterests } = calculateTotalActivity(entries);

        // Create a table for activity details
        const activityTable = document.createElement('table');
        activityTable.className = 'employee-performance-card'; // Reusing the performance card style
        const thead = activityTable.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Activity Type';
        headerRow.insertCell().textContent = 'Actual';
        headerRow.insertCell().textContent = 'Target';
        headerRow.insertCell().textContent = '% Achieved';

        const tbody = activityTable.createTBody();

        // Populate activity rows
        for (const [activity, actual] of Object.entries(totalActivity)) {
            const target = employeeTargets[activity] || 0;
            const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) + '%' : 'N/A';
            const row = tbody.insertRow();
            row.insertCell().textContent = activity;
            row.insertCell().textContent = actual;
            row.insertCell().textContent = target;
            row.insertCell().textContent = percentage;
        }

        reportDisplay.appendChild(activityTable);

        // Display Product Interests
        if (productInterests.length > 0) {
            const productInterestDiv = document.createElement('div');
            productInterestDiv.className = 'product-interest-summary';
            productInterestDiv.innerHTML = `<h3>Products of Interest:</h3><p>${productInterests.join(', ')}</p>`;
            reportDisplay.appendChild(productInterestDiv);
        }
    }


    // Render All Entries for Filter (d5.PNG)
    function renderAllEntries(entries) {
        reportDisplay.innerHTML = '<h2>All Entries for Filter</h2>';

        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No entries found for the current filters.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'all-entries-table'; // You might want to create a specific style for this

        const thead = table.createTHead();
        const headerRow = thead.insertRow();

        // Dynamically get all unique headers from the filtered entries
        const allHeaders = new Set();
        entries.forEach(entry => {
            Object.keys(entry).forEach(header => allHeaders.add(header));
        });
        const sortedHeaders = Array.from(allHeaders).sort(); // Sort headers alphabetically

        sortedHeaders.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            sortedHeaders.forEach(header => {
                const cell = row.insertCell();
                let cellValue = entry[header] || '';
                // Format date columns if they match date headers
                if ((header === HEADER_DATE || header === HEADER_TIMESTAMP || header === HEADER_NEXT_FOLLOW_UP_DATE) && cellValue) {
                    cellValue = formatDate(cellValue);
                }
                cell.textContent = cellValue;
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // --- Tab Switching Logic ---
    function showTab(tabButtonId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => button.classList.remove('active'));
        // Hide all content sections
        document.querySelectorAll('.report-section').forEach(section => section.style.display = 'none');

        // Activate the clicked tab button
        const clickedButton = document.getElementById(tabButtonId);
        if (clickedButton) {
            clickedButton.classList.add('active');
        }

        // Show the corresponding section and render report
        if (tabButtonId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            renderAllBranchSnapshot();
            // Reset main report filters
            branchSelect.value = "";
            employeeSelect.value = "";
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
        } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            renderOverallStaffPerformanceReport();
            // Reset main report filters
            branchSelect.value = "";
            employeeSelect.value = "";
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
        } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            renderNonParticipatingBranches();
            // Reset main report filters
            branchSelect.value = "";
            employeeSelect.value = "";
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
        } else if (tabButtonId === 'detailedCustomerViewTabBtn') { // NEW
            detailedCustomerViewSection.style.display = 'block';
            // Populate dropdowns specifically for detailed customer view
            populateDropdown(customerViewBranchSelect, allUniqueBranches);
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select Employee --</option>'; // Clear employee dropdown initially
            customerCanvassedList.innerHTML = '<p>Select criteria and click \'Search\' to view canvassed customers.</p>';
            customerDetailsContent.innerHTML = '<p>Click on a customer from the list to view their details.</p>';
            customerActivityDate.value = ''; // Clear date input
        } else if (tabButtonId === 'employeeManagementTabBtn') { // NEW
            employeeManagementSection.style.display = 'block';
            // Clear any previous messages
            displayEmployeeManagementMessage('', false);
        } else if (tabButtonId === 'dailyProgressMonitoringTabBtn') { // NEW
            dailyProgressMonitoringSection.style.display = 'block';
            populateDropdown(progressBranchSelect, allUniqueBranches);
            progressEmployeeSelect.innerHTML = '<option value="">-- All Employees --</option>';
            dailyProgressDisplay.innerHTML = '<p>Select a date, branch, and/or employee and click "Show Progress".</p>';
            // Set default date to today
            progressDateInput.value = formatDate(new Date());
        }
    }

    // --- Main Report Section Button Event Listeners ---
    if (viewBranchVisitLeaderboardBtn) {
        viewBranchVisitLeaderboardBtn.addEventListener('click', () => {
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewBranchVisitLeaderboardBtn.classList.add('active');
            renderBranchLeaderboard('Visit');
        });
    }

    if (viewBranchCallLeaderboardBtn) {
        viewBranchCallLeaderboardBtn.addEventListener('click', () => {
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewBranchCallLeaderboardBtn.classList.add('active');
            renderBranchLeaderboard('Call');
        });
    }

    if (viewStaffParticipationBtn) {
        viewStaffParticipationBtn.addEventListener('click', () => {
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewStaffParticipationBtn.classList.add('active');
            renderStaffParticipation();
        });
    }

    // Event Listener for Employee Summary Button (d4.PNG)
    if (viewEmployeeSummaryBtn) {
        viewEmployeeSummaryBtn.addEventListener('click', () => {
            const selectedEmployeeCode = employeeSelect.value;
            if (selectedEmployeeCode) {
                document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
                viewEmployeeSummaryBtn.classList.add('active');
                renderEmployeeSummary(selectedEmployeeCodeEntries);
            } else {
                displayMessage("Please select an employee first to view summary.", 'error');
            }
        });
    }

    // Event Listener for All Entries Button (d5.PNG)
    if (viewAllEntriesBtn) {
        viewAllEntriesBtn.addEventListener('click', () => {
            const selectedBranch = branchSelect.value;
            const selectedEmployeeCode = employeeSelect.value;
            let entriesToDisplay = [];

            if (selectedBranch && selectedEmployeeCode) {
                // Both branch and employee selected
                entriesToDisplay = allCanvassingData.filter(entry =>
                    entry[HEADER_BRANCH_NAME] === selectedBranch &&
                    entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
                );
            } else if (selectedBranch) {
                // Only branch selected
                entriesToDisplay = allCanvassingData.filter(entry =>
                    entry[HEADER_BRANCH_NAME] === selectedBranch
                );
            } else {
                displayMessage("Please select a branch or an employee to view all entries.", 'error');
                return;
            }

            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewAllEntriesBtn.classList.add('active');
            renderAllEntries(entriesToDisplay);
        });
    }

    // Render Branch Leaderboard (Visits or Calls)
    function renderBranchLeaderboard(activityType) {
        const selectedBranch = branchSelect.value;
        if (!selectedBranch) {
            reportDisplay.innerHTML = '<p>Please select a branch to view the leaderboard.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Branch ${activityType} Leaderboard for ${selectedBranch}</h2>`;

        const branchEmployeesActivity = {}; // {employeeCode: count}

        allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
            .forEach(entry => {
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                const actType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';

                if (!branchEmployeesActivity[employeeCode]) {
                    branchEmployeesActivity[employeeCode] = { count: 0, name: employeeCodeToNameMap[employeeCode] || employeeCode };
                }

                if ((activityType === 'Visit' && actType === 'visit') ||
                    (activityType === 'Call' && actType === 'calls')) {
                    branchEmployeesActivity[employeeCode].count++;
                }
            });

        const sortedLeaderboard = Object.values(branchEmployeesActivity).sort((a, b) => b.count - a.count);

        if (sortedLeaderboard.length === 0) {
            reportDisplay.innerHTML += `<p class="no-participation-message">No ${activityType.toLowerCase()}s recorded for this branch.</p>`;
            return;
        }

        const table = document.createElement('table');
        table.className = 'leaderboard-table'; // You might want to create a specific style for this
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Rank';
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = activityType + ' Count';

        const tbody = table.createTBody();
        sortedLeaderboard.forEach((employee, index) => {
            const row = tbody.insertRow();
            row.insertCell().textContent = index + 1;
            row.insertCell().textContent = employee.name;
            row.insertCell().textContent = employee.count;
        });

        reportDisplay.appendChild(table);
    }


    // Render Staff Participation (d3.PNG)
    function renderStaffParticipation() {
        const selectedBranch = branchSelect.value;
        if (!selectedBranch) {
            reportDisplay.innerHTML = '<p>Please select a branch to view staff participation.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Staff Participation for ${selectedBranch}</h2>`;

        // Get all unique employee codes associated with the selected branch from canvassing data
        const employeesInSelectedBranch = [...new Set(
            allCanvassingData
            .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
            .map(entry => entry[HEADER_EMPLOYEE_CODE])
        )].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });


        if (employeesInSelectedBranch.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No employees found with activity in this branch.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'participation-table'; // You might want to create a specific style for this
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Total Entries';
        headerRow.insertCell().textContent = 'Last Activity Date';

        const tbody = table.createTBody();

        employeesInSelectedBranch.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const employeeEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode && entry[HEADER_BRANCH_NAME] === selectedBranch);
            const totalEntries = employeeEntries.length;

            let lastActivityDate = 'N/A';
            if (totalEntries > 0) {
                // Find the latest date among the employee's entries
                const latestEntry = employeeEntries.reduce((latest, current) => {
                    const latestDate = new Date(latest[HEADER_DATE]);
                    const currentDate = new Date(current[HEADER_DATE]);
                    return currentDate > latestDate ? current : latest;
                });
                lastActivityDate = formatDate(latestEntry[HEADER_DATE]);
            }

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = totalEntries;
            row.insertCell().textContent = lastActivityDate;
        });

        reportDisplay.appendChild(table);
    }


    // --- Detailed Customer View Logic ---
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        if (selectedBranch) {
            // Filter employees for the selected branch based on canvassing data
            const employeeCodesInBranch = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);
            const uniqueEmployeeCodes = [...new Set(employeeCodesInBranch)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, uniqueEmployeeCodes, true);
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select Employee --</option>';
        }
        customerCanvassedList.innerHTML = '<p>Select criteria and click \'Search\' to view canvassed customers.</p>';
        customerDetailsContent.innerHTML = '<p>Click on a customer from the list to view their details.</p>';
    });

    document.getElementById('customerCanvassedSearchBtn').addEventListener('click', () => {
        const branch = customerViewBranchSelect.value;
        const employeeCode = customerViewEmployeeSelect.value;
        const date = customerActivityDate.value; // YYYY-MM-DD format from input

        if (!branch && !employeeCode && !date) {
            displayMessage("Please select at least one filter (Branch, Employee, or Date) for Customer View.", 'error');
            return;
        }

        let filteredCustomers = allCanvassingData;

        if (branch) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
        }
        if (employeeCode) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
        }
        if (date) {
            // Convert entry date to YYYY-MM-DD for comparison
            filteredCustomers = filteredCustomers.filter(entry => formatDate(entry[HEADER_DATE]) === date);
        }

        displayCanvassedCustomers(filteredCustomers);
    });

    function displayCanvassedCustomers(customers) {
        customerCanvassedList.innerHTML = '';
        customerDetailsContent.innerHTML = '<p>Click on a customer from the list to view their details.</p>';

        if (customers.length === 0) {
            customerCanvassedList.innerHTML = '<p class="no-participation-message">No canvassed customers found for the selected criteria.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'canvassed-customer-list';
        customers.forEach((customer, index) => {
            const li = document.createElement('li');
            li.textContent = `${customer[HEADER_PROSPECT_NAME]} (${formatDate(customer[HEADER_DATE])})`;
            li.addEventListener('click', () => showCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function showCustomerDetails(customer) {
        customerDetailsContent.innerHTML = `
            <h3>${customer[HEADER_PROSPECT_NAME]}</h3>
            <p><strong>Branch:</strong> ${customer[HEADER_BRANCH_NAME]}</p>
            <p><strong>Employee:</strong> ${employeeCodeToNameMap[customer[HEADER_EMPLOYEE_CODE]] || customer[HEADER_EMPLOYEE_CODE]} (${customer[HEADER_DESIGNATION]})</p>
            <p><strong>Date:</strong> ${formatDate(customer[HEADER_DATE])}</p>
            <p><strong>Activity Type:</strong> ${customer[HEADER_ACTIVITY_TYPE]}</p>
            <p><strong>Type of Customer:</strong> ${customer[HEADER_TYPE_OF_CUSTOMER]}</p>
            <p><strong>How Contacted:</strong> ${customer[HEADER_HOW_CONTACTED]}</p>
            <p><strong>Phone:</strong> ${customer[HEADER_PHONE_NUMBER_WHATSAPP]}</p>
            <p><strong>Address:</strong> ${customer[HEADER_ADDRESS]}</p>
            <p><strong>Profession:</strong> ${customer[HEADER_PROFESSION]}</p>
            <p><strong>DOB/WD:</strong> ${customer[HEADER_DOB_WD]}</p>
            <p><strong>Product Interested:</strong> ${customer[HEADER_PRODUCT_INTERESTED]}</p>
            <p><strong>Remarks:</strong> ${customer[HEADER_REMARKS]}</p>
            <p><strong>Next Follow-up Date:</strong> ${formatDate(customer[HEADER_NEXT_FOLLOW_UP_DATE])}</p>
            <p><strong>Relation With Staff:</strong> ${customer[HEADER_RELATION_WITH_STAFF]}</p>
            <p><strong>Family Details 1 (Wife/Husband Name):</strong> ${customer[HEADER_FAMILY_DETAILS_1] || 'N/A'}</p>
            <p><strong>Family Details 2 (Wife/Husband Job):</strong> ${customer[HEADER_FAMILY_DETAILS_2] || 'N/A'}</p>
            <p><strong>Family Details 3 (Children Names):</strong> ${customer[HEADER_FAMILY_DETAILS_3] || 'N/A'}</p>
            <p><strong>Family Details 4 (Children Details):</strong> ${customer[HEADER_FAMILY_DETAILS_4] || 'N/A'}</p>
            <p><strong>Profile of Customer:</strong> ${customer[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</p>
        `;
    }


    // --- Google Apps Script (GAS) Communication ---
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage(`Sending data to server for ${action}...`, false);
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
                throw new Error(`Server error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(`${action.replace(/_/g, ' ')} successful! ${result.message || ''}`, false);
                await processData(); // Re-fetch data to update dropdowns and reports
                return true;
            } else {
                displayEmployeeManagementMessage(`${action.replace(/_/g, ' ')} failed: ${result.message || 'Unknown error.'}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Network error or failed to send data: ${error.message}`, true);
            return false;
        }
    }


    // --- Employee Management Form Event Listeners ---

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault(); // Prevent default form submission

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            // Basic validation
            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH_NAME] || !employeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required to add an employee.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset(); // Clear the form
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName) {
                displayEmployeeManagementMessage('Branch Name is required for bulk employee entry.', true);
                return;
            }
            if (!employeeDetails) {
                displayEmployeeManagementMessage('Employee Details are required for bulk entry.', true);
                return;
            }

            const lines = employeeDetails.split('\n').map(line => line.trim()).filter(line => line !== '');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Expect at least Name and Code
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


    // --- Daily Progress Monitoring Logic ---

    // Event listener for the "Show Progress" button
    if (showDailyProgressBtn) {
        showDailyProgressBtn.addEventListener('click', () => {
            const selectedDate = progressDateInput.value; // YYYY-MM-DD
            const selectedBranch = progressBranchSelect.value;
            const selectedEmployee = progressEmployeeSelect.value;

            if (!selectedDate) {
                displayMessage("Please select a date for Daily Progress Monitoring.", 'error');
                return;
            }

            renderDailyProgress(selectedDate, selectedBranch, selectedEmployee);
        });
    }

    // Event listener for Branch selection in Daily Progress Monitoring
    if (progressBranchSelect) {
        progressBranchSelect.addEventListener('change', () => {
            const selectedBranch = progressBranchSelect.value;
            if (selectedBranch) {
                const employeeCodesInBranch = allCanvassingData
                    .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                    .map(entry => entry[HEADER_EMPLOYEE_CODE]);
                const uniqueEmployeeCodes = [...new Set(employeeCodesInBranch)].sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                });
                populateDropdown(progressEmployeeSelect, uniqueEmployeeCodes, true);
            } else {
                // If "All Branches" is selected, populate with all unique employees
                populateDropdown(progressEmployeeSelect, allUniqueEmployees, true);
            }
        });
    }


    function renderDailyProgress(date, branch, employeeCode) {
        dailyProgressDisplay.innerHTML = `<h2>Daily Progress for ${formatDate(date)}</h2>`;

        let filteredEntries = allCanvassingData.filter(entry => formatDate(entry[HEADER_DATE]) === date);

        if (branch) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
        }
        if (employeeCode) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
        }

        if (filteredEntries.length === 0) {
            dailyProgressDisplay.innerHTML += '<p class="no-participation-message">No activity recorded for the selected date and filters.</p>';
            return;
        }

        // Aggregate data per employee for the selected day
        const employeeDailySummary = {}; // {employeeCode: {name, branch, designation, totalVisit, totalCall, totalReference, totalNewCustomerLeads, entries: []}}

        filteredEntries.forEach(entry => {
            const empCode = entry[HEADER_EMPLOYEE_CODE];
            const empName = employeeCodeToNameMap[empCode] || empCode;
            const empBranch = entry[HEADER_BRANCH_NAME];
            const empDesignation = employeeCodeToDesignationMap[empCode] || 'Default';

            if (!employeeDailySummary[empCode]) {
                employeeDailySummary[empCode] = {
                    name: empName,
                    branch: empBranch,
                    designation: empDesignation,
                    totalVisit: 0,
                    totalCall: 0,
                    totalReference: 0,
                    totalNewCustomerLeads: 0,
                    entries: [] // Store individual entries if needed for detailed view
                };
            }

            const trimmedActivityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER] ? entry[HEADER_TYPE_OF_CUSTOMER].trim().toLowerCase() : '';

            if (trimmedActivityType === 'visit') {
                employeeDailySummary[empCode].totalVisit++;
            } else if (trimmedActivityType === 'calls') {
                employeeDailySummary[empCode].totalCall++;
            } else if (trimmedActivityType === 'referance') {
                employeeDailySummary[empCode].totalReference++;
            }

            if (trimmedTypeOfCustomer === 'new') {
                employeeDailySummary[empCode].totalNewCustomerLeads++;
            }
            employeeDailySummary[empCode].entries.push(entry);
        });

        // Create table to display summary
        const table = document.createElement('table');
        table.className = 'daily-progress-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            'Employee Name', 'Branch', 'Designation', 'Visits', 'Calls', 'References', 'New Customer Leads'
        ];
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        Object.values(employeeDailySummary).sort((a, b) => a.name.localeCompare(b.name)).forEach(summary => {
            const row = tbody.insertRow();
            row.insertCell().textContent = summary.name;
            row.insertCell().textContent = summary.branch;
            row.insertCell().textContent = summary.designation;
            row.insertCell().textContent = summary.totalVisit;
            row.insertCell().textContent = summary.totalCall;
            row.insertCell().textContent = summary.totalReference;
            row.insertCell().textContent = summary.totalNewCustomerLeads;
        });
        dailyProgressDisplay.appendChild(table);

        // Optionally, display detailed entries below the summary table if selected
        // For example, a button to show all raw entries for the day, or expand individual employee rows
        dailyProgressDisplay.innerHTML += `<h3 style="margin-top: 30px;">Detailed Activities for ${formatDate(date)}</h3>`;
        renderAllEntries(filteredEntries); // Reuse renderAllEntries to show raw data
    }


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Default to showing the All Branch Snapshot tab
});
