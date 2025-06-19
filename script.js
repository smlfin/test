document.addEventListener('DOMContentLoaded', () => {
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 
    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYF0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=2120288173&single=true&output=csv"; // Marked as UNUSED for clarity, won't be fetched for reports

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

// ------------------------
    const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code'; // Must match your Master Employee Sheet column header
    const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name'; // Must match your Master Employee Sheet column header
    const MASTER_HEADER_DESIGNATION = 'Designation';     // Must match your Master Employee Sheet column header
    const MASTER_HEADER_BRANCH = 'Branch';               // Must match your Master Employee Sheet column header
    // Core Display and Status Elements (existing)
    const reportDisplay = document.getElementById('reportDisplay');
    const statusMessage = document.getElementById('statusMessage');

    // Main Content Sections to toggle (existing)
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    // NEW: Performance Summary Dashboard Section
    const performanceSummarySection = document.getElementById('performanceSummarySection');
    const performanceSummaryDashboardContainer = document.getElementById('performanceSummaryDashboard'); // Get the dashboard container

    // Tab buttons for main navigation (existing)
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const branchPerformanceTabBtn = document.getElementById('branchPerformanceTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    // NEW: Performance Summary Tab Button
    const performanceSummaryTabBtn = document.getElementById('performanceSummaryTabBtn');


    // Dropdowns & Filter Panels (existing)
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    // NEW: Dashboard Branch Select
    const dashboardBranchSelect = document.getElementById('dashboardBranchSelect');

    // View Option Buttons (from your provided list - ensure these are present in your script)
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
    const viewBranchVisitLeaderboardBtn = document.getElementById('viewBranchVisitLeaderboardBtn');
    const viewBranchCallLeaderboardBtn = document.getElementById('viewBranchCallLeaderboardBtn');
    const viewStaffParticipationBtn = document.getElementById('viewStaffParticipationBtn');


    // Detailed Customer View Specific Elements (existing)
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');
    const customerCard1 = document.getElementById('customerCard1');
    const customerCard2 = document.getElementById('customerCard2');
    const customerCard3 = document.getElementById('customerCard3');


    // Employee Management Form Elements (existing)
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');


    // Download Buttons (existing)
    const downloadDetailedCustomerReportBtn = document.getElementById('downloadDetailedCustomerReportBtn'); // If you have this button
    const downloadOverallStaffPerformanceReportBtn = document.getElementById('downloadOverallStaffPerformanceReportBtn');
    // NEW: Performance Summary Dashboard Download Button
    const downloadPerformanceSummaryReportBtn = document.getElementById('downloadPerformanceSummaryReportBtn');

    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    // NEW: Master Employee List Data (will remain empty for now as per your request, until you decide to integrate it)
    let allMasterEmployees = []; // Data from the Master Employee List sheet
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES or canvassing data
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
    // Helper to clear and display messages in a specific div
    function displayMessage(message, type = 'info') {
        // Corrected: Use 'statusMessage' which is declared at the top of DOMContentLoaded
        if (statusMessage) { // Ensure the element exists
            statusMessage.innerHTML = `<div class="message ${type}">${message}</div>`;
            statusMessage.style.display = 'block';
            setTimeout(() => {
                statusMessage.innerHTML = ''; // Clear message
                statusMessage.style.display = 'none';
            }, 5000); // Hide after 5 seconds
        } else {
            console.error("Error: 'statusMessage' element not found in the DOM.");
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
        // Populate dashboard branch select (NEW)
        populateDropdown(dashboardBranchSelect, allUniqueBranches);
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        showTab('allBranchSnapshotTabBtn'); // Your current default starting tab
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
// Function to render Branch Performance Report (d3.PNG style)
    function renderBranchPerformanceReport(selectedBranchName) {
        reportDisplay.innerHTML = ''; // Clear previous content

        if (!selectedBranchName) {
            reportDisplay.innerHTML = '<h2>Branch Performance Report</h2><p>Please select a branch from the dropdown above to view its performance report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Branch Performance Report: ${selectedBranchName} (This Month)</h2>`;
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling
        
        const table = document.createElement('table');
        table.className = 'performance-table';
        
        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        
        // Main Headers
        headerRow.insertCell().textContent = 'Employee Name';
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

        // Get unique employees who have activity in the selected branch this month
        const employeesInSelectedBranchWithActivityThisMonth = [...new Set(allCanvassingData
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entry[HEADER_BRANCH_NAME] === selectedBranchName &&
                       entryDate.getMonth() === currentMonth &&
                       entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

        if (employeesInSelectedBranchWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += `<p>No employee activity found for ${selectedBranchName} for the current month.</p>`;
            return;
        }

        employeesInSelectedBranchWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                entry[HEADER_BRANCH_NAME] === selectedBranchName &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );
            const { totalActivity } = calculateTotalActivity(employeeActivities);
            
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0;
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
    // --- NEW: Function to generate and download the Overall Staff Performance Report as CSV ---
    function downloadOverallStaffPerformanceReportCSV() {
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Get all employees who have had activity this month
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
            displayMessage("No employee activity found for the current month to download.", 'info');
            return;
        }

        // Define metrics for the performance table
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        const csvRows = [];

        // Add main headers
        let headers = ['Employee Name', 'Branch Name', 'Designation'];
        metrics.forEach(metric => {
            headers.push(`${metric} Actual`, `${metric} Target`, `${metric} %`);
        });
        csvRows.push(headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',')); // Quote headers

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );
            const { totalActivity } = calculateTotalActivity(employeeActivities); // Use existing calculation
            
            const targets = TARGETS[designation] || TARGETS['Default']; // Use existing targets
            const performance = calculatePerformance(totalActivity, targets); // Use existing performance calculation

            let rowData = [employeeName, branchName, designation];
            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0;
                let percentValue = performance[metric];
                let displayPercent;

                if (isNaN(percentValue) || targetValue === 0) {
                    displayPercent = 'N/A';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                }
                if (actualValue === 0 && targetValue > 0) {
                    displayPercent = '0%';
                }
                rowData.push(actualValue, targetValue, displayPercent);
            });
            csvRows.push(rowData.map(cell => {
                // Ensure values are properly quoted if they contain commas or quotes
                const stringCell = String(cell);
                return `"${stringCell.replace(/"/g, '""')}"`;
            }).join(','));
        });

        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) { // Feature detection
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', 'Overall_Staff_Performance_Report.csv');
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            displayMessage("Overall Staff Performance Report downloaded successfully!", 'success');
        } else {
            // Fallback for browsers that don't support download attribute
            displayMessage("Your browser does not support automatic downloads. Please copy the data manually.", 'error'); // Optionally display the CSV data for manual copying
            console.log(csvString);
        }
    }
    // --- END NEW ---
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
                                let percentValue = performance[metric]; // Raw numerical percentage
                                let displayPercent;
                                let progressWidth;
                                let progressBarClass;

                                if (isNaN(percentValue) || targetValue === 0) { // Check for NaN or if target is 0
                                    displayPercent = 'N/A';
                                    progressWidth = 0;
                                    progressBarClass = 'no-activity';
                                } else {
                                    displayPercent = `${Math.round(percentValue)}%`; // Round to nearest whole number
                                    progressWidth = Math.min(100, Math.round(percentValue)); // Round for width
                                    progressBarClass = getProgressBarClass(percentValue); // Use original float for color
                                }
                                // Special handling for 0 actuals with positive targets
                                if (actualValue === 0 && targetValue > 0) {
                                    displayPercent = '0%';
                                    progressWidth = 0;
                                    progressBarClass = 'danger'; // Red if 0% and target exists
                                }
                                return `
                                    <tr>
                                        <td data-label="Metric">${metric}</td>
                                        <td data-label="Actual">${actualValue}</td>
                                        <td data-label="Target">${targetValue}</td>
                                        <td data-label="Achievement (%)">${displayPercent}</td>
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

    // Function to show/hide sections based on active tab
    function showTab(activeTabId) {
        // Hide all main content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        performanceSummarySection.style.display = 'none'; // Hide new section

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-btn').forEach(button => {
            button.classList.remove('active');
        });

        // Show the relevant section and activate the clicked tab
        if (activeTabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            allBranchSnapshotTabBtn.classList.add('active');
            renderAllBranchSnapshot();
        } else if (activeTabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            allStaffOverallPerformanceTabBtn.classList.add('active');
            renderOverallStaffPerformanceReport();
        } else if (activeTabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            nonParticipatingBranchesTabBtn.classList.add('active');
            renderNonParticipatingBranches();
        } else if (activeTabId === 'branchPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            branchPerformanceTabBtn.classList.add('active');
            // This tab requires branch selection, so render a placeholder initially
            reportDisplay.innerHTML = '<h2>Branch Performance Report</h2><p>Please select a branch from the dropdown above to view its performance report.</p>';
            // Ensure the branch dropdown and employee filter panel are visible for this tab
            branchSelect.value = ''; // Clear previous branch selection
            employeeFilterPanel.style.display = 'none'; // Hide employee filter until branch selected
            // Also hide viewOptions since they are tied to employee selection for main reports
            viewOptions.style.display = 'none';
            // Populate branch dropdown for main reports section
            populateDropdown(branchSelect, allUniqueBranches);
        } else if (activeTabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'block';
            detailedCustomerViewTabBtn.classList.add('active');
            // Clear customer details content when switching to this tab
            customerCanvassedList.innerHTML = '';
            customerDetailsContent.innerHTML = '<p>Select a branch and employee to view detailed customer entries.</p>';
            customerCard1.innerHTML = '';
            customerCard2.innerHTML = '';
            customerCard3.innerHTML = '';
            // Populate branch dropdown for detailed customer view
            populateDropdown(customerViewBranchSelect, allUniqueBranches);
            // Clear previous selections
            customerViewBranchSelect.value = "";
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select --</option>';
        } else if (activeTabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            employeeManagementTabBtn.classList.add('active');
            // Optionally clear/reset forms when entering this tab
            addEmployeeForm.reset();
            bulkAddEmployeeForm.reset();
            deleteEmployeeForm.reset();
            displayEmployeeManagementMessage('', false); // Clear any old messages
        } else if (activeTabId === 'performanceSummaryTabBtn') {
            performanceSummarySection.style.display = 'block';
            performanceSummaryTabBtn.classList.add('active');
            renderPerformanceSummaryDashboard(); // Call the new function for the dashboard
        }
    }

    // Attach event listeners to tab buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    branchPerformanceTabBtn.addEventListener('click', () => showTab('branchPerformanceTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    performanceSummaryTabBtn.addEventListener('click', () => showTab('performanceSummaryTabBtn')); // Attach listener

    // Event listener for branch dropdown change for detailed customer view
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        if (selectedBranch) {
            // Filter employees who have entries in the selected branch for detailed view
            const employeesInBranch = [...new Set(allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                });
            populateDropdown(customerViewEmployeeSelect, employeesInBranch, true);
            customerViewEmployeeSelect.value = ""; // Reset employee selection
            customerCanvassedList.innerHTML = '';
            customerDetailsContent.innerHTML = '<p>Select an employee to view their detailed customer entries.</p>';
            customerCard1.innerHTML = '';
            customerCard2.innerHTML = '';
            customerCard3.innerHTML = '';
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select --</option>';
            customerCanvassedList.innerHTML = '';
            customerDetailsContent.innerHTML = '<p>Select a branch and employee to view detailed customer entries.</p>';
            customerCard1.innerHTML = '';
            customerCard2.innerHTML = '';
            customerCard3.innerHTML = '';
        }
    });

    // Event listener for employee dropdown change for detailed customer view
    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        if (selectedBranch && selectedEmployeeCode) {
            renderDetailedCustomerView(selectedBranch, selectedEmployeeCode);
        } else {
            customerCanvassedList.innerHTML = '';
            customerDetailsContent.innerHTML = '<p>Select a branch and employee to view detailed customer entries.</p>';
            customerCard1.innerHTML = '';
            customerCard2.innerHTML = '';
            customerCard3.innerHTML = '';
        }
    });

    // Function to render the detailed customer view
    function renderDetailedCustomerView(branchName, employeeCode) {
        customerCanvassedList.innerHTML = ''; // Clear previous list
        customerDetailsContent.innerHTML = ''; // Clear previous details

        const employeeActivities = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode
        );

        if (employeeActivities.length === 0) {
            customerDetailsContent.innerHTML = `<p>No customer entries found for ${employeeCodeToNameMap[employeeCode] || employeeCode} in ${branchName}.</p>`;
            return;
        }

        const ul = document.createElement('ul');
        employeeActivities.forEach((entry, index) => {
            const li = document.createElement('li');
            const prospectName = entry[HEADER_PROSPECT_NAME] || 'N/A';
            const activityType = entry[HEADER_ACTIVITY_TYPE] || 'N/A';
            const date = formatDate(entry[HEADER_DATE]) || 'N/A';
            li.textContent = `${prospectName} - ${activityType} on ${date}`;
            li.dataset.index = index; // Store original index for easy lookup
            li.addEventListener('click', () => displayCustomerDetails(entry));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);

        // Display details of the first customer by default
        if (employeeActivities.length > 0) {
            displayCustomerDetails(employeeActivities[0]);
        }
    }

    // Function to display detailed customer information
    function displayCustomerDetails(customerEntry) {
        customerDetailsContent.innerHTML = ''; // Clear previous details

        const detailsHtml = `
            <h3>Customer Details: ${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</h3>
            <p><strong>Branch:</strong> ${customerEntry[HEADER_BRANCH_NAME] || 'N/A'}</p>
            <p><strong>Employee:</strong> ${employeeCodeToNameMap[customerEntry[HEADER_EMPLOYEE_CODE]] || customerEntry[HEADER_EMPLOYEE_CODE]}</p>
            <p><strong>Date:</strong> ${formatDate(customerEntry[HEADER_DATE]) || 'N/A'}</p>
            <p><strong>Activity Type:</strong> ${customerEntry[HEADER_ACTIVITY_TYPE] || 'N/A'}</p>
            <p><strong>Customer Type:</strong> ${customerEntry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</p>
            <p><strong>Lead Source:</strong> ${customerEntry[HEADER_R_LEAD_SOURCE] || 'N/A'}</p>
            <p><strong>How Contacted:</strong> ${customerEntry[HEADER_HOW_CONTACTED] || 'N/A'}</p>
            <p><strong>Phone/Whatsapp:</strong> ${customerEntry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</p>
            <p><strong>Address:</strong> ${customerEntry[HEADER_ADDRESS] || 'N/A'}</p>
            <p><strong>Profession:</strong> ${customerEntry[HEADER_PROFESSION] || 'N/A'}</p>
            <p><strong>DOB/WD:</strong> ${customerEntry[HEADER_DOB_WD] || 'N/A'}</p>
            <p><strong>Product Interested:</strong> ${customerEntry[HEADER_PRODUCT_INTERESTED] || 'N/A'}</p>
            <p><strong>Remarks:</strong> ${customerEntry[HEADER_REMARKS] || 'N/A'}</p>
            <p><strong>Next Follow-up Date:</strong> ${formatDate(customerEntry[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}</p>
            <p><strong>Relation With Staff:</strong> ${customerEntry[HEADER_RELATION_WITH_STAFF] || 'N/A'}</p>
            <p><strong>Family Details (Wife/Husband Name):</strong> ${customerEntry[HEADER_FAMILY_DETAILS_1] || 'N/A'}</p>
            <p><strong>Family Details (Wife/Husband Job):</strong> ${customerEntry[HEADER_FAMILY_DETAILS_2] || 'N/A'}</p>
            <p><strong>Family Details (Children Names):</strong> ${customerEntry[HEADER_FAMILY_DETAILS_3] || 'N/A'}</p>
            <p><strong>Family Details (Children Details):</strong> ${customerEntry[HEADER_FAMILY_DETAILS_4] || 'N/A'}</p>
            <p><strong>Profile of Customer:</strong> ${customerEntry[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</p>
        `;
        customerDetailsContent.innerHTML = detailsHtml;
    }

    // --- Employee Management Functions (existing, provided as is) ---

    // Function to add a new employee
    addEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        const newEmployee = {
            name: newEmployeeNameInput.value.trim(),
            code: newEmployeeCodeInput.value.trim(),
            branch: newBranchNameInput.value.trim(),
            designation: newDesignationInput.value.trim()
        };

        if (!newEmployee.name || !newEmployee.code || !newEmployee.branch || !newEmployee.designation) {
            displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
            return;
        }

        try {
            const response = await fetch(`${WEB_APP_URL}?action=addEmployee`, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8', // Required for Google Apps Script
                },
                body: JSON.stringify(newEmployee)
            });

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                addEmployeeForm.reset();
                // Optionally re-fetch data or update UI immediately
                await processData(); // Re-process data to update dropdowns, etc.
            } else {
                displayEmployeeManagementMessage(`Error adding employee: ${result.message}`, true);
            }
        } catch (error) {
            console.error('Error adding employee:', error);
            displayEmployeeManagementMessage(`Network error adding employee: ${error.message}`, true);
        }
    });

    // Function for bulk adding employees
    bulkAddEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        const bulkBranch = bulkEmployeeBranchNameInput.value.trim();
        const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

        if (!bulkBranch || !bulkDetails) {
            displayEmployeeManagementMessage('Branch and employee details are required for bulk add.', true);
            return;
        }

        const employeeList = bulkDetails.split('\n').map(line => {
            const parts = line.split(',');
            if (parts.length === 2) { // Assuming format: Name,Designation
                return { name: parts[0].trim(), designation: parts[1].trim(), branch: bulkBranch };
            }
            return null;
        }).filter(item => item !== null);

        if (employeeList.length === 0) {
            displayEmployeeManagementMessage('No valid employee entries found in the bulk details. Format: Name,Designation (one per line).', true);
            return;
        }

        try {
            const response = await fetch(`${WEB_APP_URL}?action=bulkAddEmployees`, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8',
                },
                body: JSON.stringify({ employees: employeeList })
            });

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                bulkAddEmployeeForm.reset();
                await processData(); // Re-process data
            } else {
                displayEmployeeManagementMessage(`Error bulk adding employees: ${result.message}`, true);
            }
        } catch (error) {
            console.error('Error bulk adding employees:', error);
            displayEmployeeManagementMessage(`Network error bulk adding employees: ${error.message}`, true);
        }
    });

    // Function to delete an employee
    deleteEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

        if (!employeeCodeToDelete) {
            displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
            return;
        }

        if (!confirm(`Are you sure you want to delete employee with code: ${employeeCodeToDelete}? This action cannot be undone.`)) {
            return; // User cancelled
        }

        try {
            const response = await fetch(`${WEB_APP_URL}?action=deleteEmployee`, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8',
                },
                body: JSON.stringify({ employeeCode: employeeCodeToDelete })
            });

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                deleteEmployeeForm.reset();
                await processData(); // Re-process data
            } else {
                displayEmployeeManagementMessage(`Error deleting employee: ${result.message}`, true);
            }
        } catch (error) {
            console.error('Error deleting employee:', error);
            displayEmployeeManagementMessage(`Network error deleting employee: ${error.message}`, true);
        }
    });


    // --- View Option Button Event Listeners (for main reports section) ---
    // These listeners handle which report to display based on button clicks
    const viewOptions = document.querySelector('.view-options'); // Assuming a container for these buttons

    if (viewOptions) {
        viewOptions.addEventListener('click', (event) => {
            // Deactivate all buttons
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            
            const clickedButton = event.target.closest('.btn');
            if (clickedButton) {
                clickedButton.classList.add('active'); // Activate clicked button

                // Determine which report to render based on the button's ID
                if (clickedButton.id === 'viewBranchPerformanceReportBtn') {
                    renderBranchPerformanceReport(branchSelect.value); // Use the currently selected branch
                } else if (clickedButton.id === 'viewEmployeeSummaryBtn') {
                    // This is handled by employeeSelect change listener, but also callable directly
                    if (selectedEmployeeCodeEntries.length > 0) {
                        renderEmployeeSummary(selectedEmployeeCodeEntries);
                    } else {
                        reportDisplay.innerHTML = '<p>Please select an employee to view their summary.</p>';
                    }
                } else if (clickedButton.id === 'viewAllEntriesBtn') {
                    if (branchSelect.value) { // Ensure a branch is selected
                        const entriesToDisplay = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchSelect.value);
                        renderAllEntries(entriesToDisplay);
                    } else {
                        reportDisplay.innerHTML = '<p>Please select a branch to view all entries.</p>';
                    }
                } else if (clickedButton.id === 'viewPerformanceReportBtn') {
                    // This button doesn't have a specific rendering function in the current script
                    // It might be a duplicate of 'viewBranchPerformanceReportBtn' or intended for a different report
                    reportDisplay.innerHTML = '<p>Performance report display logic here. Please specify which performance report you need.</p>';
                } else if (clickedButton.id === 'viewBranchVisitLeaderboardBtn') {
                    renderBranchLeaderboard('Visit', branchSelect.value);
                } else if (clickedButton.id === 'viewBranchCallLeaderboardBtn') {
                    renderBranchLeaderboard('Call', branchSelect.value);
                } else if (clickedButton.id === 'viewStaffParticipationBtn') {
                     renderStaffParticipation(branchSelect.value);
                }
            }
        });
    }

    // Function to render an employee's summary (d4.PNG style)
    function renderEmployeeSummary(employeeEntries) {
        reportDisplay.innerHTML = ''; // Clear previous content

        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity data for the selected employee in this branch for the current month.</p>';
            return;
        }

        const employeeCode = employeeEntries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const branchName = employeeEntries[0][HEADER_BRANCH_NAME];
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        reportDisplay.innerHTML = `
            <h2>Employee Summary: ${employeeName} (${designation}) - ${branchName}</h2>
            <h3>This Month's Activity</h3>
        `;

        const { totalActivity, productInterests } = calculateTotalActivity(employeeEntries);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'summary-table';
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Metric</th>
                    <th>Actual</th>
                    <th>Target</th>
                    <th>% Achieved</th>
                </tr>
            </thead>
            <tbody>
                ${Object.keys(targets).map(metric => {
                    const actualValue = totalActivity[metric] || 0;
                    const targetValue = targets[metric] || 0;
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
        `;
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);

        // Display Product Interests
        if (productInterests.length > 0) {
            const productListDiv = document.createElement('div');
            productListDiv.innerHTML = '<h3>Products Interested:</h3>';
            const ul = document.createElement('ul');
            productInterests.forEach(product => {
                const li = document.createElement('li');
                li.textContent = product;
                ul.appendChild(li);
            });
            productListDiv.appendChild(ul);
            reportDisplay.appendChild(productListDiv);
        } else {
            reportDisplay.innerHTML += '<p>No specific product interests recorded for this employee this month.</p>';
        }
    }

    // Function to render all entries for a selected branch (d2.PNG style)
    function renderAllEntries(entriesToDisplay) {
        reportDisplay.innerHTML = '<h2>All Entries</h2>';
        if (entriesToDisplay.length === 0) {
            reportDisplay.innerHTML += '<p>No entries found for the selected criteria.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'all-entries-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_DATE, HEADER_EMPLOYEE_NAME, HEADER_BRANCH_NAME, HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER, HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP,
            HEADER_PRODUCT_INTERESTED, HEADER_REMARKS
        ]; // Display a subset of relevant headers
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entriesToDisplay.forEach(entry => {
            const row = tbody.insertRow();
            row.insertCell().textContent = formatDate(entry[HEADER_DATE]) || 'N/A';
            row.insertCell().textContent = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_CODE];
            row.insertCell().textContent = entry[HEADER_BRANCH_NAME] || 'N/A';
            row.insertCell().textContent = entry[HEADER_ACTIVITY_TYPE] || 'N/A';
            row.insertCell().textContent = entry[HEADER_TYPE_OF_CUSTOMER] || 'N/A';
            row.insertCell().textContent = entry[HEADER_PROSPECT_NAME] || 'N/A';
            row.insertCell().textContent = entry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A';
            row.insertCell().textContent = entry[HEADER_PRODUCT_INTERESTED] || 'N/A';
            row.insertCell().textContent = entry[HEADER_REMARKS] || 'N/A';
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Function to render branch leaderboards (Visit/Call)
    function renderBranchLeaderboard(activityMetric, selectedBranchName) {
        reportDisplay.innerHTML = `<h2>${activityMetric} Leaderboard: ${selectedBranchName || 'All Branches'} (This Month)</h2>`;

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        let filteredEntries = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        if (selectedBranchName) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranchName);
        }

        if (filteredEntries.length === 0) {
            reportDisplay.innerHTML += `<p>No ${activityMetric.toLowerCase()} activity found for ${selectedBranchName || 'all branches'} this month.</p>`;
            return;
        }

        // Aggregate activity counts by employee
        const employeeActivityCounts = {};
        filteredEntries.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            if (employeeCode && activityType && activityType.toLowerCase() === activityMetric.toLowerCase()) {
                employeeActivityCounts[employeeCode] = (employeeActivityCounts[employeeCode] || 0) + 1;
            } else if (employeeCode && activityMetric === 'Call' && activityType.toLowerCase() === 'calls') {
                 employeeActivityCounts[employeeCode] = (employeeActivityCounts[employeeCode] || 0) + 1;
            } else if (employeeCode && activityMetric === 'Reference' && activityType.toLowerCase() === 'referance') {
                 employeeActivityCounts[employeeCode] = (employeeActivityCounts[employeeCode] || 0) + 1;
            }
        });

        // Convert to array and sort descending
        const sortedEmployees = Object.entries(employeeActivityCounts)
            .sort(([, countA], [, countB]) => countB - countA)
            .map(([employeeCode, count]) => ({
                employeeCode,
                employeeName: employeeCodeToNameMap[employeeCode] || employeeCode,
                count: count
            }));

        if (sortedEmployees.length === 0) {
            reportDisplay.innerHTML += `<p>No ${activityMetric.toLowerCase()} activity found for ${selectedBranchName || 'all branches'} this month.</p>`;
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'leaderboard-table';
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Rank</th>
                    <th>Employee Name</th>
                    <th>${activityMetric} Count</th>
                </tr>
            </thead>
            <tbody>
                ${sortedEmployees.map((employee, index) => `
                    <tr>
                        <td>${index + 1}</td>
                        <td>${employee.employeeName}</td>
                        <td>${employee.count}</td>
                    </tr>
                `).join('')}
            </tbody>
        `;
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }
    // Render Staff Participation Report
    function renderStaffParticipation(selectedBranchName) {
        reportDisplay.innerHTML = `<h2>Staff Participation Report: ${selectedBranchName || 'All Branches'} (This Month)</h2>`;

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        let relevantCanvassingData = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        if (selectedBranchName) {
            relevantCanvassingData = relevantCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranchName);
        }

        if (relevantCanvassingData.length === 0) {
            reportDisplay.innerHTML += '<p>No activity found for the current month in the selected branch.</p>';
            return;
        }

        const employeesWithActivity = new Set(relevantCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]));

        const allEmployeesInScope = new Set();
        // Add all employees from master data
        // allMasterEmployees.forEach(emp => allEmployeesInScope.add(emp[MASTER_HEADER_EMPLOYEE_CODE]));
        // Add all employees from canvassing data (who might not be in master yet)
        allCanvassingData.forEach(entry => allEmployeesInScope.add(entry[HEADER_EMPLOYEE_CODE]));

        let employeesToConsider = [...allEmployeesInScope];
        if (selectedBranchName) {
            // Filter employees by branch from allCanvassingData (as master is not fully integrated yet)
            const employeesInSelectedBranchActivities = new Set(allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranchName)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]));
            employeesToConsider = employeesToConsider.filter(code => employeesInSelectedBranchActivities.has(code));
        }

        employeesToConsider.sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        if (employeesToConsider.length === 0) {
            reportDisplay.innerHTML += '<p>No employees found for participation tracking in the selected criteria.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'participation-table';
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Branch</th>
                    <th>Designation</th>
                    <th>Participation Status (This Month)</th>
                </tr>
            </thead>
            <tbody>
                ${employeesToConsider.map(employeeCode => {
                    const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
                    const employeeBranch = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
                    const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'N/A';
                    const hasParticipated = employeesWithActivity.has(employeeCode);
                    const statusClass = hasParticipated ? 'status-participated' : 'status-non-participated';
                    const statusText = hasParticipated ? 'Participated' : 'No Activity';
                    return `
                        <tr>
                            <td>${employeeName}</td>
                            <td>${employeeBranch}</td>
                            <td>${employeeDesignation}</td>
                            <td><span class="${statusClass}">${statusText}</span></td>
                        </tr>
                    `;
                }).join('')}
            </tbody>
        `;
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }
    // --- NEW: Performance Summary Dashboard (as requested by user) ---
    function renderPerformanceSummaryDashboard() {
        if (!performanceSummaryDashboardContainer) {
            console.error("Dashboard container with ID 'performanceSummaryDashboard' not found.");
            return;
        }

        performanceSummaryDashboardContainer.innerHTML = ''; // Clear existing content
        performanceSummaryDashboardContainer.innerHTML = '<h2>Performance Summary Dashboard</h2>';

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Filter data for the current month/year
        const currentMonthActivities = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        // 1. Total Staff
        // Assuming 'allUniqueEmployees' contains all employees known to the system from canvassing data
        const totalStaff = allUniqueEmployees.length; 

        // 2. Tracked – BRANCH (Count & %)
        const branchesWithActivity = new Set(currentMonthActivities.map(entry => entry[HEADER_BRANCH_NAME]));
        const trackedBranchesCount = branchesWithActivity.size;
        const totalPredefinedBranches = PREDEFINED_BRANCHES.length;
        const trackedBranchesPercentage = totalPredefinedBranches > 0 ? ((trackedBranchesCount / totalPredefinedBranches) * 100).toFixed(2) : 0;

        let dashboardHtml = `
            <div class="dashboard-section">
                <h3>Overall Summary</h3>
                <table class="summary-table">
                    <tr>
                        <td><strong>Total Staff:</strong></td>
                        <td>${totalStaff}</td>
                    </tr>
                    <tr>
                        <td><strong>Tracked Branches:</strong></td>
                        <td>${trackedBranchesCount} out of ${totalPredefinedBranches} (${trackedBranchesPercentage}%)</td>
                    </tr>
                </table>
            </div>
            <div class="dashboard-section">
                <h3>Activity Performance Summary (This Month)</h3>
                <table class="summary-table activity-summary-table">
                    <thead>
                        <tr>
                            <th>Activity Type</th>
                            <th>Total Employees Participated (Count & %)</th>
                            <th>Fully Completed (Count & %)</th>
                            <th>Partially Completed (Count & %)</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        const activityTypeMapping = {
            'Visit': 'visit',
            'Call': 'calls',
            'Reference': 'referance', // Note the typo 'referance' in your sheet for 'Reference'
            'New Customer Leads': 'new' // This is based on HEADER_TYPE_OF_CUSTOMER === 'new'
        };

        metrics.forEach(metric => {
            let employeesParticipated = new Set();
            let fullyCompletedCount = 0;
            let partiallyCompletedCount = 0;
            const employeesInCurrentMonth = [...new Set(currentMonthActivities.map(entry => entry[HEADER_EMPLOYEE_CODE]))];

            employeesInCurrentMonth.forEach(employeeCode => {
                const employeeActivities = currentMonthActivities.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
                const { totalActivity } = calculateTotalActivity(employeeActivities);
                const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
                const targets = TARGETS[designation] || TARGETS['Default'];
                const performance = calculatePerformance(totalActivity, targets);

                let actualForMetric = 0;
                let targetForMetric = 0;

                if (metric === 'New Customer Leads') {
                    actualForMetric = totalActivity['New Customer Leads'] || 0;
                    targetForMetric = targets['New Customer Leads'] || 0;
                    // For 'New Customer Leads', an employee "participated" if they have any new customer lead activity.
                    const hasNewLeadActivity = employeeActivities.some(entry => (entry[HEADER_TYPE_OF_CUSTOMER] || '').trim().toLowerCase() === 'new');
                    if (hasNewLeadActivity) {
                        employeesParticipated.add(employeeCode);
                    }
                } else {
                    // For Visit, Call, Reference, participation is based on having any entry for that activity type
                    actualForMetric = totalActivity[metric] || 0;
                    targetForMetric = targets[metric] || 0;
                    const hasActivity = employeeActivities.some(entry => 
                        (entry[HEADER_ACTIVITY_TYPE] || '').trim().toLowerCase() === activityTypeMapping[metric]
                    );
                    if (hasActivity) {
                        employeesParticipated.add(employeeCode);
                    }
                }
                
                // Calculate fully/partially completed based on performance percentage against targets
                const metricPerformance = performance[metric];

                if (targetForMetric > 0) { // Only consider completion if there's a target
                    if (metricPerformance >= 100) {
                        fullyCompletedCount++;
                    } else if (metricPerformance > 0 && metricPerformance < 100) {
                        partiallyCompletedCount++;
                    }
                } else if (actualForMetric > 0) {
                    // If no target, but actual activity exists, consider it "partially completed"
                    // Or you might decide to count this differently. For now, putting it in partially.
                    // This is a design choice; you might want a separate category for "achieved without target"
                    partiallyCompletedCount++; 
                }
            });

            const totalEmployeesParticipatedCount = employeesParticipated.size;
            const totalEmployeesOverall = allUniqueEmployees.length; // Total staff in the system

            const participatedPercentage = totalEmployeesOverall > 0 ? ((totalEmployeesParticipatedCount / totalEmployeesOverall) * 100).toFixed(2) : 0;
            const fullyCompletedPercentage = totalEmployeesOverall > 0 ? ((fullyCompletedCount / totalEmployeesOverall) * 100).toFixed(2) : 0;
            const partiallyCompletedPercentage = totalEmployeesOverall > 0 ? ((partiallyCompletedCount / totalEmployeesOverall) * 100).toFixed(2) : 0;


            dashboardHtml += `
                <tr>
                    <td>${metric}</td>
                    <td>${totalEmployeesParticipatedCount} (${participatedPercentage}%)</td>
                    <td>${fullyCompletedCount} (${fullyCompletedPercentage}%)</td>
                    <td>${partiallyCompletedCount} (${partiallyCompletedPercentage}%)</td>
                </tr>
            `;
        });

        dashboardHtml += `
                    </tbody>
                </table>
            </div>
        `;

        performanceSummaryDashboardContainer.innerHTML += dashboardHtml;
    }


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Your current default starting tab
}); // This is the closing brace for DOMContentLoaded
